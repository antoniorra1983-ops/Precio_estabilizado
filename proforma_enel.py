"""Lectura de proformas Enel EFE Valparaíso (PROFORMA-Q/L y CSP-Q/L).

Autoridad de cargos: columna F de las cuatro hojas PROFORMA. Las hojas
Informe, RELIQ_*, CUT_CSP, etc. son antecedentes y NO cargos nuevos.
El MES DE PROFORMA se conserva para el total facturado, pero PE, SSCC,
MT, PD y la liquidación de peaje individualizado se atribuyen al MES ANTERIOR.
Una línea con mes de reliquidación escrito explícitamente usa ese mes.
Esto es clasificación temporal: no mueve ni duplica importes en CLP.
"""
from __future__ import annotations

import csv
from datetime import date, timedelta
from io import BytesIO, StringIO
import re
import unicodedata
from xml.etree import ElementTree as ET
from zipfile import ZipFile

try:
    import streamlit as st
except ImportError:
    # GitHub Actions procesa Excel sin interfaz ni dependencia de Streamlit.
    class _StreamlitSinInterfaz:
        @staticmethod
        def fragment(func):
            return func
    st = _StreamlitSinInterfaz()

XML = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
SITIOS = {'Quilpué': 'Q', 'Limache': 'L'}
MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
         'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
MESES_CORTOS = dict(zip(('ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'),range(1,13)))


def _norm(texto):
    return ''.join(c for c in unicodedata.normalize('NFKD',str(texto or ''))
                   if not unicodedata.combining(c)).upper().strip()


def periodo_archivo(nombre):
    """Soporta EFE_VALPO_SEP2026, EFE_VALPO_2609, EFE_VALPO_2026-09."""
    n = _norm(nombre)
    m = re.search(r'(20\d{2})[-_ ](0[1-9]|1[0-2])',n)
    if m:
        return f'{int(m[1]):04d}-{int(m[2]):02d}'
    m = re.search(r'(?<!\d)(\d{2})(0[1-9]|1[0-2])(?:[PD])?(?!\d)',n)
    if m:
        return f'20{m[1]}-{m[2]}'
    m = re.search(r'(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)[A-Z]*[-_ ]*(20\d{2}|\d{2})(?!\d)',n)
    if m:
        a=int(m[2]); a=a if a>=2000 else 2000+a
        return f'{a:04d}-{MESES_CORTOS[m[1]]:02d}'
    return None


def _periodo_celda(valor):
    if isinstance(valor, (int,float)) and 30000 < valor < 70000:
        d=date(1899,12,30)+timedelta(days=int(valor))
        return d.strftime('%Y-%m')
    s=str(valor or '').strip()
    m=re.search(r'(\d{1,2})/(\d{1,2})/(20\d{2})',s)
    if m:
        return f'{int(m[3]):04d}-{int(m[2]):02d}'
    m=re.search(r'\b(20\d{2})-(0[1-9]|1[0-2])(?:-\d{2})?\b',s)
    if m:
        return f'{m[1]}-{m[2]}'
    return None


def _float(valor):
    try:
        return float(valor) if valor is not None and valor!='' else None
    except (ValueError,TypeError):
        return None


class _Excel:
    """Lector ligero OOXML; utiliza valores ya calculados en el Excel.

    No ejecuta macros ni evalúa fórmulas; celdas F sin resultado guardado
    se informan como pendientes, nunca como importe cero.
    """
    def __init__(self,data):
        self.z=ZipFile(BytesIO(data))
        root=ET.fromstring(self.z.read('xl/workbook.xml'))
        rel=ET.fromstring(self.z.read('xl/_rels/workbook.xml.rels'))
        targets={r.attrib['Id']:r.attrib['Target'] for r in rel}
        self.paths={}
        for e in root.find(XML+'sheets'):
            target=targets[e.attrib[RID]].lstrip('/')
            if not target.startswith('xl/'):
                target='xl/'+target
            self.paths[_norm(e.attrib['name'])]=target
        self.shared=[]
        if 'xl/sharedStrings.xml' in self.z.namelist():
            with self.z.open('xl/sharedStrings.xml') as fh:
                for _,node in ET.iterparse(fh,events=('end',)):
                    if node.tag==XML+'si':
                        self.shared.append(''.join(t.text or '' for t in node.iter(XML+'t')))
                        node.clear()

    def filas(self,nombre):
        path=self.paths.get(_norm(nombre))
        if not path:
            return None
        resultado={}
        with self.z.open(path) as fh:
            for _,e in ET.iterparse(fh, events=('end',)):
                if e.tag!=XML+'row':
                    continue
                fila=int(e.attrib.get('r',0)); celdas={}
                for c in e.iter(XML+'c'):
                    celda=c.attrib.get('r','')
                    mt=re.match(r'([A-Z]+)\d+',celda)
                    if not mt:
                        continue
                    col=mt[1]
                    v=c.find(XML+'v'); v=v.text if v is not None else None
                    tipo=c.attrib.get('t')
                    if tipo=='s' and v is not None:
                        v=self.shared[int(v)]
                    elif tipo=='inlineStr':
                        ins=c.find(XML+'is')
                        v=''.join(x.text or '' for x in ins.iter(XML+'t')) if ins is not None else ''
                    elif v is not None and tipo not in ('str','e'):
                        try: v=float(v)
                        except ValueError: pass
                    if v is not None:
                        celdas[col]=v
                resultado[fila]=celdas
                e.clear()
        return resultado


# Conceptos respaldados por balance CEN del mes anterior a la proforma.
# No extender esta regla a Tx Nacional/Zonal corriente ni al peaje distribución.
DESFASADOS = frozenset((
    'Servicios complementarios', 'Mínimos técnicos (MT)',
    'Partidas y detenciones (PD)', 'Precio Estabilizado',
    'Liquidación de peaje',
))


def mes_anterior(periodo):
    """Mes del balance que origina una proforma, incluso enero -> diciembre."""
    yy, mm = map(int, periodo.split('-'))
    return f'{yy - (mm == 1):04d}-{12 if mm == 1 else mm - 1:02d}'


def _mes_mencionado_en_reliquidacion(nombre):
    """Si la partida dice 'reliquidación mes de agosto de 2026', respetarlo."""
    s = _norm(nombre)
    if 'RELIQUID' not in s and 'LIQUIDAC' not in s:
        return None
    patron = r'\b(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)[A-Z]*\s+(?:DE\s+)?(20\d{2})\b'
    m = re.search(patron, s)
    if m:
        return f'{int(m.group(2)):04d}-{MESES_CORTOS[m.group(1)]:02d}'
    m = re.search(r'\b(20\d{2})[-/](0[1-9]|1[0-2])\b', s)
    return f'{m.group(1)}-{m.group(2)}' if m else None


def _grupo(nombre):
    s=_norm(nombre)
    if 'RELIQUIDAC' in s: return 'Reliquidación anterior'
    if 'CONSUMO DE ENERGIA' in s: return 'Energía contractual'
    if 'POTENCIA' in s: return 'Potencia de punta'
    if 'PUBLICO' in s: return 'Servicio público CSP'
    if 'DISTRIBUCION' in s: return 'Peaje distribución'
    if 'PEAJE' in s and ('CLIENTE INDIV' in s or 'TX NACIONAL' in s or 'LIQUIDAC' in s):
        return 'Liquidación de peaje'
    if 'SOBRECOST' in s and re.search(r'\bMT\b', s):
        return 'Mínimos técnicos (MT)'
    if 'SOBRECOST' in s and re.search(r'\bPD\b', s):
        return 'Partidas y detenciones (PD)'
    if 'ESTABILIZ' in s or 'COMPENSACION PE' in s: return 'Precio Estabilizado'
    if 'SERVICIOS COMPLEMENTARIOS' in s: return 'Servicios complementarios'
    if 'TRANSMISION' in s or 'TX NACIONAL' in s: return 'Transmisión'
    if 'RH' in s: return 'Reservas hidráulicas (RH)'
    return 'Otros cargos'


def origen_cargo(nombre, mes_proforma):
    """Clasifica fecha de origen sin cambiar el importe de la línea de proforma.

    Sólo PE, SSCC, MT, PD y liquidación de peaje retroceden un mes por
    regla. Las reliquidaciones explícitas de otros meses usan el mes escrito.
    """
    grupo = _grupo(nombre)
    if grupo == 'Reliquidación anterior':
        explicito = _mes_mencionado_en_reliquidacion(nombre)
        if explicito:
            return explicito, 'Mes explícito de reliquidación en la proforma'
        return mes_anterior(mes_proforma), 'Reliquidación sin mes explícito: mes anterior'
    if grupo in DESFASADOS:
        return mes_anterior(mes_proforma), 'Liquidación del balance CEN anterior'
    return mes_proforma, 'Cargo del período corriente de la proforma'


def _valor_fijo(filas, campo, desde, hasta):
    for n in range(desde,hasta+1):
        r=filas.get(n,{})
        if _norm(r.get('E'))==_norm(campo):
            return _float(r.get('F'))
    return None


def _mwh(filas):
    for n in (18,19,20):
        s=str(filas.get(n,{}).get('D') or '')
        m=re.search(r'([\d.,]+)\s*MWh\s+Consumidos',s,re.I)
        if m:
            return float(m[1].replace(',','.'))
    return None


def _mw(filas):
    for n in range(18,24):
        s=str(filas.get(n,{}).get('D') or filas.get(n,{}).get('C') or '')
        m=re.search(r'POTENCIA\s+DE\s+PUNTA:\s*([\d.,]+)\s*MW',s,re.I)
        if m:
            return float(m[1].replace(',','.'))
    return None


def _cargos(filas, hoja, sitio, periodo, mwh, potencia_mw=None, inicio=16, fin=43):
    salida=[]
    for n in range(inicio,fin+1):
        r=filas.get(n,{})
        nombre=str(r.get('D') or '').strip()
        if not nombre:
            continue
        if not ('CARGO' in _norm(nombre) or 'POTENCIA DE PUNTA' in _norm(nombre)
                or 'SOBRECOSTO' in _norm(nombre) or 'COMPENSAC' in _norm(nombre)
                or 'SERVICIOS COMPLEMENTARIOS' in _norm(nombre)
                or 'COSTO OPORTUNIDAD' in _norm(nombre)
                or 'PEAJE' in _norm(nombre) or 'RELIQUIDAC' in _norm(nombre)):
            continue
        monto=_float(r.get('F'))
        # No se crean montos a partir de la columna G/H: solo F es monto de proforma.
        unidad=(monto/(mwh*1000)) if monto is not None and mwh and mwh>0 else None
        potencia_unit=(monto/(potencia_mw*1000)) if (monto is not None and potencia_mw
                     and potencia_mw>0 and _grupo(nombre)=='Potencia de punta') else None
        periodo_origen, criterio = origen_cargo(nombre, periodo)
        salida.append({'Período':periodo,'Mes de proforma':periodo,
                       'Mes de origen':periodo_origen,
                       'Criterio período':criterio,
                       'Suministro':sitio,'Grupo':_grupo(nombre),
                       'Concepto':nombre,'Monto CLP':monto,
                       'Equivalente CLP/kWh':unidad,'Potencia CLP/kW/mes aprox.':potencia_unit,
                       'Referencia G':_float(r.get('G')),
                       'Variación H':_float(r.get('H')),
                       'Origen':f'{hoja}!F{n}'})
    return salida


def parsear_proforma(datos, archivo):
    libro=_Excel(datos)
    nombre_periodo=periodo_archivo(archivo)
    resultado={'archivo':archivo,'periodo':None,'suministros':{},'avisos':[],
               'hojas_disponibles':list(libro.paths.keys())}
    # El mes de balance declarado por ENEL en CARGOS_CEN!C2 o D2 sirve como
    # contraste documental; no altera el mes de factura ni inventa montos.
    hoja_cen=libro.filas('CARGOS_CEN')
    resultado['balance_cen_declarado']=(
        (_periodo_celda(hoja_cen.get(2,{}).get('C')) or
         _periodo_celda(hoja_cen.get(2,{}).get('D'))) if hoja_cen else None)
    # Fecha que acompaña la tabla de peajes individualizados del anexo.
    # Puede ser la referencia usada por Enel, no necesariamente el período
    # facturable; informar sin sobrescribir el criterio indicado por EFE.
    resultado['mes_referencia_peaje_anexo']=(
        _periodo_celda(hoja_cen.get(12,{}).get('C')) if hoja_cen else None)
    for sitio,sufijo in SITIOS.items():
        principal=f'PROFORMA-{sufijo}'
        hoja=libro.filas(principal)
        if hoja is None:
            resultado['avisos'].append(f'No existe hoja {principal}: {sitio} pendiente.')
            continue
        periodo=_periodo_celda(hoja.get(3,{}).get('C')) or _periodo_celda(hoja.get(15,{}).get('D')) or nombre_periodo
        if not periodo:
            raise ValueError(f'{principal}: no se identificó período en celda C3 ni nombre {archivo!r}')
        if nombre_periodo and nombre_periodo!=periodo:
            raise ValueError(f'{archivo}: el mes del nombre ({nombre_periodo}) difiere del mes de {principal} ({periodo}).')
        if resultado['periodo'] and resultado['periodo']!=periodo:
            raise ValueError('Quilpué y Limache indican períodos de proforma distintos.')
        resultado['periodo']=periodo
        consumo=_mwh(hoja)
        potencia=_mw(hoja)
        cargos=_cargos(hoja,principal,sitio,periodo,consumo,potencia)
        if not cargos:
            raise ValueError(f'{principal}: no se identificaron cargos en la columna F.')
        neto=_valor_fijo(hoja,'NETO',40,50)
        iva=_valor_fijo(hoja,'I.V.A.',40,50)
        total=_valor_fijo(hoja,'TOTAL',40,50)
        sumado=sum(c['Monto CLP'] for c in cargos if c['Monto CLP'] is not None)
        if neto is None or total is None:
            resultado['avisos'].append(f'{principal}: faltan NETO/TOTAL calculados en F. No se inventan.')
        elif abs(sumado-neto)>max(2,abs(neto)*1e-5):
            resultado['avisos'].append(f'{principal}: suma de líneas ({sumado:.0f}) diferente del NETO publicado ({neto:.0f}).')
        if total is not None and iva is not None and neto is not None and abs(total-neto-iva)>2:
            resultado['avisos'].append(f'{principal}: NETO + IVA no concilia con TOTAL (diferencia {total-neto-iva:.2f}).')
        csp_hoja=f'PROFORMA CSP-{sufijo}'
        csp=libro.filas(csp_hoja)
        cargos_csp=[]; neto_csp=None; iva_csp=None; total_csp=None
        if csp:
            csp_periodo=_periodo_celda(csp.get(3,{}).get('C')) or periodo
            if csp_periodo!=periodo:
                resultado['avisos'].append(f'{csp_hoja}: mes {csp_periodo} diferente del principal {periodo}. No se suma CSP.')
            else:
                cargos_csp=_cargos(csp,csp_hoja,sitio,periodo,consumo,potencia,inicio=16,fin=36)
                neto_csp=_valor_fijo(csp,'NETO',36,44)
                iva_csp=_valor_fijo(csp,'I.V.A.',36,44)
                total_csp=_valor_fijo(csp,'TOTAL',36,44)
                if total_csp is None:
                    resultado['avisos'].append(f'{csp_hoja}: CSP sin TOTAL calculado; queda pendiente.')
                elif neto_csp is not None and abs(sum(x['Monto CLP'] or 0 for x in cargos_csp)-neto_csp)>2:
                    resultado['avisos'].append(f'{csp_hoja}: detalle CSP no concilia con su NETO.')
        else:
            resultado['avisos'].append(f'{csp_hoja} no disponible: CSP no se incorpora al total.')
        total_consolidado=total+total_csp if total is not None and total_csp is not None else None
        resultado['suministros'][sitio]={
            'consumo_mwh':consumo,'potencia_mw':potencia,
            'cargos':cargos+cargos_csp,
            'neto_principal':neto,'iva_principal':iva,'total_principal':total,
            'neto_csp':neto_csp,'iva_csp':iva_csp,'total_csp':total_csp,
            'total_proforma':total_consolidado,
            'cargos_detalle':sumado,
            'estado':'Completo' if total_consolidado is not None else 'Parcial',
        }
    if not resultado['suministros']:
        raise ValueError('Este archivo no contiene PROFORMA-Q ni PROFORMA-L.')
    resultado['periodo']=resultado['periodo'] or nombre_periodo
    resultado['mes_reliquidacion']=mes_anterior(resultado['periodo'])
    if (resultado['balance_cen_declarado'] is not None and
            resultado['balance_cen_declarado'] != resultado['mes_reliquidacion']):
        resultado['avisos'].append(
            'CARGOS_CEN!C2/D2 declara balance '+resultado['balance_cen_declarado']+
            ', diferente del mes anterior a la proforma ('+
            resultado['mes_reliquidacion']+'). Revisar antes de conciliar.')
    if (resultado['mes_referencia_peaje_anexo'] is not None and
            resultado['mes_referencia_peaje_anexo'] != resultado['mes_reliquidacion']):
        resultado['avisos'].append(
            'Anexo CARGOS_CEN!C12: referencia de peaje '+
            resultado['mes_referencia_peaje_anexo']+
            ' no coincide con balance M-1 '+resultado['mes_reliquidacion']+
            '. Se conserva M-1 para la imputación solicitada, pero verificar '
            'la vigencia de la fuente del peaje individualizado.')
    return resultado


def _es(x,nd=0):
    if x is None: return 'Pendiente'
    return f'{x:,.{nd}f}'.replace(',','#').replace('.',',').replace('#','.')


def _csv_detalle(cargos):
    buff=StringIO()
    if cargos:
        fields=list(cargos[0])
        wr=csv.DictWriter(buff,fieldnames=fields,delimiter=';')
        wr.writeheader()
        wr.writerows(cargos)
    return ('\ufeff'+buff.getvalue()).encode('utf-8')


# Los selectores año/mes/suministro actualizan solo la pestaña Proforma.
@st.fragment
def mostrar_proformas_enel(archivos, obtener_proforma=None):
    """Pestaña independiente: año, mes y suministro, con GitHub o carga sesión."""
    st.subheader('🧾 Proformas Enel · Quilpué y Limache')
    st.caption('Detalle facturado en la proforma y sus anexos CSP. Fuente: Excel Enel en '
               'Datos/Proformas_enel/AAAA (todos los meses juntos). '
               'No se suman Informe ni RELIQ_Q/L por duplicado.')
    st.code('Datos/Proformas_enel/2026/EFE_VALPO_SEP2026.xlsx',language='text')
    if not archivos:
        st.info('No hay proformas cargadas. Guarda los Excel en GitHub en la carpeta indicada '
                'o súbelos en Datos y ajustes → Archivos y cargas.')
        return
    registros=[]
    for a in archivos:
        mes=periodo_archivo(a.name)
        if mes:
            registros.append((mes,a))
    if not registros:
        st.warning('Los archivos encontrados no tienen un mes reconocible en el nombre. '
                   'Usa EFE_VALPO_SEP2026.xlsx o EFE_VALPO_2026-09.xlsx.')
        return
    periodos=sorted({p for p,_ in registros})
    años=sorted({int(p[:4]) for p in periodos})
    a,b,c=st.columns([1,1,1.4])
    with a:
        year=st.selectbox('Año de proforma',años,index=len(años)-1,key='proforma_anio')
    meses=sorted({int(p[5:7]) for p in periodos if int(p[:4])==year})
    with b:
        mes=st.selectbox('Mes de proforma',meses,index=len(meses)-1,
                         format_func=lambda m: MESES[m-1],key='proforma_mes')
    with c:
        sitio=st.selectbox('Suministro',['Ambos','Quilpué','Limache'],key='proforma_sitio')
    periodo=f'{year:04d}-{mes:02d}'
    candidatos=[r for p,r in registros if p==periodo]
    if not candidatos:
        st.info('No existe proforma para el mes elegido.')
        return
    # En presencia de revisiones, el archivo lo selecciona el usuario explícitamente.
    elegido=st.selectbox('Archivo / versión',range(len(candidatos)),
                         format_func=lambda i:candidatos[i].name,key='proforma_version')
    archivo=candidatos[elegido]
    try:
        resultado=(obtener_proforma(archivo) if obtener_proforma is not None
                   else _parse_cached(archivo.getvalue(), archivo.name))
    except Exception as exc:
        st.error(f'No se pudo leer la proforma {archivo.name}: {type(exc).__name__}: {exc}')
        return
    if resultado['periodo']!=periodo:
        st.error('El período dentro del Excel difiere del seleccionado. Revisa el nombre y la hoja C3.')
        return
    st.info('**Mes de proforma/factura:** '+MESES[mes-1]+' '+str(year)+
            ' · **Origen de PE, SSCC, MT, PD y Liquidación de Peaje:** '+
            MESES[int(resultado['mes_reliquidacion'][5:7])-1]+' '+
            resultado['mes_reliquidacion'][:4]+'.')
    sitios=['Quilpué','Limache'] if sitio=='Ambos' else [sitio]
    completos=[s for s in sitios if s in resultado['suministros']]
    suma=lambda clave: sum((resultado['suministros'][s].get(clave) or 0) for s in completos)
    ok=bool(completos) and len(completos)==len(sitios)
    completitud=ok and all(resultado['suministros'][s]['total_proforma'] is not None for s in sitios)
    if not completos:
        st.warning('No se encontró la hoja de este suministro en la proforma seleccionada.')
        return
    cols=st.columns(4)
    cols[0].metric('Neto principal', '$'+_es(suma('neto_principal')) if ok else 'Parcial')
    cols[1].metric('IVA principal', '$'+_es(suma('iva_principal')) if ok else 'Parcial')
    cols[2].metric('CSP adicional', '$'+_es(suma('total_csp')) if ok and all(resultado['suministros'][s]['total_csp'] is not None for s in sitios) else 'Pendiente')
    cols[3].metric('TOTAL PROFORMA + CSP', '$'+_es(suma('total_proforma')) if completitud else 'Parcial')
    st.caption('El total incluye el TOTAL principal (con IVA indicado) y el TOTAL de las '
               'hojas CSP. El IVA de CSP solo se considera si figura en esa hoja.')
    rows=[];totales=[]
    for s in completos:
        v=resultado['suministros'][s]
        rows+=v['cargos']
        totales.append({'Suministro':s,'Consumo MWh':v['consumo_mwh'],
            'Potencia punta MW':v['potencia_mw'],'Neto principal $':v['neto_principal'],
            'IVA principal $':v['iva_principal'],'Total principal $':v['total_principal'],
            'CSP neto $':v['neto_csp'],'CSP IVA $':v['iva_csp'],
            'CSP total $':v['total_csp'],'TOTAL proforma + CSP $':v['total_proforma'],
            'Estado':v['estado']})
    st.markdown('#### Totales por suministro')
    st.dataframe(totales,hide_index=True,use_container_width=True)
    st.markdown('#### Todos los cargos de la proforma')
    st.caption('**Mes de proforma:** mes de facturación. **Mes de origen:** el balance '
               'anterior para PE, SSCC, MT, PD y Liquidación de Peaje; corriente para '
               'los demás, salvo reliquidación con mes escrito explícitamente. '
               '**Monto CLP:** columna F original; **Equivalente CLP/kWh:** monto F dividido '
               'por los kWh de la proforma, no necesariamente la tarifa oficial contractual. '
               '**Potencia CLP/kW/mes aprox.** se calcula con los MW redondeados que muestra la proforma. '
               'G y H son columnas de referencia del libro y no se suman.')
    st.dataframe(rows,hide_index=True,use_container_width=True,
                 column_config={'Monto CLP':st.column_config.NumberColumn(format='%.0f'),
                    'Equivalente CLP/kWh':st.column_config.NumberColumn(format='%.5f'),
                    'Potencia CLP/kW/mes aprox.':st.column_config.NumberColumn(format='%.3f'),
                    'Referencia G':st.column_config.NumberColumn(format='%.0f'),
                    'Variación H':st.column_config.NumberColumn(format='%.0f')})
    for aviso in resultado['avisos']:
        st.warning(aviso)
    if not completitud:
        st.warning('El total mostrado está incompleto: falta al menos una hoja o total CSP. '
                   'No se reemplazan faltantes por ceros.')
    st.download_button('Descargar detalle proforma CSV',data=_csv_detalle(rows),
                       file_name=f'Proforma_Enel_{periodo}_{sitio}.csv',mime='text/csv',
                       key='proforma_exportar_csv')
    with st.expander('Trazabilidad y criterios de reliquidación'):
        st.write('Archivo: '+archivo.name)
        st.write('Mes de proforma: '+MESES[mes-1]+' '+str(year))
        st.write('Solo se suman importes del mes de proforma indicados en PROFORMA-Q/L y '
                 'PROFORMA CSP-Q/L. La partida «Reliquidación mes de agosto» de septiembre '
                 'ya está dentro del NETO principal; no se agrega otra vez desde RELIQ_Q/L.')
        st.write('**Regla fija:** en la proforma de septiembre, PE, SSCC, MT, PD y '
                 'Liquidación de Peaje pertenecen al balance de agosto; en la proforma '
                 'de octubre corresponden a septiembre. En enero corresponde a diciembre '
                 'del año anterior. Las partidas mensuales de transmisión nacional/zonal '
                 'y el peaje de distribución no se desplazan por esta regla.')
        st.write('Cargos_CEN!C2/D2 (balance informado): '+
                 str(resultado['balance_cen_declarado'] or 'No indicado'))
        st.write('Cargos_CEN!C12 (referencia documental peaje): '+
                 str(resultado['mes_referencia_peaje_anexo'] or 'No indicado'))
        st.write('Cargos de transmisión y CSP también pueden estar dentro del Resumen calculado: '
                 'el TOTAL PROFORMA se muestra separado y nunca se suma al total teórico.')


def _parse_cached_impl(data,nombre):
    return parsear_proforma(data,nombre)

# Cache de tipos simples: el resultado es dict/list/str/float, no clases __main__.
try:
    import streamlit as _st
    _parse_cached=_st.cache_data(show_spinner=False,max_entries=24)(_parse_cached_impl)
except ImportError:
    _parse_cached=_parse_cached_impl
