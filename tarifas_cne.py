"""Consulta histórica de tarifas reguladas CNE para EFE Valparaíso.

Datos versionados en Historico_Tarifas_EFE_2023_2026.json, distribuidos con
esta actualización. Lee nuevas resoluciones SÓLO cuando un histórico validado
se incorpora a GitHub. La búsqueda de publicaciones no altera tarifas.
"""
from __future__ import annotations

from datetime import date
from io import StringIO
from pathlib import Path
import csv
import json
import logging

import pandas as pd
import streamlit as st

BASE = Path(__file__).resolve().parent
FILENAME = "Historico_Tarifas_EFE_2023_2026.json"
MESES = ('Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic')
SITIOS = ("Quilpué", "Limache")


def _data_source():
    """La copia en repo es persistente; la subida desde navegador no lo es."""
    candidates = [BASE / "Datos" / "Tarifas_CNE" / FILENAME,
                  BASE / FILENAME,
                  BASE / "Datos" / FILENAME]
    for path in candidates:
        if path.is_file():
            return path
    return None


def _check(data):
    if not isinstance(data, dict):
        raise ValueError("El archivo JSON debe contener un objeto.")
    for key in ('transmision', 'servicio_publico', 'suministros'):
        if key not in data:
            raise ValueError(f"Falta la sección obligatoria '{key}'.")
    for kind in ('transmision', 'servicio_publico'):
        items = data[kind]
        if not isinstance(items, list):
            raise ValueError(f"{kind} debe ser una lista.")
        ranges = []
        for item in items:
            if not isinstance(item, dict) or not item.get('desde') or not item.get('hasta') or not item.get('rex') or not item.get('url'):
                raise ValueError(f"Registro incompleto en {kind}: faltan vigencia, REX o fuente.")
            if item['desde'] > item['hasta']:
                raise ValueError(f"Vigencia inválida: {item['rex']}.")
            for old_start, old_end, old_rex in ranges:
                if item['desde'] <= old_end and old_start <= item['hasta']:
                    raise ValueError(f"Vigencias superpuestas en {kind}: {old_rex} y {item['rex']}.")
            ranges.append((item['desde'], item['hasta'], item['rex']))
    return data


@st.cache_data(show_spinner=False)
def _load_repo(path, mtime_ns):
    return _check(json.loads(Path(path).read_text(encoding='utf-8-sig')))


def cargar_historial(archivo=None):
    if archivo is not None:
        # Si se carga manualmente, hay que conservar archivo en GitHub para persistir.
        return _check(json.loads(archivo.getvalue().decode('utf-8-sig'))), 'Carga temporal de esta sesión'
    path = _data_source()
    if path is None:
        raise FileNotFoundError(f"No se encontró {FILENAME} en la raíz ni en Datos/Tarifas_CNE/.")
    return _load_repo(str(path), path.stat().st_mtime_ns), str(path.relative_to(BASE))


def vigencia(registros, periodo):
    return next((r for r in registros if str(r['desde']) <= periodo <= str(r['hasta'])), None)


def tarifa_sitio(tx, cp, sitio):
    """Componentes auditables; se preserva None cuando falta respaldo oficial."""
    result = {'Suministro': sitio, 'Sistema': 'C',
              'Tensión': '110 kV' if sitio == 'Quilpué' else '13 kV',
              'Clasificación': 'Individualizado' if sitio == 'Quilpué' else 'No individualizado',
              'Zonal ($/kWh)': (tx.get('qu') if sitio == 'Quilpué' else tx.get('li')) if tx else None,
              'Nacional ($/kWh)': tx.get('nac') if tx else None,
              'Exenciones generales ($/kWh)': tx.get('ex') if tx else None,
              'SSCC infraestructura ($/kWh)': tx.get('sscc') if tx else None,
              'CSP >5000 kWh ($/kWh)': cp.get('total') if cp else None,
              'CSP base ($/kWh)': cp.get('base') if cp else None,
              'FET ($/kWh)': cp.get('fet') if cp else None,
              'Tabla 6 NI ($/kWh)': tx.get('retiro') if tx and sitio == 'Limache' else None,
              'Tabla 5 NI ($/kWh)': tx.get('reduccion') if tx and sitio == 'Limache' else None,
              'REX Transmisión': tx.get('rex') if tx else 'Sin resolución incorporada',
              'REX Servicio Público': cp.get('rex') if cp else 'Sin resolución incorporada',
              'Fuente Transmisión': tx.get('url') if tx else None,
              'Fuente Servicio Público': cp.get('url') if cp else None,
              'Observaciones': tx.get('obs','') if tx else 'No existe resolución de transmisión validada para este período'}
    # Nunca convertir tarifa ausente en cero (ej. enero-junio de 2024).
    mandatory = ['Zonal ($/kWh)', 'Nacional ($/kWh)',
                 'Exenciones generales ($/kWh)', 'SSCC infraestructura ($/kWh)',
                 'CSP >5000 kWh ($/kWh)']
    result['Subtotal base referencial ($/kWh)'] = (
        round(sum(result[c] for c in mandatory), 6)
        if all(result[c] is not None for c in mandatory) else None)
    return result


def tabla_periodos(registro, years, months):
    out=[]
    for y in sorted(years):
        for m in sorted(months):
            period=f'{y:04d}-{m:02d}'
            tx=vigencia(registro['transmision'], period)
            cp=vigencia(registro['servicio_publico'],period)
            for sitio in SITIOS:
                t=tarifa_sitio(tx,cp,sitio)
                t={'Mes':period, **t}
                out.append(t)
    return out


def consumos_balances(balances):
    """Opcional: usa retiros identificados por MERVAL_ELSOL/MERVAL_LIMACHE.

    Los balances están sujetos a los filtros de Datos/Análisis energético;
    no se infieren retiros para meses faltantes.
    """
    known={}
    for balance in balances or []:
        periodo=str(getattr(balance,'label',''))
        if not (len(periodo)==7 and periodo[4]=='-'):
            continue
        for rec in (getattr(balance,'detalle',None) or []):
            clave=str(rec.get('clave','')).upper()
            sitio= 'Quilpué' if clave=='MERVAL_ELSOL' else 'Limache' if clave=='MERVAL_LIMACHE' else None
            if not sitio:
                continue
            try:
                kwh=float(rec.get('kwh') or 0)
            except (TypeError, ValueError):
                continue
            if kwh > 0:
                known[(periodo,sitio)]=known.get((periodo,sitio),0.)+kwh
    return known


def _csv(records):
    if not records:
        return b''
    b=StringIO()
    w=csv.DictWriter(b,fieldnames=list(records[0]),delimiter=';',extrasaction='ignore')
    w.writeheader()
    for rec in records:
        w.writerow({k:str(v).replace('.',',') if isinstance(v,float) else ('' if v is None else v)
                    for k,v in rec.items()})
    return ('\ufeff'+b.getvalue()).encode('utf-8')


def valorizar(filas, consumo_por_periodo, incluir_ni6=False, incluir_ni5=False):
    """No calcula montos sin kWh ni tarifas completas; no incluye otros cargos CEN."""
    salida=[]
    for f in filas:
        kwh=consumo_por_periodo.get((f['Mes'],f['Suministro']))
        base=f['Subtotal base referencial ($/kWh)']
        cond6=f['Tabla 6 NI ($/kWh)'] if incluir_ni6 and f['Suministro']=='Limache' else None
        cond5=f['Tabla 5 NI ($/kWh)'] if incluir_ni5 and f['Suministro']=='Limache' else None
        # Si se activó una condición pero no existe la tarifa oficial, no calcular.
        missing_cond=(incluir_ni6 and f['Suministro']=='Limache' and cond6 is None) or (incluir_ni5 and f['Suministro']=='Limache' and cond5 is None)
        valid= kwh is not None and kwh>0 and base is not None and not missing_cond
        unit=(base + (cond6 or 0.) + (cond5 or 0.)) if valid else None
        salida.append({'Mes':f['Mes'],'Suministro':f['Suministro'],
                       'Consumo (kWh)':kwh,'Tarifa base ($/kWh)':base,
                       'Tabla 6 condicional ($/kWh)':cond6,'Tabla 5 condicional ($/kWh)':cond5,
                       'Tarifa del escenario ($/kWh)':round(unit,6) if unit is not None else None,
                       'Costo regulado referencial ($)':round(kwh*unit,2) if valid else None,
                       'REX Transmisión':f['REX Transmisión'],'REX Servicio Público':f['REX Servicio Público'],
                       'Estado':'Calculado (referencial)' if valid else 'Pendiente: faltan kWh o tarifa/condición'})
    return salida


def mostrar_tarifas_cne(balances=None):
    st.subheader('📚 Tarifas reguladas e histórico CNE · 2023 en adelante')
    st.caption('Tarifas independientes de archivos CEN; valor de kWh aplicable por suministro, '
               'vigencia y resolución exenta. Quilpué: individualizado, 110 kV. Limache: no individualizado, 13 kV. Sistema C.')
    upload=st.file_uploader('Actualizar catálogo validado (JSON; carga de prueba temporal)',type=['json'],key='tarifas_cne_json_subido')
    try:
        catalogo, fuente=cargar_historial(upload)
    except Exception as exc:
        st.error(f'No se pudo leer el histórico regulado: {exc}')
        st.info(f'Confirma que {FILENAME} esté en la carpeta principal del repositorio.')
        return
    st.caption('Catálogo cargado: '+fuente+' · versión '+str(catalogo.get('version','no indicada')))
    years=sorted({int(r['desde'][:4]) for r in catalogo['transmision'] if r.get('desde')}
                 | {int(r['hasta'][:4]) for r in catalogo['transmision'] if r.get('hasta')})
    today=date.today()
    defaults=[today.year] if today.year in years else years[-1:]
    a,b=st.columns(2)
    with a:
        selected_years=st.multiselect('Años (independientes de Datos)', years, default=defaults, key='tarifas_cne_years')
    with b:
        selected_months=st.multiselect('Meses',list(range(1,13)),default=list(range(1,13)),
                                       format_func=lambda m:MESES[m-1],key='tarifas_cne_months')
    if not selected_years or not selected_months:
        st.info('Selecciona uno o más años y meses.')
        return
    rows=tabla_periodos(catalogo, selected_years,selected_months)
    st.markdown('#### Tarifas por mes y suministro')
    visible=['Mes','Suministro','Tensión','Clasificación','Zonal ($/kWh)','Nacional ($/kWh)',
             'Exenciones generales ($/kWh)','SSCC infraestructura ($/kWh)',
             'CSP >5000 kWh ($/kWh)','Tabla 6 NI ($/kWh)','Tabla 5 NI ($/kWh)',
             'Subtotal base referencial ($/kWh)','REX Transmisión','REX Servicio Público']
    st.dataframe([{k:rec.get(k) for k in visible} for rec in rows],hide_index=True,use_container_width=True)
    if any(rec['Subtotal base referencial ($/kWh)'] is None for rec in rows):
        st.warning('Hay períodos con tarifa no publicada o no identificada. Un guion en la resolución NO es cero; no se calculará subtotal.')
    st.caption('El subtotal incluye nacional + zonal + exenciones generales + infraestructura SSCC + servicio público '
               '(tramo >5.000 kWh). Es un escenario técnico, no necesariamente un cargo adicional a facturar; '
               'revisa reglas contractuales, descuentos y conceptos ya cobrados.')
    st.markdown('#### Cálculo referencial usando el consumo de cada suministro')
    mapeo=consumos_balances(balances)
    st.caption('Se proponen los kWh de balances CEN que estén cargados. Puedes corregirlos o ingresar consumos '
               'de la proyección operativa: la tabla no transforma automáticamente SEAT/SAF en retiro contractual.')
    consume=[{'Mes':r['Mes'],'Suministro':r['Suministro'],
              'Consumo kWh':mapeo.get((r['Mes'],r['Suministro']))} for r in rows]
    df_consumo = pd.DataFrame(consume)
    df_consumo['Consumo kWh'] = pd.to_numeric(df_consumo['Consumo kWh'], errors='coerce')
    editado=st.data_editor(df_consumo,use_container_width=True,hide_index=True,
             disabled=['Mes','Suministro'],num_rows='fixed',key='tarifas_cne_consumos')
    consumo={}
    for r in editado.to_dict('records'):
        try:
            val=float(r.get('Consumo kWh'))
            if pd.notna(val) and val>0:
                consumo[(r['Mes'],r['Suministro'])]=val
        except (TypeError, ValueError):
            pass
    with st.expander('Escenarios especiales para Limache (Tablas 5 y 6)',expanded=False):
        st.warning('No están incluidos por defecto: su aplicación requiere validar la clasificación '
                   'y excepciones de los artículos transitorios, especialmente la Tabla 5.')
        t6=st.checkbox('Incluir Tabla 6 NI (escenario sujeto a validación)',value=False,key='tarifas_cne_ni6')
        t5=st.checkbox('Incluir Tabla 5 NI (escenario sujeto a validación)',value=False,key='tarifas_cne_ni5')
    calculado=valorizar(rows,consumo,t6,t5)
    st.dataframe(calculado,hide_index=True,use_container_width=True)
    totals={s:sum(x['Costo regulado referencial ($)'] or 0 for x in calculado
                  if x['Suministro']==s and x['Costo regulado referencial ($)'] is not None)
            for s in SITIOS}
    valid={s:any(x['Suministro']==s and x['Costo regulado referencial ($)'] is not None for x in calculado) for s in SITIOS}
    c1,c2=st.columns(2)
    for c,s in [(c1,'Quilpué'),(c2,'Limache')]:
        c.metric('Subtotal referencial · '+s,
                 ('${:,.0f}'.format(totals[s]).replace(',','.')) if valid[s] else 'Sin cálculo')
    st.download_button('⬇ Exportar histórico de tarifas CSV',_csv([{k:v for k,v in row.items()} for row in rows]),
                       file_name='EFE_Historico_Tarifas_CNE.csv',mime='text/csv',key='tarifas_cne_download_hist')
    st.download_button('⬇ Exportar valorización CSV',_csv(calculado),file_name='EFE_Valorizacion_Tarifas_CNE.csv',
                       mime='text/csv',key='tarifas_cne_download_val')
    with st.expander('Resoluciones y documentos fuente'):
        for group,items in [('Transmisión',catalogo['transmision']),('Servicio Público',catalogo['servicio_publico'])]:
            st.markdown('**'+group+'**')
            for item in items:
                st.markdown(f"- REX {item['rex']} · {item['desde']} a {item['hasta']} · [Documento oficial]({item['url']})")
    st.markdown('#### Actualizaciones por nuevas resoluciones')
    st.info('La revisión de nuevas publicaciones se realiza en la pestaña **Actualizaciones CEN** y mediante '
            'el seguimiento semanal. Para incorporar una tarifa NUEVA o rectificación, primero se debe validar '
            'el texto oficial; después se agrega un registro con la vigencia correspondiente al JSON de GitHub. '
            'No se extrapolan ni reemplazan automáticamente las tarifas existentes.')
    st.caption('Exenciones CI (PagoExenCI) se mantiene SOLO en Quilpué; SSCC operacionales/BESS del CEN '
               'son independientes de los  $/kWh por infraestructura SSCC. No se suman aquí a los subtotales.')
