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
import math

import pandas as pd
import streamlit as st

BASE = Path(__file__).resolve().parent
FILENAME = "Historico_Tarifas_EFE_2023_2026.json"
MESES = ('Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic')
SITIOS = ("Quilpué", "Limache")
VERSION_ENLACE_CNE = "V5-2026-10-27"
VERSION_FILTROS_CNE = "V6-2026-10-27"
# V5 de transferencia sigue siendo compatible con proyeccion_energia.py V5.


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


def _consumos_compartidos_proyeccion():
    """Resultados del motor operativo, ya calculados en Proyección de energía."""
    state = st.session_state.get('efe_proyeccion_tarifas') or {}
    return state.get('consumos', {}) if isinstance(state, dict) else {}


@st.cache_data(ttl=1800, show_spinner=False, max_entries=36)
def _consumos_automaticos_git_ano(anio, meses):
    """Respaldo independiente con el MISMO cálculo de Proyección de energía.

    No se utiliza el motor distinto de Costo de energía; preserva fechas, UMR,
    histórico y perfiles equivalentes al módulo de proyección principal.
    """
    from calendar import monthrange
    from costo_energia import _datos_operacionales
    from proyeccion_energia import _consumos_mismo_modelo, catalogo
    anio = int(anio)
    meses = tuple(sorted(set(int(m) for m in meses)))
    if not meses:
        return {}, []
    try:
        disponibles = {m for item in catalogo('seat')
                        for y, m in [item['period']] if y == anio}
    except Exception as exc:
        return {}, [f'No se pudo consultar el catálogo SEAT del repositorio: {exc}']
    if not disponibles:
        return {}, [f'No hay Excel SEAT disponibles del año {anio} en Datos-Operacionales.']
    # Se incluyen meses con mediciones hasta el último mes seleccionado
    # para disponer del mismo perfil de entrenamiento que Proyección.
    meses_fuente = sorted({m for m in disponibles if m <= max(meses)})
    if not meses_fuente:
        return {}, [f'No hay datos de entrenamiento SEAT anteriores a los meses elegidos de {anio}.']
    (seat, saf, odom, kms), ns, nu, avisos = _datos_operacionales(
        [anio], meses_fuente, True, None, None)
    inicio = date(anio, min(meses), 1)
    ultimo_mes = max(meses)
    fin = date(anio, ultimo_mes, monthrange(anio, ultimo_mes)[1])
    mapa, incidencias = _consumos_mismo_modelo(
        seat, saf, odom, kms, [anio], meses_fuente, inicio, fin)
    mapa = {clave: rec for clave, rec in mapa.items()
            if int(clave[0][5:7]) in meses}
    mapa = {clave: dict(rec, origen='Datos-Operacionales (mismo modelo Proyección de energía)')
            for clave, rec in mapa.items()}
    return mapa, list(avisos) + list(incidencias)


def _consumos_excel_exportados(archivos):
    """Respaldo: exportaciones xlsx de Proyección de energía; no requiere red.

    Acepta los archivos originales del botón Descargar proyección:
    proyeccion_energia_seat.xlsx y proyeccion_energia_limache.xlsx.
    """
    consumos, advertencias = {}, []
    for archivo in archivos or []:
        nombre = str(archivo.name)
        name = nombre.casefold()
        sitio = 'Limache' if 'limache' in name else ('Quilpué' if 'seat' in name or 'quilp' in name else None)
        if not sitio:
            advertencias.append(nombre+': no indica seat / quilpue / limache; renómbralo para reconocer el suministro.')
            continue
        try:
            from io import BytesIO
            tabla = pd.read_excel(BytesIO(archivo.getvalue()), sheet_name='Real y proyectado')
            obligatorios = {'Fecha','kWh','Origen'}
            if not obligatorios.issubset(set(tabla.columns)):
                raise ValueError('faltan columnas Fecha, kWh u Origen de la exportación original')
            tabla['Fecha'] = pd.to_datetime(tabla['Fecha'], errors='coerce')
            tabla['kWh'] = pd.to_numeric(tabla['kWh'], errors='coerce')
            tabla = tabla.dropna(subset=['Fecha','kWh'])
            tabla = tabla[tabla['kWh'] >= 0].copy()
            tabla['Mes'] = tabla['Fecha'].dt.strftime('%Y-%m')
            for mes, grupo in tabla.groupby('Mes'):
                real = float(grupo.loc[grupo['Origen']=='Real', 'kWh'].sum())
                proy = float(grupo.loc[grupo['Origen']=='Proyectado','kWh'].sum())
                clave = (mes,sitio)
                if clave in consumos:
                    advertencias.append(nombre+': mes '+mes+' duplicado; se conservó el último archivo.')
                consumos[clave] = {'kwh':real+proy, 'real_kwh':real,
                    'proyectado_kwh':proy,'dias':int(grupo['Fecha'].dt.normalize().nunique()),
                    'origen':'Excel exportado desde Proyección de energía ('+nombre+')'}
        except Exception as exc:
            advertencias.append(nombre+': '+str(exc))
    return consumos, advertencias


def _consumos_csv_exportados(archivos):
    """Importa el CSV único generado desde la misma proyección visible.

    No infiere ni rellena con cero períodos que no están en el archivo.
    """
    result, warnings = {}, []
    for archivo in archivos or []:
        if not str(archivo.name).casefold().endswith('.csv'):
            continue
        try:
            from io import BytesIO
            df = pd.read_csv(BytesIO(archivo.getvalue()), sep=';', encoding='utf-8-sig',
                             dtype={'Mes':str,'Suministro':str})
            needed={'Mes','Suministro','Consumo kWh','Real kWh','Proyectado kWh'}
            if not needed.issubset(set(df.columns)):
                raise ValueError('CSV sin columnas del botón «Exportar consumos para Tarifas CNE»')
            for _, rec in df.iterrows():
                mes = str(rec['Mes']).strip()
                site = str(rec['Suministro']).strip()
                if len(mes) != 7 or mes[4] != '-' or site not in SITIOS:
                    warnings.append(f'{archivo.name}: período o suministro no reconocido ({mes}; {site})')
                    continue
                val = float(rec['Consumo kWh'])
                real = float(rec['Real kWh'])
                proy = float(rec['Proyectado kWh'])
                if not all(math.isfinite(v) and v>=0 for v in (val, real, proy)):
                    continue
                if abs(val-(real+proy)) > max(0.1, val * 0.00001):
                    raise ValueError(f'{mes} {site}: total kWh no coincide con real + proyectado')
                dias=rec.get('Dias')
                dias=int(dias) if pd.notna(dias) and str(dias).strip() else None
                result[(mes,site)] = {'kwh':val,'real_kwh':real,
                    'proyectado_kwh':proy, 'dias':dias,
                    'origen':'Exportación de Proyección de energía ('+str(archivo.name)+')'}
        except Exception as exc:
            warnings.append(f'{archivo.name}: error de lectura: {exc}')
    return result, warnings


def consumos_para_tarifas(filas, balances, fuente='Proyección de energía', automatico=True, excel=None):
    """Un suministro/mes = un único kWh sin duplicar Balance y Proyección."""
    compartido = _consumos_compartidos_proyeccion()
    excel = excel or {}
    balance = consumos_balances(balances)
    requiere = {(r['Mes'], r['Suministro']) for r in filas}
    git = {}
    advertencias = []
    if automatico and fuente == 'Proyección de energía':
        # Solo consulta años de la selección para los que falta alguna proyección.
        faltantes = {clave for clave in requiere if clave not in compartido and clave not in excel}
        years = sorted({int(periodo[:4]) for periodo, sitio in faltantes})
        for y in years:
            meses_faltantes = tuple(sorted({int(periodo[5:7]) for periodo, sitio in faltantes
                                           if int(periodo[:4]) == y}))
            try:
                valores, avisos = _consumos_automaticos_git_ano(y, meses_faltantes)
                git.update(valores)
                if avisos:
                    advertencias.extend(f'{y}: {x}' for x in avisos[:5])
            except Exception as exc:
                advertencias.append(f'{y}: no se pudo leer Datos-Operacionales: {exc}')
    output = {}
    for clave in sorted(requiere):
        pp = compartido.get(clave)
        xx = excel.get(clave)
        gg = git.get(clave)
        bb = balance.get(clave)
        if fuente == 'Balance CEN':
            if bb is not None:
                output[clave] = {'kwh':float(bb),'origen':'Balance CEN (retiro)', 'dias':None,
                                 'real_kwh':float(bb), 'proyectado_kwh':0.}
                continue
        if xx is not None and xx.get('kwh') is not None:
            output[clave] = xx
        elif pp is not None and pp.get('kwh') is not None:
            output[clave] = pp
        elif gg is not None and gg.get('kwh') is not None:
            output[clave] = gg
        elif bb is not None:
            output[clave] = {'kwh':float(bb),'origen':'Balance CEN (retiro)', 'dias':None,
                             'real_kwh':float(bb),'proyectado_kwh':0.}
    return output, advertencias


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
    st.caption('Conexión a Proyección de energía · ' + VERSION_ENLACE_CNE + ' · Filtros mensuales ' + VERSION_FILTROS_CNE)
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
    with st.expander('Importar consumos de Proyección (respaldo sin conexión)', expanded=False):
        st.caption('Recomendado: en Proyección de energía usa «Exportar consumos de '
                   'Quilpué y Limache para Tarifas CNE (CSV)» y sube aquí ese único archivo. '
                   'También admite los Excel individuales descargados desde la proyección.')
        subida_proyeccion=st.file_uploader('CSV de consumos o Excel de Proyección de energía',
                    type=['csv','xlsx'], accept_multiple_files=True, key='tarifas_cne_excel_proyeccion')
    excel_proy, avisos_excel = _consumos_excel_exportados(
        [f for f in (subida_proyeccion or []) if f.name.casefold().endswith('.xlsx')])
    csv_proy, avisos_csv = _consumos_csv_exportados(subida_proyeccion)
    excel_proy.update(csv_proy)
    avisos_excel.extend(avisos_csv)
    years = sorted({y for r in catalogo['transmision']
                    if r.get('desde') and r.get('hasta')
                    for y in range(int(r['desde'][:4]), int(r['hasta'][:4]) + 1)})
    if not years:
        st.error('No hay años con resoluciones validadas en el catálogo.')
        return
    today = date.today()
    compartidos = _consumos_compartidos_proyeccion()
    periodos_activos = {**compartidos, **excel_proy}
    periodos_proy = sorted({mes for mes, _sitio in periodos_activos if int(mes[:4]) in years})
    mes_sugerido = periodos_proy[-1] if periodos_proy else f'{today.year:04d}-{today.month:02d}'
    try:
        anio_sugerido, mes_sugerido_num = (int(x) for x in mes_sugerido.split('-'))
    except (ValueError, TypeError):
        anio_sugerido, mes_sugerido_num = today.year, today.month
    if anio_sugerido not in years:
        anio_sugerido, mes_sugerido_num = years[-1], today.month

    st.markdown('#### 📅 Filtrar período de tarifas y consumo')
    st.caption('El período se selecciona **independientemente** de la proyección. '
               'Si eliges un solo mes, se mostrarán únicamente Quilpué y Limache de ese mes. '
               'La sincronización de consumo no bloquea estos filtros.')
    modo = st.radio(
        'Meses que deseas visualizar',
        ['Un mes', 'Varios meses', 'Períodos de Proyección'],
        horizontal=True, key='tarifas_cne_modo_periodos',
    )
    if modo == 'Un mes':
        ca, cm = st.columns(2)
        with ca:
            anio_unico = st.selectbox('Año', years, index=years.index(anio_sugerido),
                                     key='tarifas_cne_anio_unico')
        with cm:
            mes_unico = st.selectbox('Mes', list(range(1, 13)),
                                    index=mes_sugerido_num - 1,
                                    format_func=lambda m: MESES[m - 1],
                                    key='tarifas_cne_mes_unico')
        selected_years, selected_months = [anio_unico], [mes_unico]
    elif modo == 'Varios meses':
        ca, cm = st.columns(2)
        with ca:
            selected_years = st.multiselect('Años', years, default=[anio_sugerido],
                                            key='tarifas_cne_anios_multiples')
        with cm:
            todos_los_meses = st.checkbox('Todos los meses', value=False,
                                         key='tarifas_cne_todos_los_meses')
            if todos_los_meses:
                selected_months = list(range(1, 13))
                st.caption('Meses incluidos: enero a diciembre')
            else:
                selected_months = st.multiselect('Meses', list(range(1, 13)),
                                                 default=[mes_sugerido_num],
                                                 format_func=lambda m: MESES[m - 1],
                                                 key='tarifas_cne_meses_multiples')
    else:
        # No se usa el producto cartesiano años x meses para no inventar períodos.
        if not periodos_proy:
            st.warning('No hay meses de Proyección disponibles en esta sesión. '
                       'Elige «Un mes» o «Varios meses», o importa el CSV de consumos.')
            return
        selected_years = sorted({int(p[:4]) for p in periodos_proy})
        selected_months = sorted({int(p[5:7]) for p in periodos_proy})
        st.info('Meses presentes en Proyección: ' + ', '.join(periodos_proy))

    if not selected_years or not selected_months:
        st.info('Selecciona al menos un año y un mes.')
        return
    rows = tabla_periodos(catalogo, selected_years, selected_months)
    if modo == 'Períodos de Proyección':
        periodos_validos = set(periodos_proy)
        rows = [r for r in rows if r['Mes'] in periodos_validos]
    st.caption(f'Meses visibles: {len({r["Mes"] for r in rows})} · '
               f'Suministros: Quilpué y Limache · Filas: {len(rows)}')
    if not rows:
        st.warning('No existen meses coincidentes con el histórico de tarifas.')
        return

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
    import calendar
    import hashlib
    sincronizar = st.checkbox(
        'Usar automáticamente los kWh de Proyección de energía (sin cambiar los filtros)',
        value=True, key='tarifas_cne_sincronizar_periodos',
    )
    if sincronizar:
        preferencia = 'Proyección de energía'
        st.caption('Se priorizan los consumos operacionales del mes seleccionado. '
                   'Si no están compartidos, se buscarán en Datos-Operacionales. '
                   'El año y el mes elegidos NO cambian.')
    else:
        preferencia = st.radio('Fuente prioritaria del consumo',
                              ['Proyección de energía', 'Balance CEN'],
                              horizontal=True, key='tarifas_cne_fuente_consumo')
    completar=st.checkbox('Completar automáticamente meses no visibles usando SEAT/UMR de Datos-Operacionales',
                          value=True,key='tarifas_cne_completar_operacional',
                          disabled=preferencia!='Proyección de energía')
    with st.spinner('Obteniendo kWh desde Proyección de energía y fuentes operacionales…'):
        consumos_cargados, avisos_consumo=consumos_para_tarifas(
            rows, balances, preferencia, completar and preferencia=='Proyección de energía', excel=excel_proy)
    avisos_consumo = list(avisos_excel) + list(avisos_consumo)
    compartidos = _consumos_compartidos_proyeccion()
    claves_compartidas = {(r['Mes'], r['Suministro']) for r in rows if
                          (r['Mes'], r['Suministro']) in compartidos}
    with st.expander('🔎 Verificación del enlace y versión instalada', expanded=not bool(consumos_cargados)):
        st.write('Versión instalada en Tarifas CNE:', VERSION_ENLACE_CNE)
        try:
            from proyeccion_energia import VERSION_ENLACE_CNE as version_proy
        except ImportError:
            version_proy = 'Versión anterior: falta actualizar proyeccion_energia.py'
        st.write('Versión instalada en Proyección de energía:', version_proy)
        if version_proy != VERSION_ENLACE_CNE:
            st.error('Versiones descoordinadas. Actualiza AMBOS módulos en la raíz de GitHub y reinicia Streamlit.')
        st.write('Meses/suministros publicados por Proyección:', len(compartidos))
        st.write('Meses/suministros importados desde archivo:', len(excel_proy))
        st.write('Meses/suministros que Tarifas CNE recibió:', len(consumos_cargados))
        muestra = []
        for (mes,sitio), r in sorted({**compartidos, **excel_proy}.items()):
            muestra.append({'Mes':mes,'Suministro':sitio,'kWh':r.get('kwh'),
                            'Real':r.get('real_kwh'),'Proyectado':r.get('proyectado_kwh'),
                            'Fuente':r.get('origen')})
        if muestra:
            st.dataframe(muestra, hide_index=True, use_container_width=True)
        else:
            st.warning('No existe un conjunto de consumos compartidos. Usa el CSV de respaldo o revisa en Proyección los mensajes de SEAT/UMR.')
    if excel_proy:
        st.info(f'Se importaron {len(excel_proy)} registros mensuales desde Excel de Proyección de energía.')
    if claves_compartidas:
        st.success(f'Conectado con Proyección de energía: {len(claves_compartidas)} '
                   'registros mensuales de Quilpué/Limache encontrados.')
    else:
        st.info('El mes seleccionado no está entre los consumos compartidos de la sesión. '
                'Se intentará obtener automáticamente desde Datos-Operacionales o podrás importar el CSV.')
    if avisos_consumo or st.session_state.get('efe_proyeccion_tarifas', {}).get('errores'):
        with st.expander('Diagnóstico de consumos y archivos SEAT/UMR', expanded=not bool(consumos_cargados)):
            estado = st.session_state.get('efe_proyeccion_tarifas') or {}
            st.write('Período de Proyección:', estado.get('desde','no disponible'),
                     '→', estado.get('hasta','no disponible'))
            st.write('Registros compartidos:', len(compartidos),
                     '| Registros encontrados en Tarifas CNE:', len(consumos_cargados))
            for aviso in list(estado.get('errores') or []) + avisos_consumo:
                st.write('• ' + str(aviso))
    if consumos_cargados:
        st.caption('Consumos disponibles: '+', '.join(
            f"{m} {sitio}: {rec.get('kwh',0):,.1f} kWh ({rec.get('origen','')})"
            for (m,sitio),rec in sorted(consumos_cargados.items())[:8]))
    if not consumos_cargados:
        st.error('No se obtuvieron kWh para los meses seleccionados. Revisa en Proyección de energía '
                 'que exista un total calculado y que los filtros de fecha coincidan. '
                 'Los kWh del editor pueden introducirse manualmente si es necesario.')
    st.caption('Prioridad automática: Proyección de energía de esta sesión → datos SEAT/UMR de GitHub '
               'con el mismo modelo → retiros del Balance CEN, si existen. '
               'Nunca se suman entre sí. Puedes corregir manualmente los kWh.')
    consume=[]
    parciales=[]
    for r in rows:
        key=(r['Mes'],r['Suministro'])
        c=consumos_cargados.get(key, {})
        kwh=c.get('kwh')
        ndias=c.get('dias')
        y,m=map(int,r['Mes'].split('-'))
        total_dias=calendar.monthrange(y,m)[1]
        cobertura = (f'{ndias}/{total_dias}' if ndias is not None else 'Balance / no informado')
        if kwh is not None and ndias is not None and ndias < total_dias:
            parciales.append(f"{r['Mes']} · {r['Suministro']} ({ndias}/{total_dias} días)")
        consume.append({'Mes':r['Mes'],'Suministro':r['Suministro'],
              'Consumo kWh':kwh,
              'Origen kWh':c.get('origen','Sin datos: ingresar manualmente'),
              'Real kWh':c.get('real_kwh'),
              'Proyectado kWh':c.get('proyectado_kwh'),
              'Cobertura':cobertura})
    if parciales:
        st.warning('Hay consumos que solo cubren una parte del mes. Se valorizará la energía disponible, '
                   'NO se extrapolará automáticamente al mes completo: '+', '.join(parciales[:8])+'.')
    df_consumo = pd.DataFrame(consume)
    for campo in ('Consumo kWh','Real kWh','Proyectado kWh'):
        df_consumo[campo] = pd.to_numeric(df_consumo[campo], errors='coerce')
    # Reiniciar el estado del editor solo cuando cambian los valores de entrada;
    # impide que una edición de períodos anteriores se propague a otra consulta.
    huella = hashlib.sha256(df_consumo.to_json(orient='records').encode('utf-8')).hexdigest()[:12]
    editado=st.data_editor(df_consumo,use_container_width=True,hide_index=True,
             disabled=['Mes','Suministro','Origen kWh','Real kWh','Proyectado kWh','Cobertura'],
             num_rows='fixed',key='tarifas_cne_consumos_'+huella)
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
