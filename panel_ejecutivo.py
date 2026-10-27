"""Panel ejecutivo compacto para EFE Valparaíso.

No recalcula ni mezcla cargos oficiales: utiliza la proyección ya publicada,
las tarifas CNE versionadas y el costo contractual que calculó costo_energia.
Los importes ausentes son None, nunca ceros. No suma SSCC, PSUF, PE ni CI
porque dichos rubros tienen mecanismos y fuentes independientes.
"""
from __future__ import annotations

import calendar
import math
from datetime import date
from io import StringIO
import csv

import pandas as pd
import streamlit as st

VERSION_PANEL = "Panel ejecutivo v1.0 · 2026-10-27"
SITIOS = ("Quilpué", "Limache")
MESES = ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")


def _num(value):
    try:
        f = float(value)
        return f if math.isfinite(f) else None
    except (TypeError, ValueError):
        return None


def _es(num, decimals=0):
    num = _num(num)
    if num is None:
        return "—"
    return f"{num:,.{decimals}f}".replace(",", "#").replace(".", ",").replace("#", ".")


def _costo_contractual(mes, sitio, rows):
    """Solo toma costos calculados por la pestaña contractual, sin estimarlos."""
    year, month = (int(p) for p in mes.split('-'))
    for r in rows or []:
        if (int(r.get('Año') or 0), int(r.get('Mes') or 0), r.get('Suministro')) == (year, month, sitio):
            val = _num(r.get('Costo CLP'))
            return val, r.get('Estado', 'Sin estado')
    return None, 'Pendiente en Costo de energía'


def _obtener_consu(mes, balances):
    from tarifas_cne import _consumos_compartidos_proyeccion, _consumos_automaticos_git_ano, consumos_balances
    yy, mm = (int(p) for p in mes.split('-'))
    compartidos = _consumos_compartidos_proyeccion()
    encontrados = {(periodo,sitio):dict(rec)
                   for (periodo,sitio),rec in compartidos.items()
                   if periodo == mes and sitio in SITIOS and isinstance(rec,dict)}
    avisos = []
    # Los valores compartidos de la propia proyección tienen máxima prioridad.
    pendientes = [sitio for sitio in SITIOS if (mes, sitio) not in encontrados]
    if pendientes:
        try:
            remoto, incidencias = _consumos_automaticos_git_ano(yy, (mm,))
            avisos.extend(incidencias)
            for sitio in pendientes:
                rec = remoto.get((mes,sitio))
                if rec and _num(rec.get('kwh')) is not None:
                    encontrados[(mes,sitio)] = dict(rec)
        except Exception as exc:
            avisos.append(f'No se pudo consultar el modelo operativo: {type(exc).__name__}: {exc}')
    if any((mes,s) not in encontrados for s in SITIOS):
        for (periodo, sitio), value in consumos_balances(balances).items():
            if periodo == mes and (mes,sitio) not in encontrados:
                encontrados[(mes,sitio)] = {
                    'kwh':value, 'real_kwh':value, 'proyectado_kwh':0,
                    'dias':None, 'origen':'Balance CEN (retiro; no proyección)'
                }
    return encontrados, avisos


def resumen_periodo(year, month, sitios, balances, catalogo, costo_rows):
    """Función sin renderizado; útil para pruebas deterministas."""
    from tarifas_cne import tabla_periodos
    mes = f'{year:04d}-{month:02d}'
    tarifas = {(r['Mes'],r['Suministro']):r for r in tabla_periodos(catalogo,[year],[month])}
    energia, avisos = _obtener_consu(mes, balances)
    result=[]
    for sitio in sitios:
        datum = energia.get((mes,sitio), {})
        kwh = _num(datum.get('kwh'))
        real = _num(datum.get('real_kwh'))
        proy = _num(datum.get('proyectado_kwh'))
        tarif = tarifas.get((mes,sitio), {})
        unit = _num(tarif.get('Subtotal base referencial ($/kWh)'))
        cne_total = kwh * unit if (kwh is not None and unit is not None) else None
        energia_clp, estado_costo = _costo_contractual(mes, sitio, costo_rows)
        # No se construye un total general: faltan pagos PE/SSCC/PSUF/CI y pueden
        # coexistir conceptos de transmisión ya facturados en otra partida.
        result.append({
            'Mes':mes, 'Suministro':sitio,
            'Consumo real (kWh)':real,
            'Consumo proyectado (kWh)':proy,
            'Total energía (kWh)':kwh,
            'Precio regulado base ($/kWh)':unit,
            'Cargos CNE referenciales ($)':cne_total,
            'Costo energía contractual ($)':energia_clp,
            'Estado costo energía':estado_costo,
            'REX transmisión':tarif.get('REX Transmisión','Sin resolución'),
            'REX servicio público':tarif.get('REX Servicio Público','Sin resolución'),
            'Fuente consumo':datum.get('origen','Sin consumos disponibles'),
            'Cobertura (días)':datum.get('dias'),
            'Días del mes':calendar.monthrange(year,month)[1],
        })
    return result, avisos


def _csv(records):
    out=StringIO()
    if not records:
        return b''
    writer=csv.DictWriter(out,fieldnames=list(records[0]),delimiter=';')
    writer.writeheader()
    for r in records:
        writer.writerow({k:('' if v is None else
                             str(v).replace('.',',') if isinstance(v,float) else str(v))
                        for k,v in r.items()})
    return ('\ufeff'+out.getvalue()).encode('utf-8')


def mostrar_panel_ejecutivo(balances=None, archivos_repositorio=None):
    st.subheader('Resumen mensual · Quilpué y Limache')
    st.caption('Selecciona un período. Aquí solo se muestran consumos, tarifas y costos trazables; '
               'los cálculos técnicos completos siguen en las demás secciones.')
    from tarifas_cne import cargar_historial
    try:
        catalogo, fuente_hist = cargar_historial()
    except Exception as exc:
        st.error(f'No se puede abrir el historial CNE validado: {exc}')
        st.info('Sube Historico_Tarifas_EFE_2023_2026.json a la raíz de GitHub.')
        return
    years=sorted({yy for r in catalogo['transmision']
                  for yy in range(int(r['desde'][:4]), int(r['hasta'][:4])+1)})
    if not years:
        st.warning('No hay resoluciones en el historial CNE.')
        return
    hoy=date.today()
    shared=st.session_state.get('efe_proyeccion_tarifas',{}) or {}
    periodos=sorted({k[0] for k in (shared.get('consumos') or {})
                    if isinstance(k,tuple) and len(k)==2 and isinstance(k[0],str)
                    and len(k[0])==7})
    pred=periodos[-1] if periodos else f'{hoy.year:04d}-{hoy.month:02d}'
    suggested_y=int(pred[:4]) if int(pred[:4]) in years else years[-1]
    suggested_m=int(pred[5:7])
    a,b,c=st.columns([1,1,1.4])
    with a:
        year=st.selectbox('Año',years,index=years.index(suggested_y),key='panel_periodo_ano')
    with b:
        month=st.selectbox('Mes',list(range(1,13)),index=suggested_m-1,
                           format_func=lambda x: MESES[x-1],key='panel_periodo_mes')
    with c:
        option=st.selectbox('Suministro',['Ambos',*SITIOS],key='panel_suministro')
    selected=SITIOS if option=='Ambos' else (option,)
    costo_rows=st.session_state.get('efe_costo_energia_resumen', [])
    with st.spinner('Consultando el consumo y la tarifa del período seleccionado…'):
        info, avisos=resumen_periodo(year,month,selected,balances,catalogo,costo_rows)

    # Panel de indicadores. Ausencias se muestran como raya y nunca como 0.
    completos_kwh=all(r['Total energía (kWh)'] is not None for r in info)
    completos_tar=all(r['Cargos CNE referenciales ($)'] is not None for r in info)
    completos_contrato=all(r['Costo energía contractual ($)'] is not None for r in info)
    total_kwh=sum(r['Total energía (kWh)'] for r in info) if completos_kwh else None
    total_real=sum((r['Consumo real (kWh)'] or 0) for r in info) if completos_kwh else None
    total_proy=sum((r['Consumo proyectado (kWh)'] or 0) for r in info) if completos_kwh else None
    cargos=sum(r['Cargos CNE referenciales ($)'] for r in info) if completos_tar else None
    energia_clp=sum(r['Costo energía contractual ($)'] for r in info) if completos_contrato else None
    k1,k2,k3,k4=st.columns(4)
    k1.metric('Consumo del período',(_es(total_kwh)+' kWh') if total_kwh is not None else 'Pendiente')
    k2.metric('Energía proyectada',(_es(total_proy)+' kWh') if total_proy is not None else 'Pendiente')
    k3.metric('Costo de energía (contrato)',('$'+_es(energia_clp)) if energia_clp is not None else 'Pendiente')
    k4.metric('Cargos CNE estimados',('$'+_es(cargos)) if cargos is not None else 'Pendiente')

    # Precios y consumos en una tabla compacta con exactamente los meses elegidos.
    st.markdown('#### Comparación de suministros')
    visibles=[]
    for r in info:
        visibles.append({
            'Suministro':r['Suministro'],
            'Real kWh':r['Consumo real (kWh)'],
            'Proyectado kWh':r['Consumo proyectado (kWh)'],
            'Total kWh':r['Total energía (kWh)'],
            'Tarifa CNE $/kWh':r['Precio regulado base ($/kWh)'],
            'Cargos CNE $':r['Cargos CNE referenciales ($)'],
            'Energía contrato $':r['Costo energía contractual ($)'],
        })
    st.dataframe(pd.DataFrame(visibles),hide_index=True,use_container_width=True,
                 column_config={
                     'Real kWh':st.column_config.NumberColumn(format='%.0f'),
                     'Proyectado kWh':st.column_config.NumberColumn(format='%.0f'),
                     'Total kWh':st.column_config.NumberColumn(format='%.0f'),
                     'Tarifa CNE $/kWh':st.column_config.NumberColumn(format='%.3f'),
                     'Cargos CNE $':st.column_config.NumberColumn(format='%.0f'),
                     'Energía contrato $':st.column_config.NumberColumn(format='%.0f'),
                 })
    # Aclaración sobre la diferencia entre total facturado e importes estimados.
    if not completos_kwh:
        st.warning('Falta el consumo de al menos un suministro. No se lo reemplaza por cero.')
    if any(r['Cobertura (días)'] is not None and r['Cobertura (días)'] < r['Días del mes'] for r in info):
        st.warning('La energía disponible cubre solamente una parte del mes seleccionado. '
                   'No representa necesariamente el consumo de un mes completo.')
    if not completos_tar:
        st.warning('No todas las tarifas tienen un valor oficial identificado para el mes. '
                   'Un cargo sin tarifa no se calcula como cero.')
    if not completos_contrato:
        st.info('El costo contractual requiere que la sección «Costo de energía» complete '
                'su cálculo de CPI y dólar. No se estima aquí con cifras supuestas.')
    st.caption('**Importante:** los cargos CNE son un escenario regulado, no el total facturado. '
               'No están incluidos los pagos individuales de PE, SSCC operacionales/BESS, '
               'PSUF, MT/PD, Exenciones CI ni eventuales cargos ya incorporados en factura. '
               'No se suman importes que puedan duplicar cobros.')
    def _mes_en_tarifas():
        st.session_state['tarifas_cne_modo_periodos']='Un mes'
        st.session_state['tarifas_cne_anio_unico']=year
        st.session_state['tarifas_cne_mes_unico']=month
    col1,col2=st.columns([1,1])
    with col1:
        st.button('Usar este mes en Tarifas CNE',on_click=_mes_en_tarifas,
                  key='panel_ir_tarifas',use_container_width=True)
    with col2:
        st.download_button('Descargar resumen CSV',data=_csv(info),
                           file_name=f'EFE_resumen_{year}_{month:02d}.csv',
                           mime='text/csv',key='panel_resumen_csv',use_container_width=True)
    with st.expander('Ver fuentes, resoluciones y disponibilidad de datos',expanded=False):
        st.caption('Catálogo CNE: '+fuente_hist+' · '+VERSION_PANEL)
        st.dataframe(pd.DataFrame([{
            'Suministro':r['Suministro'], 'Origen consumo':r['Fuente consumo'],
            'Cobertura (días)':r['Cobertura (días)'],
            'REX Transmisión':r['REX transmisión'],
            'REX Servicio Público':r['REX servicio público'],
            'Estado costo contractual':r['Estado costo energía'],
        } for r in info]),hide_index=True,use_container_width=True)
        if avisos:
            st.caption('Avisos de lectura SEAT/UMR:')
            for aviso in avisos[:10]:
                st.write('• '+str(aviso))
        st.caption('Los filtros de dólar se mantienen independientes. '
                   'Para comparar varios meses, usa Proyección de energía o Tarifas CNE.')
