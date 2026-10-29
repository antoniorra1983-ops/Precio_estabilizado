"""Panel ejecutivo compacto para EFE Valparaíso.

Totalización estimativa por mes de factura para EFE Valparaíso.
No duplica importes: utiliza proyección, tarifas CNE, costo contractual y
reliquidaciones del mes anterior. Los valores ausentes son None, nunca ceros.
Archivo 4 y escenarios condicionados NO se suman a los importes del Archivo 9.
"""
from __future__ import annotations

import calendar
import math
from datetime import date
from io import StringIO
import csv

import pandas as pd
import streamlit as st

from sscc_archivo4 import calcular_netos_empresa

VERSION_PANEL = "Panel ejecutivo v4.0 · SSCC+CO neto archivo 4 · 2026-10-28"
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



def _completar_costo_contractual(resultados, registros_costo):
    """Integra el cálculo contractual al panel sin depender de otras pestañas.

    Primero reutiliza el valor de «Costo de energía» solo cuando corresponde
    exactamente al mismo período, suministro y consumo. Si esa pestaña calcula
    otro período o utiliza otros kWh, valoriza los kWh del Resumen con el mismo
    motor CPI/dólar contractual de costo_energia.py.

    Nunca sustituye CPI/dólar faltante por cero ni presenta un valor inventado.
    """
    mensajes=[]
    pendientes=[]
    for fila in resultados:
        mes, sitio = fila['Mes'], fila['Suministro']
        energia_panel = _num(fila['Total energía (kWh)'])
        original = next((r for r in (registros_costo or [])
                         if (f"{int(r.get('Año') or 0):04d}-{int(r.get('Mes') or 0):02d}" == mes
                             and r.get('Suministro') == sitio)), None)
        costo_guardado = _num(original.get('Costo CLP')) if original else None
        kwh_guardado = _num(original.get('kWh')) if original else None
        # El costo previo es un dato ya calculado y auditable. Si falló el
        # enlace SEAT/UMR→Tarifas, conservar sus kWh como fuente explícita;
        # así el Resumen no esconde un costo disponible en la pestaña original.
        if energia_panel is None and kwh_guardado is not None and kwh_guardado >= 0:
            fila['Total energía (kWh)'] = kwh_guardado
            fila['Fuente consumo'] = 'Costo de energía (SEAT/SAF LI, respaldo)'
            fila['Cobertura (días)'] = None
            energia_panel = kwh_guardado
            mensajes.append(f'{sitio}: kWh recuperados del cálculo de Costo de energía '
                            'porque el enlace de Proyección no proporcionó este mes.')
        mismos_kwh = (energia_panel is not None and kwh_guardado is not None
                      and abs(energia_panel - kwh_guardado) <= max(0.1, energia_panel*0.001))
        fila['CPI aplicado'] = _num(original.get('CPI')) if original else None
        fila['Precio contrato (US$/MWh)'] = _num(original.get('US$/MWh')) if original else None
        fila['Dólar aplicado ($/US$)'] = _num(original.get('CLP/USD')) if original else None
        fila['Costo energía (US$)'] = _num(original.get('Costo USD')) if original else None
        fila['Fuente costo energía'] = original.get('Fuente dólar', '') if original else ''
        if costo_guardado is not None and mismos_kwh:
            fila['Costo energía contractual ($)'] = costo_guardado
            fila['Estado costo energía'] = original.get('Estado', 'Calculado') + ' · pestaña Costo de energía'
            continue
        if energia_panel is None or energia_panel < 0:
            fila['Costo energía contractual ($)'] = None
            fila['Estado costo energía'] = 'Sin consumo para valorizar'
            continue
        if original and costo_guardado is not None and not mismos_kwh:
            mensajes.append(f'{sitio}: el consumo de Costo de energía difiere del Resumen; '
                            'se recalcula el precio contractual para los kWh del Resumen.')
        pendientes.append(fila)
    if not pendientes:
        return resultados, mensajes

    # IMPORTANTE: usar las mismas funciones que «Costo de energía», no otra tasa.
    try:
        from costo_energia import (cpi_referencia, indices_cpi, precio_mwh,
                                   _tipos_cambio)
        cpis, _ = indices_cpi()
    except Exception as exc:
        for fila in pendientes:
            fila['Costo energía contractual ($)']=None
            fila['Estado costo energía']='Error al consultar CPI: '+str(exc)[:90]
        return resultados, mensajes + ['Consulta CPI: '+str(exc)[:120]]

    # Respetar exactamente las opciones manuales seleccionadas en Costo de energía.
    estado = st.session_state
    cpi_sust = _num(estado.get('ce_cpi_oct25'))
    if cpi_sust is not None and cpi_sust > 0:
        cpis[(2025, 10)] = cpi_sust
    for fila in pendientes:
        yy, mm=map(int, fila['Mes'].split('-'))
        ref_y, ref_m=cpi_referencia(yy, mm)
        cpi_manual=_num(estado.get(f'ce_cpi_manual_{ref_y}_{ref_m}'))
        if cpi_manual is not None and cpi_manual > 0:
            cpis[(ref_y,ref_m)] = cpi_manual

    modo = estado.get('ce_dolar_mode', 'Mes de consumo')
    if modo not in ('Mes de consumo','Mes siguiente (facturación)','Manual / escenario'):
        modo = 'Mes de consumo'
    dolar_manual = _num(estado.get('ce_fx_manual')) or 0.0
    dolar_respaldo = _num(estado.get('ce_fx_backup')) or 0.0
    meses = sorted({tuple(map(int, r['Mes'].split('-'))) for r in pendientes})
    try:
        cambios, errores_fx = _tipos_cambio(meses, modo, dolar_manual, dolar_respaldo)
        mensajes.extend(errores_fx)
    except Exception as exc:
        cambios = {}
        mensajes.append('Consulta de dólar: '+str(exc)[:180])
    for fila in pendientes:
        yy, mm=map(int, fila['Mes'].split('-'))
        cpi=cpis.get(cpi_referencia(yy,mm))
        cpi=_num(cpi)
        monto_usd=None
        precio=None
        if cpi is not None and cpi > 0:
            try:
                precio=precio_mwh(cpi)
                monto_usd=fila['Total energía (kWh)']/1000*precio
            except (ValueError,TypeError,ZeroDivisionError):
                pass
        tasa, fuente_dolar = cambios.get((yy,mm), (None,'No disponible'))
        tasa=_num(tasa)
        monto_clp=(monto_usd*tasa if monto_usd is not None and tasa is not None and tasa>0 else None)
        fila['CPI aplicado']=cpi
        fila['Precio contrato (US$/MWh)']=precio
        fila['Dólar aplicado ($/US$)']=tasa
        fila['Costo energía (US$)']=monto_usd
        fila['Costo energía contractual ($)']=monto_clp
        fila['Fuente costo energía']=fuente_dolar
        fila['Estado costo energía']=(
            'Sin CPI contractual' if precio is None else
            'Sin dólar publicado' if monto_clp is None else
            'Calculado · escenario FX' if 'manual' in fuente_dolar.lower() else
            'Calculado · CPI y dólar observado'
        )
    return resultados, mensajes


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




# ────────────────── Tablero de cargos por concepto ──────────────────
# Mantener separados: cargo regulado por kWh, pago individual CEN y costo
# contractual. Sin una asignación oficial NO se reemplaza un monto por cero.
# Los pagos del Balance CEN se muestran con su signo cuando corresponde.

CARGOS_REGULADOS = (
    ('Transmisión Nacional', 'Nacional ($/kWh)', 'Tarifa REX CNE × kWh', 'Regulado · estimado'),
    ('Transmisión Zonal C', 'Zonal ($/kWh)', 'Sistema C: Quilpué 110 kV / Limache 13 kV', 'Regulado · estimado'),
    ('Exenciones generales', 'Exenciones generales ($/kWh)', 'Tabla 4; distinta de Exenciones CI', 'Regulado · referencial'),
    ('Infraestructura SSCC', 'SSCC infraestructura ($/kWh)', 'Tabla 7; distinto de SSCC operacional/BESS', 'Regulado · estimado'),
    ('Servicio Público (CSP/FET)', 'CSP >5000 kWh ($/kWh)', 'CSP base + FET, sin duplicarlos', 'Regulado · escenario'),
)


def _reporte_mes(reportes, mes):
    return next((r for r in (reportes or []) if isinstance(r, dict)
                 and r.get('periodo') == mes), None)


def _potencia_para_factura(reportes, mes_factura):
    """Prioriza el mes de factura explícito de Balance2 (no el filtro global)."""
    return next((r for r in (reportes or []) if isinstance(r, dict)
                 and r.get('mes_factura') == mes_factura
                 and r.get('periodo') == _mes_anterior(mes_factura)), None)


def _registro_sitio(valor, nota, estado, origen, *, tarifa_clp_kwh=None,
                    tipo_tarifa=None, base_kwh=None, potencia_mw=None,
                    tarifa_clp_kw_mes=None):
    """Monto y unidad de cálculo, sin convertir un cargo de potencia a $/kWh.

    tarifa_clp_kwh: tarifa efectivamente expresada en $/kWh, o indicador
    explícitamente denominado efectivo/referencial, según tipo_tarifa.
    """
    return {'valor': _num(valor), 'nota': nota, 'estado': estado, 'origen': origen,
            'tarifa_clp_kwh': _num(tarifa_clp_kwh),
            'tipo_tarifa': tipo_tarifa,
            'base_kwh': _num(base_kwh),
            'potencia_mw': _num(potencia_mw),
            'tarifa_clp_kw_mes': _num(tarifa_clp_kw_mes)}


def _mes_anterior(mes_factura):
    """Mes del balance CEN liquidado al mes siguiente (incluye dic -> ene)."""
    year, month = map(int, mes_factura.split('-'))
    return f'{year - (month == 1):04d}-{12 if month == 1 else month - 1:02d}'


def _nombre_mes(periodo):
    yy, mm = map(int, periodo.split('-'))
    return f'{MESES[mm - 1]} {yy}'


def _retiros_del_balance(balance):
    """kWh por suministro del balance anterior; ausencia NO equivale a cero.

    Replica el criterio de balance_consumption_by_site del motor del CEN.
    Los valores aquí son retiros originales de la reliquidación y nunca el
    consumo eléctrico del mes de factura.
    """
    import unicodedata

    def normal(txt):
        text = unicodedata.normalize('NFKD', str(txt or ''))
        return ''.join(c for c in text if not unicodedata.combining(c)).upper()

    if balance is None:
        return {s: None for s in SITIOS}
    totales = {s: 0.0 for s in SITIOS}
    encontrados = {s: False for s in SITIOS}
    for reg in getattr(balance, 'detalle', None) or []:
        if not isinstance(reg, dict):
            continue
        clave = normal(reg.get('clave'))
        descripcion = normal(reg.get('descripcion'))
        sitio = ('Quilpué' if clave == 'MERVAL_ELSOL' or 'QUILPUE' in (clave + descripcion)
                 else 'Limache' if clave == 'MERVAL_LIMACHE' or 'LIMACHE' in (clave + descripcion)
                 else None)
        if not sitio:
            continue
        valor = _num(reg.get('kwh'))
        if valor is not None and valor >= 0:
            encontrados[sitio] = True
            totales[sitio] += valor
    return {s: totales[s] if encontrados[s] else None for s in SITIOS}


def desglose_todos_los_costos(info, catalogo, balances=None, sscc_reports=None,
                              potencia_reports=None, exenciones_ci_reports=None,
                              mt_pd_directos=None, sscc_companies=None,
                              reliquidaciones=None):
    """Desglose por MES DE FACTURA sin confundirlo con el MES DE BALANCE.

    El balance previo alimenta PE, SSCC 4/9 y MT/PD; la energía contractual
    y las tarifas del mes elegido permanecen en su propio período. El Archivo
    4 de SSCC se presenta SOLO como estimación, nunca se suma al Archivo 9.
    """
    from tarifas_cne import tabla_periodos
    if not info:
        return []
    mes=info[0]['Mes']  # mes elegido = factura
    mes_balance=_mes_anterior(mes)
    year, month=map(int,mes.split('-'))
    tarifas={(r['Mes'],r['Suministro']):r
             for r in tabla_periodos(catalogo,[year],[month])}
    # Fuente aislada para Resumen: el filtro global de períodos no debe impedir
    # consultar septiembre cuando el usuario seleccionó factura octubre.
    ctx = reliquidaciones if isinstance(reliquidaciones, dict) else {}
    balances_cen = ctx.get('balances', balances) or []
    reportes_9 = ctx.get('sscc_reports', sscc_reports) or []
    reportes_4 = ctx.get('sscc_companies', sscc_companies) or []
    reportes_mtpd = ctx.get('mt_pd_directos', mt_pd_directos) or []
    reportes_potencia = ctx.get('potencia_reports', potencia_reports) or []
    reportes_ci = ctx.get('exenciones_ci_reports', exenciones_ci_reports) or []
    proformas = ctx.get('proformas_enel', []) or []
    saldo_pe=next((b for b in balances_cen if getattr(b,'label',None)==mes_balance),None)
    sscc=_reporte_mes(reportes_9,mes_balance)
    sscc4=_reporte_mes(reportes_4,mes_balance)
    mtpd=_reporte_mes(reportes_mtpd,mes_balance)
    retiro_mes_balance=_retiros_del_balance(saldo_pe)
    # La propia función parse_potencia_cen publica mes_factura=mes_balance+1.
    # Por eso se toma el Balance2 cuya factura corresponde al mes seleccionado.
    # Nunca se reemplaza con un Balance2 de otro período.
    potencia=_potencia_para_factura(reportes_potencia,mes)
    # Liquidación de peaje / PagoExenCI pertenece al balance previo, NO
    # al mes corriente que se eligió en el filtro de factura.
    exenciaci=_reporte_mes(reportes_ci,mes_balance)
    proforma_actual=_reporte_mes(proformas,mes)
    tarjetas=[]
    def add(nombre, grupo, by_site, explicacion='', incluir_total=False, balance_previo=False):
        # No reemplazar importes ausentes por cero ni sumar cargos solapados.
        tarjetas.append({'nombre':nombre,'grupo':grupo,'por_sitio':by_site,
                         'explicacion':explicacion,'incluir_total':incluir_total,
                         'mes_balance':mes_balance if balance_previo else mes,
                         'mes_factura':mes})

    # Energía contractual, calculada por el mismo motor CPI/dólar de costo_energia.
    add('Energía contractual','Energía',{
        r['Suministro']:_registro_sitio(
            r.get('Costo energía contractual ($)'),
            f"{_es(r.get('Total energía (kWh)'))} kWh · "
            f"{_es(r.get('Precio contrato (US$/MWh)'),2)} US$/MWh · "
            f"{_es(r.get('Dólar aplicado ($/US$)'),2)} $/US$",
            r.get('Estado costo energía','Pendiente'),
            r.get('Fuente costo energía','Contrato Enel · CPI y dólar'),
            tarifa_clp_kwh=(
                _num(r.get('Costo energía contractual ($)')) / _num(r.get('Total energía (kWh)'))
                if _num(r.get('Costo energía contractual ($)')) is not None
                and _num(r.get('Total energía (kWh)')) is not None
                and _num(r.get('Total energía (kWh)')) > 0 else None),
            tipo_tarifa='Costo efectivo',
            base_kwh=r.get('Total energía (kWh)'))
        for r in info},'Contrato con CPI reajustado y dólar aplicado al mes seleccionado',
        incluir_total=True)

    # Cada componente del catálogo es una tarjeta independiente; nunca sumar
    # el subtotal del catálogo a estos rubros, porque los contiene.
    for nombre,campo,explicacion,estado in CARGOS_REGULADOS:
        valores={}
        for r in info:
            sitio=r['Suministro']
            tr=tarifas.get((mes,sitio),{})
            kwh=_num(r.get('Total energía (kWh)'))
            tarifa=_num(tr.get(campo))
            # CNE CSP tramo alto: si kWh conocidos bajo umbral, requiere validar tramo.
            tramo_pendiente=(campo=='CSP >5000 kWh ($/kWh)' and kwh is not None and kwh <= 5000)
            monto=(kwh*tarifa if kwh is not None and kwh>=0 and tarifa is not None
                   and not tramo_pendiente else None)
            estado_sitio=('Validar tramo CSP' if tramo_pendiente else estado if monto is not None
                          else 'Falta tarifa o kWh')
            fuente=(tr.get('REX Servicio Público') if campo.startswith('CSP')
                    else tr.get('REX Transmisión'))
            valores[sitio]=_registro_sitio(monto,
                f"{_es(kwh)} kWh × {_es(tarifa,3)} $/kWh" if not tramo_pendiente
                else 'Validar tramo de consumo aplicable',
                estado_sitio,fuente or 'Resolución no identificada',
                tarifa_clp_kwh=tarifa, tipo_tarifa='Tarifa CNE', base_kwh=kwh)
        add(nombre,'Regulados CNE',valores,explicacion,incluir_total=True)

    # Tablas de régimen transitorio: solo aplicables a Limache no individualizado,
    # pero condicionadas a validación. No se integran a ningún total automático.
    for nombre,campo in [('Régimen transitorio · Tabla 5','Tabla 5 NI ($/kWh)'),
                         ('Régimen transitorio · Tabla 6','Tabla 6 NI ($/kWh)')]:
        valores={}
        for r in info:
            sitio=r['Suministro']
            if sitio!='Limache':
                valores[sitio]=_registro_sitio(None,'No aplica: cliente individualizado',
                                               'No aplica','REX CNE')
                continue
            tr=tarifas.get((mes,sitio),{})
            u=_num(tr.get(campo))
            kwh=_num(r.get('Total energía (kWh)'))
            monto=kwh*u if kwh is not None and kwh>=0 and u is not None else None
            valores[sitio]=_registro_sitio(monto,
                f'{_es(kwh)} kWh × {_es(u,5)} $/kWh',
                'Escenario condicionado · no confirmado',tr.get('REX Transmisión','REX CNE'),
                tarifa_clp_kwh=u, tipo_tarifa='Tarifa condicionada', base_kwh=kwh)
        add(nombre,'Condicionados',valores,
            'Solo Limache no individualizado; verificar aplicación contractual')

    # PE factura del mes elegido corresponde al BALANCE CEN PREVIO.
    # No se multiplica la tarifa de septiembre por la proyección de octubre.
    pe_valores={}
    for r in info:
        sitio=r['Suministro']
        kwh=retiro_mes_balance[sitio]
        tarifa=_num(getattr(saldo_pe,'tarifa_clp_mwh',None))
        monto=kwh/1000.0*tarifa if kwh is not None and kwh>=0 and tarifa is not None else None
        pe_valores[sitio]=_registro_sitio(monto,
            f'Balance {_nombre_mes(mes_balance)}: {_es(kwh)} kWh × {_es(tarifa,2)} $/MWh ÷ 1.000',
            'Referencial SEN · no asignación oficial' if monto is not None else
            'Pendiente: balance y retiro EFE del mes anterior',
            f'Balance CEN {_nombre_mes(mes_balance)} · factura {_nombre_mes(mes)}',
            tarifa_clp_kwh=tarifa / 1000.0 if tarifa is not None else None,
            tipo_tarifa='Tasa SEN referencial', base_kwh=kwh)
    add('Precio Estabilizado (PE)','CEN mensual',pe_valores,
        'Estimación con tasa media SEN y kWh del balance CEN anterior; NO es monto individual oficial',
        incluir_total=True, balance_previo=True)

    # Una sola línea del Archivo 9 para EFE: suma TODOS los TIPO_PAGO
    # asociados al suministrador ENEL_GENERACION y al retiro de cada sitio.
    # Incluye SC_SSCC, BESS, CO, CO_ERNC, CPF(+/-), CSF(+/-),
    # CTF(+/-), CRA y REA. Los componentes quedan disponibles únicamente
    # para la auditoría interna; jamás se agregan como tarjetas separadas.
    sscc_valores={}
    for r in info:
        sitio=r['Suministro']
        v=((sscc or {}).get('resumen') or {}).get(sitio)
        if isinstance(v, dict) and v:
            importes={codigo: _num(monto) for codigo,monto in v.items()}
            validos=all(importe is not None for importe in importes.values())
            # El parser del Archivo 9 exige encontrar las dos claves de EFE.
            # Si un reporte externo trae importes no numéricos no se suman
            # como cero: se marca el cargo como pendiente.
            importe=sum(importes.values()) if validos else None
            reporte_suma=_num(((sscc or {}).get('suma_tipos') or {}).get(sitio))
            conciliado=(reporte_suma is None or importe is None or
                        abs(reporte_suma-importe)<0.01)
            if not conciliado:
                importe=None
            fuente=(sscc or {}).get('archivo','Archivo 9 CEN')
            kwh_balance=retiro_mes_balance.get(sitio)
            efectivo=(importe/kwh_balance if importe is not None
                      and kwh_balance is not None and kwh_balance>0 else None)
            registro=_registro_sitio(
                importe,
                f'Balance {_nombre_mes(mes_balance)} · suma de tipos de pago por retiro del Archivo 9',
                ('Pago CEN por retiro · total de tipos' if importe is not None
                 else 'Archivo 9 con importes incompletos o no conciliados'),
                fuente,
                tarifa_clp_kwh=efectivo,
                tipo_tarifa='Equivalente efectivo Archivo 9',
                base_kwh=kwh_balance)
            registro['componentes']=importes
            registro['filas_origen']=[x for x in ((sscc or {}).get('detalle') or [])
                                       if x.get('punto')==sitio]
            sscc_valores[sitio]=registro
        else:
            sscc_valores[sitio]=_registro_sitio(None,
                f'No hay asignación por retiro del balance {_nombre_mes(mes_balance)}',
                'Pendiente Archivo 9 del mes anterior','PLABACOM')
    add('Total Archivo 9 · pagos por retiro','CEN mensual',sscc_valores,
        'Suma de todos los tipos de pago para los retiros EFE en el Archivo 9 del mes anterior; '
        'no equivale necesariamente al concepto SSCC facturado por Enel',
        incluir_total=True, balance_previo=True)

    # Archivo 4: costo neto global SSCC + CO = (PAGA-RECIBE) para ambos.
    # CO negativo se mantiene como crédito, jamás se usa valor absoluto.
    # Solo referencia prorrateada; excluida del total para no duplicar Archivo 9.
    sscc4_valores={}
    netos_4 = None
    try:
        if sscc4:
            netos_4 = calcular_netos_empresa(sscc4)
    except (ValueError, TypeError, KeyError):
        netos_4 = None
    total_enel=_num(getattr(saldo_pe,'retiro_enel_mwh',None))
    for r in info:
        sitio=r['Suministro']
        kwh=retiro_mes_balance[sitio]
        valido=(netos_4 is not None and total_enel is not None and total_enel > 0
                and kwh is not None and kwh >= 0)
        participacion=kwh/(total_enel*1000.0) if valido else None
        importe_sscc=(netos_4['SC_SSCC']['costo_neto']*participacion if valido else None)
        importe_co=(netos_4['CO']['costo_neto']*participacion if valido else None)
        estimado=importe_sscc+importe_co if valido else None
        sscc4_valores[sitio]=_registro_sitio(estimado,
            (f'Balance {_nombre_mes(mes_balance)} · PAGA−RECIBE SSCC: {_es(importe_sscc,2)} $; '
             f'CO: {_es(importe_co,2)} $ · prorrateo { _es(kwh) } kWh; sin BESS') if valido
            else f'Pendiente Archivo 4 y Balance CEN de {_nombre_mes(mes_balance)}',
            'Referencia neta SSCC + CO · NO sumar con Archivo 9' if valido else
            'Pendiente Archivo 4 o Balance CEN',
            (sscc4 or {}).get('archivo','Archivo 4 CEN'))
    add('SSCC + CO · Archivo 4 (neto estimado)','Referencias CEN',sscc4_valores,
        'Neto Enel = (PAGA−RECIBE) SSCC + (PAGA−RECIBE) CO; prorrateo de retiros. '
        'Puede ser negativo; NO es asignación individual ni incluye BESS. Fuera del total.',
        balance_previo=True)

    psuf_valores={}
    ci_valores={}
    mt_valores={}
    pd_valores={}
    for r in info:
        sitio=r['Suministro']
        d=((potencia or {}).get('detalle') or {}).get(sitio)
        monto=_num(d.get('monto_retiro_psuf_clp')) if isinstance(d,dict) else None
        potencia_mw=_num(d.get('potencia_retiro_mw')) if isinstance(d,dict) else None
        tarifa_kw_mes=_num(d.get('precio_nudo_clp_kw_mes')) if isinstance(d,dict) else None
        psuf_valores[sitio]=_registro_sitio(monto,
            f"Balance {_nombre_mes(mes_balance)}: {_es(potencia_mw,4)} MW × "
            f"1.000 × {_es(tarifa_kw_mes,3)} $/kW/mes · Balance2!G",
            'CEN · pago de retiro' if monto is not None else
            f'Pendiente Balance2 para factura {_nombre_mes(mes)}',
            (potencia or {}).get('archivo',
                f'Balance de potencia de {_nombre_mes(mes_balance)}'),
            potencia_mw=potencia_mw, tarifa_clp_kw_mes=tarifa_kw_mes)
        if exenciaci is not None:
            ci_monto=_num(exenciaci.get('total_clp')) if sitio=='Quilpué' else 0.0
        else:
            ci_monto=None
        ci_valores[sitio]=_registro_sitio(ci_monto,
            'Liquidación de peaje / Exenciones CI del balance '+_nombre_mes(mes_balance)+
            ' · Criterio EFE: Quilpué 100 %, Limache 0 %' if exenciaci else
            'Sin PagoExenCI del balance '+_nombre_mes(mes_balance),
            'CEN · asignación por regla EFE' if ci_monto is not None else
            'Pendiente PagoExenCI del mes anterior',
            (exenciaci or {}).get('archivo','PagoExenCI'))
        directos=((mtpd or {}).get('resumen') or {}).get(sitio) or {}
        for name,out in [('MT',mt_valores),('PD',pd_valores)]:
            value=_num(directos.get(name)) if isinstance(directos,dict) else None
            out[sitio]=_registro_sitio(value,
                f'Balance {_nombre_mes(mes_balance)} · asignación individual por retiro',
                'CEN · asignación individual' if value is not None else
                'Pendiente detalle CEN del mes anterior',
                (mtpd or {}).get('archivo','Detalle de MT/PD del CEN'))
    add('Potencia PSUF','CEN mensual',psuf_valores,
        'Balance2 del mes anterior; potencia en MW, tarifa en $/kW/mes; '
        'el pago final puede diferir del cobro de Enel',
        incluir_total=True, balance_previo=True)
    add('Exenciones CI · Liquidación de peaje','CEN mensual',ci_valores,
        'PagoExenCI del mes anterior; solo Quilpué según criterio EFE, separado de '
        'las exenciones generales de transmisión',
        incluir_total=True, balance_previo=True)

    # Peaje CI que figura expresamente en la proforma de la factura: se
    # visualiza para conciliar con PagoExenCI, no se agrega al total calculado.
    peaje_proforma={}
    for r in info:
        sitio=r['Suministro']
        lineas=((proforma_actual or {}).get('suministros') or {}).get(sitio,{}).get('cargos',[])
        candidatas=[c for c in lineas if c.get('Grupo')=='Liquidación de peaje'
                    and c.get('Mes de origen')==mes_balance]
        importes=[_num(c.get('Monto CLP')) for c in candidatas]
        monto=sum(importes) if candidatas and all(v is not None for v in importes) else None
        peaje_proforma[sitio]=_registro_sitio(monto,
            'Origen '+_nombre_mes(mes_balance)+' · factura '+_nombre_mes(mes)+
            ' · partida de peaje CI de la proforma, NO sumar de nuevo',
            'Proforma · para conciliación, no acumulable' if monto is not None else
            'No figura peaje CI en la proforma seleccionada',
            (proforma_actual or {}).get('archivo','Proforma Enel'))
    add('Liquidación de peaje CI · Proforma','Referencias CEN',peaje_proforma,
        'Monto real informado en la proforma, excluido del total calculado para '
        'no duplicar otras partidas de transmisión/Exenciones CI',balance_previo=True)
    add('Mínimos Técnicos (MT)','CEN mensual',mt_valores,
        'Solo importes asignados oficialmente por retiro del balance anterior',
        incluir_total=True, balance_previo=True)
    add('Partidas y Detenciones (PD)','CEN mensual',pd_valores,
        'Solo importes asignados oficialmente por retiro del balance anterior',
        incluir_total=True, balance_previo=True)
    return tarjetas


def _total_tarjeta(tarjeta, sitios):
    """Nunca usar suma parcial como si fuera el monto completo del período."""
    aplicables=[tarjeta['por_sitio'].get(s,{}) for s in sitios
                if tarjeta['por_sitio'].get(s,{}).get('estado') != 'No aplica']
    # "No aplica" no es un cobro cero: se excluye exclusivamente de la suma.
    # Esto permite visualizar Tablas 5/6 de Limache aun seleccionando Ambos.
    if not aplicables or any(r.get('valor') is None for r in aplicables):
        return None
    return sum(r['valor'] for r in aplicables)


def totalizar_cobros(tarjetas, sitios):
    """Suma una sola vez cada cargo compatible y conserva los pendientes.

    Solo participan tarjetas con incluir_total=True. Si falta un importe para
    un suministro, el monto disponible es PARCIAL; no se imputa cero. Los
    créditos negativos del CEN conservan su signo en la suma.
    """
    sitios=tuple(sitios)
    incluidas=[t for t in tarjetas if t.get('incluir_total')]
    por_sitio={}
    detalle=[]
    for sitio in sitios:
        subtotal=0.0
        validos=0
        no_aplican=0
        pendientes=[]
        for tarjeta in incluidas:
            dato=(tarjeta.get('por_sitio') or {}).get(sitio) or {}
            estado=str(dato.get('estado') or '')
            valor=_num(dato.get('valor'))
            no_aplica=estado.strip().lower()=='no aplica'
            if no_aplica:
                incluir='No aplica'
                no_aplican+=1
            elif valor is None:
                incluir='Pendiente'
                pendientes.append(tarjeta['nombre'])
            else:
                incluir='Sumado'
                subtotal+=valor
                validos+=1
            detalle.append({
                'Mes factura':tarjeta.get('mes_factura',''),
                'Mes balance':tarjeta.get('mes_balance',''),
                'Suministro':sitio,
                'Concepto':tarjeta['nombre'],
                'Grupo':tarjeta['grupo'],
                'Monto CLP':valor if incluir=='Sumado' else None,
                'Tarifa/efectivo ($/kWh)':dato.get('tarifa_clp_kwh'),
                'Clase valor unitario':dato.get('tipo_tarifa'),
                'Consumo base (kWh)':dato.get('base_kwh'),
                'Potencia (MW)':dato.get('potencia_mw'),
                'Tarifa potencia ($/kW/mes)':dato.get('tarifa_clp_kw_mes'),
                'Inclusión':incluir,
                'Estado fuente':estado,
                'Origen':dato.get('origen',''),
            })
        por_sitio[sitio]={
            'subtotal':subtotal if validos else None,
            'n_disponibles':validos,
            'pendientes':pendientes,
            'completo':(validos>0 or no_aplican>0) and not pendientes,
        }
    total=sum(v['subtotal'] for v in por_sitio.values() if v['subtotal'] is not None)
    disponibles=sum(v['n_disponibles'] for v in por_sitio.values())
    return {
        'subtotal':total if disponibles else None,
        'completo':bool(disponibles) and all(v['completo'] for v in por_sitio.values()),
        'por_sitio':por_sitio,
        'pendientes':[(sitio,nombre) for sitio,v in por_sitio.items() for nombre in v['pendientes']],
        'detalle':detalle,
        'conceptos':len(incluidas),
    }


def _mostrar_total_cobros(tarjetas, sitios):
    """Presenta total estimado o subtotal identificado sin confundirlo con factura."""
    balance=totalizar_cobros(tarjetas,sitios)
    st.markdown('### Total de cobros del mes de facturación')
    total=balance['subtotal']
    etiqueta=('Total estimado de cobros' if balance['completo'] else
              'Subtotal identificado · parcial')
    texto=('$'+_es(total)) if total is not None else 'Sin montos disponibles'
    psuf=next((t for t in tarjetas if t.get('nombre')=='Potencia PSUF'),None)
    total_psuf=_total_tarjeta(psuf,sitios) if psuf else None
    col_total,col_potencia=st.columns([2,1])
    with col_total:
        st.metric(etiqueta,texto)
    with col_potencia:
        st.metric('Potencia PSUF · incluida en la suma',
                  '$'+_es(total_psuf) if total_psuf is not None else 'Pendiente')
        if psuf:
            st.caption('Balance: '+_nombre_mes(psuf['mes_balance'])+
                       ' · Factura: '+_nombre_mes(psuf['mes_factura']))
    if len(sitios)>1:
        columnas=st.columns(len(sitios))
        for col,sitio in zip(columnas,sitios):
            r=balance['por_sitio'][sitio]
            texto_sitio=('$'+_es(r['subtotal'])) if r['subtotal'] is not None else 'Sin datos'
            col.metric(sitio+(' · estimado' if r['completo'] else ' · parcial'),texto_sitio)
    st.caption('Suma de importes en CLP con su signo: energía contractual + cargos regulados '
               '+ PE referencial del balance anterior + total de pagos del Archivo 9 '
               '+ potencia PSUF (Balance2 del mes anterior) + Exenciones CI/peaje '
               'del mes anterior + MT + PD. '
               'Los cargos regulados y el PE son '
               'estimaciones: **no equivale al total facturado por Enel**.')
    if balance['pendientes']:
        st.warning(f"Faltan {len(balance['pendientes'])} importe(s) para completar el total. "
                   'La cifra mostrada es solamente el subtotal disponible; '
                   'ningún faltante se considera $0.')
    if total is None:
        st.info('No se ha identificado ningún monto sumable para este período y suministro.')
    st.caption('**No se suman:** SSCC Archivo 4 (referencia que duplicaría Archivo 9), '
               'liquidación de peaje CI de la proforma (referencia para conciliar), '
               'Tablas 5/6 condicionadas ni subtotal CNE (ya desglosado arriba). '
               'Los conceptos pendientes y las posibles coincidencias con cargos contractuales '
               'requieren conciliación con la factura antes de validar un total a pagar.')
    with st.expander('Ver cómo se calcula el total y qué conceptos faltan',expanded=False):
        st.dataframe(pd.DataFrame(balance['detalle']),hide_index=True,
                     use_container_width=True,
                     column_config={
                         'Monto CLP':st.column_config.NumberColumn(format='%.0f'),
                         'Tarifa/efectivo ($/kWh)':st.column_config.NumberColumn(format='%.5f'),
                         'Consumo base (kWh)':st.column_config.NumberColumn(format='%.0f'),
                         'Potencia (MW)':st.column_config.NumberColumn(format='%.4f'),
                         'Tarifa potencia ($/kW/mes)':st.column_config.NumberColumn(format='%.3f'),
                     })
        st.download_button('Descargar composición del total CSV',
                           data=_csv(balance['detalle']),
                           file_name='EFE_total_cobros_'+tarjetas[0]['mes_factura']+'.csv',
                           mime='text/csv',key='panel_total_cobros_csv',use_container_width=True)
    return balance


def _tarjeta_ui(tarjeta, sitios):
    v=_total_tarjeta(tarjeta,sitios)
    with st.container(border=True):
        es_sscc_9 = tarjeta['nombre'] == 'Total Archivo 9 · pagos por retiro'
        st.metric(tarjeta['nombre'], ('$'+_es(v,2 if es_sscc_9 else 0)) if v is not None else 'Pendiente')
        if tarjeta.get('mes_balance') != tarjeta.get('mes_factura'):
            st.caption('Balance: '+_nombre_mes(tarjeta['mes_balance'])+
                       ' → Factura: '+_nombre_mes(tarjeta['mes_factura']))
        if v is None:
            st.caption('Falta el dato de uno o más suministros' if len(sitios)>1 else
                       tarjeta['por_sitio'].get(sitios[0],{}).get('estado','Sin datos'))
        else:
            estados=[tarjeta['por_sitio'][s].get('estado','') for s in sitios]
            st.caption(' · '.join(dict.fromkeys(estados))[:115])
        # Una tasa aplicada NO se suma entre suministros: se muestra por separado.
        tarifas=[]
        for sitio in sitios:
            dato=tarjeta['por_sitio'].get(sitio,{})
            tasa=_num(dato.get('tarifa_clp_kwh'))
            if tasa is not None:
                clase=dato.get('tipo_tarifa') or 'Tarifa'
                prefijo=f'{sitio}: ' if len(sitios)>1 else ''
                tarifas.append(f'{prefijo}{clase} {_es(tasa,5)} $/kWh')
        if tarifas:
            st.caption(' | '.join(tarifas))
        if tarjeta['nombre']=='Potencia PSUF':
            potencias=[]
            for sitio in sitios:
                dato=tarjeta['por_sitio'].get(sitio,{})
                mw=_num(dato.get('potencia_mw'))
                precio=_num(dato.get('tarifa_clp_kw_mes'))
                if mw is not None and precio is not None:
                    prefijo=f'{sitio}: ' if len(sitios)>1 else ''
                    potencias.append(f'{prefijo}{_es(mw,4)} MW × {_es(precio,3)} $/kW/mes')
            if potencias:
                st.caption(' | '.join(potencias))
        if len(sitios)>1:
            st.caption(' | '.join(f'{s}: '+('$'+_es(tarjeta['por_sitio'].get(s,{}).get('valor'), 2 if es_sscc_9 else 0)
                                           if tarjeta['por_sitio'].get(s,{}).get('valor') is not None
                                           else '—') for s in sitios))


def _mostrar_sscc_archivo_9(tarjetas, sitios):
    """Una sola tarjeta con la suma de los pagos por retiro del Archivo 9.

    La selección Ambos/Quilpué/Limache determina qué importes participan;
    septiembre del CEN se muestra en la factura de octubre. Nunca se
    presentan aquí tarjetas adicionales para SC_SSCC, BESS u otros códigos.
    """
    tarjeta=next((t for t in tarjetas if t.get('nombre')==
                  'Total Archivo 9 · pagos por retiro'),None)
    if not tarjeta:
        return
    sitios=tuple(sitios)
    valor=_total_tarjeta(tarjeta,sitios)
    st.markdown('### Archivo 9 · total de pagos EFE')
    st.metric('Total Archivo 9 (sin IVA)',
              '$'+_es(valor,2) if valor is not None else 'Pendiente')
    st.caption('Balance '+_nombre_mes(tarjeta['mes_balance'])+' → Proforma '
               +_nombre_mes(tarjeta['mes_factura'])+
               ' · Todos los tipos de pago CEN agregados en un único importe.')
    if valor is not None:
        retiros=[_num((tarjeta.get('por_sitio') or {}).get(s,{}).get('base_kwh'))
                 for s in sitios]
        if retiros and all(kwh is not None and kwh>0 for kwh in retiros):
            st.caption('Equivalente efectivo: '+_es(valor/sum(retiros),6)+' $/kWh '
                       '(calculado con los retiros del balance anterior).')
    else:
        st.warning('Falta el Archivo 9 o algún importe por retiro para los suministros seleccionados.')
    st.caption('Este total es el agregado de pagos del Archivo 9; su coincidencia '
               'con los cargos facturados por Enel debe verificarse en la proforma.')


def mostrar_tarjetas_costos(info, catalogo, balances=None, sscc_reports=None,
                             potencia_reports=None, exenciones_ci_reports=None,
                             mt_pd_directos=None, sscc_companies=None,
                             reliquidaciones=None):
    """Las tarjetas presentan cada concepto sin sumar cobros no comparables."""
    tarjetas=desglose_todos_los_costos(info,catalogo,balances,sscc_reports,
                                     potencia_reports,exenciones_ci_reports,mt_pd_directos,
                                     sscc_companies=sscc_companies,
                                     reliquidaciones=reliquidaciones)
    sitios=[r['Suministro'] for r in info]
    mes_factura=info[0]['Mes']
    mes_balance=_mes_anterior(mes_factura)
    # Mostrar una sola tarjeta: la suma completa del Archivo 9.
    _mostrar_sscc_archivo_9(tarjetas, sitios)
    _mostrar_total_cobros(tarjetas,sitios)
    st.markdown('### 💰 Costos por concepto')
    st.caption(f'Factura {_nombre_mes(mes_factura)}: energía contractual y tarifas reguladas '
               f'del mes elegido. PE, SSCC (archivos 4/9), potencia PSUF, MT/PD y liquidación de peaje CI '
               f'proceden del balance {_nombre_mes(mes_balance)}. '
               'Tarifas por energía en $/kWh; potencia en $/kW/mes. '
               'Archivo 4 = estimación neta SSCC + CO (PAGA−RECIBE), sin sumarse al total; '
               'Archivo 9 = cobro único de pagos por retiro.')
    for grupo in ('Energía','Regulados CNE','CEN mensual','Referencias CEN','Condicionados'):
        items=[r for r in tarjetas if r['grupo']==grupo
               and r['nombre']!='Total Archivo 9 · pagos por retiro']
        if not items:
            continue
        etiquetas={'Energía':'Costo del consumo',
                   'Regulados CNE':'Transmisión y cargos regulados',
                   'CEN mensual':'Compensaciones y pagos del CEN',
                   'Referencias CEN':'Referencia neta SSCC + CO · Archivo 4 (no facturable)',
                   'Condicionados':'Cargos sujetos a condición contractual'}
        if grupo=='Condicionados':
            # Los escenarios de régimen transitorio no deben ocupar espacio
            # principal ni interpretarse como cobros definitivos.
            with st.expander('Ver Tablas 5 y 6: solo Limache, aplicación condicionada',expanded=False):
                for i in range(0,len(items),2):
                    cols=st.columns(min(2,len(items)-i))
                    for col,item in zip(cols,items[i:i+2]):
                        with col:
                            _tarjeta_ui(item,sitios)
            continue
        st.markdown('#### '+etiquetas[grupo])
        for i in range(0,len(items),3):
            cols=st.columns(min(3,len(items)-i))
            for col,item in zip(cols,items[i:i+3]):
                with col:
                    _tarjeta_ui(item,sitios)
    # La tabla de auditoría conserva importes negativos cuando el pago CEN
    # tiene signo. Los estados de No aplica y Pendiente se distinguen.
    with st.expander('Ver detalle por suministro, fórmula y fuente',expanded=False):
        filas=[]
        for tarjeta in tarjetas:
            for sitio in sitios:
                dato=tarjeta['por_sitio'].get(sitio,{})
                filas.append({'Grupo':tarjeta['grupo'],'Concepto':tarjeta['nombre'],
                    'Mes factura':tarjeta['mes_factura'],
                    'Mes balance utilizado':tarjeta['mes_balance'],
                    'Suministro':sitio,'Monto CLP':dato.get('valor'),
                    'Tarifa/efectivo ($/kWh)':dato.get('tarifa_clp_kwh'),
                    'Clase valor unitario':dato.get('tipo_tarifa'),
                    'Consumo base (kWh)':dato.get('base_kwh'),
                    'Potencia (MW)':dato.get('potencia_mw'),
                    'Tarifa potencia ($/kW/mes)':dato.get('tarifa_clp_kw_mes'),
                    'Estado':dato.get('estado'),'Cálculo / nota':dato.get('nota'),
                    'Fuente':dato.get('origen')})
        st.dataframe(pd.DataFrame(filas),hide_index=True,use_container_width=True,
                     column_config={
                         'Monto CLP':st.column_config.NumberColumn(format='%.0f'),
                         'Tarifa/efectivo ($/kWh)':st.column_config.NumberColumn(format='%.5f'),
                         'Consumo base (kWh)':st.column_config.NumberColumn(format='%.0f'),
                         'Potencia (MW)':st.column_config.NumberColumn(format='%.4f'),
                         'Tarifa potencia ($/kW/mes)':st.column_config.NumberColumn(format='%.3f'),
                     })
        st.download_button('⬇ Descargar detalle de todos los costos CSV',
                           data=_csv(filas),file_name='EFE_desglose_costos_'+info[0]['Mes']+'.csv',
                           mime='text/csv',key='panel_v3_csv_costos')
    st.caption('El total de cobros mostrado arriba es referencial y puede ser parcial: '
               'PE es una estimación SEN, los pagos CEN conservan sus signos y la factura '
               'del suministrador puede contener cobros incluidos en otros componentes.')
    return tarjetas

# Solo se redibuja este panel al cambiar mes, año o suministro.
# No vuelve a ejecutar app.py ni a procesar otras pestañas.
@st.fragment
def mostrar_panel_ejecutivo(balances=None, archivos_repositorio=None,
                           sscc_reports=None, potencia_reports=None,
                           exenciones_ci_reports=None, mt_pd_directos=None,
                           sscc_companies=None, cargar_reliquidacion=None):
    st.subheader('Resumen mensual · Quilpué y Limache')
    st.caption('Selecciona el mes de FACTURA. Los cargos CEN por reliquidación '
               '(PE, SSCC archivos 4/9, MT/PD y liquidación de peaje CI) corresponden al mes anterior; '
               'los costos energéticos y tarifas mantienen su período propio.')
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
    # Los filtros de año y mes se mantienen; el suministro se selecciona
    # en una fila propia para que las tres vistas sean siempre visibles.
    a,b=st.columns(2)
    with a:
        year=st.selectbox('Año',years,index=years.index(suggested_y),key='panel_periodo_ano')
    with b:
        month=st.selectbox('Mes',list(range(1,13)),index=suggested_m-1,
                           format_func=lambda x: MESES[x-1],key='panel_periodo_mes')
    option=st.radio('Ver suministro', ('Ambos', *SITIOS), horizontal=True,
                    key='panel_suministro_vista_v2', index=0,
                    help='Ambos: suma Quilpué y Limache. Las otras opciones muestran solo el suministro elegido.')
    selected=SITIOS if option=='Ambos' else (option,)
    with st.expander('Datos y velocidad', expanded=False):
        if st.button('Actualizar cálculos de los archivos publicados', key='panel_refresh_cen'):
            st.session_state.pop('efe_reliquidaciones_mes_cache', None)
            st.session_state.pop('efe_panel_periodos_cache', None)
            st.rerun(scope='fragment')
        st.caption('El cambio de mes o suministro reutiliza los cálculos leídos en esta sesión. '
                   'Si reemplazaste archivos en GitHub, reinicia Streamlit para renovar el catálogo.')
    mes_factura=f'{year:04d}-{month:02d}'
    mes_reliquidacion=_mes_anterior(mes_factura)
    st.info('**Factura:** '+_nombre_mes(mes_factura)+' · **Reliquidaciones CEN:** '+
            _nombre_mes(mes_reliquidacion)+
            ' (PE, SSCC 4/9, potencia PSUF, MT, PD y liquidación de peaje CI).')
    st.caption('Vista actual: '+('Quilpué + Limache' if option=='Ambos' else option)
               +' · Se actualizan indicadores, costos y tabla del suministro seleccionado.')
    costo_rows=st.session_state.get('efe_costo_energia_resumen', [])
    # Cachear por mes AMBOS suministros: cambiar la vista no debe consultar
    # nuevamente CPI, dólar ni el modelo de energía. Las entradas se invalidan
    # al cambiar los insumos del cálculo.
    estado_costo=tuple(sorted((k,str(v)) for k,v in st.session_state.items()
                             if k.startswith(('ce_cpi','ce_fx','ce_dolar'))))
    firma=(year,month,
           repr(costo_rows),repr((shared.get('consumos') or {})),
           tuple((b.label,b.consumo_efe_kwh,b.retiro_enel_mwh)
                 for b in (balances or [])),estado_costo,
           str(fuente_hist))
    memoria=st.session_state.setdefault('efe_panel_periodos_cache',{})
    guardado=memoria.get((year,month))
    if guardado is not None and guardado['firma']==firma:
        ambos=guardado['info']
        avisos=list(guardado['avisos'])
    else:
        with st.spinner('Calculando consumo, CPI y dólar del mes…'):
            ambos, avisos=resumen_periodo(year,month,SITIOS,balances,catalogo,costo_rows)
            ambos, avisos_costo=_completar_costo_contractual(ambos, costo_rows)
            avisos.extend(avisos_costo)
        memoria[(year,month)]={'firma':firma,'info':ambos,'avisos':list(avisos)}
        while len(memoria)>24:
            memoria.pop(next(iter(memoria)))
    info=[fila.copy() for fila in ambos if fila['Suministro'] in selected]

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

    # Consultar solamente los archivos del balance CEN anterior. La carga
    # independiente permite factura de octubre con filtro global solo octubre.
    reliquidaciones=None
    if callable(cargar_reliquidacion):
        try:
            with st.spinner('Consultando cargos CEN del mes anterior…'):
                reliquidaciones=cargar_reliquidacion(mes_reliquidacion)
            for aviso in (reliquidaciones or {}).get('avisos', [])[:8]:
                st.caption('Aviso CEN: '+str(aviso))
        except Exception as exc:
            import logging
            logging.exception('EFE_PANEL_EJECUTIVO: fallo al consultar mes anterior')
            reliquidaciones={'balances':[], 'sscc_reports':[], 'sscc_companies':[],
                             'mt_pd_directos':[], 'exenciones_ci_reports':[],
                             'proformas_enel':[], 'potencia_reports':[]}
            st.warning('No se pudieron consultar los archivos CEN del mes anterior; '
                       'las tarjetas respectivas quedan pendientes. '+str(exc)[:150])
    # Un error en tarjetas no debe ocultar el resto del resumen.
    try:
        mostrar_tarjetas_costos(info,catalogo,balances=balances,sscc_reports=sscc_reports,
            potencia_reports=potencia_reports,exenciones_ci_reports=exenciones_ci_reports,
            mt_pd_directos=mt_pd_directos, sscc_companies=sscc_companies,
            reliquidaciones=reliquidaciones)
    except Exception as exc:
        import logging
        logging.exception('EFE_PANEL_EJECUTIVO: error en tarjetas')
        st.warning('No se pudieron representar las tarjetas CEN, pero el resumen principal '
                   'continúa disponible.')
        with st.expander('Diagnóstico de tarjetas de costos'):
            st.exception(exc)

    # Precios y consumos en una tabla compacta con exactamente los meses elegidos.
    st.markdown('#### Comparación de suministros' if option=='Ambos' else f'#### Detalle de {option}')
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
            'Estado energía':r['Estado costo energía'],
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
        pendientes_costo=[f"{r['Suministro']}: {r['Estado costo energía']}"
                          for r in info if r['Costo energía contractual ($)'] is None]
        st.info('Costo contractual no disponible para el período seleccionado: '
                + ' · '.join(pendientes_costo)
                + '. No se reemplazan CPI ni dólar ausentes por cifras supuestas.')
    st.caption('**Importante:** las reliquidaciones PE, SSCC 4/9, MT/PD y peaje CI muestran el '
               'balance CEN anterior al mes de factura. La suma del Resumen es un total '
               'estimado o parcial, no un importe facturado; Archivo 4 no se suma al Archivo 9. '
               'Verifica superposiciones y aplicabilidad con la factura antes de usarla como total a pagar.')
    col1,col2=st.columns([1,1])
    with col1:
        # Este botón es una excepción intencional: modifica otra pestaña, por lo
        # que requiere una ejecución completa para actualizar Tarifas CNE.
        if st.button('Usar este mes en Tarifas CNE',
                     key='panel_ir_tarifas',use_container_width=True):
            st.session_state['tarifas_cne_modo_periodos']='Un mes'
            st.session_state['tarifas_cne_anio_unico']=year
            st.session_state['tarifas_cne_mes_unico']=month
            st.rerun()
    with col2:
        st.download_button('Descargar resumen CSV',data=_csv(info),
                           file_name=f"EFE_resumen_{'Ambos' if option=='Ambos' else 'Quilpue' if option=='Quilpué' else 'Limache'}_{year}_{month:02d}.csv",
                           mime='text/csv',key='panel_resumen_csv',use_container_width=True)
    with st.expander('Ver fuentes, resoluciones y disponibilidad de datos',expanded=False):
        st.caption('Catálogo CNE: '+fuente_hist+' · '+VERSION_PANEL)
        st.dataframe(pd.DataFrame([{
            'Suministro':r['Suministro'], 'Origen consumo':r['Fuente consumo'],
            'Cobertura (días)':r['Cobertura (días)'],
            'REX Transmisión':r['REX transmisión'],
            'REX Servicio Público':r['REX servicio público'],
            'Estado costo contractual':r['Estado costo energía'],
            'CPI aplicado':r.get('CPI aplicado'),
            'US$/MWh':r.get('Precio contrato (US$/MWh)'),
            'Dólar aplicado ($/US$)':r.get('Dólar aplicado ($/US$)'),
            'Costo USD':r.get('Costo energía (US$)'),
            'Fuente del dólar':r.get('Fuente costo energía'),
        } for r in info]),hide_index=True,use_container_width=True)
        if avisos:
            st.caption('Avisos de lectura SEAT/UMR:')
            for aviso in avisos[:10]:
                st.write('• '+str(aviso))
        st.caption('El costo contractual se calcula por período con el CPI y el dólar '
                   'de la pestaña Costo de energía, aunque sus fechas sean otras. '
                   'Los filtros de dólar se mantienen independientes. '
                   'Para comparar varios meses, usa Proyección de energía o Tarifas CNE.')
