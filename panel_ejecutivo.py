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

VERSION_PANEL = "Panel ejecutivo compacto v5.3 · Dólar mensual independiente · 2026-10-29"
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
            (sscc4 or {}).get('archivo','Archivo 4 CEN'),
            tarifa_clp_kwh=(estimado/kwh if valido and kwh > 0 else None),
            tipo_tarifa='Costo efectivo estimado', base_kwh=kwh)
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
            (exenciaci or {}).get('archivo','PagoExenCI'),
            tarifa_clp_kwh=(ci_monto/retiro_mes_balance[sitio]
                if ci_monto is not None and retiro_mes_balance[sitio] is not None
                and retiro_mes_balance[sitio] > 0 else None),
            tipo_tarifa='Costo efectivo', base_kwh=retiro_mes_balance[sitio])
        directos=((mtpd or {}).get('resumen') or {}).get(sitio) or {}
        for name,out in [('MT',mt_valores),('PD',pd_valores)]:
            value=_num(directos.get(name)) if isinstance(directos,dict) else None
            out[sitio]=_registro_sitio(value,
                f'Balance {_nombre_mes(mes_balance)} · asignación individual por retiro',
                'CEN · asignación individual' if value is not None else
                'Pendiente detalle CEN del mes anterior',
                (mtpd or {}).get('archivo','Detalle de MT/PD del CEN'),
                tarifa_clp_kwh=(value/retiro_mes_balance[sitio]
                    if value is not None and retiro_mes_balance[sitio] is not None
                    and retiro_mes_balance[sitio] > 0 else None),
                tipo_tarifa='Costo efectivo', base_kwh=retiro_mes_balance[sitio])
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
            (proforma_actual or {}).get('archivo','Proforma Enel'),
            tarifa_clp_kwh=(monto/retiro_mes_balance[sitio]
                if monto is not None and retiro_mes_balance[sitio] is not None
                and retiro_mes_balance[sitio] > 0 else None),
            tipo_tarifa='Costo efectivo', base_kwh=retiro_mes_balance[sitio])
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


def _indicadores_precios_energia(info):
    """Precios contractuales, no promedios simples de suministros distintos.

    Precio base: constante contractual del módulo costo_energia.
    Precio reajustado CPI: misma tarifa US$/MWh usada en el costo del período.
    Precio final: monto contractual / kWh; o conversión explícita US$/MWh * FX / 1000.
    Para Ambos, las tarifas finales y CPI se ponderan por kWh. Si falta un dato,
    no se anuncia un promedio incompleto como válido para los dos suministros.
    """
    try:
        from costo_energia import PRECIO_BASE_USD_MWH
        base = _num(PRECIO_BASE_USD_MWH)
    except (ImportError, AttributeError, TypeError, ValueError):
        base = None

    filas = list(info or [])

    def media_ponderada(campo):
        if not filas:
            return None
        valores = []
        for fila in filas:
            kwh = _num(fila.get('Total energía (kWh)'))
            valor = _num(campo(fila))
            if kwh is None or kwh < 0 or valor is None:
                return None
            valores.append((valor, kwh))
        total_kwh = sum(k for _, k in valores)
        if total_kwh > 0:
            return sum(v * k for v, k in valores) / total_kwh
        # Sin consumo, solo se muestra una tarifa si es idéntica en todos.
        return valores[0][0] if all(abs(v-valores[0][0]) < 1e-9 for v, _ in valores) else None

    def dolar_comun_o_ponderado():
        # El promedio mensual CLP/USD no depende de los kWh de EFE.
        # Puede mostrarse incluso cuando un suministro no tiene consumo.
        if not filas:
            return None
        tasas = [_num(fila.get('Dólar aplicado ($/US$)')) for fila in filas]
        if any(v is None for v in tasas):
            return None
        if all(abs(v - tasas[0]) < 1e-9 for v in tasas):
            return tasas[0]
        return media_ponderada(lambda fila: fila.get('Dólar aplicado ($/US$)'))

    def unitario_final(fila):
        kwh = _num(fila.get('Total energía (kWh)'))
        costo = _num(fila.get('Costo energía contractual ($)'))
        if kwh is not None and kwh > 0 and costo is not None:
            return costo / kwh
        precio = _num(fila.get('Precio contrato (US$/MWh)'))
        dolar = _num(fila.get('Dólar aplicado ($/US$)'))
        if precio is not None and dolar is not None and dolar > 0:
            return precio * dolar / 1000.0
        return None

    return {
        'base_usd_mwh': base,
        'ajustado_usd_mwh': media_ponderada(
            lambda fila: fila.get('Precio contrato (US$/MWh)')),
        # Misma tasa CLP/USD utilizada para valorizar el costo contractual.
        # Si hay dos suministros, ponderar según sus consumos; nunca sumar divisas.
        'dolar_promedio_clp_usd': dolar_comun_o_ponderado(),
        'final_clp_kwh': media_ponderada(unitario_final),
    }


@st.cache_data(ttl=3600, show_spinner=False, max_entries=72)
def _dolar_independiente_mes(anio, mes, modo, manual, respaldo):
    """Promedio CLP/USD del período elegido, independiente de los kWh/CPI.

    Reutiliza exactamente el motor del módulo Costo de energía. Si su fuente
    principal no responde, consulta la serie histórica de mindicador.cl.
    El respaldo manual solo se presenta como escenario, nunca como observado.
    """
    from costo_energia import _tipos_cambio
    tasa, fuente = None, 'Sin cotización publicada'
    try:
        cambios, _ = _tipos_cambio([(anio, mes)], modo, manual, respaldo)
        tasa, fuente = cambios.get((anio, mes), (None, 'Sin cotización publicada'))
    except Exception:
        pass
    tasa = _num(tasa)
    if modo == 'Manual / escenario':
        return (tasa if tasa is not None and tasa > 0 else None,
                'Escenario manual' if tasa is not None and tasa > 0 else 'Sin valor manual')
    if tasa is not None and tasa > 0 and 'manual' not in str(fuente).lower():
        return tasa, fuente

    # Segunda fuente con la serie publicada (no se usan promedios de otros meses).
    objetivo_anio, objetivo_mes = anio, mes
    if modo == 'Mes siguiente (facturación)':
        if objetivo_mes == 12:
            objetivo_anio, objetivo_mes = anio + 1, 1
        else:
            objetivo_mes += 1
    try:
        from urllib.request import Request, urlopen
        from datetime import date as _date
        import json
        url = f'https://mindicador.cl/api/dolar/{objetivo_anio}'
        req = Request(url, headers={'Accept':'application/json',
                                    'User-Agent':'EFE-Valparaiso-Resumen/1.0'})
        with urlopen(req, timeout=8) as respuesta:
            datos = json.load(respuesta)
        dias = {}
        hoy = _date.today()
        for fila in datos.get('serie', []):
            try:
                f = _date.fromisoformat(str(fila['fecha'])[:10])
                valor = _num(fila['valor'])
            except (KeyError, ValueError, TypeError):
                continue
            if (f.year, f.month) == (objetivo_anio, objetivo_mes) and f <= hoy \
                    and valor is not None and valor > 0:
                dias[f] = valor
        if dias:
            return (sum(dias.values())/len(dias),
                    f'Dólar observado promedio {objetivo_anio}-{objetivo_mes:02d} (respaldo)')
    except Exception:
        pass
    # Si el usuario configuró explícitamente un respaldo, se conserva su signo
    # referencial. Nunca inventar el promedio a partir de otro período.
    if tasa is not None and tasa > 0:
        return tasa, fuente
    return None, 'Sin cotización publicada para el período'


def _completar_dolar_resumen(info):
    """Rellena el CLP/USD faltante incluso cuando no hay costo contractual.

    La tasa se obtiene por período, NO por suministro. El valor no modifica
    los costos ya calculados ni los montos CEN ni la totalización.
    """
    filas = list(info or [])
    if not filas:
        return filas
    faltantes = [r for r in filas
                 if _num(r.get('Dólar aplicado ($/US$)')) is None]
    if not faltantes:
        return filas
    estado = st.session_state
    modo = estado.get('ce_dolar_mode', 'Mes de consumo')
    if modo not in ('Mes de consumo','Mes siguiente (facturación)','Manual / escenario'):
        modo = 'Mes de consumo'
    manual = _num(estado.get('ce_fx_manual')) or 0.0
    respaldo = _num(estado.get('ce_fx_backup')) or 0.0
    for periodo in {r.get('Mes') for r in faltantes if r.get('Mes')}:
        try:
            anio, mes = map(int, str(periodo).split('-'))
            tasa, fuente = _dolar_independiente_mes(anio, mes, modo, manual, respaldo)
        except (ValueError, TypeError):
            tasa, fuente = None, 'Sin cotización publicada'
        # Cuando otra fila ya trae el dólar de ese mismo mes, es válido
        # reutilizarlo entre puntos de retiro en lugar de ocultar la tarjeta.
        if tasa is None:
            disponibles = {_num(r.get('Dólar aplicado ($/US$)'))
                           for r in filas if r.get('Mes') == periodo}
            disponibles.discard(None)
            if len(disponibles) == 1:
                tasa = next(iter(disponibles))
                fuente = 'Misma tasa contractual del período'
        if tasa is not None and tasa > 0:
            for r in faltantes:
                if r.get('Mes') == periodo:
                    r['Dólar aplicado ($/US$)'] = tasa
                    r['Fuente dólar indicador'] = fuente
    return filas


def _mostrar_cabecera_energia(info):
    """Renderizar kWh y precios ANTES de cargar reliquidaciones o tarjetas CEN.

    Esto desacopla los indicadores contractuales de fallos/esperas en SSCC,
    PE y otros conceptos. No recalcula ni modifica ningún monto contractual.
    """
    filas = list(info or [])
    if not filas:
        st.info('No hay información de energía para el período seleccionado.')
        return
    valores = [_num(f.get('Total energía (kWh)')) for f in filas]
    consumo = sum(valores) if all(v is not None for v in valores) else None
    precios = _indicadores_precios_energia(filas)

    st.markdown('#### Consumo y precios de energía')
    a, b, c, d, e = st.columns([1.4, 1.0, 1.2, 1.2, 1.3])
    with a:
        st.metric('Consumo', _es(consumo) + ' kWh' if consumo is not None else '—')
        if len(filas) > 1:
            for f in filas:
                v = _num(f.get('Total energía (kWh)'))
                st.caption(f"{f['Suministro']}: " + (_es(v) + ' kWh' if v is not None else '—'))
    with b:
        precio = precios['base_usd_mwh']
        st.metric('Precio base', _es(precio, 2) + ' US$/MWh' if precio is not None else '—')
    with c:
        precio = precios['ajustado_usd_mwh']
        st.metric('Precio con CPI', _es(precio, 3) + ' US$/MWh' if precio is not None else '—')
    with d:
        dolar = precios['dolar_promedio_clp_usd']
        # La modalidad manual usa un tipo de cambio fijado por el usuario,
        # no un promedio observado mensual; la etiqueta lo distingue.
        modo = st.session_state.get('ce_dolar_mode', 'Mes de consumo')
        fuentes = [str(r.get('Fuente dólar indicador') or r.get('Fuente costo energía') or '')
                   for r in filas]
        es_manual = (modo == 'Manual / escenario' or
                     any('manual' in f.lower() for f in fuentes))
        etiqueta = 'Dólar aplicado (manual)' if es_manual else 'Dólar promedio'
        st.metric(etiqueta, _es(dolar, 2) + ' $/US$' if dolar is not None else '—')
        if dolar is None:
            st.caption('Sin cotización disponible para el período')
    with e:
        precio = precios['final_clp_kwh']
        st.metric('Precio CPI + dólar', _es(precio, 4) + ' $/kWh' if precio is not None else '—')


def _mostrar_total_cobros(tarjetas, sitios, consumo_kwh=None,
                          consumos_por_sitio=None, info=None):
    """Total de cobros, sin duplicar indicadores contractuales de cabecera."""
    balance = totalizar_cobros(tarjetas, sitios)
    total = balance['subtotal']
    etiqueta = 'Total estimado' if balance['completo'] else 'Subtotal parcial'
    consumos_por_sitio = consumos_por_sitio or {}

    if len(sitios) == 2:
        columnas = st.columns([2, 1, 1])
        with columnas[0]:
            st.metric(etiqueta, '$' + _es(total) if total is not None else '—')
        for col, sitio in zip(columnas[1:], sitios):
            dato = balance['por_sitio'][sitio]
            monto = dato['subtotal']
            with col:
                st.metric(sitio, '$' + _es(monto) if monto is not None else '—')
    else:
        st.metric(etiqueta, '$' + _es(total) if total is not None else '—')

    if balance['pendientes']:
        st.caption(f'{len(balance["pendientes"])} cargo(s) pendiente(s) · subtotal parcial')
    return balance

def _tarjeta_ui(tarjeta, sitios):
    """Tarjeta numérica: concepto, importe y precio unitario aplicable."""
    nombre = tarjeta['nombre']
    etiquetas = {
        'Energía contractual': 'Energía contractual',
        'Total Archivo 9 · pagos por retiro': 'Archivo 9 · total',
        'Exenciones CI · Liquidación de peaje': 'Liquidación de peaje CI',
        'SSCC + CO · Archivo 4 (neto estimado)': 'Archivo 4 · SSCC + CO (est.)',
        'Liquidación de peaje CI · Proforma': 'Peaje proforma (ref.)',
        'Servicio Público (CSP/FET)': 'Servicio público CSP/FET',
    }
    monto = _total_tarjeta(tarjeta, sitios)
    dos_decimales = nombre == 'Total Archivo 9 · pagos por retiro'
    with st.container(border=True):
        st.metric(etiquetas.get(nombre, nombre),
                  '$' + _es(monto, 2 if dos_decimales else 0) if monto is not None else '—')
        # La tarifa NO se suma entre suministros; se muestra individualmente
        # cuando difiere. No se exponen archivos, fórmulas ni estados técnicos.
        tarifas = []
        for sitio in sitios:
            dato = tarjeta['por_sitio'].get(sitio, {})
            tasa = _num(dato.get('tarifa_clp_kwh'))
            if tasa is not None:
                prefijo = f'{sitio}: ' if len(sitios) > 1 else ''
                tipo = str(dato.get('tipo_tarifa') or '')
                # Las tasas SEN referenciales NO se presentan como costos efectivos.
                if 'estimad' in tipo.lower():
                    sufijo = ' estimado'
                elif 'referencial' in tipo.lower():
                    sufijo = ' referencial'
                elif 'efectivo' in tipo.lower():
                    sufijo = ' efectivo'
                else:
                    sufijo = ''
                tarifas.append(f'{prefijo}{_es(tasa, 5)} $/kWh{sufijo}')
        if tarifas:
            st.caption('  ·  '.join(tarifas))
        if nombre == 'Potencia PSUF':
            potencias = []
            for sitio in sitios:
                dato = tarjeta['por_sitio'].get(sitio, {})
                mw = _num(dato.get('potencia_mw'))
                precio = _num(dato.get('tarifa_clp_kw_mes'))
                if mw is not None or precio is not None:
                    prefijo = f'{sitio}: ' if len(sitios) > 1 else ''
                    medicion = []
                    if mw is not None:
                        medicion.append(_es(mw, 3) + ' MW')
                    if precio is not None:
                        medicion.append(_es(precio, 3) + ' $/kW/mes')
                    potencias.append(prefijo + ' · '.join(medicion))
            if potencias:
                st.caption('  ·  '.join(potencias))

def _mostrar_sscc_archivo_9(tarjetas, sitios):
    """Compatibilidad: mismo importe único, sin una segunda tarjeta."""
    tarjeta = next((t for t in tarjetas if t.get('nombre') ==
                    'Total Archivo 9 · pagos por retiro'), None)
    if tarjeta:
        _tarjeta_ui(tarjeta, sitios)

def mostrar_tarjetas_costos(info, catalogo, balances=None, sscc_reports=None,
                           potencia_reports=None, exenciones_ci_reports=None,
                           mt_pd_directos=None, sscc_companies=None,
                           reliquidaciones=None, consumo_kwh=None,
                           consumos_por_sitio=None):
    """Cuadrícula ejecutiva de importes; auditoría queda en módulos técnicos."""
    tarjetas = desglose_todos_los_costos(
        info, catalogo, balances, sscc_reports, potencia_reports,
        exenciones_ci_reports, mt_pd_directos,
        sscc_companies=sscc_companies, reliquidaciones=reliquidaciones)
    sitios = [r['Suministro'] for r in info]
    _mostrar_total_cobros(tarjetas, sitios, consumo_kwh=consumo_kwh,
                         consumos_por_sitio=consumos_por_sitio, info=info)
    st.markdown('#### Costos por concepto')
    # Cargos efectivos del tablero; el Archivo 9 se muestra exactamente una vez.
    for grupo in ('Energía', 'Regulados CNE', 'CEN mensual'):
        items = [r for r in tarjetas if r['grupo'] == grupo]
        for inicio in range(0, len(items), 3):
            columnas = st.columns(min(3, len(items) - inicio))
            for columna, item in zip(columnas, items[inicio:inicio + 3]):
                with columna:
                    _tarjeta_ui(item, sitios)
    # No sumar referenciales al total general; separar visualmente.
    referencias = [r for r in tarjetas if r['grupo'] == 'Referencias CEN']
    if referencias:
        st.markdown('#### Referenciales · no incluidos en el total')
        for inicio in range(0, len(referencias), 3):
            columnas = st.columns(min(3, len(referencias) - inicio))
            for columna, item in zip(columnas, referencias[inicio:inicio + 3]):
                with columna:
                    _tarjeta_ui(item, sitios)
    return tarjetas

# Solo se redibuja este panel al cambiar mes, año o suministro.
# No vuelve a ejecutar app.py ni a procesar otras pestañas.
@st.fragment
def mostrar_panel_ejecutivo(balances=None, archivos_repositorio=None,
                           sscc_reports=None, potencia_reports=None,
                           exenciones_ci_reports=None, mt_pd_directos=None,
                           sscc_companies=None, cargar_reliquidacion=None):
    st.subheader('Resumen')
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
    filtros = st.columns([1, 1, 2])
    with filtros[0]:
        year = st.selectbox('Año', years, index=years.index(suggested_y),
                            key='panel_periodo_ano')
    with filtros[1]:
        month = st.selectbox('Mes', list(range(1, 13)), index=suggested_m-1,
                             format_func=lambda x: MESES[x-1], key='panel_periodo_mes')
    with filtros[2]:
        option = st.radio('Suministro', ('Ambos', *SITIOS), horizontal=True,
                          key='panel_suministro_vista_v2', index=0)
    selected = SITIOS if option == 'Ambos' else (option,)
    mes_factura = f'{year:04d}-{month:02d}'
    mes_reliquidacion = _mes_anterior(mes_factura)
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
    info=_completar_dolar_resumen(info)

    # Primero mostrar los indicadores solicitados. No dependen del Archivo 4, 9,
    # MT/PD ni de que el desglose de CEN se ejecute correctamente.
    _mostrar_cabecera_energia(info)

    # Solo consumo total; los importes de energía y los demás cobros
    # aparecen una vez en la cuadrícula inferior.
    completos_kwh = bool(info) and all(_num(r.get('Total energía (kWh)')) is not None
                                       for r in info)
    total_kwh = sum(r['Total energía (kWh)'] for r in info) if completos_kwh else None

    # Consultar solamente los archivos del balance CEN anterior. La carga
    # independiente permite factura de octubre con filtro global solo octubre.
    reliquidaciones=None
    if callable(cargar_reliquidacion):
        try:
            with st.spinner('Consultando cargos CEN del mes anterior…'):
                reliquidaciones=cargar_reliquidacion(mes_reliquidacion)
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
            reliquidaciones=reliquidaciones, consumo_kwh=total_kwh,
            consumos_por_sitio={r['Suministro']:r.get('Total energía (kWh)')
                               for r in info})
    except Exception as exc:
        import logging
        logging.exception('EFE_PANEL_EJECUTIVO: error en tarjetas')
        st.warning('No se pudieron representar las tarjetas CEN, pero el resumen principal '
                   'continúa disponible.')
        with st.expander('Diagnóstico de tarjetas de costos'):
            st.exception(exc)


    # El CSV sigue disponible sin convertir el Resumen en una pantalla de auditoría.
    st.download_button(
        'Descargar resumen CSV', data=_csv(info),
        file_name=f"EFE_resumen_{'Ambos' if option=='Ambos' else 'Quilpue' if option=='Quilpué' else 'Limache'}_{year}_{month:02d}.csv",
        mime='text/csv', key='panel_resumen_csv')
