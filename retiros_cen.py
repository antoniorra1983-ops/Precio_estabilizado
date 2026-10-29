"""Consulta rápida de retiros medidos del Balance Valorizado del CEN.

No inventa consumos; no usa la proforma como fuente de kWh.
El tipo de medidor se lee de columna J: L, L_D, L_S únicamente.
"""
from __future__ import annotations

import csv
from io import StringIO
from typing import Callable

TIPOS_RETIRO = ("L", "L_D", "L_S")
PUNTOS = {"Quilpué": "MERVAL_ELSOL", "Limache": "MERVAL_LIMACHE"}


def datos_retiros(balance: object) -> dict:
    """Concilia medidores EFE con agregado Enel; todas las cifras en kWh.

    Enel: total de la columna L ENEL_GENERACION en Balance Valorizado,
    únicamente para tipos de medidor J=L/L_D/L_S (calculado por el parser).
    EFE: filas individuales de esas mismas claves y tipos. Se preserva el
    signo de rectificaciones del CEN y se excluyen filas de otro suministrador.
    """
    enel_mwh = getattr(balance, "retiros_enel_por_tipo", {}) or {}
    por_punto = {p: {t: 0.0 for t in TIPOS_RETIRO} for p in PUNTOS}
    encontrados = {p: set() for p in PUNTOS}
    auditoria = []
    for r in getattr(balance, "detalle", []) or []:
        tipo = str(r.get("tipo", "")).upper().strip()
        clave = str(r.get("clave", "")).upper().strip()
        if tipo not in TIPOS_RETIRO:
            continue
        sitio = next((p for p, k in PUNTOS.items() if clave == k), None)
        if not sitio:
            continue
        valor = float(r["kwh"])
        por_punto[sitio][tipo] += valor
        encontrados[sitio].add(tipo)
        auditoria.append({
            "Fila Balance Valorizado": r.get("fila"),
            "Suministro": sitio,
            "Clave": clave,
            "Tipo medidor": tipo,
            "Retiro (kWh)": valor,
            "Barra": r.get("barra"),
            "Fuente": getattr(balance, "archivo", "")
        })
    datos = {}
    for punto in (*PUNTOS, "Total EFE", "Total Enel"):
        datos[punto] = {}
        for tipo in TIPOS_RETIRO:
            if punto in PUNTOS:
                datos[punto][tipo] = por_punto[punto][tipo] if tipo in encontrados[punto] else None
            elif punto == "Total EFE":
                presentes = [por_punto[p][tipo] for p in PUNTOS if tipo in encontrados[p]]
                datos[punto][tipo] = sum(presentes) if presentes else None
            else:
                dato = enel_mwh.get(tipo)
                datos[punto][tipo] = float(dato) * 1000 if dato is not None else None
        if punto in PUNTOS and not encontrados[punto]:
            datos[punto]["TOTAL"] = None
        else:
            datos[punto]["TOTAL"] = sum(v for k, v in datos[punto].items()
                                          if k in TIPOS_RETIRO and v is not None)
    if not any(encontrados.values()):
        datos["Total EFE"]["TOTAL"] = None
    enel_total = datos["Total Enel"]["TOTAL"]
    for punto, fila in datos.items():
        fila["MWH"] = fila["TOTAL"] / 1000 if fila["TOTAL"] is not None else None
        fila["PARTICIPACION"] = (fila["TOTAL"] / enel_total * 100
                                 if fila["TOTAL"] is not None and enel_total and enel_total > 0
                                 else None)
    return {
        "mes_balance": getattr(balance, "label", ""),
        "archivo": getattr(balance, "archivo", ""),
        "datos": datos,
        "auditoria": auditoria,
        "avisos": list(getattr(balance, "avisos", []) or []),
    }


def _csv(datos: list[dict]) -> bytes:
    salida = StringIO()
    if datos:
        writer = csv.DictWriter(salida, fieldnames=list(datos[0]), delimiter=';')
        writer.writeheader()
        for item in datos:
            writer.writerow({k: str(v).replace('.', ',') if isinstance(v, float)
                             else v for k, v in item.items()})
    return ('\ufeff' + salida.getvalue()).encode('utf-8')


def mostrar_retiros_cen(archivos: list, obtener_balance: Callable, tipos=TIPOS_RETIRO):
    """Fragmento: solo recalcula la pestaña Retiros al cambiar sus filtros."""
    import streamlit as st

    @st.fragment
    def _pantalla():
        st.subheader("Retiros CEN · EFE Valparaíso y Enel Generación")
        st.caption("Fuente única: Balance Valorizado de cada Balance CEN. "
                   "Tipo de medidor (columna J): L, L_D y L_S. "
                   "El tipo R y los demás quedan excluidos de estos consumos.")
        periodos = sorted({_periodo(a.name) for a in archivos if _periodo(a.name)})
        if not periodos:
            st.warning("No hay archivos Balance CEN disponibles. Agrega archivos en "
                       "Datos/Balances/AAAA/ en GitHub o cárgalos en Datos y ajustes.")
            return
        años = sorted({int(p[:4]) for p in periodos})
        col1,col2,col3=st.columns(3)
        with col1:
            año=st.selectbox("Año del balance", años, index=len(años)-1,
                              key='retiros_anio_v1')
        meses=sorted({int(p[5:7]) for p in periodos if p.startswith(f'{año}-')})
        MESES=("Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
        with col2:
            mes=st.selectbox("Mes del balance", meses, index=len(meses)-1,
                              format_func=lambda m: MESES[m-1], key='retiros_mes_v1')
        with col3:
            suministro=st.selectbox("Suministro", ("Ambos", "Quilpué", "Limache"),
                                     key='retiros_suministro_v1')
        periodo=f'{año:04d}-{mes:02d}'
        archivo=next((a for a in archivos if _periodo(a.name)==periodo), None)
        if archivo is None:
            st.info("No se encontró Balance para el período seleccionado.")
            return
        try:
            with st.spinner("Consultando el balance del mes seleccionado..."):
                resumen=datos_retiros(obtener_balance(archivo))
        except Exception as exc:
            st.error(f"No se pudo procesar {archivo.name}: {type(exc).__name__}: {exc}")
            return
        fuente=resumen['archivo'] or archivo.name
        st.caption(f"**Archivo:** {fuente} · **Balance:** {periodo} · "
                   "**Facturación de reliquidaciones:** mes siguiente")
        info=resumen['datos']
        puntos=["Quilpué", "Limache", "Total EFE", "Total Enel"] if suministro=='Ambos' else [suministro, "Total Enel"]
        def mostrar(x, dec=2):
            return ('—' if x is None else f'{x:,.{dec}f}'.replace(',','X').replace('.',',').replace('X','.'))
        c1,c2,c3,c4=st.columns(4)
        c1.metric('Quilpué (MWh)',mostrar(info['Quilpué']['MWH'],3))
        c2.metric('Limache (MWh)',mostrar(info['Limache']['MWH'],3))
        c3.metric('EFE total (MWh)',mostrar(info['Total EFE']['MWH'],3))
        c4.metric('Enel total (MWh)',mostrar(info['Total Enel']['MWH'],3))
        filas=[]
        for p in puntos:
            dato=info[p]
            filas.append({"Suministro":p,
                          "L (kWh)":dato['L'],
                          "L_D (kWh)":dato['L_D'],
                          "L_S (kWh)":dato['L_S'],
                          "TOTAL (kWh)":dato['TOTAL'],
                          "TOTAL (MWh)":dato['MWH'],
                          "% retiros Enel":dato['PARTICIPACION']})
        st.markdown('#### Desglose por tipo de medidor')
        st.dataframe(filas,hide_index=True,use_container_width=True)
        st.caption('Un guion indica que no se identificó un registro del tipo indicado; '
                   'no se reemplaza con cero. Los montos son retiros netos con el signo registrado en el balance.')
        st.download_button('Descargar retiros del mes (CSV)',_csv(filas),
                           f'Retiros_CEN_{periodo}.csv','text/csv',key='retiros_csv_v1')
        with st.expander('Auditoría: filas utilizadas en Balance Valorizado'):
            detalles=[r for r in resumen['auditoria']
                      if suministro=='Ambos' or r['Suministro']==suministro]
            if detalles:
                st.dataframe(detalles,hide_index=True,use_container_width=True)
                st.download_button('Descargar filas auditables (CSV)',_csv(detalles),
                                   f'Retiros_CEN_detalle_{periodo}.csv','text/csv',key='retiros_detalle_csv')
            else:
                st.info('No hay filas para las claves de este suministro y los tipos L / L_D / L_S.')
        with st.expander('Advertencias de lectura del Balance CEN'):
            for aviso in resumen['avisos']:
                st.write('• '+str(aviso))
    _pantalla()


def _periodo(name: str) -> str | None:
    import re
    name = str(name).upper()
    m = re.search(r'(?<!\d)(\d{2})(0[1-9]|1[0-2])[PD]?(?!\d)', name)
    if m:
        return f'20{m.group(1)}-{m.group(2)}'
    m = re.search(r'(20\d{2})[-_ ](0[1-9]|1[0-2])(?:[^0-9]|$)',name)
    return f'{m.group(1)}-{m.group(2)}' if m else None
