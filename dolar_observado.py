"""Dólar observado CLP/USD con filtros propios, independientes del CEN.

Valores históricos reproducidos de Banco Central de Chile por indicadores.cl
(sin API key). No se usa el tipo de cambio para transformar cargos CEN.
"""
from __future__ import annotations

from datetime import datetime, date
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo
import json
import math

import plotly.graph_objects as go
import streamlit as st

_API = "https://indicadores.cl/api/v1/historico/dolar/{anio}.json"
_SANTIAGO = ZoneInfo("America/Santiago")
_MESES = ("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic")


def _valor_num(v):
    if isinstance(v, bool):
        raise ValueError("Valor no numérico")
    if isinstance(v, str):
        s = v.strip().replace("$", "").replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        v = s
    n = float(v)
    if not math.isfinite(n) or not (0 < n < 10000):
        raise ValueError("Tipo de cambio fuera de rango")
    return n


def _fecha_chile(raw):
    # La fuente principal entrega YYYY-MM-DD (día local de Chile).
    # Soporta también datetimes UTC de proveedores secundarios.
    s = str(raw)
    if "T" in s:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        if dt.tzinfo is not None:
            dt = dt.astimezone(_SANTIAGO)
        return dt.date()
    return date.fromisoformat(s[:10])


def _interpretar_respuesta(payload, anio):
    if not isinstance(payload, dict) or not isinstance(payload.get("serie"), list):
        raise ValueError("El servicio no devolvió una serie de dólar observado válida")
    datos = {}
    for fila in payload["serie"]:
        try:
            f = _fecha_chile(fila["fecha"])
            v = _valor_num(fila["valor"])
        except (ValueError, TypeError, KeyError, OverflowError):
            continue
        if f.year == anio:
            datos[f] = v
    if not datos:
        raise ValueError(f"No hay observaciones de dólar observado para {anio}")
    return tuple((f.isoformat(), datos[f]) for f in sorted(datos))


@st.cache_data(ttl=3600, show_spinner=False, max_entries=50)
def _serie_anual(anio):
    """Serie publicada por año. Solo se cachean respuestas válidas; nunca cifras ficticias."""
    anio = int(anio)
    hoy = datetime.now(_SANTIAGO).date()
    # Fallback: los endpoints /dolar.json y /historico/dolar.json mantienen
    # la misma serie si falla el archivo individual del año.
    urls = [_API.format(anio=anio)]
    if anio == hoy.year:
        urls.append("https://indicadores.cl/api/v1/dolar.json")
    urls.append("https://indicadores.cl/api/v1/historico/dolar.json")
    errors = []
    for url in urls:
        try:
            request = Request(url, headers={"User-Agent": "EFEValparaiso-DolarHistorico/1.0", "Accept": "application/json"})
            with urlopen(request, timeout=20) as response:
                obj = json.loads(response.read())
            return _interpretar_respuesta(obj, anio)
        except Exception as exc:
            errors.append(f"{url.rsplit('/', 1)[-1]}: {exc}")
    raise ValueError("No fue posible consultar las cotizaciones publicadas: " + "; ".join(errors))


def _construir_tablas(series_por_anio, anios_seleccion, meses_seleccion, hoy=None):
    """Devuelve registros filtrados y resúmenes mensuales/anuales sin rellenar días no publicados."""
    hoy = hoy or datetime.now(_SANTIAGO).date()
    years = set(map(int, anios_seleccion))
    months = set(map(int, meses_seleccion))
    diaria = []
    for year in sorted(years):
        for fecha_iso, valor in series_por_anio.get(year, ()):
            f = date.fromisoformat(fecha_iso)
            if f.year == year and f.month in months and f <= hoy:
                diaria.append({"Fecha": f, "Año": f.year, "Mes": f.month,
                               "Dólar observado ($/USD)": float(valor)})
    diaria.sort(key=lambda fila: fila["Fecha"])
    mensual, anual = [], []
    for year in sorted(years):
        for mes in sorted(months):
            items = [r for r in diaria if r["Año"] == year and r["Mes"] == mes]
            if not items:
                continue
            vals = [r["Dólar observado ($/USD)"] for r in items]
            mensual.append({
                "Año": year, "Mes": mes, "Período": f"{_MESES[mes - 1]} {year}",
                "Promedio ($/USD)": sum(vals) / len(vals),
                "Primer valor ($/USD)": vals[0], "Último valor ($/USD)": vals[-1],
                "Mínimo ($/USD)": min(vals), "Máximo ($/USD)": max(vals),
                "Cotizaciones": len(vals), "Última fecha": items[-1]["Fecha"],
                "Estado": "Parcial" if (year, mes) >= (hoy.year, hoy.month) else "Cerrado"})
        year_rows = [r for r in diaria if r["Año"] == year]
        if year_rows:
            vals = [r["Dólar observado ($/USD)"] for r in year_rows]
            anual.append({"Año": year, "Promedio período seleccionado ($/USD)": sum(vals) / len(vals),
                          "Cotizaciones": len(vals), "Meses con datos": len({r["Mes"] for r in year_rows}),
                          "Observado hasta": year_rows[-1]["Fecha"]})
    return diaria, mensual, anual


def _formato_es(num):
    return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _csv_bytes(rows):
    # BOM para que Excel de Chile interprete acentos y separadores correctamente.
    import csv
    from io import StringIO
    if not rows:
        return b""
    buffer = StringIO()
    w = csv.DictWriter(buffer, fieldnames=list(rows[0]), delimiter=";")
    w.writeheader()
    for row in rows:
        w.writerow({k: (v.isoformat() if isinstance(v, date) else
                         (str(v).replace(".", ",") if isinstance(v, float) else v))
                    for k, v in row.items()})
    return ("\ufeff" + buffer.getvalue()).encode("utf-8")


def _filtros_dolar_independientes(hoy=None):
    """Selecciona años/meses exclusivamente del dólar, aunque no existan archivos CEN."""
    hoy = hoy or datetime.now(_SANTIAGO).date()
    with st.container(border=True):
        st.markdown("#### 📅 Períodos del dólar observado")
        st.caption("Estos filtros son **independientes** de Análisis energético y de los "
                   "archivos cargados. Puedes consultar otros meses y años sin modificar el CEN.")
        ca, cm = st.columns(2)
        with ca:
            years = st.multiselect(
                "Años del dólar (uno o varios)",
                options=list(range(2000, hoy.year + 1)),
                default=[hoy.year],
                key="dolar_ind_anios",
            )
        with cm:
            all_months = st.checkbox(
                "Todos los meses (año completo)", value=False,
                key="dolar_ind_todos_meses",
            )
            if all_months:
                months = list(range(1, 13))
                st.caption("Meses: enero a diciembre")
            else:
                months = st.multiselect(
                    "Meses del dólar (uno o varios)",
                    options=list(range(1, 13)),
                    default=[hoy.month],
                    format_func=lambda mes: _MESES[mes - 1],
                    key="dolar_ind_meses",
                )
        st.caption("Período seleccionado: " +
                   (", ".join(map(str, sorted(years))) if years else "sin años") + " · " +
                   (", ".join(_MESES[m-1] for m in sorted(months)) if months else "sin meses") + ".")
    return sorted(set(years)), sorted(set(months))


def mostrar_dolar_observado():
    """Pestaña con controles propios: no recibe períodos del módulo CEN."""
    st.subheader("💵 Dólar observado por mes y año")
    st.caption("Referencia: dólar observado CLP por US$1, Banco Central de Chile, "
               "a través de un archivo histórico público cotejado con esa serie. "
               "Se actualiza automáticamente (caché de hasta 1 hora).")
    years, months = _filtros_dolar_independientes()
    if not years or not months:
        st.warning("Selecciona al menos un año y un mes en los filtros de esta pestaña.")
        return
    if st.button("🔄 Actualizar cotizaciones", key="dolar_refrescar"):
        _serie_anual.clear()
        st.rerun()
    series, failed = {}, []
    with st.spinner("Consultando dólar observado histórico…"):
        for year in years:
            try:
                series[year] = _serie_anual(year)
            except Exception as exc:
                failed.append(f"{year}: {exc}")
    if failed:
        st.warning("No se pudieron consultar estos años: " + "; ".join(failed) +
                   ". No se sustituyen por valores estimados ni por datos de otros años.")
    diaria, mensual, anual = _construir_tablas(series, years, months)
    if not diaria:
        st.warning("Sin cotizaciones publicadas para la selección actual. "
                   "Si es un período futuro, no hay información disponible todavía.")
        return
    ult = diaria[-1]
    promedio = sum(r["Dólar observado ($/USD)"] for r in diaria) / len(diaria)
    ultimo_mes = next((m for m in mensual if (m["Año"], m["Mes"]) ==
                       (ult["Año"], ult["Mes"])), None)
    met = st.columns(4)
    met[0].metric("Última cotización seleccionada", f"${_formato_es(ult['Dólar observado ($/USD)'])}",
                  f"{ult['Fecha']:%d-%m-%Y}")
    met[1].metric("Promedio del último mes seleccionado",
                  f"${_formato_es(ultimo_mes['Promedio ($/USD)'])}" if ultimo_mes else "—",
                  ultimo_mes["Período"] + (" · parcial" if ultimo_mes["Estado"] == "Parcial" else ""))
    met[2].metric("Promedio de todos los períodos seleccionados", f"${_formato_es(promedio)}")
    met[3].metric("Días con cotización", str(len(diaria)))
    if len(failed) > 0:
        st.caption("Los promedios anteriores solo incluyen los años descargados correctamente.")
    st.caption("**Método:** promedio aritmético de las cotizaciones efectivamente publicadas en "
               "días hábiles; no se repiten cotizaciones en sábados, domingos o festivos. "
               "El promedio multimes es ponderado implícitamente por el número de días cotizados, "
               "no por el número de meses.")
    st.markdown("#### Promedio del dólar por mes")
    mostrar_mes = [{"Período": r["Período"], "Promedio ($/USD)": round(r["Promedio ($/USD)"], 2),
                    "Último día publicado": r["Última fecha"].strftime("%d-%m-%Y"),
                    "Último valor ($/USD)": round(r["Último valor ($/USD)"], 2),
                    "Mínimo ($/USD)": round(r["Mínimo ($/USD)"], 2),
                    "Máximo ($/USD)": round(r["Máximo ($/USD)"], 2),
                    "Días cotizados": r["Cotizaciones"], "Estado": r["Estado"]}
                   for r in mensual]
    st.dataframe(mostrar_mes, use_container_width=True, hide_index=True)
    if len(mensual) > 1:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=[r["Período"] for r in mensual],
                             y=[r["Promedio ($/USD)"] for r in mensual], name="Promedio mensual"))
        fig.update_layout(title="Evolución del dólar observado promedio", xaxis_title="Mes",
                          yaxis_title="CLP/USD", height=350)
        st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})
    if len(years) > 1 or len(months) > 1:
        st.markdown("#### Resumen por año (solo meses seleccionados)")
        st.dataframe([{"Año": r["Año"],
                       "Promedio ($/USD)": round(r["Promedio período seleccionado ($/USD)"], 2),
                       "Meses con datos": r["Meses con datos"],
                       "Días cotizados": r["Cotizaciones"]} for r in anual],
                     use_container_width=True, hide_index=True)
    st.markdown("#### Dólar observado por día")
    if len(mensual) > 1:
        pick = st.selectbox("Mostrar detalle diario", ["Todos los meses"] +
                            [f"{r['Período']}" for r in mensual], key="dolar_detalle_mes")
        if pick != "Todos los meses":
            row = next(r for r in mensual if r["Período"] == pick)
            detalle = [r for r in diaria if (r["Año"], r["Mes"]) == (row["Año"], row["Mes"])]
        else:
            detalle = diaria
    else:
        detalle = diaria
    st.dataframe([{"Fecha": r["Fecha"].strftime("%d-%m-%Y"),
                   "Dólar observado ($/USD)": round(r["Dólar observado ($/USD)"], 2)}
                  for r in reversed(detalle)], hide_index=True, use_container_width=True)
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("⬇ Descargar detalle diario CSV", _csv_bytes(diaria),
                           "Dolar_observado_diario_EFE.csv", "text/csv", key="dolar_csv_dia")
    with c2:
        st.download_button("⬇ Descargar promedios mensuales CSV", _csv_bytes(mensual),
                           "Dolar_observado_mensual_EFE.csv", "text/csv", key="dolar_csv_mes")
    st.caption("Datos: https://indicadores.cl/api/ (serie cotejada con Banco Central de Chile). "
               "Para la liquidación con Enel, verificar en el contrato si corresponde utilizar "
               "dólar observado de un día específico o promedio mensual; la pestaña no altera cargos CEN.")
