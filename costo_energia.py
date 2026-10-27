"""Costo de energía activa EFE: Quilpué (SEAT) y Limache (SAF LI).

Contrato mostrado por usuario: PE0 = 37,8 US$/MWh (día=noche),
reajuste cada enero/julio por CPI-U CUUR0000SA0 (no desestacionalizado):
  PE = 37,8 * CPI_octubre_anterior / 271,696  [enero-junio]
  PE = 37,8 * CPI_abril_mismo_año / 271,696 [julio-diciembre]
El CPI de octubre 2025 NO se publicó: no se imputa automáticamente.
USD/CLP: dólar observado promedio del mes de consumo (escenario opcional por factura).
La energía SEAT y SAF LI se tratan como aproximaciones de suministros, NO
necesariamente equivalentes a retiros contractuales MERVAL_ELSOL/MERVAL_LIMACHE.
"""
from __future__ import annotations

from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from io import BytesIO, StringIO
from urllib.request import Request, urlopen
import csv
import re

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from proyeccion_energia import (catalogo, descargar_archivo, leer_seat, leer_umr,
                                por_dia, periodo, tipo_dia, fmt, MESES)
from dolar_observado import _serie_anual as dolar_serie_anual

PRECIO_BASE_USD_MWH = 37.8
CPI_BASE = 271.696  # junio 2021, indicado explícitamente en contrato
CPI_SERIE = "CUUR0000SA0"
# Respaldo de observaciones históricas publicadas (FRED CPIAUCNS/BLS CUUR0000SA0).
# Octubre de 2025 intencionalmente NO existe: cierre federal de EE.UU.
CPI_CONFIRMADO = {
    (2021, 4): 267.054, (2021, 10): 276.589,
    (2022, 4): 289.109, (2022, 10): 298.012,
    (2023, 4): 303.363, (2023, 10): 307.671,
    (2024, 4): 313.548, (2024, 10): 315.664,
    (2025, 4): 320.795,
    (2026, 4): 333.020,
}


def cpi_referencia(anio: int, mes: int) -> tuple[int, int]:
    """Para cualquier mes, obtener el semestre del reajuste contractual."""
    if 1 <= mes <= 6:
        return anio - 1, 10
    if 7 <= mes <= 12:
        return anio, 4
    raise ValueError("Mes inválido")


def precio_mwh(cpi: float) -> float:
    if cpi is None or not np.isfinite(cpi) or cpi <= 0:
        raise ValueError("CPI no disponible")
    return PRECIO_BASE_USD_MWH * float(cpi) / CPI_BASE


def _mes_siguiente(y, m):
    return (y + 1, 1) if m == 12 else (y, m + 1)


@st.cache_data(ttl=86400, max_entries=5, show_spinner=False)
def _serie_cpi_fred():
    """FRED CPIAUCNS: misma serie BLS especificada en el contrato, sin API key."""
    sources = [
        "https://fred.stlouisfed.org/graph/fredgraph.csv?id=CPIAUCNS",
        "https://fred.stlouisfed.org/graph/fredgraph.csv?cosd=2021-01-01&id=CPIAUCNS",
    ]
    errors = []
    for url in sources:
        try:
            req = Request(url, headers={"User-Agent": "Mozilla/5.0 EFE-Valparaiso-CPI/1.0",
                                        "Accept": "text/csv, text/plain, */*"})
            with urlopen(req, timeout=16) as response:
                content = response.read().decode("utf-8-sig")
            rows = list(csv.DictReader(StringIO(content)))
            datos = {}
            for row in rows:
                try:
                    fecha = date.fromisoformat(str(row.get("DATE") or row.get("observation_date") or "")[:10])
                    val = float(row.get("CPIAUCNS") or "")
                    if 1900 <= fecha.year <= 2100 and np.isfinite(val) and val > 0:
                        datos[(fecha.year, fecha.month)] = val
                except (ValueError, TypeError):
                    continue
            if (2021, 6) not in datos or abs(datos[(2021, 6)] - CPI_BASE) > .01:
                raise ValueError("La serie recibida no corresponde al CPI-U no ajustado; faltó la base de 2021")
            return datos
        except Exception as exc:
            errors.append(str(exc))
    raise RuntimeError(" / ".join(errors)[:600])


def indices_cpi():
    """Devuelve referencias verificadas y, si se puede, la serie íntegra publicada."""
    datos = dict(CPI_CONFIRMADO)
    estado = "Referencias BLS/FRED verificadas, incluidas en la aplicación"
    try:
        live = _serie_cpi_fred()
        datos.update(live)
        estado = "Serie histórica actualizada desde FRED (BLS CPIAUCNS)"
    except Exception as exc:
        estado += f" (consulta en línea no disponible: {str(exc)[:130]})"
    # Sin CPI de octubre 2025: aun si una API lo rellenara incorrectamente, no aceptarlo.
    datos.pop((2025, 10), None)
    return datos, estado


@st.cache_data(ttl=3600, max_entries=120, show_spinner=False)
def _parse_operacional_cached(kind, nombre, payload):
    return leer_seat(payload, nombre) if kind == "seat" else leer_umr(payload, nombre)


def _datos_operacionales(anos, meses, usar_repo, subida_seat, subida_umr):
    """Carga tanto SEAT como UMR; evita duplicar el mismo mes y las fechas repetidas."""
    found = {"seat": {}, "umr": {}}
    warnings = []
    if usar_repo:
        for kind in found:
            try:
                entries = catalogo(kind)
                for f in entries:
                    y, m = f["period"]
                    if y in anos and m in meses:
                        found[kind][(y, m)] = (f["name"], f["url"], None)
            except Exception as exc:
                warnings.append(f"Catálogo {kind} de GitHub: {exc}")
    for kind, files in (("seat", subida_seat), ("umr", subida_umr)):
        for f in files or []:
            y, m = periodo(f.name)
            if y in anos and m in meses:
                found[kind][(y, m)] = (f.name, None, f.getvalue())
            else:
                warnings.append(f"Se omitió '{f.name}': el período no está seleccionado o no se identifica")
    seat_rows, saf_rows, od_rows, kms_rows = [], [], [], []
    jobs = [(kind, *t) for kind, dic in found.items() for t in dic.values()]
    def read(j):
        kind, name, url, data = j
        raw = data if data is not None else descargar_archivo(url)
        return kind, name, _parse_operacional_cached(kind, name, raw)
    with ThreadPoolExecutor(max_workers=4) as pool:
        pending = {pool.submit(read, j): j for j in jobs}
        for future in as_completed(pending):
            try:
                kind, name, pair = future.result()
                for rr in pair[0] + pair[1]:
                    rr["_archivo"] = name
                if kind == "seat":
                    seat_rows.extend(pair[0]); saf_rows.extend(pair[1])
                else:
                    od_rows.extend(pair[0]); kms_rows.extend(pair[1])
            except Exception as exc:
                warnings.append(f"Error {pending[future][1]}: {exc}")
    dfs = [por_dia(rows, anos, meses) for rows in (seat_rows, saf_rows, od_rows, kms_rows)]
    return dfs, len(found["seat"]), len(found["umr"]), warnings


def _perfil_quilpue(seat, odom, kms, umr):
    """Modelo de Proyección de energía SEAT: tracción + 12 kV."""
    if seat.empty or odom.empty or kms.empty:
        return pd.DataFrame(), "Se necesitan SEAT y UMR (odómetro y KM-Servicio)."
    data = kms.merge(odom, on="Fecha", how="inner").merge(seat, on="Fecha", how="inner")
    req = (data["TrenKmP"] > 0) & (data["Odom"] > 0) & ((data["Traccion"] + data["E12"]) > 0)
    data = data[req].copy()
    if data.empty:
        return pd.DataFrame(), "No coinciden días SEAT con odómetro y kilómetros de servicios UMR."
    data["Tipo"] = data["Fecha"].map(tipo_dia)
    data["IDE_tr"] = data["Traccion"] / data["Odom"]
    profile = data.groupby("Tipo").agg(Dias=("Fecha", "count"), KmP=("TrenKmP", "median"),
                                        IDE=("IDE_tr", "median"), E12=("E12", "median")).reset_index()
    profile["UMR"] = profile["Tipo"].map(umr)
    profile["kWh_dia"] = profile["IDE"] * profile["KmP"] / profile["UMR"] + profile["E12"]
    real = data.groupby("Fecha")[['Traccion', 'E12']].sum()
    real["kWh_dia"] = real["Traccion"] + real["E12"]
    return (profile, real), None


def _perfil_limache(saf):
    if saf.empty:
        return pd.DataFrame(), "No hay hoja SAFs de Limache con consumo válido."
    data = saf[(saf["Taller"] + saf["TGSN"]) > 0].copy()
    if data.empty:
        return pd.DataFrame(), "No hay días con consumos SAF LI positivos."
    data["Tipo"] = data["Fecha"].map(tipo_dia)
    profile = data.groupby("Tipo").agg(Dias=("Fecha", "count"), Taller=("Taller", "median"),
                                        TGSN=("TGSN", "median")).reset_index()
    profile["kWh_dia"] = profile["Taller"] + profile["TGSN"]
    real = data.groupby("Fecha")[["Taller", "TGSN"]].sum()
    real["kWh_dia"] = real["Taller"] + real["TGSN"]
    return (profile, real), None


def proyectar_dos_suministros(seat, saf, odom, kms, inicio, fin, umr=None):
    """Puro: devuelve consumos diarios reales/proyectados por punto; valores ausentes no son cero."""
    umr = umr or {"Laboral": .964, "Sábado": .964, "Domingo/Festivo": .964}
    profiles = {}
    errors = {}
    for name, result in (("Quilpué", _perfil_quilpue(seat, odom, kms, umr)),
                         ("Limache", _perfil_limache(saf))):
        pair, err = result
        if err:
            errors[name] = err
            continue
        profile, real = pair
        profiles[name] = (profile.set_index("Tipo")["kWh_dia"].to_dict(),
                          real["kWh_dia"].to_dict(), profile)
    rows = []
    for when in pd.date_range(inicio, fin):
        day = pd.Timestamp(when).normalize()
        typ = tipo_dia(day)
        for point, (pf, actual, _) in profiles.items():
            if day in actual:
                val, source = float(actual[day]), "Real"
            elif typ in pf:
                val, source = float(pf[typ]), "Proyectado"
            else:
                continue
            if np.isfinite(val) and val >= 0:
                rows.append({"Fecha": day.date(), "Suministro": point,
                             "kWh": val, "Origen": source, "Tipo día": typ})
    return pd.DataFrame(rows), {x: p[2] for x, p in profiles.items()}, errors


def _tipos_cambio(meses_consumo, modo, dolar_manual, escenario):
    """Devuelve {periodo:(clp_usd, fuente)}; no fabrica cotizaciones futuras."""
    target = {p: _mes_siguiente(*p) if modo == "Mes siguiente (facturación)" else p for p in meses_consumo}
    if modo == "Manual / escenario":
        value = dolar_manual if dolar_manual > 0 else None
        return {p: (value, "Escenario manual" if value else "Sin valor") for p in target}, []
    years = sorted({y for y, m in target.values() if y <= date.today().year})
    serie = {}
    errors = []
    for year in years:
        try:
            serie[year] = dolar_serie_anual(year)
        except Exception as exc:
            errors.append(f"Dólar {year}: {str(exc)[:160]}")
    monthly = defaultdict(list)
    for year, values in serie.items():
        for fecha_txt, value in values:
            try:
                d = date.fromisoformat(fecha_txt)
                if d <= date.today() and d.year == year:
                    monthly[(year, d.month)].append(float(value))
            except (ValueError, TypeError):
                pass
    out = {}
    for p, key in target.items():
        if monthly.get(key):
            out[p] = (sum(monthly[key]) / len(monthly[key]), "Dólar observado promedio " + f"{key[0]}-{key[1]:02}")
        elif escenario > 0:
            out[p] = (escenario, "Escenario manual (sin promedio publicado)")
        else:
            out[p] = (None, "No disponible")
    return out, errors


def valorizar_mensual(diaria, cpis, cambios):
    """Calcula por mes/suministro; nunca trata datos o referencias faltantes como cero."""
    if diaria.empty:
        return pd.DataFrame()
    d = diaria.copy()
    d["Año"] = pd.to_datetime(d["Fecha"]).dt.year
    d["Mes"] = pd.to_datetime(d["Fecha"]).dt.month
    m = d.groupby(["Año", "Mes", "Suministro"], as_index=False).agg(
        **{"kWh": ("kWh", "sum"), "Días reales": ("Origen", lambda s: sum(s == "Real")),
           "Días proyectados": ("Origen", lambda s: sum(s == "Proyectado"))})
    records = []
    for row in m.to_dict("records"):
        y, mes = int(row["Año"]), int(row["Mes"])
        ref = cpi_referencia(y, mes)
        cpi = cpis.get(ref)
        rate = precio_mwh(cpi) if cpi is not None else None
        fx, fx_source = cambios.get((y, mes), (None, "No disponible"))
        usd = row["kWh"] / 1000 * rate if rate is not None else None
        clp = usd * fx if usd is not None and fx is not None else None
        records.append({**row, "Período": f"{MESES[mes-1]} {y}",
                        "CPI referencia": f"{MESES[ref[1]-1]} {ref[0]}",
                        "CPI": cpi, "US$/MWh": rate,
                        "CLP/USD": fx, "Fuente dólar": fx_source,
                        "Costo USD": usd, "Costo CLP": clp,
                        "Estado": ("Sin CPI contractual" if rate is None else
                                   "Sin dólar publicado" if fx is None else
                                   "Escenario FX" if "manual" in fx_source.lower() else "Calculado")})
    return pd.DataFrame(records).sort_values(["Año", "Mes", "Suministro"])


def _excel(tab, diario, perfiles, refs):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        tab.to_excel(writer, sheet_name="Costo mensual", index=False)
        diario.to_excel(writer, sheet_name="Energia diaria", index=False)
        pd.DataFrame([{"CPI referencia": f"{MESES[m-1]} {y}", "Índice": val}
                      for (y, m), val in sorted(refs.items())]).to_excel(writer, sheet_name="CPI", index=False)
        for name, data in perfiles.items():
            data.to_excel(writer, sheet_name="Perfil " + ("Quilpue" if name == "Quilpué" else "Limache"), index=False)
    return out.getvalue()


def mostrar_costo_energia():
    st.subheader("💰 Costo del consumo de energía · Quilpué y Limache")
    st.caption("Energía operativa SEAT/UMR + dólar observado + reajuste CPI del contrato Enel. "
               "Cálculo referencial de energía activa, independiente de los otros cargos CEN.")
    with st.expander("1 · Histórico energético para construir la proyección", expanded=True):
        remoto = st.checkbox("Leer históricos automáticamente de Datos-Operacionales (GitHub)",
                              value=True, key="ce_repo")
        u1, u2 = st.columns(2)
        with u1:
            upload_seat = st.file_uploader("Excel SEAT adicionales o de reemplazo", type=["xls", "xlsx", "xlsm"],
                                           accept_multiple_files=True, key="ce_upseat")
        with u2:
            upload_umr = st.file_uploader("Excel UMR adicionales o de reemplazo", type=["xls", "xlsx", "xlsm"],
                                          accept_multiple_files=True, key="ce_upumr")
        month_pairs = set()
        if remoto:
            try:
                month_pairs.update(f["period"] for f in catalogo("seat"))
            except Exception as exc:
                st.warning(f"No se pudo leer el catálogo SEAT: {exc}")
        month_pairs.update(periodo(f.name) for f in (list(upload_seat or []) + list(upload_umr or [])))
        month_pairs = sorted({(y, m) for y, m in month_pairs if y and m})
        if not month_pairs:
            st.warning("No hay históricos SEAT visibles. Comprueba la conexión GitHub o carga SEAT y UMR.")
            return
        years = sorted({y for y, _ in month_pairs})
        available_months = sorted({m for _, m in month_pairs})
        a, b = st.columns(2)
        with a:
            sel_years = st.multiselect("Años de la muestra histórica", years,
                                       default=[years[-1]], key="ce_hist_years")
        with b:
            default_m = sorted({m for y, m in month_pairs if y == years[-1]})[-6:]
            sel_months = st.multiselect("Meses de la muestra histórica", available_months,
                                        default=default_m, format_func=lambda n: MESES[n-1],
                                        key="ce_hist_months")
        u_cols = st.columns(3)
        U = {}
        for col, name in zip(u_cols, ("Laboral", "Sábado", "Domingo/Festivo")):
            with col:
                U[name] = st.number_input(f"UMR {name} (%)", 50.0, 100.0, 96.4,
                                           .1, key="ce_umr_"+name) / 100.0
        st.caption("Se utilizan medianas de jornadas Laboral, Sábado y Domingo/Festivo. "
                   "Puedes emplear más años para mejorar el perfil.")
    if not sel_years or not sel_months:
        st.info("Selecciona años y meses para la muestra histórica.")
        return
    today = datetime.now().date()
    start_default = date(today.year, today.month, 1)
    end_default = (pd.Timestamp(start_default) + pd.DateOffset(months=3) - pd.Timedelta(days=1)).date()
    st.markdown("#### 2 · Período de consumo que deseas valorizar")
    selection = st.date_input("Desde / hasta (meses de consumo)", (start_default, end_default),
                              key="ce_consumption_dates")
    if not isinstance(selection, (list, tuple)) or len(selection) != 2:
        st.info("Selecciona las dos fechas del período.")
        return
    first, last = selection
    if last < first or (last - first).days > 1095:
        st.warning("Selecciona un período ordenado de hasta tres años.")
        return
    st.markdown("#### 3 · Reajuste CPI y conversión a pesos")
    cpis, cpi_status = indices_cpi()
    st.caption(cpi_status)
    needed = sorted({cpi_referencia(int(day.year), int(day.month)) for day in pd.date_range(first, last, freq="D")})
    if not needed:
        needed = [cpi_referencia(first.year, first.month)]
    if (2025, 10) in needed:
        st.warning("El CPI-U de octubre de 2025 no fue publicado por el BLS. "
                   "El reajuste enero-junio 2026 requiere acordar con Enel un índice de reemplazo. "
                   "Deja el campo en 0 para mostrar ese semestre como 'sin CPI contractual'.")
        cpi_alt = st.number_input("CPI sustituto octubre 2025, solo si existe acuerdo con Enel",
                                  min_value=0.0, max_value=1000.0, value=0.0, step=0.001,
                                  format="%.3f", key="ce_cpi_oct25")
        if cpi_alt > 0:
            cpis[(2025, 10)] = float(cpi_alt)
            st.info("Se está aplicando un CPI de reemplazo ingresado manualmente, "
                    "no un valor publicado por el BLS para octubre de 2025.")
    missing = [(y, m) for (y, m) in needed if (y, m) not in cpis and (y, m) != (2025, 10)]
    if missing:
        with st.expander("CPI faltantes: ingreso manual documentado", expanded=True):
            st.caption("Solo ingresa índices oficiales o acordados contractualmente. "
                       "Sin CPI no se inventará precio para ese semestre.")
            for year, mon in missing:
                val = st.number_input(f"CPI {MESES[mon-1]} {year}", min_value=0.0, max_value=1000.0,
                                      value=0.0, step=0.001, format="%.3f",
                                      key=f"ce_cpi_manual_{year}_{mon}")
                if val > 0:
                    cpis[(year, mon)] = val
    exchange_mode = st.radio("Dólar para convertir US$ a CLP", ["Mes de consumo",
                               "Mes siguiente (facturación)", "Manual / escenario"],
                             horizontal=True, key="ce_dolar_mode")
    manual_fx = 0.0
    backup_fx = 0.0
    if exchange_mode == "Manual / escenario":
        manual_fx = st.number_input("Dólar de escenario para todos los meses ($/US$)",
                                    min_value=0.0, value=0.0, step=1.0, key="ce_fx_manual")
    else:
        backup_fx = st.number_input("Dólar de escenario si un mes aún no tiene promedio publicado (opcional)",
                                    min_value=0.0, value=0.0, step=1.0, key="ce_fx_backup")
    st.caption("El extracto contractual proporcionado establece el reajuste CPI, pero no "
               "especifica qué fecha de dólar debe utilizarse. Por defecto se utiliza "
               "el promedio observado del mes de consumo como estimación, no como "
               "criterio contractual confirmado.")
    with st.spinner("Leyendo históricos y calculando ambas proyecciones energéticas…"):
        try:
            (seat, saf, odom, kms), n_seat, n_umr, errors = _datos_operacionales(
                sel_years, sel_months, remoto, upload_seat, upload_umr)
        except Exception as exc:
            st.error(f"No fue posible leer datos operacionales: {exc}")
            return
    st.caption(f"Libros seleccionados: {n_seat} SEAT y {n_umr} UMR.")
    if errors:
        with st.expander(f"Advertencias de lectura ({len(errors)})"):
            for err in errors:
                st.write(err)
    daily, profiles, no_data = proyectar_dos_suministros(seat, saf, odom, kms, first, last, U)
    if no_data:
        with st.expander("Cobertura energética incompleta", expanded=True):
            for supply, err in no_data.items():
                st.warning(f"{supply}: {err}")
    if daily.empty:
        st.error("No se consiguió proyectar consumo para ningún suministro. "
                 "Revisa que los Excel SEAT y UMR tengan fechas en común.")
        return
    months = sorted(set((d.year, d.month) for d in daily["Fecha"]))
    with st.spinner("Consultando dólar observado para los meses seleccionados…"):
        ex, ex_errors = _tipos_cambio(months, exchange_mode, manual_fx, backup_fx)
    if ex_errors:
        with st.expander("Problemas de consulta del dólar"):
            for e in ex_errors:
                st.write(e)
    monthly = valorizar_mensual(daily, cpis, ex)
    st.markdown("#### 4 · Costo estimado de energía activa")
    summary = []
    for point in ("Quilpué", "Limache"):
        d = monthly[monthly["Suministro"] == point]
        if d.empty:
            continue
        summary.append({"Suministro": point,
                        "Energía (kWh)": float(d["kWh"].sum()),
                        "Costo US$ calculable": float(d["Costo USD"].dropna().sum()) if d["Costo USD"].notna().all() else None,
                        "Costo CLP calculable": float(d["Costo CLP"].dropna().sum()) if d["Costo CLP"].notna().all() else None,
                        "Meses": len(d), "Meses incompletos": int((d["Estado"] != "Calculado").sum())})
    left, right, total = st.columns(3)
    def render_metric(area, element):
        row = next((x for x in summary if x["Suministro"] == area), None)
        if row is None:
            element.metric(area, "Sin datos")
            return
        money = row["Costo CLP calculable"]
        element.metric(area, "$" + fmt(money, 0) if money is not None else "No valorizable",
                       fmt(row["Energía (kWh)"], 0) + " kWh")
    render_metric("Quilpué", left)
    render_metric("Limache", right)
    calculated = [r["Costo CLP calculable"] for r in summary]
    total_amt = sum(calculated) if summary and all(v is not None for v in calculated) else None
    total.metric("Total energía", "$" + fmt(total_amt, 0) if total_amt is not None else "Incompleto",
                 "No incluye otros cargos CEN")
    if any(r["Meses incompletos"] for r in summary):
        st.warning("Hay meses con CPI o dólar faltante, o un dólar de escenario manual. "
                   "Los totales solo se presentan cuando todos los meses tienen valores y "
                   "debes revisar los estados por período.")
    st.dataframe([{"Suministro": x["Suministro"], "Energía kWh": round(x["Energía (kWh)"], 2),
                   "Costo USD": round(x["Costo US$ calculable"], 2) if x["Costo US$ calculable"] is not None else None,
                   "Costo CLP": round(x["Costo CLP calculable"], 2) if x["Costo CLP calculable"] is not None else None,
                   "Meses": x["Meses"]} for x in summary], use_container_width=True, hide_index=True)
    st.markdown("#### Detalle mensual — CPI, precio, dólar y costos")
    display = monthly.copy()
    for col in ["kWh", "CPI", "US$/MWh", "CLP/USD", "Costo USD", "Costo CLP"]:
        display[col] = display[col].round(3 if col in ("CPI", "US$/MWh") else 2)
    st.dataframe(display.drop(columns=["Año", "Mes"]), hide_index=True, use_container_width=True)
    if monthly["Año"].nunique() > 1:
        st.markdown("#### Comparación anual por suministro")
        anual = []
        for (yr, point), grp in monthly.groupby(["Año", "Suministro"], sort=True):
            anual.append({"Año": int(yr), "Suministro": point,
                          "Energía kWh": round(float(grp["kWh"].sum()), 2),
                          "Costo US$": round(float(grp["Costo USD"].sum()), 2) if grp["Costo USD"].notna().all() else None,
                          "Costo CLP": round(float(grp["Costo CLP"].sum()), 2) if grp["Costo CLP"].notna().all() else None,
                          "Meses con información": len(grp)})
        st.dataframe(anual, use_container_width=True, hide_index=True)
    if not monthly.empty:
        f = go.Figure()
        for point in ("Quilpué", "Limache"):
            rows = monthly[monthly["Suministro"] == point]
            f.add_trace(go.Bar(name=point, x=rows["Período"], y=rows["Costo CLP"],
                               hovertemplate="%{x}<extra>" + point + "</extra>"))
        f.update_layout(barmode="group", title="Costo estimado de energía activa por mes",
                        yaxis_title="$ CLP", xaxis_title="Mes", height=420)
        st.plotly_chart(f, use_container_width=True, config={"displaylogo": False})
    with st.expander("Ver energía real / proyectada día a día y perfiles"):
        st.dataframe(daily, hide_index=True, use_container_width=True)
        for name, pr in profiles.items():
            st.markdown("**" + name + "**")
            st.dataframe(pr, hide_index=True, use_container_width=True)
    st.download_button("⬇ Descargar cálculo completo en Excel", _excel(monthly, daily, profiles, cpis),
                       "EFE_costo_energia_Quilpue_Limache.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="ce_download")
    st.info("**Alcance:** energía SEAT se asocia provisionalmente con Quilpué y SAF LI con Limache. "
            "Debes confirmar que esos consumos equivalgan a los retiros contractuales que factura Enel; "
            "la medición SAF LI puede cubrir solo parte del consumo de Limache. "
            "El contrato fija el precio en S/E Alto Jahuel 220 kV, por lo que no se aplican "
            "factores nodales no informados. No se agregan pérdidas, peajes, "
            "potencia, compensación PE, SSCC, MT/PD ni exenciones CI.")
    st.caption("Fuentes: contrato ENEL_EFE_PPA_final_firmado (extracto aportado); "
               "CPI-U BLS CUUR0000SA0 / FRED CPIAUCNS; "
               "dólar observado del módulo ya integrado; energía de Datos-Operacionales.")
