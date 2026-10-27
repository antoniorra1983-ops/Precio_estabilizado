"""Proyección de energía operativa trasladada de Datos-Operacionales.

Fuente principal: Excel SEAT y UMR en el repositorio público
antoniorra1983-ops/Datos-Operacionales, o archivos subidos por el usuario.
Independiente de los balances económicos del CEN.
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import timedelta
from io import BytesIO
import json
import re
import unicodedata
from urllib.request import Request, urlopen

import holidays
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

REPO = "antoniorra1983-ops/Datos-Operacionales"
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MES_TXT = dict(zip(["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"], range(1, 13)))
FERIADOS = holidays.country_holidays("CL")


def normal(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii").upper()


def numero(v):
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v) if np.isfinite(float(v)) else 0.0
    s = str(v).strip().replace(" ", "").replace("$", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def periodo(nombre):
    n = normal(nombre)
    year = re.search(r"20\d{2}", n)
    year = int(year.group(0)) if year else None
    month = next((i for s, i in MES_TXT.items() if s in n), None)
    if month is None:
        m = re.match(r"(0[1-9]|1[0-2])[-_ ]", n)
        month = int(m.group(1)) if m else None
    return year, month


def descargar(url):
    req = Request(url, headers={"User-Agent": "EFE-Operacionales-Proyeccion/1.0", "Accept": "application/vnd.github+json"})
    with urlopen(req, timeout=40) as r:
        return r.read()


@st.cache_data(ttl=1200, show_spinner=False)
def catalogo(tipo):
    url = f"https://api.github.com/repos/{REPO}/contents/data/{tipo}?ref=main"
    data = json.loads(descargar(url))
    return [{"name": f["name"], "url": f["download_url"], "period": periodo(f["name"])}
            for f in data if f.get("type") == "file" and f.get("download_url")
            and f["name"].lower().endswith((".xls", ".xlsx", ".xlsm")) and all(periodo(f["name"]))]


@st.cache_data(ttl=1800, max_entries=120, show_spinner=False)
def descargar_archivo(url):
    return descargar(url)


def leer(data, filename, **kw):
    engine = "xlrd" if filename.lower().endswith(".xls") else "openpyxl"
    return pd.read_excel(BytesIO(data), engine=engine, **kw)


def leer_seat(data, filename):
    seat, saf = [], []
    d = leer(data, filename, header=None)
    if d.shape[1] >= 8:
        fechas = pd.to_datetime(d.iloc[:, 1], errors="coerce").dt.normalize()
        for i in d.index[fechas.notna()]:
            seat.append({"Fecha": fechas.loc[i], "Traccion": numero(d.iloc[i, 5]),
                         "E12": numero(d.iloc[i, 7]), "Total": numero(d.iloc[i, 3])})
    try:
        ex = pd.ExcelFile(BytesIO(data), engine="xlrd" if filename.lower().endswith(".xls") else "openpyxl")
        sheet = next((s for s in ex.sheet_names if "SAF" in normal(s)), None)
        if sheet:
            s = leer(data, filename, sheet_name=sheet, header=None)
            if s.shape[1] >= 8:
                fechas = pd.to_datetime(s.iloc[:, 1], errors="coerce").dt.normalize()
                for i in s.index[fechas.notna()]:
                    t, g = numero(s.iloc[i, 6]), numero(s.iloc[i, 7])
                    if t > 0 or g > 0:
                        saf.append({"Fecha": fechas.loc[i], "Taller": t, "TGSN": g})
    except (ValueError, KeyError, IndexError):
        pass
    return seat, saf


def leer_umr(data, filename):
    odometros, servicios = [], []
    ex = pd.ExcelFile(BytesIO(data), engine="xlrd" if filename.lower().endswith(".xls") else "openpyxl")
    for sn in ex.sheet_names:
        raw = leer(data, filename, sheet_name=sn, header=None)
        if raw.empty:
            continue
        header = next((i for i in range(min(100, len(raw)))
                       if any(x in normal(raw.iloc[i].tolist()) for x in ("FECHA", "ODO", "KILOM"))), None)
        if header is not None:
            tab = leer(data, filename, sheet_name=sn, header=header)
            tab.columns = [normal(c).strip() for c in tab.columns]
            fc = next((c for c in tab.columns if "FECHA" in c), None)
            oc = next((c for c in tab.columns if "ODO" in c), None)
            if fc and oc:
                fechas = pd.to_datetime(tab[fc], errors="coerce").dt.normalize()
                for i in tab.index[fechas.notna()]:
                    val = numero(tab.loc[i, oc])
                    if val > 0:
                        odometros.append({"Fecha": fechas.loc[i], "Odom": val})
        if "SERV" in normal(sn) and "KM" in normal(sn) and raw.shape[1] > 23:
            fechas = pd.to_datetime(raw.iloc[4:, 0], errors="coerce").ffill()
            for i in raw.index[4:]:
                fecha = fechas.get(i)
                ruta = str(raw.iloc[i, 5]).strip().upper()
                if pd.notna(fecha) and re.match(r"^[A-Z]{2,3}-[A-Z]{2,3}$", ruta):
                    servicios.append({"Fecha": pd.Timestamp(fecha).normalize(),
                                      "TrenKmP": numero(raw.iloc[i, 23]),
                                      "TrenKmR": numero(raw.iloc[i, 22])})
    return odometros, servicios


def por_dia(rows, years, months):
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).dropna(subset=["Fecha"])
    df["Fecha"] = pd.to_datetime(df["Fecha"]).dt.normalize()
    df = df[df["Fecha"].dt.year.isin(years) & df["Fecha"].dt.month.isin(months)]
    if "_archivo" in df.columns:
        # Los libros SEAT de meses contiguos pueden repetir fechas: nunca sumarlas dos veces.
        grouped = df.groupby(["Fecha", "_archivo"], as_index=False).sum(numeric_only=True)
        grouped["_prioridad"] = grouped.apply(
            lambda row: int(periodo(row["_archivo"]) == (row["Fecha"].year, row["Fecha"].month)), axis=1)
        grouped = grouped.sort_values(["Fecha", "_prioridad", "_archivo"], ascending=[True, False, True])
        return grouped.drop_duplicates("Fecha").drop(columns=["_archivo", "_prioridad"])
    return df.groupby("Fecha", as_index=False).sum(numeric_only=True)


def tipo_dia(d):
    fecha = pd.Timestamp(d).date()
    return "Domingo/Festivo" if fecha.weekday() == 6 or fecha in FERIADOS else ("Sábado" if fecha.weekday() == 5 else "Laboral")


def fmt(v, n=0):
    return f"{v:,.{n}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _consumos_mismo_modelo(seat, saf, odom, kms, anos, meses, desde, hasta, umr=None):
    """Calcula ambos suministros con las MISMAS fórmulas y filtros de la pantalla.

    Entrada: listas Excel o DataFrames ya normalizados. Salida por (YYYY-MM, punto).
    No requiere visitar otra pestaña ni utiliza el motor diferente de costo_energia.
    """
    anos, meses = list(anos), list(meses)
    umr = umr or {t: .964 for t in ('Laboral', 'Sábado', 'Domingo/Festivo')}

    def preparar(data):
        if isinstance(data, pd.DataFrame):
            if data.empty:
                return data.copy()
            frame = data.copy()
            frame['Fecha'] = pd.to_datetime(frame['Fecha']).dt.normalize()
            return frame[frame['Fecha'].dt.year.isin(anos) & frame['Fecha'].dt.month.isin(meses)].copy()
        return por_dia(data, anos, meses)

    se, sa, od, km = map(preparar, (seat, saf, odom, kms))
    perfiles, reales, incidencias = {}, {}, []

    # Quilpué: se replica literalmente el cálculo SEAT total de mostrar_proyeccion().
    if any(d.empty for d in (se, od, km)):
        incidencias.append('Quilpué: se necesitan SEAT y UMR (odómetro y KM-Servicio).')
    else:
        d = km.merge(od, on='Fecha', how='inner').merge(se, on='Fecha', how='inner')
        d = d[(d['TrenKmP'] > 0) & (d['Odom'] > 0) &
              ((d['Traccion'] + d['E12']) > 0)].copy()
        if d.empty:
            incidencias.append('Quilpué: SEAT y UMR no comparten días válidos de tracción, odómetro y kilómetros.')
        else:
            d['Tipo'] = d['Fecha'].map(tipo_dia)
            d['IDE_tr'] = d['Traccion'] / d['Odom']
            pro = d.groupby('Tipo').agg(
                TrenKmP=('TrenKmP', 'median'), IDE_tr=('IDE_tr', 'median'),
                E12=('E12', 'median')).reset_index()
            pro['UMR'] = pro['Tipo'].map(umr)
            pro['total'] = pro['IDE_tr'] * pro['TrenKmP'] / pro['UMR'] + pro['E12']
            pro = pro[pd.to_numeric(pro['total'], errors='coerce').notna() & (pro['total'] >= 0)]
            perfiles['Quilpué'] = pro.set_index('Tipo')['total'].to_dict()
            actual = d.groupby('Fecha')[['Traccion', 'E12']].sum()
            reales['Quilpué'] = (actual['Traccion'] + actual['E12']).to_dict()

    # Limache: igual a mostrar_proyeccion() (hoja SAFs: Taller + TGSN).
    if sa.empty:
        incidencias.append('Limache: faltan consumos de la hoja SAFs de SEAT.')
    else:
        d = sa[(sa['Taller'] + sa['TGSN']) > 0].copy()
        if d.empty:
            incidencias.append('Limache: los registros SAFs no contienen consumos positivos.')
        else:
            d['Tipo'] = d['Fecha'].map(tipo_dia)
            pro = d.groupby('Tipo').agg(
                Taller=('Taller', 'median'), TGSN=('TGSN', 'median')).reset_index()
            pro['total'] = pro['Taller'] + pro['TGSN']
            perfiles['Limache'] = pro.set_index('Tipo')['total'].to_dict()
            actual = d.groupby('Fecha')[['Taller', 'TGSN']].sum()
            reales['Limache'] = (actual['Taller'] + actual['TGSN']).to_dict()

    registros = []
    for dia in pd.date_range(desde, hasta):
        tipo = tipo_dia(dia)
        stamp = pd.Timestamp(dia).normalize()
        for sitio in ('Quilpué', 'Limache'):
            if sitio not in perfiles:
                continue
            if stamp in reales[sitio]:
                energia, origen = float(reales[sitio][stamp]), 'Real'
            elif tipo in perfiles[sitio]:
                energia, origen = float(perfiles[sitio][tipo]), 'Proyectado'
            else:
                continue
            if np.isfinite(energia) and energia >= 0:
                registros.append({'Fecha': stamp, 'Suministro': sitio,
                                  'kWh': energia, 'Origen': origen})
    if not registros:
        return {}, incidencias
    diario = pd.DataFrame(registros)
    diario['Mes'] = diario['Fecha'].dt.strftime('%Y-%m')
    resumen = {}
    for (mes, sitio), grupo in diario.groupby(['Mes', 'Suministro']):
        real = float(grupo.loc[grupo['Origen'] == 'Real', 'kWh'].sum())
        proyectado = float(grupo.loc[grupo['Origen'] == 'Proyectado', 'kWh'].sum())
        resumen[(mes, sitio)] = {
            'kwh': real + proyectado, 'real_kwh': real,
            'proyectado_kwh': proyectado, 'dias': int(grupo['Fecha'].nunique()),
            'origen': 'Proyección de energía (mismo modelo visualizado)',
        }
    return resumen, incidencias


def _registrar_consumos_para_tarifas(seat, saf, odom, kms, anos, meses, desde, hasta, umr):
    """Publica todos los meses para las dos estaciones usando el modelo de pantalla."""
    resumen, incidencias = _consumos_mismo_modelo(
        seat, saf, odom, kms, anos, meses, desde, hasta, umr)
    st.session_state['efe_proyeccion_tarifas'] = {
        'consumos': resumen, 'desde': desde.isoformat(), 'hasta': hasta.isoformat(),
        'anios': list(anos), 'meses': list(meses), 'errores': incidencias,
        'nota': 'SEAT tracción+12 kV = Quilpué; SAF LI Taller+TGSN = Limache. Proyección operacional.',
    }
    return resumen


def _guardar_valores_visibles(rows, sitio, fecha_desde, fecha_hasta):
    """Una segunda comprobación: lo que se VE en las métricas se publica tal cual.

    Si se ha seleccionado solo Tracción, no se reemplaza el SEAT total de Quilpué.
    Este método sincroniza también ediciones manuales del UMR y la fecha seleccionada.
    """
    if not rows:
        return
    estado = st.session_state.get('efe_proyeccion_tarifas') or {}
    resumen = dict(estado.get('consumos', {}))
    data = pd.DataFrame(rows)
    data['Mes'] = pd.to_datetime(data['Fecha']).dt.strftime('%Y-%m')
    for mes, grupo in data.groupby('Mes'):
        real = float(grupo.loc[grupo['Origen'] == 'Real', 'kWh'].sum())
        proy = float(grupo.loc[grupo['Origen'] == 'Proyectado', 'kWh'].sum())
        resumen[(mes, sitio)] = {
            'kwh': real+proy, 'real_kwh': real, 'proyectado_kwh': proy,
            'dias': int(grupo['Fecha'].nunique()),
            'origen': 'Proyección de energía (valores visibles)',
        }
    estado['consumos'] = resumen
    estado['desde'], estado['hasta'] = fecha_desde.isoformat(), fecha_hasta.isoformat()
    st.session_state['efe_proyeccion_tarifas'] = estado


def mostrar_proyeccion():
    # No borrar el resultado compartido en cada rerun: Streamlit vuelve a ejecutar
    # todas las pestañas al cambiar cualquier filtro. Un error temporal de red
    # no debe vaciar los kWh ya calculados y mostrados correctamente.
    st.subheader("🔮 Proyección de energía operativa")
    st.caption("Modelo adaptado desde Datos-Operacionales. Fuente SEAT/UMR o carga manual; independiente del Balance CEN.")
    with st.expander("Fuentes de datos y filtros del histórico", expanded=True):
        remote = st.checkbox("Leer Excel del repositorio Datos-Operacionales", value=True, key="pe_ops_remote")
        a, b = st.columns(2)
        with a:
            upload_seat = st.file_uploader("Agregar Excel SEAT", type=["xls", "xlsx", "xlsm"],
                                           accept_multiple_files=True, key="pe_ops_seat")
        with b:
            upload_umr = st.file_uploader("Agregar Excel UMR", type=["xls", "xlsx", "xlsm"],
                                          accept_multiple_files=True, key="pe_ops_umr")
        catalogs = {"seat": [], "umr": []}
        if remote:
            try:
                catalogs = {"seat": catalogo("seat"), "umr": catalogo("umr")}
            except Exception as exc:
                st.warning(f"No se pudo consultar Datos-Operacionales: {exc}. Prueba la carga manual.")
        available = [r["period"] for items in catalogs.values() for r in items]
        available += [periodo(f.name) for f in list(upload_seat or []) + list(upload_umr or [])]
        available = [(y, m) for y, m in available if y and m]
        years = sorted({y for y, _ in available})
        months = sorted({m for _, m in available})
        if not years:
            st.info("No hay Excel SEAT/UMR disponibles. Sube archivos o habilita GitHub.")
            return
        selection_years = st.multiselect("Años del histórico", years, default=[years[-1]], key="pe_ops_years")
        selection_months = st.multiselect("Meses del histórico", months, default=months,
                                          format_func=lambda x: MESES[x-1], key="pe_ops_months")
        if not selection_years or not selection_months:
            st.info("Selecciona al menos un mes y un año.")
            return
    selected = {"seat": [], "umr": []}
    if remote:
        for kind in selected:
            chosen = {}
            for f in catalogs[kind]:
                if f["period"][0] in selection_years and f["period"][1] in selection_months:
                    chosen[f["period"]] = (f["name"], f["url"], None)
            selected[kind] = list(chosen.values())
    for kind, uploads in (("seat", upload_seat), ("umr", upload_umr)):
        for f in uploads or []:
            y, m = periodo(f.name)
            if y in selection_years and m in selection_months:
                selected[kind] = [r for r in selected[kind] if periodo(r[0]) != (y, m)]
                selected[kind].append((f.name, "", f.getvalue()))
    if not selected["seat"]:
        st.warning("No hay archivos SEAT para los períodos seleccionados.")
        return
    seat, saf, odom, kms = [], [], [], []
    errors = []
    jobs = [(k, *f) for k, fs in selected.items() for f in fs]
    def work(item):
        kind, name, url, content = item
        data = content if content is not None else descargar_archivo(url)
        return kind, name, leer_seat(data, name) if kind == "seat" else leer_umr(data, name)
    with st.spinner(f"Procesando {len(jobs)} Excel operacionales…"):
        with ThreadPoolExecutor(max_workers=4) as pool:
            pending = {pool.submit(work, j): j for j in jobs}
            for f in as_completed(pending):
                try:
                    kind, name, result = f.result()
                    if kind == "seat":
                        for row in result[0] + result[1]:
                            row["_archivo"] = name
                        seat.extend(result[0]); saf.extend(result[1])
                    else:
                        for row in result[0] + result[1]:
                            row["_archivo"] = name
                        odom.extend(result[0]); kms.extend(result[1])
                except Exception as exc:
                    errors.append(f"{pending[f][1]}: {exc}")
    if errors:
        with st.expander(f"Archivos sin procesar ({len(errors)})"):
            for msg in errors: st.write(msg)
    st.caption(f"Archivos utilizados: {len(selected['seat'])} SEAT y {len(selected['umr'])} UMR. "
               "Solo se calcula con datos provenientes de GitHub o subidos en esta pestaña.")
    # Registrar datos mensuales ANTES de posibles retornos de la vista individual.
    # Así Limache puede compartirse aun si la vista SEAT no tiene UMR suficientes.
    # Más abajo se sobrescriben con las fechas/UMR exactas de la vista cuando
    # el usuario dispone de una proyección calculada.
    try:
        _datos_compartibles = seat or saf
        _fechas_base = [pd.Timestamp(x['Fecha']).date() for x in _datos_compartibles
                        if pd.notna(x.get('Fecha'))]
        if _fechas_base:
            _min_base, _max_base = min(_fechas_base), max(_fechas_base)
            _previo = st.session_state.get('efe_proyeccion_tarifas') or {}
            _rango_previo = (str(_previo.get('desde','')),str(_previo.get('hasta','')))
            _anios_previos = _previo.get('anios') or []
            _meses_previos = _previo.get('meses') or []
            if (_rango_previo[0] and _rango_previo[1]
                and list(selection_years) == list(_anios_previos)
                and list(selection_months) == list(_meses_previos)):
                from datetime import date as _date
                _inicio_base = _date.fromisoformat(_rango_previo[0])
                _fin_base = _date.fromisoformat(_rango_previo[1])
            else:
                _inicio_base, _fin_base = _min_base, _max_base
            _registrar_consumos_para_tarifas(
                seat, saf, odom, kms, selection_years, selection_months,
                _inicio_base, _fin_base, {
                    lab: (float(st.session_state.get(f'pe_ops_umr_{lab}', 96.4)) / 100
                          if st.session_state.get('pe_ops_manual_umr', False) else .964)
                    for lab in ('Laboral','Sábado','Domingo/Festivo')})
    except Exception as exc:
        st.warning('No se pudo preparar el enlace automático para Tarifas CNE: '+str(exc))
    source = st.radio("Fuente de energía", ["SEAT (tracción + 12 kV)", "Limache (SAF LI)"],
                      horizontal=True, key="pe_ops_source")
    es_seat = source.startswith("SEAT")
    if es_seat:
        se, od, km = [por_dia(r, selection_years, selection_months) for r in (seat, odom, kms)]
        if any(d.empty for d in (se, od, km)):
            st.warning("Para SEAT faltan datos coincidentes SEAT y UMR con odómetro y hoja KM-Servicio.")
            return
        d = km.merge(od, on="Fecha", how="inner").merge(se, on="Fecha", how="inner")
        d = d[(d["TrenKmP"] > 0) & (d["Odom"] > 0) & ((d["Traccion"] + d["E12"]) > 0)].copy()
        if d.empty:
            st.warning("No hay días con Tren-Km, odómetro y energía coincidentes.")
            return
        u = {"Laboral": .964, "Sábado": .964, "Domingo/Festivo": .964}
        if st.checkbox("Ajustar UMR por jornada", key="pe_ops_manual_umr"):
            cols = st.columns(3)
            for c, lab in zip(cols, u):
                u[lab] = c.number_input(f"UMR {lab} (%)", 50., 100., 96.4, .1,
                                         key=f"pe_ops_umr_{lab}") / 100.
        d["Tipo"] = d["Fecha"].map(tipo_dia)
        d["IDE_tr"] = d["Traccion"] / d["Odom"]
        profile = d.groupby("Tipo").agg(Dias=("Fecha", "size"), TrenKmP=("TrenKmP", "median"),
            IDE_tr=("IDE_tr", "median"), E12=("E12", "median")).reset_index()
        profile["UMR"] = profile["Tipo"].map(u)
        profile["Odom_proy"] = profile["TrenKmP"] / profile["UMR"]
        profile["Traccion"] = profile["IDE_tr"] * profile["Odom_proy"]
        comps = [("Tracción", "Traccion"), ("12 kV", "E12")]
        mode = st.radio("Energía a proyectar", ["Total", "Tracción"], horizontal=True, key="pe_ops_mode")
        wanted = comps[:1] if mode == "Tracción" else comps
        label, with_odom = mode, True
    else:
        d = por_dia(saf, selection_years, selection_months)
        if d.empty:
            st.warning("No se encontraron datos Limache (SAF LI) en la hoja SAFs de los SEAT seleccionados.")
            return
        d = d[(d["Taller"] + d["TGSN"]) > 0].copy()
        d["Tipo"] = d["Fecha"].map(tipo_dia)
        profile = d.groupby("Tipo").agg(Dias=("Fecha", "size"), Taller=("Taller", "median"),
                                        TGSN=("TGSN", "median")).reset_index()
        comps = [("Taller", "Taller"), ("SAF LI TGSN", "TGSN")]
        wanted = comps
        label, with_odom = "Principal Limache", False
    if d.empty or profile.empty:
        st.info("No hay datos históricos suficientes para proyectar.")
        return
    start, last = d["Fecha"].min().date(), d["Fecha"].max().date()
    chosen = st.date_input("Período (real + proyección)", (start, last + timedelta(days=30)),
                           key="pe_ops_dates")
    if not isinstance(chosen, (tuple, list)) or len(chosen) != 2:
        st.info("Selecciona dos fechas para la proyección.")
        return
    first, end = chosen
    if first > end or (end - first).days > 1095:
        st.warning("El período debe estar ordenado y no superar tres años.")
        return
    # Compartir kWh mensuales de ambos puntos, independientemente de si aquí
    # se está visualizando SEAT o SAF LI. Siempre usar SEAT TOTAL para Quilpué.
    _umr_para_cne = u if es_seat else {
        lab: (float(st.session_state.get(f"pe_ops_umr_{lab}", 96.4)) / 100
              if st.session_state.get("pe_ops_manual_umr", False) else .964)
        for lab in ("Laboral", "Sábado", "Domingo/Festivo")}
    try:
        _shared_cne = _registrar_consumos_para_tarifas(
            seat, saf, odom, kms, selection_years, selection_months,
            first, end, _umr_para_cne)
        if _shared_cne:
            st.caption(f"Consumos preparados para Tarifas CNE: {len(_shared_cne)} "
                       "registros mensuales, separando Quilpué y Limache.")
        for incidencia in st.session_state.get('efe_proyeccion_tarifas', {}).get('errores', []):
            st.caption('⚠ ' + incidencia)
    except Exception as exc:
        st.warning("No fue posible compartir los consumos con Tarifas CNE: " + str(exc))
    profile_map = profile.set_index("Tipo").to_dict("index")
    real = d.groupby("Fecha")[[c for _, c in comps] + (["Odom"] if with_odom else [])].sum().to_dict("index")
    totals = {lab: {"Real": 0., "Proyectado": 0.} for lab, _ in comps}
    o_real = o_proj = 0.
    n_real = n_proj = 0
    rows = []
    for dt in pd.date_range(first, end):
        day = pd.Timestamp(dt).normalize()
        typ = tipo_dia(dt)
        if day in real:
            values, origin = real[day], "Real"
            n_real += 1
            for lab, col in comps: totals[lab][origin] += values[col]
            if with_odom: o_real += values["Odom"]
            energy = sum(values[col] for _, col in wanted)
        elif typ in profile_map:
            values, origin = profile_map[typ], "Proyectado"
            n_proj += 1
            for lab, col in comps: totals[lab][origin] += values[col]
            if with_odom: o_proj += values["Odom_proy"]
            energy = sum(values[col] for _, col in wanted)
        else:
            continue
        rows.append({"Fecha": dt, "kWh": energy, "Origen": origin})
    if not rows:
        st.warning("No existen días reales ni proyectables en el período.")
        return
    # Fuente de verdad: exactamente los kWh mostrados por esta pestaña.
    # Quilpué se valoriza con SEAT total; el modo 'Tracción' solo es exploratorio.
    if not (es_seat and mode == 'Tracción'):
        _guardar_valores_visibles(rows, 'Quilpué' if es_seat else 'Limache', first, end)
    else:
        st.caption('Tarifas CNE usa SEAT TOTAL (tracción + 12 kV), no el filtro exploratorio Tracción.')
    tr = sum(totals[k]["Real"] for k, _ in wanted)
    tp = sum(totals[k]["Proyectado"] for k, _ in wanted)
    k = st.columns(4)
    k[0].metric("Total del período", fmt(tr+tp, 2) + " kWh")
    k[1].metric("Real", fmt(tr, 2) + " kWh", f"{n_real} días")
    k[2].metric("Proyectado", fmt(tp, 2) + " kWh", f"{n_proj} días")
    k[3].metric("Promedio diario", fmt((tr+tp)/len(rows), 2) + " kWh")
    table = [{"Componente": lab + " (kWh)", **totals[lab]} for lab, _ in comps]
    if with_odom:
        table.insert(0, {"Componente": "Odómetro (km)", "Real": o_real, "Proyectado": o_proj})
    tabla = pd.DataFrame(table)
    tabla["Total período"] = tabla["Real"] + tabla["Proyectado"]
    st.markdown("#### Real + proyectado por componente")
    st.dataframe(tabla, use_container_width=True, hide_index=True)
    st.markdown("#### Perfil por tipo de día (mediana histórica)")
    st.dataframe(profile.round(3), use_container_width=True, hide_index=True)
    df = pd.DataFrame(rows)
    fig = go.Figure()
    for origen, color, dash in (("Real", "#1f4e79", "solid"), ("Proyectado", "#E85500", "dash")):
        sub = df[df["Origen"] == origen]
        if not sub.empty:
            fig.add_trace(go.Scatter(x=sub["Fecha"], y=sub["kWh"], mode="lines+markers",
                                     name=origen, line=dict(color=color, dash=dash)))
    fig.update_layout(title=f"{label} — real y proyección", xaxis_title="Fecha", yaxis_title="kWh",
                      hovermode="x unified", height=420)
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})
    buff = BytesIO()
    with pd.ExcelWriter(buff, engine="openpyxl") as xl:
        df.to_excel(xl, sheet_name="Real y proyectado", index=False)
        tabla.to_excel(xl, sheet_name="Resumen", index=False)
        profile.to_excel(xl, sheet_name="Perfil de jornadas", index=False)
    st.download_button("⬇ Descargar proyección (Excel)", buff.getvalue(),
                       file_name="proyeccion_energia_" + ("seat" if es_seat else "limache") + ".xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="pe_ops_download")
    st.caption("Método de Datos-Operacionales: mediana por jornada; Tracción = IDE × (Tren-Km / UMR). "
               "No corresponde a un cargo económico del CEN.")
