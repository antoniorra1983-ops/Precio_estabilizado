"""Aplicación web Compensación PE EFE Valparaíso, versión Cloud autocontenida.

Motor de cálculo y pantalla integrados. No requiere importar cen_core.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from io import BytesIO, StringIO
import csv
import re
import unicodedata
from xml.etree import ElementTree as ET
from zipfile import ZipFile

XML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = "{" + XML + "}"
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def normalize(s: object) -> str:
    s = unicodedata.normalize("NFKD", str(s or ""))
    return "".join(c for c in s if not unicodedata.combining(c)).strip().upper()


def excel_date(serial: float) -> date:
    return date(1899, 12, 30) + timedelta(days=int(serial))


class WorkbookReader:
    """Lee XLSX/XLSM mediante el XML comprimido, sin alterar el original."""
    def __init__(self, data: bytes):
        self.zip = ZipFile(BytesIO(data))
        root = ET.fromstring(self.zip.read("xl/workbook.xml"))
        rel_root = ET.fromstring(self.zip.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: r.attrib["Target"] for r in rel_root}
        self.sheets: dict[str, str] = {}
        for e in root.find(NS + "sheets"):
            target = targets[e.attrib[REL_ID]].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            self.sheets[e.attrib["name"]] = target
        self._strings = []
        if "xl/sharedStrings.xml" in self.zip.namelist():
            with self.zip.open("xl/sharedStrings.xml") as f:
                for _, el in ET.iterparse(f, events=("end",)):
                    if el.tag == NS + "si":
                        self._strings.append("".join(t.text or "" for t in el.iter(NS + "t")))
                        el.clear()

    def sheet(self, name: str):
        match = next((s for s in self.sheets if normalize(s) == normalize(name)), None)
        if not match:
            raise ValueError(f"No existe la hoja '{name}'. Disponibles: {', '.join(self.sheets)}")
        with self.zip.open(self.sheets[match]) as f:
            for _, elem in ET.iterparse(f, events=("end",)):
                if elem.tag != NS + "row":
                    continue
                row = {}
                for c in elem.iter(NS + "c"):
                    ref = re.match(r"([A-Z]+)\d+", c.attrib.get("r", ""))
                    if not ref:
                        continue
                    col = ref.group(1)
                    v = c.find(NS + "v")
                    inline = c.find(NS + "is")
                    value = v.text if v is not None else None
                    kind = c.attrib.get("t")
                    if kind == "s" and value is not None:
                        value = self._strings[int(value)]
                    elif kind == "inlineStr" and inline is not None:
                        value = "".join(t.text or "" for t in inline.iter(NS + "t"))
                    elif value is not None and kind not in ("str", "e", "inlineStr", "s"):
                        try:
                            value = float(value)
                        except (TypeError, ValueError):
                            pass
                    if value is not None:
                        row[col] = value
                number = int(elem.attrib.get("r", 0))
                elem.clear()
                yield number, row


def num(v: object) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    if v is None or str(v).strip() == "":
        return 0.0
    raw = str(v).replace("$", "").strip().replace(" ", "")
    if "," in raw and "." in raw:
        raw = raw.replace(".", "").replace(",", ".") if raw.rfind(",") > raw.rfind(".") else raw.replace(",", "")
    elif "," in raw:
        raw = raw.replace(",", ".")
    return float(raw)


@dataclass
class BalanceResult:
    label: str
    compensacion_sen_clp: float
    retiro_sen_mwh: float
    tarifa_clp_mwh: float
    consumo_efe_kwh: float
    detalle: list[dict]
    retiros_por_tipo: dict[str, float]
    metodo: str
    avisos: list[str] = field(default_factory=list)

    @property
    def proyeccion_clp(self) -> float:
        return self.consumo_efe_kwh / 1000.0 * self.tarifa_clp_mwh


def month_from_filename(filename: str) -> str:
    name = filename or ""
    m = re.search(r"(?<!\d)(\d{2})(0[1-9]|1[0-2])[PD]?(?!\d)", name.upper())
    if m:
        return f"20{m.group(1)}-{m.group(2)}"
    return "Sin período identificado"


def parse_balance(data: bytes, filename: str, keys: list[str]) -> BalanceResult:
    reader = WorkbookReader(data)
    mvals = []
    for idx, d in reader.sheet("Cuadro de Energía"):
        if normalize(d.get("A")) == "TOTAL GENERAL":
            mvals.append((idx, num(d.get("L")), num(d.get("M"))))
    if len(mvals) != 1:
        raise ValueError("No se encontró una única fila 'Total general' en 'Cuadro de Energía'.")
    totalrow, extra_ingresos, comp_signed = mvals[0]
    if abs(comp_signed) < 1:
        raise ValueError("La compensación agregada es cero/no está calculada; verificar datos guardados en Excel.")
    # En los archivos CEN observados, M es compensación con signo negativo para generadores
    # y L es contrapartida positiva. Para presupuesto de retiros se toma magnitud.
    comp = abs(comp_signed)
    aggregate = {"L": 0.0, "L_D": 0.0, "R": 0.0}
    wanted = {normalize(k): k for k in keys if str(k).strip()}
    found = []
    for idx, d in reader.sheet("Balance Valorizado"):
        if idx == 1:
            continue
        kind = normalize(d.get("J"))
        kwh = num(d.get("I"))
        if kind in aggregate:
            aggregate[kind] -= kwh  # signo CEN: los retiros son negativos
        key = normalize(d.get("A"))
        if key in wanted:
            if kind not in ("L", "L_D", "R"):
                continue
            found.append({
                "clave": str(d.get("A", "")), "descripcion": str(d.get("H", "")),
                "barra": f"{d.get('B', '')} {d.get('C', '')} kV",
                "suministrador": str(d.get("F", "")),
                "tipo": kind, "fila": idx, "kwh": max(0., -kwh),
                "origen": "Balance Valorizado",
            })
    total_kwh = sum(aggregate.values())
    if total_kwh <= 0:
        raise ValueError("No se obtuvieron retiros netos positivos L/L_D/R. Revisar el archivo.")
    rate = comp / (total_kwh / 1000.0)
    warnings = [
        "El monto se basa en el total mensual agregado 'Compensación precio nudo/estabilizado' del Cuadro de Energía, que no identifica por sí solo el componente exclusivo PMGD ni la asignación horaria de EFE.",
        "La proyección usa una tasa media simple del SEN sobre el retiro mensual de EFE. No equivale al cargo oficial por retiro: puede variar por asignación horaria, balance definitivo, reliquidaciones y contrato.",
    ]
    if not found:
        warnings.append("No se encontraron las claves de retiro de EFE. Introduce el consumo manualmente o corrige las claves.")
    if abs(extra_ingresos - comp) > max(1., comp * 0.001):
        warnings.append("El monto absoluto de M no coincide con la contrapartida L; comprobar cierre del balance.")
    if any(not "ENEL GENERACION" in normalize(item["suministrador"]) for item in found):
        warnings.append("Una o más claves seleccionadas no figuran con Enel Generación como suministrador; revisa la correspondencia comercial.")
    return BalanceResult(
        label=month_from_filename(filename), compensacion_sen_clp=comp,
        retiro_sen_mwh=total_kwh/1000, tarifa_clp_mwh=rate,
        consumo_efe_kwh=sum(f["kwh"] for f in found), detalle=found,
        retiros_por_tipo={k:v/1000 for k,v in aggregate.items()},
        metodo=f"Cuadro de Energía!M{totalrow} (valor absoluto) / retiros netos J=L, L_D, R de Balance Valorizado",
        avisos=warnings,
    )


def parse_invoice(data: bytes, filename: str) -> dict:
    x = WorkbookReader(data)
    month = None
    lines: dict[str, dict] = {}
    for i, d in x.sheet("Cargos_CEN"):
        if i == 2 and isinstance(d.get("D"), (float, int)):
            month = excel_date(float(d["D"])).strftime("%Y-%m")
        key = normalize(d.get("B"))
        if not key.startswith("MERVAL"):
            continue
        name = str(d.get("B", ""))
        if name not in lines:
            lines[name] = {"punto": name, "kwh": 0., "cargo_clp": 0.}
        lines[name]["kwh"] += max(0., num(d.get("G")))
        lines[name]["cargo_clp"] += num(d.get("N"))
    kwh = sum(x["kwh"] for x in lines.values())
    cargo = sum(x["cargo_clp"] for x in lines.values())
    if not lines or kwh <= 0:
        raise ValueError("No se localizaron consumos/cargos MERVAL en la hoja Cargos_CEN.")
    return {"archivo":filename,"mes_balance":month or month_from_filename(filename),
            "consumo_kwh":kwh,"compensacion_clp":cargo,"tarifa_clp_kwh":cargo/kwh,
            "puntos":list(lines.values()),"fuente":"Cargos_CEN: columna G (kWh), N ($)"}


def parse_price_quillota(data: bytes) -> list[dict]:
    x = WorkbookReader(data)
    prices = []
    for r, d in x.sheet("DECRETOS PESTABILIZADO"):
        if normalize(d.get("B")) == "QUILLOTA" and num(d.get("C")) == 220:
            prices.append({"bloque":d.get("D", ""),"valor_clp_kwh":num(d.get("F")),
                           "decreto":str(d.get("G", ""))})
    return prices


def parse_hourly_csv(data: bytes) -> dict:
    """CSV permite retiro_quilpue_mwh + retiro_limache_mwh o el antiguo retiro_efe_mwh."""
    decoded = None
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            decoded = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if not decoded:
        raise ValueError("No se pudo decodificar CSV.")
    dialect = csv.Sniffer().sniff(decoded[:3000], delimiters=",;\t")
    table = csv.DictReader(StringIO(decoded), dialect=dialect)
    headers = {normalize(k).lower() for k in (table.fieldnames or [])}
    required = {"fecha", "hora", "compensacion_sen_clp", "retiro_sen_mwh"}
    if not required.issubset(headers):
        raise ValueError("Faltan fecha, hora, compensacion_sen_clp o retiro_sen_mwh.")
    split = {"retiro_quilpue_mwh", "retiro_limache_mwh"}.issubset(headers)
    if not split and "retiro_efe_mwh" not in headers:
        raise ValueError("Faltan retiro_quilpue_mwh y retiro_limache_mwh (o retiro_efe_mwh).")
    totals = {s: {"mwh": 0.0, "clp": 0.0} for s in SITE_NAMES} if split else {"EFE sin desglose": {"mwh": 0., "clp": 0.}}
    output = []
    n = missing = 0
    for entry in table:
        item = {normalize(k).lower():v for k,v in entry.items()}
        if not any(str(v or "").strip() for v in item.values()):
            continue
        total_sen = num(item["retiro_sen_mwh"])
        if total_sen <= 0:
            missing += 1
            continue
        tariff = num(item["compensacion_sen_clp"]) / total_sen
        consumptions = ({"Quilpué": num(item["retiro_quilpue_mwh"]),
                         "Limache": num(item["retiro_limache_mwh"])} if split
                        else {"EFE sin desglose": num(item["retiro_efe_mwh"])})
        for name, energy in consumptions.items():
            charge = energy * tariff
            totals[name]["mwh"] += energy
            totals[name]["clp"] += charge
            output.append({"fecha": item["fecha"], "hora": item["hora"], "punto": name,
                           "tarifa_sen_clp_mwh": tariff, "retiro_mwh": energy,
                           "cargo_clp": charge})
        n += 1
    if not n:
        raise ValueError("No hay horas válidas en el archivo.")
    return {"clp": sum(v["clp"] for v in totals.values()),
            "mwh": sum(v["mwh"] for v in totals.values()),
            "horas": n, "omitidas": missing, "detalle": output, "puntos": totals, "separado": split}


def csv_bytes(rows: list[dict]) -> bytes:
    if not rows:return b""
    buf = StringIO()
    writer = csv.DictWriter(buf, fieldnames=list(rows[0].keys()), delimiter=";")
    writer.writeheader()
    for row in rows:writer.writerow(row)
    return "\ufeff".encode("utf-8") + buf.getvalue().encode("utf-8")



# Separación de puntos de retiro comercial utilizados por EFE.
SITE_NAMES = ("Quilpué", "Limache")
MONTHS_ES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
             "agosto", "septiembre", "octubre", "noviembre", "diciembre")


def following_month(yyyymm: str) -> str:
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", str(yyyymm)):
        return "Sin período identificado"
    year, month = (int(v) for v in yyyymm.split("-"))
    return f"{year + (month == 12):04d}-{month % 12 + 1:02d}"


def readable_month(yyyymm: str) -> str:
    try:
        year, month = (int(v) for v in yyyymm.split("-"))
        return f"{MONTHS_ES[month - 1].capitalize()} de {year}"
    except (ValueError, IndexError, AttributeError):
        return "Período no identificado"


def balance_consumption_by_site(balance: BalanceResult) -> dict[str, float]:
    amounts = {s: 0.0 for s in SITE_NAMES}
    for entry in balance.detalle:
        key = normalize(entry.get("clave"))
        desc = normalize(entry.get("descripcion"))
        if key == "MERVAL_ELSOL" or "QUILPUE" in key or "QUILPUE" in desc:
            amounts["Quilpué"] += float(entry["kwh"])
        elif key == "MERVAL_LIMACHE" or "LIMACHE" in key or "LIMACHE" in desc:
            amounts["Limache"] += float(entry["kwh"])
    return amounts


def invoice_data_by_site(invoice: dict) -> dict[str, dict]:
    values = {s: {"kwh": 0.0, "cargo_clp": 0.0, "tarifa_clp_mwh": 0.0} for s in SITE_NAMES}
    for item in invoice.get("puntos", []):
        name = normalize(item.get("punto"))
        site = "Quilpué" if "QUILPUE" in name else "Limache" if "LIMACHE" in name else None
        if site:
            values[site]["kwh"] += float(item["kwh"])
            values[site]["cargo_clp"] += float(item["cargo_clp"])
    for site in SITE_NAMES:
        x = values[site]
        if x["kwh"]:
            x["tarifa_clp_mwh"] = x["cargo_clp"] / x["kwh"] * 1000
    return values


def projected_by_site(kwh: dict[str, float], rates: dict[str, float],
                      transfer_pct: float = 100.0, adjustments: dict[str, float] | None = None,
                      sensitivity_pct: float = 20.0) -> list[dict]:
    adjustments = adjustments or {}
    result = []
    for site in SITE_NAMES:
        energy = float(kwh.get(site, 0))
        rate = float(rates.get(site, 0))
        adjustment = float(adjustments.get(site, 0))
        base = energy / 1000.0 * rate * transfer_pct / 100.0
        low, high = sorted((base * (1 - sensitivity_pct/100) + adjustment,
                            base * (1 + sensitivity_pct/100) + adjustment))
        result.append({"punto": site, "kwh": energy, "mwh": energy/1000,
                       "tarifa_clp_mwh": rate, "tarifa_clp_kwh": rate/1000,
                       "base_clp": base, "ajuste_clp": adjustment,
                       "cargo_clp": base + adjustment, "bajo_clp": low, "alto_clp": high})
    return result


import streamlit as st

st.set_page_config(page_title="Compensación PE | EFE Valparaíso", page_icon="⚡", layout="wide")
st.markdown("""<style>
.block-container{max-width:1250px;padding-top:1.5rem}
h1{letter-spacing:-.04em}
.stMetric{border:1px solid #dae2ed;border-radius:12px;padding:14px}
</style>""", unsafe_allow_html=True)
st.title("⚡ Compensación por Precio Estabilizado")
st.caption("EFE Valparaíso · Suministros Quilpué y Limache por separado · Versión 1.2 Cloud")

with st.sidebar:
    st.header("1 · Subir archivos")
    balance_uploads = st.file_uploader("Balances CEN", type=["xlsx", "xlsm"],
        accept_multiple_files=True, help="Ejemplo: Balance_2609P.xlsm, balance de septiembre de 2026.")
    invoice_uploads = st.file_uploader("Proformas / facturas Enel", type=["xlsx", "xlsm"],
        accept_multiple_files=True, help="Archivo EFE_VALPO_SEP2026.xlsx con hoja Cargos_CEN.")
    price_upload = st.file_uploader("Precios estabilizados (opcional)", type=["xlsx", "xlsm"])
    st.divider()
    st.header("2 · Retiros EFE")
    clave_text = st.text_area("Claves en Balance Valorizado (una por línea)",
        "MERVAL_ELSOL\nMERVAL_LIMACHE", height=90,
        help="MERVAL_ELSOL corresponde a Quilpué y MERVAL_LIMACHE a Limache.")
    keys = [k.strip() for k in clave_text.replace(",", "\n").splitlines() if k.strip()]
    st.caption("Las planillas se leen en el servidor de la aplicación. No se ejecutan macros.")

@st.cache_data(show_spinner=False)
def calc_balance(data: bytes, name: str, keys_tuple: tuple[str, ...]):
    return parse_balance(data, name, list(keys_tuple))

@st.cache_data(show_spinner=False)
def calc_invoice(data: bytes, name: str):
    return parse_invoice(data, name)

balances, invoices, errors = [], [], []
for item in balance_uploads or []:
    try:
        balances.append(calc_balance(item.getvalue(), item.name, tuple(keys)))
    except Exception as exc:
        errors.append(f"Balance {item.name}: {exc}")
for item in invoice_uploads or []:
    try:
        invoices.append(calc_invoice(item.getvalue(), item.name))
    except Exception as exc:
        errors.append(f"Factura {item.name}: {exc}")
balances.sort(key=lambda b: b.label)
invoices.sort(key=lambda i: i["mes_balance"])
for error in errors:
    st.error(error)

def clp(number: float) -> str:
    return "$" + f"{number:,.0f}".replace(",", ".")

def number_es(number: float, decimals: int = 2) -> str:
    return f"{number:,.{decimals}f}".replace(",", "_").replace(".", ",").replace("_", ".")

if not balances and not invoices:
    st.info("Sube un balance CEN y, si quieres comparar, una proforma de Enel. "
            "También puedes indicar consumos y tarifas manualmente para Quilpué y Limache.")

calc_tab, history_tab, hourly_tab, method_tab = st.tabs([
    "📊 Proyección por punto", "📈 Histórico por punto", "🕒 Cálculo horario", "📖 Metodología"])

with calc_tab:
    st.subheader("Proyección de Compensación PE")
    chosen = None
    if balances:
        choice = st.selectbox("Selecciona el balance CEN de referencia", range(len(balances)),
            format_func=lambda ix: f"{readable_month(balances[ix].label)} · {number_es(balances[ix].tarifa_clp_mwh)} $/MWh",
            index=len(balances)-1)
        chosen = balances[choice]

    # El mes del balance no es necesariamente el de la factura.
    if chosen and re.fullmatch(r"\d{4}-\d{2}", chosen.label):
        default_period = chosen.label
    elif invoices and re.fullmatch(r"\d{4}-\d{2}", invoices[-1]["mes_balance"]):
        default_period = following_month(invoices[-1]["mes_balance"])
    else:
        default_period = f"{date.today().year}-{date.today().month:02d}"
    default_date = date.fromisoformat(default_period + "-01")
    period_date = st.date_input("Mes del balance que deseas proyectar", value=default_date,
        format="DD/MM/YYYY", key=f"periodo_{chosen.label if chosen else 'manual'}")
    period = f"{period_date.year:04d}-{period_date.month:02d}"
    invoice_month = following_month(period)
    mcol1, mcol2 = st.columns(2)
    mcol1.info(f"**Mes del balance CEN:** {readable_month(period)}")
    mcol2.info(f"**Mes de la factura proyectada:** {readable_month(invoice_month)}")
    if chosen and chosen.label != period:
        st.warning(f"Estás usando una tasa del balance de **{readable_month(chosen.label)}** para proyectar "
                   f"**{readable_month(period)}**. Es una extrapolación, no el valor publicado del mes objetivo.")

    latest_invoice = invoices[-1] if invoices else None
    latest_invoice_data = invoice_data_by_site(latest_invoice) if latest_invoice else None
    if latest_invoice and latest_invoice_data:
        st.caption(f"Última factura subida: {readable_month(following_month(latest_invoice['mes_balance']))} "
                   f"(balance {readable_month(latest_invoice['mes_balance'])}).")

    options = []
    if chosen:
        options.append("Promedio mensual SEN obtenido del balance CEN")
    if latest_invoice:
        options.append("Tarifa efectiva de última factura (separada por punto)")
    options.append("Ingresar tarifas manuales por punto")
    rate_source = st.radio("Origen de la compensación unitaria", options)
    if rate_source.startswith("Promedio"):
        rates = {s: chosen.tarifa_clp_mwh for s in SITE_NAMES}
    elif rate_source.startswith("Tarifa efectiva"):
        rates = {s: latest_invoice_data[s]["tarifa_clp_mwh"] for s in SITE_NAMES}
    else:
        st.markdown("**Compensación unitaria a proyectar ($/MWh)**")
        rate_columns = st.columns(2)
        rates = {}
        for site, col in zip(SITE_NAMES, rate_columns):
            seed_rate = (chosen.tarifa_clp_mwh if chosen else
                (latest_invoice_data[site]["tarifa_clp_mwh"] if latest_invoice_data else 0.0))
            with col:
                rates[site] = st.number_input(f"Tarifa {site} ($/MWh)",
                    value=float(round(seed_rate, 3)), step=100.0, format="%.3f",
                    key=f"tasa_{site}_{period}")

    energy_options = []
    if chosen:
        energy_options.append("Retiros del balance CEN por punto")
    if latest_invoice:
        energy_options.append("Consumos de la última factura por punto")
    energy_options.append("Ingresar consumos manuales por punto")
    energy_source = st.radio("Origen del consumo", energy_options)
    balances_by_site = balance_consumption_by_site(chosen) if chosen else None
    if energy_source.startswith("Retiros"):
        usages = balances_by_site.copy()
    elif energy_source.startswith("Consumos"):
        usages = {s: latest_invoice_data[s]["kwh"] for s in SITE_NAMES}
    else:
        st.markdown("**Consumo estimado de cada punto (kWh)**")
        usage_columns = st.columns(2)
        usages = {}
        for site, col in zip(SITE_NAMES, usage_columns):
            seed = (balances_by_site[site] if balances_by_site else
                (latest_invoice_data[site]["kwh"] if latest_invoice_data else 0.0))
            with col:
                usages[site] = st.number_input(f"Consumo {site} (kWh)", min_value=0.0,
                    value=round(float(seed), 2), step=1000.0, format="%.2f",
                    key=f"consumo_{site}_{period}")
    if chosen and energy_source.startswith("Retiros"):
        detected = sum(usages.values())
        if abs(detected-chosen.consumo_efe_kwh) > 0.01:
            st.warning("Algunas claves del balance seleccionadas no corresponden a Quilpué o Limache. "
                       "Revisa la identificación de los retiros; el total EFE del balance no coincide con los puntos mostrados.")
        for site in SITE_NAMES:
            if not usages[site]:
                st.warning(f"No se encontró retiro para {site} en el balance seleccionado.")

    with st.expander("Ajustes, reliquidaciones y escenarios", expanded=False):
        transfer = st.number_input("Porcentaje de traslado contractual (%)", 0.0, 200.0, 100.0, 1.0)
        s1, s2 = st.columns(2)
        with s1:
            adjust_quilpue = st.number_input("Ajuste Quilpué ($)", value=0.0, step=10000.0, format="%.0f")
        with s2:
            adjust_limache = st.number_input("Ajuste Limache ($)", value=0.0, step=10000.0, format="%.0f")
        sensitivity = st.slider("Variación escenarios bajo/alto (%)", 0, 80, 20, 5)
    rows = projected_by_site(usages, rates, transfer,
        {"Quilpué": adjust_quilpue, "Limache": adjust_limache}, sensitivity)
    total_charge = sum(x["cargo_clp"] for x in rows)
    total_energy = sum(x["kwh"] for x in rows)
    total_low = sum(x["bajo_clp"] for x in rows)
    total_high = sum(x["alto_clp"] for x in rows)

    st.markdown(f"#### Resultado · Factura de {readable_month(invoice_month)}")
    col1, col2, col3 = st.columns(3)
    col1.metric("Compensación Quilpué", clp(rows[0]["cargo_clp"]))
    col2.metric("Compensación Limache", clp(rows[1]["cargo_clp"]))
    col3.metric("Total EFE (sin IVA)", clp(total_charge))

    detail_rows = [{"Mes balance": readable_month(period),
                    "Mes factura": readable_month(invoice_month), "Punto": x["punto"],
                    "Consumo (kWh)": round(x["kwh"],2),
                    "Consumo (MWh)": round(x["mwh"],4),
                    "Compensación ($/MWh)": round(x["tarifa_clp_mwh"],3),
                    "Compensación ($/kWh)": round(x["tarifa_clp_kwh"],6),
                    "Ajuste ($)": round(x["ajuste_clp"]),
                    "Cargo proyectado ($)": round(x["cargo_clp"])} for x in rows]
    weighted_rate = (sum(x["kwh"] * x["tarifa_clp_mwh"] for x in rows) / total_energy
                     if total_energy > 0 else 0.0)
    detail_rows.append({"Mes balance": readable_month(period),
                        "Mes factura": readable_month(invoice_month),
                        "Punto": "TOTAL EFE", "Consumo (kWh)": round(total_energy,2),
                        "Consumo (MWh)": round(total_energy/1000,4),
                        "Compensación ($/MWh)": round(weighted_rate,3),
                        "Compensación ($/kWh)": round(weighted_rate/1000,6),
                        "Ajuste ($)": round(adjust_quilpue+adjust_limache),
                        "Cargo proyectado ($)": round(total_charge)})
    st.dataframe(detail_rows, hide_index=True, use_container_width=True)
    st.caption("Los importes corresponden al balance seleccionado o a los supuestos ingresados. "
               "No constituyen una asignación oficial individual del CEN.")

    st.markdown("**Escenarios por punto, sin IVA**")
    scenario = [{"Punto": x["punto"], "Bajo ($)": round(x["bajo_clp"]),
                 "Medio ($)": round(x["cargo_clp"]), "Alto ($)": round(x["alto_clp"])} for x in rows]
    scenario.append({"Punto": "TOTAL EFE", "Bajo ($)": round(total_low),
                     "Medio ($)": round(total_charge), "Alto ($)": round(total_high)})
    st.dataframe(scenario, hide_index=True, use_container_width=True)

    if chosen:
        with st.expander("Trazabilidad del balance y claves detectadas"):
            a, b = st.columns(2)
            a.metric("Compensación agregada SEN ($)", clp(chosen.compensacion_sen_clp))
            b.metric("Retiros SEN (MWh)", number_es(chosen.retiro_sen_mwh,3))
            st.write("**Período de la fuente:**", readable_month(chosen.label))
            st.dataframe(chosen.detalle, hide_index=True, use_container_width=True)
            for note in chosen.avisos:
                st.warning(note)
    else:
        st.warning("Sin balance CEN. El resultado es referencial y depende exclusivamente de las tarifas y consumos ingresados o históricos.")
    st.download_button("⬇ Descargar cálculo separado en CSV",
        data=csv_bytes([{
            "mes_balance": period, "mes_factura":invoice_month,
            "mes_referencia_tarifa": chosen.label if chosen and rate_source.startswith("Promedio") else
                (latest_invoice["mes_balance"] if latest_invoice and rate_source.startswith("Tarifa efectiva") else "Manual"),
            "fuente_tarifa": rate_source, "fuente_consumo": energy_source,
            "punto": x["punto"], "retiro_kwh": x["kwh"],
            "tarifa_clp_mwh": x["tarifa_clp_mwh"],
            "factor_traslado_pct": transfer, "ajuste_clp": x["ajuste_clp"],
            "cargo_estimado_clp": round(x["cargo_clp"]),
            "escenario_bajo_clp": round(x["bajo_clp"]), "escenario_alto_clp": round(x["alto_clp"]),
            "condicion": "Proyección, no asignación oficial individual del CEN",
        } for x in rows] + [{
            "mes_balance": period, "mes_factura": invoice_month,
            "mes_referencia_tarifa": chosen.label if chosen else "Manual",
            "fuente_tarifa": rate_source, "fuente_consumo": energy_source,
            "punto": "TOTAL EFE", "retiro_kwh": total_energy,
            "tarifa_clp_mwh": "Promedio ponderado",
            "factor_traslado_pct": transfer, "ajuste_clp": adjust_quilpue+adjust_limache,
            "cargo_estimado_clp": round(total_charge),
            "escenario_bajo_clp": round(total_low), "escenario_alto_clp": round(total_high),
            "condicion": "Proyección, no asignación oficial individual del CEN",
        }]), file_name=f"compensacion_PE_Quilpue_Limache_balance_{period}_factura_{invoice_month}.csv",
        mime="text/csv")

with history_tab:
    st.subheader("Histórico separado de Quilpué y Limache")
    if invoices:
        table = []
        for invoice in invoices:
            per_site = invoice_data_by_site(invoice)
            for site in SITE_NAMES:
                x = per_site[site]
                table.append({"Balance CEN": readable_month(invoice["mes_balance"]),
                              "Factura Enel": readable_month(following_month(invoice["mes_balance"])),
                              "Punto": site, "Consumo (kWh)": round(x["kwh"],2),
                              "Cargo facturado ($)": round(x["cargo_clp"]),
                              "Tasa efectiva ($/kWh)": round(x["tarifa_clp_mwh"]/1000,6)})
            table.append({"Balance CEN": readable_month(invoice["mes_balance"]),
                          "Factura Enel": readable_month(following_month(invoice["mes_balance"])),
                          "Punto": "TOTAL EFE", "Consumo (kWh)": round(invoice["consumo_kwh"],2),
                          "Cargo facturado ($)": round(invoice["compensacion_clp"]),
                          "Tasa efectiva ($/kWh)": round(invoice["tarifa_clp_kwh"],6)})
        st.markdown("**Cobros efectivamente incluidos en proformas / facturas cargadas**")
        st.dataframe(table, hide_index=True, use_container_width=True)
        st.download_button("Descargar histórico de facturas CSV", csv_bytes(table),
                           "historico_facturas_quilpue_limache.csv", "text/csv")
    else:
        st.info("Carga una proforma Enel para ver los cargos facturados por punto.")

    if balances:
        table = []
        for b in balances:
            split = balance_consumption_by_site(b)
            for site in SITE_NAMES:
                amount = split[site] / 1000 * b.tarifa_clp_mwh
                table.append({"Balance CEN": readable_month(b.label),
                              "Factura estimada": readable_month(following_month(b.label)),
                              "Punto": site, "Retiro (kWh)": round(split[site],2),
                              "Tasa media SEN ($/MWh)": round(b.tarifa_clp_mwh,3),
                              "Proyección ($)": round(amount)})
            table.append({"Balance CEN": readable_month(b.label),
                          "Factura estimada": readable_month(following_month(b.label)),
                          "Punto": "TOTAL EFE", "Retiro (kWh)": round(sum(split.values()),2),
                          "Tasa media SEN ($/MWh)": round(b.tarifa_clp_mwh,3),
                          "Proyección ($)": round(sum(split.values())/1000*b.tarifa_clp_mwh)})
        st.markdown("**Proyecciones elaboradas a partir de los balances CEN**")
        st.dataframe(table, hide_index=True, use_container_width=True)
        st.download_button("Descargar histórico de balances CSV", csv_bytes(table),
                           "historico_balances_quilpue_limache.csv", "text/csv")
    if invoices and balances:
        st.markdown("**Comparación de proyección vs. factura por punto cuando coinciden los meses**")
        actuals = {item["mes_balance"]: item for item in invoices}
        comparisons = []
        for b in balances:
            if b.label not in actuals:
                continue
            actual_by_site = invoice_data_by_site(actuals[b.label])
            projected_energy = balance_consumption_by_site(b)
            for site in SITE_NAMES:
                estimated = projected_energy[site] / 1000 * b.tarifa_clp_mwh
                real = actual_by_site[site]["cargo_clp"]
                comparisons.append({"Balance CEN": readable_month(b.label), "Punto": site,
                                    "Factura ($)": round(real), "Proyección ($)":round(estimated),
                                    "Diferencia ($)":round(estimated-real)})
        if comparisons:
            st.dataframe(comparisons, hide_index=True, use_container_width=True)
        else:
            st.caption("Los balances y facturas cargados corresponden a meses distintos; no se comparan indebidamente.")

with hourly_tab:
    st.subheader("Compensación horaria por punto")
    st.write("Carga un CSV con la compensación y el retiro total del SEN, más las energías horarias de "
             "Quilpué y Limache por separado. La aplicación calcula el cargo para cada punto y su total.")
    st.download_button("⬇ Descargar plantilla horaria CSV", data=csv_bytes([{
        "fecha":"2026-09-01", "hora":"00:00", "compensacion_sen_clp":"",
        "retiro_sen_mwh":"", "retiro_quilpue_mwh":"", "retiro_limache_mwh":""
    }]), file_name="plantilla_horaria_quilpue_limache.csv", mime="text/csv")
    hourly_file = st.file_uploader("Subir CSV horario", type=["csv"])
    if hourly_file:
        try:
            result = parse_hourly_csv(hourly_file.getvalue())
            st.metric("Compensación horaria calculada, total ($)", clp(result["clp"]))
            st.caption(f"Horas procesadas: {result['horas']}; horas omitidas: {result['omitidas']}")
            overview = [{"Punto": name, "Retiro (MWh)": round(x["mwh"],3),
                         "Compensación ($)": round(x["clp"])} for name,x in result["puntos"].items()]
            st.dataframe(overview, hide_index=True, use_container_width=True)
            st.dataframe(result["detalle"], hide_index=True, use_container_width=True)
            st.download_button("Descargar desglose horario CSV", csv_bytes(result["detalle"]),
                               "desglose_horario_quilpue_limache.csv", "text/csv")
            if not result["separado"]:
                st.warning("Este CSV sólo contiene retiro total EFE. Para desglosar Quilpué y Limache "
                           "utiliza la nueva plantilla con ambas columnas.")
        except Exception as exc:
            st.error(f"No se pudo procesar el archivo horario: {exc}")
    st.caption("Las compensaciones horarias deben corresponder a los montos asignables a retiros "
               "según el CEN. Sin verificar ese supuesto, el resultado sigue siendo indicativo.")

with method_tab:
    st.subheader("Metodología y trazabilidad")
    st.latex(r"T_{PE,mes}=\frac{|C_{SEN,mes}|}{R_{SEN,mes}}\quad [\mathrm{CLP/MWh}]")
    st.latex(r"Cargo_{s,proyectado}=\frac{E_{s,mes}[\mathrm{kWh}]}{1000}\times T_{s,mes}[\mathrm{CLP/MWh}]")
    st.markdown("""
**Puntos mostrados:** `MERVAL_ELSOL` se agrupa en **Quilpué** y `MERVAL_LIMACHE` se
agrupa en **Limache**. En el archivo histórico de Enel los nombres son `MERVAL QUILPUE` y
`MERVAL LIMACHE`. Los dos cargos se calculan independientemente y luego se suman.

**Períodos:** la pantalla y los archivos descargados separan el **mes del balance CEN** y el
**mes de facturación Enel**. Se utiliza un mes de desfase, tal como indicó el suministrador;
si cambian las condiciones contractuales, revisa esa hipótesis.

**Balance CEN:** se usa la magnitud de `Cuadro de Energía`, columna M, fila `Total general`,
y el volumen neto de retiros tipificados `L`, `L_D`, `R` en `Balance Valorizado`.
La tasa resultante es un promedio agregado de presupuesto, **no** necesariamente el valor
oficial de compensación PMGD o el cargo por retiro adjudicado a EFE.

**Factura histórica:** hoja `Cargos_CEN`, columna G para kWh y columna N para cargo ($).
La última tasa efectiva de facturación se calcula **por separado** para Quilpué y Limache.

**Precaución:** el balance preliminar puede estar sujeto a reliquidaciones. Además, el
componente de la columna M puede incluir valorizaciones adicionales al cargo PMGD.
No usar el resultado para afirmar que el CEN asignó ese monto específico a EFE.

**Privacidad:** los archivos Excel se procesan en el servidor de Streamlit donde alojes la
aplicación, sin ejecutar macros. Asegúrate de contar con autorización empresarial para su
procesamiento en servicios externos.
""")
    if price_upload:
        try:
            prices = parse_price_quillota(price_upload.getvalue())
            if prices:
                st.markdown("**Precios estabilizados Quillota 220 kV (referencia documental)**")
                st.dataframe(prices, hide_index=True, use_container_width=True)
                st.info("Son precios estabilizados de generación, no compensaciones unitarias para retiro EFE.")
            else:
                st.warning("No se encontraron valores de Quillota 220 kV en el archivo cargado.")
        except Exception as exc:
            st.error(f"No pude procesar precios estabilizados: {exc}")
