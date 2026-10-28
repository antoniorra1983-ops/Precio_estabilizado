"""Aplicación web Compensación PE EFE Valparaíso, versión Cloud autocontenida 2.16: fuentes estrictas Datos/ o cargas en Streamlit; filtros sin períodos ficticios.

Motor de cálculo y pantalla integrados. No requiere importar cen_core.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import date, timedelta
from io import BytesIO, StringIO
from pathlib import Path
from functools import lru_cache
import csv
import re
import unicodedata
import json
import os
from urllib.request import Request, urlopen
from urllib.parse import quote
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
    retiro_enel_mwh: float
    retiros_enel_por_tipo: dict[str, float]
    metodo: str
    avisos: list[str] = field(default_factory=list)
    sobrecostos_enel_cen: dict = field(default_factory=dict)

    @property
    def proyeccion_clp(self) -> float:
        return self.consumo_efe_kwh / 1000.0 * self.tarifa_clp_mwh


def month_from_filename(filename: str) -> str:
    name = filename or ""
    m = re.search(r"(?<!\d)(\d{2})(0[1-9]|1[0-2])[PD]?(?!\d)", name.upper())
    if m:
        return f"20{m.group(1)}-{m.group(2)}"
    m_ym = re.search(r"(20\d{2})[-_ ](0[1-9]|1[0-2])(?:[^0-9]|$)", name)
    if m_ym:
        return f"{m_ym.group(1)}-{m_ym.group(2)}"
    nombres = {"ENE":1,"FEB":2,"MAR":3,"ABR":4,"MAY":5,"JUN":6,"JUL":7,
               "AGO":8,"SEP":9,"OCT":10,"NOV":11,"DIC":12}
    m_texto = re.search(r"(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)[A-Z]*[ _-]?(20\d{2}|\d{2})(?!\d)",name.upper())
    if m_texto:
        year = int(m_texto.group(2))
        return f"{year if year >= 2000 else 2000+year:04d}-{nombres[m_texto.group(1)]:02d}"
    return "Sin período identificado"


def parse_balance(data: bytes, filename: str, keys: list[str]) -> BalanceResult:
    reader = WorkbookReader(data)
    mvals = []
    enel_costos = []
    for idx, d in reader.sheet("Cuadro de Energía"):
        if normalize(d.get("A")) == "TOTAL GENERAL":
            mvals.append((idx, num(d.get("L")), num(d.get("M"))))
        if normalize(d.get("B")) == "ENEL_GENERACION":
            enel_costos.append({"fila": idx, "pd_neto_clp": num(d.get("O")),
                                "sobrecostos_energia_netos_clp": num(d.get("P"))})
    if len(mvals) != 1:
        raise ValueError("No se encontró una única fila 'Total general' en 'Cuadro de Energía'.")
    if len(enel_costos) != 1:
        raise ValueError("No se encontró una única fila ENEL_GENERACION en Cuadro de Energía.")
    totalrow, extra_ingresos, comp_signed = mvals[0]
    if abs(comp_signed) < 1:
        raise ValueError("La compensación agregada es cero/no está calculada; verificar datos guardados en Excel.")
    # En los archivos CEN observados, M es compensación con signo negativo para generadores
    # y L es contrapartida positiva. Para presupuesto de retiros se toma magnitud.
    comp = abs(comp_signed)
    aggregate = {"L": 0.0, "L_D": 0.0, "R": 0.0}
    enel_retiros = {"L": 0.0, "L_D": 0.0, "R": 0.0}
    wanted = {normalize(k): k for k in keys if str(k).strip()}
    found = []
    for idx, d in reader.sheet("Balance Valorizado"):
        if idx == 1:
            continue
        kind = normalize(d.get("J"))
        kwh = num(d.get("I"))
        if kind in aggregate:
            aggregate[kind] -= kwh  # signo CEN: los retiros son negativos
            if normalize(d.get("L")) == "ENEL_GENERACION":
                enel_retiros[kind] -= kwh  # solo retiros atribuidos a Enel Generación
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
    enel_total_kwh = sum(enel_retiros.values())
    if enel_total_kwh <= 0:
        raise ValueError("No se encontraron retiros positivos de ENEL_GENERACION en Balance Valorizado, columna L.")
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
        retiro_enel_mwh=enel_total_kwh / 1000.0,
        retiros_enel_por_tipo={k:v / 1000 for k,v in enel_retiros.items()},
        metodo=f"Cuadro de Energía!M{totalrow} (valor absoluto) / retiros netos J=L, L_D, R de Balance Valorizado",
        avisos=warnings,
        sobrecostos_enel_cen=enel_costos[0],
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


def mes_corto_grafico(yyyymm: str) -> str:
    """Etiqueta legible de tres letras más año; conserva el orden mediante YYYY-MM."""
    if not re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])", str(yyyymm)):
        return "Período no identificado"
    anio, mes = (int(v) for v in yyyymm.split("-"))
    return f"{MONTHS_ES[mes - 1][:3].capitalize()} {anio}"


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



# --------------------------------------------------
# Servicios Complementarios: pagos por retiro CEN.
# Los importes del archivo 9 son asignaciones individuales, NO una tarifa. BESS se agrupa en SSCC.
SSCC_CODES = ("SC_SSCC", "BESS", "CO", "CO_ERNC", "CPF(+)", "CPF(-)",
              "CSF(+)", "CSF(-)", "CTF(+)", "CTF(-)", "CRA", "REA")
SSCC_KEYS = {"Quilpué": "MERVAL_ELSOL", "Limache": "MERVAL_LIMACHE"}
SSCC_BARS = {"Quilpué": "QUILPUE", "Limache": "S.PEDRO"}

# BESS es un componente de Servicios Complementarios y se presenta unido a SC_SSCC.
# Los códigos originales se conservan en el detalle de auditoría para comprobar el origen.
SSCC_INCLUIDOS = ("SC_SSCC", "BESS")
SSCC_CATEGORIAS = ("SSCC",) + tuple(c for c in SSCC_CODES if c not in SSCC_INCLUIDOS)


def agrupar_sscc(reporte: dict) -> dict[str, dict[str, float]]:
    """Agrupa SC_SSCC y BESS como un único SSCC; nunca contabiliza BESS dos veces."""
    agrupado = {}
    for sitio in SITE_NAMES:
        pagos = reporte["resumen"][sitio]
        agrupado[sitio] = {"SSCC": pagos["SC_SSCC"] + pagos["BESS"]}
        agrupado[sitio].update({c: pagos[c] for c in SSCC_CATEGORIAS if c != "SSCC"})
        if abs(sum(agrupado[sitio].values()) - sum(pagos.values())) > 0.01:
            raise ValueError(f"La conciliación de conceptos falló para {sitio}.")
    return agrupado


def exportar_sscc_agrupado(reporte: dict) -> list[dict]:
    """CSV principal: una sola fila SSCC por suministro (SC_SSCC + BESS)."""
    salidas = []
    periodo = reporte["periodo"]
    agrupado = agrupar_sscc(reporte)
    for sitio in SITE_NAMES:
        for codigo in SSCC_CATEGORIAS:
            origenes = SSCC_INCLUIDOS if codigo == "SSCC" else (codigo,)
            items = [x for x in reporte["detalle"] if x["punto"] == sitio and x["codigo"] in origenes]
            salidas.append({
                "mes_balance": periodo,
                "mes_factura": following_month(periodo),
                "suministrador": "ENEL_GENERACION",
                "punto": sitio,
                "clave": SSCC_KEYS[sitio],
                "concepto": codigo,
                "incluye": "+".join(origenes),
                "monto_clp": round(agrupado[sitio][codigo], 5),
                "archivo_cen": reporte["archivo"],
                "filas_origen": ", ".join(str(x["fila"]) for x in items),
            })
    return salidas


def parse_sscc_retiros(data: bytes, filename: str) -> dict:
    reader = WorkbookReader(data)
    if not any(normalize(k) == 'PAGO_RETIRO' for k in reader.sheets):
        raise ValueError("Archivo no compatible: falta la hoja PAGO_RETIRO.")
    totals = {site: {code: 0.0 for code in SSCC_CODES} for site in SSCC_KEYS}
    items = []
    header_checked = False
    for rownum, row in reader.sheet('PAGO_RETIRO'):
        if rownum == 1:
            expected = {'E': 'SUMINISTRADOR', 'G': 'CLAVE', 'H': 'TIPO_PAGO', 'J': 'PAGO'}
            for col, name in expected.items():
                if normalize(row.get(col)) != name:
                    raise ValueError(f"Encabezado inesperado en {col}1; se esperaba {name}.")
            header_checked = True
            continue
        key = normalize(row.get('G'))
        site = next((s for s, k in SSCC_KEYS.items() if key == k), None)
        if site is None or normalize(row.get('E')) != 'ENEL_GENERACION':
            continue
        if normalize(row.get('B')) != SSCC_BARS[site]:
            raise ValueError(f"Barra inesperada para {site}, fila {rownum}: {row.get('B')}")
        code = normalize(row.get('H'))
        if code not in SSCC_CODES:
            raise ValueError(f"Tipo_pago desconocido para {site}, fila {rownum}: {code}")
        # Un PAGO realmente vacío no equivale a un cero oficial.
        raw_pago = row.get('J')
        if raw_pago is None or str(raw_pago).strip() == '':
            raise ValueError(f'Importe PAGO vacío en PAGO_RETIRO!J{rownum} ({site}, {code}).')
        value = num(raw_pago)
        totals[site][code] += value
        items.append({'punto':site, 'clave':key, 'barra':str(row.get('B')),
                      'codigo':code, 'valor_clp':value, 'fila':rownum,
                      'suministrador':str(row.get('E'))})
    if not header_checked:
        raise ValueError("No se pudo validar el encabezado de PAGO_RETIRO.")
    missing = [s for s in SSCC_KEYS if not any(x['punto'] == s for x in items)]
    if missing:
        raise ValueError('No se encontraron registros de: ' + ', '.join(missing))
    return {
        'periodo': month_from_filename(filename),
        'archivo': filename, 'resumen': totals, 'detalle': items,
        'subtotal_base': {s: d['SC_SSCC'] + d['BESS'] for s, d in totals.items()},
        'suma_tipos': {s: sum(d.values()) for s, d in totals.items()},
        'faltantes': {s: [c for c in SSCC_CODES if not any(x['punto']==s and x['codigo']==c for x in items)] for s in SSCC_KEYS},
    }


def parse_sscc_empresa(data: bytes, filename: str) -> dict:
    reader = WorkbookReader(data)
    result = {'periodo': month_from_filename(filename), 'archivo': filename}
    for sheetname, code in (('PAGO_EMPRESAS_SC_SSCC', 'SC_SSCC'), ('PAGO_EMPRESAS_CO', 'CO')):
        company_rows = []
        for idx, row in reader.sheet(sheetname):
            if normalize(row.get('A')) == 'ENEL_GENERACION':
                company_rows.append({'recibe': num(row.get('B')), 'paga': num(row.get('C')),
                                     'neto': num(row.get('D')), 'fila': idx})
        if len(company_rows) != 1:
            raise ValueError(f"Se esperaban 1 fila para Enel en {sheetname}; encontradas: {len(company_rows)}")
        result[code] = company_rows[0]
    return result


def parse_sscc_invoice(data: bytes, filename: str) -> dict:
    """Vincula la línea SSCC de la PROFORMA con el balance CEN previo.

    En EFE_VALPO_SEP2026.xlsx: Cargos_CEN!C2 = 01-08-2026 (balance),
    mientras la proforma es septiembre. Se tolera D2 para otras plantillas,
    pero nunca se inventa un período si ambas celdas están vacías.
    """
    reader = WorkbookReader(data)
    mon = None
    for idx, row in reader.sheet('Cargos_CEN'):
        if idx == 2:
            raw = row.get('C')
            if not isinstance(raw, (int, float)):
                raw = row.get('D')
            if isinstance(raw, (int, float)):
                mon = excel_date(raw).strftime('%Y-%m')
            break
    if mon is None:
        raise ValueError('No se pudo identificar el mes balance en Cargos_CEN!C2/D2.')
    mes_factura = following_month(mon)
    mes_archivo = month_from_filename(filename)
    if mes_archivo != 'Sin período identificado' and mes_archivo != mes_factura:
        raise ValueError(
            f'Períodos incompatibles en {filename}: balance {mon}, '
            f'factura esperada {mes_factura}, nombre indica {mes_archivo}.')
    amounts = {}
    for sheetname, site in (('PROFORMA-Q','Quilpué'), ('PROFORMA-L', 'Limache')):
        values = []
        for _, row in reader.sheet(sheetname):
            if normalize(row.get('D')) != 'SERVICIOS COMPLEMENTARIOS':
                continue
            raw = row.get('F')
            if raw is None or str(raw).strip() == '':
                raise ValueError(f'Importe SSCC vacío en {sheetname}, columna F.')
            values.append(num(raw))
        if len(values) != 1:
            raise ValueError('No hay una única línea de Servicios Complementarios en '+sheetname)
        amounts[site] = values[0]
    return {'periodo_balance':mon,'mes_factura':mes_factura,
            'importes':amounts,'archivo':filename}


from sscc_archivo4 import calcular_netos_empresa


def calcular_prorrateo_sscc(empresa: dict, balance: BalanceResult,
                            bess_unitario_clp_kwh: float = 0.0) -> dict:
    """Referencia EFE = (SSCC neto + CO neto) × participación en retiros Enel.

    NETO para costo = PAGA - RECIBE, manteniendo importes negativos.
    BESS opcional es una referencia histórica SEPARADA del Archivo 4.
    El resultado no es una asignación oficial del CEN y NO se suma al Archivo 9.
    """
    if empresa["periodo"] != balance.label:
        raise ValueError("El archivo 4 y el balance deben pertenecer al mismo mes.")
    if not isinstance(bess_unitario_clp_kwh, (float, int)) or bess_unitario_clp_kwh < 0:
        raise ValueError("La tasa de referencia BESS no puede ser negativa.")
    retiro_enel_kwh = balance.retiro_enel_mwh * 1000.0
    if retiro_enel_kwh <= 0:
        raise ValueError("Los retiros totales de Enel son cero.")
    saldos = calcular_netos_empresa(empresa)
    energies = balance_consumption_by_site(balance)
    details = []
    for site in SITE_NAMES:
        retiro_kwh = energies[site]
        participacion = retiro_kwh / retiro_enel_kwh
        sscc_neto = saldos["SC_SSCC"]["costo_neto"] * participacion
        co_neto = saldos["CO"]["costo_neto"] * participacion
        base_neta = sscc_neto + co_neto
        bess_estimado = retiro_kwh * bess_unitario_clp_kwh
        details.append({
            "punto": site,
            "clave": SSCC_KEYS[site],
            "retiro_kwh": retiro_kwh,
            "participacion_enel_pct": participacion * 100.0,
            "sscc_base_clp": sscc_neto,   # campo histórico: ahora costo neto, no PAGA bruto
            "sscc_neto_estimado_clp": sscc_neto,
            "co_neto_estimado_clp": co_neto,
            "sscc_co_neto_estimado_clp": base_neta,
            "bess_estimado_clp": bess_estimado,
            "sscc_estimado_clp": base_neta + bess_estimado,
        })
    return {
        "periodo": empresa["periodo"],
        "mes_factura": following_month(empresa["periodo"]),
        "archivo_4": empresa["archivo"],
        "pago_enel_sc_sscc_clp": saldos["SC_SSCC"]["paga"],
        "recibe_enel_sc_sscc_clp": saldos["SC_SSCC"]["recibe"],
        "costo_neto_sc_sscc_clp": saldos["SC_SSCC"]["costo_neto"],
        "paga_enel_co_clp": saldos["CO"]["paga"],
        "recibe_enel_co_clp": saldos["CO"]["recibe"],
        "costo_neto_co_clp": saldos["CO"]["costo_neto"],
        "costo_neto_total_clp": saldos["total_costo_neto"],
        "retiro_enel_mwh": balance.retiro_enel_mwh,
        "retiros_enel_por_tipo": balance.retiros_enel_por_tipo,
        "bess_unitario_clp_kwh": bess_unitario_clp_kwh,
        "bess_incluido_como_sup": bool(bess_unitario_clp_kwh),
        "detalle": details,
        "sscc_base_total": sum(x["sscc_neto_estimado_clp"] for x in details),
        "co_neto_total": sum(x["co_neto_estimado_clp"] for x in details),
        "sscc_co_neto_total": sum(x["sscc_co_neto_estimado_clp"] for x in details),
        "sscc_total_estimado": sum(x["sscc_estimado_clp"] for x in details),
    }


def tasa_bess_referencia_previa(reportes: list[dict], balances: list[BalanceResult],
                                periodo_objetivo: str) -> tuple[float, str]:
    """Usa solo un mes anterior, nunca el mes objetivo, evitando fuga de datos futuros."""
    alternativas = []
    for reporte in reportes:
        per = reporte["periodo"]
        if per >= periodo_objetivo or not re.fullmatch(r"\d{4}-\d{2}", per):
            continue
        b = next((b for b in balances if b.label == per), None)
        if not b:
            continue
        total_kwh = sum(balance_consumption_by_site(b).values())
        if total_kwh <= 0:
            continue
        tasa = sum(reporte["resumen"][s]["BESS"] for s in SITE_NAMES) / total_kwh
        alternativas.append((per, tasa))
    if not alternativas:
        return 0.0, "No se cargaron datos previos para calcular una tasa BESS histórica."
    per, tasa = max(alternativas, key=lambda item: item[0])
    return tasa, f"Tasa BESS histórica calculada desde el archivo 9 y balance de {readable_month(per)}."



# --------------------------------------------------
# Mínimos Técnicos (MT) y Partidas y Detenciones (PD)
# El archivo de facturación EFE permite conocer el costo histórico por punto.
# La tarifa efectiva histórica es sólo un indicador para presupuesto, NO una tarifa CEN.
MT_PD_CODES = {"MT": ("K", "Mínimos Técnicos"),
               "PD": ("M", "Partidas y Detenciones")}


def parse_mt_pd_invoice(data: bytes, filename: str) -> dict:
    """Lee MT (K), PD (M) y energía (G) de Cargos_CEN por sitio.

    Quilpué normalmente trae consumo en fila MERVALIMEL y cargos en fila
    MERVAL_ELSOL. Se consolidan por el nombre de suministro de la columna B.
    """
    reader = WorkbookReader(data)
    amounts = {s: {"kwh": 0.0, "MT": 0.0, "PD": 0.0, "filas": []}
               for s in SITE_NAMES}
    period = None
    checked = False
    found = set()
    started = False
    for rownum, row in reader.sheet("Cargos_CEN"):
        if rownum == 2 and isinstance(row.get("D"), (float, int)):
            period = excel_date(float(row["D"])).strftime("%Y-%m")
        if rownum == 4:
            if (normalize(row.get("K")) != "MINIMO TECNICO"
                    or normalize(row.get("M")) != "P_D"):
                raise ValueError("No se reconocen las cabeceras MT/PD en K4 y M4 de Cargos_CEN.")
            checked = True
            continue
        if rownum <= 5:
            continue
        name = normalize(row.get("B"))
        if started and name and not name.startswith("MERVAL "):
            break  # Fin del bloque de cargos por retiro; no leer las tablas posteriores.
        if not name.startswith("MERVAL "):
            continue
        started = True
        site = ("Quilpué" if "QUILPUE" in name
                else "Limache" if "LIMACHE" in name else None)
        if site is None:
            continue
        # Las claves IVTE identifican las filas que pertenecen a los suministros.
        key = normalize(row.get("D"))
        if site == "Quilpué" and key not in ("MERVAL_ELSOL", "MERVALIMEL"):
            continue
        if site == "Limache" and key != "MERVAL_LIMACHE":
            continue
        found.add(site)
        values = amounts[site]
        energy = max(0.0, num(row.get("G")))
        mt, pd = num(row.get("K")), num(row.get("M"))
        values["kwh"] += energy
        values["MT"] += mt
        values["PD"] += pd
        values["filas"].append({"fila": rownum, "clave": key,
                               "energia_kwh": energy, "mt_clp": mt, "pd_clp": pd})
    if not checked or not found:
        raise ValueError("No se halló un bloque MT/PD compatible en Cargos_CEN.")
    if not period:
        raise ValueError("No se encontró el mes de balance en Cargos_CEN!D2.")
    warnings = []
    for site in SITE_NAMES:
        if site not in found or amounts[site]["kwh"] <= 0:
            warnings.append(f"{site}: falta consumo histórico; no se puede obtener su tasa efectiva.")
        energy = amounts[site]["kwh"]
        amounts[site]["tarifas_clp_kwh"] = {
            code: amounts[site][code] / energy if energy else None
            for code in MT_PD_CODES
        }
    return {
        "archivo": filename, "mes_balance": period,
        "mes_factura": following_month(period),
        "puntos": amounts, "avisos": warnings,
        "fuente": "Cargos_CEN: D2 (mes), B/D (punto), G (kWh), K (MT), M (PD)",
    }


def calcular_mt_pd(historico: dict | None, retiros_kwh: dict[str, float],
                   tasas_clp_kwh: dict[str, dict[str, float]],
                   variacion_pct: dict[str, float], periodo: str,
                   sensibilidad_pct: float = 20.0) -> dict:
    """Proyección por consumo: retiro previsto x tasa histórica ajustada.

    Los resultados no representan el prorrateo real por servicio del CEN.
    """
    items = []
    for site in SITE_NAMES:
        kwh = float(retiros_kwh.get(site, 0.0))
        if kwh < 0:
            raise ValueError(f"El consumo proyectado de {site} no puede ser negativo.")
        for code, (_, label) in MT_PD_CODES.items():
            unit = float(tasas_clp_kwh[site][code])
            factor = 1.0 + float(variacion_pct.get(code, 0.0)) / 100.0
            rate = unit * factor
            cargo = kwh * rate
            low, high = sorted((cargo * (1-sensibilidad_pct/100.0),
                                cargo * (1+sensibilidad_pct/100.0)))
            items.append({"punto": site, "codigo": code, "concepto": label,
                          "retiro_kwh": kwh, "tasa_historica_clp_kwh": unit,
                          "variacion_pct": float(variacion_pct.get(code, 0.0)),
                          "tasa_proyectada_clp_kwh": rate,
                          "cargo_estimado_clp": cargo,
                          "escenario_bajo_clp": low, "escenario_alto_clp": high})
    return {"mes_balance": periodo, "mes_factura": following_month(periodo),
            "mes_fuente": historico["mes_balance"] if historico else "Manual",
            "archivo_fuente": historico["archivo"] if historico else "Manual",
            "detalle": items,
            "total_mt": sum(x["cargo_estimado_clp"] for x in items if x["codigo"] == "MT"),
            "total_pd": sum(x["cargo_estimado_clp"] for x in items if x["codigo"] == "PD"),
            "total": sum(x["cargo_estimado_clp"] for x in items)}



# --------------------------------------------------
# MT y PD: sin facturas. Referencia a balance CEN y asignación directa cuando exista.
# El archivo 'Balance_YYMMP' presenta dos saldos de EMPRESA (columnas O/P) con signo.
# No son cargos individuales; P NO es idéntico a Mínimos Técnicos.


def estimar_sobrecostos_cen(balance: BalanceResult) -> dict:
    if balance.retiro_enel_mwh <= 0:
        raise ValueError("Balance sin retiros válidos para Enel Generación.")
    totales = balance.sobrecostos_enel_cen
    if not totales:
        raise ValueError("El balance no contiene sobrecostos de ENEL_GENERACION.")
    energies = balance_consumption_by_site(balance)
    e_total = balance.retiro_enel_mwh * 1000.0
    records = []
    for site in SITE_NAMES:
        kwh = energies[site]
        if kwh <= 0:
            raise ValueError(f"No hay retiro de {site} en el Balance Valorizado.")
        participacion = kwh / e_total
        records.append({
            "punto": site, "clave": SSCC_KEYS[site], "retiro_kwh": kwh,
            "participacion_pct": participacion * 100,
            "ref_pd_clp": abs(totales["pd_neto_clp"]) * participacion,
            "ref_sobrecostos_energia_clp": abs(totales["sobrecostos_energia_netos_clp"]) * participacion,
        })
    return {
        "periodo": balance.label, "mes_factura": following_month(balance.label),
        "ret_enel_mwh": balance.retiro_enel_mwh, "global_enel": totales,
        "detalle": records,
        "total_referencia_pd": sum(r["ref_pd_clp"] for r in records),
        "total_referencia_energia": sum(r["ref_sobrecostos_energia_clp"] for r in records),
        "metodo": "Magnitud de Cuadro de Energía O/P (Enel) × participación del retiro mensual sobre Enel. Aproximación, NO reparto CEN oficial.",
        "estado": "REFERENCIAL / NO FACTURABLE",
    }


def _canon_cen_col(label: object) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", normalize(label)).strip("_")


def _tipo_mt_pd_cen(texto: object) -> str | None:
    value = _canon_cen_col(texto)
    if value in ("MT", "MINIMO_TECNICO", "MINIMOS_TECNICOS", "CARGO_MT"):
        return "MT"
    if value in ("PD", "P_D", "SOBRECOSTO_PD", "SOBRECOSTOS_PD", "PARTIDA_DETENCION", "PARTIDAS_Y_DETENCIONES", "PARTIDAS_DETENCIONES"):
        return "PD"
    if value in ("SOBRECOSTOS_ENERGIA", "SOBRECOSTO_ENERGIA"):
        return "SOBRECOSTOS_ENERGIA"  # Nunca asignar automáticamente a MT.
    return None


def parse_mt_pd_cen_retiros(data: bytes, filename: str) -> dict:
    """Lee tablas CEN por retiro solo si declaran concepto y monto explícitos.

    Tablas tipo PAGO_RETIRO: cabeceras 'clave' o 'clave_medidor', 'tipo_pago'
    o 'concepto' y 'pago' o 'monto', con 'suministrador' opcional.
    No interpreta a ciegas pagos generales ni sobregastos en otras columnas.
    """
    reader = WorkbookReader(data)
    resultado = {s: {"MT": None, "PD": None, "SOBRECOSTOS_ENERGIA": None} for s in SITE_NAMES}
    detalle = []
    posibles = []
    for sheet in reader.sheets:
        mapping = None
        encabezado = None
        for i,row in reader.sheet(sheet):
            if i > 15:
                break
            t = {_canon_cen_col(v): c for c, v in row.items()}
            def col(*opts):
                return next((t[o] for o in opts if o in t), None)
            a = col("CLAVE", "CLAVE_MEDIDOR", "CLAVE_RETIRO", "PUNTO_RETIRO")
            b = col("TIPO_PAGO", "CONCEPTO", "TIPO", "CARGO")
            c = col("PAGO", "MONTO", "VALOR", "MONTO_CLP", "PAGO_CLP")
            if a and b and c:
                mapping = {"clave":a,"concepto":b,"monto":c,
                           "empresa":col("SUMINISTRADOR", "EMPRESA", "NOMBRE_CORTO_EMPRESA")}
                encabezado=i
                break
        if not mapping:
            continue
        posibles.append(sheet)
        for i,row in reader.sheet(sheet):
            if i <= encabezado:
                continue
            key = normalize(row.get(mapping["clave"]))
            site = next((site for site, code in SSCC_KEYS.items() if key == code), None)
            if site is None:
                continue
            if mapping["empresa"] and normalize(row.get(mapping["empresa"])) != "ENEL_GENERACION":
                continue
            code = _tipo_mt_pd_cen(row.get(mapping["concepto"]))
            if not code:
                continue
            amount = num(row.get(mapping["monto"]))
            if resultado[site][code] is None:
                resultado[site][code] = 0.0
            resultado[site][code] += amount
            detalle.append({"punto":site,"clave":key,"concepto":code,"monto_clp":amount,
                            "archivo":filename,"hoja":sheet,"fila":i})
    if not posibles:
        raise ValueError("El archivo no posee cabeceras reconocidas de clave + tipo_pago/concepto + pago/monto; no se inventarán equivalencias.")
    if not detalle:
        raise ValueError("No se encontraron asignaciones MT/PD para MERVAL_ELSOL y MERVAL_LIMACHE. El archivo puede corresponder a otros pagos laterales.")
    return {"periodo":month_from_filename(filename),"archivo":filename,
            "resumen":resultado,"detalle":detalle,"hojas_revisadas":posibles}

# --------------------------------------------------
# Potencia de Suficiencia (PSUF): archivo CEN de balance definitivo de potencia.
# La hoja Balance2 incluye varias tablas paralelas; B:G es la única utilizada.
# El identificador comercial de Quilpué difiere del de Balance de Energía.
# El Pago PSUF aparece NEGATIVO cuando corresponde a retiros. Se conserva signo.

POTENCIA_IDENTIFICADORES = {
    "Quilpué": {"nombres": {"METRO REGIONAL DE VALPARAISO SA", "MERVAL QUILPUE", "MERVAL_ELSOL"},
               "barra": "QUILPUE", "nivel": "110"},
    "Limache": {"nombres": {"MERVAL LIMACHE", "MERVAL_LIMACHE"},
                "barra": "S.PEDRO", "nivel": "013"},
}


def parse_potencia_cen(data: bytes, filename: str) -> dict:
    libro = WorkbookReader(data)
    registro = {s: [] for s in POTENCIA_IDENTIFICADORES}
    fecha_balance = None
    fijacion = None
    cabeceras = {}
    for fila, c in libro.sheet("Balance2"):
        if fila == 6 and isinstance(c.get("C"), (int, float)):
            fecha_balance = excel_date(float(c["C"])).strftime("%Y-%m")
        if fila == 10:
            fijacion = str(c.get("E", ""))
        if fila == 11:
            cabeceras = c
        if fila < 12:
            continue
        nombre = normalize(c.get("B"))
        barra = normalize(c.get("C"))
        tipo = normalize(c.get("D"))
        for sitio, identificador in POTENCIA_IDENTIFICADORES.items():
            if (nombre not in identificador["nombres"]
                or not barra.startswith(identificador["barra"])
                or not barra.endswith(identificador["nivel"]) or tipo not in ("L", "L_D", "R")):
                continue
            mw_con_signo = num(c.get("E"))
            tarifa_kw_mes = num(c.get("F"))
            pago_firmado = num(c.get("G"))
            if tarifa_kw_mes <= 0:
                raise ValueError(f"Tarifa de potencia no válida en Balance2!F{fila}")
            calculado = mw_con_signo * 1000.0 * tarifa_kw_mes
            registro[sitio].append({
                "suministro": sitio,
                "descripcion_cen": str(c.get("B")),
                "barra_cen": str(c.get("C")),
                "tipo_cen": tipo,
                "potencia_mw_con_signo": mw_con_signo,
                "potencia_retiro_mw": max(0.0, -mw_con_signo),
                "precio_nudo_clp_kw_mes": tarifa_kw_mes,
                "pago_psuf_con_signo_clp": pago_firmado,
                "monto_retiro_psuf_clp": max(0.0, -pago_firmado),
                "recalculado_con_signo_clp": calculado,
                "diferencia_calculo_clp": pago_firmado - calculado,
                "fila": fila,
                "hoja": "Balance2",
            })
    esperadas = {"B": "DESCRIPCION", "C": "BARRA", "D": "TIPO CLAVE",
                 "E": "MEDIDA (MW)", "F": "PRECIO NUDO ($/KW/MES)", "G": "PAGO PSUF ($)"}
    for col, esperado in esperadas.items():
        if normalize(cabeceras.get(col)) != esperado:
            raise ValueError(f"Encabezado de Balance2!{col}11 desconocido: {cabeceras.get(col)!r}. "
                             "Revisar estructura antes de aplicar la fórmula.")
    avisos = []
    detalle = {}
    for sitio, registros in registro.items():
        if len(registros) != 1:
            avisos.append(f"{sitio}: se encontraron {len(registros)} filas de retiro identificadas "
                          "por nombre y barra. No se certifica total para ese suministro.")
        if len(registros) == 1:
            r = registros[0]
            if r["potencia_mw_con_signo"] > 0:
                avisos.append(f"{sitio}: la medida tiene signo positivo; revisar antes de interpretarla como retiro.")
            if r["pago_psuf_con_signo_clp"] > 0:
                avisos.append(f"{sitio}: PSUF positivo; revisar antes de interpretarlo como cargo.")
            if abs(r["diferencia_calculo_clp"]) > max(1.0, abs(r["pago_psuf_con_signo_clp"]) * 1e-6):
                avisos.append(f"{sitio}: el Pago PSUF publicado difiere de MW × 1000 × tarifa; revisar fila {r['fila']}.")
            detalle[sitio] = r
    periodo = fecha_balance or month_from_filename(filename)
    if not re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])", periodo):
        avisos.append("No se pudo establecer de forma confiable el mes del balance de potencia.")
    return {"archivo": filename, "periodo": periodo,
            "mes_factura": following_month(periodo),
            "fijacion": fijacion, "detalle": detalle, "avisos": avisos,
            "completo": len(detalle) == len(POTENCIA_IDENTIFICADORES),
            "total_monto_retiro_clp": sum(r["monto_retiro_psuf_clp"] for r in detalle.values()),
            "total_potencia_retiro_mw": sum(r["potencia_retiro_mw"] for r in detalle.values()),
            "fuente": "Balance2 (columnas B:G): Medida MW × 1.000 × Precio Nudo ($/kW/mes) = Pago PSUF ($)",
            "nota": "El archivo registra pago PSUF por retiro con signo negativo. "
                    "Se presenta su magnitud positiva para presupuestar; el cobro contractual "
                    "final del suministrador podría tener ajustes o condiciones propias."}


# Exenciones a Clientes Individualizados: liquidación por cliente y suministrador.
# No se confunde con Precio Estabilizado ni con Servicios Complementarios.
def parse_exenciones_ci(data: bytes, filename: str) -> dict:
    periodo = month_from_filename(filename)
    if not re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", periodo):
        raise ValueError("No se pudo determinar el mes (ejemplo de nombre: PagoExenCI2609.xlsm).")
    reader = WorkbookReader(data)
    hoja = "Pago x Exenciones a CI"
    filas = list(reader.sheet(hoja))
    fila_titulos = next((r for n, r in filas if n == 10), {})
    if normalize(fila_titulos.get("AW")) != "TOTAL GENERAL":
        raise ValueError("Estructura de Exenciones CI no reconocida: falta 'Total general' en AW10.")
    coincidencias = []
    for numero_fila, fila in filas:
        proveedor = normalize(fila.get("B"))
        cliente = normalize(fila.get("C"))
        cliente_efe = ("METRO REGIONAL DE VALPARAISO" in cliente
                       or cliente in ("EFE VALPARAISO S.A.", "EFE VALPARAISO SA"))
        if proveedor == "ENEL_GENERACION" and cliente_efe:
            detalle = []
            subtotal = 0.0
            for c, etiqueta in fila_titulos.items():
                if c in ("A", "B", "C", "AW") or not etiqueta:
                    continue
                monto = num(fila.get(c))
                if abs(monto) > 0.00001:
                    detalle.append({"empresa":str(etiqueta), "importe_clp":monto, "columna":c})
                subtotal += monto
            declarado = num(fila.get("AW"))
            if abs(subtotal-declarado) > max(1.0, abs(declarado)*0.00001):
                raise ValueError(f"El total AW{numero_fila} no coincide con la suma por transmisor.")
            coincidencias.append({"fila":numero_fila, "cliente":fila.get("C"),
                                   "suministrador":fila.get("B"), "total_clp":declarado,
                                   "detalle":detalle})
    if not coincidencias:
        raise ValueError("No existe un pago de ENEL_GENERACION para METRO REGIONAL DE VALPARAÍSO S.A. en este archivo.")
    if len(coincidencias) != 1:
        raise ValueError(f"Se detectaron {len(coincidencias)} coincidencias EFE; se requiere conciliarlas antes de sumar.")
    item = coincidencias[0]
    prorrata_individual = None
    for numero, fila in reader.sheet("Prorrata Individual"):
        if "METRO REGIONAL DE VALPARAISO" in normalize(fila.get("B")):
            prorrata_individual = num(fila.get("D"))
            break
    abreviaturas = ("ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
                    "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")
    abreviatura = abreviaturas[int(periodo[5:7])-1]
    prorrata_suministro = None
    for numero, fila in reader.sheet("Prorrata Suministro"):
        if (normalize(fila.get("B")) == abreviatura
                and normalize(fila.get("C")) == "ENEL_GENERACION"
                and "METRO REGIONAL DE VALPARAISO" in normalize(fila.get("D"))):
            prorrata_suministro = num(fila.get("E"))
            break
    return {"periodo":periodo, "archivo":filename, "hoja":hoja,
            "suministrador":item["suministrador"], "cliente":item["cliente"],
            "total_clp":item["total_clp"], "fila":item["fila"],
            "detalle":item["detalle"], "prorrata_individual":prorrata_individual,
            "prorrata_suministro":prorrata_suministro}


def asignar_exenciones_ci_suministros(reporte: dict) -> dict[str, float]:
    """Regla de negocio EFE: Exenciones CI se atribuye 100 % a Quilpué.

    El archivo CEN PagoExenCI liquida a nivel de cliente (Metro Regional de
    Valparaíso), por lo que esta desagregación proviene del criterio informado
    por EFE, no de dos pagos individuales del CEN. No se prorratea por consumo.
    """
    total = float(reporte["total_clp"])
    return {"Quilpué": total, "Limache": 0.0, "Total EFE": total}


import streamlit as st

# Carpeta activa única: Datos/<tipo>/<año>/, con los nombres indicados.
# No se leen las antiguas carpetas en minúsculas ni Excel sueltos.
# No se incluye datos/ en minúsculas ni archivos antiguos fuera de Datos/.
# Archivos >25 MiB: original.xlsx.part001, part002... en la MISMA carpeta.
CARPETAS_CEN = {
    "balance": "Balances",
    "precio": "Precios_estabilizados",
    "sscc4": "Sscc_archivo4",
    "sscc9": "Sscc_archivo9",
    "mtpd": "Mt_pd",
    "potencia": "Potencia",
    "exenciaci": "Exenciones_CI",
    "horario": "Horarios",
    "factura": "Facturas_enel",
    "proforma": "Proformas_enel",
}
# Se exige Datos/<categoría>/<año>/ para evitar fuentes heredadas.
CARPETAS_A_CATEGORIAS = {v.casefold(): k for k, v in CARPETAS_CEN.items()}


def categoria_carpeta_cen(nombre: str) -> str | None:
    """Tolera mayúsculas y separadores en carpetas creadas desde iPad/GitHub.

    Sólo acepta categorías CEN conocidas dentro de la carpeta principal Datos/.
    """
    exacta = CARPETAS_A_CATEGORIAS.get(nombre.casefold())
    if exacta:
        return exacta
    canonico = re.sub(r"[^A-Z0-9]", "", normalize(nombre))
    equivalencias = {re.sub(r"[^A-Z0-9]", "", normalize(v)): k
                     for k,v in CARPETAS_CEN.items()}
    equivalencias.update({"SSCC4": "sscc4", "SSCCARCHIVO4": "sscc4",
                          "ARCHIVO4SSCC": "sscc4", "SSCC9": "sscc9",
                          "SSCCARCHIVO9": "sscc9", "ARCHIVO9SSCC": "sscc9"})
    return equivalencias.get(canonico)

@dataclass(frozen=True)
class ArchivoPartesSesion:
    """Libro reconstruido exclusivamente a partir de partes subidas en Streamlit."""
    name: str
    contenido: bytes
    anio: int | None = None

    def getvalue(self) -> bytes:
        return self.contenido


def reconstruir_partes_subidas(subidas) -> tuple[dict[str, list[ArchivoPartesSesion]], list[dict]]:
    """Agrupa .part001, .part002... y valida el Excel antes de incorporarlo."""
    categorias = {k: [] for k in CARPETAS_CEN}
    estado = []
    patron = re.compile(r"^(.*\.(?:xlsx|xlsm))\.part(\d{3})$", re.I)
    grupos = {}
    for parte in (subidas or []):
        encontrado = patron.fullmatch(parte.name)
        if not encontrado:
            estado.append({"Archivo": parte.name, "Partes": "—",
                           "Estado": "Nombre no válido", "Detalle": "Usa Excel.xlsx.part001, .part002..."})
            continue
        base, numero = encontrado.group(1), int(encontrado.group(2))
        grupos.setdefault(base.casefold(), {"nombre": base, "partes": []})["partes"].append((numero, parte))
    for grupo in grupos.values():
        base = grupo["nombre"]
        partes = sorted(grupo["partes"], key=lambda x: x[0])
        indices = [numero for numero, _ in partes]
        if indices != list(range(1, len(indices)+1)):
            estado.append({"Archivo": base, "Partes": str(indices),
                           "Estado": "Faltan partes", "Detalle": "Selecciona TODAS las partes del mismo Excel."})
            continue
        try:
            contenido = b"".join(parte.getvalue() for _, parte in partes)
            if contenido.startswith(b"version https://git-lfs.github.com/spec/v1"):
                raise ValueError("Solo se cargó un puntero Git LFS, no el Excel")
            with ZipFile(BytesIO(contenido)) as z:
                requerido = {"[Content_Types].xml", "xl/workbook.xml"}
                faltantes = requerido.difference(z.namelist())
                if faltantes:
                    raise ValueError("Faltan componentes del Excel: " + ", ".join(sorted(faltantes)))
            categoria = clasificar_archivo_antiguo(base)
            periodo = month_from_filename(base)
            if categoria not in ("sscc9", "potencia", "balance", "precio", "sscc4", "mtpd", "exenciaci", "proforma"):
                raise ValueError("Nombre no reconocido como archivo CEN; revisa el nombre original")
            if not re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", periodo):
                raise ValueError("Falta período año-mes en el nombre del archivo")
            categorias[categoria].append(ArchivoPartesSesion(base, contenido, int(periodo[:4])))
            estado.append({"Archivo": base, "Partes": ", ".join(f"{n:03d}" for n in indices),
                           "Estado": "Cargado y reconstruido", "Detalle": f"{len(contenido)/1048576:.2f} MiB · {CARPETAS_CEN[categoria]} · {periodo}"})
        except Exception as exc:
            estado.append({"Archivo": base, "Partes": ", ".join(f"{n:03d}" for n in indices),
                           "Estado": "Error de reconstrucción", "Detalle": str(exc)})
    return categorias, estado


def directorios_datos_publicados() -> list[Path]:
    """Busca TODAS las carpetas Datos compatibles dentro del proyecto desplegado.

    Una carpeta Datos vacía junto a app.py no oculta otra carpeta Datos que
    sí contiene los libros del CEN. No lee carpetas de otros proyectos.
    """
    inicio = Path(__file__).resolve().parent
    candidatos = [inicio, *list(inicio.parents)[:3]]
    actual = Path.cwd().resolve()
    if actual == inicio or inicio in actual.parents or actual in inicio.parents:
        candidatos.extend([actual, *list(actual.parents)[:2]])
    encontrados = []
    for base in candidatos:
        if base == Path(base.anchor):
            continue
        for nombre in ("Datos", "datos"):
            carpeta = base / nombre
            if carpeta.is_dir() and carpeta.resolve() not in encontrados:
                encontrados.append(carpeta.resolve())
    return encontrados


def directorio_datos_publicados() -> Path:
    """Ruta principal para mostrar en el diagnóstico de la interfaz."""
    encontrados = directorios_datos_publicados()
    if encontrados:
        return max(encontrados, key=lambda p: sum(1 for f in p.rglob('*') if f.is_file()))
    return Path(__file__).resolve().parent / "Datos"


# Fuente alternativa: archivos realmente PUBLICADOS en el repositorio GitHub.
# La consulta remota NO agrega meses si GitHub no tiene esos archivos.
REPOSITORIO_GITHUB = os.environ.get("EFE_GITHUB_REPO", "antoniorra1983-ops/Precio_estabilizado")
RAMA_GITHUB = os.environ.get("EFE_GITHUB_BRANCH", "main")


@dataclass(frozen=True)
class ArchivoGithubRemoto:
    name: str
    rutas: tuple[str, ...]  # rutas relativas en el árbol de GitHub
    categoria: str
    anio: int | None
    carpeta: str
    raiz: str = "GitHub remoto"

    def getvalue(self) -> bytes:
        return _leer_partes_remotas(self.rutas)


@st.cache_data(ttl=600, show_spinner=False, max_entries=3)
def _arbol_publico_github(repo: str, rama: str) -> dict:
    url = f"https://api.github.com/repos/{repo}/git/trees/{quote(rama, safe='')}?recursive=1"
    peticion = Request(url, headers={"Accept":"application/vnd.github+json", "User-Agent":"EFE-CEN-Streamlit/2.16"})
    try:
        with urlopen(peticion, timeout=7) as respuesta:
            resultado = json.load(respuesta)
        if resultado.get("truncated"):
            raise ValueError("El listado de GitHub está truncado; usa los datos del despliegue local.")
        return resultado
    except Exception as exc:
        # También se cachea el error; evita repetir la misma espera en cada filtro.
        return {"error": str(exc)}


@st.cache_data(ttl=900, show_spinner=False, max_entries=20)
def _descargar_archivo_publico(repo: str, rama: str, ruta: str) -> bytes:
    nombre = quote(ruta, safe='/')
    url = f"https://raw.githubusercontent.com/{repo}/{quote(rama, safe='')}/{nombre}"
    peticion = Request(url, headers={"User-Agent":"EFE-CEN-Streamlit/2.16"})
    with urlopen(peticion, timeout=45) as respuesta:
        datos = respuesta.read()
    if datos.startswith(b"version https://git-lfs.github.com/spec/v1"):
        raise ValueError(f"{ruta} es un puntero Git LFS, no el archivo real")
    return datos


@st.cache_data(ttl=900, show_spinner=False, max_entries=15)
def _leer_partes_remotas(rutas: tuple[str, ...]) -> bytes:
    return b"".join(_descargar_archivo_publico(REPOSITORIO_GITHUB, RAMA_GITHUB, ruta) for ruta in rutas)


def descubrir_archivos_github_remoto() -> tuple[dict[str, list[ArchivoGithubRemoto]], list[str], dict]:
    """Consulta el árbol público cuando el despliegue no contiene todos los Excel.

    Sólo admite Datos/<categoría>/<AAAA>/archivo.xlsx[.partNNN].
    """
    grupos = {k: [] for k in CARPETAS_CEN}
    avisos = []
    estado = {"Repositorio":REPOSITORIO_GITHUB, "Rama":RAMA_GITHUB,
              "Estado":"Sin consultar", "Partes vistas":0}
    arbol = _arbol_publico_github(REPOSITORIO_GITHUB, RAMA_GITHUB)
    if arbol.get("error"):
        estado["Estado"] = "Consulta no disponible"
        avisos.append("No se pudo consultar el árbol público de GitHub (puede ser privado, "
                      f"existir un problema de red o estar en otra rama): {arbol['error']}")
        return grupos, avisos, estado
    estado["Estado"] = "Conectado"
    completos = {}
    divididos = {}
    rutas_fuera = []
    patron = re.compile(r"^(.*\.(?:xlsx|xlsm|csv))\.part(\d{3})$", re.I)
    for item in arbol.get("tree", []):
        if item.get("type") != "blob":
            continue
        ruta = item.get("path", "")
        secciones = ruta.split("/")
        if len(secciones) not in (4, 5) or secciones[0].casefold() != "datos":
            if re.search(r"\.part\d{3}$", ruta, re.I):
                rutas_fuera.append(ruta)
            continue
        cat = categoria_carpeta_cen(secciones[1])
        if cat is None or not re.fullmatch(r"20\d{2}", secciones[2]):
            continue
        # Carpeta oficial de proformas: Datos/Proformas_enel/AAAA/archivo.xlsx.
        # Se mantiene lectura de rutas históricas /AAAA/MM/ solo para compatibilidad.
        if len(secciones) == 5 and (cat != 'proforma' or
                 not re.fullmatch(r'0[1-9]|1[0-2]',secciones[3])):
            continue
        nombre = secciones[-1]
        mes = month_from_filename(nombre)
        if len(secciones)==5 and mes[5:7]!=secciones[3]:
            continue
        if not re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", mes):
            continue
        if int(mes[:4]) != int(secciones[2]):
            continue
        subcarpeta = '/'.join(secciones[1:-1])
        llave = (cat, int(secciones[2]), subcarpeta, nombre)
        m = patron.fullmatch(nombre)
        if m:
            clave = (cat, int(secciones[2]), subcarpeta, m.group(1))
            divididos.setdefault(clave, []).append((int(m.group(2)), ruta))
            estado["Partes vistas"] += 1
        elif nombre.lower().endswith((".xlsx", ".xlsm", ".csv")):
            completos[(cat, int(secciones[2]), subcarpeta, nombre)] = ruta
    for (cat, anio, carpeta, nombre), ruta in completos.items():
        grupos[cat].append(ArchivoGithubRemoto(nombre, (ruta,), cat, anio, carpeta))
    for clave, partes in divididos.items():
        cat, anio, carpeta, nombre = clave
        if clave in completos:
            continue
        partes.sort(key=lambda item:item[0])
        numeros = [n for n,_ in partes]
        if numeros != list(range(1,len(numeros)+1)):
            avisos.append(f"GitHub remoto: Datos/{carpeta}/{anio}/{nombre} tiene partes incompletas: {numeros}")
            continue
        rutas = tuple(ruta for _,ruta in partes)
        grupos[cat].append(ArchivoGithubRemoto(nombre,rutas,cat,anio,carpeta))
    estado["Libros disponibles"] = sum(map(len,grupos.values()))
    estado["Partes fuera de Datos/<tipo>/<año>/"] = rutas_fuera[:15]
    if rutas_fuera:
        avisos.append(f"GitHub remoto contiene {len(rutas_fuera)} archivo(s) .part fuera de "
                      "Datos/<tipo>/<año>/. Revisa las rutas mostradas en el diagnóstico.")
    return grupos, avisos, estado


@dataclass(frozen=True)
class ArchivoRepositorio:
    name: str
    rutas: tuple[str, ...]
    categoria: str = ""
    anio: int | None = None
    carpeta: str = ""
    raiz: str = "Datos"

    def getvalue(self) -> bytes:
        return _leer_archivos_repositorio(self.rutas)


@lru_cache(maxsize=36)
def _leer_archivos_repositorio(rutas: tuple[str, ...]) -> bytes:
    return b"".join(Path(ruta).read_bytes() for ruta in rutas)


def comprobar_integridad_excel(rutas: tuple[str, ...], nombre: str) -> tuple[bool, str]:
    """Verifica que todas las partes formen un Excel ZIP completo, no truncado.

    El ZIP se verifica por índice central, formato mínimo y referencias internas.
    No se descomprimen millones de filas durante el inventario.
    """
    if not rutas:
        return False, "No existen partes para reconstruir el Excel."
    try:
        data = _leer_archivos_repositorio(rutas)
        if not data:
            return False, "El archivo está vacío."
        if data.startswith(b"version https://git-lfs.github.com/spec/v1"):
            return False, "Se encontró un puntero Git LFS, no el contenido Excel."
        if not data.startswith(b"PK\x03\x04"):
            return False, ("El contenido no comienza como XLSX/XLSM; puede faltar la "
                           "primera parte o se subió un archivo diferente.")
        with ZipFile(BytesIO(data)) as z:
            contenido = set(z.namelist())
            necesarios = {"[Content_Types].xml", "xl/workbook.xml"}
            faltan = sorted(necesarios - contenido)
            if faltan:
                return False, "El Excel reconstruido no contiene " + ", ".join(faltan)
    except Exception as e:
        return False, ("No se pudo reconstruir el Excel: " + str(e) +
                       ". Comprueba que subiste TODAS las partes y sin cambiar nombres.")
    return True, f"Excel válido ({len(rutas)} parte(s), {len(data)/1048576:.2f} MiB)."


def diagnosticar_partes(raiz: Path) -> list[dict]:
    """Diagnóstico con rutas y motivos, accesible desde la pestaña Datos."""
    if not raiz.is_dir():
        return [{"Ruta": str(raiz), "Archivo": "—", "Partes": "—",
                 "Estado": "No existe la carpeta Datos", "Detalle": "Comprueba rama y ruta de app.py"}]
    patron = re.compile(r"^(.*\.(?:xlsx|xlsm))\.part(\d{3})$", re.I)
    grupos = {}
    errores = []
    for ruta in raiz.rglob('*'):
        if not ruta.is_file():
            continue
        if ruta.suffix.casefold() == '.zip':
            errores.append({"Ruta": str(ruta.parent.relative_to(raiz)), "Archivo": ruta.name,
                            "Partes": "—", "Estado": "ZIP sin descomprimir",
                            "Detalle": "Extrae el ZIP y sube .part001, .part002, etc."})
        m = patron.fullmatch(ruta.name)
        if m:
            grupos.setdefault((ruta.parent, m.group(1)), []).append((int(m.group(2)), ruta))
        elif '.part' in ruta.name.lower():
            errores.append({"Ruta": str(ruta.parent.relative_to(raiz)), "Archivo": ruta.name,
                            "Partes": "—", "Estado": "Nombre inválido",
                            "Detalle": "Usa exactamente archivo.xlsx.part001, .part002…"})
    for (carpeta, nombre), datos in sorted(grupos.items(), key=lambda kv: str(kv[0])):
        datos.sort()
        indices = [v for v, _ in datos]
        detalle = ''
        estado = 'Correcto'
        if indices != list(range(1,len(indices)+1)):
            estado = 'Partes faltantes o repetidas'
            detalle = 'Secuencia encontrada: ' + ', '.join(f'{n:03d}' for n in indices)
        else:
            valide, detalle = comprobar_integridad_excel(tuple(str(p) for _,p in datos),nombre)
            if not valide:
                estado = 'No se puede reconstruir'
        rel=carpeta.relative_to(raiz)
        partes=f'{len(datos)}: ' + ', '.join(f'{n:03d}' for n in indices)
        if len(rel.parts)!=2 or not re.fullmatch(r'20\d{2}', rel.parts[1]):
            estado = 'Carpeta incorrecta'
            detalle += ' · Debe ubicarse en Datos/<tipo>/AAAA/'
        elif categoria_carpeta_cen(rel.parts[0]) is None:
            estado = 'Categoría incorrecta'
            detalle += ' · Usa un nombre de carpeta reconocido.'
        errores.append({"Ruta": str(rel), "Archivo": nombre,
                        "Partes": partes, "Estado": estado, "Detalle": detalle})
    return errores


def clasificar_archivo_antiguo(nombre: str) -> str | None:
    n = normalize(nombre)
    if n.startswith("BALANCE_"):
        return "balance"
    if "PRECIOS ESTABILIZADOS" in n:
        return "precio"
    if n.startswith("9_PAGOS_RETIROS"):
        return "sscc9"
    if n.startswith("4_") and ("REMUNERACION" in n or "PAGOS_RETIROS" in n):
        return "sscc4"
    if n.startswith("BDEF") and "DETALLE" in n:
        return "potencia"
    if "PAGOEXENCI" in re.sub(r"[^A-Z0-9]", "", n) or "EXENCIONES" in n and "CI" in n:
        return "exenciaci"
    if n.startswith("EFE_VALPO_"):
        return "proforma"
    if "MT_PD" in n or "PAGOS_LATERALES" in n:
        return "mtpd"
    if n.endswith(".CSV") and "HORARIO" in n:
        return "horario"
    return None


def extraer_anio_nombre(nombre: str) -> int | None:
    per = month_from_filename(nombre)
    if re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", per):
        return int(per[:4])
    patron = re.search(r"(?:^|[^0-9])(20\d{2})(?:[^0-9]|$)", nombre)
    return int(patron.group(1)) if patron else None


def _descubrir_archivos_en_raiz(raiz: Path) -> tuple[dict[str, list[ArchivoRepositorio]], list[str]]:
    categorias: dict[str, list[ArchivoRepositorio]] = {c: [] for c in CARPETAS_CEN}
    avisos: list[str] = []
    if not raiz.is_dir():
        return categorias, avisos

    patron_partes = re.compile(r"^(.*\.(?:xlsx|xlsm|csv))\.part(\d{3})$", re.I)
    # Agrupar SIEMPRE por carpeta relativa, evitando unir partes entre años.
    agrupadas: dict[Path, dict[str, list[tuple[int,Path]]]] = {}
    completos: dict[Path, dict[str, Path]] = {}
    for ruta in raiz.rglob("*"):
        if not ruta.is_file() or ruta.name.startswith("."):
            continue
        rel_dir = ruta.parent.relative_to(raiz)
        trozos = patron_partes.fullmatch(ruta.name)
        if trozos:
            agrupadas.setdefault(rel_dir, {}).setdefault(trozos.group(1), []).append((int(trozos.group(2)), ruta))
        elif ruta.suffix.casefold() in (".xlsx", ".xlsm", ".csv"):
            completos.setdefault(rel_dir, {})[ruta.name] = ruta
        else:
            avisos.append(f"Se omitió {ruta.relative_to(raiz)}: formato no compatible.")

    registros: list[tuple[Path,str,tuple[str,...]]] = []
    for carpeta in sorted(set(agrupadas)|set(completos),key=str):
        enteros = completos.get(carpeta, {})
        for nombre, ruta in enteros.items():
            registros.append((carpeta, nombre, (str(ruta),)))
        for nombre, pedazos in agrupadas.get(carpeta, {}).items():
            if nombre in enteros:
                avisos.append(f"{carpeta}/{nombre}: existe el original; no se utilizan sus partes.")
                continue
            pedazos.sort(key=lambda x: x[0])
            ids = [n for n,_ in pedazos]
            if ids != list(range(1,len(ids)+1)):
                avisos.append(f"{carpeta}/{nombre}: partes incompletas {ids}. Deben ser 001, 002, etc. sin saltos.")
                continue
            rutas = tuple(str(p) for _,p in pedazos)
            valido, motivo = comprobar_integridad_excel(rutas, nombre)
            if not valido:
                avisos.append(f"{carpeta}/{nombre}: {motivo} Archivo omitido; vuelve a subir las partes.")
                continue
            registros.append((carpeta,nombre,rutas))

    candidatos: dict[tuple[str,int,str],ArchivoRepositorio] = {}
    prioridades: dict[tuple[str,int,str],int] = {}
    for carpeta, nombre, rutas in registros:
        niveles = carpeta.parts
        if not niveles:
            avisos.append(f"{nombre}: archivo sin categoría/año, ignorado. Muévelo a Datos/<tipo>/AAAA/.")
            continue
        else:
            cat = categoria_carpeta_cen(niveles[0])
            if cat is None:
                avisos.append(f"{carpeta}/{nombre}: categoría no reconocida. "
                              f"Para el archivo 4 usa Datos/Sscc_archivo4/AAAA/nombre.xlsx.")
                continue
            if niveles[0] != CARPETAS_CEN[cat]:
                avisos.append(f"{carpeta}/{nombre}: carpeta reconocida como {CARPETAS_CEN[cat]}, "
                              f"pero se recomienda renombrarla a {CARPETAS_CEN[cat]}.")
            detectado = extraer_anio_nombre(nombre)
            # Compatibilidad con carpetas antiguas por año/mes; nuevas proformas
            # deben subirse solo a Proformas_enel/AAAA/nombre.xlsx.
            if (cat == 'proforma' and len(niveles) == 3
                    and re.fullmatch(r'20\d{2}',niveles[1])
                    and re.fullmatch(r'0[1-9]|1[0-2]',niveles[2])):
                prioridad = 4
                anno = int(niveles[1])
                periodo_nombre = month_from_filename(nombre)
                if not re.fullmatch(r'20\d{2}-(?:0[1-9]|1[0-2])',periodo_nombre):
                    avisos.append(f'{carpeta}/{nombre}: falta mes reconocible en el nombre; se omitió.')
                    continue
                if periodo_nombre != f'{anno:04d}-{niveles[2]}':
                    avisos.append(f'{carpeta}/{nombre}: carpeta /AAAA/MM no coincide con el archivo ({periodo_nombre}); se omitió.')
                    continue
            elif len(niveles) == 2 and re.fullmatch(r"20\d{2}", niveles[1]):
                prioridad = 3  # Prefiere archivo organizado en carpeta del año.
                anno = int(niveles[1])
                if detectado and anno != detectado:
                    avisos.append(f"{carpeta}/{nombre}: el año {anno} de la carpeta no coincide con "
                                  f"{detectado} del archivo; se omitió.")
                    continue
            elif len(niveles) == 1 and detectado:
                prioridad = 1  # Compatibilidad; no debe desplazar a la ruta anual.
                anno = detectado
                avisos.append(f"{carpeta}/{nombre}: archivo reconocido gracias al período del nombre, "
                              f"pero muévelo a Datos/{CARPETAS_CEN[cat]}/{anno}/ para ordenar los años.")
            else:
                avisos.append(f"{carpeta}/{nombre}: coloca el archivo dentro de "
                              f"Datos/{CARPETAS_CEN[cat]}/AAAA/. El nombre también debe indicar el mes (ej. 2609).")
                continue
        objeto = ArchivoRepositorio(nombre, rutas, cat, anno, str(carpeta), raiz.name)
        # El mismo nombre de archivo puede existir en años diferentes.
        identificador = (cat, anno or 0, nombre.casefold())
        if identificador in candidatos:
            if prioridad <= prioridades[identificador]:
                avisos.append(f"Duplicado {nombre} en {carpeta}; se conserva {candidatos[identificador].carpeta}.")
                continue
            avisos.append(f"Duplicado {nombre}: se prefiere la carpeta actual {carpeta}.")
        candidatos[identificador] = objeto
        prioridades[identificador] = prioridad

    for ar in candidatos.values():
        categorias[ar.categoria].append(ar)
    for ar_list in categorias.values():
        ar_list.sort(key=lambda ar: (ar.anio or 0,month_from_filename(ar.name),ar.name.casefold()))
    return categorias,avisos


def descubrir_archivos_repositorio() -> tuple[dict[str, list[ArchivoRepositorio]], list[str]]:
    """Lee SOLO Datos/ del despliegue actual.

    No incluye archivos de ejemplo, de este chat, de la carpeta antigua
    datos/ ni libros sueltos fuera de las subcarpetas configuradas.
    """
    encontrados = {cat: [] for cat in CARPETAS_CEN}
    avisos = []
    carpetas = directorios_datos_publicados()
    vistos = set()
    for raiz in carpetas:
        hallados, advertencias = _descubrir_archivos_en_raiz(raiz)
        for cat, archivos in hallados.items():
            for archivo in archivos:
                llave = (cat, archivo.anio, archivo.name.casefold())
                if llave not in vistos:
                    encontrados[cat].append(archivo)
                    vistos.add(llave)
        avisos.extend([f"{raiz}: {msg}" for msg in advertencias])
    return encontrados, avisos


def fusionar_archivos(locales, subidos):
    """Solo combina archivos subidos en esta sesión y del repositorio desplegado: año+nombre.

    Los archivos subidos a la aplicación prevalecen sobre una copia con el
    mismo nombre que esté guardada en el repositorio. No inyecta datos de prueba.
    """
    result = {(getattr(a, 'anio', None) or extraer_anio_nombre(a.name), a.name.casefold()): a
              for a in (locales or [])}
    for a in subidos or []:
        key = (getattr(a, 'anio', None) or extraer_anio_nombre(a.name), a.name.casefold())
        result[key] = a
    return sorted(result.values(), key=lambda a: (month_from_filename(a.name), a.name.casefold()))


def opciones_periodo_archivos(archivos):
    return sorted(archivos,key=lambda a: (month_from_filename(a.name), a.name.casefold()))

# Los resultados nacen exclusivamente de los archivos que el usuario agrega a
# la aplicación, y opcionalmente los que haya guardado en Datos/ en su repositorio.
# No se incluye ningún Excel histórico en el paquete instalable.

# v2.16: lectura de partes del repositorio, incluso cuando hay carpetas intermedias.
# Se limita a rutas dentro de una carpeta Datos real del repositorio desplegado.
def descubrir_partes_github_flexible() -> tuple[dict[str, list[ArchivoRepositorio]], list[dict]]:
    """Reconoce y valida .part001 ... .partNNN bajo Datos de GitHub.

    No usa archivos compartidos por ChatGPT ni datos de ejemplo; solo los presentes
    en el disco del despliegue. Evita depender de un nivel exacto de subcarpetas.
    """
    resultados = {k: [] for k in CARPETAS_CEN}
    diagnostico = []
    base_app = Path(__file__).resolve().parent
    ubicaciones = [base_app, base_app.parent, Path.cwd().resolve()]
    carpetas_datos = directorios_datos_publicados()
    # Algunos proyectos sitúan app.py dentro de una subcarpeta: explorar la raíz
    # visible, sin recorrer carpetas de otros proyectos ni /.
    for ubicacion in ubicaciones:
        if not ubicacion.is_dir():
            continue
        for d in ubicacion.glob('*/Datos'):
            if d.is_dir() and d.resolve() not in carpetas_datos:
                carpetas_datos.append(d.resolve())
    vistas = set()
    patron = re.compile(r"^(.*\.(?:xlsx|xlsm))\.part(\d{3})$", re.I)
    for raiz in carpetas_datos:
        grupos = {}
        for f in raiz.rglob('*'):
            if not f.is_file():
                continue
            m = patron.fullmatch(f.name)
            if not m:
                continue
            nombre = m.group(1)
            # Las partes deben estar juntas. Evitar mezclas de distintas carpetas/años.
            clave = (str(f.parent.resolve()), nombre.casefold())
            g = grupos.setdefault(clave, {'nombre':nombre, 'partes':[], 'raiz':raiz})
            g['partes'].append((int(m.group(2)), f))
        for grupo in grupos.values():
            nombre = grupo['nombre']
            partes = sorted(grupo['partes'], key=lambda z:z[0])
            numeros = [n for n,_ in partes]
            carpeta = partes[0][1].parent
            relativa = carpeta.relative_to(raiz)
            segmentos = relativa.parts
            categoria = next((categoria_carpeta_cen(p) for p in segmentos
                               if categoria_carpeta_cen(p)), None)
            categoria = categoria or clasificar_archivo_antiguo(nombre)
            periodo = month_from_filename(nombre)
            anno = int(periodo[:4]) if re.fullmatch(r'20\d{2}-(?:0[1-9]|1[0-2])',periodo) else None
            fila={'Carpeta encontrada': str(carpeta), 'Archivo': nombre,
                  'Partes': ', '.join(f'{n:03d}' for n in numeros),
                  'Categoría':CARPETAS_CEN.get(categoria,'Sin reconocer'),
                  'Mes':periodo, 'Estado':'Pendiente', 'Detalle':''}
            if not categoria:
                fila.update(Estado='Omitido', Detalle='No se pudo identificar el tipo de archivo CEN')
            elif not anno:
                fila.update(Estado='Omitido', Detalle='No se identificó mes y año en el nombre original')
            elif (categoria == 'proforma' and len(segmentos) >= 3 and
                  re.fullmatch(r'0[1-9]|1[0-2]', segmentos[-1]) and
                  segmentos[-1] != periodo[5:7]):
                fila.update(Estado='Omitido', Detalle='Mes de carpeta diferente al mes del Excel')
            elif any(re.fullmatch(r'20\d{2}',p) and int(p)!=anno for p in segmentos):
                fila.update(Estado='Omitido', Detalle='El año de la carpeta no coincide con el nombre del Excel')
            elif numeros != list(range(1,len(numeros)+1)):
                fila.update(Estado='Omitido', Detalle='Partes incompletas o duplicadas: deben comenzar por 001 y ser consecutivas')
            else:
                rutas = tuple(str(f) for _, f in partes)
                correcto, detalle = comprobar_integridad_excel(rutas, nombre)
                if not correcto:
                    fila.update(Estado='Error en reconstrucción', Detalle=detalle)
                else:
                    identificador=(categoria,anno,nombre.casefold())
                    if identificador in vistas:
                        fila.update(Estado='Duplicado', Detalle='Ya se incorporó otra copia del mismo Excel')
                    else:
                        vistas.add(identificador)
                        resultados[categoria].append(ArchivoRepositorio(
                            nombre,rutas,categoria,anno,str(relativa),str(raiz)))
                        fila.update(Estado='Reconstruido y disponible',
                                    Detalle=f'{len(partes)} partes encontradas; Excel válido')
            diagnostico.append(fila)
    return resultados, diagnostico


archivos_repositorio, advertencias_repositorio = descubrir_archivos_repositorio()
# Recuperación flexible: integra las partes que estaban en GitHub pero eran
# ignoradas por una estructura de carpetas distinta a la ruta canónica.
archivos_flexibles, diagnostico_flexible = descubrir_partes_github_flexible()
for _cat, _archivos in archivos_flexibles.items():
    _existentes = {(x.name.casefold(),x.anio) for x in archivos_repositorio[_cat]}
    archivos_repositorio[_cat].extend(
        x for x in _archivos if (x.name.casefold(),x.anio) not in _existentes)

# Si SSCC 9 o Potencia no aparecen en el despliegue, revisar también GitHub
# remoto. Los archivos siguen siendo exclusivamente los publicados por el usuario.
archivos_remotos, avisos_remotos, estado_github = ({k: [] for k in CARPETAS_CEN}, [], {"Estado":"No fue necesario"})
if (not archivos_repositorio.get("sscc9") or not archivos_repositorio.get("potencia")
        or not archivos_repositorio.get("exenciaci")
        or not archivos_repositorio.get("proforma")):
    archivos_remotos, avisos_remotos, estado_github = descubrir_archivos_github_remoto()
# La copia local prevalece sobre una copia idéntica encontrada en GitHub remoto.
for _cat, _items in archivos_remotos.items():
    _locales = {(x.name.casefold(), x.anio) for x in archivos_repositorio[_cat]}
    archivos_repositorio[_cat].extend(x for x in _items if (x.name.casefold(), x.anio) not in _locales)
advertencias_repositorio.extend(avisos_remotos)

st.set_page_config(page_title="EFE Valparaíso | Panel energético", page_icon="⚡", layout="wide")
st.markdown("""<style>
.block-container{max-width:1420px;padding-top:1.0rem;padding-bottom:2rem}
h1{letter-spacing:-.035em}
.stMetric{border:1px solid #dce4ee;border-radius:12px;padding:12px}
[data-testid="stTabs"] button{font-size:.92rem}
</style>""", unsafe_allow_html=True)
st.title("⚡ EFE Valparaíso | Gestión energética")
st.caption("Quilpué · Limache  |  Consumos, costos y cargos oficiales  |  Panel ejecutivo v1.0")

# Seis secciones principales; las pantallas técnicas originales se conservan
# en subpestañas. No se altera el motor de cálculo ni el flujo de lectura CEN.
resumen_principal, energia_principal, cargos_principal, proforma_principal, tarifas_principal, dolar_principal, configuracion_principal = st.tabs([
    "🏠 Resumen", "⚡ Energía", "💰 Costos y cargos", "🧾 Proforma", "📚 Tarifas CNE",
    "💵 Dólar observado", "⚙️ Datos y ajustes"
])

with resumen_principal:
    # Se llena al final del archivo para que las cargas CEN y las proyecciones
    # estén disponibles, sin importar cuál pestaña tenga abierta el usuario.
    espacio_resumen = st.empty()

with proforma_principal:
    espacio_proforma = st.empty()

with energia_principal:
    proyeccion_operativa_tab, energia_tab = st.tabs(["🔮 Proyección de energía", "📊 Consumo CEN"])

with cargos_principal:
    costos_activos, pe_grupo, cen_grupo, historico_grupo = st.tabs([
        "💡 Costo de energía", "📊 Precio Estabilizado", "🏛️ Cargos CEN", "📈 Histórico PE"
    ])
    with costos_activos:
        costo_energia_tab = st.container()
    with pe_grupo:
        calc_tab, hourly_tab = st.tabs(["Compensación PE", "Detalle horario"])
    with cen_grupo:
        sscc_tab, potencia_tab, exenciaci_tab, mt_pd_tab = st.tabs([
            "SSCC", "Potencia", "Exenciones CI", "Mínimos técnicos y PD"
        ])
    with historico_grupo:
        history_tab = st.container()

with tarifas_principal:
    tarifas_cne_tab = st.container()

with dolar_principal:
    dolar_observado_tab = st.container()

with configuracion_principal:
    datos_tab, monitoreo_cen_tab, method_tab = st.tabs([
        "📁 Archivos y cargas", "📡 Actualizaciones oficiales", "📖 Metodología"
    ])

with datos_tab:
    st.subheader("1 · Cargar información del CEN")
    st.write("**Sube los archivos aquí una sola vez**. Los módulos trabajan con esta misma selección; "
             "no se usan los Excel enviados anteriormente al chat ni resultados numéricos precargados.")
    # El contenido de Datos/ del repositorio se incluye SIEMPRE.
    # En la v2.5 el checkbox estaba desactivado por defecto, por lo que los
    # libros correctamente guardados en GitHub no aparecían en el inventario.
    usar_repositorio = True
    total_repo = sum(len(items) for items in archivos_repositorio.values())
    if total_repo:
        st.success(f"Archivos detectados desde GitHub: {total_repo} libro(s) CEN (local o remoto). "
                   "Ya están disponibles para todas las pestañas y sus filtros.")
    else:
        st.warning("No encontré archivos en Datos/ dentro del despliegue de Streamlit. "
                   "Comprueba que los Excel estén confirmados (commit) en la misma rama que app.py, "
                   "dentro de Datos/Balances/AAAA, Datos/Sscc_archivo4/AAAA, "
                   "Datos/Sscc_archivo9/AAAA o Datos/Potencia/AAAA. "
                   "Si acabas de subirlos, reinicia o redespliega la aplicación.")
    st.caption("Solo se leen libros ubicados en Datos/<tipo>/<año> del despliegue de GitHub, "
               "además de los subidos directamente en esta sesión. "
               "Los archivos guardados en GitHub se cargan automáticamente al abrir la aplicación; "
               "puedes añadir o reemplazar archivos temporalmente usando los controles de esta pestaña. "
               "Las cargas manuales solo permanecen en la sesión y no se publican en GitHub.")

    col_a, col_b = st.columns(2)
    with col_a:
        balance_uploads = st.file_uploader("Balances de Energía (Balance_...)  ·  .xlsm/.xlsx",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_balance")
        price_uploads = st.file_uploader("Precios estabilizados  ·  .xlsm/.xlsx",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_precio")
        sscc_empresa_uploads = st.file_uploader("SSCC · Archivo 4 (pagos globales de Enel)",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_sscc4")
        sscc_uploads = st.file_uploader("SSCC · Archivo 9 (asignación por retiro, incluye BESS)",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_sscc9")
    with col_b:
        potencia_uploads = st.file_uploader("Potencia · BDef Detalle CEN",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_potencia")
        exenciaci_uploads = st.file_uploader("Exenciones a Clientes Individualizados · PagoExenCI",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_exenciaci")
        mtpd_cen_uploads = st.file_uploader("MT y PD · Asignaciones por retiro (si existen)",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_mtpd")
        horarios_uploads = st.file_uploader("Compensaciones horarias · CSV",
            type=["csv"], accept_multiple_files=True, key="datos_horarios")
        invoice_uploads = st.file_uploader("Factura Enel (solo conciliación, opcional)",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_factura")
        proforma_uploads = st.file_uploader("🧾 Proformas Enel Quilpué y Limache (con CSP)",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_proforma",
            help="Para dejarlo permanente: Datos/Proformas_enel/2026/EFE_VALPO_SEP2026.xlsx en GitHub. Todos los meses juntos en la carpeta del año.")
    st.markdown("##### Cargar archivos divididos (.part001, .part002…)")
    st.caption("Si GitHub aún no los muestra, aquí puedes probar las partes directamente. "
               "Selecciona todas las partes del mismo Excel de una vez. "
               "La carga se mantiene solo durante esta sesión; no modifica GitHub.")
    partes_uploads = st.file_uploader(
        "Partes binarias de Excel CEN", type=None, accept_multiple_files=True,
        key="datos_partes", help="Ejemplo: 9_Pagos_Retiros_2609_def.xlsx.part001")
    with st.expander("Configuración de identificadores de EFE", expanded=False):
        clave_text = st.text_area("Claves de retiro en Balance Valorizado (una por línea)",
            "MERVAL_ELSOL\nMERVAL_LIMACHE", height=88, key="datos_claves")
    keys = [k.strip() for k in clave_text.replace(",", "\n").splitlines() if k.strip()]
    st.info("**Seguridad:** los archivos se procesan en el servidor donde está alojado Streamlit. "
            "No se ejecutan macros. Evita guardar datos comerciales en repositorios públicos.")
    if st.button("🧹 Limpiar archivos subidos en esta sesión", key="limpiar_cargas_v27"):
        # Los ficheros del repositorio nunca se modifican desde este botón.
        for widget_key in ("datos_balance", "datos_precio", "datos_sscc4", "datos_sscc9",
                           "datos_potencia", "datos_exenciaci", "datos_mtpd", "datos_horarios", "datos_factura", "datos_proforma", "datos_partes"):
            st.session_state.pop(widget_key, None)
        st.rerun()

archivos_partes_sesion, estado_partes_sesion = reconstruir_partes_subidas(partes_uploads)
with datos_tab:
    if estado_partes_sesion:
        st.markdown("##### Validación de partes subidas a esta sesión")
        st.dataframe(estado_partes_sesion, hide_index=True, use_container_width=True)
        if any(item["Estado"].startswith("Error") or item["Estado"] in ("Faltan partes", "Nombre no válido")
               for item in estado_partes_sesion):
            st.error("Hay archivos divididos que no se reconstruyeron. Revisa la tabla.")
        else:
            st.success("Los archivos divididos de esta sesión están listos para calcular.")

subidas = {"balance":balance_uploads or [], "precio":price_uploads or [],
           "sscc4":sscc_empresa_uploads or [], "sscc9":sscc_uploads or [],
           "potencia":potencia_uploads or [], "exenciaci":exenciaci_uploads or [], "mtpd":mtpd_cen_uploads or [],
           "horario":horarios_uploads or [], "factura":invoice_uploads or [],
           "proforma":proforma_uploads or []}
# Precedencia: GitHub -> partes de la sesión -> archivos completos de la sesión.
# La carga manual más reciente prevalece si el nombre coincide.
archivos_activos = {}
for categoria in CARPETAS_CEN:
    fuentes = archivos_repositorio.get(categoria, []) if usar_repositorio else []
    fuentes = fusionar_archivos(fuentes, archivos_partes_sesion.get(categoria, []))
    archivos_activos[categoria] = fusionar_archivos(fuentes, subidas[categoria])

with datos_tab:
    st.divider()
    st.markdown("#### Archivos que utilizará la aplicación")
    inventario = []
    for categoria, archivos in archivos_activos.items():
        for archivo in archivos:
            periodo = month_from_filename(archivo.name)
            inventario.append({"Categoría":CARPETAS_CEN[categoria],
                               "Archivo":archivo.name,
                               "Mes del balance":readable_month(periodo) if periodo.startswith("20") else "Sin identificar",
                               "Procedencia":("GitHub remoto" if isinstance(archivo, ArchivoGithubRemoto)
                                             else "Datos/ del repositorio") if isinstance(archivo, (ArchivoRepositorio, ArchivoGithubRemoto))
                                             else ("Partes cargadas en la sesión" if isinstance(archivo, ArchivoPartesSesion)
                                                   else "Subido en la aplicación")})
    if inventario:
        st.dataframe(inventario, hide_index=True, use_container_width=True)
        st.caption(f"{len(inventario)} archivo(s) disponibles. Cada archivo se registra una sola vez por categoría, "
                   "año y nombre. Las distintas versiones de un balance se concilian por período.")
    else:
        st.warning("Todavía no hay archivos disponibles. Sube los Excel del CEN arriba "
                   "para que aparezcan los valores en las demás pestañas.")
    if usar_repositorio:
        for aviso in advertencias_repositorio:
            st.warning(aviso)
    with st.expander("🔎 Diagnóstico de partes subidas a GitHub", expanded=True):
        st.caption("Busca las carpetas Datos del despliegue y, si faltan libros, intenta una consulta "
                   "al repositorio público remoto. No utiliza archivos enviados únicamente por ChatGPT.")
        ruta_datos_real = directorio_datos_publicados()
        st.code(str(ruta_datos_real), language=None)
        if ruta_datos_real.is_dir():
            carpetas = sorted(p.name for p in ruta_datos_real.iterdir())
            st.caption("Carpetas y archivos encontrados en Datos/: " + (
                ", ".join(carpetas[:30]) if carpetas else "carpeta vacía"))
        else:
            st.error("La carpeta Datos/ no existe en este despliegue. "
                     "Comprueba la rama de Streamlit Cloud, la ubicación de app.py y que los cambios estén publicados.")
        st.markdown("**Carpetas Datos realmente encontradas en el despliegue**")
        carpetas_detectadas = directorios_datos_publicados()
        if carpetas_detectadas:
            st.dataframe([{"Carpeta":str(ruta),
                           "Archivos físicos":sum(1 for f in ruta.rglob('*') if f.is_file())}
                          for ruta in carpetas_detectadas],
                         hide_index=True, use_container_width=True)
        else:
            st.warning("No existe ninguna carpeta Datos en el despliegue local.")
        st.markdown("**Comprobación independiente de las partes guardadas en GitHub**")
        if diagnostico_flexible:
            st.dataframe(diagnostico_flexible,hide_index=True,use_container_width=True)
            st.caption("En esta tabla se ven las rutas completas de las partes físicamente presentes en el despliegue. "
                       "Si no aparece ninguna, Streamlit no dispone de esas partes en el código desplegado. "
                       "Comprueba el commit, la rama y que no estén solo en un ZIP.")
        else:
            st.error("El despliegue actual no tiene archivos .part001 bajo ninguna carpeta Datos encontrada. "
                     "Verifica que GitHub tenga ambas partes y que estén confirmadas en la misma rama desplegada.")
        st.markdown("**Consulta al repositorio remoto**")
        st.json(estado_github, expanded=False)
        st.caption("Si GitHub remoto detecta partes y el despliegue no, comprueba la rama publicada "
                   "en Streamlit Cloud. Si GitHub no las encuentra, revisa las rutas y el commit.")
        tabla_partes = []
        for _ruta in carpetas_detectadas:
            for _fila in diagnosticar_partes(_ruta):
                tabla_partes.append({"Raíz Datos":str(_ruta), **_fila})
        if tabla_partes:
            st.dataframe(tabla_partes, hide_index=True, use_container_width=True)
            if all(fila["Estado"] == "Correcto" for fila in tabla_partes):
                st.success("Todos los grupos .part encontrados se pueden reconstruir.")
            else:
                st.warning("Hay partes que requieren corrección. Revisa Estado y Detalle.")
        else:
            st.info("No se encontraron archivos .part en Datos/ del despliegue actual. "
                    "Comprueba que hayas hecho Commit y que Streamlit use la misma rama.")

# Un único filtro global compartido por todos los módulos.
anios_disponibles = sorted({int(month_from_filename(a.name)[:4])
    for archivos in archivos_activos.values() for a in archivos
    if re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", month_from_filename(a.name))})
meses_disponibles = sorted({int(month_from_filename(a.name)[5:7])
    for archivos in archivos_activos.values() for a in archivos
    if re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", month_from_filename(a.name))})
# Los filtros se dibujan DENTRO de Análisis energético: en iPad la barra
# lateral puede estar oculta y los usuarios no veían los controles.
# Son filtros comunes: al cambiar aquí los períodos, se actualizan TODOS
# los cálculos (PE, SSCC, MT/PD, Potencia e histórico energético).
with energia_tab:
    st.markdown("### 📅 Selección de años y meses")
    st.caption("Elige uno, dos, varios o todos los períodos. Esta selección también "
               "filtra las demás pestañas del programa. Solo se ofrecen períodos "
               "presentes en archivos de GitHub / Datos o subidos en esta sesión.")
    with st.container(border=True):
        col_anos, col_meses = st.columns(2)
        with col_anos:
            st.markdown("**Años del balance**")
            todos_anos = st.checkbox("Todos los años disponibles", value=True,
                                    key="f_todos_anos")
            if todos_anos:
                anos_seleccion = anios_disponibles
                st.caption("Seleccionados: " + (
                    ", ".join(map(str, anos_seleccion)) or "sin datos disponibles"))
            else:
                anos_seleccion = st.multiselect(
                    "Seleccionar uno o varios años",
                    anios_disponibles,
                    default=anios_disponibles[-1:] if anios_disponibles else [],
                    key="f_anos",
                )
        meses_por_anos = sorted({int(month_from_filename(archivo.name)[5:7])
            for archivos in archivos_activos.values() for archivo in archivos
            if re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", month_from_filename(archivo.name))
            and int(month_from_filename(archivo.name)[:4]) in anos_seleccion})
        with col_meses:
            st.markdown("**Meses del balance**")
            todos_meses = st.checkbox("Todos los meses disponibles", value=True,
                                     key="f_todos_meses")
            if todos_meses:
                meses_seleccion = meses_por_anos
                st.caption("Seleccionados: " + (
                    ", ".join(MONTHS_ES[m-1].capitalize() for m in meses_seleccion)
                    or "sin meses disponibles"))
            else:
                meses_seleccion = st.multiselect(
                    "Seleccionar uno o varios meses", meses_por_anos,
                    default=meses_por_anos,
                    format_func=lambda m: MONTHS_ES[m-1].capitalize(),
                    key="f_meses",
                )
        meses_balance_disponibles = sorted({month_from_filename(a.name)
            for a in archivos_activos['balance']
            if re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", month_from_filename(a.name))})
        meses_balance_elegidos = [p for p in meses_balance_disponibles
            if int(p[:4]) in anos_seleccion and int(p[5:7]) in meses_seleccion]
        st.info(f"**Balances energéticos disponibles en la selección: "
                f"{len(meses_balance_elegidos)} mes(es).** "
                "Los períodos sin balance de energía no se consideran como consumo cero.")
        if not meses_balance_elegidos:
            st.caption("Carga balances en la pestaña Datos o cambia la selección "
                       "para ver el análisis energético.")

# Resumen lateral informativo; los controles están visibles en la pestaña.
with st.sidebar:
    st.header("📅 Períodos activos")
    st.caption("Cambia los filtros desde la pestaña **Análisis energético**.")
    st.write("**Años:** " + (", ".join(map(str, anos_seleccion)) or "ninguno"))
    st.write("**Meses:** " + (", ".join(MONTHS_ES[m-1].capitalize()
                                     for m in meses_seleccion) or "ninguno"))
    st.caption(f"Archivos cargados: {sum(map(len, archivos_activos.values()))}")

def periodo_seleccionado(periodo: str) -> bool:
    return bool(re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", periodo)
        and int(periodo[:4]) in anos_seleccion and int(periodo[5:7]) in meses_seleccion)

def filtrar_periodos(archivos):
    return [a for a in archivos if periodo_seleccionado(month_from_filename(a.name))]

def un_archivo_por_mes(archivos, categoria):
    """Evita sumar dos veces preliminar y definitivo del mismo mes/categoría.

    El inventario de la pestaña Datos conserva todas las cargas para auditoría.
    """
    elegidos = {}
    avisos = []
    for a in archivos:
        mes = month_from_filename(a.name)
        n = normalize(a.name)
        prioridad = 2 if ("_DEF" in n or "DEFINITIVO" in n or re.search(r"\d{4}D(?:\.|_|$)", n)) else (0 if "PRELIMINAR" in n or re.search(r"\d{4}P(?:\.|_|$)", n) else 1)
        revision = int(re.search(r"(?:^|_)V(\d+)(?:\.|_|$)", n).group(1)) if re.search(r"(?:^|_)V(\d+)(?:\.|_|$)", n) else 0
        llave = mes
        if llave in elegidos:
            anterior, puntaje = elegidos[llave]
            if (prioridad, revision, n) > puntaje:
                avisos.append(f"{CARPETAS_CEN[categoria]} · {readable_month(mes)}: "
                              f"se utiliza {a.name} en lugar de {anterior.name}.")
                elegidos[llave] = (a, (prioridad, revision, n))
            else:
                avisos.append(f"{CARPETAS_CEN[categoria]} · {readable_month(mes)}: "
                              f"se conserva {anterior.name}; {a.name} no se suma.")
        else:
            elegidos[llave] = (a, (prioridad, revision, n))
    return [x[0] for x in sorted(elegidos.values(),key=lambda x: month_from_filename(x[0].name))], avisos

archivos_filtrados = {}
avisos_version = []
for categoria in CARPETAS_CEN:
    activos_periodo = filtrar_periodos(archivos_activos[categoria])
    if categoria == "factura":
        archivos_filtrados[categoria] = activos_periodo
    else:
        archivos_filtrados[categoria], avisos_cat = un_archivo_por_mes(activos_periodo,categoria)
        avisos_version.extend(avisos_cat)

with datos_tab:
    if avisos_version:
        with st.expander("Versiones repetidas de un mismo mes (sin doble conteo)"):
            for aviso in avisos_version:
                st.caption(aviso)

balances_catalogo_completo = archivos_filtrados["balance"]
balance_uploads = balances_catalogo_completo
invoice_uploads = archivos_filtrados["factura"]
sscc_uploads = archivos_filtrados["sscc9"]
mtpd_cen_uploads = archivos_filtrados["mtpd"]
potencia_uploads = archivos_filtrados["potencia"]
exenciaci_uploads = archivos_filtrados["exenciaci"]
price_uploads = archivos_filtrados["precio"]
sscc_empresa_uploads = archivos_filtrados["sscc4"]
horarios_disponibles = archivos_filtrados["horario"]

with datos_tab:
    st.markdown("#### Períodos activos en todas las pestañas")
    periodos_activos = sorted({month_from_filename(a.name)
        for files in archivos_activos.values() for a in files
        if periodo_seleccionado(month_from_filename(a.name))})
    if periodos_activos:
        st.write(", ".join(readable_month(p) for p in periodos_activos))
    else:
        st.info("No hay datos para los meses y años seleccionados.")
    st.markdown("#### Archivos fuente por módulo y período")
    trazabilidad = []
    for categoria, archivos in archivos_filtrados.items():
        for archivo in archivos:
            trazabilidad.append({
                "Módulo": CARPETAS_CEN[categoria],
                "Mes del archivo": readable_month(month_from_filename(archivo.name)),
                "Archivo utilizado": archivo.name,
                "Origen": "GitHub remoto" if isinstance(archivo, ArchivoGithubRemoto)
                          else ("GitHub / Datos" if isinstance(archivo, (ArchivoRepositorio, ArchivoGithubRemoto))
                          else "Cargado en esta sesión"),
            })
    if trazabilidad:
        st.dataframe(trazabilidad, hide_index=True, use_container_width=True)
    else:
        st.caption("No hay archivos fuente para los períodos seleccionados.")
    sin_periodo = [a.name for files in archivos_activos.values() for a in files
                   if month_from_filename(a.name) == "Sin período identificado"]
    if sin_periodo:
        st.warning("Archivos sin mes identificable (no se incluirán en los cálculos): "
                   + ", ".join(sin_periodo[:12])
                   + ". Comprueba que el nombre incluya el año y mes del balance CEN.")

@st.cache_data(show_spinner=False)
def calc_potencia(data: bytes, name: str):
    return parse_potencia_cen(data, name)

@st.cache_data(show_spinner=False)
def _calc_balance_cached(data: bytes, name: str, keys_tuple: tuple[str, ...]) -> dict:
    """Guardar un diccionario serializable y no una clase definida en __main__."""
    resultado = parse_balance(data, name, list(keys_tuple))
    return asdict(resultado)


def calc_balance(data: bytes, name: str, keys_tuple: tuple[str, ...]) -> BalanceResult:
    """Reconstruir BalanceResult para mantener compatibilidad con la aplicación."""
    datos = _calc_balance_cached(data, name, keys_tuple)
    return BalanceResult(**datos)

@st.cache_data(show_spinner=False)
def calc_invoice(data: bytes, name: str):
    return parse_invoice(data, name)

@st.cache_data(show_spinner=False)
def calc_mt_pd_cen_retiros(data: bytes, name: str):
    return parse_mt_pd_cen_retiros(data, name)

@st.cache_data(show_spinner=False)
def calc_sscc(data: bytes, name: str):
    return parse_sscc_retiros(data, name)

@st.cache_data(show_spinner=False)
def calc_sscc_company(data: bytes, name: str):
    return parse_sscc_empresa(data, name)

@st.cache_data(show_spinner=False)
def calc_exenciones_ci(data: bytes, name: str):
    return parse_exenciones_ci(data, name)

@st.cache_data(show_spinner=False)
def calc_sscc_bill(data: bytes, name: str):
    return parse_sscc_invoice(data, name)


balances, invoices, errors = [], [], []
exenciones_ci_reports = []
for item in exenciaci_uploads or []:
    try:
        exenciones_ci_reports.append(calc_exenciones_ci(item.getvalue(), item.name))
    except Exception as exc:
        errors.append(f"Exenciones CI {item.name}: {exc}")
exenciones_ci_reports.sort(key=lambda r: r["periodo"])
potencia_reports = []
for item in potencia_uploads or []:
    try:
        potencia_reports.append(calc_potencia(item.getvalue(), item.name))
    except Exception as exc:
        errors.append(f"Potencia CEN {item.name}: {exc}")
potencia_reports.sort(key=lambda r: r["periodo"])
mt_pd_directos = []
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
for item in mtpd_cen_uploads or []:
    try:
        mt_pd_directos.append(calc_mt_pd_cen_retiros(item.getvalue(), item.name))
    except Exception as exc:
        errors.append(f"MT/PD CEN {item.name}: {exc}")
mt_pd_directos.sort(key=lambda i: i["periodo"])
# El archivo de ~35 MB se analiza una sola vez por contenido, gracias al caché de Streamlit.
sscc_reports = []
sscc_companies = []
sscc_bills = []
for item in sscc_uploads or []:
    try:
        sscc_reports.append(calc_sscc(item.getvalue(), item.name))
    except Exception as exc:
        errors.append(f"SSCC {item.name}: {exc}")
sscc_reports.sort(key=lambda r: r['periodo'])
for item in sscc_empresa_uploads or []:
    try:
        sscc_companies.append(calc_sscc_company(item.getvalue(), item.name))
    except Exception as exc:
        errors.append(f"Archivo 4 {item.name}: {exc}")
sscc_companies.sort(key=lambda r: (r["periodo"], r["archivo"]))
# Conciliación SSCC del Archivo 9 (balance M) con proforma/factura M+1.
# La carpeta Proformas_enel no depende del filtro global del MES BALANCE;
# por ejemplo, Archivo 9 septiembre se compara con proforma OCT2026.
meses_factura_sscc = {following_month(r['periodo']) for r in sscc_reports}
fuentes_conciliacion = [*(invoice_uploads or []),
                        *(archivos_activos.get('proforma') or [])]
archivos_ya_leidos = set()
for item in fuentes_conciliacion:
    nombre = item.name.casefold()
    if nombre in archivos_ya_leidos:
        continue
    archivos_ya_leidos.add(nombre)
    # Si el nombre identifica mes, evitar abrir proformas ajenas al Archivo 9.
    periodo_nombre = month_from_filename(item.name)
    if (periodo_nombre != 'Sin período identificado'
            and periodo_nombre not in meses_factura_sscc):
        continue
    try:
        informe = calc_sscc_bill(item.getvalue(), item.name)
        if informe['mes_factura'] in meses_factura_sscc:
            sscc_bills.append(informe)
    except Exception as exc:
        # La proforma sigue visible en su propia pestaña; en Datos queda
        # advertido si no se puede usar para conciliar Servicios Complementarios.
        errors.append(f'Conciliación SSCC {item.name}: {type(exc).__name__}: {exc}')
for error in errors:
    st.error(error)

# Diagnóstico transparente: reconoce la reconstrucción y la lectura efectiva.
with datos_tab:
    st.markdown("#### Estado final · Excel divididos SSCC y Potencia")
    estados_archivos_grandes = []
    procesados_por_categoria = {
        "sscc9": {(r.get("periodo"), r.get("archivo")) for r in sscc_reports},
        "potencia": {(r.get("periodo"), r.get("archivo")) for r in potencia_reports},
    }
    for categoria in ("sscc9", "potencia"):
        for archivo in archivos_activos.get(categoria, []):
            partes = (len(archivo.rutas) if isinstance(archivo, ArchivoRepositorio)
                      else ("Sesión" if isinstance(archivo, ArchivoPartesSesion) else "Completo"))
            periodo = month_from_filename(archivo.name)
            if not periodo_seleccionado(periodo):
                estado = "Encontrado, pero excluido por los filtros"
            elif (periodo, archivo.name) in procesados_por_categoria[categoria]:
                estado = "Reconstruido y procesado"
            elif any(archivo.name in err for err in errors):
                estado = "Error al leer: revisa el mensaje rojo"
            else:
                estado = "Archivo encontrado, no seleccionado para este período"
            estados_archivos_grandes.append({
                "Categoría": CARPETAS_CEN[categoria],
                "Archivo": archivo.name,
                "Origen": ("GitHub" if isinstance(archivo,(ArchivoRepositorio,ArchivoGithubRemoto)) else "Aplicación"),
                "Partes": partes,
                "Período": periodo,
                "Resultado": estado,
            })
    if estados_archivos_grandes:
        st.dataframe(estados_archivos_grandes, hide_index=True, use_container_width=True)
    else:
        st.warning("No se encontraron libros SSCC Archivo 9 ni Potencia. "
                   "Consulta el diagnóstico de la carpeta Datos/ de GitHub o carga las partes "
                   "desde esta misma pestaña para identificar el problema.")

# Diagnóstico del archivo 4: visible incluso cuando los filtros o errores
# de lectura impiden mostrar el cálculo económico en la pestaña SSCC.
with datos_tab:
    st.markdown("#### Estado de carga · SSCC — Archivo 4")
    candidatos_4 = archivos_activos.get("sscc4", [])
    if not candidatos_4:
        st.warning("No se detectó ningún archivo 4 en el repositorio desplegado ni en las "
                   "cargas de esta sesión. Ruta recomendada: "
                   "`Datos/Sscc_archivo4/2026/4_REMUNERACION_SC_CO_Y_Pagos_Retiros_2609_def.xlsx`. "
                   "Si ya está en GitHub, comprueba que esté en la misma rama del despliegue "
                   "y reinicia Streamlit.")
    else:
        procesados = {(r['periodo'], r['archivo']) for r in sscc_companies}
        errores_4 = [e for e in errors if e.startswith('Archivo 4 ')]
        filas_estado = []
        for item in candidatos_4:
            periodo = month_from_filename(item.name)
            ruta = (f"Datos/{item.carpeta}/{item.name}" if isinstance(item, (ArchivoRepositorio, ArchivoGithubRemoto))
                    else "Carga manual desde Streamlit")
            if periodo == "Sin período identificado":
                estado = "Mes no identificado en el nombre"
            elif not periodo_seleccionado(periodo):
                estado = "Oculto por los filtros de meses/años"
            elif (periodo, item.name) in procesados:
                estado = "Leído y procesado correctamente"
            elif any(item.name in e for e in errores_4):
                estado = "Error al leer el Excel (ver mensaje rojo)"
            else:
                estado = "No seleccionado o reemplazado por otra versión del mismo mes"
            filas_estado.append({"Archivo": item.name, "Período": readable_month(periodo),
                                 "Ubicación": ruta, "Estado": estado})
        st.dataframe(filas_estado, use_container_width=True, hide_index=True)
        st.caption("Este inventario se construye únicamente con archivos presentes en Datos/ "
                   "del despliegue o cargados desde el aplicativo.")
        if sscc_companies:
            st.success(f"Archivo 4 procesado: {len(sscc_companies)} período(s) "
                       "incluido(s) en los filtros actuales.")
        elif not errores_4:
            st.info("Si el archivo figura como 'Oculto por los filtros', ve a "
                    "Análisis energético y selecciona su mes y año, o activa 'Todos'.")

def clp(number: float) -> str:
    return "$" + f"{number:,.0f}".replace(",", ".")


def clp_sscc9(number: float) -> str:
    """Importes exactos (centavos) del pago por retiro, sin truncarlos."""
    return "$" + f"{number:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")

def number_es(number: float, decimals: int = 2) -> str:
    return f"{number:,.{decimals}f}".replace(",", "_").replace(".", ",").replace("_", ".")


# -----------------------------------------------------------------------------
# Dashboard energético multianual (solo Balances oficiales del CEN).
# No mezcla consumos con cargos y no requiere facturas de Enel.
# -----------------------------------------------------------------------------
def periodo_archivo_balance(archivo) -> str | None:
    """Lee el mes por nombre, sin abrir el pesado XLSM."""
    etiqueta = month_from_filename(archivo.name)
    return etiqueta if re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", etiqueta) else None


def preferencia_version_balance(archivo) -> tuple[int, str]:
    """Preferir balance definitivo a preliminar cuando existe más de una versión."""
    nombre = normalize(archivo.name)
    if "DEFINITIVO" in nombre or "_DEF" in nombre or re.search(r"\d{4}D(?:\.|_|$)", nombre):
        prioridad = 3
    elif "PRELIMINAR" in nombre or re.search(r"\d{4}P(?:\.|_|$)", nombre):
        prioridad = 1
    else:
        prioridad = 2
    return prioridad, nombre


def catalogo_balances_por_mes(archivos: list) -> tuple[dict[str, object], list[str]]:
    """Un único balance por mes; informa duplicados y nombres sin fecha."""
    elegidos = {}
    avisos = []
    for archivo in archivos:
        periodo = periodo_archivo_balance(archivo)
        if periodo is None:
            avisos.append(f"No se puede identificar el mes de {archivo.name}; revisa el nombre del archivo.")
            continue
        anterior = elegidos.get(periodo)
        if anterior is None:
            elegidos[periodo] = archivo
        else:
            preferido = max((anterior, archivo), key=preferencia_version_balance)
            elegidos[periodo] = preferido
            avisos.append(f"Hay varias versiones para {readable_month(periodo)}: "
                          f"se usa {preferido.name} en el análisis energético.")
    return elegidos, avisos


def datos_energeticos_resumidos(balances_seleccionados: list[BalanceResult]) -> list[dict]:
    """Una fila por período. MWh y porcentajes calculados desde los retiros físicos."""
    resumen = []
    for b in sorted(balances_seleccionados, key=lambda x: x.label):
        puntos = balance_consumption_by_site(b)
        q, l = puntos["Quilpué"] / 1000.0, puntos["Limache"] / 1000.0
        total = q + l
        if not re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", b.label):
            continue
        year, month = map(int, b.label.split("-"))
        resumen.append({
            "periodo": b.label, "anio": year, "mes_num": month,
            "Quilpué (MWh)": q, "Limache (MWh)": l, "Total EFE (MWh)": total,
            "Enel (MWh)": b.retiro_enel_mwh, "SEN (MWh)": b.retiro_sen_mwh,
            "Participación EFE / Enel (%)": total / b.retiro_enel_mwh * 100 if b.retiro_enel_mwh else 0,
        })
    return resumen

if not any(archivos_activos.values()):
    with energia_tab:
        st.info("Carga archivos en Datos para activar el análisis. La aplicación no muestra valores de ejemplo.")
    st.stop()

with energia_tab:
    st.subheader("Análisis energético · EFE Valparaíso")
    st.caption("Origen: archivos cargados en Datos. Usa los filtros de años y meses que aparecen arriba. "
               "Los meses sin archivos no se interpretan como consumo cero.")
    catalogo, avisos_catalogo = catalogo_balances_por_mes(balances_catalogo_completo)
    for aviso in avisos_catalogo:
        st.warning(aviso)
    if not catalogo:
        st.info("No hay balances CEN para los meses y años seleccionados. "
                "Cárgalos en Datos y ajusta el filtro global.")
    else:
        resultados_energia = []
        for per, archivo in sorted(catalogo.items()):
            try:
                resultados_energia.append(calc_balance(archivo.getvalue(), archivo.name, tuple(keys)))
            except Exception as exc:
                st.warning(f"No se pudo leer {archivo.name} ({per}): {exc}")
        datos_energia = datos_energeticos_resumidos(resultados_energia)
        if not datos_energia:
            st.info("No se pudo calcular consumo desde los balances elegidos.")
        else:
            q_total = sum(r["Quilpué (MWh)"] for r in datos_energia)
            l_total = sum(r["Limache (MWh)"] for r in datos_energia)
            efe_total = q_total + l_total
            a, b, c, d = st.columns(4)
            a.metric("Quilpué · total elegido",f"{number_es(q_total,2)} MWh")
            b.metric("Limache · total elegido",f"{number_es(l_total,2)} MWh")
            c.metric("EFE · total elegido",f"{number_es(efe_total,2)} MWh")
            d.metric("Meses disponibles",str(len(datos_energia)))
            st.markdown("#### Consumo por mes y por suministro")
            # Los meses se muestran en palabras con tres letras; el año evita
            # confundir, por ejemplo, "Sep 2025" con "Sep 2026".
            graf = [{"Mes":mes_corto_grafico(r["periodo"]),
                     "Quilpué":r["Quilpué (MWh)"],
                     "Limache":r["Limache (MWh)"]} for r in datos_energia]
            try:
                import pandas as pd
                tabla_graf = pd.DataFrame(graf).set_index("Mes")
                st.bar_chart(tabla_graf)
            except ImportError:
                st.dataframe(graf, hide_index=True, use_container_width=True)
            st.dataframe([{
                "Mes balance":readable_month(r["periodo"]),
                "Mes factura":readable_month(following_month(r["periodo"])),
                "Quilpué (MWh)":round(r["Quilpué (MWh)"],3),
                "Limache (MWh)":round(r["Limache (MWh)"],3),
                "Total EFE (MWh)":round(r["Total EFE (MWh)"],3),
                "Total Enel (MWh)":round(r["Enel (MWh)"],3),
                "Participación EFE en Enel (%)":round(r["Participación EFE / Enel (%)"],5)
            } for r in datos_energia], hide_index=True,use_container_width=True)
            resumen_anual = []
            for anio in sorted({r["anio"] for r in datos_energia}):
                sub = [r for r in datos_energia if r["anio"] == anio]
                resumen_anual.append({"Año":anio,"Meses con balance":len(sub),
                                     "Quilpué (MWh)":round(sum(r["Quilpué (MWh)"] for r in sub),3),
                                     "Limache (MWh)":round(sum(r["Limache (MWh)"] for r in sub),3),
                                     "Total EFE (MWh)":round(sum(r["Total EFE (MWh)"] for r in sub),3)})
            st.markdown("#### Comparación entre años seleccionados")
            st.dataframe(resumen_anual,hide_index=True,use_container_width=True)
            st.download_button("⬇ Descargar análisis energético filtrado CSV",
                               csv_bytes(datos_energia),"EFE_analisis_energetico_filtrado.csv","text/csv")

with calc_tab:
    st.subheader("Compensación por Precio Estabilizado (PE)")
    st.caption("Solo balances CEN incorporados en Datos/ de GitHub o cargados en esta sesión. "
               "No se crean meses mediante fechas manuales ni se utilizan archivos enviados por ChatGPT.")
    if not balances:
        st.info("No hay balances CEN válidos para los períodos seleccionados. "
                "Carga el archivo Balance del mes en Datos/ de GitHub o en la pestaña Datos. "
                "No se muestran importes ni períodos de ejemplo.")
    else:
        st.markdown("#### Meses de PE respaldados por archivos")
        archivos_por_periodo = {month_from_filename(a.name): a for a in balance_uploads}
        resumen_fuentes = [{
            "Mes balance": readable_month(b.label),
            "Mes factura asociado": readable_month(following_month(b.label)),
            "Archivo": archivos_por_periodo[b.label].name if b.label in archivos_por_periodo else "",
            "Origen": ("GitHub / Datos" if isinstance(archivos_por_periodo.get(b.label), (ArchivoRepositorio, ArchivoGithubRemoto))
                       else "Cargado en esta sesión"),
        } for b in balances]
        st.dataframe(resumen_fuentes, hide_index=True, use_container_width=True)
        choice = st.selectbox(
            "Selecciona un mes que tenga Balance CEN disponible",
            range(len(balances)),
            index=len(balances) - 1,
            format_func=lambda ix: f"{readable_month(balances[ix].label)} · "
                                   f"{number_es(balances[ix].tarifa_clp_mwh)} $/MWh",
            key="pe_balance_fuente_v27",
        )
        chosen = balances[choice]
        period = chosen.label
        invoice_month = following_month(period)
        source_file = archivos_por_periodo.get(period)
        a, b = st.columns(2)
        a.info(f"**Mes del balance CEN:** {readable_month(period)}")
        b.info(f"**Mes de factura asociado:** {readable_month(invoice_month)}")
        if source_file:
            source = "GitHub / Datos" if isinstance(source_file, (ArchivoRepositorio, ArchivoGithubRemoto)) else "Cargado en Streamlit"
            st.caption(f"Fuente utilizada: {source_file.name} · {source}")

        usages = balance_consumption_by_site(chosen)
        rates = {site: chosen.tarifa_clp_mwh for site in SITE_NAMES}
        rows = projected_by_site(usages, rates, 100.0,
                                 {site: 0.0 for site in SITE_NAMES}, 20)
        total_charge = sum(x["cargo_clp"] for x in rows)
        total_energy = sum(x["kwh"] for x in rows)
        col1, col2, col3 = st.columns(3)
        col1.metric("Compensación Quilpué", clp(rows[0]["cargo_clp"]))
        col2.metric("Compensación Limache", clp(rows[1]["cargo_clp"]))
        col3.metric("Total EFE (sin IVA)", clp(total_charge))
        detail_rows = [{
            "Mes balance": readable_month(period),
            "Mes factura asociado": readable_month(invoice_month),
            "Archivo de origen": source_file.name if source_file else "",
            "Suministro": x["punto"], "Retiro (kWh)": round(x["kwh"], 2),
            "Compensación media SEN ($/MWh)": round(x["tarifa_clp_mwh"], 3),
            "Estimación ($)": round(x["cargo_clp"]),
        } for x in rows]
        detail_rows.append({
            "Mes balance": readable_month(period),
            "Mes factura asociado": readable_month(invoice_month),
            "Archivo de origen": source_file.name if source_file else "",
            "Suministro": "TOTAL EFE", "Retiro (kWh)": round(total_energy, 2),
            "Compensación media SEN ($/MWh)": round(chosen.tarifa_clp_mwh, 3),
            "Estimación ($)": round(total_charge),
        })
        st.dataframe(detail_rows, hide_index=True, use_container_width=True)
        st.warning("Estimación mediante el promedio mensual del SEN y retiros del balance CEN. "
                   "No equivale a la asignación horaria definitiva por retiro.")
        if any(usages[site] == 0 for site in SITE_NAMES):
            st.warning("Hay retiros de EFE no identificados en este balance; verifica las claves comerciales.")
        with st.expander("Trazabilidad del cálculo y resultados del balance"):
            c1, c2 = st.columns(2)
            c1.metric("Compensación agregada SEN ($)", clp(chosen.compensacion_sen_clp))
            c2.metric("Retiros SEN (MWh)", number_es(chosen.retiro_sen_mwh, 3))
            st.dataframe(chosen.detalle, hide_index=True, use_container_width=True)
            for aviso in chosen.avisos:
                st.warning(aviso)
        st.download_button("⬇ Descargar cálculo del balance seleccionado (CSV)",
            data=csv_bytes(detail_rows),
            file_name=f"Compensacion_PE_EFE_{period}.csv", mime="text/csv",
            key="descarga_pe_v27")

with history_tab:
    st.caption("Histórico filtrado según los períodos seleccionados en la barra lateral. Fuente: Datos.")
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
    st.write("Carga en la pestaña Datos un CSV con la compensación y los retiros horarios del SEN, "
             "Quilpué y Limache. La aplicación calcula el cargo para cada punto y su total.")
    st.download_button("⬇ Descargar plantilla horaria CSV", data=csv_bytes([{
        "fecha":"", "hora":"", "compensacion_sen_clp":"",
        "retiro_sen_mwh":"", "retiro_quilpue_mwh":"", "retiro_limache_mwh":""
    }]), file_name="plantilla_horaria_quilpue_limache.csv", mime="text/csv")
    hourly_file = None
    if horarios_disponibles:
        hidx = st.selectbox("Archivo horario guardado",list(range(len(horarios_disponibles))),
            index=len(horarios_disponibles)-1,
            format_func=lambda i: f"{readable_month(month_from_filename(horarios_disponibles[i].name))} · {horarios_disponibles[i].name}")
        hourly_file = horarios_disponibles[hidx]
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
    if not hourly_file:
        st.info("Sube el archivo CSV horario desde la pestaña Datos para analizarlo aquí.")
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

**Prorrateo SSCC + CO desde Archivo 4:** se leen las hojas `PAGO_EMPRESAS_SC_SSCC`
y `PAGO_EMPRESAS_CO` para `ENEL_GENERACION`. Para cada una se calcula costo neto
`PAGA (C) − RECIBE (B)`, conservando el signo; la columna D original del CEN utiliza
la convención contraria (`RECIBE−PAGA`). Se suman SSCC neto y CO neto, y el total
se multiplica por la proporción de retiros de EFE dentro de los retiros de Enel en
el mismo balance. Es **una estimación**, no un cobro individual oficial de EFE,
y se excluye del total general porque el Archivo 9 se considera por separado.

**BESS en estimación:** el Archivo 4 no tiene asignación individual BESS. La pestaña
SSCC puede sumar una tasa BESS histórica de un mes anterior como **supuesto**
separado del neto SSCC+CO. Si no existe historial, la estimación no incluye BESS.
El Archivo 9 sigue calculándose de forma independiente y como un único cargo con
sus 12 tipos de pago en Resumen.

**Potencia de Suficiencia:** archivo de Balance Definitivo Detalle del CEN, hoja `Balance2`;
solo se consideran las filas identificadas por **nombre comercial Y barra**. La columna E es
potencia medida en MW (retiros negativos), F es precio nudo en CLP/kW/mes y G es Pago PSUF
en CLP, normalmente negativo para retiros. El importe presentado corresponde a magnitud
del PSUF individual, **no automáticamente al precio contractual facturado** por Enel.

**Privacidad:** las cargas de Datos/ en la aplicación duran la sesión, mientras que los archivos de un repositorio persisten hasta borrarlos. Los archivos Excel se procesan en el servidor de Streamlit donde alojes la
aplicación, sin ejecutar macros. **Los archivos guardados en carpetas de GitHub son visibles si el repositorio es público**. Asegúrate de contar con autorización empresarial para publicarlos y procesarlos en servicios externos.
""")
    if price_uploads:
        precio_idx = st.selectbox("Mes de precios estabilizados",range(len(price_uploads)),
            index=len(price_uploads)-1,
            format_func=lambda i: f"{readable_month(month_from_filename(price_uploads[i].name))} · {price_uploads[i].name}",
            key="precio_mes_repositorio")
        price_upload = price_uploads[precio_idx]
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


with sscc_tab:
    st.caption("Fuente: Datos · Años y meses filtrados globalmente. BESS está incluido una sola vez en SSCC.")
    st.subheader('Servicios Complementarios (SSCC) · Quilpué y Limache')
    st.caption('Se muestra una estimación construida con el archivo 4 y, si lo cargas, '
               'la asignación individual del archivo 9. No son el mismo resultado.')
    if sscc_reports:
        st.markdown("#### Histórico oficial SSCC de los períodos seleccionados")
        historial_sscc = []
        for r in sscc_reports:
            for sitio in SITE_NAMES:
                historial_sscc.append({"Balance CEN":readable_month(r["periodo"]),
                        "Factura":readable_month(following_month(r["periodo"])),
                        "Suministro":sitio,
                        "SSCC con BESS ($)":round(r["subtotal_base"][sitio],2),
                        "Archivo":r["archivo"]})
        st.dataframe(historial_sscc,hide_index=True,use_container_width=True)
        st.download_button("Descargar histórico SSCC filtrado",csv_bytes(historial_sscc),
                           "EFE_SSCC_periodos_filtrados.csv","text/csv")
    st.markdown('### 1. ¿Qué significa el archivo 4?')
    sscc_company = None
    if sscc_companies:
        seleccion = st.selectbox('Mes de balance del archivo 4',
            list(range(len(sscc_companies))),index=len(sscc_companies)-1,
            format_func=lambda i: f"{readable_month(sscc_companies[i]['periodo'])} · {sscc_companies[i]['archivo']}",
            key='sscc4_period_selector')
        sscc_company = sscc_companies[seleccion]
    st.write('El **archivo 4** muestra los pagos y remuneraciones de **Enel Generación en todo el SEN**, '
             '**no lo que se cobra directamente a EFE Valparaíso**. Para estimar un cargo de EFE '
             'hay que combinarlo con los retiros del Balance CEN.')
    if sscc_company:
        periodo_empresa = sscc_company['periodo']
        st.info(f'**Mes del balance:** {readable_month(periodo_empresa)}   |   '
                f'**Mes de factura estimado:** {readable_month(following_month(periodo_empresa))}')
        saldos_empresa = calcular_netos_empresa(sscc_company)
        st.markdown('#### Paso 1 · Saldos netos SSCC y CO de Enel')
        st.caption('Se utilizan **PAGA − RECIBE** para ambos conceptos; CO negativo reduce el costo. '
                   'Los saldos pertenecen a Enel a nivel SEN, no son cobros individuales de EFE.')
        filas_empresa=[]
        for codigo, nombre, hoja in (
            ('SC_SSCC','Servicios Complementarios','PAGO_EMPRESAS_SC_SSCC'),
            ('CO','Costos de Oportunidad','PAGO_EMPRESAS_CO')):
            c=saldos_empresa[codigo]
            fila=sscc_company[codigo]['fila']
            filas_empresa.append({
                'Concepto':nombre,
                'Enel PAGA ($)':round(c['paga'],2),
                'Enel RECIBE ($)':round(c['recibe'],2),
                'Costo neto PAGA−RECIBE ($)':round(c['costo_neto'],2),
                'Hoja CEN':hoja,
                'Celdas':f'B{fila}:D{fila}',
            })
        filas_empresa.append({
            'Concepto':'TOTAL SSCC + CO',
            'Enel PAGA ($)':round(saldos_empresa['total_paga'],2),
            'Enel RECIBE ($)':round(saldos_empresa['total_recibe'],2),
            'Costo neto PAGA−RECIBE ($)':round(saldos_empresa['total_costo_neto'],2),
            'Hoja CEN':'—','Celdas':'—'})
        st.dataframe(filas_empresa,hide_index=True,use_container_width=True)
        st.metric('Costo neto Enel · SSCC + CO',clp(saldos_empresa['total_costo_neto']))
        st.caption('La columna D original del CEN representa RECIBE−PAGA, por eso tiene el signo '
                   'contrario al costo neto mostrado. CO puede ser un abono neto.')
        compatible_balances = [b for b in balances if b.label == periodo_empresa]
        if not compatible_balances:
            st.warning('**Falta el Balance CEN del mismo mes.** Para repartir el pago general de Enel '
                       f'debes cargar un Balance de {readable_month(periodo_empresa)} '
                       'correspondiente al mismo mes.')
        else:
            try:
                selected_balance_index = st.selectbox(
                    'Balance de retiros usado para la estimación',
                    list(range(len(compatible_balances))),
                    format_func=lambda i: f"{readable_month(compatible_balances[i].label)} · "
                                          f"{number_es(compatible_balances[i].retiro_enel_mwh,3)} MWh Enel",
                    key='sscc_balance_para_prorrateo')
                selected_balance=compatible_balances[selected_balance_index]
                st.markdown('#### Paso 2 · Calcular la participación de cada suministro')
                st.write('El Balance CEN indica cuánta energía retiró **Enel en total** '
                         'y cuánta corresponde a **Quilpué y Limache**. '
                         'El programa reparte de forma aproximada el pago global según esos retiros.')
                st.metric('Retiros totales de Enel',
                          number_es(selected_balance.retiro_enel_mwh,3)+' MWh')
                st.caption('Balance Valorizado: suministrador ENEL_GENERACION, '
                           'tipos de retiro L, L_D y R. Se consideran magnitudes positivas de retiros.')
                referencia_bess, fuente_bess = tasa_bess_referencia_previa(
                    sscc_reports, balances, periodo_empresa)
                # Solo se emplean valores BESS de archivos CEN realmente disponibles.
                # No se permiten tarifas introducidas manualmente como si fueran datos fuente.
                bess_unit = float(referencia_bess) if referencia_bess else 0.0
                with st.expander('Origen documental de BESS incluido en SSCC'):
                    st.write('**BESS ya es parte de SSCC.** La referencia histórica se calcula '
                             'exclusivamente con los archivos CEN cargados en Datos.')
                    if bess_unit:
                        st.metric('Referencia histórica BESS ($/kWh)',number_es(bess_unit,6))
                        st.caption(fuente_bess)
                    else:
                        st.warning('No hay archivos anteriores compatibles para obtener BESS. '
                                   'La estimación se limita al neto SSCC + CO del archivo 4; no incluye BESS.')
                prorrateo=calcular_prorrateo_sscc(sscc_company,selected_balance,bess_unit)
                table_participacion=[]
                for rec in prorrateo['detalle']:
                    table_participacion.append({
                        'Suministro':rec['punto'],
                        'Retiro EFE (MWh)':number_es(rec['retiro_kwh']/1000,3),
                        'Parte de los retiros Enel (%)':number_es(rec['participacion_enel_pct'],6),
                        'SSCC neto ($)':clp(rec['sscc_neto_estimado_clp']),
                        'CO neto ($)':clp(rec['co_neto_estimado_clp']),
                        'SSCC + CO neto ($)':clp(rec['sscc_co_neto_estimado_clp'])})
                st.dataframe(table_participacion,hide_index=True,use_container_width=True)
                st.markdown('**Así se obtiene el monto estimado:**')
                st.code('[(PAGA − RECIBE) SSCC + (PAGA − RECIBE) CO] × (MWh del suministro ÷ MWh totales Enel)',language=None)
                st.markdown('#### Paso 3 · ¿Cuánto se estima para EFE?')
                if bess_unit:
                    st.caption('Se muestra neto SSCC + CO; adicionalmente se incorpora BESS como supuesto histórico, no como dato del Archivo 4.')
                else:
                    st.warning('**Estimación referencial:** incluye SSCC neto + CO neto del Archivo 4, '
                               'sin BESS histórico. No es el cobro definitivo de EFE.')
                k1,k2,k3=st.columns(3)
                k1.metric('Quilpué · estimado',clp(prorrateo['detalle'][0]['sscc_estimado_clp']))
                k2.metric('Limache · estimado',clp(prorrateo['detalle'][1]['sscc_estimado_clp']))
                k3.metric('Total EFE · neto estimado',clp(prorrateo['sscc_total_estimado']))
                st.caption('**No es el cobro oficial del CEN.** El prorrateo mensual puede diferir '
                           'del reparto real por servicios, horarios, zonas y otros criterios.')
                report_same_month = next((r for r in sscc_reports
                                          if r['periodo']==prorrateo['periodo']),None)
                if report_same_month:
                    st.markdown('#### Paso 4 · Comparación con el archivo 9 (asignación individual)')
                    st.success('Se encontró el archivo 9 del mismo mes. '
                               'Ese archivo contiene los importes asignados a los retiros de EFE.')
                    st.info('Archivo 4: saldo neto global de Enel prorrateado; Archivo 9: asignación '
                            'por retiro para EFE. No son importes directamente conciliables.')
                    total_9 = sum(report_same_month.get('suma_tipos', {}).values())
                    st.metric('Archivo 9 · total de pagos de EFE (12 tipos)',clp(total_9))
                    st.caption('Se mantiene el Archivo 9 como un único cobro en Resumen. '
                               'Archivo 4, incluso al incorporar CO, no se suma al total de factura.')
                export_rows=[{
                    'mes_balance':prorrateo['periodo'],
                    'mes_factura':prorrateo['mes_factura'],
                    'metodo':'Costo neto (PAGA-RECIBE) SSCC + CO, prorrateado por retiros Enel',
                    'punto':rec['punto'],'clave':rec['clave'],
                    'pago_global_enel_sc_sscc_clp':prorrateo['pago_enel_sc_sscc_clp'],
                    'recibe_global_enel_sc_sscc_clp':prorrateo['recibe_enel_sc_sscc_clp'],
                    'neto_global_enel_sc_sscc_clp':prorrateo['costo_neto_sc_sscc_clp'],
                    'paga_global_enel_co_clp':prorrateo['paga_enel_co_clp'],
                    'recibe_global_enel_co_clp':prorrateo['recibe_enel_co_clp'],
                    'neto_global_enel_co_clp':prorrateo['costo_neto_co_clp'],
                    'neto_global_total_enel_clp':prorrateo['costo_neto_total_clp'],
                    'retiro_total_enel_mwh':prorrateo['retiro_enel_mwh'],
                    'retiro_punto_mwh':rec['retiro_kwh']/1000,
                    'participacion_enel_pct':rec['participacion_enel_pct'],
                    'sscc_neto_estimado_clp':rec['sscc_neto_estimado_clp'],
                    'co_neto_estimado_clp':rec['co_neto_estimado_clp'],
                    'sscc_co_neto_estimado_clp':rec['sscc_co_neto_estimado_clp'],
                    'referencia_bess_incluida_en_sscc_clp_kwh':bess_unit,
                    'sscc_estimado_clp':rec['sscc_estimado_clp'],
                    'caracter_resultado':'Neto SSCC+CO sin BESS' if bess_unit==0 else 'Neto SSCC+CO con supuesto histórico BESS',
                    'archivo_4':sscc_company['archivo'],
                    'archivo_balance':selected_balance.label,
                } for rec in prorrateo['detalle']]
                st.download_button('⬇ Descargar explicación y estimación (CSV)',
                    data=csv_bytes(export_rows),
                    file_name=f"estimacion_SSCC_CO_neto_{prorrateo['periodo']}_factura_{prorrateo['mes_factura']}.csv",
                    mime='text/csv',key='export_sscc_prorrateo')
            except Exception as exc:
                st.error(f'No se pudo calcular el prorrateo con el archivo 4: {exc}')
    else:
        if archivos_activos.get("sscc4"):
            st.warning('Hay archivos 4 registrados en **Datos**, pero ninguno quedó disponible '
                       'en esta pestaña. Comprueba el filtro de años/meses en Análisis energético '
                       'y los mensajes de procesamiento del diagnóstico de la pestaña Datos.')
        else:
            st.info('No se detectó ningún archivo 4. Súbelo en **Datos → SSCC · Archivo 4**, '
                    'o ubícalo en GitHub dentro de **Datos/Sscc_archivo4/2026/**. '
                    'Agrega además el Balance CEN del mismo mes para calcular un prorrateo. '
                    'El archivo 9 es opcional para validar el importe individual.')
    st.divider()
    st.markdown('### 2. Asignación individual del CEN · Archivo 9')
    st.caption('El archivo 9_Pagos_Retiros se filtra por suministrador ENEL_GENERACION '
               'y claves MERVAL_ELSOL (Quilpué) / MERVAL_LIMACHE (Limache).')
    if not sscc_reports:
        st.info('Carga un archivo 9 del CEN desde la pestaña Datos para ver los importes oficiales individualizados.')
    else:
        report_idx = st.selectbox('Balance SSCC a consultar', list(range(len(sscc_reports))),
            index=len(sscc_reports)-1,
            format_func=lambda i: f"{readable_month(sscc_reports[i]['periodo'])} · {sscc_reports[i]['archivo']}")
        report = sscc_reports[report_idx]
        period = report['periodo']
        invoice_month = following_month(period)
        col1,col2 = st.columns(2)
        col1.info(f'**Balance CEN:** {readable_month(period)}')
        col2.info(f'**Factura esperada:** {readable_month(invoice_month)}')
        st.info('**BESS está incorporado dentro de SSCC.** El programa suma el registro '
                '`SC_SSCC` con el registro `BESS` y presenta **un solo cargo SSCC por suministro**, '
                'sin duplicarlo. Los demás tipos de pago del CEN se muestran separados para conciliación.')
        agrupado = agrupar_sscc(report)
        table = []
        for code in SSCC_CATEGORIAS:
            q = agrupado['Quilpué'][code]
            l = agrupado['Limache'][code]
            table.append({'Concepto':code,'Quilpué ($)':round(q,2),
                          'Limache ($)':round(l,2),'Total EFE ($)':round(q+l,2)})
        st.dataframe(table,hide_index=True,use_container_width=True)
        def _monto_sscc_completo(sitio):
            ausentes = report['faltantes'].get(sitio, [])
            return (agrupado[sitio]['SSCC'] if not any(
                    codigo in ausentes for codigo in ('SC_SSCC', 'BESS')) else None)
        qbase = _monto_sscc_completo('Quilpué')
        lbase = _monto_sscc_completo('Limache')
        totalbase = qbase + lbase if qbase is not None and lbase is not None else None
        a,b,c=st.columns(3)
        a.metric('SSCC · Quilpué',clp_sscc9(qbase) if qbase is not None else 'Pendiente')
        b.metric('SSCC · Limache',clp_sscc9(lbase) if lbase is not None else 'Pendiente')
        c.metric('Total SSCC EFE (sin IVA)',clp_sscc9(totalbase) if totalbase is not None else 'Pendiente')
        st.caption('Estos importes proceden de PAGO_RETIRO, columna J; filtro ENEL_GENERACION '
                   'y claves MERVAL_ELSOL / MERVAL_LIMACHE. Corresponden a la factura del mes siguiente.')
        total_all = sum(sum(c.values()) for c in agrupado.values())
        st.write('**Total informativo de todos los tipos de pago, incluido SSCC una sola vez**: '
                 f"Quilpué {clp(sum(agrupado['Quilpué'].values()))}; "
                 f"Limache {clp(sum(agrupado['Limache'].values()))}; "
                 f"total {clp(total_all)}. "
                 'Los demás tipos no se incorporan automáticamente a SSCC facturado por Enel.')
        if any(report['faltantes'][site] for site in SITE_NAMES):
            st.warning('Faltan tipos de pago para alguna clave; revisar hoja PAGO_RETIRO: '+
                '; '.join(f"{site}: {', '.join(report['faltantes'][site])}" for site in SITE_NAMES
                          if report['faltantes'][site]))
        # Construir una estimación $/kWh por sitio a partir del retiro del balance, si existe.
        match_balance = next((b for b in balances if b.label == period), None)
        if match_balance:
            energies=balance_consumption_by_site(match_balance)
            st.markdown('#### Referencias unitarias del balance (no tarifas oficiales fijas)')
            rates=[]
            for site in SITE_NAMES:
                kwh=energies[site]
                rate=report['subtotal_base'][site]/kwh if kwh else 0.
                rates.append({'Punto':site,'Retiro balance (kWh)':round(kwh,2),
                              'SSCC ($)':round(agrupado[site]['SSCC'],2),
                              'Equivalente ($/kWh)':round(rate,6) if kwh else None})
            st.dataframe(rates,hide_index=True,use_container_width=True)
            st.caption('El SSCC incluye BESS: se divide el cargo combinado por los retiros del Balance Valorizado del mismo mes.')
        else:
            st.caption('Para obtener un equivalente $/kWh por punto, carga también el archivo '
                       f'Balance_{period.replace("-", "")[2:]}P.xlsm del mismo mes.')
        st.download_button('⬇ Descargar cargos consolidados (CSV)',
            data=csv_bytes(exportar_sscc_agrupado(report)),
            file_name=f'SSCC_consolidado_EFE_balance_{period}_factura_{invoice_month}.csv', mime='text/csv')
        with st.expander('Auditoría del origen (códigos y filas originales del CEN)'):
            st.caption('Para verificar el cálculo, se conservan los registros originales. '
                       'SC_SSCC y BESS se agrupan en un único SSCC en pantalla y en el CSV consolidado.')
            st.dataframe(report['detalle'], hide_index=True, use_container_width=True)
            st.download_button('Descargar detalle técnico original (CSV)',
                data=csv_bytes([{'mes_balance':period,'mes_factura':invoice_month,
                                 'punto':x['punto'],'clave':x['clave'],
                                 'barra':x['barra'],'tipo_pago_original':x['codigo'],
                                 'pago_original_clp':round(x['valor_clp'],5),
                                 'fuente':report['archivo'],'fila':x['fila']}
                                for x in report['detalle']]),
                file_name=f'SSCC_auditoria_origen_{period}.csv', mime='text/csv')
        empresa_del_mes = next((x for x in sscc_companies if x['periodo']==period),None)
        if empresa_del_mes:
            with st.expander('Respaldo: totales de Enel Generación en el archivo 4'):
                st.write('Estos importes son **a nivel de la empresa Enel Generación en todo el SEN**; '
                         'no deben sumarse a los cargos de Quilpué y Limache.')
                netos_del_mes=calcular_netos_empresa(empresa_del_mes)
                company_table=[{'Concepto':c,'Empresa': 'ENEL_GENERACION',
                                'Recibe ($)':round(netos_del_mes[c]['recibe'],2),
                                'Paga ($)':round(netos_del_mes[c]['paga'],2),
                                'Costo neto PAGA−RECIBE ($)':round(netos_del_mes[c]['costo_neto'],2)}
                               for c in ('SC_SSCC','CO')]
                company_table.append({'Concepto':'TOTAL SSCC + CO','Empresa':'ENEL_GENERACION',
                                      'Recibe ($)':round(netos_del_mes['total_recibe'],2),
                                      'Paga ($)':round(netos_del_mes['total_paga'],2),
                                      'Costo neto PAGA−RECIBE ($)':round(netos_del_mes['total_costo_neto'],2)})
                st.dataframe(company_table,hide_index=True,use_container_width=True)
                st.caption('Referencia a nivel de empresa: no se suma al Archivo 9.')
        elif sscc_companies:
            st.caption('No hay un archivo 4 del mismo mes que el archivo 9 seleccionado.')
        matching_bills=[i for i in sscc_bills if i['periodo_balance'] == period]
        if matching_bills:
            with st.expander('Comparación con facturas Enel del mismo período'):
                for bill in matching_bills:
                    st.write(f"**Factura {readable_month(bill['mes_factura'])}** — {bill['archivo']}")
                    st.dataframe([{'Punto':site,'SSCC proforma ($)':round(bill['importes'][site],2),
                                   'SSCC (incluye BESS) ($)':round(agrupado[site]['SSCC'],2),
                                   'Diferencia ($)':round(bill['importes'][site]-agrupado[site]['SSCC'],2)}
                                  for site in SITE_NAMES],hide_index=True,use_container_width=True)
                    st.caption('Diferencia no implica automáticamente error: Enel podría incluir otros tipos, ajustes o reliquidaciones.')
        elif sscc_bills:
            st.caption('La factura subida corresponde a otro mes de balance. No se comparan importes de meses distintos.')



with potencia_tab:
    st.caption("Fuente: Datos · Años y meses filtrados globalmente.")
    st.subheader("Potencia de Suficiencia · Quilpué y Limache")
    st.write("Carga el archivo **BDef Detalle** del CEN. No se requiere factura: "
             "la aplicación lee directamente **Balance2**, columnas B a G.")
    st.caption("Criterio de identificación: Quilpué = METRO REGIONAL DE VALPARAISO SA, "
               "barra QUILPUE 110; Limache = MERVAL LIMACHE, barra S.PEDRO 13. "
               "No se suman todas las filas que mencionan EFE, porque pueden pertenecer a otras regiones.")
    if potencia_reports:
        resumen_potencia = []
        for resultado in potencia_reports:
            for sitio in SITE_NAMES:
                rec = resultado["detalle"].get(sitio)
                resumen_potencia.append({"Balance":readable_month(resultado["periodo"]),
                       "Factura":readable_month(resultado["mes_factura"]),
                       "Suministro":sitio,
                       "Potencia (MW)":round(rec["potencia_retiro_mw"],5) if rec else None,
                       "PSUF CEN ($)":round(rec["monto_retiro_psuf_clp"],2) if rec else None})
        st.markdown("#### Histórico de potencia seleccionado")
        st.dataframe(resumen_potencia,hide_index=True,use_container_width=True)
        st.download_button("Descargar histórico PSUF filtrado",csv_bytes(resumen_potencia),
                           "EFE_potencia_periodos_filtrados.csv","text/csv")
    if not potencia_reports:
        st.info("Sube un archivo de **Potencia BDef Detalle** en la pestaña Datos, sección "
                "**Potencia de Suficiencia** en la pestaña Datos.")
    else:
        opciones = list(range(len(potencia_reports)))
        idx = st.selectbox("Selecciona el balance de potencia", opciones,
            index=len(opciones)-1,
            format_func=lambda i: f"{readable_month(potencia_reports[i]['periodo'])} · {potencia_reports[i]['archivo']}",
            key="potencia_v19_selector")
        resultado = potencia_reports[idx]
        st.info(f"**Balance CEN:** {readable_month(resultado['periodo'])} · "
                f"**Factura de referencia:** {readable_month(resultado['mes_factura'])} · "
                f"**Fijación precios:** {resultado['fijacion'] or 'No identificada'}")
        tabla=[]
        for sitio in SITE_NAMES:
            rec=resultado["detalle"].get(sitio)
            if rec is None:
                tabla.append({"Suministro":sitio,"Potencia retirada (MW)":None,
                              "Precio nudo ($/kW/mes)":None,"Monto PSUF del CEN ($)":None,
                              "Origen":"Sin coincidencia validada"})
            else:
                tabla.append({"Suministro":sitio,
                              "Potencia retirada (MW)":round(rec["potencia_retiro_mw"],4),
                              "Precio nudo ($/kW/mes)":round(rec["precio_nudo_clp_kw_mes"],2),
                              "Monto PSUF del CEN ($)":round(rec["monto_retiro_psuf_clp"],2),
                              "Origen":f"Balance2!B{rec['fila']}:G{rec['fila']}"})
        st.dataframe(tabla,hide_index=True,use_container_width=True)
        if resultado["completo"]:
            a,b,c=st.columns(3)
            a.metric("Potencia total EFE (MW)",number_es(resultado["total_potencia_retiro_mw"],4))
            b.metric("PSUF total EFE ($)",clp(resultado["total_monto_retiro_clp"]))
            c.metric("Mes de factura",readable_month(resultado["mes_factura"]))
        else:
            st.warning("Falta identificar uno de los dos retiros. No debe considerarse "
                       "el subtotal como total EFE hasta resolver las coincidencias.")
        st.latex(r"\mathrm{Pago\ PSUF}=\mathrm{Potencia\ (MW)}\times 1000\times"
                 r"\mathrm{Precio\ Nudo\ (CLP/kW/mes)}")
        st.warning("**Interpretación del signo:** el CEN registra los retiros y el Pago PSUF con "
                   "signo negativo. Se presenta su magnitud positiva como referencia económica. "
                   "Esto NO demuestra que el importe sea idéntico al cargo contratado con Enel "
                   "ni incluye eventuales ajustes o reliquidaciones.")
        for aviso in resultado["avisos"]:
            st.warning(aviso)
        with st.expander("Detalle auditable: valores originales y conciliación de la fórmula"):
            st.dataframe([{"Punto":r["suministro"],"Descripción original CEN":r["descripcion_cen"],
                           "Barra":r["barra_cen"],"Tipo":r["tipo_cen"],
                           "Medida con signo (MW)":r["potencia_mw_con_signo"],
                           "Precio nudo ($/kW/mes)":r["precio_nudo_clp_kw_mes"],
                           "Pago PSUF con signo ($)":r["pago_psuf_con_signo_clp"],
                           "Producto calculado ($)":r["recalculado_con_signo_clp"],
                           "Diferencia ($)":r["diferencia_calculo_clp"],
                           "Hoja":r["hoja"],"Fila":r["fila"]}
                          for r in resultado["detalle"].values()],hide_index=True,use_container_width=True)
        st.download_button("⬇ Descargar potencia Quilpué y Limache (CSV)",
            data=csv_bytes([{"mes_balance":resultado["periodo"],
                  "mes_factura":resultado["mes_factura"],
                  "archivo_cen":resultado["archivo"],
                  "suministro":s,
                  "descripcion_cen":r["descripcion_cen"] if r else "NO IDENTIFICADO",
                  "barra":r["barra_cen"] if r else "",
                  "tipo":r["tipo_cen"] if r else "",
                  "potencia_retiro_mw":r["potencia_retiro_mw"] if r else None,
                  "precio_nudo_clp_kw_mes":r["precio_nudo_clp_kw_mes"] if r else None,
                  "pago_psuf_original_con_signo_clp":r["pago_psuf_con_signo_clp"] if r else None,
                  "monto_psuf_magnitud_clp":r["monto_retiro_psuf_clp"] if r else None,
                  "estado":"Registro CEN individual" if r else "No identificado", 
                  "hoja":r["hoja"] if r else "", "fila":r["fila"] if r else ""}
                  for s in SITE_NAMES for r in [resultado["detalle"].get(s)]]),
            file_name=f"Potencia_PSUF_EFE_{resultado['periodo']}.csv",
            mime="text/csv", key="potencia_v19_export")
        st.caption("Para contrastar el cargo de potencia de la factura se debe revisar el "
                   "contrato de suministro, los peajes y otros componentes que facture Enel.")


with exenciaci_tab:
    st.subheader("Pago por Exenciones a Clientes Individualizados (CI)")
    st.caption("Fuente: PagoExenCI del CEN desde Datos/Exenciones_CI/AAAA/ "
               "o cargado en Datos. Solo se muestran los años y meses disponibles.")
    st.info("**Criterio de asignación EFE:** este cargo corresponde exclusivamente "
            "a **Quilpué (100 %)**. **Limache no paga (0 %)**. "
            "No se reparte por MWh retirados ni se requiere el Balance de Energía "
            "para asignar el pago.")
    st.caption("El archivo PagoExenCI informa un pago agregado por cliente y suministrador. "
               "La clasificación Quilpué/Limache se basa en el criterio de EFE, "
               "no en un desglose por suministro publicado por el CEN. "
               "Este concepto se presenta por separado de PE, SSCC/BESS, MT/PD y PSUF.")

    if not exenciones_ci_reports:
        st.info("No hay archivos PagoExenCI en los períodos seleccionados. "
                "Cárgalos en Datos → Exenciones CI o guárdalos en "
                "Datos/Exenciones_CI/AAAA/ de GitHub.")
    else:
        historico = []
        for rep in exenciones_ci_reports:
            asignacion = asignar_exenciones_ci_suministros(rep)
            historico.append({
                "Mes balance": readable_month(rep["periodo"]),
                "Período": rep["periodo"],
                "Mes factura": readable_month(following_month(rep["periodo"])),
                "Quilpué ($)": round(asignacion["Quilpué"], 2),
                "Limache ($)": round(asignacion["Limache"], 2),
                "Total EFE ($)": round(asignacion["Total EFE"], 2),
                "Archivo CEN": rep["archivo"],
                "Criterio": "Quilpué 100 % · Limache 0 % (EFE)",
            })

        st.markdown("#### Resumen de períodos seleccionados")
        t_quilpue = sum(r["Quilpué ($)"] for r in historico)
        t_limache = sum(r["Limache ($)"] for r in historico)
        t_efe = sum(r["Total EFE ($)"] for r in historico)
        total_col1, total_col2, total_col3 = st.columns(3)
        total_col1.metric("Quilpué · Exenciones CI", clp(t_quilpue))
        total_col2.metric("Limache · Exenciones CI", clp(t_limache))
        total_col3.metric("Total EFE", clp(t_efe))
        st.dataframe(historico, hide_index=True, use_container_width=True)

        indice = st.selectbox(
            "Mes a consultar", range(len(exenciones_ci_reports)),
            format_func=lambda i: f"{readable_month(exenciones_ci_reports[i]['periodo'])} — {exenciones_ci_reports[i]['archivo']}",
            index=len(exenciones_ci_reports)-1, key="exenciaci_mes")
        rep = exenciones_ci_reports[indice]
        asignacion = asignar_exenciones_ci_suministros(rep)
        st.markdown("#### Detalle del mes seleccionado")
        col_q, col_l, col_total = st.columns(3)
        col_q.metric("Quilpué · 100 %", clp(asignacion["Quilpué"]))
        col_l.metric("Limache · 0 %", clp(asignacion["Limache"]))
        col_total.metric("Total EFE", clp(asignacion["Total EFE"]))
        col_p, col_s = st.columns(2)
        col_p.metric("Prorrata individual CEN (%)",
            number_es(rep["prorrata_individual"]*100, 6) if rep["prorrata_individual"] is not None else "No informada")
        col_s.metric("Participación de suministro Enel (%)",
            number_es(rep["prorrata_suministro"]*100, 2) if rep["prorrata_suministro"] is not None else "No informada")
        st.caption(f"Archivo de origen: {rep['archivo']} · Hoja: '{rep['hoja']}' · "
                   f"fila {rep['fila']} · importe total: columna AW. "
                   "El monto CEN no se multiplica por el 100 % de suministro de Enel: "
                   "ese porcentaje solo identifica su participación en el cliente.")

        with st.expander("Detalle por empresa de transmisión y trazabilidad del CEN"):
            st.dataframe([
                {"Empresa": r["empresa"], "Importe ($)": round(r["importe_clp"], 2),
                 "Origen": f"{rep['hoja']}!{r['columna']}{rep['fila']}"}
                for r in rep["detalle"]], hide_index=True, use_container_width=True)
            st.caption("Los importes negativos se mantienen con su signo. "
                       "La suma por empresa corresponde al total CEN atribuido a Quilpué "
                       "según el criterio de EFE; Limache permanece en $0.")

        st.download_button("⬇ Descargar Exenciones CI por suministro (CSV)",
                           data=csv_bytes(historico), file_name="EFE_Exenciones_CI_Quilpue.csv",
                           mime="text/csv", key="exenciaci_export")

with mt_pd_tab:
    st.caption("Fuente: Datos · Años y meses filtrados globalmente. Cargos referenciales y asignaciones oficiales se distinguen.")
    st.subheader("Mínimos Técnicos y Partidas y Detenciones · Solo datos del CEN")
    st.success("**No necesitas facturas ni proformas de Enel para este módulo.** "
               "El Balance CEN se utiliza para estimar los sobrecostos y una planilla CEN "
               "por retiro, si está disponible, para leer las asignaciones individuales.")
    st.warning("**No confundir:** `Sobrecostos Energía` (P) no es idéntico a `Mínimos Técnicos`. "
               "Las columnas O/P del Cuadro de Energía son saldos globales de Enel con signo, "
               "no montos cobrables directamente a EFE. El prorrateo por retiros mensuales "
               "es un escenario referencial; requiere el detalle de asignación para un cargo verificable.")
    if balances:
        st.markdown("#### Referencias mensuales MT/PD (no son cargos oficiales)")
        mtpd_comparativo = []
        for b in balances:
            p = estimar_sobrecostos_cen(b)
            for x in p["detalle"]:
                mtpd_comparativo.append({"Balance":readable_month(b.label),
                     "Suministro":x["punto"], "Sobrecosto PD referencial ($)":round(x["ref_pd_clp"]),
                     "Sobrecostos Energía referenciales ($)":round(x["ref_sobrecostos_energia_clp"])})
        st.dataframe(mtpd_comparativo,hide_index=True,use_container_width=True)
    selected = None
    if balances:
        idx=st.selectbox("Selecciona balance de energía del CEN", range(len(balances)),
            index=len(balances)-1,
            format_func=lambda i: f"{readable_month(balances[i].label)} · Balance CEN",
            key="mtpd_v18_balance")
        selected=balances[idx]
        st.info(f"**Balance:** {readable_month(selected.label)} · "
                f"**Factura de referencia:** {readable_month(following_month(selected.label))}")
        st.markdown("#### 1. Importes publicados por el CEN a nivel de Enel")
        costs=selected.sobrecostos_enel_cen
        st.dataframe([
            {"Concepto CEN":"Sobrecosto PD", "Hoja / celda":f"Cuadro de Energía!O{costs['fila']}",
             "Saldo Enel con signo ($)":round(costs['pd_neto_clp'],2)},
            {"Concepto CEN":"Sobrecostos Energía (no necesariamente MT)",
             "Hoja / celda":f"Cuadro de Energía!P{costs['fila']}",
             "Saldo Enel con signo ($)":round(costs['sobrecostos_energia_netos_clp'],2)},
        ],hide_index=True,use_container_width=True)
        st.markdown("#### 2. Proyección aproximada con retiros del Balance CEN")
        projection=estimar_sobrecostos_cen(selected)
        st.caption("Se utiliza la **magnitud absoluta** de cada saldo del CEN como referencia de tamaño "
                   "y se prorratea con la energía mensual retirada por EFE sobre la energía de Enel. "
                   "Se conserva el signo original arriba. Esta magnitud **no es** una asignación "
                   "oficial ni un cargo presupuestario validado.")
        st.dataframe([{"Suministro":rec['punto'],"Retiros (kWh)":round(rec['retiro_kwh'],2),
                       "Participación de Enel (%)":round(rec['participacion_pct'],6),
                       "Referencia de Sobrecosto PD ($)":round(rec['ref_pd_clp']),
                       "Referencia de Sobrecostos Energía ($)":round(rec['ref_sobrecostos_energia_clp'])}
                      for rec in projection['detalle']],hide_index=True,use_container_width=True)
        st.caption(f"**Retiros Enel:** {number_es(selected.retiro_enel_mwh,3)} MWh. "
                   "Esta tabla no entrega un valor certificado de Mínimos Técnicos.")
        st.download_button("⬇ Exportar estimación referencial (CSV)",
            data=csv_bytes([{
                "mes_balance":projection['periodo'], "mes_factura":projection['mes_factura'],
                "suministro":r['punto'],"clave":r['clave'],"retiro_kwh":r['retiro_kwh'],
                "participacion_enel_pct":r['participacion_pct'],
                "magnitud_referencial_pd_clp":round(r['ref_pd_clp'],2),
                "magnitud_referencial_sobrecostos_energia_clp":round(r['ref_sobrecostos_energia_clp'],2),
                "estado":"NO ES CARGO OFICIAL NI MT EXACTO",
                "fuente_pd":f"Cuadro de Energía!O{costs['fila']}",
                "fuente_energia":f"Cuadro de Energía!P{costs['fila']}"}
                for r in projection['detalle']]),
            file_name=f"CEN_sobrecostos_referencial_{projection['periodo']}.csv",mime="text/csv",
            key="mtpd_v18_prorrateo")
    else:
        st.info("Carga un Balance CEN para leer los sobrecostos de "
                "Enel y los retiros de Quilpué y Limache.")
    st.divider()
    st.markdown("#### 3. Asignación oficial individual desde otro archivo CEN")
    st.caption("Si cargas una planilla del CEN con columnas **clave / tipo_pago / pago**, el "
               "programa buscará exclusivamente MERVAL_ELSOL y MERVAL_LIMACHE. "
               "No incluirá SSCC o BESS como si fueran MT/PD.")
    if mt_pd_directos:
        choices = list(range(len(mt_pd_directos)))
        if selected:
            matches = [i for i in choices if mt_pd_directos[i]['periodo']==selected.label]
            starting = matches[-1] if matches else len(choices)-1
        else:
            starting = len(choices)-1
        which=st.selectbox("Detalle CEN de MT/PD",choices,index=starting,
            format_func=lambda i:mt_pd_directos[i]['archivo'],key="mtpd_v18_detail")
        detail=mt_pd_directos[which]
        st.info(f"**Mes de balance:** {readable_month(detail['periodo'])} · "
                f"**Factura:** {readable_month(following_month(detail['periodo']))}")
        rows=[]
        for s in SITE_NAMES:
            vals=detail['resumen'][s]
            rows.append({"Punto":s,"MT explícito ($)":round(vals['MT'],2) if vals['MT'] is not None else None,
                         "PD explícito ($)":round(vals['PD'],2) if vals['PD'] is not None else None,
                         "Otros sobrecostos energía ($)":round(vals['SOBRECOSTOS_ENERGIA'],2) if vals['SOBRECOSTOS_ENERGIA'] is not None else None})
        st.dataframe(rows,hide_index=True,use_container_width=True)
        if not all(detail['resumen'][s]['MT'] is not None and detail['resumen'][s]['PD'] is not None for s in SITE_NAMES):
            st.warning("La planilla no identifica ambos conceptos MT y PD para los dos suministros. "
                       "No se completan celdas faltantes con saldos globales ni se convierte automáticamente "
                       "Sobrecostos Energía en MT.")
        with st.expander("Origen auditable del archivo CEN"):
            st.dataframe(detail['detalle'],hide_index=True,use_container_width=True)
        st.download_button("⬇ Exportar asignación individual del CEN (CSV)",
            data=csv_bytes([{**r,"mes_balance":detail['periodo'],
                             "mes_factura":following_month(detail['periodo'])} for r in detail['detalle']]),
            file_name=f"CEN_MT_PD_detalle_{detail['periodo']}.csv",mime="text/csv",
            key="mtpd_v18_cen_export")
    else:
        st.info("Con los archivos actuales aún falta el **detalle CEN de pagos laterales por retiro** "
                "que identifique MT y PD. El Balance permite estimar magnitudes, pero no "
                "reconstruir con certeza el monto exacto facturable a EFE.")


# Proyección operativa de Datos-Operacionales (pestaña independiente)
from proyeccion_energia import mostrar_proyeccion
with proyeccion_operativa_tab:
    mostrar_proyeccion()


# Dólar observado: filtros propios de año y mes, independientes del CEN.
# El dólar no depende de los Excel cargados en Datos ni modifica filtros globales.
with dolar_observado_tab:
    try:
        from dolar_observado import mostrar_dolar_observado
        mostrar_dolar_observado()
    except Exception as exc:
        import logging as _log_dolar
        _log_dolar.exception("EFE_DOLAR_OBSERVADO: error al cargar módulo")
        st.error("No se puede abrir Dólar observado; las demás pestañas continúan funcionando.")
        st.caption("Tipo de error: " + type(exc).__name__ +
                   ". Consulta Streamlit Cloud → Manage app → Logs para ver el detalle.")


# Costo de energía activa: SEAT/SAF LI + reajuste CPI + dólar observado.
from costo_energia import mostrar_costo_energia
with costo_energia_tab:
    mostrar_costo_energia()


# Seguimiento de publicaciones oficiales de Transmisión Nacional/Zonal, SSCC y
# Servicio Público. No modifica los importes validados de la calculadora.
with monitoreo_cen_tab:
    try:
        from monitoreo_cen import mostrar_monitoreo_cen
        mostrar_monitoreo_cen(archivos_repositorio)
    except Exception as exc:
        import logging as _log_cen
        _log_cen.exception("EFE_ACTUALIZACIONES_CEN: error de revisión de fuentes oficiales")
        st.error("La consulta CEN no se pudo realizar; las demás pestañas siguen disponibles.")
        st.caption("Error: " + type(exc).__name__ + ". Revisa Streamlit Cloud → Logs.")


# Histórico tarifario verificado: REX CNE 2023–2026; tarifas autónomas
# respecto del inventario de Excel CEN. Los escenarios condicionados no se
# facturan automáticamente.
with tarifas_cne_tab:
    try:
        from tarifas_cne import mostrar_tarifas_cne
        mostrar_tarifas_cne(balances=balances)
    except Exception as exc:
        import logging as _log_tarifas
        _log_tarifas.exception("EFE_TARIFAS_CNE: error en módulo de histórico regulado")
        st.error("No se puede cargar Tarifas CNE; las demás pestañas siguen disponibles.")
        st.caption("Error: " + type(exc).__name__ + ". Revisa Streamlit Cloud → Logs.")


# Resumen de FACTURA: PE, SSCC 4/9, MT/PD y liquidación de peaje
# (PagoExenCI y proforma) son del MES ANTERIOR a la factura, incluso
# cuando los filtros globales muestran otro mes.
# La carga es perezosa: solo el mes consultado. El resultado se mantiene en
# memoria durante la sesión y no se reabre el Excel grande por cada cambio de filtro.
# Memoria por sesión/ejecución, no serializa Excel ni reabre el Archivo 9.
# Una nueva ejecución completa recrea este caché y recoge los cambios en GitHub.
# Al cambiar año/mes/suministro DENTRO del fragmento, se reutilizan los meses ya leídos.
@lru_cache(maxsize=24)
def cargar_reliquidaciones_resumen(mes_balance: str) -> dict:
    resultado = {'balances': [], 'sscc_reports': [], 'sscc_companies': [],
                 'mt_pd_directos': [], 'potencia_reports': [],
                 'exenciones_ci_reports': [], 'proformas_enel': [], 'avisos': []}
    tareas = (
        ('balance', 'balances', lambda archivo: calc_balance(
            archivo.getvalue(), archivo.name, tuple(keys))),
        ('sscc9', 'sscc_reports', lambda archivo: calc_sscc(
            archivo.getvalue(), archivo.name)),
        ('sscc4', 'sscc_companies', lambda archivo: calc_sscc_company(
            archivo.getvalue(), archivo.name)),
        ('mtpd', 'mt_pd_directos', lambda archivo: calc_mt_pd_cen_retiros(
            archivo.getvalue(), archivo.name)),
        ('potencia', 'potencia_reports', lambda archivo: calc_potencia(
            archivo.getvalue(), archivo.name)),
        ('exenciaci', 'exenciones_ci_reports', lambda archivo: calc_exenciones_ci(
            archivo.getvalue(), archivo.name)),
    )
    for categoria, destino, parser in tareas:
        disponibles=[archivo for archivo in archivos_activos.get(categoria, [])
                     if month_from_filename(archivo.name)==mes_balance]
        elegidos, avisos = un_archivo_por_mes(disponibles, categoria)
        resultado['avisos'].extend(avisos)
        for archivo in elegidos:
            try:
                dato=parser(archivo)
                # Nunca mezclar archivos etiquetados con otra fecha.
                identificador=(dato.label if categoria=='balance' else dato.get('periodo'))
                if identificador!=mes_balance:
                    raise ValueError(f'Período incompatible: {identificador}')
                if categoria=='potencia' and dato.get('mes_factura')!=following_month(mes_balance):
                    raise ValueError('Mes de factura del Balance2 no coincide con el siguiente al balance')
                resultado[destino].append(dato)
            except Exception as exc:
                resultado['avisos'].append(
                    f'{archivo.name}: {type(exc).__name__}: {str(exc)[:130]}')

    # Proforma con mes de factura M+1: mostrar peaje individualizado de M
    # sólo como referencia, no sumarlo otra vez sobre la tarifa de transmisión.
    from proforma_enel import periodo_archivo, _parse_cached
    mes_factura = following_month(mes_balance)
    proformas = [a for a in archivos_activos.get('proforma', [])
                 if periodo_archivo(a.name) == mes_factura]
    seleccion, avisos = un_archivo_por_mes(proformas, 'proforma')
    resultado['avisos'].extend(avisos)
    for archivo in seleccion:
        try:
            leida = _parse_cached(archivo.getvalue(), archivo.name)
            if leida.get('periodo') != mes_factura:
                raise ValueError('Mes del Excel distinto de la factura '+mes_factura)
            resultado['proformas_enel'].append(leida)
        except Exception as exc:
            resultado['avisos'].append(
                f'Proforma {archivo.name}: {type(exc).__name__}: {str(exc)[:130]}')
    return resultado


# La proforma tiene sus propios filtros año/mes/suministro y no depende
# de que el balance CEN anterior esté seleccionado en Análisis energético.
with espacio_proforma.container():
    try:
        from proforma_enel import mostrar_proformas_enel
        mostrar_proformas_enel(archivos_activos.get('proforma', []))
    except Exception as exc:
        import logging as _log_proforma
        _log_proforma.exception('EFE_PROFORMA: fallo al construir pestaña')
        st.error('No se pudo construir Proforma; las demás pestañas siguen disponibles.')
        with st.expander('Diagnóstico de proforma'):
            st.exception(exc)

# Panel principal: información ejecutiva reunida después de todos los módulos.
with espacio_resumen.container():
    try:
        from panel_ejecutivo import mostrar_panel_ejecutivo
        mostrar_panel_ejecutivo(
            balances=balances, archivos_repositorio=archivos_repositorio,
            sscc_reports=sscc_reports, potencia_reports=potencia_reports,
            exenciones_ci_reports=exenciones_ci_reports, mt_pd_directos=mt_pd_directos,
            sscc_companies=sscc_companies,
            cargar_reliquidacion=cargar_reliquidaciones_resumen)
    except Exception as exc:
        import logging as _log_panel
        _log_panel.exception("EFE_PANEL_EJECUTIVO: fallo en resumen")
        st.error("No se pudo construir el resumen. Las demás secciones siguen disponibles.")
        with st.expander("Diagnóstico de la pantalla principal"):
            st.write(f"{type(exc).__name__}: {exc}")
