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
    archivo: str = ""
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


TIPOS_MEDIDOR_RETIRO = ("L", "L_D", "L_S")  # Balance Valorizado!J; NO utilizar R.


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
    aggregate = {medidor: 0.0 for medidor in TIPOS_MEDIDOR_RETIRO}
    enel_retiros = {medidor: 0.0 for medidor in TIPOS_MEDIDOR_RETIRO}
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
            if kind not in TIPOS_MEDIDOR_RETIRO:
                continue
            # EFE bajo contrato Enel: igual criterio de suministrador que el
            # denominador de Enel, leído de Balance Valorizado columna L.
            if normalize(d.get("L")) != "ENEL_GENERACION":
                continue
            found.append({
                "clave": str(d.get("A", "")), "descripcion": str(d.get("H", "")),
                "barra": f"{d.get('B', '')} {d.get('C', '')} kV",
                "suministrador": str(d.get("F", "")),
                "tipo": kind, "fila": idx, "kwh": -kwh,
                "origen": "Balance Valorizado",
            })
    total_kwh = sum(aggregate.values())
    if total_kwh <= 0:
        raise ValueError("No se obtuvieron retiros netos positivos L/L_D/L_S. Revisar el archivo.")
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
        metodo=f"Cuadro de Energía!M{totalrow} (valor absoluto) / retiros netos J=L, L_D, L_S de Balance Valorizado",
        archivo=filename,
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
        if normalize(entry.get('tipo')) not in TIPOS_MEDIDOR_RETIRO:
            continue
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


