"""Aplicación web Compensación PE EFE Valparaíso, versión Cloud autocontenida 2.6: autolectura de GitHub Datos/, cargas manuales y filtros globales por mes/año.

Motor de cálculo y pantalla integrados. No requiere importar cen_core.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from io import BytesIO, StringIO
from pathlib import Path
from functools import lru_cache
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
        value = num(row.get('J'))
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
    reader = WorkbookReader(data)
    mon = 'Sin período identificado'
    amounts = {}
    for idx, row in reader.sheet('Cargos_CEN'):
        if idx == 2 and isinstance(row.get('D'), (int, float)):
            mon = excel_date(row['D']).strftime('%Y-%m')
            break
    for sheetname, site in (('PROFORMA-Q','Quilpué'), ('PROFORMA-L', 'Limache')):
        values = [num(r.get('F')) for _,r in reader.sheet(sheetname)
                  if normalize(r.get('D')) == 'SERVICIOS COMPLEMENTARIOS']
        if len(values) != 1:
            raise ValueError('No hay una única línea de Servicios Complementarios en '+sheetname)
        amounts[site] = values[0]
    return {'periodo_balance':mon,'mes_factura':following_month(mon),
            'importes':amounts,'archivo':filename}

def calcular_prorrateo_sscc(empresa: dict, balance: BalanceResult,
                            bess_unitario_clp_kwh: float = 0.0) -> dict:
    """Estimación mensual: C46 archivo 4 x retiro EFE / retiros Enel del balance.

    BESS se considera ÚNICAMENTE como supuesto dentro de SSCC, nunca se suma a
    la línea oficial del archivo 9. El valor cero significa proyección parcial.
    """
    if empresa["periodo"] != balance.label:
        raise ValueError("El archivo 4 y el balance deben pertenecer al mismo mes.")
    if not isinstance(bess_unitario_clp_kwh, (float, int)) or bess_unitario_clp_kwh < 0:
        raise ValueError("La tasa de referencia BESS no puede ser negativa.")
    retiro_enel_kwh = balance.retiro_enel_mwh * 1000.0
    if retiro_enel_kwh <= 0:
        raise ValueError("Los retiros totales de Enel son cero.")
    paga_enel = float(empresa["SC_SSCC"]["paga"])
    energies = balance_consumption_by_site(balance)
    details = []
    for site in SITE_NAMES:
        retiro_kwh = energies[site]
        participacion = retiro_kwh / retiro_enel_kwh
        base = paga_enel * participacion
        bess_estimado = retiro_kwh * bess_unitario_clp_kwh
        details.append({
            "punto": site,
            "clave": SSCC_KEYS[site],
            "retiro_kwh": retiro_kwh,
            "participacion_enel_pct": participacion * 100.0,
            "sscc_base_clp": base,
            "bess_estimado_clp": bess_estimado,
            "sscc_estimado_clp": base + bess_estimado,
        })
    return {
        "periodo": empresa["periodo"],
        "mes_factura": following_month(empresa["periodo"]),
        "archivo_4": empresa["archivo"],
        "pago_enel_sc_sscc_clp": paga_enel,
        "retiro_enel_mwh": balance.retiro_enel_mwh,
        "retiros_enel_por_tipo": balance.retiros_enel_por_tipo,
        "bess_unitario_clp_kwh": bess_unitario_clp_kwh,
        "bess_incluido_como_sup": bool(bess_unitario_clp_kwh),
        "detalle": details,
        "sscc_base_total": sum(x["sscc_base_clp"] for x in details),
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


import streamlit as st

# Versión 2.3: carpeta raíz Datos/ con D mayúscula.
# Datos/Balances/2026/, Datos/Sscc_archivo4/2026/, etc.
# Se admite temporalmente datos/ (d minúscula) para migraciones.
# Archivos >25 MiB: original.xlsx.part001, part002... en la MISMA carpeta.
CARPETAS_CEN = {
    "balance": "Balances",
    "precio": "Precios_estabilizados",
    "sscc4": "Sscc_archivo4",
    "sscc9": "Sscc_archivo9",
    "mtpd": "Mt_pd",
    "potencia": "Potencia",
    "horario": "Horarios",
    "factura": "Facturas_enel",
}
# El reconocimiento permite las dos formas de la carpeta raíz y nombres internos de versiones anteriores.
CARPETAS_A_CATEGORIAS = {v.casefold(): k for k, v in CARPETAS_CEN.items()}

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
    if n.startswith("EFE_VALPO_"):
        return "factura"
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
            registros.append((carpeta,nombre,tuple(str(p) for _,p in pedazos)))

    candidatos: dict[tuple[str,int,str],ArchivoRepositorio] = {}
    prioridades: dict[tuple[str,int,str],int] = {}
    for carpeta, nombre, rutas in registros:
        niveles = carpeta.parts
        if not niveles:   # Compatibilidad con la versión 2.0, con carpetas raíz planas.
            cat = clasificar_archivo_antiguo(nombre)
            anno = extraer_anio_nombre(nombre)
            if cat is None:
                avisos.append(f"{nombre}: no se reconoce categoría en {raiz.name}/; mueve a la carpeta apropiada.")
                continue
            prioridad = 0
        else:
            cat = CARPETAS_A_CATEGORIAS.get(niveles[0].casefold())
            if cat is None:
                avisos.append(f"{carpeta}/{nombre}: carpeta no reconocida; consulta README.")
                continue
            if len(niveles)>2 or (len(niveles)==2 and not re.fullmatch(r"20\d{2}", niveles[1])):
                avisos.append(f"{carpeta}/{nombre}: usa {raiz.name}/{CARPETAS_CEN[cat]}/AAAA/nombre.xlsx.")
                continue
            anno = int(niveles[1]) if len(niveles)==2 else extraer_anio_nombre(nombre)
            prioridad = 2 if niveles[0] == CARPETAS_CEN[cat] else 1
            detectado = extraer_anio_nombre(nombre)
            if anno and detectado and anno != detectado:
                avisos.append(f"{carpeta}/{nombre}: el año {anno} de la carpeta no coincide con {detectado} del archivo; se omitió.")
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
    """Lee Datos/ y, durante la transición, datos/; Datos/ tiene prioridad.

    Mantiene registros históricos de ambos directorios si no están duplicados.
    """
    raiz_proyecto = Path(__file__).resolve().parent
    por_archivo: dict[tuple[str, int, str], ArchivoRepositorio] = {}
    avisos: list[str] = []
    for nombre_raiz in ("datos", "Datos"):
        encontrados, avisos_en_raiz = _descubrir_archivos_en_raiz(raiz_proyecto / nombre_raiz)
        avisos.extend(avisos_en_raiz)
        for categoria, archivos in encontrados.items():
            for archivo in archivos:
                clave = (categoria, archivo.anio or 0, archivo.name.casefold())
                # Iterar Datos/ al final hace que prevalezca en duplicados.
                por_archivo[clave] = archivo
    resultado: dict[str, list[ArchivoRepositorio]] = {c: [] for c in CARPETAS_CEN}
    for archivo in por_archivo.values():
        resultado[archivo.categoria].append(archivo)
    for lista in resultado.values():
        lista.sort(key=lambda a: (a.anio or 0, month_from_filename(a.name), a.name.casefold()))
    return resultado, avisos


def fusionar_archivos(locales, subidos):
    """Conserva el histórico: un archivo por categoría, nombre y año.

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
archivos_repositorio, advertencias_repositorio = descubrir_archivos_repositorio()

st.set_page_config(page_title="Análisis CEN | EFE Valparaíso", page_icon="⚡", layout="wide")
st.markdown("""<style>
.block-container{max-width:1300px;padding-top:1.3rem}
h1{letter-spacing:-.035em}
.stMetric{border:1px solid #dae2ed;border-radius:12px;padding:12px}
</style>""", unsafe_allow_html=True)
st.title("⚡ Cargos CEN y análisis energético")
st.caption("EFE Valparaíso · Quilpué y Limache · Versión 2.6 · GitHub automático + cargas manuales")

# Crear las pestañas antes de leer/transformar los Excel permite alojar todos
# los cargadores en Datos sin romper el resto del motor ya implementado.
datos_tab, energia_tab, calc_tab, history_tab, hourly_tab, sscc_tab, mt_pd_tab, potencia_tab, method_tab = st.tabs([
    "📁 Datos", "🔋 Análisis energético", "📊 Precio Estabilizado", "📈 Histórico PE",
    "🕒 Cálculo horario PE", "🔧 Servicios Complementarios", "⚙️ Mínimos Técnicos y PD",
    "⚡ Potencia PSUF", "📖 Metodología"])

with datos_tab:
    st.subheader("1 · Cargar información del CEN")
    st.write("**Sube los archivos aquí una sola vez**. Los módulos trabajan con esta misma selección; "
             "no se usan los Excel enviados anteriormente al chat ni resultados numéricos precargados.")
    # v2.6: el contenido de Datos/ en el repositorio se incluye SIEMPRE.
    # En la v2.5 el checkbox estaba desactivado por defecto, por lo que los
    # libros correctamente guardados en GitHub no aparecían en el inventario.
    usar_repositorio = True
    total_repo = sum(len(items) for items in archivos_repositorio.values())
    if total_repo:
        st.success(f"GitHub conectado mediante archivos desplegados: {total_repo} archivo(s) CEN detectado(s) en Datos/. "
                   "Ya están disponibles para todas las pestañas y sus filtros.")
    else:
        st.warning("No encontré archivos en Datos/ dentro del despliegue de Streamlit. "
                   "Comprueba que los Excel estén confirmados (commit) en la misma rama que app.py, "
                   "dentro de Datos/Balances/AAAA, Datos/Sscc_archivo4/AAAA, "
                   "Datos/Sscc_archivo9/AAAA o Datos/Potencia/AAAA. "
                   "Si acabas de subirlos, reinicia o redespliega la aplicación.")
    st.caption("Los archivos guardados en GitHub se cargan automáticamente al abrir la aplicación; "
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
        mtpd_cen_uploads = st.file_uploader("MT y PD · Asignaciones por retiro (si existen)",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_mtpd")
        horarios_uploads = st.file_uploader("Compensaciones horarias · CSV",
            type=["csv"], accept_multiple_files=True, key="datos_horarios")
        invoice_uploads = st.file_uploader("Factura Enel (solo conciliación, opcional)",
            type=["xlsx", "xlsm"], accept_multiple_files=True, key="datos_factura")
    with st.expander("Configuración de identificadores de EFE", expanded=False):
        clave_text = st.text_area("Claves de retiro en Balance Valorizado (una por línea)",
            "MERVAL_ELSOL\nMERVAL_LIMACHE", height=88, key="datos_claves")
    keys = [k.strip() for k in clave_text.replace(",", "\n").splitlines() if k.strip()]
    st.info("**Seguridad:** los archivos se procesan en el servidor donde está alojado Streamlit. "
            "No se ejecutan macros. Evita guardar datos comerciales en repositorios públicos.")

subidas = {"balance":balance_uploads or [], "precio":price_uploads or [],
           "sscc4":sscc_empresa_uploads or [], "sscc9":sscc_uploads or [],
           "potencia":potencia_uploads or [], "mtpd":mtpd_cen_uploads or [],
           "horario":horarios_uploads or [], "factura":invoice_uploads or []}
archivos_activos = {categoria: fusionar_archivos(
    archivos_repositorio.get(categoria, []) if usar_repositorio else [], subidas[categoria])
    for categoria in CARPETAS_CEN}

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
                               "Procedencia":"Datos/ del repositorio" if isinstance(archivo, ArchivoRepositorio)
                                             else "Subido en la aplicación"})
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

# Un único filtro global compartido por todos los módulos.
anios_disponibles = sorted({int(month_from_filename(a.name)[:4])
    for archivos in archivos_activos.values() for a in archivos
    if re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", month_from_filename(a.name))})
meses_disponibles = sorted({int(month_from_filename(a.name)[5:7])
    for archivos in archivos_activos.values() for a in archivos
    if re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", month_from_filename(a.name))})
with st.sidebar:
    st.header("📅 Filtros para todas las pestañas")
    st.caption("Selecciona uno, dos o todos los meses y años. El mismo filtro se aplica a PE, "
               "SSCC, MT/PD, potencia e histórico energético.")
    todos_anos = st.checkbox("Todos los años", value=True, key="f_todos_anos")
    if todos_anos:
        anos_seleccion = anios_disponibles
        st.caption("Años: " + (", ".join(map(str, anos_seleccion)) or "sin datos"))
    else:
        anos_seleccion = st.multiselect("Seleccionar años", anios_disponibles,
                                       default=anios_disponibles[-1:] if anios_disponibles else [],
                                       key="f_anos")
    todos_meses = st.checkbox("Todos los meses", value=True, key="f_todos_meses")
    if todos_meses:
        meses_seleccion = list(range(1, 13))
    else:
        meses_seleccion = st.multiselect("Seleccionar meses", list(range(1, 13)),
                                        default=meses_disponibles,
                                        format_func=lambda m: MONTHS_ES[m-1].capitalize(), key="f_meses")
    st.divider()
    st.caption(f"Archivos cargados: {sum(map(len, archivos_activos.values()))} · "
               f"Años elegidos: {len(anos_seleccion)} · Meses elegidos: {len(meses_seleccion)}")

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
    sin_periodo = [a.name for files in archivos_activos.values() for a in files
                   if month_from_filename(a.name) == "Sin período identificado"]
    if sin_periodo:
        st.warning("Archivos sin mes identificable (no se incluirán en los cálculos): "
                   + ", ".join(sin_periodo[:12])
                   + ". Utiliza nombres como Balance_2609P.xlsm o BDef Detalle Sep26_def_v2.xlsm.")

@st.cache_data(show_spinner=False)
def calc_potencia(data: bytes, name: str):
    return parse_potencia_cen(data, name)

@st.cache_data(show_spinner=False)
def calc_balance(data: bytes, name: str, keys_tuple: tuple[str, ...]):
    return parse_balance(data, name, list(keys_tuple))

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
def calc_sscc_bill(data: bytes, name: str):
    return parse_sscc_invoice(data, name)


balances, invoices, errors = [], [], []
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
for item in invoice_uploads or []:
    try:
        sscc_bills.append(calc_sscc_bill(item.getvalue(), item.name))
    except Exception:
        pass  # Proformas no compatibles se siguen leyendo normalmente en el módulo PE.
for error in errors:
    st.error(error)

def clp(number: float) -> str:
    return "$" + f"{number:,.0f}".replace(",", ".")

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
            avisos.append(f"No se puede identificar el mes de {archivo.name}; usa nombres como Balance_2609P.xlsm.")
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
    st.caption("Origen: archivos cargados en Datos. Se aplican los filtros globales de la izquierda. "
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
            graf = [{"Período":r["periodo"], "Quilpué":r["Quilpué (MWh)"],
                     "Limache":r["Limache (MWh)"]} for r in datos_energia]
            try:
                import pandas as pd
                tabla_graf = pd.DataFrame(graf).set_index("Período")
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
    st.subheader("Proyección de Compensación PE")
    st.caption("Fuente: pestaña Datos · Se aplican los años y meses elegidos en el filtro global.")
    if not balances and not invoices:
        st.info("No hay Balance CEN o factura de conciliación en los períodos elegidos. "
                "Carga un balance desde Datos y selecciónalo en los filtros globales.")
    else:
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
        "fecha":"2026-09-01", "hora":"00:00", "compensacion_sen_clp":"",
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

**Prorrateo SSCC desde archivo 4:** se toma el importe bruto `PAGA` de Enel
Generación en `PAGO_EMPRESAS_SC_SSCC` (no el saldo `Paga+Recibe`). Del `Balance
Valorizado` se suman los retiros negativos `fisico_kwh`, tipos `L`, `L_D` y `R`,
con `nombre_corto_empresa = ENEL_GENERACION`. Se calcula la participación mensual
de Quilpué y Limache, cada uno por separado. Este procedimiento es una **estimación**
y no reproduce necesariamente los criterios de asignación horaria del CEN.

**BESS dentro de SSCC:** el archivo 4 no suministra la asignación individual BESS.
Se ofrece tasa BESS histórica o editable como **supuesto** dentro del único SSCC
estimado. Si queda en cero, el resultado es **parcial: solo SC_SSCC**. El archivo 9
sirve para comprobar la asignación individual oficial `SC_SSCC + BESS`, sin doble suma.

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
        empresa_sscc = sscc_company['SC_SSCC']
        st.markdown('#### Paso 1 · Identificar el pago general de Enel')
        e1,e2=st.columns(2)
        e1.metric('Enel PAGA al mecanismo SSCC', clp(empresa_sscc['paga']),
                  help='Archivo 4, hoja PAGO_EMPRESAS_SC_SSCC, columna C de la fila de Enel Generación. Es el importe utilizado para el prorrateo referencial.')
        e2.metric('Enel RECIBE por SSCC',clp(empresa_sscc['recibe']),
                  help='Archivo 4, hoja PAGO_EMPRESAS_SC_SSCC, columna B; NO es un descuento de la factura de EFE.')
        st.caption('Ambos importes pertenecen a Enel a nivel de empresa. '
                   'El importe que PAGA se usa para la simulación; el importe que RECIBE '
                   'NO se descuenta del cargo de EFE.')
        with st.expander('Ver respaldo y saldos del archivo 4 (información técnica)'):
            st.dataframe([
                {'Dato':'Enel recibe por SSCC', 'Hoja':'PAGO_EMPRESAS_SC_SSCC',
                 'Celda':f"B{empresa_sscc['fila']}", 'Monto ($)':round(empresa_sscc['recibe'],2),
                 'En el cálculo EFE':'No'},
                {'Dato':'Enel paga por SSCC','Hoja':'PAGO_EMPRESAS_SC_SSCC',
                 'Celda':f"C{empresa_sscc['fila']}",'Monto ($)':round(empresa_sscc['paga'],2),
                 'En el cálculo EFE':'Sí; base de simulación'},
                {'Dato':'Saldo neto Enel','Hoja':'PAGO_EMPRESAS_SC_SSCC',
                 'Celda':f"D{empresa_sscc['fila']}",'Monto ($)':round(empresa_sscc['neto'],2),
                 'En el cálculo EFE':'No'},
            ],hide_index=True,use_container_width=True)
            st.write('**Saldo neto:** resume pagos y remuneraciones de Enel; '
                     'no representa un descuento, un abono ni un cargo individual de EFE.')
            empresa_co=sscc_company['CO']
            st.write('**CO (costos de oportunidad):** el archivo 4 también contiene esta hoja, '
                     'pero sus montos **no se suman automáticamente** al SSCC calculado aquí.')
            st.dataframe([{'Dato':'Enel recibe por CO','Monto ($)':round(empresa_co['recibe'],2)},
                          {'Dato':'Enel paga por CO','Monto ($)':round(empresa_co['paga'],2)},
                          {'Dato':'Saldo neto Enel en CO','Monto ($)':round(empresa_co['neto'],2)}],
                         hide_index=True,use_container_width=True)
        compatible_balances = [b for b in balances if b.label == periodo_empresa]
        if not compatible_balances:
            st.warning('**Falta el Balance CEN del mismo mes.** Para repartir el pago general de Enel '
                       f'debes cargar un Balance de {readable_month(periodo_empresa)} '
                       '(por ejemplo, `Balance_2609P.xlsm`).')
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
                with st.expander('Opcional · Mejorar la estimación con una referencia histórica de BESS'):
                    st.write('**BESS ya es parte de SSCC**, no se mostrará como un segundo cargo. '
                             'El archivo 4 no entrega su importe individual para EFE. '
                             'Puedes añadir una referencia histórica expresada en $/kWh.')
                    bess_unit=st.number_input(
                        'Referencia histórica de BESS incluida en SSCC ($/kWh)',
                        min_value=0.0,value=round(float(referencia_bess),6),
                        step=0.05,format='%.6f',key='bess_rate_sup')
                    if referencia_bess:
                        st.caption(fuente_bess)
                    st.caption('Con valor 0, el resultado se limita a SC_SSCC del archivo 4 '
                               'y no equivale al total completo de SSCC.')
                prorrateo=calcular_prorrateo_sscc(sscc_company,selected_balance,bess_unit)
                table_participacion=[]
                for rec in prorrateo['detalle']:
                    table_participacion.append({
                        'Suministro':rec['punto'],
                        'Retiro EFE (MWh)':number_es(rec['retiro_kwh']/1000,3),
                        'Parte de los retiros Enel (%)':number_es(rec['participacion_enel_pct'],6),
                        'Pago repartido ($)':clp(rec['sscc_base_clp'])})
                st.dataframe(table_participacion,hide_index=True,use_container_width=True)
                st.markdown('**Así se obtiene el monto estimado:**')
                st.code('Pago SSCC de Enel × (MWh del suministro ÷ MWh totales Enel)',language=None)
                st.markdown('#### Paso 3 · ¿Cuánto se estima para EFE?')
                if bess_unit:
                    st.caption('Estimación con referencia BESS incorporada dentro del único cargo SSCC.')
                else:
                    st.warning('**Estimación parcial:** solo incluye el prorrateo SC_SSCC del archivo 4. '
                               'No incluye una estimación de BESS. No la consideres como el total '
                               'SSCC definitivo de la factura.')
                k1,k2,k3=st.columns(3)
                k1.metric('Quilpué · estimado',clp(prorrateo['detalle'][0]['sscc_estimado_clp']))
                k2.metric('Limache · estimado',clp(prorrateo['detalle'][1]['sscc_estimado_clp']))
                k3.metric('Total EFE · estimado',clp(prorrateo['sscc_total_estimado']))
                st.caption('**No es el cobro oficial del CEN.** El prorrateo mensual puede diferir '
                           'del reparto real por servicios, horarios, zonas y otros criterios.')
                report_same_month = next((r for r in sscc_reports
                                          if r['periodo']==prorrateo['periodo']),None)
                if report_same_month:
                    st.markdown('#### Paso 4 · Comparación con el archivo 9 (asignación individual)')
                    st.success('Se encontró el archivo 9 del mismo mes. '
                               'Ese archivo contiene los importes asignados a los retiros de EFE.')
                    compare=[]
                    for rec in prorrateo['detalle']:
                        site=rec['punto']
                        official_base=report_same_month['resumen'][site]['SC_SSCC']
                        official_total=report_same_month['subtotal_base'][site]
                        compare.append({
                            'Suministro':site,
                            'Estimación SC_SSCC archivo 4 ($)':clp(rec['sscc_base_clp']),
                            'SC_SSCC archivo 9 ($)':clp(official_base),
                            'Diferencia componente comparable ($)':clp(rec['sscc_base_clp']-official_base),
                            'SSCC oficial archivo 9, con BESS ($)':clp(official_total)})
                    st.dataframe(compare,hide_index=True,use_container_width=True)
                    st.caption('La diferencia corresponde al componente SC_SSCC sin BESS en ambas '
                               'columnas comparables. La última columna muestra el **total oficial SSCC '
                               'con BESS** y no debe compararse directamente con un prorrateo parcial.')
                    o1,o2,o3=st.columns(3)
                    o1.metric('Quilpué · archivo 9',clp(report_same_month['subtotal_base']['Quilpué']))
                    o2.metric('Limache · archivo 9',clp(report_same_month['subtotal_base']['Limache']))
                    o3.metric('Total EFE · archivo 9',clp(sum(report_same_month['subtotal_base'].values())))
                    st.info('**Para verificar el monto publicado por el CEN, utiliza el archivo 9.** '
                            'El archivo 4 sirve para explicar el costo global de Enel y hacer '
                            'proyecciones referenciales cuando aún no esté el archivo 9.')
                export_rows=[{
                    'mes_balance':prorrateo['periodo'],
                    'mes_factura':prorrateo['mes_factura'],
                    'metodo':'Estimación por prorrateo mensual desde archivo 4 y Balance CEN',
                    'punto':rec['punto'],'clave':rec['clave'],
                    'pago_global_enel_sc_sscc_clp':prorrateo['pago_enel_sc_sscc_clp'],
                    'retiro_total_enel_mwh':prorrateo['retiro_enel_mwh'],
                    'retiro_punto_mwh':rec['retiro_kwh']/1000,
                    'participacion_enel_pct':rec['participacion_enel_pct'],
                    'sscc_base_estimado_clp':rec['sscc_base_clp'],
                    'referencia_bess_incluida_en_sscc_clp_kwh':bess_unit,
                    'sscc_estimado_clp':rec['sscc_estimado_clp'],
                    'caracter_resultado':'Estimación sin BESS' if bess_unit==0 else 'Estimación con supuesto histórico BESS',
                    'archivo_4':sscc_company['archivo'],
                    'archivo_balance':selected_balance.label,
                } for rec in prorrateo['detalle']]
                st.download_button('⬇ Descargar explicación y estimación (CSV)',
                    data=csv_bytes(export_rows),
                    file_name=f"estimacion_SSCC_{prorrateo['periodo']}_factura_{prorrateo['mes_factura']}.csv",
                    mime='text/csv',key='export_sscc_prorrateo')
            except Exception as exc:
                st.error(f'No se pudo calcular el prorrateo con el archivo 4: {exc}')
    else:
        st.info('Carga el **archivo 4** desde la pestaña Datos para ver cuánto paga y recibe Enel. '
                'Agrega un **Balance CEN del mismo mes** para estimar la parte de Quilpué y Limache. '
                'El **archivo 9** es opcional para validar los importes individuales.')
    st.divider()
    st.markdown('### 2. Asignación individual del CEN · Archivo 9')
    st.caption('El archivo 9_Pagos_Retiros se filtra por suministrador ENEL_GENERACION '
               'y claves MERVAL_ELSOL (Quilpué) / MERVAL_LIMACHE (Limache).')
    if not sscc_reports:
        st.info('Carga el archivo 9_Pagos_Retiros_2609_def.xlsx desde la pestaña Datos para ver los importes oficiales individualizados.')
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
        qbase = agrupado['Quilpué']['SSCC']
        lbase = agrupado['Limache']['SSCC']
        totalbase = qbase + lbase
        a,b,c=st.columns(3)
        a.metric('SSCC · Quilpué',clp(qbase))
        b.metric('SSCC · Limache',clp(lbase))
        c.metric('Total SSCC EFE (sin IVA)',clp(totalbase))
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
                company_table=[{'Concepto':c,'Empresa': 'ENEL_GENERACION',
                                'Recibe ($)':round(empresa_del_mes[c]['recibe'],2),
                                'Paga ($)':round(empresa_del_mes[c]['paga'],2),
                                'Saldo ($)':round(empresa_del_mes[c]['neto'],2)}
                               for c in ('SC_SSCC','CO')]
                st.dataframe(company_table,hide_index=True,use_container_width=True)
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
        st.info("Sube **BDef Detalle Sep26_def_v2.xlsm** en la pestaña Datos, sección "
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
        st.info("Carga `Balance_2609P.xlsm` u otro Balance CEN para leer los sobrecostos de "
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
