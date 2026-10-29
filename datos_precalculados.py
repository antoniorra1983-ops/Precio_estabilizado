"""Resultados CEN preprocesados y versionados en GitHub.

Uso SOLO de datos originados en Datos/. El cache no modifica ni inventa valores.
Los archivos de usuario subidos a Streamlit no utilizan resultados de GitHub.
"""
from __future__ import annotations
from functools import lru_cache
from pathlib import Path
from hashlib import sha256
import gzip
import json
import os
import re
import subprocess

ESQUEMA = 1
ROOT = Path(__file__).resolve().parent
CARPETAS = ROOT / "Datos_procesados"
CODIGO = ("cen_motor.py", "proforma_enel.py")


@lru_cache(maxsize=1)
def firma_motor() -> str:
    digest = sha256()
    digest.update(f"precálculo-{ESQUEMA}".encode())
    for archivo in CODIGO:
        digest.update((ROOT / archivo).read_bytes())
    return digest.hexdigest()


def ruta_cache(categoria: str, periodo: str) -> Path:
    if not re.fullmatch(r"[a-z0-9_]+", categoria):
        raise ValueError("Categoría incorrecta")
    if not re.fullmatch(r"20\d\d-(?:0[1-9]|1[0-2])", periodo):
        raise ValueError("Período incorrecto")
    return CARPETAS / periodo[:4] / f"{categoria}_{periodo}.json.gz"


@lru_cache(maxsize=180)
def _leer_cache(ruta: str, tam: int, fecha: int) -> dict:
    with gzip.open(ruta, "rt", encoding="utf-8") as f:
        return json.load(f)


@lru_cache(maxsize=1)
def _git_blobs() -> dict[str, str]:
    """Lee SHA de Git desde el índice (sin abrir archivos Excel grandes).

    Si no existe .git (despliegue sin historial), se valida por tamaño y se
    confía en la actualización atómica del proceso GitHub Actions.
    """
    try:
        raw = subprocess.run(["git", "ls-files", "--stage", "-z", "--", "Datos"],
                             cwd=ROOT, capture_output=True, check=True,
                             timeout=3).stdout
        datos = {}
        for item in raw.split(b"\0"):
            if not item or b"\t" not in item:
                continue
            metadata, name = item.split(b"\t", 1)
            sha = metadata.split()[1].decode()
            datos[name.decode("utf-8", errors="replace")] = sha
        return datos
    except (OSError, subprocess.SubprocessError, IndexError):
        return {}


def _rutas(archivo) -> list[Path] | None:
    rutas = getattr(archivo, "rutas", None)
    if not rutas:
        return None  # Subida manual: siempre procesarla, no reutilizar GitHub.
    paths = []
    for r in rutas:
        p = Path(r)
        if not p.is_absolute():
            p = ROOT / p
        try:
            p.relative_to(ROOT)
        except ValueError:
            return None
        paths.append(p)
    return paths


def resultado_precargado(archivo, categoria: str, periodo: str, claves=None):
    """Devuelve dict persistente o None si falta, está obsoleto o se subió manualmente."""
    rutas = _rutas(archivo)
    if rutas is None:
        return None
    ruta = ruta_cache(categoria, periodo)
    if not ruta.is_file():
        return None
    try:
        stat = ruta.stat()
        cache = _leer_cache(str(ruta), stat.st_size, stat.st_mtime_ns)
        if cache.get("esquema") != ESQUEMA or cache.get("motor_hash") != firma_motor():
            return None
        registros = cache.get("archivos") or {}
        reg = registros.get(archivo.name)
        if not reg:
            return None
        if categoria == "balance":
            lista = sorted(k.strip().upper() for k in (claves or ()) if k.strip())
            if lista != sorted(reg.get("claves") or []):
                return None
        referencias = reg.get("origen") or []
        if len(referencias) != len(rutas):
            return None
        git = _git_blobs()
        # Comprobar identidad y versión de cada parte sin cargar el Excel en memoria.
        for p, ref in zip(rutas, referencias):
            if not p.is_file():
                # Fuente remota en GitHub, sin fichero local descargado; la
                # versión proviene del manifiesto generado por Actions.
                if getattr(archivo, "raiz", "") != "GitHub remoto":
                    return None
                continue
            if p.stat().st_size != ref.get("size"):
                return None
            relativo = str(p.relative_to(ROOT)).replace(os.sep, "/")
            if git and (relativo not in git or git[relativo] != ref.get("blob")):
                return None
        if reg.get("error"):
            raise ValueError("Precálculo de " + archivo.name + ": " + str(reg["error"]))
        return reg.get("resultado")
    except (OSError, EOFError, gzip.BadGzipFile, json.JSONDecodeError, KeyError, TypeError):
        return None


def obtener(archivo, categoria: str, parser, claves=None):
    """Lectura rápida; parser es una función que acepta bytes y nombre."""
    from cen_motor import BalanceResult
    from cen_motor import month_from_filename
    periodo = month_from_filename(archivo.name)
    dato = resultado_precargado(archivo, categoria, periodo, claves)
    if dato is not None:
        return BalanceResult(**dato) if categoria == "balance" else dato
    return parser(archivo.getvalue(), archivo.name)


def estado_procesados() -> dict:
    archivos = sorted(CARPETAS.glob("20??/*.json.gz")) if CARPETAS.exists() else []
    periodos = {m.group(1) for p in archivos
                if (m := re.search(r"(20\d{2}-(?:0[1-9]|1[0-2]))\.json\.gz$", p.name))}
    return {"archivos_mes": len(archivos), "bytes": sum(p.stat().st_size for p in archivos),
            "periodos": sorted(periodos)}


def integridad_precargada(rutas, nombre: str) -> bool:
    """Omitir reconstrucción ZIP en inventario SOLO si el precálculo es vigente."""
    from types import SimpleNamespace
    from cen_motor import month_from_filename
    mapa = {'Balances':'balance','Sscc_archivo4':'sscc4','Sscc_archivo9':'sscc9',
            'Mt_pd':'mtpd','Potencia':'potencia','Exenciones_CI':'exenciaci',
            'Facturas_enel':'factura','Proformas_enel':'proforma'}
    try:
        ruta = Path(rutas[0]).resolve()
        seg = ruta.relative_to(ROOT).parts
        if len(seg)<3 or seg[0]!='Datos':
            return False
        cat = mapa.get(seg[1])
        if not cat:
            return False
        if cat=='balance':
            claves=('MERVAL_ELSOL','MERVAL_LIMACHE')
        else:
            claves=None
        archivo=SimpleNamespace(name=nombre, rutas=tuple(rutas))
        return resultado_precargado(archivo,cat,month_from_filename(nombre),claves) is not None
    except (ValueError, TypeError, KeyError):
        return False
