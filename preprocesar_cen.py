"""Preprocesa Excel del CEN y proformas a datos persistentes por mes.

Se ejecuta en GitHub Actions. Procesa SOLO archivos nuevos/modificados (blob
Git), guarda todos los resultados y las advertencias como JSON comprimido.
No requiere instalar Streamlit ni ejecuta su interfaz.
"""
from __future__ import annotations
import argparse
from collections import defaultdict
from dataclasses import asdict
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from hashlib import sha1
import gzip
import json
import re
import subprocess
import sys
import os

from cen_motor import (
    month_from_filename, parse_balance, parse_invoice, parse_mt_pd_cen_retiros,
    parse_potencia_cen, parse_exenciones_ci, parse_sscc_retiros,
    parse_sscc_empresa,
)
from proforma_enel import parsear_proforma
from datos_precalculados import ESQUEMA, firma_motor, ruta_cache

TIPOS = {
    'Balances': ('balance', lambda b, n: asdict(parse_balance(
        b, n, ['MERVAL_ELSOL', 'MERVAL_LIMACHE']))),
    'Sscc_archivo4': ('sscc4', parse_sscc_empresa),
    'Sscc_archivo9': ('sscc9', parse_sscc_retiros),
    'Mt_pd': ('mtpd', parse_mt_pd_cen_retiros),
    'Potencia': ('potencia', parse_potencia_cen),
    'Exenciones_CI': ('exenciaci', parse_exenciones_ci),
    'Facturas_enel': ('factura', parse_invoice),
    'Proformas_enel': ('proforma', parsear_proforma),
}


def blobs_git(root: Path) -> dict[str, str]:
    try:
        raw = subprocess.run(['git','ls-files','--stage','-z','--','Datos'],
                             cwd=root, capture_output=True, check=True,
                             timeout=8).stdout
        return {str(path.decode('utf-8')): meta.split()[1].decode()
                for item in raw.split(b'\0') if item and b'\t' in item
                for meta, path in [item.split(b'\t', 1)]}
    except (OSError, subprocess.SubprocessError):
        return {}


def sha_blob(path: Path) -> str:
    data = path.read_bytes()
    return sha1(f'blob {len(data)}\0'.encode() + data).hexdigest()


def inventario(root: Path, git: dict[str, str]) -> list[dict]:
    entradas = []
    data = root / 'Datos'
    if not data.exists():
        return []
    for directorio, (categoria, _) in TIPOS.items():
        raiz = data / directorio
        if not raiz.is_dir():
            continue
        archivos = sorted(p for p in raiz.rglob('*') if p.is_file())
        grupos = defaultdict(list)
        for p in archivos:
            if p.name.startswith('.'):
                continue
            mt = re.fullmatch(r'(.+\.(?:xlsx|xlsm))\.part(\d{3})', p.name, re.I)
            if mt:
                grupos[(p.parent,mt[1])].append((int(mt[2]),p))
            elif p.suffix.lower() in ('.xlsx','.xlsm'):
                grupos[(p.parent,p.name)].append((0,p))
        for (_, nombre), partes in sorted(grupos.items(), key=lambda q: str(q[0])):
            mes = month_from_filename(nombre)
            if not re.fullmatch(r'20\d{2}-(?:0[1-9]|1[0-2])',mes):
                print(f'AVISO {nombre}: mes no identificable, omitido')
                continue
            rutas = [p for _,p in sorted(partes)]
            nums = sorted(n for n,_ in partes)
            if len(nums)>1 and nums != list(range(1,len(nums)+1)):
                raise ValueError(f'Partes incompletas: {nombre}: {nums}')
            if len(nums)==1 and nums[0] not in (0,1):
                raise ValueError(f'Partes incompletas: {nombre}: {nums}')
            # Se admite /AAAA/ o /AAAA/MM/, pero nunca mezcla de años.
            if mes[:4] not in rutas[0].parts:
                print(f'AVISO {nombre}: carpeta anual distinta o ausente; omitido')
                continue
            origen=[]
            for p in rutas:
                rel=p.relative_to(root).as_posix()
                origen.append({'ruta':rel, 'size':p.stat().st_size,
                               'blob':git.get(rel) or sha_blob(p)})
            entradas.append({'categoria':categoria,'mes':mes,'nombre':nombre,
                             'rutas':rutas,'origen':origen})
    return entradas


def leer_cache(path):
    if not path.is_file():
        return None
    try:
        with gzip.open(path, 'rt', encoding='utf8') as f:
            return json.load(f)
    except (OSError, EOFError, json.JSONDecodeError):
        return None


def guardar(path, objeto):
    path.parent.mkdir(parents=True,exist_ok=True)
    original = json.dumps(objeto, ensure_ascii=False, sort_keys=True,
                          separators=(',',':'),allow_nan=False).encode('utf-8')
    salida = BytesIO()
    with gzip.GzipFile(filename='', fileobj=salida, mode='wb',mtime=0,
                       compresslevel=7) as g:
        g.write(original)
    if path.is_file() and path.read_bytes()==salida.getvalue():
        return False
    temporal = path.with_suffix('.tmp')
    temporal.write_bytes(salida.getvalue())
    temporal.replace(path)
    return True


def procesar(root: Path, solo: set[str] | None = None) -> dict:
    root=root.resolve()
    os.chdir(root)
    git=blobs_git(root)
    disponibles=inventario(root,git)
    organizados=defaultdict(list)
    for a in disponibles:
        organizados[(a['categoria'],a['mes'])].append(a)
    firma=firma_motor()
    stats={'calculados':0,'reutilizados':0,'errores':0,'archivos_generados':0}
    for (cat,mes), sources in sorted(organizados.items()):
        if solo and mes not in solo:
            continue
        dst=ruta_cache(cat,mes)
        anterior=leer_cache(dst)
        antiguos=(anterior or {}).get('archivos',{}) if (anterior or {}).get('motor_hash')==firma else {}
        resultado={'esquema':ESQUEMA,'motor_hash':firma,'categoria':cat,
                   'periodo':mes,'archivos':{}}
        for item in sources:
            archivo=item['nombre']
            previo=antiguos.get(archivo)
            if previo and previo.get('origen') == item['origen']:
                resultado['archivos'][archivo]=previo
                stats['reutilizados']+=1
                continue
            _,parser=next(d for carpeta,d in TIPOS.items() if d[0]==cat)
            print(f'Procesando {cat} {mes}: {archivo}',flush=True)
            reg={'origen':item['origen'],'claves':['MERVAL_ELSOL','MERVAL_LIMACHE'] if cat=='balance' else []}
            try:
                datos=b''.join(p.read_bytes() for p in item['rutas'])
                value=parser(datos,archivo)
                json.dumps(value, allow_nan=False)
                reg['resultado']=value
                stats['calculados']+=1
            except Exception as exc:
                reg['error']=f'{type(exc).__name__}: {str(exc)[:320]}'
                stats['errores']+=1
                print(f'ERROR {cat} {mes} {archivo}: {reg["error"]}',file=sys.stderr)
            resultado['archivos'][archivo]=reg
        if guardar(dst,resultado):
            stats['archivos_generados']+=1
    # Retirar entradas obsoletas al eliminar fuentes del repositorio (solo en ejecución completa).
    if not solo:
        esperados={str(ruta_cache(c,m)) for c,m in organizados}
        from datos_precalculados import CARPETAS
        for prev in CARPETAS.glob('20??/*.json.gz') if CARPETAS.exists() else []:
            if str(prev) not in esperados:
                prev.unlink()
                print('Eliminado precálculo obsoleto:',prev)
    print('PRECALC RESULTADO:',json.dumps(stats,ensure_ascii=False),flush=True)
    return stats


if __name__=='__main__':
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--root',type=Path,default=Path(__file__).resolve().parent)
    ap.add_argument('--solo',action='append',help='Mes 2026-09 (repetible)')
    ap.add_argument('--strict',action='store_true',help='Fallar si un origen tiene error')
    args=ap.parse_args()
    report=procesar(args.root,set(args.solo) if args.solo else None)
    if args.strict and report['errores']:
        sys.exit(1)
