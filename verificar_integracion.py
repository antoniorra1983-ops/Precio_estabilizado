"""Validación de sintaxis y montaje de las pestañas activas."""
from pathlib import Path
import ast
from datos_precalculados import categoria_de_carpeta
ROOT=Path(__file__).resolve().parent
s=(ROOT/'app.py').read_text(encoding='utf8')
for p in ['app.py','cen_motor.py','preprocesar_cen.py','proforma_enel.py','datos_precalculados.py']:
    ast.parse((ROOT/p).read_text(encoding='utf8'),filename=p)
for literal in ['if resumen_principal.open:', 'if proforma_principal.open:',
                'if retiros_principal.open:', 'obtener_precios_estabilizados(price_upload)',
                'def obtener_precios_estabilizados(archivo):',
                'on_change="rerun", key="efe_pestanas_principales"']:
    assert literal in s,literal
for k,v in {'SSCC — Archivo 4':'sscc4','Potencia de suficiencia':'potencia',
             'Sscc_archivo9':'sscc9','Precios_estabilizados':'precio'}.items():
    assert categoria_de_carpeta(k)==v,(k,categoria_de_carpeta(k))
workflow=(ROOT/'.github/workflows/precalcular-cen.yml').read_text(encoding='utf8')
assert 'python preprocesar_cen.py --root . --strict' in workflow
print('OK: interfaz, alias, compatibilidad y workflow verificados')
