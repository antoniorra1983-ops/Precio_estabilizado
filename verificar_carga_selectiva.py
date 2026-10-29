from pathlib import Path
import ast
s=Path(__file__).with_name("app.py").read_text(encoding="utf-8")
ast.parse(s)
for literal in [
    'on_change="rerun", key="efe_pestanas_principales"',
    'if resumen_principal.open:',
    'if proforma_principal.open:',
    'if retiros_principal.open:',
    'if cargos_principal.open and cen_grupo.open and sscc_tab.open:',
    'if cargos_principal.open and costos_activos.open:',
    'for item in ((sscc_uploads or []) if cargos_principal.open else []):',
]:
    assert literal in s, literal
print("OK: carga selectiva presente, sintaxis valida")
