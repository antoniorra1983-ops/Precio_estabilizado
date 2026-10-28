"""Pruebas de SSCC + CO (Archivo 4): fórmula, signos, período, total y panel."""
from pathlib import Path
from types import ModuleType, SimpleNamespace
import ast
import math
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT))
from sscc_archivo4 import calcular_netos_empresa

empresa={
 'periodo':'2026-09', 'archivo':'4_REMUNERACION_SC_CO_2609_def.xlsx',
 'SC_SSCC':{'paga':2129859390.680467,'recibe':1595179553.245922,'neto':-534679837.4345448},
 'CO':{'paga':313867223.6396561,'recibe':520238890.992852,'neto':206371667.35319585},
}
netos=calcular_netos_empresa(empresa)
assert math.isclose(netos['SC_SSCC']['costo_neto'], 534679837.4345448,abs_tol=.001)
assert math.isclose(netos['CO']['costo_neto'], -206371667.35319585,abs_tol=.001)
assert math.isclose(netos['total_costo_neto'], 328308170.08134896,abs_tol=.001)
assert netos['CO']['costo_neto'] < 0
try:
    calcular_netos_empresa({**empresa,'CO':{'paga':100,'recibe':500,'neto':-400}})
    raise AssertionError('Se aceptó neto del CEN con signo incorrecto')
except ValueError: pass
try:
    calcular_netos_empresa({**empresa, 'CO': None})
    raise AssertionError('Se admitió CO ausente como cero')
except ValueError: pass

# Probar cálculo sin importar la interfaz Streamlit ni abrir archivos remotos.
app_ast=ast.parse((ROOT/'app.py').read_text(encoding='utf-8'))
node=next(n for n in app_ast.body if isinstance(n,ast.FunctionDef) and n.name=='calcular_prorrateo_sscc')
code=compile(ast.Module(body=[node],type_ignores=[]),str(ROOT/'app.py'),'exec')
balance=SimpleNamespace(label='2026-09',retiro_enel_mwh=1000, retiros_enel_por_tipo={})
site_consumption={'Quilpué':100000.,'Limache':20000.}
ns={'BalanceResult':object,'SITE_NAMES':('Quilpué','Limache'),
 'SSCC_KEYS':{'Quilpué':'MERVAL_ELSOL','Limache':'MERVAL_LIMACHE'},
 'balance_consumption_by_site':lambda x:site_consumption,
 'calcular_netos_empresa':calcular_netos_empresa,
 'following_month':lambda p:'2026-10' if p=='2026-09' else '2027-01'}
exec(code,ns)
pr=ns['calcular_prorrateo_sscc'](empresa,balance)
assert pr['mes_factura']=='2026-10'
assert math.isclose(pr['detalle'][0]['participacion_enel_pct'],10.,abs_tol=1e-10)
assert math.isclose(pr['detalle'][0]['sscc_neto_estimado_clp'],53467983.74345448,abs_tol=.01)
assert math.isclose(pr['detalle'][0]['co_neto_estimado_clp'],-20637166.735319585,abs_tol=.01)
assert math.isclose(pr['detalle'][0]['sscc_co_neto_estimado_clp'],32830817.0081349,abs_tol=.01)
assert math.isclose(pr['sscc_total_estimado'],netos['total_costo_neto']*.12,abs_tol=.01)
pr_bess=ns['calcular_prorrateo_sscc'](empresa,balance,2.5)
assert math.isclose(pr_bess['sscc_total_estimado']-pr['sscc_total_estimado'],300000.,abs_tol=.01)
assert math.isclose(pr_bess['detalle'][0]['co_neto_estimado_clp'],pr['detalle'][0]['co_neto_estimado_clp'],abs_tol=.001)
try:
    ns['calcular_prorrateo_sscc'](empresa,SimpleNamespace(label='2026-08',retiro_enel_mwh=1000))
    raise AssertionError('Período de balance incorrecto admitido')
except ValueError: pass

# Probar panel: 1 referencia neta y total de Archivo 9 sin alteraciones.
panel_ast=ast.parse((ROOT/'panel_ejecutivo.py').read_text(encoding='utf-8'))
funcs={'_num','_es','_reporte_mes','_potencia_para_factura','_registro_sitio',
       '_mes_anterior','_nombre_mes','_retiros_del_balance','desglose_todos_los_costos',
       '_total_tarjeta','totalizar_cobros'}
subset=ast.Module(body=[n for n in panel_ast.body if isinstance(n,ast.FunctionDef) and n.name in funcs],type_ignores=[])
const=next(ast.literal_eval(n.value) for n in panel_ast.body if isinstance(n,ast.Assign)
          and any(isinstance(t,ast.Name) and t.id=='CARGOS_REGULADOS' for t in n.targets))
ps={'math':math,'SITIOS':('Quilpué','Limache'),
    'MESES':('Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'),
    'CARGOS_REGULADOS':const,'calcular_netos_empresa':calcular_netos_empresa}
exec(compile(subset,str(ROOT/'panel_ejecutivo.py'),'exec'),ps)
tarifas=ModuleType('tarifas_cne')
tarifas.tabla_periodos=lambda cat,years,months: [dict(Mes=f'{years[0]:04d}-{months[0]:02d}',Suministro=s, **{
    'Nacional ($/kWh)':0.,'Zonal ($/kWh)':0.,'Exenciones generales ($/kWh)':0.,
    'SSCC infraestructura ($/kWh)':0.,'CSP >5000 kWh ($/kWh)':0.}) for s in ('Quilpué','Limache')]
sys.modules['tarifas_cne']=tarifas
b=SimpleNamespace(label='2026-09',detalle=[
    {'clave':'MERVAL_ELSOL','kwh':100000.},
    {'clave':'MERVAL_LIMACHE','kwh':20000.}],
    tarifa_clp_mwh=0.,retiro_enel_mwh=1000.)
info=[{'Mes':'2026-10','Suministro':s,'Total energía (kWh)':100000.,
       'Costo energía contractual ($)':10000000.,'Estado costo energía':'Calculado'}
      for s in ('Quilpué','Limache')]
tipo9=['SC_SSCC','BESS','CO','CO_ERNC','CPF(+)','CPF(-)','CSF(+)','CSF(-)','CTF(+)','CTF(-)','CRA','REA']
reporte9={'periodo':'2026-09','archivo':'9_Pagos_Retiros_2609_def.xlsx',
          'resumen':{s:{c:100. if c=='SC_SSCC' else 0. for c in tipo9}
                     for s in ('Quilpué','Limache')},
          'suma_tipos':{'Quilpué':100.,'Limache':100.},
          'detalle':[]}
tarjetas=ps['desglose_todos_los_costos'](info,{},reliquidaciones={
 'balances':[b], 'sscc_companies':[empresa], 'sscc_reports':[reporte9]})
ref=next(t for t in tarjetas if 'Archivo 4' in t['nombre'])
assert len([t for t in tarjetas if 'Archivo 4' in t['nombre']])==1
assert not ref['incluir_total']
assert ref['mes_factura']=='2026-10' and ref['mes_balance']=='2026-09'
assert math.isclose(ref['por_sitio']['Quilpué']['valor'],netos['total_costo_neto']*.10,abs_tol=.01)
assert math.isclose(ref['por_sitio']['Limache']['valor'],netos['total_costo_neto']*.02,abs_tol=.01)
assert 'CO' in ref['por_sitio']['Quilpué']['detalle'] if 'detalle' in ref['por_sitio']['Quilpué'] else True
archivo9=next(t for t in tarjetas if 'Archivo 9' in t['nombre'])
assert archivo9['incluir_total'] and math.isclose(ps['_total_tarjeta'](archivo9,('Quilpué','Limache')),200.)
resultado=ps['totalizar_cobros'](tarjetas,('Quilpué','Limache'))
assert len([r for r in resultado['detalle'] if 'Archivo 4' in r['Concepto'] and r['Inclusión']=='Sumado'])==0
assert len([r for r in resultado['detalle'] if 'Archivo 9' in r['Concepto'] and r['Inclusión']=='Sumado'])==2

# Si no llega archivo 4, la referencia debe quedar pendiente, no convertirse en cero.
no4=ps['desglose_todos_los_costos'](info,{},reliquidaciones={
    'balances':[b],'sscc_reports':[reporte9]})
pending=next(t for t in no4 if 'Archivo 4' in t['nombre'])
assert all(v['valor'] is None for v in pending['por_sitio'].values())
print('OK: fórmula SSCC neto + CO neto y signo negativo de CO')
print('OK: septiembre CEN = octubre de factura y prorrateo por retiro')
print('OK: Archivo 4 fuera del total, Archivo 9 único sin cambios')
print('OK: validación neto CEN, ausencia de CO y referencia BESS opcional')
