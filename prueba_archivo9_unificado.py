"""Pruebas sin acceso a red del cargo único Archivo 9 del Resumen."""
from pathlib import Path
from types import ModuleType, SimpleNamespace
import ast
import math
import sys

panel=Path(__file__).parent/'panel_ejecutivo.py'
node=ast.parse(panel.read_text(encoding='utf-8'))
functions={'_num','_es','_reporte_mes','_potencia_para_factura','_registro_sitio',
           '_mes_anterior','_nombre_mes','_retiros_del_balance','desglose_todos_los_costos',
           '_total_tarjeta','totalizar_cobros'}
subset=ast.Module(body=[n for n in node.body if isinstance(n,ast.FunctionDef) and n.name in functions],type_ignores=[])
assert {n.name for n in subset.body}==functions
ns={'math':math,'SITIOS':('Quilpué','Limache'),
    'MESES':('Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre')}
exec(compile(subset,str(panel),'exec'),ns)
ns['CARGOS_REGULADOS']=next(ast.literal_eval(n.value) for n in node.body if isinstance(n,ast.Assign)
                        and any(isinstance(t,ast.Name) and t.id=='CARGOS_REGULADOS' for t in n.targets))
tarifas=ModuleType('tarifas_cne')
tarifas.tabla_periodos=lambda cat,years,months: [dict(Mes=f'{years[0]:04d}-{months[0]:02d}',Suministro=s, **{
    'Nacional ($/kWh)':0.,'Zonal ($/kWh)':0.,'Exenciones generales ($/kWh)':0.,
    'SSCC infraestructura ($/kWh)':0.,'CSP >5000 kWh ($/kWh)':0.}) for s in ('Quilpué','Limache')]
sys.modules['tarifas_cne']=tarifas
# Los 24 importes conservan todos los decimales del desglose previo del Archivo 9 definitivo.
importes={
 'Quilpué':{
 'BESS':222347.0792943719,'CO':235414.800705847,'CO_ERNC':10558.14803442172,
 'CPF(+)':21065.54890442864,'CPF(-)':5663.214769264531,'CRA':3510.837668774951,
 'CSF(+)':40511.66032577585,'CSF(-)':1077.65100528629,
 'CTF(+)':85374.9768772734,'CTF(-)':409.1874142328746,
 'REA':3500.788887000678,'SC_SSCC':1389195.22256911},
 'Limache':{
 'BESS':5201.785704466969,'CO':5092.962139718979,'CO_ERNC':279.8859224807461,
 'CPF(+)':592.1663026327416,'CPF(-)':175.6761859496117,'CRA':99.14396902657269,
 'CSF(+)':812.0176174519296,'CSF(-)':22.37124876299625,
 'CTF(+)':2006.546026721033,'CTF(-)':14.08037798473681,
 'REA':154.6013323577121,'SC_SSCC':38371.2491013372},
}
reporte={'periodo':'2026-09','archivo':'9_Pagos_Retiros_2609_def.xlsx',
         'resumen':importes,'suma_tipos':{s:sum(d.values()) for s,d in importes.items()},
         'detalle':[],'faltantes':{'Quilpué':[],'Limache':[]}}
balance=SimpleNamespace(label='2026-09',detalle=[
 {'clave':'MERVAL_ELSOL','kwh':1386541.01},
 {'clave':'MERVAL_LIMACHE','kwh':33251.73}],tarifa_clp_mwh=2.,retiro_enel_mwh=3000000.)
info=[{'Mes':'2026-10','Suministro':s,'Total energía (kWh)':100000.,
       'Costo energía contractual ($)':10000000.,'Estado costo energía':'Calculado'}
      for s in ('Quilpué','Limache')]
tarjetas=ns['desglose_todos_los_costos'](info,{},reliquidaciones={'balances':[balance],
                                                          'sscc_reports':[reporte]})
sscc=[t for t in tarjetas if 'Archivo 9' in t['nombre']]
assert len(sscc)==1,sscc
sscc=sscc[0]
assert sscc['nombre']=='Total Archivo 9 · pagos por retiro'
assert sscc['incluir_total'] and sscc['mes_balance']=='2026-09' and sscc['mes_factura']=='2026-10'
assert round(sscc['por_sitio']['Quilpué']['valor'],2)==2018629.12
assert round(sscc['por_sitio']['Limache']['valor'],2)==52822.49
assert round(ns['_total_tarjeta'](sscc,('Quilpué','Limache')),2)==2071451.60
assert round(ns['_total_tarjeta'](sscc,('Quilpué',)),2)==2018629.12
assert round(ns['_total_tarjeta'](sscc,('Limache',)),2)==52822.49
for selected in [('Quilpué','Limache'),('Quilpué',),('Limache',)]:
    total=ns['totalizar_cobros'](tarjetas,selected)
    registros=[r for r in total['detalle'] if r['Concepto']==sscc['nombre']]
    assert len(registros)==len(selected),registros
    assert all(r['Inclusión']=='Sumado' for r in registros)
    assert abs(sum(r['Monto CLP'] for r in registros) - ns['_total_tarjeta'](sscc,selected))<0.001
assert ns['_mes_anterior']('2027-01')=='2026-12'
reporte_incompleto={'periodo':'2026-09','archivo':'incompleto.xlsx',
                   'resumen':{'Quilpué':{'SC_SSCC':None},'Limache':{'SC_SSCC':1}}}
tarjetas_incompletas=ns['desglose_todos_los_costos'](info,{},reliquidaciones={'balances':[balance],
                                                      'sscc_reports':[reporte_incompleto]})
sscc_incompleto=next(t for t in tarjetas_incompletas if t['nombre']==sscc['nombre'])
assert sscc_incompleto['por_sitio']['Quilpué']['valor'] is None
assert ns['_total_tarjeta'](sscc_incompleto,('Quilpué','Limache')) is None
print('OK: balance septiembre → proforma octubre')
print('OK: único cargo Archivo 9 y agregado en total exactamente una vez')
print('OK: Ambos $2.071.451,60 | Quilpué $2.018.629,12 | Limache $52.822,49')
print('OK: importes pendientes, moneda sin IVA y cierre diciembre/enero')
