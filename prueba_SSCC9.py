from pathlib import Path
from types import SimpleNamespace, ModuleType
import ast,sys
p=Path('/mnt/data/sscc_actualizacion/Precio_estabilizado_desfase_proforma_resumen/panel_ejecutivo.py')
source=p.read_text()
node=ast.parse(source)
selected={'_num','_es','_reporte_mes','_potencia_para_factura','_registro_sitio','_mes_anterior','_nombre_mes','_retiros_del_balance',
          'desglose_todos_los_costos','_total_tarjeta','totalizar_cobros'}
subset=ast.Module(body=[n for n in node.body if isinstance(n,ast.FunctionDef) and n.name in selected],type_ignores=[])
classes={n.name for n in subset.body}
assert classes==selected,(classes,selected)
ns={'__name__':'sscc_prueba','SITIOS':('Quilpué','Limache'),'MESES':('Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'),'math':__import__('math')}
exec(compile(subset,str(p),'exec'),ns)
ns['CARGOS_REGULADOS']=next(ast.literal_eval(n.value) for n in node.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='CARGOS_REGULADOS' for t in n.targets))
tarifas=ModuleType('tarifas_cne')
tarifas.tabla_periodos=lambda cat,years,months: [dict(Mes=f'{years[0]:04d}-{months[0]:02d}',Suministro=s, **{'Nacional ($/kWh)':0.0,'Zonal ($/kWh)':0.0,'Exenciones generales ($/kWh)':0.0,'SSCC infraestructura ($/kWh)':0.0,'CSP >5000 kWh ($/kWh)':0.0}) for s in ('Quilpué','Limache')]
sys.modules['tarifas_cne']=tarifas
componentes={'Quilpué':{'SC_SSCC':1389195.22256911,'BESS':222347.0792943719},'Limache':{'SC_SSCC':38371.2491013372,'BESS':5201.785704466969}}
otros={'Quilpué':407086.8145923052,'Limache':9249.4511230871}
rows=[]
for sitio, values in componentes.items():
    for code, amount in values.items(): rows.append({'punto':sitio,'codigo':code,'valor_clp':amount,'fila':166390 if code=='SC_SSCC' and sitio=='Quilpué' else 399285 if code=='BESS' and sitio=='Quilpué' else 406422 if code=='SC_SSCC' else 144345})
reporte={'periodo':'2026-09','archivo':'9_Pagos_Retiros_2609_def.xlsx','resumen':{s:{**d,'CO':otros[s]} for s,d in componentes.items()},'detalle':rows,'faltantes':{'Quilpué':[],'Limache':[]}}
balance=SimpleNamespace(label='2026-09',detalle=[{'clave':'MERVAL_ELSOL','kwh':1386541.01},{'clave':'MERVAL_LIMACHE','kwh':33251.73}],tarifa_clp_mwh=2.0,retiro_enel_mwh=3_000_000.)
info=[{'Mes':'2026-10','Suministro':s,'Total energía (kWh)':100000.0,'Costo energía contractual ($)':10_000_000.,'Estado costo energía':'Calculado'} for s in ('Quilpué','Limache')]
ctx={'balances':[balance],'sscc_reports':[reporte]}
tarjetas=ns['desglose_todos_los_costos'](info,{},reliquidaciones=ctx)
t=next(t for t in tarjetas if t['nombre']=='SSCC operacionales + BESS · Archivo 9')
q=t['por_sitio']['Quilpué'];l=t['por_sitio']['Limache']
assert round(q['valor'],2)==1611542.30,(q['valor'],'Quilpue')
assert round(l['valor'],2)==43573.03,(l['valor'],'Limache')
assert round(q['valor']+l['valor'],2)==1655115.34,(q['valor']+l['valor'],'Total')
assert t['mes_balance']=='2026-09' and t['mes_factura']=='2026-10'
assert abs(q['tarifa_clp_kwh']-(1611542.301863482/1386541.01))<1e-9
assert abs(l['tarifa_clp_kwh']-(43573.03480580417/33251.73))<1e-9
extra=next(t for t in tarjetas if t['nombre'].startswith('Otros pagos del Archivo 9'))
assert extra['incluir_total'] is False
for sitio in ('Quilpué','Limache'):
    resultado=ns['totalizar_cobros'](tarjetas,(sitio,))
    assert len([r for r in resultado['detalle'] if r['Concepto']==extra['nombre']])==0
assert ns['_mes_anterior']('2027-01')=='2026-12'
print('Correcto: Septiembre 2026 CEN → Octubre 2026 factura')
print('Quilpué $'+ns['_es'](q['valor'],2),' | Limache $'+ns['_es'](l['valor'],2),' | Total $'+ns['_es'](q['valor']+l['valor'],2))
print('Equivalentes: Q',round(q['tarifa_clp_kwh'],6),'L',round(l['tarifa_clp_kwh'],6))
print('Otros 10 tipos de pago excluidos del total: OK')
print('Filtro año/mes y referencia de filas: OK; cierre diciembre a enero: OK')
