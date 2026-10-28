"""Pruebas locales de fechas/conciliación, sin ejecutar Streamlit Cloud."""
from pathlib import Path
import ast
import sys
import types

DIR=Path(__file__).resolve().parent
sys.path.insert(0,str(DIR))
from proforma_enel import DESFASADOS, mes_anterior, origen_cargo, parsear_proforma

assert mes_anterior('2026-09')=='2026-08'
assert mes_anterior('2026-10')=='2026-09'
assert mes_anterior('2027-01')=='2026-12'
conceptos={
    'Servicios Complementarios':'Servicios complementarios',
    'Sobrecostos MT':'Mínimos técnicos (MT)',
    'Sobrecosto PD':'Partidas y detenciones (PD)',
    'Compensacion PE':'Precio Estabilizado',
    'Cargo Peaje Tx Nacional (Cliente Indiv.)':'Liquidación de peaje',
}
for texto in conceptos:
    assert origen_cargo(texto,'2026-09')[0]=='2026-08'
    assert origen_cargo(texto,'2026-10')[0]=='2026-09'
    assert origen_cargo(texto,'2027-01')[0]=='2026-12'
for texto in ('Cargo Mensual por uso del Sistema de Transmisión Nacional',
              'Cargo Mensual por uso del Sistema de Transmisión Zonal',
              'Peaje distribución', 'Potencia de punta: 4.2 MW',
              'Cargo Mensual por Consumo de Energía Activa',
              'Cargo por Servicio Público FET'):
    assert origen_cargo(texto,'2026-09')[0]=='2026-09'
assert origen_cargo('Reliquidación mes de agosto de 2026','2026-10')[0]=='2026-08'

b=Path('/mnt/data/EFE_VALPO_SEP2026.xlsx').read_bytes()
p=parsear_proforma(b,'EFE_VALPO_SEP2026.xlsx')
assert p['periodo']=='2026-09'
assert p['mes_reliquidacion']=='2026-08'
assert p['balance_cen_declarado']=='2026-08'
assert p['mes_referencia_peaje_anexo']=='2026-05'
assert p['suministros']['Quilpué']['total_proforma']==191841743.0
assert p['suministros']['Limache']['total_proforma']==6499525.0
for sitio, info in p['suministros'].items():
    grupos=[]
    for cargo in info['cargos']:
        assert cargo['Mes de proforma']=='2026-09'
        if cargo['Grupo'] in DESFASADOS:
            assert cargo['Mes de origen']=='2026-08'
            grupos.append(cargo['Grupo'])
        if cargo['Grupo'] in ('Energía contractual','Potencia de punta','Transmisión','Peaje distribución','Servicio público CSP'):
            assert cargo['Mes de origen']=='2026-09'
    for concept in ('Servicios complementarios','Mínimos técnicos (MT)',
                    'Partidas y detenciones (PD)','Precio Estabilizado'):
        assert concept in grupos,(sitio, concept)
assert any(c['Grupo']=='Liquidación de peaje' for c in p['suministros']['Quilpué']['cargos'])
assert p['suministros']['Quilpué']['total_proforma']+p['suministros']['Limache']['total_proforma']==198341268

# Pruebas de Resumen con respuestas CEN del mes previo, sin red/Streamlit.
sys.modules['streamlit']=types.ModuleType('streamlit')
fake_tarifas=types.ModuleType('tarifas_cne')
def tabla_periodos(catalogo,years,months):
    mes=f'{years[0]}-{months[0]:02d}'
    return [dict({'Mes':mes,'Suministro':s},**{
        'Nacional ($/kWh)':2,'Zonal ($/kWh)':3,
        'Exenciones generales ($/kWh)':4,'SSCC infraestructura ($/kWh)':5,
        'CSP >5000 kWh ($/kWh)':6}) for s in ('Quilpué','Limache')]
fake_tarifas.tabla_periodos=tabla_periodos
sys.modules['tarifas_cne']=fake_tarifas
from panel_ejecutivo import desglose_todos_los_costos,totalizar_cobros
info=[{'Mes':'2026-09','Suministro':s,'Total energía (kWh)':k,
       'Costo energía contractual ($)':20000.0,'Estado costo energía':'Calculado',
       'Precio contrato (US$/MWh)':30.,'Dólar aplicado ($/US$)':900.}
      for s,k in [('Quilpué',10000.),('Limache',6000.)]]
from types import SimpleNamespace
balance=SimpleNamespace(label='2026-08',tarifa_clp_mwh=1500.,retiro_enel_mwh=100.,
    detalle=[{'clave':'MERVAL_ELSOL','kwh':10000.},
             {'clave':'MERVAL_LIMACHE','kwh':6000.}])
ctx={'balances':[balance],
     'sscc_reports':[{'periodo':'2026-08','archivo':'9_Ago.xlsx',
            'resumen':{s:{'SC_SSCC':1000.,'BESS':250.,'CO':3.} for s in ('Quilpué','Limache')},
            'faltantes':{s:[] for s in ('Quilpué','Limache')}}],
     'sscc_companies':[], 'mt_pd_directos':[],
     'potencia_reports':[],
     'exenciones_ci_reports':[{'periodo':'2026-08','total_clp':555.,'archivo':'PagoExenCI_Ago.xlsm'},
                              {'periodo':'2026-09','total_clp':999.,'archivo':'PagoExenCI_Sep.xlsm'}],
     'proformas_enel':[p]}
resultado=desglose_todos_los_costos(info, {},reliquidaciones=ctx)
target={r['nombre']:r for r in resultado}
pe=target['Precio Estabilizado (PE)']
assert pe['mes_balance']=='2026-08' and pe['mes_factura']=='2026-09'
assert pe['por_sitio']['Quilpué']['valor']==15000.
ci=target['Exenciones CI · Liquidación de peaje']
assert ci['mes_balance']=='2026-08' and ci['por_sitio']['Quilpué']['valor']==555.
assert ci['por_sitio']['Limache']['valor']==0.0
peaje=target['Liquidación de peaje CI · Proforma']
assert not peaje['incluir_total']
assert abs(peaje['por_sitio']['Quilpué']['valor']-8444622.96163837)<0.01
assert peaje['por_sitio']['Limache']['valor'] is None
sscc=target['SSCC operacionales + BESS · Archivo 9']
assert sscc['por_sitio']['Quilpué']['valor']==1250.
total=totalizar_cobros(resultado,['Quilpué','Limache'])
assert all(row['Concepto']!='Liquidación de peaje CI · Proforma' for row in total['detalle'])
assert total['subtotal'] is not None
for fn in ['app.py','panel_ejecutivo.py','proforma_enel.py']:
    ast.parse((DIR/fn).read_text(encoding='utf-8'))
print('PRUEBAS OK: período de origen, 5 cargos, sumas de proforma, Exenciones CI, peaje referencial y sintaxis.')
