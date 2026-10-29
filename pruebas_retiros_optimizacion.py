"""Pruebas sin Streamlit: motor Balance y vista de Retiros CEN."""
from io import BytesIO
from pathlib import Path
import sys
import types
from openpyxl import Workbook

raiz=Path(__file__).resolve().parent
modulo=types.ModuleType('test_motor_balance')
sys.modules['test_motor_balance']=modulo
scope=modulo.__dict__
codigo=(raiz/'app.py').read_text()
exec(codigo.split('\ndef parse_invoice')[0],scope)
parse_balance=scope['parse_balance']
# Build one small compatible CEN book.
wb=Workbook()
cuadro=wb.active
cuadro.title='Cuadro de Energía'
cuadro['A1']='TOTAL GENERAL';cuadro['L1']=1000000;cuadro['M1']=-1000000
cuadro['B2']='ENEL_GENERACION';cuadro['O2']=100;cuadro['P2']=200
balance=wb.create_sheet('Balance Valorizado')

def fila(num,clave,tipo,ret_kwh,suministrador='ENEL_GENERACION'):
    balance[f'A{num}']=clave
    balance[f'B{num}']='QUILPUE' if clave=='MERVAL_ELSOL' else 'S.PEDRO'
    balance[f'C{num}']=110 if clave=='MERVAL_ELSOL' else 13
    balance[f'F{num}']='ENEL GENERACION'
    balance[f'H{num}']='MERVAL QUILPUE' if clave=='MERVAL_ELSOL' else 'MERVAL LIMACHE'
    balance[f'I{num}']=-ret_kwh
    balance[f'J{num}']=tipo
    balance[f'L{num}']=suministrador

fila(2,'MERVAL_ELSOL','L',1000000)
fila(3,'MERVAL_ELSOL','L_D',200000)
fila(4,'MERVAL_ELSOL','L_S',100000)
fila(5,'MERVAL_ELSOL','R',999999)     # should NOT count
fila(6,'MERVAL_LIMACHE','L',50000)
fila(7,'MERVAL_LIMACHE','L_S',10000)
fila(8,'MERVAL_LIMACHE','R',666666) # should NOT count
fila(9,'MERVAL_ELSOL','L',99999,'OTRO_SUMINISTRADOR')  # should NOT count for Enel/EFE
fila(10,'TERCERO','L',2000000)
fila(11,'TERCERO','L_D',1000000)
fila(12,'TERCERO','L_S',500000)
out=BytesIO();wb.save(out)
b=parse_balance(out.getvalue(),'Balance_2609D.xlsm',['MERVAL_ELSOL','MERVAL_LIMACHE'])
assert b.archivo=='Balance_2609D.xlsm'
assert set(b.retiros_enel_por_tipo)=={'L','L_D','L_S'}
assert b.retiro_enel_mwh==4860
assert b.consumo_efe_kwh==1360000
assert len(b.detalle)==5
assert not any(r['tipo']=='R' for r in b.detalle)
from retiros_cen import datos_retiros,_periodo
vista=datos_retiros(b)
assert vista['datos']['Quilpué']['TOTAL']==1300000
assert vista['datos']['Limache']['TOTAL']==60000
assert vista['datos']['Total EFE']['TOTAL']==1360000
assert vista['datos']['Total Enel']['TOTAL']==4860000
assert vista['datos']['Quilpué']['L_S']==100000
assert vista['datos']['Limache']['L_D'] is None
assert len(vista['auditoria'])==5
assert _periodo('Balance_2609D.xlsm')=='2026-09'
assert _periodo('Balance_2610D.xlsm')=='2026-10'

# Business guardrails: Archive4 remains net (PAGA-RECIBE) and not in invoiced sum.
panel=(raiz/'panel_ejecutivo.py').read_text()
assert 'SSCC + CO · Archivo 4 (neto estimado)' in panel
assert "balance_previo=True)" in panel
assert 'efe_panel_periodos_cache' in panel
assert 'efe_reliquidaciones_mes_cache' in codigo
assert '@lru_cache(maxsize=24)\ndef cargar_reliquidaciones_resumen' not in codigo
assert '@st.cache_resource(ttl=1200' in codigo
assert '"📊 Retiros CEN"' in codigo
assert 'L/L_D/L_S' in codigo
print('OK: balance L,L_D,L_S y exclusión R/otro suministrador; Enel y EFE consistentes.')
print('OK: tab Retiros CEN y cachés persistentes en sesión y bytes inmutables.')
print('Prueba: 2026-09, EFE Quilpué 1300000 kWh, Limache 60000 kWh, Enel 4860000 kWh.')
