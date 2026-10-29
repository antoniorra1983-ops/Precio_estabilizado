"""Prueba autónoma de precálculo, rellamada sin Excel y actualización al cambiar origen."""
from pathlib import Path
from tempfile import TemporaryDirectory
from shutil import copy2
from subprocess import run
from io import BytesIO
from openpyxl import Workbook
import gzip
import json
import math
import sys

ORIGEN=Path(__file__).resolve().parent

def guardado(wb, path):
    path.parent.mkdir(parents=True,exist_ok=True)
    wb.save(path)

def hacer_balance(path):
    wb=Workbook();c=wb.active;c.title='Cuadro de Energía'
    c['A1']='TOTAL GENERAL';c['L1']=100000;c['M1']=-100000
    c['B2']='ENEL_GENERACION';c['O2']=100;c['P2']=200
    b=wb.create_sheet('Balance Valorizado')
    for i,(clave,tipo,kwh) in enumerate([
        ('MERVAL_ELSOL','L',100000),('MERVAL_ELSOL','L_D',20000),
        ('MERVAL_ELSOL','L_S',30000),('MERVAL_ELSOL','R',2000000),
        ('MERVAL_LIMACHE','L',20000),('TERCERO','L',1000000)],start=2):
        b[f'A{i}']=clave;b[f'B{i}']='QUILPUE';b[f'C{i}']=110
        b[f'H{i}']=clave;b[f'I{i}']=-kwh;b[f'J{i}']=tipo
        b[f'L{i}']='ENEL_GENERACION'
    guardado(wb,path)

def hacer_archivo4(path):
    wb=Workbook();a=wb.active;a.title='PAGO_EMPRESAS_SC_SSCC'
    a['A2']='ENEL_GENERACION';a['B2']=100;a['C2']=500;a['D2']=-400
    a=wb.create_sheet('PAGO_EMPRESAS_CO')
    a['A2']='ENEL_GENERACION';a['B2']=300;a['C2']=100;a['D2']=200
    guardado(wb,path)

def hacer_archivo9(path):
    wb=Workbook();a=wb.active;a.title='PAGO_RETIRO'
    for c,v in [('E','SUMINISTRADOR'),('G','CLAVE'),('H','TIPO_PAGO'),('J','PAGO')]:
        a[f'{c}1']=v
    for i,(clave,barra,tipo,monto) in enumerate([
        ('MERVAL_ELSOL','QUILPUE','SC_SSCC',1000),('MERVAL_ELSOL','QUILPUE','BESS',200),
        ('MERVAL_ELSOL','QUILPUE','CO',300),('MERVAL_LIMACHE','S.PEDRO','SC_SSCC',100),
        ('MERVAL_LIMACHE','S.PEDRO','BESS',20),('MERVAL_LIMACHE','S.PEDRO','CO',30)
    ],start=2):
        for c,v in [('B',barra),('E','ENEL_GENERACION'),('G',clave),('H',tipo),('J',monto)]:
            a[f'{c}{i}']=v
    guardado(wb,path)

with TemporaryDirectory() as tmp:
    root=Path(tmp)
    for f in ['cen_motor.py','datos_precalculados.py','preprocesar_cen.py','proforma_enel.py','sscc_archivo4.py']:
        copy2(ORIGEN/f,root/f)
    b=root/'Datos/Balances/2026/Balance_2609D.xlsm'
    q=root/'Datos/Sscc_archivo4/2026/4_REMUNERACION_SC_CO_2609_def.xlsx'
    p=root/'Datos/Sscc_archivo9/2026/9_Pagos_Retiros_2609_def.xlsx'
    hacer_balance(b);hacer_archivo4(q);hacer_archivo9(p)
    def ejecutar():
        return run([sys.executable,str(root/'preprocesar_cen.py'),'--root',str(root)],
                   cwd=root,capture_output=True,text=True,check=True).stdout
    out=ejecutar();assert '"calculados": 3' in out,out
    out2=ejecutar();assert '"calculados": 0' in out2,out2
    assert '"reutilizados": 3' in out2,out2
    prog='''
from pathlib import Path
from datos_precalculados import obtener,integridad_precargada
from cen_motor import parse_balance,parse_sscc_empresa,parse_sscc_retiros
root=Path.cwd()
class A:
    def __init__(self,filename,cat):
        self.name=filename;self.rutas=(str(root/'Datos'/cat/'2026'/filename),)
    def getvalue(self):raise AssertionError('No se debe abrir Excel')
b=A('Balance_2609D.xlsm','Balances')
x=obtener(b,'balance',lambda data,name:parse_balance(data,name,['MERVAL_ELSOL','MERVAL_LIMACHE']),claves=['MERVAL_ELSOL','MERVAL_LIMACHE'])
assert x.consumo_efe_kwh==170000,x.consumo_efe_kwh
assert x.retiro_enel_mwh==1170,x.retiro_enel_mwh
assert integridad_precargada(b.rutas,b.name)
q=A('4_REMUNERACION_SC_CO_2609_def.xlsx','Sscc_archivo4')
y=obtener(q,'sscc4',parse_sscc_empresa)
assert y['SC_SSCC']['paga']-y['SC_SSCC']['recibe']==400
assert y['CO']['paga']-y['CO']['recibe']==-200
p=A('9_Pagos_Retiros_2609_def.xlsx','Sscc_archivo9')
z=obtener(p,'sscc9',parse_sscc_retiros)
assert z['suma_tipos']['Quilpué']==1500
assert z['suma_tipos']['Limache']==150
assert z['periodo']=='2026-09'
print('OK: los 3 resultados vienen del cache sin abrir los Excel')
'''
    res=run([sys.executable,'-c',prog],cwd=root,text=True,capture_output=True)
    assert res.returncode==0,res.stdout+'\n'+res.stderr
    print(res.stdout.strip())
    # Cambiar un archivo invalida su caché por tamaño incluso sin git.
    with q.open('ab') as f:f.write(b'CAMBIO')
    invalid='''
from pathlib import Path
from datos_precalculados import resultado_precargado
p=Path.cwd()/'Datos/Sscc_archivo4/2026/4_REMUNERACION_SC_CO_2609_def.xlsx'
class F: name=p.name;rutas=(str(p),)
assert resultado_precargado(F(),'sscc4','2026-09') is None
print('OK: cache invalidado tras cambiar fuente')
'''
    res=run([sys.executable,'-c',invalid],cwd=root,text=True,capture_output=True)
    assert res.returncode==0,res.stdout+'\n'+res.stderr
    print(res.stdout.strip())
    print('OK: hash del motor y mes de origen preservados')
