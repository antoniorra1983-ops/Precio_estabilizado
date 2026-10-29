"""Regresión de carpetas reales, Archivo 4, PSUF dividido y precios Quillota 220.

Ejecutar: python prueba_cobertura_total.py
"""
from pathlib import Path
from tempfile import TemporaryDirectory
from shutil import copy2
from subprocess import run
import gzip, json, sys, io
from openpyxl import Workbook

SRC=Path(__file__).resolve().parent

def save(wb,path):
    path.parent.mkdir(parents=True,exist_ok=True)
    wb.save(path)

def w_4(path):
    w=Workbook();a=w.active;a.title='PAGO_EMPRESAS_SC_SSCC'
    a['A2']='ENEL_GENERACION';a['B2']=200;a['C2']=700;a['D2']=-500
    a=w.create_sheet('PAGO_EMPRESAS_CO')
    a['A2']='ENEL_GENERACION';a['B2']=400;a['C2']=100;a['D2']=300
    save(w,path)

def w_potencia(path):
    w=Workbook();a=w.active;a.title='Balance2';a['C6']=45605;a['E10']='2026-09'
    for col,t in [('B','DESCRIPCION'),('C','BARRA'),('D','TIPO CLAVE'),('E','MEDIDA (MW)'),('F','PRECIO NUDO ($/KW/MES)'),('G','PAGO PSUF ($)')]:a[f'{col}11']=t
    for i,vals in enumerate([
        ('MERVAL_ELSOL','QUILPUE 110','L',-1.4,1800,-2520000),
        ('MERVAL_LIMACHE','S.PEDRO 013','L_D',-.75,1300,-975000),
    ],start=12):
        for col,val in zip('BCDEFG',vals):a[f'{col}{i}']=val
    save(w,path)

def w_precio(path):
    w=Workbook();a=w.active;a.title='DECRETOS PESTABILIZADO'
    for c,v in [('B','QUILLOTA'),('C',220),('D','1'),('F',65.2),('G','D2026')]:a[f'{c}2']=v
    save(w,path)

with TemporaryDirectory() as t:
    r=Path(t)
    for x in ['cen_motor.py','datos_precalculados.py','preprocesar_cen.py','proforma_enel.py','sscc_archivo4.py']:
        copy2(SRC/x,r/x)
    q=r/'Datos/SSCC — Archivo 4/2026/4_REMUNERACION_SC_CO_2609_def.xlsx'
    p=r/'Datos/Potencia de suficiencia/2026/BDef Detalle Sep26_def_v2.xlsm'
    x=r/'Datos/Precios_estabilizados/2026/01 Precios Estabilizados 2609D.xlsm'
    w_4(q);w_potencia(p);w_precio(x)
    b=p.read_bytes();p.unlink();paths=[p.with_name(p.name+f'.part{i:03d}') for i in (1,2)]
    for f,chunk in zip(paths,[b[:len(b)//2],b[len(b)//2:]]):f.write_bytes(chunk)
    def runit():
        result=run([sys.executable,'preprocesar_cen.py','--root','.','--strict'],cwd=r,capture_output=True,text=True)
        assert result.returncode==0,(result.stdout,result.stderr)
        return result.stdout
    stdout=runit()
    assert '"calculados": 3' in stdout,stdout
    assert '"errores": 0' in stdout,stdout
    assert 'Procesando potencia' in stdout
    assert 'Procesando precio' in stdout
    def res(cat):
        with gzip.open(r/f'Datos_procesados/2026/{cat}_2026-09.json.gz','rt',encoding='utf8') as f:return json.load(f)
    pp=res('precio')['archivos'][x.name]['resultado']
    assert pp and pp[0]['valor_clp_kwh']==65.2,pp
    ps=res('potencia')['archivos'][p.name]['resultado']
    assert ps['completo'] and round(ps['total_potencia_retiro_mw'],4)==2.15,ps
    ss=res('sscc4')['archivos'][q.name]['resultado']
    assert ss['SC_SSCC']['paga']-ss['SC_SSCC']['recibe']==500,ss
    assert ss['CO']['paga']-ss['CO']['recibe']==-300,ss
    stdout2=runit()
    assert '"calculados": 0' in stdout2,stdout2
    assert '"reutilizados": 3' in stdout2,stdout2
    script="""
from pathlib import Path
from types import SimpleNamespace
from datos_precalculados import obtener, integridad_precargada
from cen_motor import parse_potencia_cen,parse_price_quillota,parse_sscc_empresa
r=Path.cwd()
for c,n,p in [('sscc4','4_REMUNERACION_SC_CO_2609_def.xlsx','SSCC — Archivo 4'),('potencia','BDef Detalle Sep26_def_v2.xlsm','Potencia de suficiencia'),('precio','01 Precios Estabilizados 2609D.xlsm','Precios_estabilizados')]:
    origen=r/'Datos'/p/'2026'/n
    paths=tuple(str(x) for x in sorted(origen.parent.glob(n+'.part*'))) if c=='potencia' else (str(origen),)
    a=SimpleNamespace(name=n,rutas=paths)
    assert integridad_precargada(paths,n),c
    def explode(*args):raise AssertionError('Abrió Excel en '+c)
    assert obtener(a,c,explode) is not None
"""
    rr=run([sys.executable,'-c',script],cwd=r,capture_output=True,text=True)
    assert rr.returncode==0,(rr.stdout,rr.stderr)
    # Modificación de un archivo obliga solo a recalcular su categoría.
    w_precio(x)
    # Modificar contenido conservando el nombre (y archivo del mismo tamaño)
    content=x.read_bytes();x.write_bytes(content+b'X')
    stdout3=runit()
    assert '"calculados": 1' in stdout3,stdout3
    assert '"reutilizados": 2' in stdout3,stdout3
    print('OK: 3 categorías precalculadas con nombres originales del repositorio')
    print('OK: tarifas Quillota, Potencia dividida y SSCC+CO neto')
    print('OK: reapertura usa cache sin abrir Excel y cambio invalida solo 1 categoría')
