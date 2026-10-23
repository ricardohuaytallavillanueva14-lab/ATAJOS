from pathlib import Path
from datetime import datetime, date
import sys, tempfile, unittest, zipfile, json, ast, io, contextlib
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from src import ocr_sat as c
from openpyxl import Workbook, load_workbook
from lxml import etree as ET
from PIL import Image, ImageDraw

def libro_ficticio(path):
    wb=Workbook(); ws=wb.active; ws.title='2026'
    ws.append(['EXPDIENTE','numValorPred','fechEmiPred','fechNotiPred','fecha emision REC','Ffecha de notificacion REC','numValorArbitrio','fechEmiArbi','fechNotiArbi','CodPropietario','CodSigeti','CodSimi','nombre'])
    for number,em,notification in [(1,datetime(2026,9,2),None),(2,'03/09/2026',None),(3,'02/09/2026','10/09/2026'),(4,'02/09/2026','=DATE(2026,9,10)'),(5,'02/09/2026',None),(5,'02/09/2026',None)]:
        ws.append([f'16-012-{number:09d}',None,None,None,em,notification,None,None,None,None,'123',None,'PERSONA FICTICIA'])
    wb.create_sheet('2025')['A1']='SIN CAMBIOS'
    wb.save(path)

class ExcelScopeTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.source=Path(self.temp.name)/'original.xlsx'
        self.output=Path(self.temp.name)/'resultado.xlsx'
        libro_ficticio(self.source)
        self.index,_,self.epoch=c.build_index(self.source,('2026',))
    def tearDown(self):self.temp.cleanup()
    def test_valid_match_and_iso_date(self):
        d=dict(expediente='16-012-1-2026',codigo='00123',resolucion='UNO',fecha_notificacion='15/09/2026')
        target,issues=c.evaluar(d,self.index)
        self.assertEqual(issues,[]);self.assertEqual(target['celda'],'F2')
        self.assertEqual(c.fecha(target['emision']),date(2026,9,2))
    def test_scope_and_existing_values_rejected(self):
        for key in [('2025','F2'),('2026','E2'),('2026','F3'),('2026','F4'),('2026','F5'),('2026','F6')]:
            with self.subTest(key=key),self.assertRaises(ValueError):
                c.patch_xlsx(self.source,self.output,{key:date(2026,9,15)},self.epoch)
        with self.assertRaises(ValueError):
            c.patch_xlsx(self.source,self.source,{('2026','F2'):date(2026,9,15)},self.epoch)
    def test_preserves_other_cells_and_zip_parts(self):
        c.patch_xlsx(self.source,self.output,{('2026','F2'):date(2026,9,15)},self.epoch)
        with zipfile.ZipFile(self.source) as a,zipfile.ZipFile(self.output) as b:
            self.assertEqual(a.namelist(),b.namelist())
            changed=[n for n in a.namelist() if a.read(n)!=b.read(n)]
            self.assertEqual(set(changed),{'xl/styles.xml','xl/worksheets/sheet1.xml'})
            roots=[ET.fromstring(z.read('xl/worksheets/sheet1.xml')) for z in (a,b)]
            for root in roots:
                cell=root.find('.//s:c[@r="F2"]',c.N)
                if cell is not None:cell.getparent().remove(cell)
            self.assertEqual(ET.tostring(roots[0],method='c14n'),ET.tostring(roots[1],method='c14n'))
        wb=load_workbook(self.output)
        self.assertEqual(wb['2026']['F2'].value.date(),date(2026,9,15))
        self.assertEqual(wb['2026']['F5'].value,'=DATE(2026,9,10)')
        wb.close()

class ConfigurarTests(unittest.TestCase):
    def setUp(self):
        self.original=(c.HOJA,c.FECHA_EMISION)
        self.temp=tempfile.TemporaryDirectory()
        self.source=Path(self.temp.name)/'original.xlsx'
        self.output=Path(self.temp.name)/'resultado.xlsx'
        libro_ficticio(self.source)
    def tearDown(self):
        c.HOJA,c.FECHA_EMISION=self.original
        self.temp.cleanup()
    def test_rejects_bad_values(self):
        for hoja,fecha in [('26',None),(None,'texto'),('2025','02/09/2026'),('abcd',None)]:
            with self.subTest(hoja=hoja,fecha=fecha),self.assertRaises(ValueError):
                c.configurar(hoja,fecha)
        self.assertEqual((c.HOJA,c.FECHA_EMISION),self.original)
    def test_scope_follows_configuration(self):
        c.configurar('2026','03/09/2026')
        index,_,epoch=c.build_index(self.source,('2026',))
        d=dict(expediente='16-012-1-2026',codigo='123',resolucion='UNO',fecha_notificacion='15/09/2026')
        target,issues=c.evaluar(d,index)
        self.assertTrue(any('Fuera del alcance' in i for i in issues))
        d['expediente']='16-012-2-2026'
        target,issues=c.evaluar(d,index)
        self.assertEqual(issues,[]);self.assertEqual(target['celda'],'F3')
        with self.assertRaises(ValueError):
            c.patch_xlsx(self.source,self.output,{('2026','F2'):date(2026,9,15)},epoch)
        c.patch_xlsx(self.source,self.output,{('2026','F3'):date(2026,9,15)},epoch)
        self.assertTrue(self.output.exists())
        with self.assertRaises(ValueError):
            c.build_index(self.source,('2025',))

class MotorFalso:
    nombre='motor ficticio'
    def __init__(self):self.vistas=0
    def leer(self,src):
        self.vistas+=1
        Image.open(src).load()
        return dict(expediente='16-012-1-2026',codigo='123',resolucion='UNO',fecha_notificacion='15/09/2026',
                    fecha_emision='02/09/2026',fecha_texto=None,observaciones=''),{'motor':self.nombre}

class ProyectoBase(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        base=Path(self.temp.name)
        self.fotos=base/'fotos'; (self.fotos/'A'/'sub').mkdir(parents=True); (self.fotos/'B').mkdir()
        for i,rel in enumerate(['A/sub/f1.png','A/f2.png','B/f3.png']):
            Image.new('RGB',(40,30),(i*40,120,200)).save(self.fotos/rel)
        (self.fotos/'B/copia_f3.png').write_bytes((self.fotos/'B/f3.png').read_bytes())
        self.excel=base/'original.xlsx'; libro_ficticio(self.excel)
        self.salida=base/'salida'
    def tearDown(self):self.temp.cleanup()
    def proyecto(self):
        with contextlib.redirect_stdout(io.StringIO()):
            return c.Proyecto(self.fotos,self.excel,self.salida)

class ProyectoTests(ProyectoBase):
    def test_inventory_cache_and_resume(self):
        p=self.proyecto()
        llamadas=[]
        real=c.sha256
        def contador(path):
            if Path(path).suffix=='.png': llamadas.append(path)
            return real(path)
        c.sha256=contador
        try:
            with contextlib.redirect_stdout(io.StringIO()) as out:
                p.inventario()
            self.assertEqual(len(llamadas),4)
            self.assertEqual(len(p.active),3)
            self.assertIn('registros escritos: 3',out.getvalue())
            self.assertTrue((self.salida/'inventario_cache.json').exists())
            llamadas.clear()
            q=self.proyecto()
            with contextlib.redirect_stdout(io.StringIO()) as out:
                q.inventario()
            self.assertEqual(llamadas,[])
            self.assertIn('reutilizadas de la caché: 4',out.getvalue())
            self.assertIn('registros escritos: 0',out.getvalue())
            with contextlib.redirect_stdout(io.StringIO()):
                q.inventario(rehash=True)
            self.assertEqual(len(llamadas),4)
        finally:
            c.sha256=real
        cfg=json.loads((self.salida/'proyecto.json').read_text())
        self.assertEqual(cfg['fecha_emision'],'2026-09-02')
        # Carpetas creadas por la versión anterior (sin fecha_emision) siguen siendo válidas.
        cfg.pop('fecha_emision'); cfg['version']='2026-09-02-v1'
        (self.salida/'proyecto.json').write_text(json.dumps(cfg))
        self.proyecto()
        cfg['excel_sha256']='0'*64
        (self.salida/'proyecto.json').write_text(json.dumps(cfg))
        with self.assertRaises(ValueError):
            self.proyecto()
    def test_sample_processing_and_pending(self):
        p=self.proyecto()
        with contextlib.redirect_stdout(io.StringIO()):
            p.inventario()
        muestra=c.seleccionar_muestra(p.listar_fotos(),p.fotos,cantidad=2)
        ids=p.ids_de(muestra)
        self.assertEqual(len(ids),2)
        motor=MotorFalso()
        with contextlib.redirect_stdout(io.StringIO()):
            resumen=p.procesar(motor,limite=10,solo=ids)
        self.assertEqual(resumen['procesadas'],2); self.assertEqual(motor.vistas,2)
        self.assertEqual(sorted(p.records[k]['estado'] for k in ids),['POR_REVISAR','POR_REVISAR'])
        self.assertEqual(len(p.pendientes()),1)
        with contextlib.redirect_stdout(io.StringIO()):
            resumen=p.procesar(motor,limite=10)
        self.assertEqual(resumen['procesadas'],1); self.assertEqual(p.pendientes(),[])
        self.assertIn('segundos',p.records[ids[0]]['evidencia_ocr'])
        with self.assertRaises(ValueError):
            p.ids_de([self.fotos/'no_existe.png'])

class FiabilidadTests(unittest.TestCase):
    def test_vista_tinta_aisla_azul(self):
        im=Image.new('RGB',(200,60),'white'); dr=ImageDraw.Draw(im)
        dr.line((0,55,200,55),fill='black',width=3)            # línea impresa
        dr.text((5,5),'FECHA DE NOTIFICACION',fill='black')    # rótulo impreso
        dr.rectangle((120,15,160,40),fill=(40,60,200))         # trazo de bolígrafo azul
        tinta,fraccion=c.vista_tinta(im)
        self.assertIsNotNone(tinta); self.assertGreater(fraccion,0.05)
        px=tinta.load()
        self.assertEqual(px[140,28],(0,0,0))       # tinta conservada
        self.assertEqual(px[100,55],(255,255,255)) # línea negra eliminada
        self.assertEqual(px[10,8],(255,255,255))   # texto impreso eliminado
        sin,fr=c.vista_tinta(Image.new('RGB',(50,50),'white'))
        self.assertIsNone(sin); self.assertEqual(fr,0.0)
    def test_dates_in_text(self):
        self.assertEqual(c.dates_in_text('15 /09 /2026'),['2026-09-15'])
        self.assertEqual(c.dates_in_text('15 09  2026'),['2026-09-15'])
        self.assertEqual(c.dates_in_text('15092026'),['2026-09-15'])
        self.assertEqual(c.dates_in_text('15-9-2026 y 03/06/2026'),['2026-06-03','2026-09-15'])
        self.assertEqual(c.dates_in_text('31022026'),[])
        self.assertEqual(c.dates_in_text('N 15 021 000021666'),[])
    def test_nivel_confianza(self):
        v=lambda vista,fechas:{'vista':vista,'fechas':fechas,'texto':''}
        self.assertEqual(c.nivel_confianza([v('color',['2026-09-15']),v('trocr_color',['2026-09-15'])],'2026-09-15'),('ALTA',['paddle','trocr']))
        self.assertEqual(c.nivel_confianza([v('color',['2026-09-15']),v('tinta',['2026-09-15']),v('trocr_color',[])],'2026-09-15')[0],'MEDIA')
        self.assertEqual(c.nivel_confianza([v('color',['2026-09-15']),v('vlm_color',['2026-09-16'])],None)[0],'CONTRADICTORIA')
        self.assertEqual(c.nivel_confianza([v('color',[]),v('trocr_color',[])],None)[0],'SIN_FECHA')
    def test_alertas_plausibilidad(self):
        original=c.VENTANA_DIAS
        try:
            c.configurar(ventana_dias=30)
            t={'emision':'2026-09-02'}
            self.assertEqual(c.alertas({'fecha_notificacion':'15/09/2026'},t),[])
            avisos=c.alertas({'fecha_notificacion':'15/12/2026'},t)
            self.assertEqual(len(avisos),1); self.assertIn('104 días',avisos[0])
            self.assertTrue(any('domingo' in a for a in c.alertas({'fecha_notificacion':'06/09/2026'},t)))
            with self.assertRaises(ValueError): c.configurar(ventana_dias=0)
        finally:
            c.VENTANA_DIAS=original

class ExportacionLegalTests(ProyectoBase):
    def preparar(self):
        p=self.proyecto()
        with contextlib.redirect_stdout(io.StringIO()):
            p.inventario(); p.procesar(MotorFalso(),limite=10)
        return p
    def datos(self,n=1):
        return dict(expediente=f'16-012-{n}-2026',codigo='123',resolucion='UNO',fecha_notificacion='15/09/2026',fecha_emision='02/09/2026')
    def test_aprobaciones_evidencia_y_doble_revision(self):
        p=self.preparar(); rid=sorted(p.active)[0]
        p.records[rid]['confianza']='MEDIA'; p.records[rid]['evidencia_ocr']['roi_fecha']=[0.1,0.6,0.9,0.9]
        p.aprobar(rid,self.datos(),'Ana')
        r=p.records[rid]
        self.assertEqual(r['estado'],'APROBADO'); self.assertEqual(len(r['aprobaciones']),1)
        self.assertEqual(r['aprobacion']['confianza_ocr'],'MEDIA')
        self.assertTrue((self.salida/'evidencia'/f'{rid}_cargo.jpg').exists())
        self.assertTrue((self.salida/'evidencia'/f'{rid}_fecha.jpg').exists())
        with self.assertRaises(ValueError):  # misma persona, mismos datos
            p.aprobar(rid,self.datos(),'Ana')
        with contextlib.redirect_stdout(io.StringIO()):
            xlsx,csv_path=p.exportar(doble_revision='no_alta')
        self.assertIsNone(xlsx)
        filas=list(__import__('csv').DictReader(csv_path.open(encoding='utf-8-sig'),delimiter=';'))
        fila=next(f for f in filas if f['id']==rid)
        self.assertEqual(fila['resultado'],'FALTA SEGUNDA REVISION'); self.assertEqual(fila['confianza_ocr'],'MEDIA')
        self.assertIn('_fecha.jpg',fila['evidencia'])
        p.aprobar(rid,self.datos(),'Luis')
        self.assertEqual(len({a['revisor'] for a in p.aprobaciones_vigentes(p.records[rid])}),2)
        self.assertEqual(p.lista_revision('segunda'),[])
        with contextlib.redirect_stdout(io.StringIO()):
            xlsx,csv_path=p.exportar(doble_revision='no_alta')
        self.assertIsNotNone(xlsx)
        actas=sorted(self.salida.glob('acta_exportacion_*.json'))
        acta=json.loads(actas[-1].read_text(encoding='utf-8'))
        self.assertEqual(acta['celdas_escritas'],1); self.assertEqual(acta['excel_generado']['sha256'],c.sha256(xlsx))
        self.assertEqual(acta['detalle'][0]['celda'],'F2'); self.assertEqual(acta['detalle'][0]['revisor'],'Ana | Luis')
        # Una corrección posterior invalida las aprobaciones anteriores.
        corregido=self.datos(); corregido['fecha_notificacion']='16/09/2026'
        p.aprobar(rid,corregido,'Ana')
        self.assertEqual(len(p.aprobaciones_vigentes(p.records[rid])),1)
        self.assertEqual(p.lista_revision('segunda'),[rid])
        with contextlib.redirect_stdout(io.StringIO()):
            xlsx,_=p.exportar(doble_revision=True)
        self.assertIsNone(xlsx)
        with contextlib.redirect_stdout(io.StringIO()):
            xlsx,_=p.exportar(doble_revision=False)
        self.assertIsNotNone(xlsx)
        self.assertEqual(load_workbook(xlsx)['2026']['F2'].value.date(),date(2026,9,16))

class NotebookTests(unittest.TestCase):
    def setUp(self):
        sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
        import build_notebook
        self.builder=build_notebook
        self.path=Path(__file__).resolve().parents[1]/'notebooks/REC_2026_02_SEPTIEMBRE_Colab.ipynb'
    def test_notebook_has_clean_outputs_and_valid_syntax(self):
        nb=json.loads(self.path.read_text(encoding='utf-8'))
        self.assertEqual(nb['nbformat'],4)
        self.assertEqual(nb['metadata']['accelerator'],'GPU')
        self.assertEqual(nb['metadata']['colab']['gpuType'],'T4')
        cells=[x for x in nb['cells'] if x['cell_type']=='code']
        self.assertEqual(len(cells),7)
        for cell in cells:
            self.assertEqual(cell['outputs'],[]);self.assertIsNone(cell['execution_count'])
            source=''.join(cell['source'])
            ast.parse(source)
            self.assertTrue(source.startswith('# @title '),cell['id'])
        self.assertIn('@param',''.join(cells[0]['source']))
        self.assertIn('def configurar(',''.join(cells[1]['source']))
        self.assertIn('DOBLE_REVISION',''.join(cells[5]['source']))
        self.assertIn('LECTOR_VLM',''.join(cells[3]['source']))
    def test_notebook_matches_templates(self):
        nb=json.loads(self.path.read_text(encoding='utf-8'))
        self.assertEqual(nb,self.builder.render(),'Ejecuta python scripts/build_notebook.py y confirma el cuaderno.')
if __name__=='__main__':unittest.main()
