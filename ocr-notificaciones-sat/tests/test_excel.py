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
        with Image.open(src) as im: im.load()
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
        self.assertEqual(c.dates_in_text('16/09/26'),['2026-09-16'])
        self.assertEqual(c.dates_in_text('16 09 26'),['2026-09-16'])
        self.assertEqual(c.dates_in_text('21/9/26 y 22/9/26'),['2026-09-21','2026-09-22'])
        self.assertEqual(c.dates_in_text('160926'),[])
        self.assertEqual(c.dates_in_text('Siendo las 9:00 horas, del dia 11/09/2026'),['2026-09-11'])
        self.assertEqual(c.dates_in_text('15-9-2026 y 03/06/2026'),['2026-06-03','2026-09-15'])
        self.assertEqual(c.dates_in_text('31022026'),[])
        self.assertEqual(c.dates_in_text('N 15 021 000021666'),[])
    def test_nivel_confianza(self):
        v=lambda vista,fechas:{'vista':vista,'fechas':fechas,'texto':''}
        F='2026-09-15'; G='2026-09-16'
        casos=[
            # (variantes, propuesta, nivel esperado, lectores que coinciden)
            ([v('trocr_color',[F]),v('vlm_color',[F])],F,'ALTA',['trocr','vlm']),
            ([v('color',[F]),v('trocr_color',[F]),v('vlm_color',[F])],F,'ALTA',['detector','trocr','vlm']),
            ([v('color',[F]),v('trocr_color',[F])],F,'MEDIA',['detector','trocr']),       # impreso + un manuscrito
            ([v('color',[F]),v('tinta',[F]),v('contraste',[F])],F,'MEDIA',['detector']),   # solo el detector impreso
            ([v('trocr_color',[F]),v('trocr_tinta',[F])],F,'MEDIA',['trocr']),             # un solo manuscrito, dos vistas
            ([v('vlm_color',[F])],F,'MEDIA',['vlm']),
            ([v('color',[F]),v('trocr_color',[F]),v('vlm_color',[])],F,'MEDIA',['detector','trocr']),
            ([v('trocr_color',[F]),v('vlm_color',[G])],None,'CONTRADICTORIA',[]),
            ([v('color',[G]),v('trocr_color',[F]),v('vlm_color',[F])],None,'CONTRADICTORIA',[]),  # el detector contradice
            ([v('trocr_color',[F]),v('vlm_color',[F]),v('acta_negativa',[G])],F,'ALTA',['trocr','vlm']),  # el acta no interviene
            ([v('color',[]),v('trocr_color',[])],None,'SIN_FECHA',[]),
            ([],None,'SIN_FECHA',[]),
        ]
        for variantes,prop,nivel,lectores in casos:
            with self.subTest(nivel=nivel,variantes=[x['vista'] for x in variantes]):
                self.assertEqual(c.nivel_confianza(variantes,prop),(nivel,lectores))
        class M: vlm=None
        self.assertIn('no es alcanzable',c.aviso_lectores(M()))
        M.vlm=object(); self.assertIn('Qwen2.5-VL',c.aviso_lectores(M()))
    def test_alertas_plausibilidad(self):
        original=c.VENTANA_DIAS
        try:
            c.configurar(ventana_dias=30)
            t={'emision':'2026-09-02'}
            self.assertEqual(c.alertas({'fecha_notificacion':'15/09/2026','resolucion':'UNO'},t),[])
            self.assertTrue(any('resoluci' in a for a in c.alertas({'fecha_notificacion':'15/09/2026'},t)))
            avisos=c.alertas({'fecha_notificacion':'15/12/2026','resolucion':'UNO'},t)
            self.assertEqual(len(avisos),1); self.assertIn('104 días',avisos[0])
            self.assertTrue(any('domingo' in a for a in c.alertas({'fecha_notificacion':'06/09/2026','resolucion':'UNO'},t)))
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

def caja(x1,y1,x2,y2,text,score=0.9):
    return {'box':[[x1,y1],[x2,y1],[x2,y2],[x1,y2]],'text':text,'score':score}

class LocalizacionTests(unittest.TestCase):
    W,H=2200,2800
    def test_rotulos_mal_leidos_por_easyocr(self):
        rows=[caja(1000,1900,1500,1930,'FECHA DE HotECACLON:'),caja(1000,1990,1500,2020,'HORADENOIIECACKN:'),
              caja(1000,2080,1500,2110,'FOLIOS ENTREGADOS:'),caja(100,1900,600,1930,'CARACTERISTICAS DEL LUGAR')]
        roi,metodo=c.localizar_recuadro_fecha(rows,self.W,self.H)
        self.assertEqual(metodo,'rotulos_fecha_y_hora')
        self.assertEqual((roi[1],roi[3]),(1930,1990)); self.assertLess(roi[0],1000); self.assertGreater(roi[2],1500)
    def test_rotulo_partido_en_dos_detecciones(self):
        rows=[caja(1000,1900,1180,1930,'FECHA DE'),caja(1190,1901,1500,1931,'NOTIFICACION:'),caja(1000,1990,1500,2020,'HORA DE NOTIFICACION:')]
        roi,metodo=c.localizar_recuadro_fecha(rows,self.W,self.H)
        self.assertEqual(metodo,'rotulos_fecha_y_hora'); self.assertEqual((roi[1],roi[3]),(1931,1990))
    def test_solo_hora_o_nada(self):
        roi,metodo=c.localizar_recuadro_fecha([caja(1000,1990,1500,2020,'HORA DE NOTIFICACION:')],self.W,self.H)
        self.assertEqual(metodo,'rotulo_hora'); self.assertEqual(roi[3],1990); self.assertLess(roi[1],1990)
        roi,metodo=c.localizar_recuadro_fecha([caja(1000,1990,1500,2020,'FECHA Y HORA'),caja(1000,1890,1500,1920,'FOLIOS ENTREGADOS')],self.W,self.H)
        self.assertIsNone(roi)
        # 'HORA DE NOTIFICACION' no debe aceptarse como rótulo FECHA aunque se parezca.
        roi,metodo=c.localizar_recuadro_fecha([caja(1000,1990,1500,2020,'HORA DE NOTIFICACION')],self.W,self.H)
        self.assertEqual(metodo,'rotulo_hora')
    def test_acta_negativa_y_clasificacion(self):
        rows=[caja(100,1700,900,1730,'Siendo las 9:00 horas, del dia 11/09/2026 , quien suscribe'),
              caja(1000,1900,1500,1930,'FECHA DE NOTIFICACION:'),caja(1000,1990,1500,2020,'HORA DE NOTIFICACION:')]
        self.assertEqual(c.fecha_acta_negativa(rows),['2026-09-11'])
        head='EXPEDIENTE COACTIVO : 16-012-000008411-2026-SAT-H/DCC-IPA\nRESOLUCION DE EJECUCION COACTIVA'
        self.assertEqual(c.clasificar_documento(head,rows,self.W,self.H,'16-012-000008411-2026'),'REC')
        self.assertEqual(c.clasificar_documento(head,[],2480,1748,'16-012-000008411-2026'),'REC_SIN_CARGO')
        self.assertEqual(c.clasificar_documento('texto suelto',rows,self.W,self.H,None),'CARGO_SIN_ENCABEZADO')
        self.assertEqual(c.clasificar_documento('AVISO PREVIO DE NOTIFICACION',[],3307,2330,None),'ANEXO')
        self.assertEqual(c.clasificar_documento('CEDULON',[],3307,2330,None),'ANEXO')
        self.assertEqual(c.clasificar_documento('nada',[],self.W,self.H,None),'DESCONOCIDO')
    def test_parse_header_resolucion_tolerante(self):
        exp,code,res=c.parse_header('EXPEDIENTE COACTIVO : 16-012-000008411-2026-SAT-H/DCC-IPA\nCODIGO CONTRIBUYENTE: 000000066789\nResoluci6n N° UNO')
        self.assertEqual((exp,code,res),('16-012-000008411-2026-SAT-H/DCC-IPA','66789','UNO'))
        self.assertEqual(c.parse_header('RESOLUCLON N UNO')[2],'UNO')
        self.assertIsNone(c.parse_header('texto sin resolucion')[2])
        d=dict(expediente='16-012-1-2026',codigo='123',resolucion=None,fecha_notificacion='15/09/2026')
        src=Path(tempfile.mkdtemp())/'o.xlsx'; libro_ficticio(src); index,_,_=c.build_index(src,('2026',))
        self.assertEqual(c.evaluar(d,index)[1],[])
        d['resolucion']='DOS'; self.assertTrue(any('REC UNO' in i for i in c.evaluar(d,index)[1]))

class MotorMediaPagina(MotorFalso):
    """Simula una REC partida en dos fotos: encabezado en una, cargo en la otra."""
    def leer(self,src):
        with Image.open(src) as im: color=im.getpixel((0,0))
        base=dict(expediente=None,codigo=None,resolucion=None,fecha_notificacion=None,fecha_emision=None,fecha_texto=None,observaciones='')
        if color[0]==0:   # f1: encabezado sin cargo
            return dict(base,expediente='16-012-1-2026',codigo='123',resolucion='UNO'),{'motor':self.nombre,'tipo':'REC_SIN_CARGO'}
        if color[0]==40:  # f2: cargo sin encabezado
            return dict(base,fecha_notificacion='15/09/2026'),{'motor':self.nombre,'tipo':'CARGO_SIN_ENCABEZADO'}
        return dict(base),{'motor':self.nombre,'tipo':'ANEXO'}

class EmparejamientoTests(ProyectoBase):
    def test_emparejar_y_medir(self):
        p=self.proyecto()
        with contextlib.redirect_stdout(io.StringIO()):
            p.inventario(); resumen=p.procesar(MotorMediaPagina(),limite=10)
        self.assertEqual(resumen['emparejadas'],0)  # f1 y f2 están en carpetas distintas
        f1=p.ids_de([self.fotos/'A/sub/f1.png'])[0]; f2=p.ids_de([self.fotos/'A/f2.png'])[0]; f3=p.ids_de([self.fotos/'B/f3.png'])[0]
        self.assertEqual(p.records[f3]['estado'],'ANEXO')
        self.assertEqual(sorted(p.lista_revision('pendientes')),sorted([f1,f2])); self.assertIn(f3,p.lista_revision('todos'))
        # f1 y f2 están en carpetas distintas: no se emparejan. Movemos f2 junto a f1 y repetimos.
        self.assertIsNone(p.records[f2]['pareja'])
        p.records[f1]['archivos']=['A/Imagen (149).png']; p.records[f2]['archivos']=['A/Imagen (150).png']
        self.assertEqual(p.emparejar(),1)
        self.assertEqual(p.records[f2]['pareja'],f1); self.assertEqual(p.records[f1]['pareja'],f2)
        self.assertEqual(p.records[f2]['lectura']['expediente'],'16-012-1-2026')
        self.assertEqual(p.records[f2]['destino']['celda'],'F2'); self.assertEqual(p.records[f2]['problemas'],[])
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(p.resumen_alcance(),{'filas':5,'f_vacia':3,'f_registrada':2})
        # Medición contra F ya registrada: la fila 3 (16-012-3) tiene 10/09/2026.
        p.records[f2]['lectura'].update(expediente='16-012-3-2026',fecha_notificacion='10/09/2026')
        p.records[f1]['lectura'].update(fecha_notificacion='11/09/2026',expediente='16-012-3-2026')
        with contextlib.redirect_stdout(io.StringIO()):
            res=p.medir_acierto()
        self.assertEqual((res['con_referencia'],res['aciertos'],res['errores']),(2,1,1))

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
        self.assertIn('langchain',''.join(cells[0]['source']))
        self.assertIn('medir_acierto',''.join(cells[2]['source']))
    def test_notebook_matches_templates(self):
        nb=json.loads(self.path.read_text(encoding='utf-8'))
        self.assertEqual(nb,self.builder.render(),'Ejecuta python scripts/build_notebook.py y confirma el cuaderno.')
if __name__=='__main__':unittest.main()
