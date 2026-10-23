from pathlib import Path
from datetime import datetime, date
from collections import defaultdict, Counter
from zoneinfo import ZoneInfo
import base64, copy, csv, hashlib, html, io, json, os, re, time, unicodedata
import zipfile, posixpath, uuid
from PIL import Image, ImageOps
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.datetime import to_excel
from lxml import etree as ET

VERSION = '2026-09-02-v1'
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
N = {'s': NS}
EXTENSIONES = {'.jpg', '.jpeg', '.png', '.webp', '.tif', '.tiff', '.bmp'}

def ahora():
    return datetime.now(ZoneInfo('America/Lima')).isoformat(timespec='seconds')

def sha256(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()

def atomic_json(path, obj):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    with tmp.open('w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, path)

def norm(value):
    s = unicodedata.normalize('NFKD', str(value or ''))
    return re.sub(r'[^a-z0-9]', '', ''.join(c for c in s if not unicodedata.combining(c)).lower())

def codigo(value):
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = str(value or '').strip()
    return str(int(s)) if re.fullmatch(r'\d+', s) else None

def expediente(value):
    s = str(value or '').strip().upper().replace('–', '-').replace('—', '-')
    s = re.sub(r'\s+', '', s)
    m = re.fullmatch(r'(\d{2})-(\d{3})-(\d{1,12})(?:-(20\d{2})(?:-SAT-H/DCC(?:-[A-Z]+)?)?)?', s)
    if not m:
        return None
    return f'{m[1]}-{m[2]}-{int(m[3])}', int(m[4]) if m[4] else None

def fecha(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value or '').strip()
    for fmt in ('%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%d/%m/%Y %H:%M:%S'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None

def texto_celda(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return value

def build_index(excel, hojas):
    if tuple(hojas) != ('2026',):
        raise ValueError('Solo se permite la hoja 2026.')
    if not zipfile.is_zipfile(excel):
        raise ValueError('El archivo no es un XLSX válido o está incompleto. Ábrelo en Excel y guarda una copia .xlsx antes de continuar.')
    wb = load_workbook(excel, read_only=True, data_only=False, keep_links=True)
    index, resumen = defaultdict(list), []
    try:
        for sheet_name in hojas:
            if sheet_name not in wb.sheetnames or not re.fullmatch(r'20\d{2}', sheet_name):
                raise ValueError(f'Hoja anual inexistente o no válida: {sheet_name}')
            ws = wb[sheet_name]
            # Leer todas las filas, incluso las que quedan fuera del autofiltro.
            rows = ws.iter_rows(values_only=True)
            headers = [norm(v) for v in next(rows)]
            def unique(pred, label):
                found = [i for i, h in enumerate(headers) if pred(h)]
                if len(found) != 1:
                    raise ValueError(f'{sheet_name}: no se identificó una sola columna {label}. Encabezados: {headers}')
                return found[0]
            exp_col = unique(lambda h: h in {'expdiente','expediente','nexpediente','ndeexpediente','nrodeexpediente','ndeexp'}, 'expediente')
            date_col = unique(lambda h: h in {'ffechadenotificacionrec','fechadenotificacionrec','fechanotificacionrec','fechanotifrec','fechadenotificacion'}, 'notificación REC')
            code_col = unique(lambda h: h in {'codsigeti','codigodeobligsigeti'}, 'código SIGETI')
            name_cols = [i for i,h in enumerate(headers) if h in {'nombre','obligado'}]
            emission_cols = [i for i,h in enumerate(headers) if h in {'fechaemisionrec','fechadeemisionderec','fechaemisrec'}]
            if date_col != 5 or emission_cols != [4]:
                raise ValueError('Se esperaba emisión REC en E y notificación REC en F.')
            count = 0
            for row_num, row in enumerate(rows, 2):
                def at(i): return row[i] if i < len(row) else None
                parsed = expediente(at(exp_col))
                if not parsed:
                    continue
                key, explicit_year = parsed
                year = explicit_year or int(sheet_name)
                if year != int(sheet_name):
                    continue
                item = dict(hoja=sheet_name, fila=row_num, celda=f'{get_column_letter(date_col+1)}{row_num}',
                            expediente=str(at(exp_col)), codigo=codigo(at(code_col)),
                            nombre=str(at(name_cols[0]) or '') if name_cols else '',
                            fecha_actual=texto_celda(at(date_col)),
                            emision=texto_celda(at(emission_cols[0])) if emission_cols else None)
                index[(key,year)].append(item)
                count += 1
            resumen.append({'hoja':sheet_name,'columna_fecha':get_column_letter(date_col+1),'expedientes':count})
        epoch = wb.epoch
    finally:
        wb.close()
    return dict(index), resumen, epoch

def normalizar_lectura(data):
    if not isinstance(data, dict):
        raise ValueError('La respuesta no es un objeto JSON.')
    fields = ['expediente','codigo','resolucion','fecha_notificacion','fecha_emision','fecha_texto','observaciones']
    out = {}
    for field in fields:
        v = data.get(field)
        if v is not None and not isinstance(v, str):
            raise ValueError(f'Campo {field} debe ser texto o null.')
        out[field] = v.strip() if isinstance(v, str) else None
    return out

def evaluar(data, index):
    issues, target = [], None
    parsed = expediente(data.get('expediente'))
    if not parsed or parsed[1] is None:
        issues.append('Expediente completo y año ilegibles o incompletos.')
    else:
        matches = index.get(parsed, [])
        if len(matches) != 1:
            issues.append(f'Coincidencias de expediente/año en Excel: {len(matches)}. Debe existir exactamente una.')
        else:
            target = matches[0]
            if target['hoja'] != '2026' or fecha(target['emision']) != date(2026,9,2):
                issues.append('Fuera del alcance: solo 2026 con E = 02/09/2026.')
            if target['fecha_actual'] not in (None, ''):
                issues.append('La notificación ya está registrada. Se conserva sin modificar.')
            if not codigo(data.get('codigo')) or codigo(data.get('codigo')) != target['codigo']:
                issues.append('El código del contribuyente no coincide con el Excel o no se puede leer.')
    if norm(data.get('resolucion')) not in {'uno','1','01','resolucionnuno','resolucionuno'}:
        issues.append('Esta versión solo llena la notificación de la REC UNO. Otra resolución requiere revisión aparte.')
    d = fecha(data.get('fecha_notificacion'))
    if d is None:
        issues.append('Fecha de notificación incompleta, ilegible o inválida.')
    else:
        if d > datetime.now(ZoneInfo('America/Lima')).date():
            issues.append('La fecha de notificación está en el futuro.')
        em = fecha(data.get('fecha_emision'))
        excel_em = fecha(target['emision']) if target else None
        if (em and d < em) or (excel_em and d < excel_em):
            issues.append('La notificación es anterior a la emisión de la REC.')
        if target and target['fecha_actual'] not in (None, '') and fecha(target['fecha_actual']) != d:
            issues.append('La celda ya contiene otra fecha o una fórmula. No se sobrescribe.')
    return target, issues

def parse_header(text):
    # No sustituir O por 0 ni completar dígitos: un identificador dudoso se revisa.
    clean = text.upper().replace('–','-').replace('—','-')
    pattern = r'\b(\d{2}\s*-\s*\d{3}\s*-\s*\d{1,12}\s*-\s*20\d{2})(?:\s*-\s*SAT\s*-\s*H\s*/\s*DCC(?:\s*-\s*[A-Z]+)?)?'
    found = {re.sub(r'\s+','',m.group(0)) for m in re.finditer(pattern,clean)}
    exp = next(iter(found)) if len(found)==1 else None
    codes = re.findall(r'(?:CONTRIBUYENTE|SIGETI)\s*[:.;°ºN\s]*([0-9]{3,15})\b',clean)
    codes = {codigo(c) for c in codes}
    code = next(iter(codes)) if len(codes)==1 else None
    resolution = re.findall(r'RESOLUCI[OÓ]N\s*(?:N[°ºO.]?\s*)?(UNO|DOS|TRES|\d+)\b',clean)
    return exp,code,resolution[0] if len(set(resolution))==1 else None

def dates_in_text(text):
    matches = re.findall(r'(?<!\d)(\d{1,2})\s*[/ .-]\s*(\d{1,2})\s*[/ .-]\s*(20\d{2})(?!\d)',text)
    out = set()
    for day,month,year in matches:
        try: out.add(date(int(year),int(month),int(day)).isoformat())
        except ValueError: pass
    return sorted(out)

def seleccionar_muestra(paths, raiz, cantidad=30, semilla=2026):
    """Muestra reproducible repartida por carpeta principal y subcarpeta final."""
    import random
    from collections import deque
    if cantidad < 1:
        raise ValueError('La muestra debe contener al menos una foto.')
    raiz=Path(raiz).resolve()
    grupos=defaultdict(lambda:defaultdict(list))
    for p in sorted(set(Path(p).resolve() for p in paths)):
        rel=p.relative_to(raiz)
        principal=rel.parts[0] if len(rel.parts)>1 else '(raiz)'
        grupos[principal][rel.parent.as_posix()].append(p)
    rng=random.Random(semilla)
    principales=list(sorted(grupos));rng.shuffle(principales)
    colas={}
    for principal in principales:
        hojas=list(sorted(grupos[principal]));rng.shuffle(hojas)
        filas=[]
        for hoja in hojas:
            fotos=grupos[principal][hoja];rng.shuffle(fotos)
            filas.append(deque(fotos))
        colas[principal]=deque(filas)
    turnos=deque(principales); seleccion=[]
    while turnos and len(seleccion)<cantidad:
        principal=turnos.popleft()
        fila=colas[principal].popleft()
        seleccion.append(fila.popleft())
        if fila: colas[principal].append(fila)
        if colas[principal]: turnos.append(principal)
    return seleccion


class MotorOCR:
    """PaddleOCR + TrOCR; EasyOCR y DeepSeek local como alternativas."""
    def __init__(self, motor='paddle_trocr', gpu=False, trocr_model='microsoft/trocr-base-handwritten'):
        if motor not in ('paddle_trocr','easyocr','deepseek_local'):
            raise ValueError('Motor desconocido: '+str(motor))
        self.paddle=None
        self.trocr=None
        if motor=='paddle_trocr':
            from paddleocr import PaddleOCR
            import torch
            from transformers import TrOCRProcessor, VisionEncoderDecoderModel
            # CPU para Paddle evita conflictos entre sus bibliotecas CUDA y PyTorch.
            self.paddle=PaddleOCR(lang='es',ocr_version='PP-OCRv5',device='cpu',
                use_doc_orientation_classify=False,use_doc_unwarping=False,
                use_textline_orientation=False,enable_mkldnn=False)
            self.trocr_device='cuda' if gpu and torch.cuda.is_available() else 'cpu'
            self.processor=TrOCRProcessor.from_pretrained(trocr_model)
            self.trocr=VisionEncoderDecoderModel.from_pretrained(trocr_model).to(self.trocr_device).eval()
        else:
            import easyocr
            self.reader = easyocr.Reader(['es','en'],gpu=gpu if motor=='easyocr' else False)
        self.nombre='EasyOCR local'
        self.ds_model=None
        self.ds_tokenizer=None
        if motor=='paddle_trocr':
            self.nombre='PaddleOCR PP-OCRv5 + '+trocr_model+' ('+self.trocr_device+')'
        if motor=='deepseek_local':
            import torch
            if not torch.cuda.is_available() or not torch.cuda.is_bf16_supported():
                raise RuntimeError('DeepSeek-OCR-2 requiere en esta integración una GPU CUDA con bfloat16 (p. ej. L4/A100). Usa MOTOR="easyocr" para CPU o T4.')
            from transformers import AutoModel, AutoTokenizer
            model_name='deepseek-ai/DeepSeek-OCR-2'
            self.ds_tokenizer=AutoTokenizer.from_pretrained(model_name,trust_remote_code=True)
            self.ds_model=AutoModel.from_pretrained(model_name,trust_remote_code=True,
                use_safetensors=True,_attn_implementation='flash_attention_2').eval().cuda().to(torch.bfloat16)
            self.nombre='EasyOCR (encabezado) + DeepSeek-OCR-2 local (fecha)'

    def read(self,im,digits=False):
        import numpy as np
        if self.paddle is not None:
            # Paddle recibe matrices BGR; PIL entrega RGB.
            output=[]
            for result in self.paddle.predict(np.array(im.convert('RGB'))[:,:,::-1].copy()):
                for box,text,score in zip(result['rec_polys'],result['rec_texts'],result['rec_scores']):
                    output.append({'box':[[float(x),float(y)] for x,y in box],
                                   'text':str(text),'score':float(score)})
            return output
        kwargs={'detail':1,'paragraph':False,'batch_size':1}
        if digits: kwargs['allowlist']='0123456789/.- '
        rows=self.reader.readtext(np.array(im.convert('RGB')),**kwargs)
        return [{'box':[[float(x),float(y)] for x,y in box], 'text':str(text),'score':float(score)} for box,text,score in rows]

    def read_handwritten(self,im):
        import torch
        pixels=self.processor(images=im.convert('RGB'),return_tensors='pt').pixel_values.to(self.trocr_device)
        with torch.inference_mode():
            ids=self.trocr.generate(pixels,max_new_tokens=32,num_beams=4,do_sample=False)
        return self.processor.batch_decode(ids,skip_special_tokens=True)[0].strip()

    def leer(self,path):
        with Image.open(path) as src:
            if getattr(src,'n_frames',1)>1:
                raise ValueError('Imagen de varias páginas: sepárala antes de procesar.')
            im=ImageOps.exif_transpose(src).convert('RGB')
        im.thumbnail((2200,2800),Image.Resampling.LANCZOS)
        w,h=im.size
        head=self.read(im.crop((0,0,w,int(h*.32))))
        head_text='\n'.join(r['text'] for r in head)
        exp,code,res=parse_header(head_text)
        offset=int(h*.58)
        foot=self.read(im.crop((0,offset,w,h)))
        anchors=[r for r in foot if 'fecha' in norm(r['text']) and 'notif' in norm(r['text'])]
        data=dict(expediente=exp,codigo=code,resolucion=res,fecha_notificacion=None,
                  fecha_emision=None,fecha_texto=None,observaciones='')
        evidence={'encabezado':head,'pie':foot,'roi_fecha':None,'variantes_fecha':[],
                  'motor':self.nombre,'aviso':'Puntuaciones OCR no calibradas; no son probabilidades de acierto.'}
        if len(anchors)!=1:
            data['observaciones']='No se localizó un único rótulo FECHA DE NOTIFICACIÓN. Completar manualmente desde la foto.'
            return data,evidence
        box=anchors[0]['box']
        x1=min(x for x,y in box); x2=max(x for x,y in box)
        y1=min(y for x,y in box)+offset; y2=max(y for x,y in box)+offset
        lh=max(10,y2-y1)
        bottom=min(h,y2+max(32,lh*2.6))
        for r in foot:
            if 'hora' in norm(r['text']) and 'notif' in norm(r['text']):
                ry=min(y for x,y in r['box'])+offset
                if y2<ry<bottom: bottom=ry
        roi=(max(0,int(x1-25)),max(0,int(y2)),min(w,int(x2+30)),int(bottom))
        if roi[3]-roi[1]<8:
            data['observaciones']='El recorte de fecha es demasiado pequeño. Revisar manualmente.'
            return data,evidence
        evidence['roi_fecha']=[round(roi[0]/w,6),round(roi[1]/h,6),round(roi[2]/w,6),round(roi[3]/h,6)]
        crop=im.crop(roi)
        crop=crop.resize((crop.width*2,crop.height*2),Image.Resampling.LANCZOS)
        candidates=[]
        for label,view in [('color',crop),('contraste',ImageOps.autocontrast(ImageOps.grayscale(crop)).convert('RGB'))]:
            rows=self.read(view,digits=True)
            text=' '.join(r['text'] for r in sorted(rows,key=lambda r:min(x for x,y in r['box'])))
            ds=dates_in_text(text)
            evidence['variantes_fecha'].append({'vista':label,'texto':text,'fechas':ds,'detecciones':rows})
            candidates.append(ds)
            if self.trocr is not None:
                handwritten=self.read_handwritten(view)
                hd=dates_in_text(handwritten)
                evidence['variantes_fecha'].append({'vista':'trocr_'+label,'texto':handwritten,'fechas':hd})
                candidates.append(hd)
        if self.ds_model is not None:
            import tempfile, contextlib
            with tempfile.TemporaryDirectory() as td:
                image_path=Path(td)/'fecha.jpg'; crop.save(image_path)
                # eval_mode retorna el texto; no crea dibujos ni ejecuta el parser de grounding.
                with contextlib.redirect_stdout(io.StringIO()):
                    result=self.ds_model.infer(self.ds_tokenizer,prompt='<image>\nFree OCR.',
                        image_file=str(image_path),output_path=td,base_size=1024,image_size=768,
                        crop_mode=True,save_results=False,eval_mode=True)
                text=result if isinstance(result,str) else ''
                ds=dates_in_text(text)
                evidence['variantes_fecha'].append({'vista':'deepseek_local','texto':text,'fechas':ds})
                candidates.append(ds)
        nonempty={x for seq in candidates for x in seq}
        # Una lectura es una propuesta; todas requieren confirmación visual.
        if len(nonempty)==1:
            data['fecha_notificacion']=next(iter(nonempty))
            data['fecha_texto']=' | '.join(v['texto'] for v in evidence['variantes_fecha'])
            data['observaciones']='Fecha propuesta por OCR. Cotejar visualmente; la coincidencia entre variantes no garantiza exactitud.'
        elif len(nonempty)>1:
            data['observaciones']='Lecturas de fecha contradictorias: '+', '.join(sorted(nonempty))
        else:
            data['observaciones']='El OCR no obtuvo una fecha completa. Introducirla desde el recorte; no se adivina el año.'
        return data,evidence


class Proyecto:
    def __init__(self, fotos, excel, salida, hojas=('2026',)):
        self.fotos, self.excel, self.salida = Path(fotos).expanduser().resolve(), Path(excel).expanduser().resolve(), Path(salida).expanduser().resolve()
        if not self.fotos.is_dir():
            raise ValueError(f'No existe la carpeta de fotos: {self.fotos}')
        if not self.excel.is_file() or self.excel.suffix.lower() != '.xlsx':
            raise ValueError('Selecciona un archivo .xlsx existente.')
        if self.salida == self.fotos or self.fotos in self.salida.parents:
            raise ValueError('La salida debe estar fuera de la carpeta de fotografías.')
        self.salida.mkdir(parents=True,exist_ok=True)
        self.records_dir = self.salida/'registros'
        self.records_dir.mkdir(exist_ok=True)
        self.fingerprint = sha256(self.excel)
        cfg_path = self.salida/'proyecto.json'
        cfg = {'version':VERSION,'excel_sha256':self.fingerprint,'hojas':list(hojas)}
        if cfg_path.exists() and json.loads(cfg_path.read_text(encoding='utf-8')) != cfg:
            raise ValueError('Esta carpeta de resultados corresponde a otro Excel o configuración. Elige otra carpeta de salida.')
        atomic_json(cfg_path,cfg)
        self.index, self.resumen_hojas, self.epoch = build_index(self.excel,hojas)
        self.records = {}
        for p in self.records_dir.glob('*.json'):
            try:
                r = json.loads(p.read_text(encoding='utf-8'))
                self.records[r['id']] = r
            except Exception:
                raise ValueError(f'Registro de avance dañado: {p}. No se ignora para conservar la trazabilidad.') from None

    def save(self,r):
        atomic_json(self.records_dir/f"{r['id']}.json",r)
        self.records[r['id']] = r

    def inventario(self, seleccion=None):
        paths = sorted(p for p in self.fotos.rglob('*') if p.is_file() and p.suffix.lower() in EXTENSIONES and '__MACOSX' not in p.parts) if seleccion is None else sorted(set(Path(p).resolve() for p in seleccion))
        for p in paths:
            if self.fotos not in p.resolve().parents or not p.is_file() or p.suffix.lower() not in EXTENSIONES:
                raise ValueError('Archivo seleccionado fuera de la carpeta de fotos o no admitido: '+str(p))
        active = defaultdict(list)
        total_bytes = 0
        for i,p in enumerate(paths,1):
            digest = sha256(p)
            active[digest].append(p.relative_to(self.fotos).as_posix())
            total_bytes += p.stat().st_size
            if i % 250 == 0:
                print(f'Inventario: {i}/{len(paths)} fotos')
        for digest, aliases in active.items():
            r = self.records.get(digest,{'id':digest,'estado':'NUEVO','lectura':None,'aprobacion':None})
            r['archivos'] = aliases
            self.save(r)
        self.active = set(active)
        print(f'{len(paths)} fotos; {len(active)} contenidos únicos; {len(paths)-len(active)} copias idénticas; {total_bytes/1e9:.2f} GB.')
        print('Hojas:', self.resumen_hojas)
        self.estado()

    def path(self,r):
        for name in r['archivos']:
            p = (self.fotos/name).resolve()
            if self.fotos in p.parents and p.is_file():
                return p
        raise FileNotFoundError('La fotografía no está en la carpeta actual.')

    def estado(self):
        counts = Counter(self.records[k]['estado'] for k in getattr(self,'active',self.records))
        print('Estados:',dict(counts))
        return counts

    def procesar(self,motor,limite=100,reintentar_errores=False,reprocesar_pendientes=False):
        if not hasattr(self,'active'):
            raise ValueError('Ejecuta primero el inventario.')
        if limite < 1: raise ValueError('El límite debe ser mayor que cero.')
        eligible={'NUEVO'} | ({'ERROR'} if reintentar_errores else set())
        pending=[self.records[k] for k in sorted(self.active)
                 if self.records[k]['estado'] in eligible or
                 (reprocesar_pendientes and self.records[k]['estado'] in {'POR_REVISAR','PENDIENTE'}
                  and self.records[k].get('motor')!=motor.nombre)][:limite]
        for i,r in enumerate(pending,1):
            try:
                p=self.path(r)
                if sha256(p)!=r['id']: raise ValueError('La foto cambió. Repite el inventario.')
                d,evidence=motor.leer(p)
                target,issues=evaluar(d,self.index)
                if r.get('lectura') is not None:
                    r.setdefault('historial_ocr',[]).append({k:r.get(k) for k in
                        ('lectura','evidencia_ocr','motor','leido_en','estado')})
                r.update(lectura=d,estado='POR_REVISAR',destino=target,problemas=issues,
                         evidencia_ocr=evidence,motor=motor.nombre,leido_en=ahora(),error=None)
                self.save(r)
                print(f'{i}/{len(pending)}: {p.name} — '+('revisar discrepancias' if issues else 'confirmar visualmente'))
            except KeyboardInterrupt:
                print('Detenido. El avance anterior está guardado.')
                break
            except Exception as e:
                r.update(estado='ERROR',error=str(e)[:500]); self.save(r)
                print(f'{i}/{len(pending)}: error registrado en {r["archivos"][0]}')
                if 'out of memory' in str(e).lower():
                    print('Memoria insuficiente. Detenido; reinicia el entorno o usa CPU con EasyOCR.')
                    break
        self.estado()

    def aprobar(self,rid,data,revisor):
        if not revisor.strip():
            raise ValueError('Escribe el nombre o iniciales de quien revisa.')
        r = self.records[rid]
        if sha256(self.path(r)) != rid:
            raise ValueError('La imagen cambió. Ejecuta el inventario otra vez.')
        data = normalizar_lectura(data)
        target,issues = evaluar(data,self.index)
        if issues:
            raise ValueError('\n'.join(issues))
        if not target:
            raise ValueError('Falta el destino.')
        previous = r.get('aprobacion')
        if previous:
            r.setdefault('historial_aprobaciones',[]).append(previous)
        r['aprobacion'] = {'datos':data,'destino':target,'revisor':revisor.strip(),'fecha':ahora()}
        r['estado'] = 'APROBADO'
        self.save(r)

    def revision_manual(self, limite=100):
        """Permite trabajar aun sin modelos descargados ni OCR disponible."""
        pending=[self.records[k] for k in sorted(self.active) if self.records[k]['estado']=='NUEVO'][:limite]
        for r in pending:
            r.update(estado='PENDIENTE',lectura=None)
            self.save(r)
        print(f'{len(pending)} fotografías preparadas para transcripción visual.')

    def revisar(self):
        import ipywidgets as W
        from IPython.display import display, clear_output
        ids = sorted(getattr(self,'active',self.records), key=lambda k:(self.records[k]['archivos'][0]))
        ids = [k for k in ids if self.records[k]['estado'] in {'POR_REVISAR','PENDIENTE','ERROR'}]
        def prioridad(k):
            d=self.records[k].get('lectura') or {}
            t,_=evaluar(d,self.index)
            return 0 if t and fecha(t['emision'])==date(2026,9,2) and t['fecha_actual'] in (None,'') else 1
        ids.sort(key=prioridad)
        if not ids:
            print('Todavía no hay lecturas. Ejecuta el procesamiento primero.')
            return
        choice = W.Dropdown(options=[(f'{self.records[k]["estado"]} | {self.records[k]["archivos"][0]}',k) for k in ids],layout=W.Layout(width='98%'))
        output, msg = W.Output(), W.Output()
        exp = W.Text(description='Expediente:',layout=W.Layout(width='98%'),style={'description_width':'140px'})
        code = W.Text(description='Código:',style={'description_width':'140px'})
        res = W.Text(description='Resolución:',style={'description_width':'140px'})
        dt = W.Text(description='Notificación:',placeholder='DD/MM/AAAA',style={'description_width':'140px'})
        em = W.Text(description='Emisión REC:',placeholder='DD/MM/AAAA',style={'description_width':'140px'})
        who = W.Text(description='Revisor:',style={'description_width':'140px'})
        check = W.Checkbox(value=False,description='He cotejado expediente, código, resolución y fecha con la foto.',indent=False,layout=W.Layout(width='98%'))
        approve = W.Button(description='Confirmar y guardar',button_style='success')
        skip = W.Button(description='Dejar pendiente')
        next_btn = W.Button(description='Siguiente')
        rotate = W.Button(description='Girar vista 90°')
        full = W.Checkbox(value=False,description='Mostrar página completa')
        angle = [0]
        def show_image():
            with output:
                clear_output(wait=True)
                r = self.records[choice.value]
                with Image.open(self.path(r)) as src:
                    im = ImageOps.exif_transpose(src).convert('RGB').rotate(angle[0],expand=True)
                w,h = im.size
                for title,view in [('Encabezado',im.crop((0,0,w,int(h*.30)))),('Cargo de notificación',im.crop((0,int(h*.62),w,h)))]:
                    print(title)
                    view.thumbnail((1100,1000))
                    display(view)
                roi=r.get('evidencia_ocr',{}).get('roi_fecha')
                if roi and angle[0]==0:
                    crop=im.crop((int(roi[0]*w),int(roi[1]*h),int(roi[2]*w),int(roi[3]*h)))
                    crop=crop.resize((crop.width*2,crop.height*2),Image.Resampling.LANCZOS)
                    crop.thumbnail((1100,350))
                    print('Recorte que leyó el OCR: comprueba que contenga la fecha correcta.')
                    display(crop)
                if full.value:
                    im.thumbnail((1100,1700)); display(im)
                data = r.get('aprobacion',{}).get('datos') if r.get('aprobacion') else r.get('lectura')
                if data:
                    target,issues = evaluar(data,self.index)
                    print('Destino:',target)
                    print('Observaciones de lectura:',data.get('observaciones') or 'Ninguna')
                    print('Comprobaciones:',issues or 'Coinciden; falta tu comprobación visual.')
                elif r.get('error'):
                    print('Error:',r['error'])
        def load(change=None):
            r = self.records[choice.value]
            d = (r['aprobacion']['datos'] if r.get('aprobacion') else r.get('lectura')) or {}
            for widget,field in [(exp,'expediente'),(code,'codigo'),(res,'resolucion'),(dt,'fecha_notificacion'),(em,'fecha_emision')]:
                widget.value = d.get(field) or ''
            angle[0]=0; check.value=False
            with msg: clear_output()
            show_image()
        def advance(_=None):
            idx = ids.index(choice.value)
            if idx+1 < len(ids): choice.value = ids[idx+1]
            else:
                with msg: print('Fin de la lista. Puedes seleccionar otro registro arriba.')
        def confirm(_):
            with msg:
                clear_output()
                try:
                    if not check.value: raise ValueError('Marca la casilla después de cotejar la fotografía.')
                    r = self.records[choice.value]
                    d = dict(r.get('lectura') or {})
                    d.update(expediente=exp.value,codigo=code.value,resolucion=res.value,
                             fecha_notificacion=dt.value,fecha_emision=em.value)
                    self.aprobar(choice.value,d,who.value)
                    print('Confirmado y guardado. Pulsa Siguiente para continuar.')
                    check.value=False
                except Exception as e: print(str(e))
        def postpone(_):
            r=self.records[choice.value]
            if r.get('aprobacion'): r.setdefault('historial_aprobaciones',[]).append(r['aprobacion'])
            r.update(estado='PENDIENTE',aprobacion=None)
            self.save(r)
            with msg: clear_output(); print('Pendiente: esta foto no se exportará al Excel.')
        def turn(_):
            angle[0]=(angle[0]+90)%360
            show_image()
        choice.observe(load,names='value'); approve.on_click(confirm); skip.on_click(postpone)
        next_btn.on_click(advance); rotate.on_click(turn); full.observe(lambda _:show_image(),names='value')
        display(W.VBox([choice,full,rotate,output,exp,code,res,dt,em,who,check,W.HBox([approve,skip,next_btn]),msg]))
        load()

    def exportar(self):
        if sha256(self.excel) != self.fingerprint:
            raise ValueError('El Excel original cambió. Usa una carpeta de resultados nueva y vuelve a verificar las coincidencias.')
        proposals = defaultdict(list)
        report = []
        for rid in sorted(getattr(self,'active',self.records)):
            r = self.records[rid]
            a = r.get('aprobacion') if r['estado']=='APROBADO' else None
            d = (a or {}).get('datos') or r.get('lectura') or {}
            target,issues = evaluar(d,self.index)
            entry = {'id':rid,'archivos':' | '.join(r['archivos']),'estado':r['estado'],
                     'expediente':d.get('expediente'),'codigo':d.get('codigo'),'fecha':d.get('fecha_notificacion'),
                     'hoja':target['hoja'] if target else '', 'celda':target['celda'] if target else '',
                     'revisor':a['revisor'] if a else '', 'resultado':'NO EXPORTADO',
                     'detalle':'; '.join(issues) or r.get('error') or ''}
            report.append(entry)
            if a and not issues and target:
                if sha256(self.path(r)) != rid:
                    entry['detalle']='La fotografía cambió desde la aprobación.'
                    continue
                proposals[(target['hoja'],target['celda'])].append((fecha(d['fecha_notificacion']),entry))
        updates = {}
        for key, items in proposals.items():
            ds = {d for d,_ in items}
            if len(ds)>1:
                for _,entry in items:
                    entry.update(resultado='CONFLICTO',detalle='Dos fotografías aprobadas proponen fechas distintas para la misma celda.')
                continue
            updates[key]=next(iter(ds))
            for _,entry in items: entry['resultado']='INCLUIDO'
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid.uuid4().hex[:6]
        report_path = self.salida/f'revision_{stamp}.csv'
        columns = ['id','archivos','estado','expediente','codigo','fecha','hoja','celda','revisor','resultado','detalle']
        def safe(v):
            s = str(v if v is not None else '')
            return "'"+s if s.startswith(('=','+','-','@','\t','\r')) else s
        with report_path.open('w',newline='',encoding='utf-8-sig') as f:
            writer=csv.DictWriter(f,fieldnames=columns,delimiter=';')
            writer.writeheader()
            writer.writerows({k:safe(v) for k,v in r.items()} for r in report)
        if not updates:
            print('Todavía no hay fechas aprobadas sin conflictos. Informe:',report_path)
            return None,report_path
        output = self.salida/f'expedientes_actualizados_{stamp}.xlsx'
        patch_xlsx(self.excel,output,updates,self.epoch)
        print(f'Excel generado: {len(updates)} celdas confirmadas. Las demás conservan su valor anterior.')
        print(output); print(report_path)
        return output, report_path

    def paquete_revision(self, limite=20):
        """Exporta evidencia y lecturas para adjuntarlas a ChatGPT; no llama ninguna API."""
        ids=[k for k in sorted(getattr(self,'active',self.records)) if self.records[k]['estado'] in {'POR_REVISAR','PENDIENTE','ERROR'}][:limite]
        if not ids: print('No hay casos pendientes.'); return None
        stamp=datetime.now().strftime('%Y%m%d_%H%M%S')
        dest=self.salida/f'revision_chatgpt_{stamp}.zip'
        manifest=[]
        with zipfile.ZipFile(dest,'w',zipfile.ZIP_DEFLATED) as z:
            for i,k in enumerate(ids,1):
                r=self.records[k]
                with Image.open(self.path(r)) as src:
                    im=ImageOps.exif_transpose(src).convert('RGB')
                w,h=im.size
                for name,view in [('pagina',im),('encabezado',im.crop((0,0,w,int(h*.30)))),('cargo',im.crop((0,int(h*.62),w,h)))]:
                    view.thumbnail((2000,2400))
                    b=io.BytesIO(); view.save(b,'JPEG',quality=92)
                    z.writestr(f'{i:03d}_{name}.jpg',b.getvalue())
                manifest.append({'caso':i,'id':k,'archivos':r['archivos'],'lectura':r.get('lectura'),'problemas':r.get('problemas',[])})
            z.writestr('lecturas.json',json.dumps(manifest,ensure_ascii=False,indent=2))
            z.writestr('INSTRUCCIONES.txt','Compara cada lectura con la página, encabezado y cargo. Verifica expediente, código y fecha manuscrita de notificación. No completes dígitos ilegibles. Devuelve una tabla con caso, expediente, código, fecha y observaciones. Las correcciones se confirman manualmente en el cuaderno antes de exportar.')
        print(dest)
        return dest

def _patch_xlsx(source,output,updates,epoch):
    """Modifica únicamente las celdas autorizadas y sus formatos. Conserva las otras partes ZIP."""
    parser=ET.XMLParser(resolve_entities=False,no_network=True)
    def parse(b): return ET.fromstring(b,parser)
    def serial(root): return ET.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
    with zipfile.ZipFile(source) as zin:
        book=parse(zin.read('xl/workbook.xml'))
        rels=parse(zin.read('xl/_rels/workbook.xml.rels'))
        targets={e.get('Id'):posixpath.normpath(posixpath.join('xl',e.get('Target'))) if not e.get('Target').startswith('/') else e.get('Target').lstrip('/') for e in rels}
        sheet_paths={e.get('name'):targets[e.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')] for e in book.find('s:sheets',N)}
        style=parse(zin.read('xl/styles.xml'))
        xfs=style.find('s:cellXfs',N)
        formats=style.find('s:numFmts',N)
        if formats is None:
            formats=ET.Element('{'+NS+'}numFmts',count='0'); style.insert(0,formats)
        fmt_id=max([163]+[int(e.get('numFmtId')) for e in formats])+1
        ET.SubElement(formats,'{'+NS+'}numFmt',numFmtId=str(fmt_id),formatCode='dd/mm/yyyy')
        formats.set('count',str(len(formats)))
        style_cache={}
        def date_style(old):
            if old not in style_cache:
                xf=copy.deepcopy(xfs[old]); xf.set('numFmtId',str(fmt_id)); xf.set('applyNumberFormat','1')
                style_cache[old]=len(xfs); xfs.append(xf)
            return str(style_cache[old])
        changed={}
        grouped=defaultdict(dict)
        for (sheet,cell),d in updates.items(): grouped[sheet][cell]=d
        for sheet,cells in grouped.items():
            root=parse(zin.read(sheet_paths[sheet]))
            rows={int(r.get('r')):r for r in root.find('s:sheetData',N)}
            merged=root.find('s:mergeCells',N)
            if merged is not None:
                from openpyxl.utils.cell import range_boundaries
                for addr in cells:
                    col,row=re.fullmatch(r'([A-Z]+)(\d+)',addr).groups(); ci=column_index_from_string(col); ri=int(row)
                    for m in merged:
                        a,b,c,d=range_boundaries(m.get('ref'))
                        if a<=ci<=c and b<=ri<=d: raise ValueError('Celda de destino combinada: '+addr)
            for addr,d in cells.items():
                col,row_num=re.fullmatch(r'([A-Z]+)(\d+)',addr).groups()
                row=rows[int(row_num)]
                cell=next((c for c in row if c.get('r')==addr),None)
                if cell is None:
                    cell=ET.Element('{'+NS+'}c',r=addr)
                    ci=column_index_from_string(col)
                    pos=next((i for i,c in enumerate(row) if column_index_from_string(re.match('[A-Z]+',c.get('r')).group())>ci),len(row))
                    row.insert(pos,cell)
                if cell.find('s:f',N) is not None: raise ValueError('No se reemplaza una fórmula: '+addr)
                cell.set('s',date_style(int(cell.get('s','0')))); cell.set('t','n')
                for child in list(cell):
                    if child.tag in {'{'+NS+'}v','{'+NS+'}is'}: cell.remove(child)
                ET.SubElement(cell,'{'+NS+'}v').text=str(to_excel(datetime.combine(d,datetime.min.time()),epoch))
            changed[sheet_paths[sheet]]=serial(root)
        xfs.set('count',str(len(xfs)))
        changed['xl/styles.xml']=serial(style)
        tmp=Path(str(output)+'.tmp')
        try:
            with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist(): zout.writestr(info,changed.get(info.filename,zin.read(info.filename)))
            with zipfile.ZipFile(tmp) as check:
                bad=check.testzip()
                if bad: raise ValueError('Error ZIP: '+bad)
                for sheet,cells in grouped.items():
                    root=parse(check.read(sheet_paths[sheet]))
                    vals={c.get('r'):c.findtext('s:v',namespaces=N) for c in root.findall('s:sheetData/s:row/s:c',N) if c.get('r') in cells}
                    for addr,d in cells.items():
                        if float(vals[addr])!=to_excel(datetime.combine(d,datetime.min.time()),epoch):
                            raise ValueError('La verificación de exportación falló: '+addr)
            os.replace(tmp,output)
        finally:
            tmp.unlink(missing_ok=True)


def patch_xlsx(source, output, updates, epoch):
    # Segunda barrera: incluso una llamada directa debe respetar E, F y hoja.
    if Path(source).resolve() == Path(output).resolve():
        raise ValueError('La salida debe ser una copia distinta del original.')
    index, _, original_epoch = build_index(source, ('2026',))
    allowed = {(r['hoja'],r['celda']):r for items in index.values() if len(items)==1
               for r in items if fecha(r['emision'])==date(2026,9,2)
               and r['fecha_actual'] in (None,'')}
    for key,d in updates.items():
        if key not in allowed:
            raise ValueError('Celda fuera del alcance o ya ocupada: '+str(key))
        if not isinstance(d,date) or not date(2026,9,2)<=d<=datetime.now(ZoneInfo('America/Lima')).date():
            raise ValueError('Fecha de notificación inválida: '+str(d))
    if not updates:
        raise ValueError('No hay fechas confirmadas para escribir.')
    _patch_xlsx(source,output,updates,original_epoch)
