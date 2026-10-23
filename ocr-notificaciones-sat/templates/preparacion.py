# 1. Conectar Drive y comprobar que la subida coincide con el inventario.
from pathlib import Path, PureWindowsPath
import os, sys, subprocess

RUTA_FOTOS = '/content/drive/MyDrive/01_SAT_Huamanga/PROYECTO REC 1'
RUTA_RESULTADOS = '/content/drive/MyDrive/01_SAT_Huamanga/REC_2026_02SEP_RESULTADOS'
FOTOS_ESPERADAS = 3996
CONFIGURACION_LISTA = False

if os.name == 'nt':
    raise RuntimeError('Sigues conectado a tu PC. Selecciona un entorno alojado de Colab y vuelve a ejecutar.')
from google.colab import drive
drive.mount('/content/drive')
raiz = Path(RUTA_FOTOS)
if not raiz.is_dir():
    raise FileNotFoundError('No aparece la carpeta. Autoriza la misma cuenta de Google que contiene 01_SAT_Huamanga. Ruta: '+RUTA_FOTOS)
extensiones={'.jpg','.jpeg','.png','.tif','.tiff','.bmp','.webp'}
problemas=[]
fotos=[]
for carpeta,dirs,archivos in os.walk(raiz,onerror=lambda e:problemas.append(str(e))):
    dirs[:]=[d for d in dirs if d!='__MACOSX']
    fotos.extend(Path(carpeta)/n for n in archivos if Path(n).suffix.lower() in extensiones)
print('FOTOS EN DRIVE:',len(fotos),'| ESPERADAS:',FOTOS_ESPERADAS,flush=True)
if problemas:
    raise RuntimeError('No se pudieron explorar todas las carpetas: '+str(problemas[:3]))
if len(fotos)!=FOTOS_ESPERADAS:
    raise RuntimeError('La cantidad no coincide con el inventario. Comprueba si terminó la subida antes de continuar. Si cambió el conjunto de fotos, actualiza FOTOS_ESPERADAS.')
print('Instalando dependencias; el progreso aparecerá debajo.',flush=True)
deps=['paddlepaddle==3.2.0','paddleocr==3.3.2','transformers==4.57.6',
      'torch>=2.6,<3','torchvision','sentencepiece','safetensors',
      'openpyxl==3.1.5','Pillow>=10,<13','lxml>=5,<7','ipywidgets>=8,<9']
proc=subprocess.Popen([sys.executable,'-m','pip','install',*deps],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,bufsize=1)
for linea in proc.stdout: print(linea.rstrip(),flush=True)
if proc.wait()!=0:
    raise RuntimeError('La instalación falló. Comparte las últimas líneas de ERROR; no inicies el OCR todavía.')
# Verificar archivos individuales por nombre relativo y tamaño, si está el inventario.
from openpyxl import load_workbook
inventario=raiz/'inventario_fotos.xlsx'
if inventario.exists():
    wb=load_workbook(inventario,read_only=True,data_only=True)
    registros=iter(wb.active.values); cabecera=next(registros)
    esperado={}
    for valores in registros:
        r=dict(zip(cabecera,valores))
        if not r.get('FullName'): continue
        partes=PureWindowsPath(r['FullName']).parts
        if 'PROYECTO REC 1' not in partes:
            wb.close();raise ValueError('Ruta no reconocida en inventario: '+str(r['FullName']))
        rel='/'.join(partes[partes.index('PROYECTO REC 1')+1:])
        if rel in esperado: wb.close();raise ValueError('Ruta duplicada en inventario: '+rel)
        esperado[rel]=int(r['Length'])
    wb.close()
    real={p.relative_to(raiz).as_posix():p.stat().st_size for p in fotos}
    faltan=sorted(set(esperado)-set(real));extras=sorted(set(real)-set(esperado))
    distintos=[k for k in esperado.keys() & real.keys() if esperado[k]!=real[k]]
    print('Verificación por inventario:',len(faltan),'faltantes,',len(extras),'adicionales,',len(distintos),'tamaños distintos.')
    if faltan or extras or distintos:
        print('Ejemplos:',(faltan+extras+distintos)[:10])
        raise RuntimeError('La subida no coincide con el inventario. Revisa los archivos indicados.')
else:
    print('No hay inventario XLSX: solo se comprobó el total de imágenes.')
import torch
print('GPU para TrOCR:',torch.cuda.get_device_name(0) if torch.cuda.is_available() else 'No disponible; se utilizará CPU.')
Path(RUTA_RESULTADOS).mkdir(parents=True,exist_ok=True)
CONFIGURACION_LISTA=True
print('PREPARACIÓN COMPLETA. Ejecuta la segunda celda para cargar el Excel y preparar el proyecto.')
