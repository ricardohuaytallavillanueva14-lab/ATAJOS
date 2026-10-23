# @title 1. Preparar Colab: Drive, fotografías, dependencias y caché de modelos
# Rellena el formulario de la derecha y ejecuta la celda. Es seguro repetirla.
RUTA_FOTOS = '/content/drive/MyDrive/01_SAT_Huamanga/PROYECTO REC 1'  # @param {type:"string"}
RUTA_RESULTADOS = '/content/drive/MyDrive/01_SAT_Huamanga/REC_2026_02SEP_RESULTADOS'  # @param {type:"string"}
FOTOS_ESPERADAS = 3996  # @param {type:"integer"}
CACHE_MODELOS_EN_DRIVE = True  # @param {type:"boolean"}
REINSTALAR_DEPENDENCIAS = False  # @param {type:"boolean"}
# FOTOS_ESPERADAS = 0 omite la comprobación de cantidad (lotes nuevos o parciales).
# CACHE_MODELOS_EN_DRIVE guarda TrOCR y PaddleOCR en Drive para no descargarlos en cada sesión.

from pathlib import Path, PureWindowsPath
import importlib.metadata, os, sys, subprocess, time
CONFIGURACION_LISTA = False

try:
    from google.colab import drive  # noqa: F401
except ImportError:
    raise RuntimeError('Este cuaderno está preparado para Google Colab con entorno alojado. '
                       'Ábrelo desde colab.research.google.com y elige Entorno de ejecución > Cambiar tipo > GPU T4.')
if os.name == 'nt':
    raise RuntimeError('Sigues conectado a tu PC. Selecciona un entorno alojado de Colab y vuelve a ejecutar.')

drive.mount('/content/drive')
raiz = Path(RUTA_FOTOS)
if not raiz.is_dir():
    raise FileNotFoundError('No aparece la carpeta. Autoriza la misma cuenta de Google que contiene las fotos. Ruta: ' + RUTA_FOTOS)
salida = Path(RUTA_RESULTADOS)
if salida == raiz or raiz in salida.parents:
    raise ValueError('RUTA_RESULTADOS debe estar fuera de la carpeta de fotografías.')

# Caché de modelos en Drive: debe fijarse antes de importar transformers o paddleocr.
if CACHE_MODELOS_EN_DRIVE:
    cache_modelos = salida.parent / 'modelos_ocr_cache'
    (cache_modelos / 'huggingface').mkdir(parents=True, exist_ok=True)
    (cache_modelos / 'paddlex').mkdir(parents=True, exist_ok=True)
    os.environ['HF_HOME'] = str(cache_modelos / 'huggingface')
    os.environ['PADDLE_PDX_CACHE_HOME'] = str(cache_modelos / 'paddlex')
    os.environ['HF_HUB_DISABLE_SYMLINKS_WARNING'] = '1'
    print('Caché de modelos:', cache_modelos)

extensiones = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.webp'}
problemas, fotos = [], []
t0 = time.monotonic()
for carpeta, dirs, archivos in os.walk(raiz, onerror=lambda e: problemas.append(str(e))):
    dirs[:] = [d for d in dirs if d != '__MACOSX']
    fotos.extend(Path(carpeta) / n for n in archivos if Path(n).suffix.lower() in extensiones)
print(f'FOTOS EN DRIVE: {len(fotos)} | ESPERADAS: {FOTOS_ESPERADAS or "sin comprobar"} | {time.monotonic()-t0:.0f} s', flush=True)
if problemas:
    raise RuntimeError('No se pudieron explorar todas las carpetas: ' + str(problemas[:3]))
if FOTOS_ESPERADAS and len(fotos) != FOTOS_ESPERADAS:
    raise RuntimeError('La cantidad no coincide con el inventario. Comprueba si terminó la subida antes de continuar. '
                       'Si cambió el conjunto de fotos, actualiza FOTOS_ESPERADAS (0 para no comprobar).')

# Instalación: solo lo que falta o no coincide con la versión fijada. Colab ya trae torch con CUDA.
deps = ['paddlepaddle==3.2.0', 'paddleocr==3.3.2', 'easyocr>=1.7,<2', 'transformers==4.57.6',
        'torch>=2.6,<3', 'torchvision', 'sentencepiece', 'safetensors', 'accelerate>=1,<2',
        'openpyxl==3.1.5', 'Pillow>=10,<13', 'lxml>=5,<7', 'ipywidgets>=8,<9', 'tqdm>=4,<5']
def version_instalada(nombre):
    try:
        return importlib.metadata.version(nombre)
    except importlib.metadata.PackageNotFoundError:
        return None
fijas = {d.split('==')[0]: d.split('==')[1] for d in deps if '==' in d}
faltan = [n for n, v in fijas.items() if version_instalada(n) != v]
numpy_antes = version_instalada('numpy')
if REINSTALAR_DEPENDENCIAS or faltan:
    print('Instalando dependencias' + (f' (pendientes: {", ".join(faltan)})' if faltan else '') + '; el progreso aparecerá debajo.', flush=True)
    proc = subprocess.Popen([sys.executable, '-m', 'pip', 'install', '--progress-bar', 'off', *deps],
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
    for linea in proc.stdout:
        if linea.strip() and not linea.startswith('Requirement already satisfied'):
            print(linea.rstrip(), flush=True)
    if proc.wait() != 0:
        raise RuntimeError('La instalación falló. Comparte las últimas líneas de ERROR; no inicies el OCR todavía.')
    if version_instalada('numpy') != numpy_antes:
        raise RuntimeError('pip cambió la versión de numpy. Reinicia el entorno (Entorno de ejecución > Reiniciar sesión) '
                           'y vuelve a ejecutar esta celda; la segunda vez no reinstalará nada.')
else:
    print('Dependencias ya instaladas con las versiones fijadas.')
# paddlex importa langchain.docstore si langchain está instalado; Colab trae langchain 1.x, que ya no lo tiene.
if version_instalada('langchain'):
    print('Retirando langchain (incompatible con paddlex en Colab)...', flush=True)
    subprocess.run([sys.executable, '-m', 'pip', 'uninstall', '-y', '-q', 'langchain'], check=False)

# Verificar archivos individuales por nombre relativo y tamaño, si está el inventario.
from openpyxl import load_workbook
inventario = raiz / 'inventario_fotos.xlsx'
if inventario.exists():
    wb = load_workbook(inventario, read_only=True, data_only=True)
    registros = iter(wb.active.values); cabecera = next(registros)
    esperado = {}
    for valores in registros:
        r = dict(zip(cabecera, valores))
        if not r.get('FullName'):
            continue
        partes = PureWindowsPath(r['FullName']).parts
        if raiz.name not in partes:
            wb.close(); raise ValueError('Ruta no reconocida en inventario: ' + str(r['FullName']))
        rel = '/'.join(partes[partes.index(raiz.name) + 1:])
        if rel in esperado:
            wb.close(); raise ValueError('Ruta duplicada en inventario: ' + rel)
        esperado[rel] = int(r['Length'])
    wb.close()
    real = {p.relative_to(raiz).as_posix(): p.stat().st_size for p in fotos}
    faltan = sorted(set(esperado) - set(real)); extras = sorted(set(real) - set(esperado))
    distintos = [k for k in esperado.keys() & real.keys() if esperado[k] != real[k]]
    print('Verificación por inventario:', len(faltan), 'faltantes,', len(extras), 'adicionales,', len(distintos), 'tamaños distintos.')
    if faltan or extras or distintos:
        print('Ejemplos:', (faltan + extras + distintos)[:10])
        raise RuntimeError('La subida no coincide con el inventario. Revisa los archivos indicados.')
else:
    print('No hay inventario_fotos.xlsx: solo se comprobó el total de imágenes.')

import torch
print('GPU para TrOCR:', torch.cuda.get_device_name(0) if torch.cuda.is_available() else 'No disponible; se utilizará CPU (mucho más lento).')
print('CPU disponibles:', os.cpu_count())
salida.mkdir(parents=True, exist_ok=True)
CONFIGURACION_LISTA = True
print('PREPARACIÓN COMPLETA. Ejecuta la celda 2 para cargar el Excel y preparar el proyecto.')
