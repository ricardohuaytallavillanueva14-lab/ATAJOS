
# --- Fin del código incrustado. Carga del Excel y del inventario. ---
from google.colab import files, output
import shutil
output.enable_custom_widget_manager()
configurar(HOJA_LOTE, FECHA_EMISION_LOTE)
EXCEL_SHA_ESPERADO = EXCEL_SHA_ESPERADO.strip().lower()
entrada = Path(RUTA_RESULTADOS) / 'entrada'
entrada.mkdir(exist_ok=True)
EXCEL = entrada / 'original_expedientes.xlsx'
if not EXCEL.exists():
    temporal = entrada / 'original.tmp'
    if RUTA_EXCEL_EN_DRIVE.strip():
        origen = Path(RUTA_EXCEL_EN_DRIVE.strip())
        if not origen.is_file() or origen.suffix.lower() != '.xlsx':
            raise FileNotFoundError('No se encontró el .xlsx indicado en RUTA_EXCEL_EN_DRIVE: ' + str(origen))
        shutil.copyfile(origen, temporal)
    else:
        print('Selecciona el Excel original de expedientes (.xlsx). Se guardará en Drive para reanudar sin volver a subirlo.')
        subidos = files.upload()
        validos = [b for n, b in subidos.items() if n.lower().endswith('.xlsx')]
        if len(validos) != 1:
            raise ValueError('Sube exactamente un archivo .xlsx.')
        temporal.write_bytes(validos[0])
    if EXCEL_SHA_ESPERADO and sha256(temporal) != EXCEL_SHA_ESPERADO:
        temporal.unlink()
        raise ValueError('No se recibió el Excel exacto revisado (la huella SHA-256 no coincide). '
                         'Selecciona el archivo original sin volver a guardarlo en Excel, o vacía EXCEL_SHA_ESPERADO si es un lote nuevo.')
    os.replace(temporal, EXCEL)
if EXCEL_SHA_ESPERADO and sha256(EXCEL) != EXCEL_SHA_ESPERADO:
    raise ValueError('El original guardado en Drive cambió; no se iniciará el proceso.')
print('Excel:', EXCEL, '| SHA-256:', sha256(EXCEL)[:16] + '…')
proyecto = Proyecto(RUTA_FOTOS, EXCEL, RUTA_RESULTADOS, hojas=(HOJA,))
print(f'Filtro fijo: hoja {HOJA}; E = {FECHA_EMISION:%d/%m/%Y}; destino F vacía.')
proyecto.resumen_alcance()
print('Inventariando las imágenes. La primera vez lee todo el contenido; después reutiliza la caché.')
proyecto.inventario(rehash=REHACER_INVENTARIO)
print('PROYECTO LISTO. Ejecuta la celda 3 (prueba con muestra) o directamente la 4.')
