
# --- Fin del código incrustado. Carga del Excel y del inventario. ---
from google.colab import files, output
import shutil
output.enable_custom_widget_manager()
configurar(HOJA_LOTE, FECHA_EMISION_LOTE)
EXCEL_SHA_ESPERADO = EXCEL_SHA_ESPERADO.strip().lower()
entrada = Path(RUTA_RESULTADOS) / 'entrada'
entrada.mkdir(exist_ok=True)
EXCEL = entrada / 'original_expedientes.xlsx'
EXCEL_DRIVE = None  # Excel de trabajo en Drive; la celda 6 puede escribir en él con respaldo.
if RUTA_EXCEL_EN_DRIVE.strip():
    EXCEL_DRIVE = Path(RUTA_EXCEL_EN_DRIVE.strip())
    if not EXCEL_DRIVE.is_file() or EXCEL_DRIVE.suffix.lower() != '.xlsx':
        raise FileNotFoundError('No se encontró el .xlsx indicado en RUTA_EXCEL_EN_DRIVE: ' + str(EXCEL_DRIVE))
else:
    candidatos = sorted(p for p in Path(RUTA_FOTOS).glob('*.xlsx') if p.name.lower() != 'inventario_fotos.xlsx' and '_respaldo_' not in p.name)
    if len(candidatos) == 1:
        EXCEL_DRIVE = candidatos[0]
    elif len(candidatos) > 1:
        raise ValueError('Hay varios .xlsx en la carpeta de fotos; indica cuál en RUTA_EXCEL_EN_DRIVE: ' + ', '.join(p.name for p in candidatos))
if EXCEL_DRIVE is not None:
    print('Excel de trabajo en Drive:', EXCEL_DRIVE)
    if EXCEL.exists() and sha256(EXCEL) != sha256(EXCEL_DRIVE):
        raise ValueError('El Excel de Drive no coincide con la copia de entrada guardada en la carpeta de resultados. '
                         'Si cambiaste el Excel, usa otra carpeta de resultados; si lo actualizó la celda 6, la copia de entrada ya debería coincidir.')
if not EXCEL.exists():
    temporal = entrada / 'original.tmp'
    if EXCEL_DRIVE is not None:
        shutil.copyfile(EXCEL_DRIVE, temporal)
    else:
        print('No hay .xlsx en la carpeta de fotos. Selecciona el Excel de expedientes; se guardará en la carpeta de resultados.')
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
proyecto.conciliar(guardar=False)
print('Inventariando las imágenes. La primera vez lee todo el contenido; después reutiliza la caché.')
proyecto.inventario(rehash=REHACER_INVENTARIO)
print('PROYECTO LISTO. Ejecuta la celda 3 (prueba con muestra) o directamente la 4.')
