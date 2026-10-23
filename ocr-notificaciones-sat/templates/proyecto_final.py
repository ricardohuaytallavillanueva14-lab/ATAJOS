
from google.colab import files, output
output.enable_custom_widget_manager()
EXCEL_SHA_ESPERADO = '98012167fbeb97f1d28a724326dc2312f232e79acb8a19ff2e802ab04bac999f'
entrada=Path(RUTA_RESULTADOS)/'entrada'
entrada.mkdir(exist_ok=True)
EXCEL=entrada/'original_expedientes.xlsx'
if not EXCEL.exists():
    print('Selecciona el mismo Excel adjuntado en ChatGPT (termina en (1).xlsx).')
    subidos=files.upload()
    validos=[b for n,b in subidos.items() if n.lower().endswith('.xlsx') and hashlib.sha256(b).hexdigest()==EXCEL_SHA_ESPERADO]
    if len(validos)!=1:
        raise ValueError('No se recibió el Excel exacto revisado. Selecciona el archivo original adjuntado, sin volver a guardarlo en Excel.')
    temporal=entrada/'original.tmp'
    temporal.write_bytes(validos[0]); os.replace(temporal,EXCEL)
if sha256(EXCEL)!=EXCEL_SHA_ESPERADO:
    raise ValueError('El original guardado cambió; no se iniciará el proceso.')
proyecto=Proyecto(RUTA_FOTOS,EXCEL,RUTA_RESULTADOS,hojas=('2026',))
print('Filtro fijo: hoja 2026; E = 02/09/2026; destino F vacía.')
print('Inventariando todas las imágenes; puede tardar porque verifica su contenido.')
proyecto.inventario()
print('PROYECTO LISTO. Ejecuta la celda 3.')
