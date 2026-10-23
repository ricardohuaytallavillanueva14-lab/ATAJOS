# @title 2. Cargar el Excel, fijar el alcance y preparar el inventario
HOJA_LOTE = '2026'  # @param {type:"string"}
FECHA_EMISION_LOTE = '02/09/2026'  # @param {type:"string"}
EXCEL_SHA_ESPERADO = '98012167fbeb97f1d28a724326dc2312f232e79acb8a19ff2e802ab04bac999f'  # @param {type:"string"}
RUTA_EXCEL_EN_DRIVE = ''  # @param {type:"string"}
REHACER_INVENTARIO = False  # @param {type:"boolean"}
# HOJA_LOTE y FECHA_EMISION_LOTE delimitan el lote: hoja anual y columna E (emisión REC).
# EXCEL_SHA_ESPERADO vacío: no se comprueba la huella del archivo (solo para lotes nuevos).
# RUTA_EXCEL_EN_DRIVE: ruta del .xlsx en Drive; vacío = subirlo desde el equipo la primera vez.
# REHACER_INVENTARIO: vuelve a leer el contenido de todas las fotos en lugar de usar la caché.
if not globals().get('CONFIGURACION_LISTA', False):
    raise RuntimeError('Ejecuta primero la celda 1.')
# --- A continuación va el código del proyecto (src/ocr_sat.py), incrustado para no clonar nada. ---
