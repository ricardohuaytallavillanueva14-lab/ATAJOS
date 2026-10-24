# @title 2. Cargar el Excel, fijar el alcance y preparar el inventario
HOJA_LOTE = '2026'  # @param {type:"string"}
FECHA_EMISION_LOTE = '02/09/2026'  # @param {type:"string"}
EXCEL_SHA_ESPERADO = ''  # @param {type:"string"}
RUTA_EXCEL_EN_DRIVE = ''  # @param {type:"string"}
REHACER_INVENTARIO = False  # @param {type:"boolean"}
# HOJA_LOTE y FECHA_EMISION_LOTE delimitan el lote: hoja anual y columna E (emisión REC).
# RUTA_EXCEL_EN_DRIVE: ruta del Excel de expedientes en Drive. Vacío = se busca el único .xlsx que haya en la raíz de
#   RUTA_FOTOS (PROYECTO REC 1), sin contar inventario_fotos.xlsx; si no hay ninguno, se pide subirlo.
# EXCEL_SHA_ESPERADO: huella SHA-256 opcional para exigir un archivo exacto; vacío = no se comprueba. La carpeta de
#   resultados queda ligada igualmente a la huella del Excel con el que se creó (proyecto.json).
# REHACER_INVENTARIO: vuelve a leer el contenido de todas las fotos en lugar de usar la caché.
if not globals().get('CONFIGURACION_LISTA', False):
    raise RuntimeError('Ejecuta primero la celda 1.')
# --- A continuación va el código del proyecto (src/ocr_sat.py), incrustado para no clonar nada. ---
