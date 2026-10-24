# @title 2. Fijar el lote, (opcionalmente) el Excel de expedientes y preparar el inventario
HOJA_LOTE = '2026'  # @param {type:"string"}
FECHA_EMISION_LOTE = '02/09/2026'  # @param {type:"string"}
USAR_EXCEL = False  # @param {type:"boolean"}
EXCEL_SHA_ESPERADO = '07545132e6379309f69c04b5449fe389f2c7ffab08b109c62e6045f3c672ef22'  # @param {type:"string"}
RUTA_EXCEL_EN_DRIVE = ''  # @param {type:"string"}
REHACER_INVENTARIO = False  # @param {type:"boolean"}
# HOJA_LOTE y FECHA_EMISION_LOTE delimitan el lote: año del expediente y fecha de emisión impresa en la REC.
# USAR_EXCEL = False (recomendado ahora): no se carga ningún Excel; la celda 6 crea un Excel NUEVO con los datos
#   importantes de cada documento. USAR_EXCEL = True: además valida contra el Excel de expedientes y puede
#   completar la columna F en una copia (los campos siguientes solo se usan en ese caso).
# EXCEL_SHA_ESPERADO vacío: no se comprueba la huella del archivo (solo para lotes nuevos).
# RUTA_EXCEL_EN_DRIVE: ruta del .xlsx en Drive; vacío = subirlo desde el equipo la primera vez.
# REHACER_INVENTARIO: vuelve a leer el contenido de todas las fotos en lugar de usar la caché.
if not globals().get('CONFIGURACION_LISTA', False):
    raise RuntimeError('Ejecuta primero la celda 1.')
# --- A continuación va el código del proyecto (src/ocr_sat.py), incrustado para no clonar nada. ---
