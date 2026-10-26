# @title 4. Procesar todas las fotografías pendientes
LIMITE = 0  # @param {type:"integer"}
REINTENTAR_ERRORES = False  # @param {type:"boolean"}
MOTOR = 'auto'  # @param ["auto", "paddle_aislado", "paddle", "easyocr"]
PADDLE_GPU = False  # @param {type:"boolean"}
MODO_RAPIDO = False  # @param {type:"boolean"}
LECTOR_VLM = True  # @param {type:"boolean"}
ORIENTACION_AUTO = False  # @param {type:"boolean"}
VLM_DOS_VISTAS = True  # @param {type:"boolean"}
LECTOR_API = 'ninguno'  # @param ["ninguno", "gemini", "deepseek"]
API_PAUSA_S = 6  # @param {type:"integer"}
API_MODELO = 'gemini-3.8-flash'  # @param {type:"string"}
# LIMITE = 0 procesa todas las pendientes; un número procesa solo ese lote (útil para sesiones cortas).
# Las copias idénticas se leen una sola vez. Cada lectura se guarda al instante: puedes reanudar.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
pendientes = proyecto.pendientes(REINTENTAR_ERRORES)
if pendientes:
    clave_api = None
    if LECTOR_API != 'ninguno':
        try:
            from google.colab import userdata
            clave_api = userdata.get({'gemini': 'GEMINI_API_KEY', 'deepseek': 'DEEPSEEK_API_KEY'}[LECTOR_API])
        except Exception as e:
            raise RuntimeError(f'No se pudo leer el secreto de la API ({e}). En Colab: panel izquierdo > icono de llave > añade GEMINI_API_KEY (o DEEPSEEK_API_KEY) y activa el acceso del cuaderno.')
    opciones = dict(motor=MOTOR, gpu=torch.cuda.is_available(), paddle_gpu=PADDLE_GPU, modo_rapido=MODO_RAPIDO,
                    vlm='qwen' if LECTOR_VLM else None, orientacion_auto=ORIENTACION_AUTO,
                    paddle_python=globals().get('PADDLE_PYTHON'), vlm_vistas=2 if VLM_DOS_VISTAS else 1,
                    api=None if LECTOR_API == 'ninguno' else LECTOR_API, api_vistas=1, api_pausa=API_PAUSA_S,
                api_modelo=API_MODELO.strip() or None)
    if 'motor' not in globals() or any(getattr(motor, 'opciones', {}).get(k) != v for k, v in opciones.items() if k != 'gpu'):
        if clave_api: opciones['api_key'] = clave_api
        print('Cargando modelos. El primer uso los descarga (o los lee de la caché en Drive).')
        motor = MotorOCR(**opciones)
    print('Motor:', motor.nombre, '| pendientes:', len(pendientes))
    resumen = proyecto.procesar(motor, limite=LIMITE or len(pendientes), reintentar_errores=REINTENTAR_ERRORES)
else:
    print('No hay lecturas nuevas pendientes; se conservan los avances.')
# Estructura de los documentos leídos y conciliación Excel-fotos por expediente.
estructura = proyecto.sondeo_estructura()
conciliacion = proyecto.conciliar()
# Genera un informe aun si todavía no se ha aprobado ninguna fecha.
proyecto.exportar()
print('Las propuestas OCR requieren cotejo visual. Continúa con la celda 5.')
