# @title 4. Procesar todas las fotografías pendientes
LIMITE = 0  # @param {type:"integer"}
REINTENTAR_ERRORES = False  # @param {type:"boolean"}
MOTOR = 'auto'  # @param ["auto", "paddle", "easyocr"]
PADDLE_GPU = False  # @param {type:"boolean"}
MODO_RAPIDO = False  # @param {type:"boolean"}
LECTOR_VLM = True  # @param {type:"boolean"}
ORIENTACION_AUTO = False  # @param {type:"boolean"}
# LIMITE = 0 procesa todas las pendientes; un número procesa solo ese lote (útil para sesiones cortas).
# Las copias idénticas se leen una sola vez. Cada lectura se guarda al instante: puedes reanudar.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
pendientes = proyecto.pendientes(REINTENTAR_ERRORES)
if pendientes:
    opciones = dict(motor=MOTOR, gpu=torch.cuda.is_available(), paddle_gpu=PADDLE_GPU, modo_rapido=MODO_RAPIDO,
                vlm='qwen' if LECTOR_VLM else None, orientacion_auto=ORIENTACION_AUTO)
    if 'motor' not in globals() or any(getattr(motor, 'opciones', {}).get(k) != v for k, v in opciones.items() if k != 'gpu'):
        print('Cargando modelos. El primer uso los descarga (o los lee de la caché en Drive).')
        motor = MotorOCR(**opciones)
    print('Motor:', motor.nombre, '| pendientes:', len(pendientes))
    resumen = proyecto.procesar(motor, limite=LIMITE or len(pendientes), reintentar_errores=REINTENTAR_ERRORES)
else:
    print('No hay lecturas nuevas pendientes; se conservan los avances.')
# Genera un informe aun si todavía no se ha aprobado ninguna fecha.
proyecto.exportar()
print('Las propuestas OCR requieren cotejo visual. Continúa con la celda 5.')
