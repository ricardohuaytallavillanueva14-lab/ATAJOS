# @title 4. Procesar todas las fotografías pendientes
LIMITE = 0  # @param {type:"integer"}
REINTENTAR_ERRORES = False  # @param {type:"boolean"}
PADDLE_GPU = False  # @param {type:"boolean"}
MODO_RAPIDO = False  # @param {type:"boolean"}
# LIMITE = 0 procesa todas las pendientes; un número procesa solo ese lote (útil para sesiones cortas).
# Las copias idénticas se leen una sola vez. Cada lectura se guarda al instante: puedes reanudar.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
pendientes = proyecto.pendientes(REINTENTAR_ERRORES)
if pendientes:
    opciones = dict(motor='paddle_trocr', gpu=torch.cuda.is_available(), paddle_gpu=PADDLE_GPU, modo_rapido=MODO_RAPIDO)
    if 'motor' not in globals() or getattr(motor, 'opciones', {}).get('paddle_gpu') != PADDLE_GPU or motor.opciones.get('modo_rapido') != MODO_RAPIDO:
        print('Cargando modelos. El primer uso los descarga (o los lee de la caché en Drive).')
        motor = MotorOCR(**opciones)
    print('Motor:', motor.nombre, '| pendientes:', len(pendientes))
    resumen = proyecto.procesar(motor, limite=LIMITE or len(pendientes), reintentar_errores=REINTENTAR_ERRORES)
else:
    print('No hay lecturas nuevas pendientes; se conservan los avances.')
# Genera un informe aun si todavía no se ha aprobado ninguna fecha.
proyecto.exportar()
print('Las propuestas OCR requieren cotejo visual. Continúa con la celda 5.')
