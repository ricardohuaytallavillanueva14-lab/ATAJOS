# @title 3. Prueba rápida con una muestra (recomendada antes de procesar todo)
MUESTRA = 20  # @param {type:"integer"}
PADDLE_GPU = False  # @param {type:"boolean"}
MODO_RAPIDO = False  # @param {type:"boolean"}
# Procesa una muestra repartida entre carpetas para medir velocidad y calidad de lectura.
# PADDLE_GPU requiere paddlepaddle-gpu; si no está disponible, vuelve a CPU automáticamente.
# MODO_RAPIDO usa el detector PP-OCRv5 mobile: más rápido en CPU, algo menos preciso.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
if MUESTRA < 1:
    raise ValueError('MUESTRA debe ser al menos 1.')
opciones = dict(motor='paddle_trocr', gpu=torch.cuda.is_available(), paddle_gpu=PADDLE_GPU, modo_rapido=MODO_RAPIDO)
if 'motor' not in globals() or getattr(motor, 'opciones', {}).get('paddle_gpu') != PADDLE_GPU or motor.opciones.get('modo_rapido') != MODO_RAPIDO:
    print('Cargando modelos. El primer uso los descarga (o los lee de la caché en Drive).')
    motor = MotorOCR(**opciones)
print('Motor:', motor.nombre)
muestra = seleccionar_muestra(proyecto.listar_fotos(), proyecto.fotos, cantidad=MUESTRA)
ids_muestra = proyecto.ids_de(muestra)
resumen = proyecto.procesar(motor, limite=len(ids_muestra), solo=ids_muestra)
print('Resultado de la muestra:', resumen)
for k in ids_muestra:
    r = proyecto.records[k]
    d = r.get('lectura') or {}
    print(f"- {r['archivos'][0]} | {r['estado']} | exp={d.get('expediente')} cod={d.get('codigo')} "
          f"fecha={d.get('fecha_notificacion')} | {(r.get('problemas') or [r.get('error') or 'confirmar visualmente'])[0][:70]}")
print('Si el ritmo (s/foto) es aceptable, continúa con la celda 4. Las lecturas de la muestra ya quedan guardadas.')
