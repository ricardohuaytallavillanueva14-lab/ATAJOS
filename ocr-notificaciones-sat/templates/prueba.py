# @title 3. Piloto con una muestra: velocidad, confianza y acierto contra F ya registrada
MUESTRA = 200  # @param {type:"integer"}
SEMILLA = 2026  # @param {type:"integer"}
MOTOR = 'auto'  # @param ["auto", "paddle", "easyocr"]
PADDLE_GPU = False  # @param {type:"boolean"}
MODO_RAPIDO = False  # @param {type:"boolean"}
LECTOR_VLM = True  # @param {type:"boolean"}
ORIENTACION_AUTO = False  # @param {type:"boolean"}
# Procesa una muestra repartida entre carpetas para medir velocidad y calidad de lectura.
# MOTOR auto prueba PaddleOCR en un subproceso y, si se cae (Python 3.13), usa EasyOCR en GPU.
# PADDLE_GPU requiere paddlepaddle-gpu; si no está disponible, vuelve a CPU automáticamente.
# Con unas 200 fotos, parte corresponderá a filas que ya tienen F: sirven de patrón para medir el acierto real.
# El acierto de ALTA solo es concluyente con al menos 30 casos ALTA con referencia; el resumen se guarda en piloto_*.json.
# MODO_RAPIDO usa el detector PP-OCRv5 mobile: más rápido en CPU, algo menos preciso.
# LECTOR_VLM añade Qwen2.5-VL-3B como tercer lector independiente de la fecha (~7 GB en GPU T4).
# ORIENTACION_AUTO corrige fotos giradas 90/180/270 grados (más lento).
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
if MUESTRA < 1:
    raise ValueError('MUESTRA debe ser al menos 1.')
opciones = dict(motor=MOTOR, gpu=torch.cuda.is_available(), paddle_gpu=PADDLE_GPU, modo_rapido=MODO_RAPIDO,
                vlm='qwen' if LECTOR_VLM else None, orientacion_auto=ORIENTACION_AUTO)
if 'motor' not in globals() or any(getattr(motor, 'opciones', {}).get(k) != v for k, v in opciones.items() if k != 'gpu'):
    print('Cargando modelos. El primer uso los descarga (o los lee de la caché en Drive).')
    motor = MotorOCR(**opciones)
print('Motor:', motor.nombre)
muestra = seleccionar_muestra(proyecto.listar_fotos(), proyecto.fotos, cantidad=MUESTRA, semilla=SEMILLA)
ids_muestra = proyecto.ids_de(muestra)
resumen = proyecto.procesar(motor, limite=len(ids_muestra), solo=ids_muestra)
print('Resultado de la muestra:', resumen)
for k in ids_muestra:
    r = proyecto.records[k]
    d = r.get('lectura') or {}
    print(f"- {r['archivos'][0]} | {r['estado']} | {r.get('confianza') or '-'} | exp={d.get('expediente')} cod={d.get('codigo')} "
          f"fecha={d.get('fecha_notificacion')} | {(r.get('problemas') or r.get('alertas') or [r.get('error') or 'confirmar visualmente'])[0][:70]}")
from collections import Counter
print('Confianza en la muestra:', dict(Counter(proyecto.records[k].get('confianza') or 'ERROR' for k in ids_muestra)))
print('Tipos de documento:', dict(Counter(proyecto.records[k].get('tipo') or '-' for k in ids_muestra)))
print(aviso_lectores(motor))
acierto = proyecto.medir_acierto(ids_muestra)
print('Si el ritmo (s/foto) y el acierto son aceptables, continúa con la celda 4. Las lecturas de la muestra ya quedan guardadas.')
