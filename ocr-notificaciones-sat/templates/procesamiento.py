# Lee TODAS las imágenes únicas; las copias idénticas conservan sus rutas.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
REINTENTAR_ERRORES = False  # Cambia a True si deseas reintentar errores de una ejecución anterior.
pendientes=[k for k in proyecto.active if proyecto.records[k]['estado']=='NUEVO'
            or (REINTENTAR_ERRORES and proyecto.records[k]['estado']=='ERROR')]
if pendientes:
    print('Cargando modelos gratuitos. El primer uso requiere descargarlos.')
    motor=MotorOCR(motor='paddle_trocr',gpu=torch.cuda.is_available())
    proyecto.procesar(motor,limite=len(pendientes),reintentar_errores=REINTENTAR_ERRORES)
else:
    print('No hay lecturas nuevas pendientes; se conservan los avances.')
# Genera un informe aun si todavía no se ha aprobado ninguna fecha.
proyecto.exportar()
print('Las propuestas OCR requieren cotejo visual. Continúa con la celda 4.')
