# @title 6. Crear el Excel nuevo con los datos importantes (y, con Excel de expedientes, la copia con F)
SOLO_CONFIRMADAS = False  # @param {type:"boolean"}
FILTRAR_POR_EMISION = True  # @param {type:"boolean"}
DOBLE_REVISION = 'no_alta'  # @param ["ninguna", "no_alta", "todas"]
# El Excel nuevo (extraccion_*.xlsx) trae la hoja Expedientes (una fila por expediente: expediente, nombre y
# apellidos, código, fecha de emisión, fecha de notificación, verificación, confianza, alertas), la hoja Fotos
# (una fila por foto) y la hoja Resumen. SOLO_CONFIRMADAS deja en Expedientes solo las filas con cotejo humano.
# FILTRAR_POR_EMISION excluye de Expedientes los documentos cuya emisión impresa es de otro lote.
# DOBLE_REVISION solo aplica a la copia con columna F cuando se usa el Excel de expedientes.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
excel_nuevo, acta_nueva = proyecto.exportar_extraccion(solo_confirmadas=SOLO_CONFIRMADAS, filtrar_emision=FILTRAR_POR_EMISION)
carpeta = Path(RUTA_RESULTADOS).relative_to('/content/drive/MyDrive') if str(RUTA_RESULTADOS).startswith('/content/drive/MyDrive') else RUTA_RESULTADOS
print('LISTO: abre Drive >', carpeta, '>', excel_nuevo.name)
if proyecto.excel is not None:
    excel_generado, informe = proyecto.exportar(doble_revision={'ninguna': False, 'no_alta': 'no_alta', 'todas': True}[DOBLE_REVISION])
    if excel_generado:
        print('Copia del Excel de expedientes con F completada:', excel_generado.name)
    else:
        print('Aún no hay fechas aprobadas sin conflictos para la copia con F (o faltan segundas revisiones). Revisa el CSV y la celda 5.')
