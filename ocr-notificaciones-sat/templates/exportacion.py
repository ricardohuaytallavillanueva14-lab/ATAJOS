# @title 6. Exportar la copia del Excel, el informe CSV y el acta; opcionalmente escribir en el Excel de Drive
DOBLE_REVISION = 'no_alta'  # @param ["ninguna", "no_alta", "todas"]
ESCRIBIR_EN_EXCEL_DE_DRIVE = False  # @param {type:"boolean"}
# ninguna: basta una aprobación. no_alta: dos revisores distintos cuando la confianza OCR no fue ALTA (recomendado).
# todas: dos revisores distintos para cada fecha. Ejecuta tantas veces como necesites.
# ESCRIBIR_EN_EXCEL_DE_DRIVE: además de la copia en resultados, lleva las fechas confirmadas al Excel de trabajo de
#   PROYECTO REC 1 (el que se abre desde Drive), dejando antes un respaldo <nombre>_respaldo_<fecha>.xlsx en la misma carpeta.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
excel_generado, informe = proyecto.exportar(doble_revision={'ninguna': False, 'no_alta': 'no_alta', 'todas': True}[DOBLE_REVISION])
if excel_generado:
    print('LISTO: abre Drive >', Path(RUTA_RESULTADOS).relative_to('/content/drive/MyDrive') if str(RUTA_RESULTADOS).startswith('/content/drive/MyDrive') else RUTA_RESULTADOS)
    print('Copia actualizada:', excel_generado.name)
    if ESCRIBIR_EN_EXCEL_DE_DRIVE:
        if globals().get('EXCEL_DRIVE') is None:
            raise RuntimeError('No se conoce el Excel de trabajo en Drive (la celda 2 lo tomó de una subida). Indica RUTA_EXCEL_EN_DRIVE en la celda 2.')
        proyecto.actualizar_excel_en_drive(excel_generado, EXCEL_DRIVE)
        print('Las fechas confirmadas ya están en el Excel de Drive. Si vuelves a exportar, solo se añadirán las nuevas.')
else:
    print('Aún no hay fechas aprobadas sin conflictos (o faltan segundas revisiones). Confirma lecturas en la celda 5.')
