# @title 6. Exportar la copia del Excel, el informe CSV y el acta de exportación
DOBLE_REVISION = 'no_alta'  # @param ["ninguna", "no_alta", "todas"]
# ninguna: basta una aprobación. no_alta: dos revisores distintos cuando la confianza OCR no fue ALTA (recomendado).
# todas: dos revisores distintos para cada fecha. Ejecuta tantas veces como necesites.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
excel_generado, informe = proyecto.exportar(doble_revision={'ninguna': False, 'no_alta': 'no_alta', 'todas': True}[DOBLE_REVISION])
if excel_generado:
    print('LISTO: abre Drive >', Path(RUTA_RESULTADOS).relative_to('/content/drive/MyDrive') if str(RUTA_RESULTADOS).startswith('/content/drive/MyDrive') else RUTA_RESULTADOS)
    print('Copia actualizada:', excel_generado.name)
else:
    print('Aún no hay fechas aprobadas sin conflictos (o faltan segundas revisiones). Revisa el informe CSV y la celda 5.')
