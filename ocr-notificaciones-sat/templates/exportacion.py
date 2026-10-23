# @title 6. Exportar la copia del Excel y el informe CSV
# Ejecuta después de confirmar lecturas, y tantas veces como necesites.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
excel_generado, informe = proyecto.exportar()
if excel_generado:
    print('LISTO: abre Drive >', Path(RUTA_RESULTADOS).relative_to('/content/drive/MyDrive') if str(RUTA_RESULTADOS).startswith('/content/drive/MyDrive') else RUTA_RESULTADOS)
    print('Copia actualizada:', excel_generado.name)
else:
    print('Aún no hay fechas aprobadas sin conflictos. Confirma lecturas en la celda 5.')
