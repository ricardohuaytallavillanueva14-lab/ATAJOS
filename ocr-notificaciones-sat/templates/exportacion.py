# Ejecuta después de confirmar lecturas, y tantas veces como necesites.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
excel_generado,informe=proyecto.exportar()
if excel_generado:
    print('LISTO: abre Drive > 01_SAT_Huamanga > REC_2026_02SEP_RESULTADOS')
    print('Copia actualizada:',excel_generado.name)
else:
    print('Aún no hay fechas aprobadas sin conflictos. Confirma lecturas en la celda 4.')
