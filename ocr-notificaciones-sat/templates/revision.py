# Repite esta celda para actualizar la lista de pendientes.
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
proyecto.revisar()
