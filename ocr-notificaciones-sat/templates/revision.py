# @title 5. Revisión visual y confirmación (repite la celda para refrescar la lista)
MODO = 'pendientes'  # @param ["pendientes", "segunda", "todos"]
# pendientes: lecturas sin confirmar. segunda: aprobadas por un solo revisor, para que otra persona las confirme.
# El orden prioriza las filas del alcance y la confianza OCR (ALTA, MEDIA, CONTRADICTORIA, SIN_FECHA).
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
from google.colab import output as _salida
_salida.enable_custom_widget_manager()
proyecto.revisar(modo=MODO)
