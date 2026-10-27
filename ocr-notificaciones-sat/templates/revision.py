# @title 5. Revisión visual y confirmación (repite la celda para refrescar la lista)
MODO = 'pendientes'  # @param ["pendientes", "segunda", "todos"]
FORMA = 'texto'  # @param ["texto", "widget"]
REVISOR = ''  # @param {type:"string"}
LIMITE_REVISION = 0  # @param {type:"integer"}
# pendientes: lecturas sin confirmar. segunda: aprobadas por un solo revisor, para que otra persona las confirme.
# FORMA texto: muestra los recortes y pide cada decisión en un cuadro de texto (Enter confirma la propuesta,
#   DD/MM/AAAA la corrige, e edita expediente/código/resolución, p pendiente, o omitir, s salir). Funciona aunque
#   los widgets no se muestren. FORMA widget: formulario con botones (si sale en blanco, usa texto).
# Confirmar una fecha declara que la cotejaste con la foto; queda registrado con REVISOR y la hora.
# LIMITE_REVISION = 0 recorre toda la lista; un número limita la sesión.
# El orden prioriza las filas del alcance y la confianza OCR (ALTA, MEDIA, CONTRADICTORIA, SIN_FECHA).
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
if FORMA == 'texto':
    proyecto.revisar_texto(REVISOR, modo=MODO, limite=LIMITE_REVISION or None)
else:
    from google.colab import output as _salida
    _salida.enable_custom_widget_manager()
    proyecto.revisar(modo=MODO)
