# @title 7. OPCIONAL: paquete ZIP de casos pendientes para una segunda revisión
CASOS = 20  # @param {type:"integer"}
if 'proyecto' not in globals():
    raise RuntimeError('Ejecuta antes las celdas 1 y 2.')
paquete = proyecto.paquete_revision(limite=CASOS)
