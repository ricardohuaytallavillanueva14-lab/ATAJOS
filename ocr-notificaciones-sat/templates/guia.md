# Notificaciones REC — extractor de fechas desde fotografías escaneadas

Cuaderno preparado para Google Colab. Lee las fotografías de un lote de notificaciones guardadas en Drive, propone expediente, código SIGETI, resolución y fecha manuscrita de notificación con OCR gratuito (PaddleOCR + TrOCR), y, tras cotejo visual, completa la columna **F** en una **copia** del Excel de expedientes. Usa un entorno **alojado** con GPU **T4**: menú *Entorno de ejecución > Cambiar tipo de entorno de ejecución*.

**Alcance por defecto:** hoja **2026**, filas con **E = 02/09/2026** (fecha de emisión REC) y F vacía. Cambia `HOJA_LOTE` y `FECHA_EMISION_LOTE` en la celda 2 para otro lote; cada carpeta de resultados queda ligada a un Excel y a un alcance. Nunca se modifica el original ni otras hojas, filas o columnas.

## Ejecutar en orden
1. **Celda 1 · Preparación.** Monta Drive, cuenta las fotos (0 en `FOTOS_ESPERADAS` para no comprobar), instala solo las dependencias que faltan y guarda los modelos en `modelos_ocr_cache` dentro de Drive para no descargarlos en cada sesión. Si pide reiniciar el entorno, reinícialo y repite la celda.
2. **Celda 2 · Excel e inventario.** Toma el Excel desde una ruta de Drive (`RUTA_EXCEL_EN_DRIVE`) o lo sube desde tu equipo; lo guarda en `entrada/` para reanudar. Comprueba su huella SHA-256 si la indicas. Inventaría las fotos por contenido; las sesiones siguientes reutilizan `inventario_cache.json` y tardan segundos.
3. **Celda 3 · Prueba con muestra.** Procesa unas 20 fotos repartidas entre carpetas, muestra el tiempo por foto y una estimación de horas para el resto. Útil para decidir `MODO_RAPIDO` o `PADDLE_GPU` antes de la pasada completa.
4. **Celda 4 · Procesamiento.** Lee todas las pendientes con barra de progreso y tiempo estimado. `LIMITE` permite trabajar por lotes en sesiones cortas. Cada lectura se guarda al instante en Drive, así que una desconexión no pierde avances.
5. **Celda 5 · Revisión.** Muestra encabezado, cargo y el recorte que leyó el OCR; corrige los campos, escribe tu nombre, marca la casilla y pulsa **Confirmar y guardar**. Solo se admite REC UNO con expediente/año único y código SIGETI coincidente.
6. **Celda 6 · Exportación.** Genera `expedientes_actualizados_*.xlsx` y `revision_*.csv` en la carpeta de resultados. Solo escribe fechas confirmadas y sin conflictos.
7. **Celda 7 · Opcional.** ZIP con hasta N casos pendientes para una segunda revisión externa. No envía nada automáticamente.

## Reanudar y límites
Si Colab se desconecta, ejecuta de nuevo 1 y 2 (ahora rápidas gracias a las cachés) y luego 4. No ejecutes dos sesiones sobre la misma carpeta de resultados. Los errores se reintentan con `REINTENTAR_ERRORES`. Los límites de sesión de Colab siguen aplicando: con `LIMITE` puedes repartir el trabajo en varias sesiones.

**La escritura manuscrita puede leerse mal.** Una coincidencia entre motores no es una tasa de precisión medida. AVISO PREVIO, CEDULON, fotos giradas o diseños distintos pueden exigir transcripción manual. No se adivinan fechas ni expedientes. Si dos fotos confirmadas dan fechas distintas para una misma celda, esa celda queda sin modificar.

Referencias: [PaddleOCR](https://www.paddleocr.ai/latest/en/quick_start.html), [TrOCR](https://huggingface.co/docs/transformers/model_doc/trocr), [Colab](https://research.google.com/colaboratory/faq.html).
