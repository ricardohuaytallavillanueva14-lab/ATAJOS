# Contexto para supervisión de Colab

Supervisa `REC_2026_02_SEPTIEMBRE_Colab.ipynb`. El objetivo es leer las 3.996 fotos y, tras cotejo visual, completar F vacía únicamente en la hoja 2026 cuando E sea 02/09/2026.

## Ubicaciones

- Fotos: `/content/drive/MyDrive/01_SAT_Huamanga/PROYECTO REC 1` y todas sus subcarpetas.
- Salida: `/content/drive/MyDrive/01_SAT_Huamanga/REC_2026_02SEP_RESULTADOS`.
- Original guardado: `entrada/original_expedientes.xlsx`, dentro de la salida.
- Configuración: `proyecto.json` (Excel, hoja y fecha de emisión del lote).
- Caché de huellas: `inventario_cache.json` (tamaño y fecha por foto; `REHACER_INVENTARIO` la ignora).
- Modelos: `modelos_ocr_cache/` junto a la carpeta de resultados (TrOCR y PaddleOCR).
- Avance: `registros/*.json` por SHA-256 del contenido fotográfico (lecturas de cada lector, confianza, alertas, aprobaciones).
- Evidencia: `evidencia/<sha256>_fecha.jpg` y `_cargo.jpg` guardados al aprobar.
- Actas: `acta_exportacion_*.json` con huellas del original y de cada copia generada.
- Exportaciones: `expedientes_actualizados_*.xlsx` y `revision_*.csv`.

El Excel de trabajo es el `.xlsx` guardado en Drive dentro de `PROYECTO REC 1` (el que el usuario abre desde Drive). La celda 2 lo localiza, lo copia a `entrada/original_expedientes.xlsx` y liga la carpeta de resultados a su huella SHA-256 (`proyecto.json`); abrirlo y volver a guardarlo cambia la huella. Con `ESCRIBIR_EN_EXCEL_DE_DRIVE` (celda 6) las fechas confirmadas se escriben en ese mismo Excel tras guardar un respaldo `_respaldo_<fecha>.xlsx` junto a él; `proyecto.json` conserva el historial de huellas. No se usa ninguna otra copia del Excel. No guardes identificadores de Drive ni datos de contribuyentes en el repositorio.

## Funcionamiento

1. Preparación: montaje de Drive, verificación de las fotos (`FOTOS_ESPERADAS`, 0 = sin comprobar) y del inventario XLSX si existe, instalación solo de lo que falta, caché de modelos en Drive.
2. Proyecto: carga de código, `configurar(HOJA_LOTE, FECHA_EMISION_LOTE)`, Excel desde Drive o subido, índice e inventario (reutiliza la caché de huellas).
3. Sondeo y piloto: `sondeo_formato()` sobre todas las fotos sin OCR; muestra de `MUESTRA` fotos con acierto contra F registrada, `sondeo_estructura()` (campos clave: expediente, nombre, código, emisión impresa, fecha manuscrita) y `conciliar()` (filas con E = lote frente a fotos por expediente). Informes en JSON en la carpeta de resultados.
4. Procesamiento: PaddleOCR + TrOCR + Qwen2.5-VL (`LECTOR_VLM`) sobre el recuadro de fecha con tinta azul aislada; confianza ALTA/MEDIA/CONTRADICTORIA/SIN_FECHA y alertas. Todos los registros NUEVOS (o hasta `LIMITE`). `REINTENTAR_ERRORES` controla los errores anteriores. Guarda cada lectura y llama al exportador; si ya había aprobaciones recuperadas, puede producir una copia actualizada.
5. Revisión: muestra imagen, recortes, lecturas y campos; requiere nombre del revisor, cotejo y confirmación. `MODO = segunda` para el segundo revisor; la misma persona no puede firmar dos veces los mismos datos.
6. Exportación: revalida todas las aprobaciones, aplica `DOBLE_REVISION` (ninguna, no_alta, todas) y genera copia, CSV y acta JSON; con `ESCRIBIR_EN_EXCEL_DE_DRIVE` sustituye el Excel de trabajo de Drive con respaldo previo y rota la huella del proyecto.
7. Paquete opcional: hasta `CASOS` pendientes para revisión externa; no los envía automáticamente.

Estados: NUEVO, POR_REVISAR, PENDIENTE, APROBADO, ERROR. POR_REVISAR no es un fallo ni una aprobación. Menos imágenes únicas que archivos es normal si hay duplicados. Ninguna aprobación implica que todavía no hay copia actualizada; sí puede haber informe CSV.

## Validaciones

- Hoja 2026 y E = 02/09/2026; destino F vacía.
- Expediente con año y coincidencia única; código SIGETI de K coincidente; resolución UNO.
- Fecha válida, no futura y no anterior a la emisión.
- Original e imagen no cambiados desde la carga/aprobación.
- Fechas contradictorias para una misma celda bloquean su actualización.

## Supervisión y recuperación

Comprueba la ejecución real antes de iniciar celdas. No deduzcas que el servidor terminó solo porque la interfaz diga Conectando. No ejecutes sesiones paralelas contra la misma salida.

El usuario autorizó reconectar y ejecutar 1, 2 y 4 en orden, incluyendo una recarga de pestaña para recuperar la conexión. Si se pide autorización de Google o el Excel, avisa al usuario. No apruebes fechas ni ejecutes 5 y 6 por tu cuenta. No borres avances ni elimines entornos sin indicación.

Si falla, informa celda, error exacto, avance recuperado y acción propuesta. Distingue fallo de conexión, de instalación, de descarga de modelos, de lectura de imagen y falta de memoria. Si pierdes acceso a la pestaña, dilo: no afirmes supervisión continua sin acceso.

Estado conocido (24/10/2026): entorno Colab Pro con GPU T4 y Python 3.13; torch 2.11, paddlepaddle 3.2.0, paddleocr 3.3.2, transformers 4.57.6. La celda 2 funcionó: 3.996 fotos, 3.581 contenidos únicos, 415 copias idénticas, 10.701 expedientes indexados en la hoja 2026. PaddleOCR 3.2.0 se cae (segmentation fault) al inferir en ese entorno; EasyOCR en GPU funciona, por eso `MOTOR = auto` lo elige tras probar Paddle en un subproceso. La versión actual del cuaderno todavía no se ha ejecutado sobre fotos reales: ni la localización del recuadro, ni los lectores manuscritos, ni el acierto del piloto están medidos. No se ha confirmado que el OCR completo haya terminado ni su tasa de precisión.
