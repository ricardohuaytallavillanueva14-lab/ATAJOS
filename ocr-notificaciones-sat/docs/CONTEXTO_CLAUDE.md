# Contexto para supervisión de Colab

Supervisa `REC_2026_02_SEPTIEMBRE_Colab.ipynb`. El objetivo es leer las 3.996 fotos y, tras cotejo visual, completar F vacía únicamente en la hoja 2026 cuando E sea 02/09/2026.

## Ubicaciones

- Fotos: `/content/drive/MyDrive/01_SAT_Huamanga/PROYECTO REC 1` y todas sus subcarpetas.
- Salida: `/content/drive/MyDrive/01_SAT_Huamanga/REC_2026_02SEP_RESULTADOS`.
- Original guardado: `entrada/original_expedientes.xlsx`, dentro de la salida.
- Configuración: `proyecto.json` (Excel, hoja y fecha de emisión del lote).
- Caché de huellas: `inventario_cache.json` (tamaño y fecha por foto; `REHACER_INVENTARIO` la ignora).
- Modelos: `modelos_ocr_cache/` junto a la carpeta de resultados (TrOCR y PaddleOCR).
- Avance: `registros/*.json` por SHA-256 del contenido fotográfico.
- Exportaciones: `expedientes_actualizados_*.xlsx` y `revision_*.csv`.

El original esperado es el XLSX adjuntado cuyo nombre termina en `(1).xlsx`. Su SHA-256 se comprueba en la celda 2; abrirlo y volver a guardarlo puede cambiar esa huella aunque sus valores parezcan iguales.

## Funcionamiento

1. Preparación: montaje de Drive, verificación de las fotos (`FOTOS_ESPERADAS`, 0 = sin comprobar) y del inventario XLSX si existe, instalación solo de lo que falta, caché de modelos en Drive.
2. Proyecto: carga de código, `configurar(HOJA_LOTE, FECHA_EMISION_LOTE)`, Excel desde Drive o subido, índice e inventario (reutiliza la caché de huellas).
3. Prueba: muestra de `MUESTRA` fotos repartidas por carpetas; informa segundos por foto y horas estimadas.
4. Procesamiento: PaddleOCR + TrOCR, todos los registros NUEVOS (o hasta `LIMITE`). `REINTENTAR_ERRORES` controla los errores anteriores. Guarda cada lectura y llama al exportador; si ya había aprobaciones recuperadas, puede producir una copia actualizada.
5. Revisión: muestra imagen y campos; requiere nombre del revisor, cotejo y confirmación.
6. Exportación: revalida todas las aprobaciones y genera una copia y un CSV.
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

En el último reporte del usuario la celda 1 había encontrado 3.996 fotos sin diferencias y T4, pero la pestaña quedó en Conectando/Reanudando la ejecución. Este es un antecedente, no el estado actual: hay que comprobar los registros y las celdas al reconectar. No se ha confirmado desde este proyecto que el OCR completo haya terminado ni su tasa de precisión.
