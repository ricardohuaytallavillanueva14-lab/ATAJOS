# Contexto para supervisión de Colab

Supervisa `REC_2026_02_SEPTIEMBRE_Colab.ipynb`. El objetivo es leer las 3.996 fotos y, tras cotejo visual, completar F vacía únicamente en la hoja 2026 cuando E sea 02/09/2026.

## Ubicaciones

- Fotos: `/content/drive/MyDrive/01_SAT_Huamanga/PROYECTO REC 1` y todas sus subcarpetas.
- Salida: `/content/drive/MyDrive/01_SAT_Huamanga/REC_2026_02SEP_RESULTADOS`.
- Original guardado: `entrada/original_expedientes.xlsx`, dentro de la salida.
- Configuración: `proyecto.json`.
- Avance: `registros/*.json` por SHA-256 del contenido fotográfico.
- Exportaciones: `expedientes_actualizados_*.xlsx` y `revision_*.csv`.

El original esperado es el XLSX adjuntado cuyo nombre termina en `(1).xlsx`. Su SHA-256 se comprueba en la celda 2; abrirlo y volver a guardarlo puede cambiar esa huella aunque sus valores parezcan iguales.

## Funcionamiento

1. Preparación: montaje de Drive, verificación de 3.996 fotos y del inventario si existe, instalación de dependencias.
2. Proyecto: carga de código, selección del Excel, índice e inventario completo. Lee el contenido para identificar duplicados.
3. Procesamiento: PaddleOCR + TrOCR, todos los registros NUEVOS. `REINTENTAR_ERRORES` controla los errores anteriores. Guarda cada lectura y llama al exportador; si ya había aprobaciones recuperadas, puede producir una copia actualizada.
4. Revisión: muestra imagen y campos; requiere nombre del revisor, cotejo y confirmación.
5. Exportación: revalida todas las aprobaciones y genera una copia y un CSV.
6. Paquete opcional: hasta 20 casos pendientes para revisión externa; no los envía automáticamente.

Estados: NUEVO, POR_REVISAR, PENDIENTE, APROBADO, ERROR. POR_REVISAR no es un fallo ni una aprobación. Menos imágenes únicas que archivos es normal si hay duplicados. Ninguna aprobación implica que todavía no hay copia actualizada; sí puede haber informe CSV.

## Validaciones

- Hoja 2026 y E = 02/09/2026; destino F vacía.
- Expediente con año y coincidencia única; código SIGETI de K coincidente; resolución UNO.
- Fecha válida, no futura y no anterior a la emisión.
- Original e imagen no cambiados desde la carga/aprobación.
- Fechas contradictorias para una misma celda bloquean su actualización.

## Supervisión y recuperación

Comprueba la ejecución real antes de iniciar celdas. No deduzcas que el servidor terminó solo porque la interfaz diga Conectando. No ejecutes sesiones paralelas contra la misma salida.

El usuario autorizó reconectar y ejecutar 1, 2 y 3 en orden, incluyendo una recarga de pestaña para recuperar la conexión. Si se pide autorización de Google o el Excel, avisa al usuario. No apruebes fechas ni ejecutes 4 y 5 por tu cuenta. No borres avances ni elimines entornos sin indicación.

Si falla, informa celda, error exacto, avance recuperado y acción propuesta. Distingue fallo de conexión, de instalación, de descarga de modelos, de lectura de imagen y falta de memoria. Si pierdes acceso a la pestaña, dilo: no afirmes supervisión continua sin acceso.

En el último reporte del usuario la celda 1 había encontrado 3.996 fotos sin diferencias y T4, pero la pestaña quedó en Conectando/Reanudando la ejecución. Este es un antecedente, no el estado actual: hay que comprobar los registros y las celdas al reconectar. No se ha confirmado desde este proyecto que el OCR completo haya terminado ni su tasa de precisión.
