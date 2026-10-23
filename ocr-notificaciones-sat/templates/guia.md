# Notificaciones REC — hoja 2026 · emisión 02/09/2026

Cuaderno preparado para tu Excel y las **3.996 fotografías** de Drive, incluidas todas las subcarpetas. Usa un entorno **alojado de Colab**, preferentemente GPU T4; no necesitas PowerShell ni conectar el entorno local.

**Alcance fijo:** E es «fecha emision REC» y se usa como filtro. La fecha manuscrita de notificación se escribe en **F**, «Ffecha de notificacion REC». Solo hoja **2026**, filas con **E = 02/09/2026** y F vacía. El archivo revisado tiene **3.401 filas** con esa emisión y **1.950 notificaciones vacías**. Las 1.451 ya registradas se conservan. No se modifican otras hojas, filas, columnas ni el original.

## Ejecutar en orden
1. **Celda 1:** monta Drive, comprueba las fotos e instala las dependencias. Autoriza la cuenta que contiene `01_SAT_Huamanga/PROYECTO REC 1`.
2. **Celda 2:** solicita subir **solo el Excel original adjuntado en ChatGPT**, cuyo nombre termina en `(1).xlsx`. Se guarda en Drive para reanudar sin volver a subirlo. Verifica el archivo por su contenido y prepara el inventario completo.
3. **Celda 3:** procesa todas las fotos con PaddleOCR + TrOCR gratuitos. Guarda cada lectura. No utiliza una API de pago. No cierres voluntariamente el entorno durante el procesamiento.
4. **Celda 4:** revisa la imagen y corrige expediente con año, código, resolución y fecha. Escribe tu nombre/iniciales, marca la casilla y pulsa **Confirmar y guardar**. Los casos del filtro aparecen primero. Pulsa Siguiente para continuar; las lecturas ilegibles quedan pendientes. Solo se admite REC UNO con expediente/año único y código SIGETI coincidente.
5. **Celda 5:** crea una **copia actualizada del Excel** y un informe CSV en `Mi unidad/01_SAT_Huamanga/REC_2026_02SEP_RESULTADOS`. Puedes exportar por etapas. Solo escribe fechas confirmadas, sin conflictos; el OCR por sí solo no autoriza una modificación.

## Reanudar y verificar
Si Colab se desconecta, vuelve a ejecutar las celdas 1 y 2 y luego la 3: las lecturas y aprobaciones guardadas se recuperan. Se repite el inventario para detectar cambios en las imágenes; las copias idénticas se leen una sola vez y se conservan sus rutas. No ejecutes dos sesiones sobre la misma salida. Si hubo errores, cambia `REINTENTAR_ERRORES` a `True` en la celda 3. Cada error queda registrado y no impide continuar con otras imágenes, salvo falta de memoria.

Los límites de sesión de Colab siguen aplicando. La primera instalación y descarga de modelos requiere internet. Si Colab solicita reiniciar tras instalar, reinicia y ejecuta de nuevo desde la celda 1.

**La escritura manuscrita puede leerse mal.** Una coincidencia entre motores no constituye una tasa de precisión medida. AVISO PREVIO y CEDULON requieren cotejar el documento y la fecha pertinente. El recorte automático está orientado al encabezado y cargo inferior; otros diseños o fotos giradas pueden exigir transcripción manual. No se adivinan fechas ni expedientes. Si dos fotos confirmadas dan fechas distintas para una misma celda, esa celda queda sin modificar.

Puedes pedirme una segunda revisión adjuntando los casos dudosos. La última celda opcional prepara un ZIP de hasta 20 casos con fotografías; no envía datos a ChatGPT automáticamente. Las correcciones se confirman en la celda 4.

Se probó el filtro y la exportación sobre una copia del Excel real, verificando que se conservaran las demás partes del archivo. **No se ha ejecutado el OCR de las 3.996 fotos desde este chat ni medido su exactitud.**

Referencias: [PaddleOCR](https://www.paddleocr.ai/latest/en/quick_start.html), [TrOCR](https://huggingface.co/docs/transformers/model_doc/trocr), [Colab](https://research.google.com/colaboratory/faq.html).
