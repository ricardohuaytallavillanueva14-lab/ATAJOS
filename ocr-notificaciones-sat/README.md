# OCR Notificaciones SAT

Extractor de datos de páginas escaneadas pensado para **Google Colab**. Lee fotografías de notificaciones (REC) guardadas en Google Drive, propone con OCR gratuito el expediente, el código SIGETI, la resolución y la **fecha manuscrita de notificación**, y, tras un cotejo visual obligatorio, escribe esa fecha en una **copia** del Excel de expedientes. El original nunca se modifica.

[![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/ricardohuaytallavillanueva14-lab/ATAJOS/blob/main/ocr-notificaciones-sat/notebooks/REC_2026_02_SEPTIEMBRE_Colab.ipynb)

## Cómo funciona

| Etapa | Qué hace |
| --- | --- |
| 1 · Preparación | Monta Drive, cuenta las fotos, instala solo lo que falta y guarda los modelos en Drive. |
| 2 · Excel e inventario | Carga el Excel (desde Drive o subido), fija hoja y fecha de emisión, inventaría las fotos por contenido. |
| 3 · Prueba con muestra | Procesa ~20 fotos repartidas por carpetas y estima el tiempo total. |
| 4 · Procesamiento | Tres lectores independientes (PaddleOCR PP-OCRv5, TrOCR manuscrito y Qwen2.5-VL) sobre el recuadro de fecha, con tinta azul aislada, nivel de confianza y alertas. |
| 5 · Revisión | Muestra encabezado, cargo, recorte de la fecha y manuscrito aislado; el revisor corrige y confirma. Modo de segunda revisión. |
| 6 · Exportación | Copia del XLSX con las fechas confirmadas, informe CSV y acta de exportación con huellas SHA-256. |
| 7 · Opcional | ZIP con casos pendientes para una segunda revisión externa. |

Alcance por defecto: hoja `2026`, filas con columna E (emisión REC) = `02/09/2026` y columna F (notificación REC) vacía. Ambos valores se cambian en la celda 2 y quedan registrados en `proyecto.json` de la carpeta de resultados, de modo que una carpeta no se mezcla con otro Excel ni otro lote.

## Uso en Colab

1. Pulsa **Abrir en Colab** y guarda una copia en tu Drive si quieres conservar cambios.
2. Elige *Entorno de ejecución > Cambiar tipo de entorno de ejecución > GPU T4*.
3. Rellena los formularios de cada celda (rutas de Drive, cantidad de fotos, hoja, fecha de emisión) y ejecútalas en orden.
4. Tras la celda 3, mira los segundos por foto y la estimación de horas. Si supera tu sesión de Colab, usa `LIMITE` en la celda 4 para avanzar por lotes o activa `MODO_RAPIDO`.
5. Revisa en la celda 5 y exporta en la 6 tantas veces como necesites.

Rutas por defecto en Drive:

- Fotos: `/content/drive/MyDrive/01_SAT_Huamanga/PROYECTO REC 1`
- Resultados: `/content/drive/MyDrive/01_SAT_Huamanga/REC_2026_02SEP_RESULTADOS`
- Caché de modelos: `/content/drive/MyDrive/01_SAT_Huamanga/modelos_ocr_cache`

El repositorio solo contiene software y documentación. Fotografías, Excel, lecturas y aprobaciones viven en Drive. No agregues datos de contribuyentes al repositorio.

## Mejoras para Colab

- **Formularios** (`# @param`) en todas las celdas: sin editar código para cambiar rutas, lote o límites.
- **Instalación inteligente**: solo instala paquetes ausentes o con otra versión; respeta el `torch` con CUDA de Colab y avisa si hace falta reiniciar.
- **Caché de modelos en Drive** (`HF_HOME`, `PADDLE_PDX_CACHE_HOME`): TrOCR y PaddleOCR no se descargan en cada sesión.
- **Caché de inventario** (`inventario_cache.json`): al reanudar, las huellas SHA-256 se reutilizan si tamaño y fecha coinciden; los registros sin cambios no se reescriben en Drive.
- **Una sola lectura por foto**: la huella se calcula sobre los mismos bytes que recibe el OCR.
- **Progreso y estimaciones**: barra `tqdm`, segundos por foto, horas restantes y resumen por lote.
- **Celda de prueba** con muestra reproducible repartida entre carpetas.
- **Opciones de rendimiento**: `MODO_RAPIDO` (detector PP-OCRv5 mobile + reconocedor latino), `PADDLE_GPU` (requiere `paddlepaddle-gpu`; vuelve a CPU si falla) y `ORIENTACION_AUTO` para fotos giradas. TrOCR procesa las vistas del recorte en un solo lote.
- **Alcance configurable** con `configurar(hoja, fecha_emision)`, aplicado por igual al índice, la evaluación y la escritura del XLSX.
- **Compatibilidad**: las carpetas de resultados creadas con la versión anterior se abren sin cambios.

## Confiabilidad y uso legal

El sistema está diseñado para que **cada fecha escrita en el Excel sea atribuible a una persona que la cotejó con la fotografía**, y para que esa comprobación pueda demostrarse después. El OCR nunca escribe por sí solo: propone, ordena y documenta.

Cómo se lee la fecha manuscrita (celda 4):

1. PaddleOCR localiza el rótulo impreso «FECHA DE NOTIFICACIÓN» y recorta solo ese recuadro, hasta «HORA DE NOTIFICACIÓN». Las demás fechas impresas del documento (emisión, títulos ejecutivos) quedan fuera del recorte.
2. Del recuadro se extraen tres vistas: color, contraste y **tinta azul aislada**. Esta última conserva únicamente los trazos de bolígrafo y elimina el texto impreso y las líneas de la tabla, que son la principal causa de lecturas erróneas en manuscritos.
3. Tres lectores independientes leen las vistas: PaddleOCR (detección de texto), TrOCR manuscrito (Microsoft) y, si `LECTOR_VLM` está activo, el modelo de visión-lenguaje Qwen2.5-VL-3B con una instrucción que prohíbe inventar cifras.
4. Las lecturas se comparan. **ALTA**: dos o más lectores leen exactamente la misma fecha y ninguno otra. **MEDIA**: solo un lector obtuvo fecha. **CONTRADICTORIA**: lecturas distintas. **SIN_FECHA**: nadie leyó una fecha completa. Nunca se completan dígitos ni se adivina el año.
5. Alertas de plausibilidad: notificación más de `VENTANA_DIAS` (60) días después de la emisión, o en domingo. No bloquean, pero se muestran al revisor y quedan en el informe.

Controles antes de escribir (celdas 5 y 6):

- Cotejo visual obligatorio con nombre del revisor, casilla de conformidad y fecha-hora de la aprobación.
- Validación cruzada con el Excel: expediente con año único, código SIGETI coincidente, resolución UNO, hoja y emisión del lote, celda F vacía, fecha no futura ni anterior a la emisión.
- **Doble revisión** (`DOBLE_REVISION` en la celda 6): `no_alta` exige un segundo revisor distinto cuando la confianza OCR no fue ALTA; `todas` lo exige siempre. Una corrección posterior invalida las aprobaciones anteriores y vuelve a pedir la segunda firma.
- Dos fotos aprobadas con fechas distintas para la misma celda bloquean esa celda.

Cadena de evidencia que queda en Drive:

- `registros/<sha256>.json`: huella de la foto, todas las lecturas de cada lector, recorte usado, confianza, alertas y la lista completa de aprobaciones con revisor y hora.
- `evidencia/<sha256>_fecha.jpg` y `_cargo.jpg`: lo que vio el revisor al aprobar.
- `revision_*.csv`: una fila por foto con resultado, revisores, confianza, lectura OCR, alertas y evidencia.
- `acta_exportacion_*.json`: huellas SHA-256 del Excel original y de la copia generada, alcance, política de doble revisión y detalle de cada celda escrita.
- La copia del Excel solo cambia las celdas aprobadas y su formato de fecha; el resto del archivo se conserva byte a byte (verificado por las pruebas).

Límites que conviene declarar en cualquier uso formal: la exactitud de los lectores no se ha medido sobre este lote; la garantía proviene del cotejo humano y de la trazabilidad, no del OCR. Con la celda 3 puede estimarse la tasa de coincidencia en una muestra antes de la pasada completa.

## Código y mantenimiento

- `src/ocr_sat.py`: extracción, inventario, revisión, validación y escritura del XLSX.
- `templates/`: código de las siete celdas y guía incorporada. La celda 2 incrusta `src/ocr_sat.py`.
- `scripts/build_notebook.py`: reconstruye el cuaderno autocontenido.
- `notebooks/`: cuaderno listo para Colab (sin salidas).
- `tests/test_excel.py`: pruebas con datos ficticios, sin descargar modelos.
- `docs/CONTEXTO_CLAUDE.md`: contexto de supervisión y recuperación.

Desde esta carpeta:

```bash
python -m pip install -r requirements-test.txt
python -m unittest discover -s tests -v
python scripts/build_notebook.py
```

Tras cambiar cualquier plantilla o `src/ocr_sat.py`, ejecuta `build_notebook.py` y confirma el `.ipynb`: una prueba y el flujo de GitHub Actions comprueban que el cuaderno coincide con las plantillas.

## Validación y límites

Las pruebas reproducibles usan un libro ficticio para comprobar alcance, fórmulas, duplicados, celdas ocupadas, caché del inventario, procesamiento por muestra con un motor simulado, aislamiento de tinta azul, niveles de confianza, alertas, doble revisión, evidencia y acta de exportación. La escritura se verificó antes sobre una copia del Excel real: solo cambió la celda objetivo y su formato.

El OCR completo del lote y su precisión no se han medido. Las puntuaciones del OCR no son probabilidades de acierto; cada propuesta exige cotejo visual. PaddleOCR corre en CPU salvo que se active `PADDLE_GPU`; TrOCR y Qwen2.5-VL usan la GPU de Colab si existe. La integración de Qwen2.5-VL sigue la API documentada de transformers 4.57 pero no pudo ejecutarse desde este repositorio: si falla la carga, desactiva `LECTOR_VLM` y el resto sigue funcionando. Fotografías giradas, otros diseños, AVISO PREVIO y CEDULON pueden requerir transcripción manual. Los límites de sesión de Colab siguen aplicando: el avance se guarda por foto y se reanuda.

No se incluye una licencia de redistribución: el titular puede añadir la que corresponda.
