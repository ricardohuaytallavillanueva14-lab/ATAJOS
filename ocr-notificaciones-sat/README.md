# OCR Notificaciones SAT

Extractor de datos de páginas escaneadas pensado para **Google Colab**. Lee fotografías de notificaciones (REC) guardadas en Google Drive, propone con OCR gratuito el expediente, el código SIGETI, la resolución y la **fecha manuscrita de notificación**, y, tras un cotejo visual obligatorio, escribe esa fecha en una **copia** del Excel de expedientes. El original nunca se modifica.

[![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/ricardohuaytallavillanueva14-lab/ATAJOS/blob/main/ocr-notificaciones-sat/notebooks/REC_2026_02_SEPTIEMBRE_Colab.ipynb)

## Cómo funciona

| Etapa | Qué hace |
| --- | --- |
| 1 · Preparación | Monta Drive, cuenta las fotos, instala solo lo que falta y guarda los modelos en Drive. |
| 2 · Excel e inventario | Carga el Excel (desde Drive o subido), fija hoja y fecha de emisión, inventaría las fotos por contenido. |
| 3 · Prueba con muestra | Procesa ~20 fotos repartidas por carpetas y estima el tiempo total. |
| 4 · Procesamiento | PaddleOCR (PP-OCRv5) + TrOCR manuscrito sobre todas las pendientes, con progreso y reanudación. |
| 5 · Revisión | Muestra encabezado, cargo y recorte de la fecha; el revisor corrige y confirma. |
| 6 · Exportación | Copia del XLSX con las fechas confirmadas y un informe CSV. |
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
- **Opciones de rendimiento**: `MODO_RAPIDO` (detector PP-OCRv5 mobile + reconocedor latino) y `PADDLE_GPU` (requiere `paddlepaddle-gpu`; vuelve a CPU si falla). TrOCR procesa las vistas del recorte en un solo lote.
- **Alcance configurable** con `configurar(hoja, fecha_emision)`, aplicado por igual al índice, la evaluación y la escritura del XLSX.
- **Compatibilidad**: las carpetas de resultados creadas con la versión anterior se abren sin cambios.

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

Las pruebas reproducibles usan un libro ficticio para comprobar alcance, fórmulas, duplicados, celdas ocupadas, caché del inventario y procesamiento por muestra con un motor simulado. La escritura se verificó antes sobre una copia del Excel real: solo cambió la celda objetivo y su formato.

El OCR completo del lote y su precisión no se han medido. Las puntuaciones del OCR no son probabilidades de acierto; cada propuesta exige cotejo visual. PaddleOCR corre en CPU salvo que se active `PADDLE_GPU`; TrOCR usa la GPU de Colab si existe. Fotografías giradas, otros diseños, AVISO PREVIO y CEDULON pueden requerir transcripción manual. Los límites de sesión de Colab siguen aplicando: el avance se guarda por foto y se reanuda.

No se incluye una licencia de redistribución: el titular puede añadir la que corresponda.
