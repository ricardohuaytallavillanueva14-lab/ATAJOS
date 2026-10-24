# OCR Notificaciones SAT

**Modo actual (sin Excel de entrada):** el cuaderno lee las fotografías y crea un **Excel nuevo** (`extraccion_*.xlsx`) con una fila por expediente y solo los datos importantes: expediente, nombre y apellidos, código de contribuyente, fecha de emisión y fecha de notificación manuscrita, más verificación, confianza y alertas. El modo con Excel de expedientes (validar contra sus filas y completar la columna F en una copia) sigue disponible con `USAR_EXCEL = True`.

Extractor de datos de páginas escaneadas pensado para **Google Colab**. Lee fotografías de notificaciones (REC) guardadas en Google Drive, propone con OCR gratuito el expediente, el código SIGETI, la resolución y la **fecha manuscrita de notificación**, y, tras un cotejo visual obligatorio, escribe esa fecha en una **copia** del Excel de expedientes. El original nunca se modifica.

[![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/ricardohuaytallavillanueva14-lab/ATAJOS/blob/claude/hola-dec1ew/ocr-notificaciones-sat/notebooks/REC_2026_02_SEPTIEMBRE_Colab.ipynb)

El botón abre el cuaderno de la rama `claude/hola-dec1ew`, donde vive el proyecto; en `main` todavía no existe. Cuando la rama se fusione, cambia `claude/hola-dec1ew` por `main` en este enlace y en el README de la raíz.

## Qué extrae de cada fotografía

Solo los campos necesarios para relacionar la foto con su fila y completar F. Todo lo demás del documento se ignora.

| Campo | Dónde está en la REC | Uso |
| --- | --- | --- |
| Expediente coactivo (con año) | Encabezado impreso | Clave de la fila: debe existir exactamente una coincidencia en la hoja. |
| Nombre y apellidos del obligado | Encabezado impreso, «OBLIGADO (A)» | Se compara por palabras con el nombre de la fila; si no se parecen, alerta. |
| Código de contribuyente | Encabezado impreso | Debe coincidir con la columna K (SIGETI); si no, bloquea. |
| Fecha de emisión | Impresa, «Ayacucho, 02 de setiembre de 2026» | Debe coincidir con la columna E de la fila; si no, bloquea (la foto es de otro lote). |
| **Fecha de notificación** | **Manuscrita**, recuadro del cargo | El dato que se escribe en F. Es el único manuscrito: tres lectores, tinta aislada, confianza y cotejo visual. |
| Número de resolución | Impreso, «Resolución N° UNO» | Apoyo: si se lee otra distinta de UNO bloquea; si no se lee, el revisor lo confirma con una casilla. |

El sondeo de formato (celda 3, `SONDEO_FORMATO`) recorre todas las fotos sin OCR y resume orientación, tamaño y carpeta; el sondeo de estructura (tras el piloto y tras el lote) indica cuántas fotos leídas son REC completas con los cinco campos y el recuadro de fecha localizado. La conciliación compara las filas con E = fecha del lote frente a las fotos, por expediente: filas con foto, sin foto, con varias fotos, fotos de expedientes fuera del lote y fotos sin expediente legible. Los tres informes se guardan como JSON en la carpeta de resultados.

## Cómo funciona

| Etapa | Qué hace |
| --- | --- |
| 1 · Preparación | Monta Drive, cuenta las fotos, instala solo lo que falta y guarda los modelos en Drive. |
| 2 · Excel e inventario | Carga el Excel (desde Drive o subido), fija hoja y fecha de emisión, inventaría las fotos por contenido. |
| 3 · Prueba con muestra | Procesa ~20 fotos repartidas por carpetas y estima el tiempo total. |
| 4 · Procesamiento | Tres lectores independientes (PaddleOCR PP-OCRv5, TrOCR manuscrito y Qwen2.5-VL) sobre el recuadro de fecha, con tinta azul aislada, nivel de confianza y alertas. |
| 5 · Revisión | Muestra encabezado, cargo, recorte de la fecha y manuscrito aislado; el revisor corrige y confirma. Modo de segunda revisión. |
| 6 · Exportación | Excel nuevo con los datos importantes (hojas Expedientes, Fotos y Resumen) y acta con su huella SHA-256. Con Excel de expedientes, además la copia con F completada, el CSV y su acta. |
| 7 · Opcional | ZIP con casos pendientes para una segunda revisión externa. |

Alcance por defecto: hoja `2026`, filas con columna E (emisión REC) = `02/09/2026` y columna F (notificación REC) vacía. Ambos valores se cambian en la celda 2 y quedan registrados en `proyecto.json` de la carpeta de resultados, de modo que una carpeta no se mezcla con otro Excel ni otro lote.

## Excel nuevo (`exportar_extraccion`)

- **Hoja Expedientes**: una fila por expediente del lote. Columnas: Expediente, Nombre y apellidos, Código contribuyente, Fecha emisión, Fecha notificación, Resolución, Verificación («Confirmada por …», «Propuesta OCR sin verificar», «Pendiente», «CONFLICTO: fotos con fechas distintas»), Revisor(es), Confirmada el, Confianza OCR, Alertas, Problemas, Fotos, Archivo(s), Huella SHA-256. Si un expediente tiene varias fotos, prevalece la confirmada; si las fechas no coinciden, la celda queda vacía y se marca CONFLICTO.
- **Hoja Fotos**: una fila por foto única, con la misma información más el tipo de documento y la hora de lectura. Incluye anexos, errores y documentos de otros lotes.
- **Hoja Resumen**: cantidades por tipo de verificación.
- `SOLO_CONFIRMADAS` deja en Expedientes solo las filas con cotejo humano; `FILTRAR_POR_EMISION` excluye los documentos cuya emisión impresa no es la del lote. Las fechas son celdas de fecha con formato dd/mm/yyyy.
- Sin Excel de entrada, el lote se delimita por el año del expediente y la fecha de emisión impresa en cada REC; no hay validación contra la columna K ni medición de acierto contra F, por eso el cotejo visual de la celda 5 sigue siendo la garantía.

## Uso en Colab

1. Pulsa **Abrir en Colab** y guarda una copia en tu Drive si quieres conservar cambios.
2. Elige *Entorno de ejecución > Cambiar tipo de entorno de ejecución > GPU T4*.
3. Rellena los formularios de cada celda (rutas de Drive, cantidad de fotos, hoja, fecha de emisión) y ejecútalas en orden.
4. Tras la celda 3, mira los segundos por foto y la estimación de horas. Si supera tu sesión de Colab, usa `LIMITE` en la celda 4 para avanzar por lotes o activa `MODO_RAPIDO`.
5. Revisa en la celda 5 y exporta en la 6 tantas veces como necesites.

El Excel de trabajo es el archivo guardado en Drive cuya huella SHA-256 está en `EXCEL_SHA_ESPERADO` (celda 2); la celda lo copia a `entrada/` y comprueba la huella. Si usas otro lote, cambia o vacía ese valor.

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
- **Motor automático**: `MOTOR = auto` comprueba PaddleOCR en un subproceso (una caída no afecta al cuaderno) y usa EasyOCR en GPU si falla, como ocurre con Python 3.13 en Colab; la celda 1 retira `langchain`, que rompe `paddlex`.
- **Opciones de rendimiento**: `MODO_RAPIDO` (detector PP-OCRv5 mobile + reconocedor latino), `PADDLE_GPU` (requiere `paddlepaddle-gpu`; vuelve a CPU si falla) y `ORIENTACION_AUTO` para fotos giradas. TrOCR procesa las vistas del recorte en un solo lote.
- **Alcance configurable** con `configurar(hoja, fecha_emision)`, aplicado por igual al índice, la evaluación y la escritura del XLSX.
- **Compatibilidad**: las carpetas de resultados creadas con la versión anterior se abren sin cambios.

## Confiabilidad y uso legal

El sistema está diseñado para que **cada fecha escrita en el Excel sea atribuible a una persona que la cotejó con la fotografía**, y para que esa comprobación pueda demostrarse después. El OCR nunca escribe por sí solo: propone, ordena y documenta.

Cómo se lee la fecha manuscrita (celda 4):

1. El detector impreso (PaddleOCR o, si este no puede inferir en el entorno, EasyOCR) localiza los rótulos «FECHA DE NOTIFICACIÓN» y «HORA DE NOTIFICACIÓN» por coincidencia aproximada (tolera lecturas como «HotECACLON» o rótulos partidos) y recorta solo ese recuadro, a resolución completa. Las demás fechas impresas del documento quedan fuera del recorte.
2. Del recuadro se extraen tres vistas: color, contraste y **tinta azul aislada**. Esta última conserva únicamente los trazos de bolígrafo y elimina el texto impreso y las líneas de la tabla, que son la principal causa de lecturas erróneas en manuscritos.
3. Tres lectores independientes leen las vistas: el detector impreso, TrOCR manuscrito (Microsoft) y, si `LECTOR_VLM` está activo, el modelo de visión-lenguaje Qwen2.5-VL-3B con una instrucción que prohíbe inventar cifras. La fecha escrita en el acta («Siendo las … horas, del día __/__/__») se extrae aparte y **solo genera una alerta** si difiere de la del recuadro: no entra en la propuesta ni en la confianza. Motivo: en las actas bajo puerta (DOMICILIO CERRADO) esa línea lleva la fecha del aviso previo, normalmente el día anterior, así que una diferencia puede ser legítima; en NEGATIVA debería coincidir. Como la regla no se ha confirmado sobre fotos reales, se deja al revisor.
4. Las lecturas se comparan. **ALTA**: dos o más lectores leen exactamente la misma fecha y ninguno otra. **MEDIA**: solo un lector obtuvo fecha. **CONTRADICTORIA**: lecturas distintas. **SIN_FECHA**: nadie leyó una fecha completa. Nunca se completan dígitos ni se adivina el año.
5. Alertas de plausibilidad: notificación más de `VENTANA_DIAS` (60) días después de la emisión, en domingo, o resolución no leída (se confirma visualmente que es la REC UNO). No bloquean, pero se muestran al revisor y quedan en el informe.
6. Tipos de documento: las REC completas se leen; AVISO PREVIO y CEDULÓN se marcan como anexos sin fecha para F; las medias páginas (encabezado en una foto y cargo en otra) se emparejan con la foto anterior de la carpeta y el revisor las confirma viendo ambas.

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

Límites que conviene declarar en cualquier uso formal: la garantía proviene del cotejo humano y de la trazabilidad, no del OCR. La celda 3 mide el acierto del OCR contra las filas que ya tienen F registrada en el Excel (patrón de referencia) sobre una muestra de 80 fotos antes de la pasada completa.

## Licencias de los modelos

| Modelo | Licencia | Uso en este proyecto |
| --- | --- | --- |
| PaddleOCR PP-OCRv5 | Apache-2.0 | Detector impreso |
| EasyOCR | Apache-2.0 | Detector impreso alternativo |
| microsoft/trocr-base-handwritten | MIT | Lector manuscrito |
| Qwen/Qwen2.5-VL-3B-Instruct | **Qwen Research License Agreement** (Alibaba Cloud): uso de investigación y evaluación no comercial; el uso comercial requiere licencia aparte. Es distinta de la Apache-2.0 del modelo de 7B. | Lector manuscrito opcional (`LECTOR_VLM`) |

La licencia del 3B se tomó de fuentes secundarias que citan el archivo `LICENSE` del repositorio del modelo (no fue posible abrir huggingface.co desde el entorno donde se escribió esto): confírmala en https://huggingface.co/Qwen/Qwen2.5-VL-3B-Instruct/blob/main/LICENSE antes de usarlo en un procedimiento institucional. Si el uso no encaja en esa licencia, desactiva `LECTOR_VLM` (la confianza ALTA deja de ser alcanzable y todas las fechas se cotejan igual) o sustituye el modelo por `Qwen2.5-VL-7B-Instruct` (Apache-2.0), que en una T4 requiere cuantización a 4 bits y no está integrado. Si el VLM no carga en fp16 en la T4, el cuaderno captura el error, continúa sin él y lo avisa en pantalla.

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
