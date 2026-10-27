# OCR Notificaciones SAT

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
| 6 · Exportación | Copia del XLSX con las fechas confirmadas, informe CSV y acta con huellas SHA-256. Opcionalmente lleva esas fechas al Excel de trabajo de Drive, con respaldo previo. |
| 7 · Opcional | ZIP con casos pendientes para una segunda revisión externa. |

Alcance por defecto: hoja `2026`, filas con columna E (emisión REC) = `02/09/2026` y columna F (notificación REC) vacía. Ambos valores se cambian en la celda 2 y quedan registrados en `proyecto.json` de la carpeta de resultados, de modo que una carpeta no se mezcla con otro Excel ni otro lote.

## Uso en Colab

1. Pulsa **Abrir en Colab** y guarda una copia en tu Drive si quieres conservar cambios.
2. Elige *Entorno de ejecución > Cambiar tipo de entorno de ejecución > GPU T4*.
3. Rellena los formularios de cada celda (rutas de Drive, cantidad de fotos, hoja, fecha de emisión) y ejecútalas en orden.
4. Tras la celda 3, mira los segundos por foto y la estimación de horas. Si supera tu sesión de Colab, usa `LIMITE` en la celda 4 para avanzar por lotes o activa `MODO_RAPIDO`.
5. Revisa en la celda 5 y exporta en la 6 tantas veces como necesites.

El Excel de trabajo es el archivo `.xlsx` guardado en Drive dentro de la carpeta de fotos (`PROYECTO REC 1`), el mismo que se abre desde Drive. La celda 2 lo localiza sola si es el único `.xlsx` de esa carpeta (sin contar `inventario_fotos.xlsx` ni los respaldos), lo copia a `entrada/` y liga la carpeta de resultados a su huella SHA-256; `RUTA_EXCEL_EN_DRIVE` permite indicarlo a mano y `EXCEL_SHA_ESPERADO` exigir una huella concreta.

Para que las fechas confirmadas queden en ese mismo Excel de Drive, activa `ESCRIBIR_EN_EXCEL_DE_DRIVE` en la celda 6: se genera primero la copia verificada en resultados, se guarda un respaldo `<nombre>_respaldo_<fecha>.xlsx` junto al Excel de trabajo, se comprueba que nadie lo cambió desde que se cargó y solo entonces se sustituye. El proyecto pasa a trabajar con la versión actualizada (nueva huella en `proyecto.json`, con historial), así que las exportaciones siguientes solo añaden fechas nuevas. Mientras Colab escribe, no edites ese Excel en Google Sheets.

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
- **PaddleOCR en un entorno aislado**: paddlepaddle 3.2 se cae (segmentation fault) al inferir con el Python 3.13 de Colab. La celda 1 crea `/content/paddle_env` con Python 3.11 mediante `uv`, instala allí PaddleOCR y el cuaderno lo ejecuta en un proceso aparte (`PaddleWorker`) que recibe cada recorte y devuelve las detecciones. En el proceso aislado Paddle corre con oneDNN (`enable_mkldnn=True`, que evita la ruta Im2Col del conv en CPU donde cayó en Colab) y con el lado mayor del detector limitado a 960 px. Si aun así el proceso se cae, el cuaderno no muere: pasa a EasyOCR para el resto de la pasada sin reintentar Paddle y muestra las últimas líneas de su registro. `MOTOR = auto` prueba primero el entorno aislado, después PaddleOCR en el propio proceso y por último EasyOCR. La celda 1 también retira `langchain`, que rompe `paddlex`.
- **Opciones de rendimiento**: `MODO_RAPIDO` (detector PP-OCRv5 mobile + reconocedor latino), `PADDLE_GPU` (requiere `paddlepaddle-gpu`; vuelve a CPU si falla) y `ORIENTACION_AUTO` para fotos giradas. TrOCR procesa las vistas del recorte en un solo lote.
- **Alcance configurable** con `configurar(hoja, fecha_emision)`, aplicado por igual al índice, la evaluación y la escritura del XLSX.
- **Compatibilidad**: las carpetas de resultados creadas con la versión anterior se abren sin cambios.

## Confiabilidad y uso legal

El sistema está diseñado para que **cada fecha escrita en el Excel sea atribuible a una persona que la cotejó con la fotografía**, y para que esa comprobación pueda demostrarse después. El OCR nunca escribe por sí solo: propone, ordena y documenta.

Cómo se lee la fecha manuscrita (celda 4):

1. El detector impreso (PaddleOCR o, si este no puede inferir en el entorno, EasyOCR) localiza los rótulos «FECHA DE NOTIFICACIÓN» y «HORA DE NOTIFICACIÓN» por coincidencia aproximada (tolera lecturas como «HotECACLON» o rótulos partidos) y recorta solo ese recuadro, a resolución completa. Si solo se reconoce la palabra «NOTIFICACIÓN» en la mitad inferior, también sirve de ancla: con un rótulo HORA debajo es el de FECHA, con FOLIOS debajo es el de HORA, y si no se sabe se leen las dos casillas (encima y debajo) y las lecturas que no son fechas se descartan solas. Se descarta si la línea menciona EMISIÓN o VENCIMIENTO.
   - Un «NOTIFICACIÓN» suelto solo vale como ancla si es una línea corta (rótulo de tabla, no un párrafo), está en la mitad inferior, por encima de «TÉNGASE POR BIEN NOTIFICADO», no es otro rótulo («ACTA DE NOTIFICACIÓN», «PERSONAL QUE NOTIFICA») y está en la columna central del cargo o tiene HORA o FOLIOS debajo; entre varias, gana la que tiene HORA o FOLIOS debajo y, si no, la más baja. El recorte entre FECHA y HORA tiene una altura mínima de 2,5 rótulos para no cortar la fecha. Si en la primera pasada no aparece ningún rótulo, el detector repite sobre la mitad derecha del cargo ampliada al doble, porque esos rótulos son pequeños.
   - **Respaldo por posición.** Si no hay ningún rótulo, o si el recuadro anclado no dio ninguna fecha manuscrita, se recorta la franja del cargo por posición relativa en el formato REC, se separan las líneas con tinta azul y se leen de arriba abajo con TrOCR y el VLM, deteniéndose en la primera línea donde un lector manuscrito lee una fecha. Estas lecturas se marcan `por_posicion` y **nunca alcanzan ALTA por sí solas**, porque el recorte puede contener otras cifras del cargo. Si tampoco así aparece una fecha, se pide al VLM que lea solo el valor junto a «FECHA DE NOTIFICACIÓN» en la mitad inferior completa (`vlm_mitad_inferior`, también `por_posicion`).
2. Del recuadro se extraen tres vistas: color, contraste y **tinta azul aislada**. Esta última conserva únicamente los trazos de bolígrafo y elimina el texto impreso y las líneas de la tabla, que son la principal causa de lecturas erróneas en manuscritos.
3. Tres lectores independientes leen el recorte: el detector impreso (una vista: la tinta aislada), TrOCR manuscrito (tinta y color) y, si `LECTOR_VLM` está activo, Qwen2.5-VL-3B en color y, con `VLM_DOS_VISTAS`, también en contraste, con una instrucción que prohíbe inventar cifras. Cuando hay dos recuadros candidatos, el segundo solo se lee si el primero no dio fecha. La fecha escrita en el acta («Siendo las … horas, del día __/__/__») se extrae aparte y **solo genera una alerta** si difiere de la del recuadro: no entra en la propuesta ni en la confianza. Motivo: en las actas bajo puerta (DOMICILIO CERRADO) esa línea lleva la fecha del aviso previo, normalmente el día anterior, así que una diferencia puede ser legítima; en NEGATIVA debería coincidir. Como la regla no se ha confirmado sobre fotos reales, se deja al revisor.
4. Las fechas anteriores a la emisión o más de 120 días posteriores (p. ej. años 2000, 2005 o 2021 leídos por error) no se proponen ni crean contradicción; quedan en `fechas_implausibles`. Las lecturas restantes se comparan. Si dos lectores manuscritos coinciden y solo el detector impreso discrepa, prevalecen los manuscritos (el detector suele leer el «1» como «/»); las lecturas del detector con un dígito inicial perdido se descartan. **ALTA**: dos o más lectores leen exactamente la misma fecha y ninguno otra. **MEDIA**: solo un lector obtuvo fecha. **CONTRADICTORIA**: lecturas distintas. **SIN_FECHA**: nadie leyó una fecha completa. Nunca se completan dígitos ni se adivina el año.
5. Alertas de plausibilidad: notificación más de `VENTANA_DIAS` (60) días después de la emisión, en domingo, o resolución no leída (se confirma visualmente que es la REC UNO). No bloquean, pero se muestran al revisor y quedan en el informe.
6. Tipos de documento: REC (rótulo localizado), REC_CARGO_NO_UBICADO (página vertical con encabezado cuyo rótulo no se leyó: se intenta por posición), REC_SIN_CARGO (foto horizontal, media página: se empareja con la foto siguiente de la carpeta), CARGO_SIN_ENCABEZADO, ANEXO (AVISO PREVIO y CEDULÓN, sin fecha para F) y DESCONOCIDO.

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

## Lector externo por API (opcional)

`LECTOR_API` en las celdas 3 y 4 añade un lector manuscrito más por API, en formato compatible con OpenAI: `gemini` (Google, por su punto de compatibilidad con OpenAI; modelo en `API_MODELO`, por defecto `gemini-3.8-flash`, porque Google retiró `gemini-2.5-flash` para claves nuevas) o `deepseek` (`deepseek-v4-flash-vision-exp`). Ante HTTP 503 o tiempo de espera agotado reintenta hasta tres veces con esperas de 5, 10 y 20 s y lo cuenta aparte; ante HTTP 404 de modelo avisa «modelo no disponible para esta clave» y desactiva solo ese lector. Cuenta como lector manuscrito independiente, así que con Qwen y la API coincidiendo la confianza ALTA es alcanzable. Condiciones:

- **Nivel gratuito de Gemini**: no requiere tarjeta, pero limita las peticiones por minuto y por día, y Google indica que puede usar lo enviado para mejorar sus productos. `API_PAUSA_S` (6 s) respeta el límite por minuto; si la cuota diaria se agota, el lector recibe HTTP 429 varias veces seguidas, se desactiva solo para el resto de la pasada y los demás lectores siguen. El resumen de la celda 3 indica llamadas, rechazos por cuota y si quedó desactivado. Con una sola vista por recuadro, un piloto de 200 fotos ronda las 200 a 250 llamadas.

- **Solo sale el recorte de la fecha**, sin nombre, expediente ni código. Aun así es un envío a un servicio externo (servidores de DeepSeek): la decisión corresponde a la institución, y conviene revisar la política de datos del proveedor.
- La clave se guarda en los **secretos de Colab** (`GEMINI_API_KEY` o `DEEPSEEK_API_KEY`, con acceso del cuaderno activado), nunca en el repositorio, en el cuaderno ni en un chat.
- Costo orientativo: cada imagen cuesta como máximo unos 384 tokens de entrada; el lote completo con dos lecturas por foto ronda los 3,6 millones de tokens, es decir, menos de un dólar a los precios publicados en 2026. Verifica la tarifa vigente.
- Un fallo de red o de la API se registra como `ERROR` en esa lectura y no detiene la pasada; se reintenta dos veces con espera. La celda 3 informa llamadas, errores y tokens consumidos.
- Su utilidad sobre actas manuscritas no está medida: el piloto es quien lo dirá.

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
