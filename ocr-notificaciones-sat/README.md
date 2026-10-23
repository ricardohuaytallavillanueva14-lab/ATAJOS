# OCR Notificaciones SAT

Proyecto para procesar 3.996 fotografías desde Google Drive y completar fechas de notificación REC en una copia del Excel de expedientes.

[![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/ricardohuaytallavillanueva14-lab/ATAJOS/blob/main/ocr-notificaciones-sat/notebooks/REC_2026_02_SEPTIEMBRE_Colab.ipynb)

## Alcance del Excel

| Elemento | Regla |
| --- | --- |
| Hoja | `2026` |
| Columna E | Filtro por fecha de emisión **02/09/2026** |
| Columna F | Destino de la fecha de notificación, solo si está vacía |
| Cruce | Expediente y año únicos, código SIGETI de K coincidente, REC UNO |
| Aprobación | Cotejo visual antes de escribir |
| Salida | Copia nueva; el original se conserva |

El archivo revisado tiene 3.401 filas con la emisión indicada, de las cuales 1.950 tienen F vacía. El código comprueba el SHA-256 del Excel específico para evitar cargar otro archivo por accidente.

## Uso

1. Abre el botón **Abrir en Colab**. Si deseas conservar cambios, guarda una copia del cuaderno en tu Drive.
2. Utiliza un entorno alojado, preferentemente con GPU T4.
3. Ejecuta la celda 1: autoriza Drive, verifica el inventario e instala dependencias.
4. Ejecuta la celda 2: sube el Excel original si lo solicita. Se guarda para reutilizarlo. Espera el inventario completo.
5. Ejecuta la celda 3: procesa todas las imágenes únicas y guarda cada lectura.
6. Ejecuta la celda 4: coteja imágenes, corrige campos y confirma las lecturas legibles.
7. Ejecuta la celda 5: genera la copia del Excel y el informe CSV.

La celda 6 es opcional: crea un paquete de hasta 20 casos pendientes para una segunda revisión. No envía imágenes a ningún asistente automáticamente.

## Rutas de Drive

Fotos: `/content/drive/MyDrive/01_SAT_Huamanga/PROYECTO REC 1`

Resultados: `/content/drive/MyDrive/01_SAT_Huamanga/REC_2026_02SEP_RESULTADOS`

El repositorio contiene software y documentación. Las fotografías, el Excel, las lecturas y las aprobaciones se guardan en Drive. No agregues datos de contribuyentes al repositorio público.

## Reanudación

Si se pierde la conexión, reconecta y comprueba si existe una ejecución activa antes de iniciar otra. Tras perder el estado del entorno, ejecuta 1 y 2 y luego 3. Los registros guardados se recuperan; el inventario vuelve a verificar las imágenes. No ejecutes dos sesiones contra la misma salida.

Los errores anteriores no se reintentan por defecto. Cambia `REINTENTAR_ERRORES = True` en la celda 3 para reintentarlos. La celda 3 también llama a la exportación: produce un informe y puede generar una copia si ya había aprobaciones recuperadas.

## Código y mantenimiento

- `src/ocr_sat.py`: extracción, inventario, revisión, validación y escritura del XLSX.
- `templates/`: código de las seis celdas y guía incorporada.
- `scripts/build_notebook.py`: reconstruye el cuaderno autocontenido.
- `notebooks/`: cuaderno listo para Colab.
- `tests/test_excel.py`: pruebas con datos ficticios, sin descargar modelos.
- `docs/CONTEXTO_CLAUDE.md`: contexto de supervisión y recuperación.

Desde esta carpeta:

```bash
python -m pip install -r requirements-test.txt
python -m unittest discover -s tests -v
python scripts/build_notebook.py
```

Las dependencias completas están en `requirements.txt`. La celda 1 las instala en Colab. El cuaderno se ejecuta solo, sin clonar el repositorio.

## Validación y límites

Se verificó la escritura sobre una copia del Excel real: solo cambió la celda objetivo y su formato; las otras partes del XLSX se conservaron. Las pruebas reproducibles del repositorio usan un libro ficticio para comprobar el alcance, fórmulas, duplicados y celdas ocupadas.

El OCR completo de las 3.996 fotografías y su precisión aún no se han validado. PaddleOCR utiliza CPU y TrOCR usa GPU si está disponible. El recorte automático está adaptado al encabezado y cargo inferior; fotografías giradas, diseños distintos, AVISO PREVIO y CEDULON pueden requerir revisión manual. Las coincidencias entre motores no constituyen una tasa de exactitud medida. Los límites de sesión de Colab siguen aplicando.

No se incluye una licencia de redistribución: el titular puede añadir la que corresponda.
