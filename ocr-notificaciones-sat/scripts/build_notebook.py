"""Reconstruye el cuaderno autocontenido de Colab sin necesitar fotografías ni Excel.

Cada celda de código nace de templates/; la celda 2 incrusta src/ocr_sat.py para
que el cuaderno funcione sin clonar el repositorio. Ejecuta este script tras
cambiar cualquiera de esos archivos y confirma el .ipynb resultante.
"""
from pathlib import Path
import ast
import json

ROOT = Path(__file__).resolve().parents[1]
NOTEBOOK = ROOT / 'notebooks/REC_2026_02_SEPTIEMBRE_Colab.ipynb'
CELDAS = ['preparacion.py', None, 'prueba.py', 'procesamiento.py', 'revision.py', 'exportacion.py', 'paquete.py']


def render():
    def read(name):
        return (ROOT / 'templates' / name).read_text(encoding='utf-8')
    init = read('proyecto_inicio.py') + (ROOT / 'src/ocr_sat.py').read_text(encoding='utf-8') + read('proyecto_final.py')
    sources = [init if name is None else read(name) for name in CELDAS]
    cells = [dict(cell_type='markdown', metadata={}, source=read('guia.md').splitlines(True), id='guia')]
    for i, source in enumerate(sources, 1):
        ast.parse(source)
        cells.append(dict(cell_type='code', metadata={}, source=source.splitlines(True), id=f'paso{i}', outputs=[], execution_count=None))
    metadata = {
        'kernelspec': {'display_name': 'Python 3', 'language': 'python', 'name': 'python3'},
        'language_info': {'name': 'python'},
        'accelerator': 'GPU',
        'colab': {'provenance': [], 'gpuType': 'T4', 'toc_visible': True},
    }
    return dict(nbformat=4, nbformat_minor=5, metadata=metadata, cells=cells)


def build(output=NOTEBOOK):
    output.parent.mkdir(exist_ok=True)
    output.write_text(json.dumps(render(), ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(output)
    return output


if __name__ == '__main__':
    build()
