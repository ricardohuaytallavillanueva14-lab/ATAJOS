"""Reconstruye el notebook autocontenido sin necesitar fotografías ni Excel."""
from pathlib import Path
import ast
import json

ROOT = Path(__file__).resolve().parents[1]
def build():
    def read(name):
        return (ROOT / 'templates' / name).read_text(encoding='utf-8')
    init = read('proyecto_inicio.py') + (ROOT / 'src/ocr_sat.py').read_text(encoding='utf-8') + read('proyecto_final.py')
    sources = [read('preparacion.py'), init, read('procesamiento.py'), read('revision.py'), read('exportacion.py'), read('paquete.py')]
    cells = [dict(cell_type='markdown', metadata={}, source=read('guia.md').splitlines(True), id='guia')]
    for i, source in enumerate(sources, 1):
        ast.parse(source)
        cells.append(dict(cell_type='code', metadata={}, source=source.splitlines(True), id=f'paso{i}', outputs=[], execution_count=None))
    notebook = dict(nbformat=4, nbformat_minor=5, metadata={'kernelspec':{'display_name':'Python 3','language':'python','name':'python3'},'accelerator':'GPU'}, cells=cells)
    output = ROOT / 'notebooks/REC_2026_02_SEPTIEMBRE_Colab.ipynb'
    output.parent.mkdir(exist_ok=True)
    output.write_text(json.dumps(notebook, ensure_ascii=False, indent=2), encoding='utf-8')
    print(output)
    return output
if __name__ == '__main__':
    build()
