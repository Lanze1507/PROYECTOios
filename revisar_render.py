"""Use the packaged rasterizer with a native Word PDF backend on Windows.

The runtime has no bundled LibreOffice for Windows. Export is done by hidden
Word COM; its PDF is passed to render_docx.rasterize without using desktop LO.
"""
from pathlib import Path
import importlib.util
import os

ROOT = Path(__file__).parent
RENDERER = Path(r'C:\Users\joseb\.codex\plugins\cache\openai-primary-runtime\documents\26.909.12148\skills\documents\render_docx.py')
POPP = Path(r'C:\Users\joseb\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin')
if not (POPP/'pdftoppm.exe').exists():
    POPP = Path(r'C:\Users\joseb\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\bin')
os.environ['PATH'] = str(POPP) + os.pathsep + os.environ.get('PATH','')
spec = importlib.util.spec_from_file_location('render_docx', RENDERER)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
pdf = ROOT/'revision_word.pdf'
if not pdf.exists():
    raise FileNotFoundError(pdf)
module.convert_to_pdf = lambda *args, **kwargs: (str(pdf), 'Native Word COM export')
pages=module.rasterize(str(ROOT.parent/'Proyecto_Metodo_Simplex_Prototipo_1_0.docx'), str(ROOT/'revision'), 145, False, False)
print('\n'.join(pages))
