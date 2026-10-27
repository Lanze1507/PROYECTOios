"""Ejecuta simplex.py sin modificarlo y conserva su salida textual real."""
import importlib.util
import io
from contextlib import redirect_stdout
from pathlib import Path

carpeta = Path(__file__).resolve().parent
fuente = Path(r"C:\Users\joseb\Downloads\simplex.py")
spec = importlib.util.spec_from_file_location("simplex_documentado", fuente)
modulo = importlib.util.module_from_spec(spec)
spec.loader.exec_module(modulo)

for indice, caso in enumerate(modulo.CASOS_PRUEBA):
    salida = io.StringIO()
    with redirect_stdout(salida):
        modulo.resolver_caso(**caso)
    completa = salida.getvalue()
    (carpeta / f"caso_{indice + 1}_salida_completa.txt").write_text(completa, encoding="utf-8")
    extracto = (
        f"{caso['nombre']}\n"
        "Fuente: C:\\Users\\joseb\\Downloads\\simplex.py\n"
        f"Llamada ejecutada: resolver_caso(**CASOS_PRUEBA[{indice}])\n\n"
        + completa[completa.index("TABLA FINAL"):completa.index("INTERPRETACION")]
        + completa[completa.index("RESULTADO: VALIDACION CORRECTA"):]
    )
    (carpeta / f"caso_{indice + 1}_tabla_final.txt").write_text(extracto, encoding="utf-8")
    print(extracto)
