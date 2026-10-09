"""Empaqueta el tablero, los datos y los PDF en entregables/Tablero_BSC_Pedidos.zip (carpeta raíz: Trabajos)."""
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
LEEME = """TABLERO BALANCED SCORECARD - PEDIDOS OLIST

1. Descomprime este ZIP directamente en C:\\  (queda C:\\Trabajos\\datos y C:\\Trabajos\\tablero).
2. Abre C:\\Trabajos\\tablero\\Tablero_Pedidos_BSC.pbip con Power BI Desktop.
3. Si lo descomprimiste en otra carpeta: Inicio > Transformar datos > Editar parametros >
   RutaDatos = carpeta donde esta olist_orders_dataset.csv (terminada en \\) > Aplicar cambios.
4. Inicio > Actualizar.
5. Archivo > Guardar como > Archivo de Power BI (.pbix)  -> ese es el archivo para entregar.

Valores esperados y explicacion del ETL: Guia_ETL_y_Tablero_BSC.pdf
"""

destino = RAIZ / "entregables" / "Tablero_BSC_Pedidos.zip"
with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    z.writestr("Trabajos/LEEME.txt", LEEME.replace("\n", "\r\n"))
    z.write(RAIZ / "datos/originales/olist_orders_dataset.csv", "Trabajos/datos/originales/olist_orders_dataset.csv")
    for f in sorted((RAIZ / "tablero").rglob("*")):
        if f.is_file() and f.name != ".gitignore":
            z.write(f, "Trabajos/" + f.relative_to(RAIZ).as_posix())
    for pdf in ["Resumen_Metodologias_BI.pdf", "Guia_ETL_y_Tablero_BSC.pdf"]:
        z.write(RAIZ / "entregables" / pdf, "Trabajos/" + pdf)
print(destino, round(destino.stat().st_size / 1e6, 1), "MB")
