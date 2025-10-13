from flask import Flask, render_template, request
from markupsafe import Markup
import pandas as pd
import openpyxl
import os

app = Flask(__name__)

# === CARGA DEL EXCEL ===
# Leer los valores normalmente con pandas (cabecera en fila 2)
df = pd.read_excel("INVENTARIO.xlsx", header=2, sheet_name=0)

# Leer el archivo también con openpyxl para extraer hipervínculos
wb = openpyxl.load_workbook("INVENTARIO.xlsx", data_only=True)
ws = wb.active

# Si existe columna "Link", reemplazamos su contenido por los hyperlinks reales
if "Link" in df.columns:
    col_idx = list(df.columns).index("Link") + 1  # índice de columna en Excel (1-based)
    links = []
    for row in ws.iter_rows(min_row=3, min_col=col_idx, max_col=col_idx):  # fila 3 = primera fila de datos
        cell = row[0]
        if cell.hyperlink:
            links.append(cell.hyperlink.target)
        else:
            links.append(None)
    df["Link"] = links[:len(df)]

@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    if request.method == "POST":
        codigo = request.form["codigo"].strip()
        row = df[df.iloc[:, 1].astype(str).str.strip() == codigo]
        if not row.empty:
            record = row.to_dict(orient="records")[0]

            # Procesar si hay link
            if "Link" in record and pd.notna(record["Link"]):
                link = str(record["Link"]).strip()
                embed = ""

                # Si el link apunta a una carpeta (termina con /)
                if link.endswith("/"):
                    embed = f'<iframe src="{link}" width="100%" height="600px"></iframe>'

                # Si apunta a PDF
                elif link.endswith(".pdf"):
                    embed = f'<embed src="{link}" type="application/pdf" width="100%" height="600px">'

                # Si apunta a imagen
                elif link.endswith((".png", ".jpg", ".jpeg")):
                    embed = f'<img src="{link}" alt="Imagen relacionada" style="max-width:100%;">'

                # Si es otro tipo de recurso (carpeta de Google Drive, etc.)
                else:
                    embed = f'<iframe src="{link}" width="100%" height="600px"></iframe>'

                record["Documento"] = Markup(embed)

            result = record
        else:
            result = "No se encontró el código."
    return render_template("index.html", result=result)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
