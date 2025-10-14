from flask import Flask, render_template, request, jsonify
from markupsafe import Markup
import pandas as pd
import openpyxl

app = Flask(__name__)

# === CARGA DEL EXCEL ===
# Cabecera en la fila 3 (index 2)
df = pd.read_excel("INVENTARIO.xlsx", header=2, sheet_name=0)

# Leer hipervínculos con openpyxl
wb = openpyxl.load_workbook("INVENTARIO.xlsx", data_only=True)
ws = wb.active

# Extraer hipervínculos reales de la columna "Link"
if "Link" in df.columns:
    col_idx = list(df.columns).index("Link") + 1
    links = []
    for row in ws.iter_rows(min_row=4, min_col=col_idx, max_col=col_idx):
        cell = row[0]
        if cell.hyperlink:
            links.append(cell.hyperlink.target)
        else:
            links.append(None)
    df["Link"] = links[:len(df)]

def buscar_codigo(codigo):
    """Busca una fila según el código en la segunda columna."""
    codigo = codigo.strip()
    row = df[df.iloc[:, 1].astype(str).str.strip() == codigo]
    if not row.empty:
        record = row.to_dict(orient="records")[0]
        if "Link" in record and pd.notna(record["Link"]):
            link = str(record["Link"]).strip()
            record["Enlace"] = Markup(f'<a href="{link}" target="_blank">{link}</a>')
            # Incrustar documento
            embed = ""
            if link.endswith("/"):
                embed = f'<iframe src="{link}" width="100%" height="600px"></iframe>'
            elif link.endswith(".pdf"):
                embed = f'<embed src="{link}" type="application/pdf" width="100%" height="600px">'
            elif link.endswith((".png", ".jpg", ".jpeg")):
                embed = f'<img src="{link}" alt="Imagen relacionada" style="max-width:100%;">'
            else:
                embed = f'<iframe src="{link}" width="100%" height="600px"></iframe>'
            record["Documento"] = Markup(embed)
        return record
    return None

@app.route("/", methods=["GET", "POST"])
@app.route("/<codigo>", methods=["GET", "POST"])
def index(codigo=None):
    """Permite búsqueda por formulario o acceso directo desde URL."""
    result = None
    if request.method == "POST":
        codigo = request.form["codigo"].strip()
    elif codigo:
        codigo = codigo.strip()
    if codigo:
        result = buscar_codigo(codigo) or "No se encontró el código."
    return render_template("index.html", result=result)
    
# @app.route("/", methods=["GET", "POST"])
# def index():
#     result = None
#     if request.method == "POST":
#         codigo = request.form["codigo"].strip()
#         result = buscar_codigo(codigo) or "No se encontró el código."
#     return render_template("index.html", result=result)

# @app.route("/<codigo>")
# def direct_lookup(codigo):
#     """Permite acceso directo con URL tipo /A-01-3ALMINS000037"""
#     result = buscar_codigo(codigo.strip()) or "No se encontró el código."
#     return render_template("index.html", result=result)

@app.route("/autocomplete")
def autocomplete():
    """Devuelve sugerencias de códigos en formato JSON."""
    query = request.args.get("q", "").strip().lower()
    codes = df.iloc[:, 1].dropna().astype(str)
    matches = [c for c in codes if query in c.lower()][:10]  # hasta 10 sugerencias
    return jsonify(matches)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
