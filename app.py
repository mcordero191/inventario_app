from flask import Flask, render_template, request, jsonify
from markupsafe import Markup
import pandas as pd
import openpyxl

app = Flask(__name__)

# === CARGA DEL EXCEL ===
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

# === Funciones de ayuda ===

def buscar_codigo(codigo):
    """Busca una fila según el código (segunda columna)."""
    row = df[df.iloc[:, 1].astype(str).str.strip().str.lower() == codigo.lower()]
    if not row.empty:
        return preparar_registro(row.to_dict(orient="records")[0])
    return None

def buscar_descripcion(texto):
    """Busca por descripción (tercera columna o que contenga el texto)."""
    texto = texto.lower().strip()
    matches = df[df.apply(lambda x: any(texto in str(v).lower() for v in x.values), axis=1)]
    if not matches.empty:
        return [preparar_registro(r) for r in matches.to_dict(orient="records")]
    return []

def preparar_registro(record):
    """Formatea un registro y agrega documento incrustado."""
    if "Link" in record and pd.notna(record["Link"]):
        link = str(record["Link"]).strip()
        record["Enlace"] = Markup(f'<a href="{link}" target="_blank">{link}</a>')
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

def resumen_general():
    """Cuenta los elementos con/sin código y con link."""
    total = len(df)
    sin_codigo = df[df.iloc[:, 1].isna()].shape[0]
    con_codigo = total - sin_codigo
    con_codigo_link = df[(~df.iloc[:, 1].isna()) & (df["Link"].notna())].shape[0]
    return {
        "total": total,
        "sin_codigo": sin_codigo,
        "con_codigo": con_codigo,
        "con_codigo_link": con_codigo_link
    }

# === Rutas ===

@app.route("/", methods=["GET"])
@app.route("/<codigo>", methods=["GET"])
def index(codigo=None):
    result = None
    multiples = None
    stats = resumen_general()

    desc = request.args.get("desc", "").strip()
    if codigo:
        codigo = codigo.strip()
        result = buscar_codigo(codigo) or "No se encontró el código."
    elif desc:
        multiples = buscar_descripcion(desc)
        if not multiples:
            multiples = "No se encontraron coincidencias."
    return render_template("index.html", result=result, multiples=multiples, stats=stats)

@app.route("/autocomplete")
def autocomplete():
    query = request.args.get("q", "").strip().lower()
    codes = df.iloc[:, 1].dropna().astype(str)
    matches = [c for c in codes if query in c.lower()][:10]
    return jsonify(matches)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
