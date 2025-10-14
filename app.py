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

# Extraer hipervínculos reales
if "Link" in df.columns:
    col_idx = list(df.columns).index("Link") + 1
    links = []
    for row in ws.iter_rows(min_row=4, min_col=col_idx, max_col=col_idx):
        cell = row[0]
        links.append(cell.hyperlink.target if cell.hyperlink else None)
    df["Link"] = links[:len(df)]

def preparar_registro(record):
    """Agrega documento incrustado y enlace si existe."""
    if "Link" in record and pd.notna(record["Link"]):
        link = str(record["Link"]).strip()
        record["Enlace"] = Markup(f'<a href="{link}" target="_blank">{link}</a>')
        if link.endswith("/"):
            record["Documento"] = Markup(f'<iframe src="{link}" width="100%" height="600px"></iframe>')
        elif link.endswith(".pdf"):
            record["Documento"] = Markup(f'<embed src="{link}" type="application/pdf" width="100%" height="600px">')
        elif link.endswith((".png", ".jpg", ".jpeg")):
            record["Documento"] = Markup(f'<img src="{link}" style="max-width:100%;">')
        else:
            record["Documento"] = Markup(f'<iframe src="{link}" width="100%" height="600px"></iframe>')
    return record

def buscar_codigo(codigo):
    row = df[df.iloc[:, 1].astype(str).str.strip().str.lower() == codigo.lower()]
    if not row.empty:
        return preparar_registro(row.to_dict(orient="records")[0])
    return None

def buscar_descripcion(desc):
    matches = df[df.apply(lambda x: any(desc.lower() in str(v).lower() for v in x.values), axis=1)]
    return [preparar_registro(r) for r in matches.to_dict(orient="records")]

def resumen_general():
    total = len(df)
    sin_codigo = df[df.iloc[:, 1].isna()].shape[0]
    con_codigo = total - sin_codigo
    con_codigo_link = df[(~df.iloc[:, 1].isna()) & (df["Link"].notna())].shape[0]
    return dict(total=total, sin_codigo=sin_codigo, con_codigo=con_codigo, con_codigo_link=con_codigo_link)

@app.route("/", methods=["GET"])
@app.route("/<codigo>", methods=["GET"])
def index(codigo=None):
    result = None
    multiples = None
    stats = resumen_general()

    desc = request.args.get("desc", "").strip()
    codigos = sorted(df.iloc[:, 1].dropna().astype(str).unique())
    descripciones = sorted(df.iloc[:, 2].dropna().astype(str).unique())  # tercera columna = descripción

    if codigo:
        result = buscar_codigo(codigo) or "No se encontró el código."
    elif desc:
        multiples = buscar_descripcion(desc)
        if not multiples:
            multiples = "No se encontraron coincidencias."

    return render_template("index.html", result=result, multiples=multiples, stats=stats,
                           codigos=codigos, descripciones=descripciones)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
