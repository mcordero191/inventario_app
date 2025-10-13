from flask import Flask, render_template, request
from markupsafe import Markup
import pandas as pd
import os

app = Flask(__name__)

# === CARGA DEL EXCEL ===
# Asume que la cabecera está en la tercera fila (índice 2)
df = pd.read_excel("INVENTARIO.xlsx", header=2, sheet_name=0)

@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    if request.method == "POST":
        codigo = request.form["codigo"].strip()
        # Buscar en la segunda columna (índice 2)
        row = df[df.iloc[:, 1].astype(str).str.strip() == codigo]
        if not row.empty:
            record = row.to_dict(orient="records")[0]
            
            # Si existe la columna "Link", incrustar su contenido
            if "Link" in record and pd.notna(record["Link"]):
                link = str(record["Link"]).strip()
                embed = ""
                if link.endswith(".pdf"):
                    embed = f'<embed src="{link}" type="application/pdf" width="100%" height="600px">'
                elif link.endswith((".png", ".jpg", ".jpeg")):
                    embed = f'<img src="{link}" alt="Imagen relacionada" style="max-width:100%;">'
                else:
                    embed = f'<iframe src="{link}" width="100%" height="600px"></iframe>'
                record["Documento"] = Markup(embed)
            result = record
        else:
            result = "No se encontró el código."
    return render_template("index.html", result=result)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
