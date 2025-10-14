from flask import Flask, render_template, request, redirect, url_for
from markupsafe import Markup
import pandas as pd
import openpyxl

app = Flask(__name__)

# === CARGA DEL EXCEL ===
df = pd.read_excel("INVENTARIO.xlsx", header=2, sheet_name=0)

# Leer hipervínculos reales de la columna "Link"
wb = openpyxl.load_workbook("INVENTARIO.xlsx", data_only=True)
ws = wb.active
if "Link" in df.columns:
    col_idx = list(df.columns).index("Link") + 1
    links = []
    for row in ws.iter_rows(min_row=4, min_col=col_idx, max_col=col_idx):
        cell = row[0]
        links.append(cell.hyperlink.target if cell.hyperlink else None)
    df["Link"] = links[:len(df)]


# === FUNCIONES ===
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
    """Busca por código (columna 2)."""
    row = df[df.iloc[:, 1].astype(str).str.strip().str.lower() == codigo.lower()]
    if not row.empty:
        return preparar_registro(row.to_dict(orient="records")[0])
    return None


def buscar_codigo_por_descripcion(desc):
    """Busca códigos que coincidan con la descripción (columna 4)."""
    desc = desc.lower().strip()
    matches = df[df.iloc[:, 3].astype(str).str.lower().str.contains(desc, na=False)]
    codigos = matches.iloc[:, 1].dropna().astype(str).unique().tolist()
    return codigos


def resumen_general():
    """Resumen de conteo de elementos."""
    total = len(df)
    sin_codigo = df[df.iloc[:, 1].isna()].shape[0]
    con_codigo = total - sin_codigo
    con_codigo_link = df[(~df.iloc[:, 1].isna()) & (df["Link"].notna())].shape[0]
    return dict(total=total, sin_codigo=sin_codigo, con_codigo=con_codigo, con_codigo_link=con_codigo_link)


# === RUTAS PRINCIPALES ===
@app.route("/", methods=["GET"])
@app.route("/<codigo>", methods=["GET"])
def index(codigo=None):
    """Vista principal: búsqueda por código."""
    result = None
    stats = resumen_general()
    codigos = sorted(df.iloc[:, 1].dropna().astype(str).unique())
    descripciones = sorted(df.iloc[:, 3].dropna().astype(str).unique())

    if codigo:
        result = buscar_codigo(codigo)
        if not result:
            result = {
                "no_agregado": True,
                "mensaje": f"El código «{codigo}» no ha sido registrado o agregado al inventario."
            }

    return render_template("index.html", result=result, stats=stats,
                           codigos=codigos, descripciones=descripciones)


@app.route("/buscar", methods=["GET"])
def buscar():
    """Procesa la búsqueda por descripción y redirige o muestra resultados múltiples."""
    desc = request.args.get("desc", "").strip()
    if not desc:
        return redirect(url_for("index"))

    codigos = buscar_codigo_por_descripcion(desc)

    # Caso: sin ningún código asociado
    if len(codigos) == 0:
        mensaje = f"La descripción «{desc}» no tiene un código agregado al inventario."
        return render_template("seleccion.html", mensaje=mensaje)

    # Caso: un solo código → redirigir directamente
    elif len(codigos) == 1:
        codigo = codigos[0].strip()
        if codigo == "" or codigo.lower() in ["nan", "none"]:
            mensaje = f"La descripción «{desc}» no tiene un código agregado al inventario."
            return render_template("seleccion.html", mensaje=mensaje)
        else:
            return redirect(f"/{codigo}")

    # Caso: múltiples coincidencias → mostrar tabla de código + descripción
    else:
        codigos_validos = [c for c in codigos if str(c).strip() not in ["", "nan", "none"]]
        if not codigos_validos:
            mensaje = f"La descripción «{desc}» no tiene un código agregado al inventario."
            return render_template("seleccion.html", mensaje=mensaje)

        coincidencias = df[df.iloc[:, 1].astype(str).isin(codigos_validos)][[df.columns[1], df.columns[3]]]
        items = coincidencias.to_dict(orient="records")

        return render_template("seleccion.html", desc=desc, items=items,
                               col_codigo=df.columns[1], col_desc=df.columns[3])


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
