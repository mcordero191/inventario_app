from flask import Flask, render_template, request, redirect, url_for, jsonify
from markupsafe import Markup
import pandas as pd
import openpyxl
import sqlite3
import os
from datetime import datetime

app = Flask(__name__)

EXCEL_FILE = "INVENTARIO.xlsx"
DB_FILE = "estado.db"

# === CARGA DEL EXCEL ===
df = pd.read_excel(EXCEL_FILE, header=2, sheet_name=0)

# === BASE DE DATOS LOCAL ===
def init_db():
    """Crea la base si no existe."""
    if not os.path.exists(DB_FILE):
        print("Creando base de datos local:", DB_FILE)
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS inventario_estado (
            codigo TEXT PRIMARY KEY,
            estado TEXT DEFAULT 'Disponible',
            prestado_a TEXT,
            fecha_prestamo TEXT
        )
    """)
    conn.commit()
    conn.close()

init_db()


# === FUNCIONES AUXILIARES ===
def get_estado(codigo):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT estado, prestado_a, fecha_prestamo FROM inventario_estado WHERE codigo = ?", (codigo,))
    row = c.fetchone()
    if not row:
        c.execute("INSERT INTO inventario_estado (codigo, estado) VALUES (?, 'Disponible')", (codigo,))
        conn.commit()
        conn.close()
        return "Disponible", None, None
    conn.close()
    return row[0], row[1], row[2]


def actualizar_estado(codigo, estado, prestado_a=None):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    if estado == "Prestado":
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        c.execute("""
            REPLACE INTO inventario_estado (codigo, estado, prestado_a, fecha_prestamo)
            VALUES (?, ?, ?, ?)
        """, (codigo, estado, prestado_a, fecha))
    else:
        c.execute("""
            UPDATE inventario_estado
            SET estado='Disponible', prestado_a=NULL, fecha_prestamo=NULL
            WHERE codigo=?
        """, (codigo,))
    conn.commit()
    conn.close()


def preparar_registro(record, codigo):
    """Agrega estado, enlace y documento incrustado."""
    record["Codigo"] = codigo  # campo uniforme para HTML
    estado, prestado_a, fecha_prestamo = get_estado(codigo)
    record["Estado"] = estado
    record["Prestado_a"] = prestado_a if prestado_a else "-"
    record["Fecha_de_prestamo"] = fecha_prestamo
    return record


def buscar_codigo(codigo):
    row = df[df.iloc[:, 1].astype(str).str.strip().str.lower() == codigo.lower()]
    if not row.empty:
        return preparar_registro(row.to_dict(orient="records")[0], codigo)
    return None


def buscar_codigo_por_descripcion(desc):
    desc = desc.lower().strip()
    matches = df[df.iloc[:, 3].astype(str).str.lower().str.contains(desc, na=False)]
    return matches.iloc[:, 1].dropna().astype(str).unique().tolist()


# === RUTAS ===
@app.route("/", methods=["GET"])
@app.route("/<codigo>", methods=["GET"])
def index(codigo=None):
    """Vista principal: búsqueda por código o descripción."""
    q_codigo = request.args.get("codigo", "").strip()
    if q_codigo:
        return redirect(f"/{q_codigo}")

    result = None
    codigos = sorted(df.iloc[:, 1].dropna().astype(str).unique())
    descripciones = sorted(df.iloc[:, 3].dropna().astype(str).unique())

    if codigo:
        result = buscar_codigo(codigo)
        if not result:
            result = {"no_agregado": True, "mensaje": f"El código «{codigo}» no ha sido registrado o agregado al inventario."}

    return render_template("index.html", result=result, codigos=codigos, descripciones=descripciones)


@app.get("/autocomplete")
def autocomplete():
    """Devuelve sugerencias para autocompletar."""
    q = (request.args.get("q") or "").strip().lower()
    kind = (request.args.get("kind") or "code").lower()

    if kind == "desc":
        col = df.iloc[:, 3].dropna().astype(str)
    else:
        col = df.iloc[:, 1].dropna().astype(str)

    if not q:
        suggestions = col.unique().tolist()[:10]
    else:
        suggestions = [v for v in col if q in v.lower()][:10]

    return jsonify(suggestions)

@app.route("/buscar", methods=["GET"])
def buscar():
    desc = request.args.get("desc", "").strip()
    if not desc:
        return redirect(url_for("index"))

    codigos = buscar_codigo_por_descripcion(desc)

    if len(codigos) == 0:
        mensaje = f"La descripción «{desc}» no tiene un código agregado al inventario."
        return render_template("seleccion.html", mensaje=mensaje)

    if len(codigos) == 1:
        return redirect(f"/{codigos[0]}")

    # Obtener código, descripción y estado actual
    coincidencias = df[df.iloc[:, 1].astype(str).isin(codigos)][[df.columns[1], df.columns[3]]]
    # coincidencias = coincidencias.rename(columns={df.columns[1]: "Codigo", df.columns[3]: "Descripcion"})

    items = []
    for _, row in coincidencias.iterrows():
        codigo = str(row["Codigo"])
        estado, prestado_a, fecha_de_prestamo = get_estado(codigo)
        items.append({
            "Código": codigo,
            "Descripción": row["Descripcion"],
            "Estado": estado,
            "Prestado_a": prestado_a if prestado_a else "-",
            "Fecha de préstamo": fecha_de_prestamo,
        })

    return render_template("seleccion.html", desc=desc, items=items)


@app.route("/prestar/<codigo>", methods=["GET", "POST"])
def prestar(codigo):
    if request.method == "POST":
        alumno = request.form.get("alumno").strip()
        actualizar_estado(codigo, "Prestado", alumno)
        return redirect(f"/{codigo}")
    row = df[df.iloc[:, 1].astype(str).str.strip().str.lower() == codigo.lower()]
    return render_template("prestar.html", codigo=codigo, desc=row["Descripcion"].dropna().astype(str))


@app.route("/devolver/<codigo>", methods=["GET", "POST"])
def devolver(codigo):
    if request.method == "POST":
        actualizar_estado(codigo, "Disponible")
        return redirect(f"/{codigo}")
    row = df[df.iloc[:, 1].astype(str).str.strip().str.lower() == codigo.lower()]
    return render_template("devolver.html", codigo=codigo, desc=row["Descripcion"].dropna().astype(str))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
