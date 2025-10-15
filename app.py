from flask import Flask, render_template, request, redirect, url_for
from markupsafe import Markup
import pandas as pd
import openpyxl
import sqlite3
import os
from datetime import datetime

app = Flask(__name__)

# === CARGA DEL EXCEL ===
EXCEL_FILE = "INVENTARIO.xlsx"
DB_FILE = "estado.db"

df = pd.read_excel(EXCEL_FILE, header=2, sheet_name=0)

wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
ws = wb.active
if "Link" in df.columns:
    col_idx = list(df.columns).index("Link") + 1
    links = []
    for row in ws.iter_rows(min_row=4, min_col=col_idx, max_col=col_idx):
        cell = row[0]
        links.append(cell.hyperlink.target if cell.hyperlink else None)
    df["Link"] = links[:len(df)]


# === FUNCIÓN DE INICIALIZACIÓN DE BASE DE DATOS ===
def init_db():
    """Crea la base de datos si no existe y la tabla de estados."""
    if not os.path.exists(DB_FILE):
        print("🟢 Creando base de datos local:", DB_FILE)
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
    c.execute("SELECT estado, prestado_a FROM inventario_estado WHERE codigo = ?", (codigo,))
    row = c.fetchone()
    if not row:
        # si no existe, crearlo automáticamente como disponible
        c.execute("INSERT INTO inventario_estado (codigo, estado) VALUES (?, 'Disponible')", (codigo,))
        conn.commit()
        conn.close()
        return "Disponible", None
    conn.close()
    return row[0], row[1]


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
    """Agrega estado y documento incrustado."""
    if "Link" in record and pd.notna(record["Link"]):
        link = str(record["Link"]).strip()
        record["Enlace"] = Markup(f'<a href="{link}" target="_blank">{link}</a>')
        if link.endswith(".pdf"):
            record["Documento"] = Markup(f'<embed src="{link}" type="application/pdf" width="100%" height="600px">')
        elif link.endswith((".png", ".jpg", ".jpeg")):
            record["Documento"] = Markup(f'<img src="{link}" style="max-width:100%;">')
        else:
            record["Documento"] = Markup(f'<iframe src="{link}" width="100%" height="600px"></iframe>')
    estado, prestado_a = get_estado(codigo)
    record["Estado"] = estado
    record["Prestado_a"] = prestado_a if prestado_a else "-"
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
@app.route("/")
@app.route("/<codigo>")
def index(codigo=None):
    result = None
    codigos = sorted(df.iloc[:, 1].dropna().astype(str).unique())
    descripciones = sorted(df.iloc[:, 3].dropna().astype(str).unique())

    if codigo:
        result = buscar_codigo(codigo)
        if not result:
            result = {"no_agregado": True, "mensaje": f"El código «{codigo}» no ha sido registrado o agregado al inventario."}

    return render_template("index.html", result=result, codigos=codigos, descripciones=descripciones)


@app.route("/buscar")
def buscar():
    desc = request.args.get("desc", "").strip()
    codigos = buscar_codigo_por_descripcion(desc)
    if len(codigos) == 0:
        mensaje = f"La descripción «{desc}» no tiene un código agregado al inventario."
        return render_template("seleccion.html", mensaje=mensaje)
    elif len(codigos) == 1:
        return redirect(f"/{codigos[0]}")
    else:
        coincidencias = df[df.iloc[:, 1].astype(str).isin(codigos)][[df.columns[1], df.columns[3]]]
        items = coincidencias.to_dict(orient="records")
        return render_template("seleccion.html", desc=desc, items=items, col_codigo=df.columns[1], col_desc=df.columns[3])


@app.route("/prestar/<codigo>", methods=["GET", "POST"])
def prestar(codigo):
    if request.method == "POST":
        alumno = request.form.get("alumno").strip()
        actualizar_estado(codigo, "Prestado", alumno)
        return redirect(f"/{codigo}")
    return render_template("prestar.html", codigo=codigo)


@app.route("/devolver/<codigo>", methods=["GET", "POST"])
def devolver(codigo):
    if request.method == "POST":
        actualizar_estado(codigo, "Disponible")
        return redirect(f"/{codigo}")
    return render_template("devolver.html", codigo=codigo)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
