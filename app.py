from flask import Flask, render_template, request, redirect, url_for, jsonify, send_from_directory, abort
from markupsafe import Markup
import pandas as pd
import sqlite3
import os
import unicodedata
from datetime import datetime

app = Flask(__name__)

# === Paths & files ===
EXCEL_FILE = "INVENTARIO2025.xlsx"     # <-- your new Excel file
DB_FILE = "estado.db"
FOTOS_DIR = "Fotos"                    # folder at project root

# === Helpers ===
def strip_accents(s: str) -> str:
    if s is None:
        return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn")

def norm_col(s: str) -> str:
    s = strip_accents(s).lower().strip()
    out = []
    prev_us = False
    for ch in s:
        if ch.isalnum():
            out.append(ch)
            prev_us = False
        else:
            if not prev_us:
                out.append("_")
                prev_us = True
    return "".join(out).strip("_")

def choose_col(df, candidates):
    norm_map = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        k = norm_col(cand)
        if k in norm_map:
            return norm_map[k]
    return None

def build_foto_path(codigo: str):
    if not codigo or len(codigo) < 6:
        return None
    stand = codigo[0]
    xdd = codigo[:4]
    xddi = codigo[:6]
    base = os.path.join(FOTOS_DIR, f"STAND {stand}", xdd, xddi)
    for ext in (".jpg", ".jpeg", ".png", ".webp"):
        full = os.path.join(base, f"{codigo}{ext}")
        if os.path.isfile(full):
            rel = os.path.relpath(full, FOTOS_DIR)
            return rel.replace("\\", "/")
    return None

# === Load Excel (header at row 3) ===
df_raw = pd.read_excel(EXCEL_FILE, header=2, dtype=str)
df_raw.columns = [str(c).strip().lower() for c in df_raw.columns]
df = df_raw.copy()

COL_CODIGO   = choose_col(df, ["codigo", "código"])
COL_DESC     = choose_col(df, ["descripcion", "descripción"])
COL_SERIE    = choose_col(df, ["serie", "nro_serie", "numero de serie", "número de serie"])
COL_ACTFIJO  = choose_col(df, ["act fijo", "activo fijo", "act_fijo"])

# === SQLite for variable fields ===
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS inventario_estado (
            codigo TEXT PRIMARY KEY,
            estado TEXT DEFAULT 'Disponible',
            prestado_a TEXT,
            fecha_prestamo TEXT,
            num_prestamos INTEGER DEFAULT 0
        )
    """)
    conn.commit()
    conn.close()
init_db()

def get_estado(codigo):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT estado, prestado_a, fecha_prestamo, num_prestamos FROM inventario_estado WHERE codigo=?", (codigo,))
    row = c.fetchone()
    conn.close()
    if not row:
        return ("Disponible", None, None, 0)
    return row

def actualizar_estado(codigo, estado, prestado_a=None):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if estado == "Prestado":
        c.execute("""INSERT INTO inventario_estado (codigo, estado, prestado_a, fecha_prestamo, num_prestamos)
                     VALUES (?, 'Prestado', ?, ?, 1)
                     ON CONFLICT(codigo)
                     DO UPDATE SET estado='Prestado', prestado_a=excluded.prestado_a,
                     fecha_prestamo=excluded.fecha_prestamo,
                     num_prestamos=inventario_estado.num_prestamos+1""",
                     (codigo, prestado_a, now))
    else:
        c.execute("""INSERT INTO inventario_estado (codigo, estado)
                     VALUES (?, 'Disponible')
                     ON CONFLICT(codigo)
                     DO UPDATE SET estado='Disponible', prestado_a=NULL, fecha_prestamo=NULL""",
                     (codigo,))
    conn.commit()
    conn.close()

def preparar_registro(row_dict):
    codigo = str(row_dict.get(COL_CODIGO, "")).strip()
    fixed = {k: (v if v not in [None, "", "nan", "NaN"] else "-") for k, v in row_dict.items()}
    estado, prestado_a, fecha_prestamo, num_prestamos = get_estado(codigo)
    variable = {
        "Estado": estado,
        "Prestado_a": prestado_a or "-",
        "Fecha_de_prestamo": fecha_prestamo or "-",
        "Numero_prestamos": num_prestamos or 0
    }
    foto_rel = build_foto_path(codigo)
    return {"codigo": codigo, "fixed": fixed, "variable": variable, "foto_rel": foto_rel}

def buscar_por_col(valor, col_name):
    v = (valor or "").strip().lower()
    series = df[col_name].astype(str).str.strip().str.lower()
    hits = df[series == v]
    if not hits.empty:
        return preparar_registro(hits.iloc[0].to_dict())
    return None

def buscar_coincidencias(valor, col_name):
    v = (valor or "").strip().lower()
    series = df[col_name].astype(str).str.strip().str.lower()
    mask = series.str.contains(v, na=False)
    return df.loc[mask].copy()

# === Static fotos ===
@app.route("/fotos/<path:path>")
def fotos_static(path):
    safe = os.path.normpath(path).replace("\\", "/")
    return send_from_directory(FOTOS_DIR, safe)

# === Routes ===
@app.route("/", methods=["GET"])
@app.route("/<codigo>", methods=["GET"])
def index(codigo=None):
    q_codigo = request.args.get("codigo", "").strip()
    if q_codigo:
        return redirect(f"/{q_codigo}")

    result = None
    if codigo:
        result = buscar_por_col(codigo, COL_CODIGO)
        if not result:
            result = {"no_agregado": True, "mensaje": f"Código «{codigo}» no encontrado."}

    return render_template("index.html", result=result)

@app.route("/autocomplete")
def autocomplete():
    q = (request.args.get("q") or "").strip().lower()
    kind = request.args.get("kind", "code")
    if kind == "desc":
        col = df[COL_DESC]
    elif kind == "serie" and COL_SERIE:
        col = df[COL_SERIE]
    elif kind == "act" and COL_ACTFIJO:
        col = df[COL_ACTFIJO]
    else:
        col = df[COL_CODIGO]

    col = col.dropna().astype(str)
    suggestions = [v for v in col if q in v.lower()][:10]
    return jsonify(suggestions)

@app.route("/buscar", methods=["GET"])
def buscar():
    desc = request.args.get("desc", "").strip()
    serie = request.args.get("serie", "").strip()
    act = request.args.get("act", "").strip()

    if desc:
        df_hits = buscar_coincidencias(desc, COL_DESC)
        titulo = f"Descripción: {desc}"
    elif serie and COL_SERIE:
        df_hits = buscar_coincidencias(serie, COL_SERIE)
        titulo = f"Serie: {serie}"
    elif act and COL_ACTFIJO:
        df_hits = buscar_coincidencias(act, COL_ACTFIJO)
        titulo = f"Act Fijo: {act}"
    else:
        return redirect("/")

    if df_hits.empty:
        return render_template("seleccion.html", mensaje="No se encontraron coincidencias.")

    listing = []
    for _, row in df_hits.iterrows():
        code = str(row.get(COL_CODIGO, "")).strip()
        estado, _, _, nump = get_estado(code)
        listing.append({
            "codigo": code,
            "descripcion": str(row.get(COL_DESC, "") or ""),
            "serie": str(row.get(COL_SERIE, "") or "-") if COL_SERIE else "-",
            "actfijo": str(row.get(COL_ACTFIJO, "") or "-") if COL_ACTFIJO else "-",
            "estado": estado,
            "laboratorio": row.get("laboratorio"),
            "responsable": row.get("responsable"),
            "numero_prestamos": nump or 0
        })
    return render_template("seleccion.html", mensaje=None, items=listing, titulo=titulo)

@app.route("/prestar/<codigo>", methods=["GET", "POST"])
def prestar(codigo):
    if request.method == "POST":
        alumno = (request.form.get("alumno") or "").strip()
        actualizar_estado(codigo, "Prestado", alumno)
        return redirect(f"/{codigo}")
    desc = "-"
    row = df[df[COL_CODIGO].astype(str).str.lower() == codigo.lower()]
    if not row.empty:
        desc = row.iloc[0][COL_DESC]
    return render_template("prestar.html", codigo=codigo, desc=desc)

@app.route("/devolver/<codigo>", methods=["GET", "POST"])
def devolver(codigo):
    if request.method == "POST":
        actualizar_estado(codigo, "Disponible")
        return redirect(f"/{codigo}")
    desc = "-"
    row = df[df[COL_CODIGO].astype(str).str.lower() == codigo.lower()]
    if not row.empty:
        desc = row.iloc[0][COL_DESC]
    return render_template("devolver.html", codigo=codigo, desc=desc)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)