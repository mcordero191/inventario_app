from flask import Flask, render_template, request, redirect, url_for, jsonify, send_from_directory, abort, session
from markupsafe import Markup
import pandas as pd
import sqlite3, os, unicodedata
from datetime import datetime
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")  # CHANGE in prod

# === Paths & files ===
EXCEL_FILE = "static/INVENTARIO2025.xlsx"
DB_FILE = "estado.db"
FOTOS_DIR = "Fotos"

# === Roles helpers ===
def is_admin(): return session.get("role") == "admin"
def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not is_admin():
            return redirect(url_for("login", next=request.path))
        return fn(*args, **kwargs)
    return wrapper

@app.context_processor
def inject_user():
    return {"is_admin": is_admin(), "current_user": session.get("user")}

# === Helpers for Excel headers ===
def strip_accents(s): 
    if s is None: return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn")
def norm_col(s):
    s = strip_accents(str(s)).lower().strip()
    out, prev_us = [], False
    for ch in s:
        if ch.isalnum(): out.append(ch); prev_us=False
        else:
            if not prev_us: out.append("_"); prev_us=True
    return "".join(out).strip("_")
def choose_col(df, candidates):
    norm_map = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        k = norm_col(cand)
        if k in norm_map: return norm_map[k]
    return None

def build_foto_path(codigo, ubicacion):
    
    if not ubicacion or len(ubicacion) < 6:
        return None
    
    stand = ubicacion[0];
    xdd = ubicacion[:4];
    xddi = ubicacion[:]
    
    base = os.path.join(FOTOS_DIR, f"STAND {stand}", xdd, xddi)
    
    for ext in (".jpg",".jpeg",".png",".webp"):
        full = os.path.join(base, f"{codigo}{ext}")
        if os.path.isfile(full):
            rel = os.path.relpath(full, FOTOS_DIR)
            return rel.replace("\\","/")
        
    return None

# === Load Excel (header row 3) & lowercase headers ===
df_raw = pd.read_excel(EXCEL_FILE, header=2, dtype=str)
df_raw.columns = [str(c).strip().lower() for c in df_raw.columns]
df = df_raw.copy()

COL_CODIGO  = choose_col(df, ["codigo", "código"])
COL_UBI  = choose_col(df, ["codigo de ubicacion", "código de ubicación"])
COL_DESC    = choose_col(df, ["descripcion", "descripción"])
COL_SERIE   = choose_col(df, ["serie", "nro_serie", "nro de serie"])
COL_ACTFIJO = choose_col(df, ["act fijo", "activo fijo", "act_fijo", "codigo act fijo"])
if not COL_CODIGO: raise RuntimeError("Falta columna 'codigo'.")
if not COL_DESC:   raise RuntimeError("Falta columna 'descripcion'.")

# === DB init: users, states, audit ===
def db():
    return sqlite3.connect(DB_FILE)

def init_db():
    conn = db(); c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('admin','visitor'))
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS inventario_estado (
        codigo TEXT PRIMARY KEY,
        estado TEXT DEFAULT 'Disponible',
        prestado_a TEXT,
        fecha_prestamo TEXT,
        num_prestamos INTEGER DEFAULT 0,
        prestado_por TEXT,    -- username who lent last time
        devuelto_a TEXT     -- username who returned last time
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT,
        action TEXT,          -- 'prestar' | 'devolver'
        actor TEXT,           -- username performing action
        target TEXT,          -- person receiving item (for prestar) or '-' (for devolver)
        at TEXT               -- timestamp
    )""")
    conn.commit()
    # seed admin if missing
    c.execute("SELECT 1 FROM users WHERE username='admin'")
    if not c.fetchone():
        admin_pw = os.environ.get("ADMIN_PASSWORD", "changeme")
        c.execute("INSERT INTO users(username,password_hash,role) VALUES(?,?,?)",
                  ("admin", generate_password_hash(admin_pw), "admin"))
        conn.commit()
    conn.close()
init_db()

# === State helpers ===
def get_estado(codigo):
    conn = db(); c = conn.cursor()
    c.execute("""SELECT estado, prestado_a, fecha_prestamo, num_prestamos, prestado_por, devuelto_a
                 FROM inventario_estado WHERE codigo=?""", (codigo,))
    row = c.fetchone(); conn.close()
    return row or ("Disponible", None, None, 0, None, None)

def actualizar_estado(codigo, estado, actor_username, prestado_a=None):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = db(); c = conn.cursor()
    if estado == "Prestado":
        c.execute("""INSERT INTO inventario_estado (codigo, estado, prestado_a, fecha_prestamo, num_prestamos, prestado_por, devuelto_a)
                     VALUES (?, 'Prestado', ?, ?, 1, ?, NULL)
                     ON CONFLICT(codigo) DO UPDATE SET
                        estado='Prestado',
                        prestado_a=excluded.prestado_a,
                        fecha_prestamo=excluded.fecha_prestamo,
                        num_prestamos=inventario_estado.num_prestamos+1,
                        prestado_por=excluded.prestado_por,
                        devuelto_a=NULL
        """, (codigo, prestado_a, now, actor_username))
        c.execute("""INSERT INTO audit_log(codigo, action, actor, target, at)
                     VALUES(?, 'prestar', ?, ?, ?)""", (codigo, actor_username, prestado_a or "-", now))
    else:
        c.execute("""INSERT INTO inventario_estado (codigo, estado, prestado_a, fecha_prestamo, prestado_por, devuelto_a)
                     VALUES (?, 'Disponible', NULL, NULL, NULL, ?)
                     ON CONFLICT(codigo) DO UPDATE SET
                        estado='Disponible', prestado_a=NULL, fecha_prestamo=NULL, devuelto_a=excluded.devuelto_a
        """, (codigo, actor_username))
        c.execute("""INSERT INTO audit_log(codigo, action, actor, target, at)
                     VALUES(?, 'devolver', ?, '-', ?)""", (codigo, actor_username, now))
    conn.commit(); conn.close()

def preparar_registro(row_dict):
    codigo = str(row_dict.get(COL_CODIGO, "")).strip()
    ubicacion = str(row_dict.get(COL_UBI, "")).strip()
    fixed = {k: (v if v not in [None,"","nan","NaN"] else "-") for k, v in row_dict.items()}
    estado, prestado_a, fecha_prestamo, num_prestamos, prestado_por, devuelto_a = get_estado(codigo)
    variable = {
        "Estado": estado,
        "Prestado_a": prestado_a or "-",
        "Fecha_de_prestamo": fecha_prestamo or "-",
        "Numero_prestamos": num_prestamos or 0,
        "Prestado_por": prestado_por or "-",
        "Devuelto_a": devuelto_a or "-"
    }
    foto_rel = build_foto_path(codigo, ubicacion)
    return {"codigo": codigo, "fixed": fixed, "variable": variable, "foto_rel": foto_rel}

def buscar_por_col(valor, col_name):
    v = (valor or "").strip().lower()
    series = df[col_name].astype(str).str.strip().str.lower()
    hits = df[series == v]
    return preparar_registro(hits.iloc[0].to_dict()) if not hits.empty else None

def buscar_coincidencias(valor, col_name):
    v = (valor or "").strip().lower()
    series = df[col_name].astype(str).str.strip().str.lower()
    return df.loc[series.str.contains(v, na=False)].copy()

# === Static fotos ===
@app.route("/fotos/<path:path>")
def fotos_static(path):
    safe = os.path.normpath(path).replace("\\","/")
    return send_from_directory(FOTOS_DIR, safe)

# === Auth ===
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        conn = db(); c = conn.cursor()
        c.execute("SELECT password_hash, role FROM users WHERE username=?", (username,))
        row = c.fetchone(); conn.close()
        if row and check_password_hash(row[0], password):
            session["user"] = username; session["role"] = row[1]
            nxt = request.args.get("next") or url_for("index")
            return redirect(nxt)
        return render_template("login.html", error="Credenciales inválidas.", next=request.args.get("next"))
    return render_template("login.html", next=request.args.get("next"))

@app.route("/logout")
def logout():
    session.clear(); return redirect(url_for("index"))

# === User admin (Admins only) ===
@app.route("/users", methods=["GET","POST"])
@admin_required
def users():
    conn = db(); c = conn.cursor()
    if request.method == "POST":
        # create user
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        role = request.form.get("role") or "visitor"
        if username and password and role in ("admin","visitor"):
            try:
                c.execute("INSERT INTO users(username,password_hash,role) VALUES(?,?,?)",
                          (username, generate_password_hash(password), role))
                conn.commit()
            except sqlite3.IntegrityError:
                return render_template("users.html", error="El usuario ya existe.", users=list_users())
        else:
            return render_template("users.html", error="Complete todos los campos.", users=list_users())
    u = list_users(conn, c); conn.close()
    return render_template("users.html", users=u)

def list_users(conn=None, c=None):
    owns = False
    if conn is None:
        conn = db(); c = conn.cursor(); owns = True
    c.execute("SELECT username, role FROM users ORDER BY username")
    rows = [{"username": r[0], "role": r[1]} for r in c.fetchall()]
    if owns: conn.close()
    return rows

@app.route("/users/delete/<username>", methods=["POST"])
@admin_required
def delete_user(username):
    if username == "admin":  # keep a backdoor admin
        return redirect(url_for("users"))
    if username == session.get("user"):  # don't delete yourself
        return redirect(url_for("users"))
    conn = db(); c = conn.cursor()
    c.execute("DELETE FROM users WHERE username=?", (username,))
    conn.commit(); conn.close()
    return redirect(url_for("users"))

# === Main routes ===
@app.route("/", methods=["GET"])
@app.route("/<codigo>", methods=["GET"])
def index(codigo=None):
    q_codigo = request.args.get("codigo", "").strip()
    if q_codigo: return redirect(f"/{q_codigo}")
    result = None
    if codigo:
        result = buscar_por_col(codigo, COL_CODIGO)
        if not result: result = {"no_agregado": True, "mensaje": f"Código «{codigo}» no encontrado."}
    return render_template("index.html", result=result)

@app.route("/autocomplete")
def autocomplete():
    q = (request.args.get("q") or "").strip().lower()
    kind = request.args.get("kind","code")
    if kind == "desc": col = df[COL_DESC]
    elif kind == "serie" and COL_SERIE: col = df[COL_SERIE]
    elif kind == "act" and COL_ACTFIJO: col = df[COL_ACTFIJO]
    else: col = df[COL_CODIGO]
    col = col.dropna().astype(str)
    suggestions = [v for v in col if q in v.lower()][:10]
    return jsonify(suggestions)

@app.route("/buscar", methods=["GET"])
def buscar():
    desc = request.args.get("desc","").strip()
    serie = request.args.get("serie","").strip()
    act = request.args.get("act","").strip()
    if desc:
        df_hits = buscar_coincidencias(desc, COL_DESC); titulo = f"Descripción: {desc}"
    elif serie and COL_SERIE:
        df_hits = buscar_coincidencias(serie, COL_SERIE); titulo = f"Serie: {serie}"
    elif act and COL_ACTFIJO:
        df_hits = buscar_coincidencias(act, COL_ACTFIJO); titulo = f"Act Fijo: {act}"
    else:
        return redirect("/")
    if df_hits.empty: return render_template("seleccion.html", mensaje="No se encontraron coincidencias.")
    listing = []
    for _, row in df_hits.iterrows():
        code = str(row.get(COL_CODIGO,"")).strip()
        estado, _, _, nump, _, _ = get_estado(code)
        listing.append({
            "codigo": code,
            "descripcion": str(row.get(COL_DESC,"") or ""),
            "serie": str(row.get(COL_SERIE,"") or "-") if COL_SERIE else "-",
            "actfijo": str(row.get(COL_ACTFIJO,"") or "-") if COL_ACTFIJO else "-",
            "estado": estado, 
            "numero_prestamos": nump or 0
        })
    return render_template("seleccion.html", mensaje=None, items=listing, titulo=titulo)

# === Actions — Admin only ===
@app.route("/prestar/<codigo>", methods=["GET","POST"])
@admin_required
def prestar(codigo):
    if request.method == "POST":
        alumno = (request.form.get("alumno") or "").strip()
        actualizar_estado(codigo, "Prestado", session.get("user"), prestado_a=alumno)
        return redirect(f"/{codigo}")
    row = df[df[COL_CODIGO].astype(str).str.strip().str.lower() == codigo.strip().lower()]
    desc = str(row.iloc[0][COL_DESC]) if not row.empty else "(Descripción no encontrada)"
    return render_template("prestar.html", codigo=codigo, desc=desc)

@app.route("/devolver/<codigo>", methods=["GET","POST"])
@admin_required
def devolver(codigo):
    if request.method == "POST":
        actualizar_estado(codigo, "Disponible", session.get("user"))
        return redirect(f"/{codigo}")
    row = df[df[COL_CODIGO].astype(str).str.strip().str.lower() == codigo.strip().lower()]
    desc = str(row.iloc[0][COL_DESC]) if not row.empty else "(Descripción no encontrada)"
    return render_template("devolver.html", codigo=codigo, desc=desc)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)