from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

# Carga del archivo Excel (solo primera hoja)
df = pd.read_excel("INVENTARIO.xlsx", sheet_name=0)

@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    if request.method == "POST":
        codigo = request.form["codigo"].strip()
        # Buscar en la segunda columna (índice 1)
        row = df[df.iloc[:, 1].astype(str).str.strip() == codigo]
        if not row.empty:
            result = row.to_dict(orient="records")[0]
        else:
            result = "No se encontró el código."
    return render_template("index.html", result=result)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
