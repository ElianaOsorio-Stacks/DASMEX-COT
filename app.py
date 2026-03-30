from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime
import pdfkit

app = Flask(__name__)

# ==============================
# CONFIGURAR WKHTMLTOPDF
# ==============================
WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

options = {
    "enable-local-file-access": "",
    "encoding": "UTF-8",
    "page-size": "A4",
    "margin-top": "0mm",
    "margin-right": "0mm",
    "margin-bottom": "0mm",
    "margin-left": "0mm",
    "zoom": "1.0",
    "dpi": "300"
}

# ==============================
# CARGAR EXCEL
# ==============================
def cargar_catalogo():
    try:
        # Intenta primero con encabezado normal
        df_local = pd.read_excel("catalogo_nuevo.xlsx")

        # Si las columnas no vienen bien, intenta con header=2
        columnas = [str(c).strip().upper() for c in df_local.columns]

        if "PRODUCTO" not in columnas:
            df_local = pd.read_excel("catalogo_nuevo.xlsx", header=2)
            columnas = [str(c).strip().upper() for c in df_local.columns]

        # Normalizar nombres
        df_local.columns = columnas

        # Tomar solo las 4 columnas necesarias si existen
        columnas_requeridas = ["MEDIDA", "PRODUCTO", "CODIGO SKU", "CLAVE SAT"]

        # Si no vienen exactas, renombrar por posición
        if not all(col in df_local.columns for col in columnas_requeridas):
            if len(df_local.columns) >= 4:
                df_local = df_local.iloc[:, :4]
                df_local.columns = columnas_requeridas
            else:
                raise ValueError("El archivo Excel no tiene suficientes columnas.")

        df_local = df_local.dropna(subset=["PRODUCTO"])
        df_local = df_local.reset_index(drop=True)

        return df_local

    except Exception as e:
        print("❌ Error leyendo Excel:", e)
        return pd.DataFrame(columns=["MEDIDA", "PRODUCTO", "CODIGO SKU", "CLAVE SAT"])


df = cargar_catalogo()

# ==============================
# GENERADOR DE FOLIO
# ==============================
def generar_folio():
    archivo = "folio.txt"

    if not os.path.exists(archivo):
        with open(archivo, "w", encoding="utf-8") as f:
            f.write("1")

    with open(archivo, "r", encoding="utf-8") as f:
        contenido = f.read().strip()
        numero = int(contenido) if contenido.isdigit() else 1

    folio = f"DAS-{numero:04d}"

    with open(archivo, "w", encoding="utf-8") as f:
        f.write(str(numero + 1))

    return folio

# ==============================
# RUTA PRINCIPAL
# ==============================
@app.route("/")
def index():
    productos = df.to_dict(orient="records")
    return render_template("index.html", productos=productos)

# ==============================
# GENERAR PDF
# ==============================
@app.route("/generar_pdf", methods=["POST"])
def generar_pdf():
    cliente = request.form.get("cliente", "").strip()

    if not cliente:
        return "⚠️ Debes capturar el nombre del cliente."

    items = []
    subtotal = 0.0

    for i in range(len(df)):
        marcado = request.form.get(f"check_{i}")

        if marcado:
            try:
                cantidad = float(request.form.get(f"cantidad_{i}", 0) or 0)
                precio = float(request.form.get(f"precio_{i}", 0) or 0)
            except ValueError:
                cantidad = 0
                precio = 0

            if cantidad <= 0:
                continue

            producto = df.iloc[i]
            total = cantidad * precio
            subtotal += total

            items.append({
                "cantidad": int(cantidad) if cantidad.is_integer() else cantidad,
                "unidad": str(producto["MEDIDA"]).strip(),
                "medida": str(producto["MEDIDA"]).strip(),
                "producto": str(producto["PRODUCTO"]).strip(),
                "codigo": str(producto["CODIGO SKU"]).strip(),
                "clave_sat": str(producto["CLAVE SAT"]).strip(),
                "precio": f"{precio:.2f}",
                "total": f"{total:.2f}"
            })

    if not items:
        return "⚠️ No seleccionaste ningún producto con cantidad válida."

    descuento = subtotal * 0.035
    iva = (subtotal - descuento) * 0.16
    total_final = subtotal - descuento + iva

    folio = generar_folio()
    fecha = datetime.now().strftime("%d/%m/%Y")

    # Ruta absoluta a static para wkhtmltopdf
    ruta_static = os.path.abspath("static").replace("\\", "/")

    # Render del HTML del PDF
    html = render_template(
        "pdf.html",
        items=items,
        cliente=cliente,
        subtotal=f"{subtotal:.2f}",
        descuento=f"{descuento:.2f}",
        iva=f"{iva:.2f}",
        total=f"{total_final:.2f}",
        fecha=fecha,
        folio=folio,
        ruta_static=ruta_static
    )

    os.makedirs("pdfs", exist_ok=True)
    pdf_path = os.path.abspath(f"pdfs/{folio}.pdf")

    try:
        pdfkit.from_string(
            html,
            pdf_path,
            configuration=config,
            options=options
        )
    except Exception as e:
        return f"❌ Error generando PDF: {e}"

    return send_file(pdf_path, as_attachment=True)

# ==============================
# RUN
# ==============================
if __name__ == "__main__":
    app.run(debug=True)