from flask import Flask, render_template, request, send_file, redirect, url_for, flash, session
import pandas as pd
import os
import json
import re
import secrets
import unicodedata
from functools import wraps
from datetime import datetime
import pdfkit
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = "cotizador-dasmex"

CATALOGO_PATH = "catalogo_nuevo.xlsx"
COLUMNAS_CATALOGO = ["MEDIDA", "PRODUCTO", "CODIGO SKU", "CLAVE SAT"]
USUARIOS_PATH = "usuarios.json"
ADMIN_USUARIO = "DASMEXORIGIN"
ADMIN_PASSWORD = "DASMEX 2026"

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
# USUARIOS Y SESIONES
# ==============================
def limpiar_texto(valor):
    texto = unicodedata.normalize("NFKD", valor or "")
    texto = texto.encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"[^A-Za-z0-9]", "", texto)
    return texto.upper()


def cargar_usuarios():
    if not os.path.exists(USUARIOS_PATH):
        usuarios = {}
    else:
        with open(USUARIOS_PATH, "r", encoding="utf-8") as archivo:
            usuarios = json.load(archivo)

    if ADMIN_USUARIO not in usuarios:
        usuarios[ADMIN_USUARIO] = {
            "password_hash": generate_password_hash(ADMIN_PASSWORD),
            "rol": "admin",
            "nombre": "Administrador DASMEX",
            "cargo": "Administrador",
            "anio_nacimiento": ""
        }
        guardar_usuarios(usuarios)

    return usuarios


def guardar_usuarios(usuarios):
    with open(USUARIOS_PATH, "w", encoding="utf-8") as archivo:
        json.dump(usuarios, archivo, indent=4, ensure_ascii=False)


def generar_credenciales(nombre, cargo, anio_nacimiento):
    nombre_limpio = limpiar_texto(nombre)
    cargo_limpio = limpiar_texto(cargo)
    anio_limpio = re.sub(r"[^0-9]", "", anio_nacimiento or "")

    base_usuario = f"{nombre_limpio[:8]}{cargo_limpio[:5]}{anio_limpio[-2:]}" or "USUARIO"
    usuarios = cargar_usuarios()
    usuario = base_usuario
    contador = 1

    while usuario in usuarios:
        contador += 1
        usuario = f"{base_usuario}{contador}"

    password = f"DASMEX-{anio_limpio or '2026'}-{secrets.token_hex(2).upper()}"
    return usuario, password


def login_requerido(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            flash("Inicia sesion para entrar al cotizador.", "error")
            return redirect(url_for("login"))

        return func(*args, **kwargs)

    return wrapper


def admin_requerido(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if "usuario" not in session:
            flash("Inicia sesion para entrar al panel.", "error")
            return redirect(url_for("login"))

        if session.get("rol") != "admin":
            flash("Solo el administrador puede entrar a este panel.", "error")
            return redirect(url_for("index"))

        return func(*args, **kwargs)

    return wrapper

# ==============================
# CARGAR EXCEL
# ==============================
def cargar_catalogo():
    try:
        # Intenta primero con encabezado normal
        df_local = pd.read_excel(CATALOGO_PATH)

        # Si las columnas no vienen bien, intenta con header=2
        columnas = [str(c).strip().upper() for c in df_local.columns]

        if "PRODUCTO" not in columnas:
            df_local = pd.read_excel(CATALOGO_PATH, header=2)
            columnas = [str(c).strip().upper() for c in df_local.columns]

        # Normalizar nombres
        df_local.columns = columnas

        # Tomar solo las 4 columnas necesarias si existen
        columnas_requeridas = COLUMNAS_CATALOGO

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
        return pd.DataFrame(columns=COLUMNAS_CATALOGO)


def normalizar_catalogo(df_origen):
    columnas = [str(c).strip().upper() for c in df_origen.columns]
    df_normalizado = df_origen.copy()
    df_normalizado.columns = columnas

    if "PRODUCTO" not in df_normalizado.columns:
        if len(df_normalizado.columns) < 4:
            raise ValueError("El archivo Excel no tiene suficientes columnas.")

        df_normalizado = df_normalizado.iloc[:, :4]
        df_normalizado.columns = COLUMNAS_CATALOGO

    if not all(col in df_normalizado.columns for col in COLUMNAS_CATALOGO):
        if len(df_normalizado.columns) < 4:
            raise ValueError("El archivo Excel no tiene suficientes columnas.")

        df_normalizado = df_normalizado.iloc[:, :4]
        df_normalizado.columns = COLUMNAS_CATALOGO

    df_normalizado = df_normalizado[COLUMNAS_CATALOGO]
    df_normalizado = df_normalizado.dropna(subset=["PRODUCTO"])
    df_normalizado = df_normalizado.fillna("")
    df_normalizado = df_normalizado.astype(str)
    df_normalizado = df_normalizado.apply(lambda col: col.str.strip())
    df_normalizado = df_normalizado[df_normalizado["PRODUCTO"] != ""]
    df_normalizado = df_normalizado.reset_index(drop=True)

    return df_normalizado


def guardar_catalogo(df_catalogo):
    df_catalogo.to_excel(CATALOGO_PATH, index=False)


def recargar_catalogo():
    global df
    df = cargar_catalogo()
    return df


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
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        password = request.form.get("password", "")
        usuarios = cargar_usuarios()
        usuario_data = usuarios.get(usuario)

        if usuario_data and check_password_hash(usuario_data["password_hash"], password):
            session["usuario"] = usuario
            session["rol"] = usuario_data.get("rol", "usuario")
            session["nombre"] = usuario_data.get("nombre", usuario)
            flash("Sesion iniciada correctamente.", "success")
            return redirect(url_for("index"))

        flash("Usuario o contrasena incorrectos.", "error")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Sesion cerrada correctamente.", "success")
    return redirect(url_for("login"))


@app.route("/admin/usuarios", methods=["GET", "POST"])
@admin_requerido
def admin_usuarios():
    credenciales_generadas = None

    if request.method == "POST":
        nombre = request.form.get("nombre", "").strip()
        cargo = request.form.get("cargo", "").strip()
        anio_nacimiento = request.form.get("anio_nacimiento", "").strip()
        admin_password = request.form.get("admin_password", "")

        if not nombre or not cargo or not anio_nacimiento or not admin_password:
            flash("Completa nombre, cargo, ano de nacimiento y contrasena admin.", "error")
            return redirect(url_for("admin_usuarios"))

        usuarios = cargar_usuarios()
        admin_data = usuarios.get(ADMIN_USUARIO)

        if not admin_data or not check_password_hash(admin_data["password_hash"], admin_password):
            flash("La contrasena del administrador no es correcta.", "error")
            return redirect(url_for("admin_usuarios"))

        usuario, password = generar_credenciales(nombre, cargo, anio_nacimiento)
        usuarios[usuario] = {
            "password_hash": generate_password_hash(password),
            "rol": "usuario",
            "nombre": nombre,
            "cargo": cargo,
            "anio_nacimiento": anio_nacimiento
        }
        guardar_usuarios(usuarios)

        credenciales_generadas = {
            "usuario": usuario,
            "password": password,
            "nombre": nombre
        }
        flash("Usuario generado correctamente.", "success")

    usuarios = cargar_usuarios()
    return render_template(
        "admin_usuarios.html",
        usuarios=usuarios,
        credenciales_generadas=credenciales_generadas
    )


@app.route("/")
@login_requerido
def index():
    productos = df.to_dict(orient="records")
    return render_template("index.html", productos=productos)


@app.route("/agregar_producto", methods=["POST"])
@login_requerido
def agregar_producto():
    medida = request.form.get("medida", "").strip()
    producto = request.form.get("producto", "").strip()
    codigo_sku = request.form.get("codigo_sku", "").strip()
    clave_sat = request.form.get("clave_sat", "").strip()

    if not producto:
        flash("Debes capturar al menos el nombre del producto.", "error")
        return redirect(url_for("index"))

    nuevo_producto = pd.DataFrame([{
        "MEDIDA": medida,
        "PRODUCTO": producto,
        "CODIGO SKU": codigo_sku,
        "CLAVE SAT": clave_sat
    }])

    catalogo_actualizado = pd.concat([df, nuevo_producto], ignore_index=True)
    catalogo_actualizado = normalizar_catalogo(catalogo_actualizado)
    guardar_catalogo(catalogo_actualizado)
    recargar_catalogo()

    flash("Producto agregado correctamente al cotizador.", "success")
    return redirect(url_for("index"))


@app.route("/cargar_excel", methods=["POST"])
@login_requerido
def cargar_excel():
    archivo = request.files.get("archivo_excel")

    if not archivo or not archivo.filename:
        flash("Selecciona un archivo Excel para cargar.", "error")
        return redirect(url_for("index"))

    extension = os.path.splitext(archivo.filename)[1].lower()
    if extension not in [".xlsx", ".xls"]:
        flash("El archivo debe ser Excel (.xlsx o .xls).", "error")
        return redirect(url_for("index"))

    try:
        df_excel = pd.read_excel(archivo)
        try:
            catalogo_nuevo = normalizar_catalogo(df_excel)
        except ValueError:
            archivo.seek(0)
            catalogo_nuevo = normalizar_catalogo(pd.read_excel(archivo, header=2))
    except Exception as e:
        flash(f"No se pudo leer el Excel: {e}", "error")
        return redirect(url_for("index"))

    if catalogo_nuevo.empty:
        flash("El Excel no contiene productos validos para agregar.", "error")
        return redirect(url_for("index"))

    catalogo_actualizado = pd.concat([df, catalogo_nuevo], ignore_index=True)
    catalogo_actualizado = catalogo_actualizado.drop_duplicates(
        subset=["PRODUCTO", "CODIGO SKU", "CLAVE SAT"],
        keep="last"
    )
    catalogo_actualizado = normalizar_catalogo(catalogo_actualizado)
    guardar_catalogo(catalogo_actualizado)
    recargar_catalogo()

    flash(f"Se agregaron {len(catalogo_nuevo)} productos desde el Excel.", "success")
    return redirect(url_for("index"))

# ==============================
# GENERAR PDF
# ==============================
@app.route("/generar_pdf", methods=["POST"])
@login_requerido
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
