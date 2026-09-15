"""
Aplicación Flask para automatizar el cruce de datos de las
Elecciones del SENA (Votantes vs. Base de Datos de Matrículas).

Autor: Claude (Anthropic)
"""

import os
import uuid
import traceback

import pandas as pd
from flask import (
    Flask,
    render_template,
    request,
    send_file,
    flash,
    redirect,
    url_for,
)

# ---------------------------------------------------------------------------
# Configuración de la aplicación
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
RESULT_FOLDER = os.path.join(BASE_DIR, "resultados")
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.secret_key = "cambia-esta-clave-por-una-segura"  # Necesario para flash()
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["RESULT_FOLDER"] = RESULT_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB máx. por archivo

# Nombre final del archivo consolidado que descargará el usuario
NOMBRE_ARCHIVO_SALIDA = "Elecciones_Sena_Consolidado.xlsx"


def allowed_file(filename: str) -> bool:
    """Valida que el archivo tenga extensión .xlsx o .xls"""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ---------------------------------------------------------------------------
# Lógica central: cruce de datos con Pandas
# ---------------------------------------------------------------------------
def procesar_cruce(path_votantes: str, path_matriculas: str) -> pd.DataFrame:
    """
    Lee el archivo de votantes y el de matrículas del SENA, estandariza
    las columnas clave y realiza un merge tipo 'left' para anexar
    Ficha, Nombres, Apellidos y Programa de formación a cada votante.

    Parámetros
    ----------
    path_votantes : str
        Ruta al Excel de votantes (hoja 'Votantes').
    path_matriculas : str
        Ruta al Excel oficial de matrículas del SENA (cabecera en fila 5).

    Retorna
    -------
    pd.DataFrame
        DataFrame consolidado listo para exportar.
    """

    # -----------------------------------------------------------------
    # 1. Lectura del archivo de VOTANTES (hoja 'Votantes')
    # -----------------------------------------------------------------
    try:
        df_votantes = pd.read_excel(path_votantes, sheet_name="Votantes")
    except ValueError:
        # Si no existe una hoja llamada exactamente 'Votantes', se usa la primera
        df_votantes = pd.read_excel(path_votantes, sheet_name=0)

    columnas_esperadas_votantes = {"Cedula", "FechaHora", "CandidatoId"}
    faltantes = columnas_esperadas_votantes - set(df_votantes.columns)
    if faltantes:
        raise ValueError(
            f"El archivo de votantes no contiene las columnas esperadas: {faltantes}. "
            f"Columnas encontradas: {list(df_votantes.columns)}"
        )

    # -----------------------------------------------------------------
    # 2. Lectura del archivo de MATRÍCULAS (cabecera real en la fila 5,
    #    es decir header=4 porque pandas indexa desde 0)
    # -----------------------------------------------------------------
    df_matriculas = pd.read_excel(path_matriculas, header=4)

    columnas_esperadas_matriculas = {
        "FICHA",
        "NUMERO_DOCUMENTO",
        "NOMBRE",
        "PRIMER_APELLIDO",
        "SEGUNDO_APELLIDO",
        "PROGRAMA",
        "ESTADO_APRENDIZ",
    }
    faltantes_matriculas = columnas_esperadas_matriculas - set(df_matriculas.columns)
    if faltantes_matriculas:
        raise ValueError(
            "El archivo de matrículas no contiene las columnas esperadas: "
            f"{faltantes_matriculas}. Columnas encontradas: {list(df_matriculas.columns)}"
        )

    # -----------------------------------------------------------------
    # 3. Estandarización de las columnas clave (cédula / documento)
    #    Se convierten a texto y se les quita espacios en blanco para
    #    evitar fallos de cruce por formatos distintos (float vs str).
    # -----------------------------------------------------------------
    df_votantes["Cedula"] = (
        df_votantes["Cedula"]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)  # por si pandas la leyó como float
    )

    df_matriculas["NUMERO_DOCUMENTO"] = (
        df_matriculas["NUMERO_DOCUMENTO"]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

    # Quitamos posibles filas totalmente vacías de metadatos residuales
    df_matriculas = df_matriculas.dropna(subset=["NUMERO_DOCUMENTO"])
    df_matriculas = df_matriculas[df_matriculas["NUMERO_DOCUMENTO"] != "nan"]

    # Si un mismo aprendiz aparece en varias fichas, nos quedamos con el
    # primer registro para no duplicar votos en el merge.
    df_matriculas_unico = df_matriculas.drop_duplicates(
        subset=["NUMERO_DOCUMENTO"], keep="first"
    )

    # -----------------------------------------------------------------
    # 4. Merge (cruce) tipo LEFT: se conservan TODOS los votantes,
    #    y se anexa la información de matrícula cuando exista coincidencia.
    # -----------------------------------------------------------------
    columnas_a_anexar = [
        "NUMERO_DOCUMENTO",
        "FICHA",
        "NOMBRE",
        "PRIMER_APELLIDO",
        "SEGUNDO_APELLIDO",
        "PROGRAMA",
        "ESTADO_APRENDIZ",
    ]

    df_final = df_votantes.merge(
        df_matriculas_unico[columnas_a_anexar],
        how="left",
        left_on="Cedula",
        right_on="NUMERO_DOCUMENTO",
    )

    # Columna de control: indica si el votante fue encontrado en la BD
    df_final["Cruce_Estado"] = df_final["NUMERO_DOCUMENTO"].apply(
        lambda x: "Encontrado" if pd.notna(x) else "NO ENCONTRADO"
    )

    # FICHA suele quedar como float (ej. 2589631.0) por los NaN que introduce
    # el merge en los no encontrados. La limpiamos para que se vea como texto/entero.
    if "FICHA" in df_final.columns:
        df_final["FICHA"] = df_final["FICHA"].apply(
            lambda x: str(int(x)) if pd.notna(x) else ""
        )

    # Nombre completo para facilitar la lectura del reporte final
    df_final["Nombre_Completo"] = (
        df_final[["NOMBRE", "PRIMER_APELLIDO", "SEGUNDO_APELLIDO"]]
        .fillna("")
        .agg(" ".join, axis=1)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    # Reordenamos columnas para que el reporte sea más legible
    columnas_finales = [
        "Cedula",
        "Nombre_Completo",
        "PRIMER_APELLIDO",
        "SEGUNDO_APELLIDO",
        "FICHA",
        "PROGRAMA",
        "ESTADO_APRENDIZ",
        "CandidatoId",
        "FechaHora",
        "Cruce_Estado",
    ]
    columnas_finales = [c for c in columnas_finales if c in df_final.columns]
    otras_columnas = [c for c in df_final.columns if c not in columnas_finales]
    df_final = df_final[columnas_finales + otras_columnas]

    # Eliminamos la columna técnica duplicada NUMERO_DOCUMENTO
    if "NUMERO_DOCUMENTO" in df_final.columns:
        df_final = df_final.drop(columns=["NUMERO_DOCUMENTO"])

    return df_final


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/procesar", methods=["POST"])
def procesar():
    archivo_votantes = request.files.get("archivo_votantes")
    archivo_matriculas = request.files.get("archivo_matriculas")

    # -------------------- Validaciones básicas --------------------
    if not archivo_votantes or archivo_votantes.filename == "":
        flash("Debes seleccionar el archivo de Votantes.", "danger")
        return redirect(url_for("index"))

    if not archivo_matriculas or archivo_matriculas.filename == "":
        flash("Debes seleccionar el archivo de Matrículas del SENA.", "danger")
        return redirect(url_for("index"))

    if not (allowed_file(archivo_votantes.filename) and allowed_file(archivo_matriculas.filename)):
        flash("Ambos archivos deben tener formato .xlsx o .xls", "danger")
        return redirect(url_for("index"))

    # -------------------- Guardado temporal de los uploads --------------------
    id_proceso = uuid.uuid4().hex[:10]
    ruta_votantes = os.path.join(app.config["UPLOAD_FOLDER"], f"{id_proceso}_votantes.xlsx")
    ruta_matriculas = os.path.join(app.config["UPLOAD_FOLDER"], f"{id_proceso}_matriculas.xlsx")

    archivo_votantes.save(ruta_votantes)
    archivo_matriculas.save(ruta_matriculas)

    # -------------------- Procesamiento --------------------
    try:
        df_resultado = procesar_cruce(ruta_votantes, ruta_matriculas)
    except Exception as exc:  # noqa: BLE001
        traceback.print_exc()
        flash(f"Ocurrió un error al procesar los archivos: {exc}", "danger")
        return redirect(url_for("index"))
    finally:
        # Limpieza de los archivos subidos (ya no se necesitan)
        for ruta in (ruta_votantes, ruta_matriculas):
            if os.path.exists(ruta):
                os.remove(ruta)

    # -------------------- Guardar resultado consolidado --------------------
    nombre_resultado = f"{id_proceso}_{NOMBRE_ARCHIVO_SALIDA}"
    ruta_resultado = os.path.join(app.config["RESULT_FOLDER"], nombre_resultado)
    df_resultado.to_excel(ruta_resultado, index=False, sheet_name="Consolidado")

    # -------------------- Preparar vista previa (primeras filas) --------------------
    total_votantes = len(df_resultado)
    total_encontrados = int((df_resultado["Cruce_Estado"] == "Encontrado").sum())
    total_no_encontrados = total_votantes - total_encontrados

    tabla_preview = df_resultado.head(15).to_html(
        classes="table table-striped table-hover table-bordered align-middle",
        index=False,
        na_rep="-",
        border=0,
    )

    return render_template(
        "index.html",
        resultado_listo=True,
        nombre_descarga=nombre_resultado,
        tabla_preview=tabla_preview,
        total_votantes=total_votantes,
        total_encontrados=total_encontrados,
        total_no_encontrados=total_no_encontrados,
    )


@app.route("/descargar/<nombre_archivo>", methods=["GET"])
def descargar(nombre_archivo):
    ruta_resultado = os.path.join(app.config["RESULT_FOLDER"], nombre_archivo)

    if not os.path.exists(ruta_resultado):
        flash("El archivo solicitado ya no está disponible. Procesa nuevamente.", "warning")
        return redirect(url_for("index"))

    return send_file(
        ruta_resultado,
        as_attachment=True,
        download_name=NOMBRE_ARCHIVO_SALIDA,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
