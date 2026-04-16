import io
import posixpath
import re
import unicodedata
import zipfile
from xml.etree import ElementTree as ET

import pandas as pd
import qrcode
import streamlit as st
from PIL import Image


DEFAULT_BASE_URL = "https://comercio.munipachacamac.gob.pe"
DEFAULT_TOKEN = (
    "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwOi8vMTkyLjE2OC4xMC4zNzo4MDAwL2FwaS9sb2dpbiIsImlhdCI6MTcxOTUxNjAyMywiZXhwIjoxNzE5NTE5NjIzLCJuYmYiOjE3MTk1MTYwMjMsImp0aSI6IjhURWlHZ2ZQaTFOQkx5UjgiLCJzdWIiOi"
)
EXCEL_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "office": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "package": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def limpiar_nombre(nombre: str) -> str:
    nombre = re.sub(r'[\/:*?"<>|]', "", nombre or "")
    nombre = nombre.strip()
    return nombre[:50] if nombre else "SIN_NOMBRE"


def normalizar_codigo(valor) -> str:
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()
    if re.fullmatch(r"\d+\.0", texto):
        texto = texto[:-2]
    return texto


def construir_url(base_url: str, codigo: str, token: str, incluir_token: bool) -> str:
    base_limpia = base_url.rstrip("/")
    codigo_limpio = normalizar_codigo(codigo)

    if incluir_token:
        return f"{base_limpia}/{codigo_limpio}/{token}"
    return f"{base_limpia}/{codigo_limpio}"


def generar_qr_png_bytes(url: str, size_px: int) -> bytes:
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    img = img.resize((size_px, size_px), Image.LANCZOS)

    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    return buffer.getvalue()


def armar_zip(df: pd.DataFrame, base_url: str, token: str, size_px: int, incluir_token: bool) -> bytes:
    zip_buffer = io.BytesIO()
    total = len(df)
    progreso = st.progress(0, text="Generando QRs...")

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for indice, fila in enumerate(df.itertuples(index=False), start=1):
            url = construir_url(base_url, fila.codigo, token, incluir_token)
            png_bytes = generar_qr_png_bytes(url, size_px=size_px)
            filename = f"{fila.codigo} {limpiar_nombre(fila.nombre)}.png"
            zf.writestr(filename, png_bytes)
            progreso.progress(indice / total, text=f"Generando {indice}/{total}...")

    progreso.empty()
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def preparar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    salida = df.copy()
    salida["codigo"] = salida["codigo"].map(normalizar_codigo)
    salida["nombre"] = salida["nombre"].fillna("").astype(str).str.strip()
    salida = salida[salida["codigo"].astype(str).str.strip() != ""]
    salida = salida.drop_duplicates(subset=["codigo"], keep="first")
    return salida.reset_index(drop=True)


def normalizar_encabezado(valor) -> str:
    texto = str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(caracter for caracter in texto if not unicodedata.combining(caracter))
    return re.sub(r"[^a-z0-9]+", "", texto)


def detectar_columna(columnas, alias_validos: set[str]) -> int | None:
    columnas_normalizadas = [normalizar_encabezado(columna) for columna in columnas]

    for indice, columna in enumerate(columnas_normalizadas):
        if columna in alias_validos:
            return indice

    for indice, columna in enumerate(columnas_normalizadas):
        if any(alias in columna for alias in alias_validos):
            return indice

    return None


def columna_desde_referencia_excel(referencia: str) -> int:
    letras = "".join(caracter for caracter in (referencia or "").upper() if caracter.isalpha())
    indice = 0
    for letra in letras:
        indice = (indice * 26) + (ord(letra) - ord("A") + 1)
    return max(indice - 1, 0)


def hacer_columnas_unicas(columnas) -> list[str]:
    columnas_limpias = []
    usados: dict[str, int] = {}

    for indice, columna in enumerate(columnas, start=1):
        base = str(columna).strip() or f"columna_{indice}"
        repeticion = usados.get(base, 0)
        nombre = base if repeticion == 0 else f"{base}_{repeticion + 1}"
        usados[base] = repeticion + 1
        columnas_limpias.append(nombre)

    return columnas_limpias


def cargar_shared_strings_xlsx(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []

    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    valores = []

    for item in root.findall("main:si", EXCEL_NS):
        textos = [nodo.text or "" for nodo in item.findall(".//main:t", EXCEL_NS)]
        valores.append("".join(textos))

    return valores


def obtener_primera_hoja_xlsx(zf: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    relaciones = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    primer_sheet = workbook.find("main:sheets/main:sheet", EXCEL_NS)

    if primer_sheet is None:
        raise ValueError("El archivo Excel no tiene hojas disponibles.")

    relacion_id = primer_sheet.attrib.get(f"{{{EXCEL_NS['office']}}}id")
    targets = {
        relacion.attrib["Id"]: relacion.attrib["Target"]
        for relacion in relaciones.findall("package:Relationship", EXCEL_NS)
    }
    target = targets.get(relacion_id)

    if not target:
        raise ValueError("No pude ubicar la primera hoja del Excel.")

    ruta_hoja = target.lstrip("/")
    if not ruta_hoja.startswith("xl/"):
        ruta_hoja = posixpath.join("xl", ruta_hoja)

    return posixpath.normpath(ruta_hoja)


def valor_celda_xlsx(celda: ET.Element, shared_strings: list[str]) -> str:
    tipo = celda.attrib.get("t")
    valor = celda.findtext("main:v", default="", namespaces=EXCEL_NS)

    if tipo == "inlineStr":
        return "".join(nodo.text or "" for nodo in celda.findall(".//main:t", EXCEL_NS))
    if tipo == "s":
        try:
            return shared_strings[int(valor)]
        except (ValueError, IndexError):
            return valor
    if tipo == "b":
        return "TRUE" if valor == "1" else "FALSE"
    if tipo == "e":
        return ""
    return valor or ""


def leer_xlsx_sin_openpyxl(archivo) -> pd.DataFrame:
    archivo.seek(0)

    with zipfile.ZipFile(archivo) as zf:
        shared_strings = cargar_shared_strings_xlsx(zf)
        hoja_path = obtener_primera_hoja_xlsx(zf)
        hoja = ET.fromstring(zf.read(hoja_path))

    filas = []

    for fila in hoja.findall("main:sheetData/main:row", EXCEL_NS):
        celdas = {}
        max_columna = -1

        for celda in fila.findall("main:c", EXCEL_NS):
            referencia = celda.attrib.get("r", "")
            columna = columna_desde_referencia_excel(referencia) if referencia else max_columna + 1
            celdas[columna] = valor_celda_xlsx(celda, shared_strings)
            max_columna = max(max_columna, columna)

        if max_columna < 0:
            filas.append([])
            continue

        fila_actual = [""] * (max_columna + 1)
        for columna, valor in celdas.items():
            fila_actual[columna] = valor
        filas.append(fila_actual)

    while filas and not any(str(valor).strip() for valor in filas[0]):
        filas.pop(0)

    if not filas:
        raise ValueError("El Excel no contiene filas legibles.")

    cantidad_columnas = max(len(fila) for fila in filas)
    filas_normalizadas = [fila + [""] * (cantidad_columnas - len(fila)) for fila in filas]
    columnas_con_datos = [
        indice
        for indice in range(cantidad_columnas)
        if any(str(fila[indice]).strip() for fila in filas_normalizadas)
    ]

    if not columnas_con_datos:
        raise ValueError("El Excel no contiene columnas con datos legibles.")

    filas_filtradas = [
        [fila[indice] for indice in columnas_con_datos]
        for fila in filas_normalizadas
    ]
    columnas = hacer_columnas_unicas(filas_filtradas[0])

    return pd.DataFrame(filas_filtradas[1:], columns=columnas)


def leer_archivo_tabular(archivo) -> pd.DataFrame:
    archivo.seek(0)

    try:
        return pd.read_excel(archivo, dtype=str)
    except ImportError as error:
        if "openpyxl" not in str(error).lower():
            raise
        archivo.seek(0)
        return leer_xlsx_sin_openpyxl(archivo)


st.set_page_config(page_title="Generador de QRs", page_icon="QR", layout="wide")
st.title("Generador de QRs")

if "df_final" not in st.session_state:
    st.session_state.df_final = None

if "manual_codigo" not in st.session_state:
    st.session_state.manual_codigo = ""

if "manual_nombre" not in st.session_state:
    st.session_state.manual_nombre = ""


with st.sidebar:
    st.header("Configuracion")
    base_url = st.text_input(
        "Base URL",
        value=st.secrets.get("BASE_URL", DEFAULT_BASE_URL),
    )
    token = st.text_input(
        "TOKEN",
        value=st.secrets.get("TOKEN_JWT", DEFAULT_TOKEN),
        type="password",
    )
    incluir_token = st.toggle("Incluir token en la URL", value=True)
    size_px = st.select_slider(
        "Tamano del QR (px)",
        options=[512, 768, 1024, 1280, 1536, 1920],
        value=1920,
    )


tab_excel, tab_manual = st.tabs(["Carga masiva (Excel)", "Ingreso manual"])


with tab_excel:
    st.subheader("Sube tu Excel")
    archivo = st.file_uploader("Archivo Excel (.xlsx o .xlsm)", type=["xlsx", "xlsm"])

    if archivo:
        try:
            df_excel = leer_archivo_tabular(archivo)
            st.write("Vista previa:")
            st.dataframe(df_excel, use_container_width=True)

            columnas = list(df_excel.columns)
            indice_codigo = detectar_columna(columnas, {"codigo", "cod"})
            indice_nombre = detectar_columna(columnas, {"nombre", "nombres", "nombrecompleto"})

            if indice_codigo is None or indice_nombre is None:
                st.info(
                    "Puedes subir cualquier Excel mientras tenga una columna de codigo y otra de nombre. "
                    "Si no las detecto automaticamente, elige ambas columnas aqui."
                )

            col_codigo = st.selectbox(
                "Columna de codigo",
                options=columnas,
                index=indice_codigo if indice_codigo is not None else 0,
            )
            col_nombre = st.selectbox(
                "Columna de nombre",
                options=columnas,
                index=indice_nombre if indice_nombre is not None else min(1, len(columnas) - 1),
            )

            if col_codigo == col_nombre:
                st.warning("Selecciona columnas distintas para codigo y nombre.")
            else:
                df_preparado = pd.DataFrame(
                    {
                        "codigo": df_excel[col_codigo],
                        "nombre": df_excel[col_nombre],
                    }
                )
                st.session_state.df_final = preparar_dataframe(df_preparado)
        except Exception as error:
            st.error(f"No pude leer el Excel: {error}")
            st.caption("Si tu archivo es .xlsx, la app intentara leerlo incluso cuando falte openpyxl.")


with tab_manual:
    st.subheader("Generar un QR manualmente")

    with st.form("manual_qr_form", clear_on_submit=False):
        codigo_manual = st.text_input("Codigo", key="manual_codigo")
        nombre_manual = st.text_input("Nombre", key="manual_nombre")
        enviado = st.form_submit_button("Usar estos datos", type="primary")

    if enviado:
        df_manual = pd.DataFrame(
            [
                {
                    "codigo": codigo_manual,
                    "nombre": nombre_manual,
                }
            ]
        )
        st.session_state.df_final = preparar_dataframe(df_manual)

        if st.session_state.df_final.empty:
            st.warning("Ingresa al menos un codigo valido.")
        else:
            st.success("Datos manuales cargados correctamente.")


df_final = st.session_state.df_final.copy() if st.session_state.df_final is not None else None

if df_final is None:
    st.info("Sube un Excel o usa el formulario manual para habilitar la generacion.")
else:
    st.write(f"Registros validos: **{len(df_final)}**")
    st.dataframe(df_final, use_container_width=True)

    if len(df_final) == 0:
        st.warning("No hay codigos validos para generar.")
    elif incluir_token and not token:
        st.warning("Tienes activado 'Incluir token' pero el token esta vacio.")
    elif len(df_final) == 1:
        fila = df_final.iloc[0]
        url_qr = construir_url(base_url, fila["codigo"], token, incluir_token)
        png_bytes = generar_qr_png_bytes(url_qr, size_px=size_px)
        nombre_archivo = f"{fila['codigo']} {limpiar_nombre(fila['nombre'])}.png"

        st.write("Se generara un solo archivo PNG, no un ZIP.")
        st.download_button(
            label="Descargar QR en PNG",
            data=png_bytes,
            file_name=nombre_archivo,
            mime="image/png",
            type="primary",
        )
    else:
        if st.button("Generar ZIP de QRs", type="primary"):
            zip_bytes = armar_zip(df_final, base_url, token, size_px, incluir_token)
            st.success("Listo. Descarga tu ZIP:")
            st.download_button(
                label="Descargar QRS_GENERADOS.zip",
                data=zip_bytes,
                file_name="QRS_GENERADOS.zip",
                mime="application/zip",
            )
