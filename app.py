import io
import re
import zipfile

import pandas as pd
import qrcode
import streamlit as st
from PIL import Image


DEFAULT_BASE_URL = "https://comercio.munipachacamac.gob.pe"
DEFAULT_TOKEN = (
    "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwOi8vMTkyLjE2OC4xMC4zNzo4MDAwL2FwaS9sb2dpbiIsImlhdCI6MTcxOTUxNjAyMywiZXhwIjoxNzE5NTE5NjIzLCJuYmYiOjE3MTk1MTYwMjMsImp0aSI6IjhURWlHZ2ZQaTFOQkx5UjgiLCJzdWIiOi"
)


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
    archivo = st.file_uploader("Archivo .xlsx", type=["xlsx"])

    if archivo:
        try:
            df_excel = pd.read_excel(archivo, dtype=str)
            st.write("Vista previa:")
            st.dataframe(df_excel, use_container_width=True)

            columnas = list(df_excel.columns)
            col_codigo = st.selectbox(
                "Columna de codigo",
                options=columnas,
                index=columnas.index("codigo") if "codigo" in columnas else 0,
            )
            col_nombre = st.selectbox(
                "Columna de nombre",
                options=columnas,
                index=columnas.index("nombre") if "nombre" in columnas else min(1, len(columnas) - 1),
            )

            df_preparado = pd.DataFrame(
                {
                    "codigo": df_excel[col_codigo],
                    "nombre": df_excel[col_nombre],
                }
            )
            st.session_state.df_final = preparar_dataframe(df_preparado)
        except Exception as error:
            st.error(f"No pude leer el Excel: {error}")


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
