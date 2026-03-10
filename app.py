import re
import io
import zipfile
import pandas as pd
import streamlit as st
import qrcode
from PIL import Image

# ---------------------------
# Utilidades
# ---------------------------
def limpiar_nombre(nombre: str) -> str:
    nombre = re.sub(r'[\/:*?"<>|]', '', nombre or "")
    nombre = nombre.strip()
    return nombre[:50] if nombre else "SIN_NOMBRE"

def normalizar_codigo(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # Si viene como 213748.0 desde Excel, lo limpiamos
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s

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

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def armar_zip(df: pd.DataFrame, base_url: str, token: str, size_px: int, incluir_token: bool) -> bytes:
    zip_buffer = io.BytesIO()
    total = len(df)

    prog = st.progress(0, text="Generando QRs...")

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for i, row in enumerate(df.itertuples(index=False), start=1):
            codigo = row.codigo
            nombre = row.nombre

            if incluir_token:
                url = f"{base_url}/{codigo}/{token}"
            else:
                url = f"{base_url}/{codigo}"

            png_bytes = generar_qr_png_bytes(url, size_px=size_px)
            filename = f"{codigo} {limpiar_nombre(nombre)}.png"
            zf.writestr(filename, png_bytes)

            prog.progress(i / total, text=f"Generando {i}/{total}...")

    prog.empty()
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# ---------------------------
# App
# ---------------------------
st.set_page_config(page_title="Generador de QRs", page_icon="📦", layout="wide")
st.title("📦 Generador de QRs (Excel / Manual)")

# Config en sidebar
with st.sidebar:
    st.header("⚙️ Configuración")
    base_url = st.text_input("Base URL", value=st.secrets.get("BASE_URL", "https://comercio.munipachacamac.gob.pe"))
    token = st.text_input("TOKEN (secrets)", value=st.secrets.get("TOKEN_JWT", ""), type="password")
    incluir_token = st.toggle("Incluir token en la URL", value=True)
    size_px = st.select_slider("Tamaño del QR (px)", options=[512, 768, 1024, 1280, 1536, 1920], value=1920)

tab1, tab2 = st.tabs(["📁 Carga masiva (Excel)", "✍️ Ingreso manual"])

# guardamos el dataframe final en session_state para que sobreviva a los reruns
if "df_final" not in st.session_state:
    st.session_state.df_final = None

df_final = None  # variable local que se rellenará con el valor de session_state más abajo

# ---- TAB EXCEL ----
with tab1:
    st.subheader("Sube tu Excel")
    up = st.file_uploader("Archivo .xlsx", type=["xlsx"])

    if up:
        try:
            df = pd.read_excel(up, dtype=str)  # dtype=str para evitar floats raros
            st.write("Vista previa:")
            st.dataframe(df, use_container_width=True)

            # Mapeo de columnas (por si cambian nombres)
            cols = list(df.columns)
            col_codigo = st.selectbox("Columna de código", options=cols, index=cols.index("codigo") if "codigo" in cols else 0)
            col_nombre = st.selectbox("Columna de nombre", options=cols, index=cols.index("nombre") if "nombre" in cols else min(1, len(cols)-1))

            df_final = pd.DataFrame({
                "codigo": df[col_codigo].map(normalizar_codigo),
                "nombre": df[col_nombre].fillna("").astype(str).str.strip(),
            })
            st.session_state.df_final = df_final  # guardar para generación

        except Exception as e:
            st.error(f"No pude leer el Excel: {e}")

# ---- TAB MANUAL ----
with tab2:
    st.subheader("Ingresa pocos registros (tabla editable)")

    if "manual_df" not in st.session_state:
        st.session_state.manual_df = pd.DataFrame([{"codigo": "", "nombre": ""}])

    st.session_state.manual_df = st.data_editor(
        st.session_state.manual_df,
        num_rows="dynamic",
        use_container_width=True
    )

    if st.button("Usar estos datos", type="secondary"):
        tmp = st.session_state.manual_df.copy()
        tmp["codigo"] = tmp["codigo"].map(normalizar_codigo)
        tmp["nombre"] = tmp["nombre"].fillna("").astype(str).str.strip()
        st.session_state.df_final = tmp  # enlistar en session_state para uso posterior

# cargar el dataframe desde session_state si existe
if st.session_state.df_final is not None:
    df_final = st.session_state.df_final.copy()
else:
    df_final = None

if df_final is not None:
    # Limpieza y validación
    df_final = df_final.copy()
    df_final = df_final[df_final["codigo"].astype(str).str.strip() != ""]
    df_final = df_final.drop_duplicates(subset=["codigo"], keep="first")

    st.write(f"Registros válidos: **{len(df_final)}**")
    st.dataframe(df_final.head(30), use_container_width=True)

    if len(df_final) == 0:
        st.warning("No hay códigos válidos para generar.")
    else:
        if incluir_token and not token:
            st.warning("Tienes activado 'Incluir token' pero el token está vacío.")
        else:
            if st.button("🚀 Generar ZIP de QRs", type="primary"):
                zip_bytes = armar_zip(df_final, base_url, token, size_px, incluir_token)

                st.success("Listo ✅ Descarga tu ZIP:")
                st.download_button(
                    label="⬇️ Descargar QRS_GENERADOS.zip",
                    data=zip_bytes,
                    file_name="QRS_GENERADOS.zip",
                    mime="application/zip"
                )
else:
    st.info("Sube un Excel o ingresa datos manualmente para habilitar la generación.")
