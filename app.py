import pandas as pd
import qrcode
import os
import re
from PIL import Image  # Asegúrate de tener Pillow instalado

# === Configuración ===
archivo_excel = r"C:\Users\PC\Documents\Generación de qrs\codigos.xlsx"
col_codigo = "codigo"
col_nombre = "nombre"
carpeta_salida = "QRS_GENERADOS"
tamaño_qr_px = 1920  # Tamaño deseado del QR (ej. 500x500 px)

# Token fijo
TOKEN_JWT = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwOi8vMTkyLjE2OC4xMC4zNzo4MDAwL2FwaS9sb2dpbiIsImlhdCI6MTcxOTUxNjAyMywiZXhwIjoxNzE5NTE5NjIzLCJuYmYiOjE3MTk1MTYwMjMsImp0aSI6IjhURWlHZ2ZQaTFOQkx5UjgiLCJzdWIiOi"

# Función para limpiar nombres de archivo
def limpiar_nombre(nombre):
    nombre = re.sub(r'[\/:*?"<>|]', '', nombre)
    return nombre[:50]

# Crear carpeta si no existe
os.makedirs(carpeta_salida, exist_ok=True)

# Leer Excel
df = pd.read_excel(archivo_excel)

# Generar QRs
for _, fila in df.iterrows():
    codigo = str(fila[col_codigo]).strip()
    nombre_limpio = limpiar_nombre(str(fila[col_nombre]).strip())
    
    url = f"https://comercio.munipachacamac.gob.pe/{codigo}/{TOKEN_JWT}"

    # Crear QR con configuración personalizada
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=4
    )
    qr.add_data(url)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white").convert('RGB')

    img = img.resize((tamaño_qr_px, tamaño_qr_px), Image.LANCZOS)

    nombre_archivo = f"{codigo} {nombre_limpio}.png"
    ruta_archivo = os.path.join(carpeta_salida, nombre_archivo)
    img.save(ruta_archivo)

    print(f"✅ QR generado: {nombre_archivo}")

