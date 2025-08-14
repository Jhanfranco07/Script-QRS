import pandas as pd
import qrcode
import qrcode.image.svg
import os
import re

# === Configuración ===
archivo_excel = r"C:\Users\PC\Documents\Generación de qrs\codigos.xlsx"
col_codigo = "codigo"
col_nombre = "nombre"
carpeta_salida = "QRS_SVG"  # nueva carpeta para SVGs

# Token JWT fijo
TOKEN_JWT = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwOi8vMTkyLjE2OC4xMC4zNzo4MDAwL2FwaS9sb2dpbiIsImlhdCI6MTcxOTUxNjAyMywiZXhwIjoxNzE5NTE5NjIzLCJuYmYiOjE3MTk1MTYwMjMsImp0aSI6IjhURWlHZ2ZQaTFOQkx5UjgiLCJzdWIiOi"

# Limpiar nombres de archivo
def limpiar_nombre(nombre):
    nombre = re.sub(r'[\/:*?"<>|]', '', nombre)
    return nombre[:50]

# Crear carpeta si no existe
os.makedirs(carpeta_salida, exist_ok=True)

# Leer Excel
df = pd.read_excel(archivo_excel)

# Generar QRs en SVG
for _, fila in df.iterrows():
    codigo = str(fila[col_codigo]).strip()
    nombre_limpio = limpiar_nombre(str(fila[col_nombre]).strip())

    # Construir URL
    url = f"https://comercio.munipachacamac.gob.pe/{codigo}/{TOKEN_JWT}"

    # Crear QR en formato SVG
    factory = qrcode.image.svg.SvgImage
    img = qrcode.make(url, image_factory=factory)

    # Guardar como archivo .svg
    nombre_archivo = f"{codigo} {nombre_limpio}.svg"
    ruta_archivo = os.path.join(carpeta_salida, nombre_archivo)
    img.save(ruta_archivo)

    print(f"✅ QR SVG generado: {nombre_archivo}")
