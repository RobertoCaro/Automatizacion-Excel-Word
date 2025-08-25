import pandas as pd
import os
import shutil
import glob

# Ruta del archivo Excel
excel_path = r"C:\Users\Roberto Caro\Documents\GitHub\Automatizacion Excel Word\Libro1.xlsx"

# Carpeta donde están los archivos que quieres copiar
carpeta_origen = r"C:\Users\Roberto Caro\Downloads"

# Carpeta base donde se copiarán los archivos
carpeta_destino_base = r"D:\Planos Anexo 2\Modulo 1"

# Cargar el Excel
df = pd.read_excel(excel_path)

# Verifica los nombres de las columnas
print("Columnas encontradas:", df.columns)
print("Archivos en carpeta origen:")
print(os.listdir(carpeta_origen))

# Recorre cada fila
for index, row in df.iterrows():
    nombre_carpeta_objetivo = str(row["Carpeta"]).strip()  # Ajusta según el nombre real de la columna
    nombre_archivo = str(row["Archivo"]).strip()           # Ajusta según el nombre real de la columna

    ruta_destino = os.path.join(carpeta_destino_base, nombre_carpeta_objetivo)

    # Crear carpeta destino si no existe
    os.makedirs(ruta_destino, exist_ok=True)

    # Buscar archivos que contengan el nombre base
    patron_busqueda = os.path.join(carpeta_origen, f"*{nombre_archivo}*")
    coincidencias = glob.glob(patron_busqueda)

    if coincidencias:
        for archivo_encontrado in coincidencias:
            shutil.copy(archivo_encontrado, ruta_destino)
            print(f"✅ Copiado: {os.path.basename(archivo_encontrado)} → {ruta_destino}")
    else:
        print(f"❌ No se encontró archivo que contenga: {nombre_archivo}")
