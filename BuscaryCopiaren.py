import pandas as pd
import os
import shutil
import glob

# Ruta del archivo Excel
#primera columna nombre de nueva carpeta de destino
#segunda columna nombre del archivo, no tiene que estar completo
excel_path = r"C:\Users\CIGSA\Documents\GitHub\Automatizacion-Excel-Word\Libro1.xlsx"

# Carpeta donde están los archivos que quieres copiar
carpeta_origen = r"C:\Users\CIGSA\Downloads\Planos Anexo 2 A"

# Carpeta base donde se copiarán los archivos
carpeta_destino_base = r"C:\Users\CIGSA\Desktop\Test"
print(f"📁 Carpeta destino base: {carpeta_destino_base}")

# Cargar el Excel
df = pd.read_excel(excel_path)

# Verifica los nombres de las columnas
print("📊 Columnas encontradas en Excel:", df.columns)

# Recorre cada fila
for index, row in df.iterrows():
    nombre_carpeta_objetivo = str(row["Carpeta"]).strip()  # Ajusta según el nombre real de la columna
    nombre_archivo = str(row["Archivo"]).strip()           # Ajusta según el nombre real de la columna

    ruta_destino = os.path.join(carpeta_destino_base, nombre_carpeta_objetivo)

    # Crear carpeta destino si no existe
    os.makedirs(ruta_destino, exist_ok=True)

    # Buscar archivos en subcarpetas que contengan el nombre base (cualquier extensión)
    patron_busqueda = os.path.join(carpeta_origen, '**', f'*{nombre_archivo}*')
    coincidencias = glob.glob(patron_busqueda, recursive=True)

    if coincidencias:
        for archivo_encontrado in coincidencias:
            try:
                shutil.copy(archivo_encontrado, ruta_destino)
                print(f"✅ Copiado: {os.path.basename(archivo_encontrado)} → {ruta_destino}")
            except Exception as e:
                print(f"⚠️ Error al copiar {archivo_encontrado}: {e}")
    else:
        print(f"❌ No se encontró archivo que contenga: {nombre_archivo}")