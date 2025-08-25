

import os

# Ruta personalizada donde se crearán las carpetas
ruta_destino = "D:\Planos Anexo 2\Modulo 1"

# Crear la ruta si no existe
os.makedirs(ruta_destino, exist_ok=True)

# Crear carpetas numeradas del 1 al 23 dentro de la ruta personalizada
for i in range(1, 24):
    folder_path = os.path.join(ruta_destino, str(i))
    os.makedirs(folder_path, exist_ok=True)

print(f"Se han creado las carpetas del 1 al 23 en la ruta: {ruta_destino}")
