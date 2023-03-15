import os
from tkinter.filedialog import askdirectory

directorio = ""

# si el directorio es vacío se solicita al usuario que seleccione un directorio
while directorio == "":
    # Selección de directorio
    directorio = askdirectory()

# listar todos los archivos CSV del directorio
archivos = os.listdir(directorio)

# Filtra los archivos CSV
archivos = [archivo for archivo in archivos if archivo.endswith(".csv")]

# Eliminar las lineas vacías de los archivos CSV y la líne que dice '<!-- esto es porque tarda mucho en armar el excel entonces cae x time out el sql tarda aprox 13 segundo en un año-->'
for archivo in archivos:
    # Abrir el archivo
    with open(os.path.join(directorio, archivo), "r") as f:
        # Leer el archivo
        contenido = f.readlines()

    # Eliminar las lineas vacías
    contenido = [linea for linea in contenido if linea.strip() != ""]

    # Si la primer línea no comienza con id, eliminar la línea
    if not contenido[0].startswith("id"):
        contenido = contenido[1:]

    # Guardar el archivo
    with open(os.path.join(directorio, archivo), "w") as f:
        # Escribir el archivo
        f.writelines(contenido)



