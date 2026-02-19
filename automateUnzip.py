# Librerias estandar
from pathlib import Path
import zipfile
import shutil
import time
import json
import sys
import os

COUNTFILES = 0
COUNTREPORTS = 0

# Funcion para poder leer archivo .json que contiene las distintas configuraciones a utilizar
def loadConfig():
    ruta_config = "config.json"
    
    try:
        with open(ruta_config, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config
    except FileNotFoundError:
        print("Error: No se encontró el archivo config.json")
        return {}

# Convierte una ruta al formato extendido de Windows para evitar el límite de 260 caracteres.
def obtener_ruta_extendida(ruta):
    return "\\\\?\\" + str(Path(ruta).absolute())

# Funcion encargada de descomprimir el .zip descargado y dejar los archivos .pbix en el directorio indicado
def unzipAndMove(download_dir):
    global COUNTFILES, COUNTREPORTS

    # Definicion de ruta base
    ruta_base = Path(download_dir)
    ruta_base_ext = obtener_ruta_extendida(download_dir)
    
    # Monitorizacion de la descarga del .zip, se establece un timeout de 600 segundos (10min) para que se descargue el archivo .zip
    archivo_zip = None
    timeout = 1200
    print("-> Esperando a que la carpeta comprimida termine la descarga")
    for _ in range(0, timeout, 2):
        descargas = list(ruta_base.glob("OneDrive_*.zip"))
        temporales = list(ruta_base.glob("*.crdownload"))
        
        if descargas and not temporales:
            archivo_zip = max(descargas, key=os.path.getmtime)
            break
        time.sleep(2)

    if not archivo_zip:
        return False

    # Una vez detectado el termino de la descarga del .zip se procede con la descompresion y reubicacion de los archivos
    try:
        print("-> Descarga finalizada, comenzando descompresión")
        with zipfile.ZipFile(archivo_zip, 'r') as z:
            # Listamos todos los archivos dentro del ZIP
            for info in z.infolist():
                # Verificamos si es un archivo .pbix
                extensions = ('.xlsx', '.xlsm', '.xls', '.xlsb', '.xltx', '.xltm', '.csv', '.pbix')
                for extension in extensions:
                    if info.filename.lower().endswith(extension):
                        COUNTFILES += 1
                        if extension == '.pbix':
                            COUNTREPORTS += 1
                        # Extraemos solo el nombre del archivo (sin la ruta interna del zip)
                        nombre_archivo = os.path.basename(info.filename)
                        if not nombre_archivo: # Saltamos si es una carpeta
                            continue
                            
                        ruta_destino = os.path.join(ruta_base_ext, nombre_archivo)
                        
                        # Leemos el contenido y lo escribimos directamente en el directorio indicado
                        with z.open(info) as fuente, open(ruta_destino, 'wb') as destino:
                            shutil.copyfileobj(fuente, destino)
                        
                        break

        # Archivo .zip original se elimina
        archivo_zip.unlink() 

        print(f"-> Descompresión finalizada con exito, {COUNTFILES} archivos extraídos")
        print(f"----> {COUNTREPORTS} reportes")
        print(f"----> {COUNTFILES-COUNTREPORTS} fuentes de datos (archivos xlsx, csv, etc)")

        time.sleep(5)

        return True

    except Exception as e:
        print(f"Error crítico: {e}")
        return False

if __name__ == "__main__":
    CONFIG = loadConfig()
    if not CONFIG:
        sys.exit(1)
        
    download_path = CONFIG.get("DOWNLOAD_PATH")
    if not download_path:
        print("❌ Error: DOWNLOAD_PATH no definido")
        sys.exit(1)

    resultado = unzipAndMove(download_path)
    
    if resultado:
        sys.exit(0)
    else:
        sys.exit(1)