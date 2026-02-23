import subprocess
import time
import json
import sys
import os

subprocess.run("cls", shell=True)

def printTitle():
    print(fr"{AMARILLO} ___  ___ _____ {RESET}")
    print(fr"{AMARILLO}| _ )/ _ \_   _|{RESET}")
    print(fr"{AMARILLO}| _ \ (_) || |  {RESET}")
    print(fr"{AMARILLO}|___/\___/ |_|  {RESET}")                                                                                                    
    print(fr"{AMARILLO}   _  _   _ _____ ___  __  __   _ _____ ___ ____  _   ___ ___ ___  _  _ {RESET}")
    print(fr"{AMARILLO}  /_\| | | |_   _/ _ \|  \/  | /_\_   _|_ _|_  / /_\ / __|_ _/ _ \| \| |{RESET}")
    print(fr"{AMARILLO} / _ \ |_| | | || (_) | |\/| |/ _ \| |  | | / / / _ \ (__ | | (_) | .` |{RESET}")
    print(fr"{AMARILLO}/_/ \_\___/  |_| \___/|_|  |_/_/ \_\_| |___/___/_/ \_\___|___\___/|_|\_|{RESET}")
    print(fr"{AMARILLO} ___ ___ ___ {RESET}")
    print(fr"{AMARILLO}| _ \ _ )_ _|{RESET}")
    print(fr"{AMARILLO}|  _/ _ \| | {RESET}")
    print(fr"{AMARILLO}|_| |___/___|{RESET}")
    print("\n"*2)

def printHelp():
    """Imprime el menú de ayuda del script"""
    print("Uso del Orquestador de Power BI")
    print("Comando base: py main.py [opciones] [archivo.pbix]\n")
    
    print("Opciones permitidas:")
    print("  onlypublish    Omite la fase de actualización de datos y va directo a la publicación.")
    print("  nodownload     Omite la fase de descarga y descompresión de archivos en SharePoint.")
    print("  --help, -h     Muestra este menú de ayuda y sale del programa.\n")
    
    print("Parámetros dinámicos:")
    print("  [nombre].pbix  Ejecuta el flujo EXCLUSIVAMENTE para el reporte indicado.")
    print("                 (El archivo debe existir en la carpeta definida en DOWNLOAD_PATH).\n")
    
    print("Ejemplos de uso válido:")
    print("  py main.py")
    print("  py main.py nodownload")
    print("  py main.py onlypublish Reporte_Ventas.pbix")
    print("  py main.py nodownload onlypublish Finanzas.pbix\n")

# Definición de colores
VERDE = '\033[92m'
ROJO = '\033[91m'
AMARILLO = '\033[33m'
RESET = '\033[0m'

try:
    with open("config.json", "r", encoding="utf-8") as f:
        config_data = json.load(f)
        download_path = config_data.get("DOWNLOAD_PATH", "")
except Exception:
    download_path = ""

onlyPublish = False
noDownload = False
targetFile = None

# Capturamos todos los argumentos, saltando el nombre del script (índice 0) y pasando a minúsculas
argumentos = sys.argv[1:]

# 1. Búsqueda de petición de ayuda explícita
if "--help" in argumentos or "-h" in argumentos:
    printTitle()
    printHelp()
    sys.exit(0)

# 2. Validación de cantidad de parámetros (Máximo 2)
if len(argumentos) > 3:
    print(f"- {ROJO}Error: Se ingresó una cantidad inválida de parámetros.{RESET}\n")
    printHelp()
    sys.exit(1)

# 3. Asignación de banderas y detección de parámetros no permitidos
for arg in argumentos:
    arg_lower = arg.lower()
    if arg_lower == "onlypublish":
        onlyPublish = True
    elif arg_lower == "nodownload":
        noDownload = True
    elif arg_lower.endswith(".pbix"):
        targetFile = arg
    else:
        printTitle()
        print(f"- {ROJO}Error: El parámetro '{arg}' no se reconoce.{RESET}\n")
        printHelp()
        sys.exit(1)

# Validación de existencia del archivo si se proporcionó uno
if targetFile:
    if not download_path:
        print(f"- {ROJO}Error: DOWNLOAD_PATH no está definido en config.json.{RESET}")
        sys.exit(1)
        
    full_path = os.path.join(download_path, targetFile)
    if not os.path.isfile(full_path):
        printTitle()
        print(f"- {ROJO}Error: El archivo '{targetFile}' no existe en la ruta de descargas.{RESET}")
        print(f"  Ruta revisada: {full_path}\n")
        sys.exit(1)
    
    print(f"{AMARILLO}Modo Archivo Único Activado: Procesando exclusivamente '{targetFile}'{RESET}")

printTitle()

if not noDownload:
    print("* Descarga y descompresión...EJECUTANDO\n")
    try:
        subprocess.check_call(["py", "automateDownloads.py"])
    except subprocess.CalledProcessError:
        sys.exit(1)

    count = 5
    for _ in range(5):
        subprocess.run("cls", shell=True)
        printTitle()
        print(f"\nProceso de descarga y descompresión finalizado, pasando a la siguiente fase en {count}")
        count -= 1
        time.sleep(1)

    subprocess.run("cls", shell=True)
    printTitle()
    print(f"* {VERDE}Descarga y descompresión...LISTO{RESET}")
else:
    print(f"* {VERDE}Descarga y descompresión...OMITIDO{RESET}")

if onlyPublish:
    print("* Solo publicación de reportes...EJECUTANDO\n")
    try:
        cmd_pbi = ["py", "automatePBI.py", "onlyPublish"]
        if targetFile:
            cmd_pbi.append(targetFile)
            subprocess.run(["py", "configManager.py", "add-report", f"{targetFile}"])
        subprocess.check_call(cmd_pbi)
    except subprocess.CalledProcessError:
        sys.exit(1)
    if targetFile:
        subprocess.run(["py", "configManager.py", "del-report", f"{targetFile}"])
    count = 5
    for _ in range(5):
        subprocess.run("cls", shell=True)
        printTitle()
        print(f"\nProceso de publicación finalizado, pasando a la siguiente fase en {count}")
        count -= 1
        time.sleep(1)

    subprocess.run("cls", shell=True)
    printTitle()
    print(f"* {VERDE}Descarga y descompresión...LISTO{RESET}")
    print(f"* {VERDE}Solo publicación de reportes...LISTO{RESET}")
    
    sys.exit(0)

print("* Actualización y publicación de reportes...EJECUTANDO\n")
try:
    cmd_pbi = ["py", "automatePBI.py"]
    if targetFile:
        cmd_pbi.append(targetFile)
    subprocess.check_call(cmd_pbi)
except subprocess.CalledProcessError:
    sys.exit(1)
count = 5
for _ in range(5):
    subprocess.run("cls", shell=True)
    printTitle()
    print(f"\nProceso de actualización y publicación finalizado, pasando a la siguiente fase en {count}")
    count -= 1
    time.sleep(1)

subprocess.run("cls", shell=True)
printTitle()
print(f"* {VERDE}Descarga y descompresión...LISTO{RESET}")
print(f"* {VERDE}Actualización y publicación de reportes...LISTO{RESET}")