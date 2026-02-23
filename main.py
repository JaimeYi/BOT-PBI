import subprocess
import time
import sys

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
    print("Comando base: py main.py <opciones> <opciones>\n")
    print("Opciones permitidas:")
    print("  onlypublish    Omite la fase de actualización de datos y va directo a la publicación.")
    print("  nodownload     Omite la fase de descarga y descompresión de archivos en SharePoint.")
    print("  --help, -h     Muestra este menú de ayuda y sale del programa.\n")
    print("Ejemplos de uso válido:")
    print("  py main.py")
    print("  py main.py onlypublish")
    print("  py main.py nodownload onlypublish\n")

# Definición de colores
VERDE = '\033[92m'
ROJO = '\033[91m'
AMARILLO = '\033[33m'
RESET = '\033[0m'

onlyPublish = False
noDownload = False

# Capturamos todos los argumentos, saltando el nombre del script (índice 0) y pasando a minúsculas
argumentos = [arg.lower() for arg in sys.argv[1:]]

# 1. Búsqueda de petición de ayuda explícita
if "--help" in argumentos or "-h" in argumentos:
    printTitle()
    printHelp()
    sys.exit(0)

# 2. Validación de cantidad de parámetros (Máximo 2)
if len(argumentos) > 2:
    print(f"- {ROJO}Error: Se ingresó una cantidad inválida de parámetros.{RESET}\n")
    printHelp()
    sys.exit(1)

# 3. Asignación de banderas y detección de parámetros no permitidos
for arg in argumentos:
    if arg == "onlypublish":
        onlyPublish = True
    elif arg == "nodownload":
        noDownload = True
    else:
        printTitle()
        print(f"- {ROJO}Error: El parámetro '{arg}' no se reconoce.{RESET}\n")
        printHelp()
        sys.exit(1)

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
        subprocess.check_call(["py", "automatePBI.py", "onlyPublish"])
    except subprocess.CalledProcessError:
        sys.exit(1)
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
    subprocess.check_call(["py", "automatePBI.py"])
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