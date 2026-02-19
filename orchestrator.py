import subprocess
import time
import sys
import os

os.system("cls")

def printTitle():
    print(fr"{AMARILLO} ___  ___ _____ {RESET}")
    print(fr"{AMARILLO}| _ )/ _ \_   _|{RESET}")
    print(fr"{AMARILLO}| _ \ (_) || |  {RESET}")
    print(fr"{AMARILLO}|___/\___/ |_|  {RESET}")                                                                                                    
    print(fr"{AMARILLO}   _  _   _ _____ ___  __  __   _ _____ ___ ____  _   ___ ___ ___  _  _ {RESET}")
    print(fr"{AMARILLO}  /_\| | | |_   _/ _ \|  \/  | /_\_   _|_ _|_  / /_\ / __|_ _/ _ \| \| |{RESET}")
    print(fr"{AMARILLO} / _ \ |_| | | || (_) | |\/| |/ _ \| |  | | / / / _ \ (__ | | (_) | .` |{RESET}")
    print(fr"{AMARILLO}/_/ \_\___/  |_| \___/|_|  |_/_/ \_\_| |___/___/_/ \_\___|___\___/|_|\_|{RESET}")
    print(fr"{AMARILLO} ___ ___ ___{RESET} ")
    print(fr"{AMARILLO}| _ \ _ )_ _{RESET}|")
    print(fr"{AMARILLO}|  _/ _ \| |{RESET} ")
    print(fr"{AMARILLO}|_| |___/___{RESET}|")
    print("\n"*2)

# Definición de colores
VERDE = '\033[92m'
ROJO = '\033[91m'
AMARILLO = '\033[93m'
RESET = '\033[0m'

printTitle()
print("* Descarga y descompresión...EJECUTANDO\n")
try:
    subprocess.check_call(["py", "automateDownloads.py"])
except subprocess.CalledProcessError:
    sys.exit(1)

count = 5
for _ in range(5):
    os.system("cls")
    printTitle()
    print(f"\nProceso de descarga y descompresión finalizado, pasando a la siguiente fase en {count}")
    count -= 1
    time.sleep(1)

os.system("cls")
printTitle()
print(f"* {VERDE}Descarga y descompresión...LISTO{RESET}")
print("* Actualización y publicación de reportes...EJECUTANDO\n")
try:
    subprocess.check_call(["py", "automatePBI.py"])
except subprocess.CalledProcessError:
    sys.exit(1)

count = 5
for _ in range(5):
    os.system("cls")
    printTitle()
    print(f"\nProceso de actualización y descompresión finalizado, pasando a la siguiente fase en {count}")
    count -= 1
    time.sleep(1)

os.system("cls")
printTitle()
print(f"* {VERDE}Descarga y descompresión...LISTO{RESET}")
print(f"* {VERDE}Actualización y publicación de reportes...LISTO{RESET}")