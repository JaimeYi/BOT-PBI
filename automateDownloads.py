from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium import webdriver
import json
import time
import sys
import os

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

# Cargamos las configuraciones
CONFIG = loadConfig()

# Definición de colores
ROJO = '\033[91m'
RESET = '\033[0m'

#=============================================#
#          CONFIGURACION DEL DRIVER           #
#=============================================#

# Seleccionamos sesion creada para el bot (en caso de no existir se crea una sesion nueva)
bot_user_data = os.path.join(os.getcwd(), 'chrome_bot_profile')
download_dir = CONFIG["DOWNLOAD_PATH"]

chrome_options = Options()

# Se establecen preferencias para no depender de interacciones humanas al descargar
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True
}

# Mas configuraciones
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_experimental_option("detach", True)
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--profile-directory=Default")
chrome_options.add_argument(f"--user-data-dir={bot_user_data}")

# Inicializacion del driver
driver = webdriver.Chrome(options=chrome_options)


#=============================================#
#             PROCESO DE DESCARGA             #
#=============================================#

try:
    # Apertura de la pagina
    driver.get(CONFIG['SP_URL'])
    wait = WebDriverWait(driver, 25)
    print("-> Sesión de Google Chrome iniciada, ingresando en SharePoint")
    
    # Localizar y presionar el boton 'More' (...)
    moreOptions_selector = 'button[data-automationid="more"]'
    moreOptions = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, moreOptions_selector)))
    driver.execute_script("arguments[0].click();", moreOptions)

    # Localizar y presionar el boton de Descargar
    download_selector = 'button[data-automationid="downloadCommand"]'
    download = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, download_selector)))
    download.click()
    print("-> Proceso de descarga iniciado")
    
    # Damos un margen para que el navegador procese el inicio del archivo
    time.sleep(5)

    # Ejecutamos codigo para monitorizar y descomprimir descarga
    os.system("py automateUnzip.py")

    # Cerramos el navegador
    driver.close()


except TimeoutException:
    print(f"\n{ROJO}ERROR: SharePoint tardó demasiado en responder.{RESET}")
    print(f"{ROJO}Posible causa: Botón 'More' o 'Download' no aparecieron. Revisa la conexión.{RESET}")
    driver.quit()
    sys.exit(1)

except NoSuchElementException as e:
    print(f"\n{ROJO}ERROR: No se encontró un elemento en la página.{RESET}")
    print(f"{ROJO}Detalle: {e.msg}{RESET}")
    driver.quit()
    sys.exit(1)

except WebDriverException as e:
    # Aquí capturamos el error de "Symbols not available"
    print(f"\n{ROJO}ERROR DE NAVEGADOR: El driver de Chrome colapsó.{RESET}")
    print(f"{ROJO}Causa probable: Versión de Chrome incompatible o conflicto con el perfil de usuario.{RESET}")
    driver.quit()
    sys.exit(1)

except Exception as e:
    # El "atrapa-todo" de seguridad por si es algo no relacionado a Selenium
    print(f"\n{ROJO}ERROR INESPERADO: {e}{RESET}")
    driver.quit()
    sys.exit(1)

