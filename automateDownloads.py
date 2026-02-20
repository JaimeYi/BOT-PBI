from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium import webdriver
import subprocess
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
AMARILLO = '\033[93m'
RESET = '\033[0m'

def login_microsoft(driver, wait, email, password):
    print("-> Verificando estado de sesión en Microsoft...")
    try:
        campos_visibles = WebDriverWait(driver, 10).until(
            EC.visibility_of_any_elements_located((By.XPATH, "//input[@name='loginfmt' or @name='passwd']"))
        )
        
        # Tomamos el primer elemento visible de la lista
        campo_activo = campos_visibles[0]
        nombre_campo = campo_activo.get_attribute("name")
        
        # 2. Lógica dinámica basada en la UI visible
        if nombre_campo == "loginfmt":
            print("-> Se requiere correo. Ingresando correo...")
            campo_activo.clear()
            campo_activo.send_keys(email)
            
            btn_next = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            btn_next.click()
            
            # Tras avanzar, esperamos a que la contraseña se vuelva visible
            campo_activo = wait.until(EC.visibility_of_element_located((By.NAME, "passwd")))
            time.sleep(1) # Pausa por animación CSS
            
        elif nombre_campo == "passwd":
            print("-> Ingresando contraseña...")
            time.sleep(1)
            
        # 3. Flujo unificado para contraseña
        campo_activo.send_keys(password)
        
        btn_sign_in = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        btn_sign_in.click()

        # 4. Captura de errores de contraseña
        try:
            error_element = WebDriverWait(driver, 3).until(
                EC.visibility_of_element_located((By.ID, "passwordError"))
            )
            print(f"-> {ROJO}ERROR DE CREDENCIALES: {error_element.text}{RESET}")
            print(f"{ROJO}Actualiza la contraseña de SharePoint en config.json.{RESET}")
            driver.quit()
            sys.exit(1)
        except TimeoutException:
            pass # Sin error visible, continuamos

        # 5. Pantalla final de sesión
        btn_no = wait.until(EC.element_to_be_clickable((By.ID, "idBtn_Back")))
        btn_no.click()
        
        print("-> Inicio de sesión exitoso")
        
    except TimeoutException:
        pass

#=============================================#
#          CONFIGURACION DEL DRIVER           #
#=============================================#

# Seleccionamos sesion creada para el bot (en caso de no existir se crea una sesion nueva)
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

# Inicializacion del driver
driver = webdriver.Chrome(options=chrome_options)


#=============================================#
#             PROCESO DE DESCARGA             #
#=============================================#

try:
    # Definimos una lista con las URLs a procesar
    urls_sharepoint = [CONFIG.get('SP_URL'), CONFIG.get('SP_URL_PARAMS')]
    wait = WebDriverWait(driver, 25)

    sp_email = CONFIG["CREDENTIALS"]["SHAREPOINT"]["EMAIL"]
    sp_password = CONFIG["CREDENTIALS"]["SHAREPOINT"]["PASSWORD"]

    # Iteramos sobre cada URL
    for url in urls_sharepoint:
        # Validación de seguridad: si alguna URL está vacía en el config, la saltamos
        if not url:
            print("-> URL no definida en config.json, pasando a la siguiente")
            continue

        driver.get(url)
        time.sleep(5)
        login_microsoft(driver, wait, sp_email, sp_password)

        time.sleep(3)
        
        # Localizar y presionar el boton 'More' (...)
        moreOptions_selector = 'button[data-automationid="more"]'
        moreOptions = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, moreOptions_selector)))
        driver.execute_script("arguments[0].click();", moreOptions)

        # Localizar y presionar el boton de Descargar
        download_selector = 'button[data-automationid="downloadCommand"]'
        download = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, download_selector)))
        download.click()
        if url == CONFIG.get('SP_URL'):
            print("-> Proceso de descarga de reportes y archivos de datos iniciado...")
        else:
            print("-> Proceso de descarga de parámetros iniciado...")
        
        # Damos un margen para que el navegador procese el inicio del archivo
        time.sleep(5)

        # Ejecutamos codigo para monitorizar y descomprimir ESTA descarga
        subprocess.check_call(["py", "automateUnzip.py"])

    print("-> Todas las descargas finalizadas. Cerrando navegador.")
    driver.quit()


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

