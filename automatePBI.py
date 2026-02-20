from email.message import EmailMessage
from pywinauto import Application
from datetime import datetime 
from pathlib import Path
import subprocess
import pyautogui
import pyperclip
import requests
import smtplib
import psutil
import time
import json
import os
import re

DEBUG = True

STATES = {
    "apertura": "",
    "publicacion": "",
    "actualizacion": ""
}

TOTALFILES = 0
WARNINGFILES = 0
FAILEDFILES = 0
SUCCESSFILES = 0
EXITOSOS = []
FALLIDOS = []
WARNING = []

# Definición de colores
VERDE = '\033[92m'
ROJO = '\033[91m'
AMARILLO = '\033[93m'
RESET = '\033[0m'


def force_kill_powerbi():
    # 1. Intento agresivo con taskkill (incluye hijos /T y forzado /F)
    subprocess.run(["taskkill", "/F", "/IM", "PBIDesktop.exe", "/T"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
    # 2. Verificación de muerte real (Wait for death)
    # A veces el kernel de Windows tarda unos segundos en liberar el handle
    timeout = 10
    start_time = time.time()
    
    while True:
        process_exists = False
        for proc in psutil.process_iter(['name']):
            if proc.name() == "PBIDesktop.exe":
                process_exists = True
                break
        
        if not process_exists:
            return True
            
        if time.time() - start_time > timeout:
            return False
            
        time.sleep(1)

# Funcion para poder leer archivo .json que contiene las distintas configuraciones a utilizar
def loadConfig():
    global ROJO, RESET
    ruta_config = "config.json"
    
    try:
        with open(ruta_config, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config
    except FileNotFoundError:
        print(f"{ROJO}Error: No se encontró el archivo config.json{RESET}")
        return {}

# Funcion que escribe los errores de la ejecucion en un archivo .txt
def log(archivo_pbi, fase, detalle):
    global ROJO, RESET
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] ARCHIVO: {archivo_pbi} | FASE: {fase} | MOTIVO: {detalle}\n"

    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir, exist_ok=True)
    
    log_filename = f"errores_publicacion_{time.strftime("%Y-%m-%d")}.txt"
    log_path = os.path.join(log_dir, log_filename)
    
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(log_line)
    except Exception as e:
        print(f"{ROJO}No se pudo escribir en el log de errores: {e} {RESET}")

def defineEmoji(result):
    if result == 0: # No error
        return "✅"
    elif result == 1: # Ejecutandose
        return "⏳"
    elif result == 2: # Error
        return "❌"
    else:
        return "⚠️"

def teamsNotification(archivo, faseDeEjecucion, resultado, error):
    global STATES, WEBHOOK, WORKSPACE, AMARILLO, RESET

    """
    Envía una notificación usando el nuevo sistema de Workflows de Teams.
    """
    if faseDeEjecucion == 0:
        STATES["apertura"] = defineEmoji(resultado)
    elif faseDeEjecucion == 1:
        STATES["actualizacion"] = defineEmoji(resultado)
    else:
        STATES["publicacion"] = defineEmoji(resultado)

    status_general = "Procesando..."
    if STATES["publicacion"] == "✅": status_general = "🚀 Publicación Exitosa"

    if resultado == 2:
        status_general = "⚠️ Error en el flujo"
        payload = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": {
                        "type": "AdaptiveCard",
                        "body": [
                            {
                                "type": "TextBlock",
                                "size": "Medium",
                                "weight": "Bolder",
                                "text": f"{status_general} - {archivo}"
                            },
                            {
                                "type": "FactSet",
                                "facts": [
                                    {"title": "Apertura:", "value": f"{STATES["apertura"]}"},
                                    {"title": "Actualización:", "value": f"{STATES["actualizacion"]}"},
                                    {"title": "Publicación:", "value": f"{STATES["publicacion"]}"},
                                    {"title": "Workspace:","value":f"{WORKSPACE}"},
                                    {"title": "Causa:", "value":f"{error}"}
                                ]
                            }
                        ],
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "version": "1.4"
                    }
                }
            ]
        }

    else:
        if resultado == 3:
            status_general = "⚠️ Publicación realizada con advertencias"
        payload = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": {
                        "type": "AdaptiveCard",
                        "body": [
                            {
                                "type": "TextBlock",
                                "size": "Medium",
                                "weight": "Bolder",
                                "text": f"{status_general} - {archivo}"
                            },
                            {
                                "type": "FactSet",
                                "facts": [
                                    {"title": "Apertura:", "value": f"{STATES["apertura"]}"},
                                    {"title": "Actualización:", "value": f"{STATES["actualizacion"]}"},
                                    {"title": "Publicación:", "value": f"{STATES["publicacion"]}"},
                                    {"title": "Workspace:","value":f"{WORKSPACE}"}
                                ]
                            }
                        ],
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "version": "1.4"
                    }
                }
            ]
        }

    try:
        response = requests.post(
            WEBHOOK, 
            data=json.dumps(payload),
            headers={'Content-Type': 'application/json'}
        )
        if response.status_code != 202:
            print(f"{AMARILLO}Error en Workflow: {response.status_code} - {response.text}{AMARILLO}")
    except Exception as e:
        print(f"{AMARILLO}Error de conexión: {e}{AMARILLO}")

def cleanPath(ruta):
    # Buscamos de derecha a izquierda (rfind) la última posición de ambos separadores
    idx_slash = ruta.rfind('\\')
    idx_puntos = ruta.rfind('...')
    
    # Evaluamos cuál de los dos separadores está más cerca del final
    if idx_slash > idx_puntos:
        # El corte se hace 1 posición después de la barra '\'
        resultado = ruta[idx_slash + 1:]
        
    elif idx_puntos != -1:
        # El corte se hace 3 posiciones después, para saltar los '...'
        resultado = ruta[idx_puntos + 3:]
        
    else:
        # Por seguridad, si no existe ninguno de los dos, devolvemos la ruta intacta
        resultado = ruta
        
    # strip() limpia cualquier espacio "fantasma" que haya quedado al principio o al final
    return resultado.strip()

def loginOffice(main_window):
    # Seleccionamos la opcion 'Cuenta de Microsoft'
    btnMicrosoftAccount = main_window.child_window(title_re="^(Cuenta de Microsoft|Cuenta de organización)", control_type="Text", found_index=0)
    btnMicrosoftAccount.click_input()

    # Presionamos el boton 'Iniciar sesion'
    btnLogin = main_window.child_window(title_re="^(Iniciar sesión)", control_type="Button", found_index=0)
    btnLogin.click_input()

    # Seleccionamos la ventana que se abre para iniciar sesion
    modalOffice = main_window.child_window(title_re=f".*(https://login.microsoftonline.com/).*", control_type="Window", found_index=0)
    modalOffice.set_focus()
    # Buscamos la opcion que corresponda a la cuenta de desarrollo y la clickeamos
    btnDevAccount = modalOffice.child_window(title_re=f".*({CONFIG["CREDENTIALS"]["DEV"]["EMAIL"]}).*", control_type="Button", found_index=0)
    btnDevAccount.click_input()

    time.sleep(3)

    passwordField = modalOffice.child_window(title_re="Escriba la contraseña", control_type="Edit", found_index=0)
    passwordField.click_input()
    # Copiamos la clave desde las configuraciones y la pegamos en el campo de contraseña requerido al iniciar sesion
    pyperclip.copy(CONFIG["CREDENTIALS"]["DEV"]["PASSWORD"])
    pyautogui.hotkey('ctrl', 'v')
    btnLogin = modalOffice.child_window(title_re="^(Iniciar sesión)", control_type="Button", found_index=0)
    btnLogin.click_input()

    # Ya teniendo todo listo, presionamos el boton 'Conectar' para asi terminar el proceso de inicio de sesion
    btnConnect = main_window.child_window(title_re="^(Conectar)", control_type="Button", found_index=0)
    btnConnect.click_input()

    time.sleep(2)

def loginBDD(main_window, database):
    global CONFIG

    # Seleccionamos el campo 'usuario' e introducimos el nombre de usuario guardado
    usernameField = main_window.child_window(title="Nombre de usuario", control_type="Edit", found_index=0)
    usernameField.click_input()
    pyautogui.hotkey('ctrl','a')
    pyperclip.copy(CONFIG["CREDENTIALS"]["DATABASES"][database]["USERNAME"])
    pyautogui.hotkey('ctrl', 'v')
    
    # Seleccionamos el campo 'contraseña' e introducimos la contraseña guardada
    passwordField = main_window.child_window(title="Contraseña", control_type="Edit", found_index=0)
    passwordField.click_input()
    pyautogui.hotkey('ctrl','a')
    pyperclip.copy(CONFIG["CREDENTIALS"]["DATABASES"][database]["PASSWORD"])
    pyautogui.hotkey('ctrl', 'v')
    
    # Presionamos boton 'Conectar'
    btnConnect = main_window.child_window(title_re="(Conectar|Guardar)", control_type="Button", found_index=0)
    btnConnect.click_input()
    time.sleep(2)

def update(main_window, name_file):
    teamsNotification(name_file, 1, 1, '')

    tab_inicio = main_window.child_window(title="Inicio", control_type="TabItem", found_index=0)
    if tab_inicio.exists(timeout=5):
        tab_inicio.click_input()
        time.sleep(1)

    groupBoxQuerys = main_window.child_window(title="Consultas", control_type="Group", found_index=0)

    # Se presiona el botón de 'Actualizar'
    btn_refresh = groupBoxQuerys.child_window(title="Actualizar", control_type="Button", found_index=1)
    btn_refresh.click_input()

    btn_refresh.wait('visible', timeout=10)
    btn_refresh.click_input()

    time.sleep(2)


    # Se selecciona el panel de actualizacion desplegado
    dialogo_progreso = main_window.child_window(title_re=".*(Actualizar).*", control_type="Window", found_index=0)

    
    # Timeout para asegurar detectar el panel de actualizacion
    if dialogo_progreso.exists(timeout=20, retry_interval=2):
        # Bucle para detener el flujo del script hasta que se termine de actualizar el reporte (flujo continua cuando panel de actualizacion desaparece)
        while True:
            if not dialogo_progreso.exists():
                time.sleep(3)
                # Doble verificacion para asegurar que realmente el proceso termino
                if not dialogo_progreso.exists():
                    teamsNotification(name_file, 1, 0, '')
                    return True
            else:
                try:
                    time.sleep(2)
                    # Verificacion de si aparece modal de SharePoint
                    sp_label = main_window.child_window(title="SharePoint", control_type="Text")
                    if sp_label.exists(timeout=0):
                        loginOffice(main_window)
                    
                    # Verificacion de si aparece modal de Salesforce
                    sf_label = main_window.child_window(title="Acceder a Salesforce",control_type="Text")
                    if sf_label.exists(timeout=0):
                        raise ValueError("Se requiere intervención (Salesforce)|actualizacion")
                    
                    # Verificacion de si aparece modal de Base de Datos
                    bdd_label = main_window.child_window(title="Base de datos MySQL",control_type="Text")
                    if bdd_label.exists(timeout=0):
                        btnBDD = main_window.child_window(title_re="^(Base de datos)", control_type="Text", found_index=0)
                        btnBDD.click_input()
                        
                        # Se itera sobre las credenciales almacenadas de bases de datos, si se encuentra la credencial requerida se introducen las
                        # respectivas credenciales, en caso contrario se arroja una excepción
                        found = False
                        for database in CONFIG["CREDENTIALS"]["DATABASES"]:
                            dbIndicator = main_window.child_window(title_re=f".*{database}.*", control_type="Text", found_index=0)
                            if dbIndicator.exists(timeout=1):
                                found = True
                                loginBDD(main_window, database)
                        
                        if not found:
                                raise ValueError("No se encontró la credencial correspondiente a la base de datos requerida|actualizacion")

                    # Verificacion de si aparece modal de Contenido Web
                    www_label = main_window.child_window(title="Acceder a contenido web",control_type="Text")
                    if www_label.exists(timeout=0):
                        loginOffice(main_window)
                        ValueError("Se requiere intervención (Acceso a contenido Web)|actualizacion")

                    time.sleep(1) 

                    # Verificacion si existe algun texto de error
                    if dialogo_progreso.child_window(title_re=".*(error|errores|credenciales|Folder).*", control_type="Text", found_index=0).exists(timeout=0):
                        messages = [el.window_text() for el in dialogo_progreso.descendants(control_type="Text")]

                        # Comprobamos si existe algun error por falta de archivo o por MySQL Host (credenciales invalidas MySql)
                        fileOrFolderFlag = False
                        sqlHostFlag = False
                        for i in messages:
                            if "File or Folder" in i:
                                fileOrFolderFlag = True
                                break
                            if "MySQL" in i:
                                sqlHostFlag = True
                                break
                            

                        # Condicional para establecer la logica necesaria en caso de haber un error por falta de archivo
                        if fileOrFolderFlag:
                            messages = [el for el in messages if "File or Folder" in el]
                            # Establecemos expresion regular para filtrar errores
                            regPattern = r".*(Reporteria C&G Grupo Ruedas).*(03. REPORTES).*"
                            regPattern = re.compile(regPattern, re.IGNORECASE)

                            invalidFiles = 0
                            validFiles = 0
                            namesOfValidFiles = []

                            # Se itera sobre los distintos mensajes que hayan respecto a falta de archivos
                            for elem in messages:
                                # Si la ruta especificada corresponde con la expresion regular el problema sera solucionable, en caso contrario no
                                if re.search(regPattern, elem):
                                    path = elem.split("'")[1]
                                    fileName = Path(path).name.lower()
                                    validFiles += 1
                                    if fileName not in namesOfValidFiles:
                                        namesOfValidFiles.append(fileName)
                                else:
                                    invalidFiles += 1
                                

                            # Condicional para abarcar logica necesaria en caso de haber errores reparables
                            if validFiles > 0:
                                # Cerramos el modal de actualizacion
                                closeButton = dialogo_progreso.child_window(title="Cerrar", control_type="Button", found_index=0)
                                closeButton.click_input()

                                # Presionamos boton desplegable en la seccion 'Transformar datos'
                                transformDataButton = main_window.child_window(title_re=".*(Transformar datos).*", control_type="Button", found_index=0)
                                transformDataButton.click_input()

                                # Presionamos la opcion 'Configuracion de origen de datos'
                                configSourceDataButton = main_window.child_window(title="Configuración de origen de datos", control_type="MenuItem", found_index=0)
                                configSourceDataButton.click_input()

                                # Obtenemos una copia de la lista de origenes de datos locales existentes para no tener errores por referencias obsoletas
                                configSourceDataWindow = main_window.child_window(title_re=".*(Configuración de origen de datos).*", control_type="Window", found_index=0)
                                listSources = main_window.child_window(control_type="List", found_index=1)
                                listSources.set_focus()
                                listSources.type_keys("^{END}")
                                extensions = ('xlsx', 'xlsm', 'xls', 'xlsb', 'xltx', 'xltm', 'csv')
                                dataSourcesList = [t for item in listSources if any(ext in (t := item.window_text()) for ext in extensions)]

                                # Iteramos sobre los distintos archivos que pueden estar generando el problema
                                modifiedPaths = 0
                                for file in dataSourcesList:
                                    time.sleep(3)
                                    # Seleccionamos el elemento dentro de la lista
                                    searchBar = configSourceDataWindow.child_window(title='Buscar configuración de origen de datos', control_type="Edit", found_index=0)
                                    pyperclip.copy(cleanPath(file))
                                    searchBar.click_input()
                                    pyautogui.hotkey('ctrl', 'a')
                                    pyautogui.hotkey('ctrl', 'v')

                                    selectedSource = listSources.child_window(title=file, control_type="ListItem", found_index=0)
                                    selectedSource.click_input()

                                    # Entramos a la seccion 'Cambiar origen'
                                    changeSource = configSourceDataWindow.child_window(title="Cambiar origen...", control_type="Button", found_index=0)
                                    changeSource.click_input()

                                    time.sleep(3)
                                    
                                    # Obtenemos la ruta actual del archivo
                                    changePathOfFileField = configSourceDataWindow.child_window(title_re=".*(Ruta de acceso de archivo).*", control_type="Edit", found_index=0)
                                    changePathOfFileField.click_input()

                                    # Verificamos si el archivo al que apunta la ruta que estamos revisando esta dentro de la lista de archivos permitidos
                                    newPath = "N/A"
                                    currentFilename = Path(changePathOfFileField.get_value()).name.lower()
                                    if currentFilename in namesOfValidFiles:
                                        newPath = os.path.join(CONFIG["DOWNLOAD_PATH"], currentFilename)

                                    # Si el archivo es valido se cambia la ruta para apuntar a la real ruta de este archivo, en caso contrario se deja tal cual estaba
                                    if os.path.exists(newPath):
                                        pyautogui.hotkey('ctrl', 'a')
                                        pyperclip.copy(str(newPath))
                                        pyautogui.hotkey('ctrl','v')
                                        modifiedPaths += 1
                                    else:
                                        invalidFiles += 1

                                    # Presionamos el boton 'Aceptar'
                                    btnAccept = configSourceDataWindow.child_window(title="Aceptar", control_type="Button", found_index=0)
                                    btnAccept.click_input()

                                time.sleep(3)

                                # Cerramos la seccion donde estabamos trabajando
                                btnClose = configSourceDataWindow.child_window(title="Cerrar", control_type="Button", found_index=0)
                                btnClose.click_input()

                                time.sleep(2)
                                # Si se modifico alguna ruta aplicamos los cambios y esperamos a que se carguen los datos
                                if modifiedPaths > 0:
                                    unapplyChangesStatusBar = main_window.child_window(title_re=".*Tiene cambios pendientes en sus consultas que no se han aplicado.*", control_type="StatusBar", found_index=0)
                                    btnApply = unapplyChangesStatusBar.child_window(title_re="(Aplicar).*", control_type="Button", found_index=0)
                                    btnApply.click_input()

                                    time.sleep(3)

                                    loadModal = main_window.child_window(title_re=".*(Carga).*", control_type="Pane")
                                    while True:
                                        if not loadModal.exists(timeout=20, retry_interval=2):
                                            break
                            # Si existen archivos no validos arrojamos una excepcion, en caso contrario guardamos los cambios y retornamos False para repetir el proceso de actualizacion
                            if invalidFiles > 0:
                                raise ValueError("no se han logrado encontrar los archivos especificados en los origenes de datos")
                            else:
                                btnSave = main_window.child_window(title="Guardar", control_type="Button", found_index=0)
                                btnSave.click_input()
                                time.sleep(20)
                                return False
                        
                        elif sqlHostFlag:
                            # Cerramos el modal de actualizacion
                            closeButton = dialogo_progreso.child_window(title="Cerrar", control_type="Button", found_index=0)
                            closeButton.click_input()

                            # Presionamos boton desplegable en la seccion 'Transformar datos'
                            transformDataButton = main_window.child_window(title_re=".*(Transformar datos).*", control_type="Button", found_index=0)
                            transformDataButton.click_input()

                            # Presionamos la opcion 'Configuracion de origen de datos'
                            configSourceDataButton = main_window.child_window(title="Configuración de origen de datos", control_type="MenuItem", found_index=0)
                            configSourceDataButton.click_input()

                            configSourceDataWindow = main_window.child_window(title_re=".*(Configuración de origen de datos).*", control_type="Window", found_index=0)
                            listSources = main_window.child_window(control_type="List", found_index=1)
                            listSources.set_focus()
                            listSources.type_keys("^{END}")
                            extensions = ('xlsx', 'xlsm', 'xls', 'xlsb', 'xltx', 'xltm', 'csv')
                            dataSourcesList = [t for item in listSources if not any(ext in (t := item.window_text()) for ext in extensions)]
                            
                            changes = False
                            for source in dataSourcesList:
                                searchBar = configSourceDataWindow.child_window(title='Buscar configuración de origen de datos', control_type="Edit", found_index=0)
                                searchBar.click_input()
                                pyautogui.hotkey('ctrl', 'a')
                                searchBar.type_keys(source)

                                for database in CONFIG["CREDENTIALS"]["DATABASES"]:
                                    if database in source:
                                        changes = True
                                        selectedSource = listSources.child_window(title=source, control_type="ListItem", found_index=0)
                                        selectedSource.click_input()

                                        configSourceDataWindow.child_window(title="Borrar permisos", control_type="Button", found_index=0).click_input()
                                        configSourceDataWindow.child_window(title="Eliminar", control_type="Button", found_index=0).click_input()

                                        break
                            
                            btnClose = configSourceDataWindow.child_window(title="Cerrar", control_type="Button", found_index=0)
                            btnClose.click_input()

                            if changes:
                                btnSave = main_window.child_window(title="Guardar", control_type="Button", found_index=0)
                                btnSave.click_input()
                                time.sleep(20)
                                return False
                            else:
                                raise ValueError("no se han logrado encontrar las credenciales para la base de datos problematica")

                        # Solo si existe, traemos el texto para el log
                        error_el = dialogo_progreso.child_window(title_re=".*(error|errores|credenciales|Folder).*", control_type="Text", found_index=0)
                        raise ValueError(f"{error_el.window_text()}|actualizacion")

                except Exception as e:
                    # 1. Si el error es uno de los nuestros (ya trae el marcador), lo lanzamos fuera de inmediato
                    # Esto evitará que caiga en el try/except anidado de abajo.
                    if "|actualizacion" in str(e):
                        raise e

                    try:
                        if not dialogo_progreso.exists(timeout=0) and not fileOrFolderFlag:
                            teamsNotification(name_file, 1, 0, '')
                            return True
                        elif fileOrFolderFlag:
                            teamsNotification(name_file, 1, 2, e)
                            raise ValueError(f"Error de origen de datos: {e}|actualizacion")
                    except Exception as nested_e:
                        if "|actualizacion" in str(nested_e):
                            raise nested_e
                        
                        return True
    else:
        return False

def automateWorkflow(file):
    global SUCCESSFILES, EXITOSOS, FALLIDOS, WARNINGFILES, ROJO, RESET, AMARILLO, VERDE

    name_file = os.path.basename(file).replace(".pbix", "")

    try:
        #=============================================#
        #             APERTURA DE ARCHIVO             #
        #=============================================#
        os.startfile(file)
        
        # Apertura de archivo seleccionado
        app = Application(backend="uia").connect(path="PBIDesktop.exe", timeout=60)
        teamsNotification(name_file, 0, 0, '')
        
        # Guardamos la ventana de power bi en una variable (para poder manipularla)
        main_window = app.window(title_re=f".*{name_file}.*")
        
        # Timeout para esperar a que la ventana de power bi termine de cargar
        main_window.wait('ready', timeout=300)
        main_window.set_focus()

        time.sleep(30)



        #=============================================#
        #           ACTUALIZACION DE ARCHIVO          #
        #=============================================#
        
        print("-> Iniciando proceso de actualización")

        # Cantidad de intentos que lleva la actualizacion
        updateTrys = 0
        
        # Cantidad de intentos permitidos que se le daran a la actualizacion
        attempts = 3

        while updateTrys < attempts:
            update_result = update(main_window, name_file)
            
            if update_result is True:
                break
            else:
                if updateTrys >= 0:
                    print(f"-> {AMARILLO}Hubo un problema al iniciar la actualización, reintentando...{RESET}")
                    updateTrys += 1
                    if not force_kill_powerbi():
                        raise ValueError("No se logró eliminar el proceso anterior de PBI Desktop|actualizacion")
                    time.sleep(5)
                    if updateTrys == attempts:
                        raise ValueError("No se logró presionar botón de actualización o panel de actualización no se desplegó|actualizacion")
                    os.startfile(file)

                    app = Application(backend="uia").connect(path="PBIDesktop.exe", timeout=60)
                    teamsNotification(name_file, 0, 0, '')
                    
                    main_window = app.window(title_re=f".*{name_file}.*")
                    
                    main_window.wait('ready', timeout=300)
                    main_window.set_focus()
            

        #=============================================#
        #           PUBLICACION DE ARCHIVO            #
        #=============================================#
        time.sleep(3)

        print(f"-> {VERDE}Proceso de actualización finalizado de manera exitosa{RESET}")
        print("-> Iniciando proceso de publicación")

        # Presionamos el boton 'Publicar'
        btn_publish = main_window.child_window(title_re="^(Publicar|Publish)", control_type="Button", found_index=0)
        btn_publish.click_input()

        teamsNotification(name_file, 2, 1, '')

        updateLink = "ms-pbi://pbi.microsoft.com/Views/KoForm.htm"
        modal_save = main_window.child_window(title_re=f".*{updateLink}.*", control_type="Pane", found_index=0)
        
        if modal_save.exists(timeout=60):
            # Buscamos el boton 'Guardar' dentro del modal que aparece antes de publicar
            btn_save = modal_save.child_window(title_re="^(Guardar|Save).*", control_type="Button",found_index=0)
            
            # Verificamos si el modal aparecio, en caso de aparecer presionamos el boton 'Guardar'
            if btn_save.exists(timeout=3):
                btn_save.click_input()
                # Esperamos a que el modal desaparezca para procesar el guardado
                modal_save.wait_not('exists', timeout=15)
            else:
                # Si no encuentra el texto 'Guardar', intentamos con la tecla Enter
                modal_save.set_focus()
                pyautogui.press('enter')
        else:
            raise ValueError("No se encontró el panel de confirmación|publicacion")

        # Guardamos en una variable el modal correspondiente al menu de seleccion al WORKSPACE donde se desea subir el reporte
        url_workspace = "ms-pbi.pbi.microsoft.com/minerva/desktopDialogHost.html"
        ventana_workspace = main_window.child_window(title_re=f".*{url_workspace}.*", control_type="Pane", found_index=0)

        # Buscamos y seleccionamos el campo editable dentro del modal de muestra las opciones de WORKSPACEs
        search_field = ventana_workspace.child_window(control_type="Edit", found_index=0)
        search_field.click_input()
        
        # Pegamos el nombre del WORKSPACE
        pyperclip.copy(WORKSPACE)
        pyautogui.hotkey('ctrl', 'v')
        
        # Pausa para que Power BI procese la búsqueda en la nube
        time.sleep(2)

        # Buscamos el workspace seleccionado en la lista desplegada y la seleccionamos
        workspace_item = ventana_workspace.child_window(title=WORKSPACE, control_type="ListItem")
        workspace_item.select()

        # Presionamos el boton 'Seleccionar'
        workspacesLabels = ventana_workspace.descendants(control_type="Text")
        patronConfirmSelection = re.compile(r"(Seleccionar|Select)", re.IGNORECASE)
        foundConfirmSelection = [el for el in workspacesLabels if patronConfirmSelection.match(el.window_text())]

        # Bloque condicional para cubrir caso donde el reporte ya existe en el workspace
        if foundConfirmSelection:
            btnReplace = ventana_workspace.child_window(title_re="^(Seleccionar|Select)", control_type="Button", found_index=0)
            btnReplace.click_input()
        
        # Ciclo para mantener en pausa el programa mientras se publica el informe, además en cada iteración se comprueba 
        # si aparece un modal para sobreescribir el informe si es que este ya habia sido publicado anteriormente. Finalmente
        # cuando termina el proceso de publicacion, se navega hasta el ultimo boton y se cierra el modal
        while True:
            time.sleep(1)
            publishLabels = ventana_workspace.descendants(control_type="Text")

            # Buscamos si existe algun texto que diga 'reemplazar'
            patronReplace = re.compile(r".*(reemplazar|replace).*", re.IGNORECASE)
            foundReplace = [el for el in publishLabels if patronReplace.match(el.window_text())]

            # Buscamos si existe algun texto que diga 'El archivo se publicó, pero se desconectó' (texto de advertencia)
            patronProblema = re.compile(r"El archivo se publicó, pero se desconectó", re.IGNORECASE)
            foundProblema = [el for el in publishLabels if patronProblema.match(el.window_text())]

            # Buscamos si existe algun texto que diga 'completada correctamente.'
            patronReady = re.compile(r".*completada correctamente.*", re.IGNORECASE)
            foundReady = [el for el in publishLabels if patronReady.match(el.window_text())]

            # Si se encuentra un texto de reemplazar buscamos el boton que diga 'reemplazar' y lo presionamos
            if foundReplace:
                btnReplace = ventana_workspace.child_window(title_re="^(Reemplazar|Replace)", control_type="Button", found_index=0)
                btnReplace.click_input()
            
            # Si se encuentra un texto de advertencia notificamos por teams, modificamos las respectivas variables y terminamos el ciclo
            if foundProblema:
                teamsNotification(name_file, 2, 3, 'Problema con los origenes de datos al publicar')
                WARNINGFILES += 1
                WARNING.append(name_file)
                print("-> Proceso de publicación finalizado con advertencias")
                break
            # Si se encuentra un texto de exito notificamos por teams, modificamos las respectivas variables y terminamos el ciclo
            if foundReady:
                btnReady = ventana_workspace.child_window(title_re="^(Entendido)", control_type="Button", found_index=0)
                btnReady.click_input()
                teamsNotification(name_file, 2, 0, '')
                SUCCESSFILES += 1
                EXITOSOS.append(name_file)
                print(f"-> {VERDE}Proceso de publicación finalizado de manera exitosa{RESET}")
                break
    
    except ValueError as e:
        FALLIDOS.append(name_file)
        error_msg = str(e)
        causa = "Desconocida"
        fase = "Proceso"
        
        if '|' in error_msg:
            section = error_msg.split('|')
            fase = section[1]
            causa = section[0]
            
            if fase == 'actualizacion':
                print(f"-> {ROJO}Proceso de actualización fallido{RESET}")
                teamsNotification(name_file, 1, 2, causa)
            elif fase == 'publicacion':
                print(f"-> {ROJO}Proceso de publicación fallido{RESET}")
                teamsNotification(name_file, 2, 2, causa)
        
        # ANOTACIÓN EN TXT
        if 'found_index is specified as' in causa:
            causa = "Elemento no encontrado"
        log(name_file, fase, causa)
        print(f"-> {ROJO}Revisar logs de ejecución para más información{RESET}")

    except Exception as e:
        FALLIDOS.append(name_file)
        teamsNotification(name_file, "General", 2, str(e))
        log(name_file, "-> Error Inesperado", str(e))
        print(f"-> {ROJO}Error crítico registrado, revisar logs de ejecución para más información{RESET}")
    
    force_kill_powerbi()

if __name__ == "__main__":
    CONFIG = loadConfig()

    # Corresponde a la ruta absoluta donde se encuentra el directorio que contiene los reportes PBI a publicar
    ROUTE = CONFIG["DOWNLOAD_PATH"]

    # Corresponde a el nombre especifico del workspace donde seran publicados los reportes
    WORKSPACE = CONFIG["WORKSPACE"]

    # URL donde se envian las peticiones post con el mensaje a entregar por canal de Teams
    WEBHOOK = CONFIG["WEBHOOK_URL"]

    skip = ["PBI_EEFF_GRUPO RUEDAS 2025.pbix",
            "PBI_EEFF_GRUPO RUEDAS 2026.pbix"]

    # Se llama a la funcion principal de la automatizacion en este ciclo que itera sobre todos los archivos 
    # existentes dentro del directorio indicado
    for file in os.listdir(ROUTE):
        #if '.pbix' not in file or file in skip:
        if '.pbix' not in file:
            continue

        TOTALFILES += 1
        STATES["apertura"] = STATES["actualizacion"] = STATES["publicacion"] = "⚪"

        print(f"Trabajando en {file}")
        automateWorkflow(os.path.join(ROUTE,file))



    '''
    RESUMEN DE EJECUCION
    '''
    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "TextBlock",
                            "size": "Medium",
                            "weight": "Bolder",
                            "text": "📊 Resumen de ejecución"
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                {"title": "📁 Archivos totales:", "value": TOTALFILES},
                                {"title": "✅ Sin errores:", "value": SUCCESSFILES},
                                {"title": "⚠️ Con advertencias:", "value": WARNINGFILES},
                                {"title": "❌ Con errores:", "value": TOTALFILES-SUCCESSFILES-WARNINGFILES}
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "size": "Small",
                            "text": "Para más información acerca de los errores, consultar logs de la ejecución."
                        }
                    ],
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.4"
                }
            }
        ]
    }

    try:
        response = requests.post(
            WEBHOOK, 
            data=json.dumps(payload),
            headers={'Content-Type': 'application/json'}
        )
        if response.status_code != 202:
            print(f"{AMARILLO}Error en Workflow: {response.status_code} - {response.text}{AMARILLO}")
    except Exception as e:
        print(f"{AMARILLO}Error de conexión: {e}{AMARILLO}")

    if not DEBUG:
        #=============================================#
        #   ENVIAR CORREO CON RESUMEN DE EJECUCION    #
        #=============================================#

        # Plantilla para el correo a mandar
        html_template = """
            <html>
            <body style="font-family: Arial, sans-serif; color: #333;">
                <h2 style="color: #004a99;">Resumen de Publicación Power BI</h2>
                <p><strong>Fecha de ejecución:</strong> {fecha}</p>
                
                <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
                    <tr style="background-color: #f2f2f2;">
                        <th style="border: 1px solid #ddd; padding: 8px;">Total Procesados</th>
                        <th style="border: 1px solid #ddd; padding: 8px; color: #28a745;">Exitosos</th>
                        <th style="border: 1px solid #ddd; padding: 8px; color: #FFFF00;">Con advertencias</th>
                        <th style="border: 1px solid #ddd; padding: 8px; color: #dc3545;">Con Errores</th>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{total}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{exitosos}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{advertencias}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{fallidos}</td>
                    </tr>
                </table>

                <h3 style="color: #28a745;">✅ Archivos Publicados:</h3>
                <ul>
                    {lista_exitosos}
                </ul>

                <h3 style="color: #FFFF00;">⚠️ Archivos con advertencia:</h3>
                <ul>
                    {lista_advertencias}
                </ul>

                <h3 style="color: #dc3545;">❌ Archivos con Fallos:</h3>
                <ul>
                    {lista_fallidos}
                </ul>

                <hr>
                <p style="font-size: 12px; color: #777;">Este es un correo automático generado por el BOT de Automatización.</p>
            </body>
            </html>
        """

        # Seteamos datos necesarios para el correo (cuerpo, destinatario, etc)
        msg = EmailMessage()
        msg.set_content("Cuerpo del correo")
        msg["Subject"] = f"BOT PBI - Resumen ejecución - {datetime.now().strftime("%d/%m/%Y")}"
        msg["From"] = CONFIG["CREDENTIALS"]["DEV"]["EMAIL"]
        msg["To"] = CONFIG["ADDRESSE_MAIL"]

        # Extraccion de datos para incluir en el correo
        fecha_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        li_exitosos = "".join([f"<li>{file}</li>" for file in EXITOSOS]) if EXITOSOS else "<li>Ninguno</li>"
        li_advertencias = "".join([f"<li>{file}</li>" for file in WARNING]) if WARNING else "<li>Ninguno</li>"
        li_fallidos = "".join([f"<li>{file}</li>" for file in FALLIDOS]) if FALLIDOS else "<li>Ninguno</li>"

        # Rellenamos campos pendientes del formato HTML 
        htmlBody = html_template.format(
            fecha=fecha_str,
            total=len(EXITOSOS) + len(FALLIDOS) + len(WARNING),
            exitosos=len(EXITOSOS),
            advertencias=len(WARNING),
            fallidos=len(FALLIDOS),
            lista_exitosos=li_exitosos,
            lista_advertencias=li_advertencias,
            lista_fallidos=li_fallidos
        )

        msg.set_content("Favor de ver este correo en un cliente compatible con HTML.") # Texto alternativo por si el dispositivo desde donde se lee el correo no es compatible con HTML
        msg.add_alternative(htmlBody, subtype='html')

        # Envio del correo con credenciales de la cuenta de desarrollo
        try:
            mailserver = smtplib.SMTP('smtp.office365.com', 587)
            mailserver.starttls()
            mailserver.login(CONFIG["CREDENTIALS"]["DEV"]["EMAIL"], CONFIG["CREDENTIALS"]["DEV"]["PASSWORD"])
            mailserver.send_message(msg)
            mailserver.quit()
        except Exception as e:
            print(f"{ROJO}Error: {e}{RESET}")