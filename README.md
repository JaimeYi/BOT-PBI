## 1. Resumen del Funcionamiento

Este proyecto consiste en un flujo de automatización robótica de procesos diseñado para la descarga, actualización y publicación desatendida de reportes de Power BI (`.pbix`). El sistema simula interacciones de usuario para extraer fuentes de datos actualizadas desde SharePoint, procesar los reportes localmente gestionando credenciales de origen y modales de error, y finalmente publicar los resultados en un Workspace de Power BI Service. El ciclo concluye con notificaciones de estado enviadas a través de Microsoft Teams y correo electrónico.

---

## 2. Arquitectura y Módulos

El sistema opera bajo una arquitectura modular orientada a la línea de comandos (CLI), dividida en los siguientes componentes principales:

1. **`main.py` (Orquestador Principal):** Punto de entrada del sistema. Parsea los parámetros de ejecución, valida las solicitudes, orquesta la llamada a los subprocesos en el orden correcto y controla el manejo de excepciones de alto nivel.
2. **`configManager.py` (Gestor de Configuración):** Interfaz CLI dedicada a la administración segura del archivo `config.json`. Previene errores de sintaxis y permite la configuración de rutas, credenciales y reportes directamente desde la consola.
3. **`automateDownloads.py`:** Módulo basado en Selenium WebDriver. Automatiza el inicio de sesión en el entorno de Microsoft 365, navega hacia las rutas de SharePoint especificadas y ejecuta la descarga de paquetes de datos y parámetros.
4. **`automateUnzip.py`:** Monitor de sistema de archivos. Detecta la finalización de las descargas y procede a la extracción y reubicación exclusiva de formatos autorizados (Excel, CSV, PBIX) hacia el directorio de trabajo operativo.
5. **`automatePBI.py`:** Módulo núcleo de procesamiento. Utiliza la librería Pywinauto para tomar control de la interfaz de Power BI Desktop, aplicar actualizaciones de datos, inyectar credenciales (Bases de Datos, Web, SharePoint), enrutar conexiones locales y ejecutar la publicación del reporte.

---

## 3. Requisitos del Sistema e Instalación

Para asegurar la correcta ejecución del bot, el entorno anfitrión debe cumplir con las siguientes especificaciones:

### Requisitos de Entorno
* **Sistema Operativo:** Windows 10, Windows 11 o Windows Server (Requisito estricto de compatibilidad para Pywinauto y Power BI Desktop).
* **Software Base:**
  * Python 3.9 o superior.
  * Google Chrome (Última versión estable).
  * Microsoft Power BI Desktop.

### Instalación de Dependencias
Ejecute el siguiente comando en su terminal para instalar las librerías requeridas por el entorno de Python:

```bash
pip install selenium pywinauto pyautogui pyperclip requests psutil

```

---

## 4. Configuración Inicial (configManager.py)

Toda la parametrización del entorno (rutas, accesos, credenciales) se almacena en el archivo `config.json`. Si el archivo no existe, la ejecución de cualquier comando de `configManager.py` generará automáticamente una plantilla estructural. Por motivos de seguridad y estandarización, la manipulación de este archivo debe realizarse **exclusivamente** mediante el gestor.

Para visualizar el listado completo de comandos disponibles:

```cmd
python configManager.py --help

```

### Comandos de Configuración Frecuentes:

**1. Definición de variables globales:**

```cmd
python configManager.py set DOWNLOAD_PATH "C:\Ruta\Hacia\Directorio\Descargas"
python configManager.py set WORKSPACE "Nombre_Workspace_Destino"
python configManager.py set SP_URL "[https://empresa.sharepoint.com/ruta_datos](https://empresa.sharepoint.com/ruta_datos)"
python configManager.py set SP_URL_PARAMS "[https://empresa.sharepoint.com/ruta_parametros](https://empresa.sharepoint.com/ruta_parametros)"

```

**2. Actualización de credenciales de sistema:**

```cmd
python configManager.py set-cred DEV usuario@empresa.com "ContraseñaSegura"
python configManager.py set-cred SHAREPOINT usuario@empresa.com "ContraseñaSegura"

```

**3. Administración de orígenes de datos (Bases de datos y Web):**

```cmd
python configManager.py add-db 192.168.0.1 admin_sql "ClaveSQL"
python configManager.py add-web "[https://intranet.empresa.com](https://intranet.empresa.com)" admin "ClaveWeb"

```

**4. Administración de reportes a publicar de forma exclusiva (modo 'solo publicación'):**

```cmd
python configManager.py add-report "Reporte.pbix"
python configManager.py del-report "Reporte.pbix"

```

**5. Exclusión de reportes específicos (Omisión durante el ciclo):**
Los archivos añadidos mediante este comando serán ignorados por el orquestador general.

```
python configManager.py add-skip-report "Reporte_Obsoleto.pbix"
python configManager.py del-skip-report "Reporte_Obsoleto.pbix"

```

**6. Auditoría de configuración actual:**

```cmd
python configManager.py list

```

*(Nota: Para habilitar las alertas vía Microsoft Teams, es necesario generar una URL mediante un flujo de "Webhook entrante" en el canal deseado y registrarla en el sistema utilizando el comando: `python configManager.py set WEBHOOK_URL "URL_GENERADA"`).*

---

## 5. Ejecución y Modo de Uso (main.py)

La herramienta `main.py` evalúa de forma preliminar que existan las configuraciones mínimas obligatorias para operar. Previo a iniciar cualquier flujo, se debe garantizar que **no existan instancias previas de Power BI Desktop en ejecución** y que el equipo no reciba interacción de hardware (mouse/teclado) durante el proceso automatizado.

Para consultar el manual de operación en consola:

```cmd
python main.py --help

```

### Escenarios de Ejecución Permitidos:

**Ejecución de flujo completo (Estándar):**
Ejecuta secuencialmente descarga, descompresión, actualización de datos y publicación para todos los reportes, omitiendo aquellos detallados en la lista de exclusión (`SKIP`).

```cmd
python main.py

```

**Omitir fase de descarga:**
Ideal para reprocesamientos cuando los archivos fuente ya se encuentran actualizados en el directorio local.

```cmd
python main.py nodownload

```

**Solo Publicación:**
Omite el refresco de orígenes de datos. Destinado a actualizaciones rápidas de formatos visuales o medidas DAX. Es importante tener en consideración que los únicos archivos que serán tomados en cuenta para esta fase serán aquellos que se encuentren en `config.json`, en la sección `ONLY_PUBLISH` (para modificar estos registros utilizar comandos indicados en la sección 4) 

```cmd
python main.py onlypublish

```

**Ejecución sobre un archivo específico:**
Aísla la ejecución a un único reporte, ignorando el resto de archivos contenidos en el directorio. El archivo debe existir físicamente en la ruta definida en `DOWNLOAD_PATH`.

```cmd
python main.py Nombre_Del_Reporte.pbix

```

**Ejecución combinada (Ejemplo avanzado):**
Publica un archivo específico omitiendo la fase de descarga en SharePoint y la actualización de orígenes de datos. (Al pasar un archivo como parámetro dinámico, el sistema gestionará su adición temporal a la lista de publicación autorizada).

```cmd
python main.py nodownload onlypublish Nombre_Del_Reporte.pbix

```

---

## 6. Resultados y Registro de Eventos (Logs)

Al finalizar la ejecución del proceso, el orquestador generará las siguientes salidas operativas:

1. **Notificación Push (Teams):** Envío de una Tarjeta Adaptable (Adaptive Card) al canal configurado detallando la cantidad de archivos procesados, éxitos y fallos.
2. **Reporte por Correo Electrónico:** Emisión de un resumen en formato HTML estructurado a la dirección registrada en el parámetro `ADDRESSE_MAIL`.
3. **Registro de Errores Técnicos:** En caso de interrupciones o fallos de lectura de orígenes de datos, el sistema registrará la traza del error en la carpeta local `/logs`, almacenando el detalle con estampa de tiempo (timestamp) para su posterior auditoría.