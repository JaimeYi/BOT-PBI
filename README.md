## Resumen del Funcionamiento

Este proyecto es un flujo de automatización diseñado para descargar, actualizar y publicar reportes de Power BI (`.pbix`) de forma desatendida. El bot simula interacciones humanas para extraer las últimas fuentes de datos desde SharePoint, procesar los reportes localmente manejando credenciales y modales de error, y finalmente publicarlos en un Workspace de Power BI en la nube, notificando el resultado a través de Microsoft Teams y correo electrónico.

---

## Arquitectura y Módulos

El sistema está dividido en cuatro componentes principales orquestados secuencialmente:

1. **`orchestrator.py` (Script Principal):** Es el punto de entrada. Maneja la interfaz de consola, ejecuta los subprocesos en orden y controla las detenciones en caso de errores críticos.
2. **`automateDownloads.py`:** Utiliza **Selenium** para abrir una sesión controlada de Google Chrome, navegar a SharePoint y descargar el archivo `.zip` con los reportes y fuentes de datos.
3. **`automateUnzip.py`:** Monitoriza la carpeta de descargas. Una vez detectado el archivo final, lo descomprime, extrae exclusivamente archivos relevantes (Excel, CSV, PBIX) a la ruta raíz y elimina el archivo comprimido original.
4. **`automatePBI.py`:** El núcleo del procesamiento. Utiliza **Pywinauto** para tomar control del sistema operativo, abrir Power BI Desktop, refrescar los orígenes de datos, inyectar credenciales (SQL, Web, SharePoint) si son requeridas, corregir rutas locales rotas y publicar el resultado.
5. **`config.json`:** Archivo centralizado de credenciales y rutas.

---

## Requisitos y Especificaciones

Para poder ejecutar este bot en un servidor o máquina local, se debe cumplir con el siguiente entorno:

### Requisitos de Sistema

* **Sistema Operativo:** Windows 10/11 o Windows Server (Requerido por Pywinauto y Power BI Desktop).
* **Software Base:**
* Python 3.9 o superior.
* Google Chrome (Última versión).
* Microsoft Power BI Desktop (Instalación clásica `.exe`, no la versión de Microsoft Store para evitar problemas de rutas).



### Dependencias de Python

Se deben instalar las siguientes librerías externas ejecutando:

```bash
pip install selenium pywinauto pyautogui pyperclip requests psutil

```

---

## Configuración (`config.json`)

Antes de la primera ejecución, se debe completar el archivo `config.json`. **Nota:** Este archivo contiene credenciales sensibles y NO debe ser subido a repositorios públicos.

| Parámetro | Descripción |
| --- | --- |
| `DOWNLOAD_PATH` | Ruta absoluta de Windows donde se descargarán y procesarán los archivos (ej. `C:\\Users\\bot\\Downloads`). |
| `WORKSPACE` | Nombre exacto del espacio de trabajo en Power BI Service donde se publicará. |
| `ADDRESSE_MAIL` | Correo electrónico que recibirá el reporte con el resumen de ejecución. |
| `SP_URL` | URL directa de SharePoint donde se aloja la carpeta desde la cual se busca descargar los reportes. |
| `WEBHOOK_URL` | URL de integración para enviar tarjetas a Microsoft Teams. |
| `CREDENTIALS` | Bloques de usuario/contraseña para entornos DEV, Base de datos, SharePoint y Salesforce. |

```json
{
    "DOWNLOAD_PATH": "",
    "WORKSPACE": "",
    "WEBHOOK_URL": "",
    "ADDRESSE_MAIL": "",
    "SP_URL": "",
    "CREDENTIALS": {
        "DEV": {
        "EMAIL": "",
        "PASSWORD": ""
        },
        "SALESFORCE": {
            "USERNAME": "",
            "PASSWORD" : ""
        },
        "SHAREPOINT": {
            "EMAIL": "",
            "PASSWORD": ""
        },
        "DATABASES": {
            "192.168.0.1": {
                "USERNAME": "",
                "PASSWORD": ""
            }
        },
        "WEBSITES": {
            "https://tusitioweb.com/": {
                "USERNAME": "",
                "PASSWORD": ""
            }
        }
    }
}
```

### ¿Cómo obtener la `WEBHOOK_URL` de Teams?

Para que el bot envíe notificaciones a un canal de Teams, debes configurar un flujo de trabajo:

1. Abre Microsoft Teams y ve al canal donde deseas recibir las alertas.
2. Haz clic en los tres puntos `(...)` en la esquina superior derecha del canal y selecciona **Flujos de trabajo** (Workflows).
3. Busca la plantilla **"Publicar en un canal cuando se reciba una solicitud de webhook"** (Post to a channel when a webhook request is received).
4. Sigue los pasos para darle un nombre a la conexión (ej. "Bot PBI") y selecciona el equipo y canal.
5. Al finalizar, el sistema generará una **URL larga**. Cópiala y pégala exactamente igual en el campo `"WEBHOOK_URL"` de tu `config.json`.
6. En caso de tener inconvenientes al realizar este proceso desde Microsoft Teams, también puede realizarse en el sitio web de Power Automate.

---

## 🚀 Modo de Uso

1. **Preparación:** Asegúrate de que ninguna instancia de Power BI Desktop esté abierta y que tu ratón y teclado no vayan a ser utilizados, ya que el bot tomará el control del cursor.
2. **Ejecución:** Abre una terminal (CMD o PowerShell) en el directorio del proyecto y ejecuta el orquestador:
```cmd
python orchestrator.py
```
3. **Opcional:** En caso de querer ejecutar uno de los flujos de trabajo de manera independiente basta con ejecutar el respectivo código.


3. **Monitoreo:** El script mostrará una interfaz en consola indicando la fase actual. Puedes minimizar la consola, pero **no minimices ni interactúes con las ventanas de Chrome o Power BI** que el bot abra.
4. **Resultados:** Al finalizar, revisa tu canal de Teams o tu bandeja de entrada para ver el resumen de ejecución. Si hubo fallos, revisa la carpeta `/logs` generada automáticamente por el sistema.
