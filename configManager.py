import argparse
import json
import sys
import os

# Definición de colores para la interfaz
VERDE = '\033[92m'
ROJO = '\033[91m'
AMARILLO = '\033[33m'
CELESTE = '\033[96m'
RESET = '\033[0m'

class ConfigManager:
    def __init__(self, filepath="config.json"):
        self.filepath = filepath
        self.config = self._load_config()

    def _load_config(self):
        # Si el archivo no existe, lo creamos físicamente de inmediato
        if not os.path.exists(self.filepath):
            print(f"\n{AMARILLO}[INFO]{RESET} Archivo '{self.filepath}' no encontrado. Generando plantilla base...")
            
            plantilla = {
                "DOWNLOAD_PATH": "", "WORKSPACE": "", "WEBHOOK_URL": "",
                "ADDRESSE_MAIL": "", "SP_URL": "", "SP_URL_PARAMS": "",
                "CREDENTIALS": {
                    "DEV": {"EMAIL": "", "PASSWORD": ""},
                    "SALESFORCE": {"USERNAME": "", "PASSWORD": ""},
                    "SHAREPOINT": {"EMAIL": "", "PASSWORD": ""},
                    "DATABASES": {},
                    "WEBSITES": {}
                },
                "ONLY_PUBLISH": [],
                "SKIP": []
            }
            
            # Forzamos la escritura en el disco duro
            with open(self.filepath, 'w', encoding='utf-8') as f:
                json.dump(plantilla, f, indent=4, ensure_ascii=False)
                
            print(f"{VERDE}[SUCCESS]{RESET} Plantilla generada exitosamente en el directorio.\n")
            return plantilla
            
        # Si existe, lo leemos de manera estándar
        with open(self.filepath, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _save_config(self):
        with open(self.filepath, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)

    # ==========================================
    # 1. ROOT SETTINGS (Configuraciones base)
    # ==========================================
    def update_root_setting(self, key, value):
        self.config[key] = value
        self._save_config()

    # ==========================================
    # 2. CREDENCIALES FIJAS (Dev, Salesforce, SharePoint)
    # ==========================================
    def update_fixed_credential(self, service, user_value, password):
        service = service.upper()
        # Manejo inteligente de las llaves distintas en el JSON
        if service in ["DEV", "SHAREPOINT"]:
            self.config["CREDENTIALS"][service]["EMAIL"] = user_value
        elif service == "SALESFORCE":
            self.config["CREDENTIALS"][service]["USERNAME"] = user_value
            
        self.config["CREDENTIALS"][service]["PASSWORD"] = password
        self._save_config()

    # ==========================================
    # 3. CREDENCIALES DINÁMICAS (Bases de datos y Sitios Web)
    # ==========================================
    def add_or_update_database(self, host, username, password):
        self.config["CREDENTIALS"]["DATABASES"][host] = {"USERNAME": username, "PASSWORD": password}
        self._save_config()

    def delete_database(self, host):
        if host in self.config["CREDENTIALS"]["DATABASES"]:
            del self.config["CREDENTIALS"]["DATABASES"][host]
            self._save_config()
            return True
        return False

    def add_or_update_website(self, url, username, password):
        self.config["CREDENTIALS"]["WEBSITES"][url] = {"USERNAME": username, "PASSWORD": password}
        self._save_config()

    def delete_website(self, url):
        if url in self.config["CREDENTIALS"]["WEBSITES"]:
            del self.config["CREDENTIALS"]["WEBSITES"][url]
            self._save_config()
            return True
        return False

    # ==========================================
    # 4. REPORTES A PUBLICAR (Only_Publish)
    # ==========================================
    def add_report_to_publish(self, report_name):
        # 1. Obtenemos la ruta configurada
        download_path = self.config.get("DOWNLOAD_PATH", "")
        
        # 2. Validamos que la ruta base exista en la configuración
        if not download_path:
            raise ValueError("La variable 'DOWNLOAD_PATH' no está definida. Configúrala primero usando el comando 'set'.")
            
        # 3. Construimos la ruta absoluta y validamos si el archivo existe físicamente
        full_path = os.path.join(download_path, report_name)
        if not os.path.isfile(full_path):
            raise FileNotFoundError(f"El archivo '{report_name}' no se encontró en el directorio:\n    {download_path}")

        # 4. Si pasa las validaciones, lo añadimos (evitando duplicados)
        if report_name not in self.config["ONLY_PUBLISH"]:
            self.config["ONLY_PUBLISH"].append(report_name)
            self._save_config()
            return True
            
        return False

    def remove_report_from_publish(self, report_name):
        if report_name in self.config["ONLY_PUBLISH"]:
            self.config["ONLY_PUBLISH"].remove(report_name)
            self._save_config()
            return True
        return False
    
    # ==========================================
    # 5. REPORTES A SALTAR (SKIP)
    # ==========================================
    def add_report_to_skip(self, report_name):
        # 1. Obtenemos la ruta configurada
        download_path = self.config.get("DOWNLOAD_PATH", "")
        
        # 2. Validamos que la ruta base exista en la configuración
        if not download_path:
            raise ValueError("La variable 'DOWNLOAD_PATH' no está definida. Configúrala primero usando el comando 'set'.")
            
        # 3. Construimos la ruta absoluta y validamos si el archivo existe físicamente
        full_path = os.path.join(download_path, report_name)
        if not os.path.isfile(full_path):
            raise FileNotFoundError(f"El archivo '{report_name}' no se encontró en el directorio:\n    {download_path}")

        # 4. Si pasa las validaciones, lo añadimos (evitando duplicados)
        if report_name not in self.config["SKIP"]:
            self.config["SKIP"].append(report_name)
            self._save_config()
            return True
            
        return False

    def remove_report_from_skip(self, report_name):
        if report_name in self.config["SKIP"]:
            self.config["SKIP"].remove(report_name)
            self._save_config()
            return True
        return False

# ==========================================
# INTERFAZ DE LÍNEA DE COMANDOS (CLI)
# ==========================================
def main():
    parser = argparse.ArgumentParser(
        description="Gestor de Configuración del Bot Power BI\nAdministra el archivo config.json de manera segura.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    
    subparsers = parser.add_subparsers(dest="comando", help="Comandos disponibles")

    # Comando: list
    parser_list = subparsers.add_parser("list", help="Muestra un resumen de TODA la configuración")

    # Comando: set (Para configuraciones base)
    parser_set = subparsers.add_parser("set", help="Define una variable global (ej. WORKSPACE, DOWNLOAD_PATH)")
    parser_set.add_argument("key", choices=["DOWNLOAD_PATH", "WORKSPACE", "WEBHOOK_URL", "ADDRESSE_MAIL", "SP_URL", "SP_URL_PARAMS"], help="Variable a modificar")
    parser_set.add_argument("value", help="Nuevo valor")

    # Comando: set-cred (Para credenciales fijas)
    parser_cred = subparsers.add_parser("set-cred", help="Actualiza credenciales de DEV, SALESFORCE o SHAREPOINT")
    parser_cred.add_argument("servicio", choices=["DEV", "SALESFORCE", "SHAREPOINT"], type=str.upper, help="Servicio a actualizar")
    parser_cred.add_argument("usuario", help="Correo electrónico o Username")
    parser_cred.add_argument("password", help="Contraseña del servicio")

    # Comandos: Bases de Datos
    parser_add_db = subparsers.add_parser("add-db", help="Añade/Actualiza credenciales de una BD")
    parser_add_db.add_argument("host", help="IP o Nombre del Host")
    parser_add_db.add_argument("usuario", help="Nombre de usuario SQL")
    parser_add_db.add_argument("password", help="Contraseña SQL")

    parser_del_db = subparsers.add_parser("del-db", help="Elimina una BD")
    parser_del_db.add_argument("host", help="IP o Nombre del Host a eliminar")

    # Comandos: Sitios Web
    parser_add_web = subparsers.add_parser("add-web", help="Añade/Actualiza credenciales de un Sitio Web")
    parser_add_web.add_argument("url", help="URL exacta del sitio web")
    parser_add_web.add_argument("usuario", help="Nombre de usuario")
    parser_add_web.add_argument("password", help="Contraseña")

    parser_del_web = subparsers.add_parser("del-web", help="Elimina un Sitio Web")
    parser_del_web.add_argument("url", help="URL del sitio web a eliminar")

    # Comandos: Reportes
    parser_add_rep = subparsers.add_parser("add-report", help="Añade un reporte a ONLY_PUBLISH")
    parser_add_rep.add_argument("reporte", help="Nombre del archivo (ej. Ventas.pbix)")

    parser_del_rep = subparsers.add_parser("del-report", help="Elimina un reporte de ONLY_PUBLISH")
    parser_del_rep.add_argument("reporte", help="Nombre del archivo a eliminar")

    parser_add_skip_rep = subparsers.add_parser("add-skip-report", help="Añade un reporte a SKIP")
    parser_add_skip_rep.add_argument("reporte", help="Nombre del archivo (ej. Stock.pbix)")

    parser_del_skip_rep = subparsers.add_parser("del-skip-report", help="Elimina un reporte de SKIP")
    parser_del_skip_rep.add_argument("reporte", help="Nombre del archivo a eliminar")

    # Validar si no se pasaron argumentos
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)

    args = parser.parse_args()
    config = ConfigManager()

    # --- Lógica de Ejecución CLI ---
    if args.comando == "list":
        print(f"\n{CELESTE}=== ESTADO ACTUAL DE CONFIG.JSON ==={RESET}")
        print(f"{VERDE}\n[VARIABLES GLOBALES]{RESET}")
        for key in ["DOWNLOAD_PATH", "WORKSPACE", "WEBHOOK_URL", "ADDRESSE_MAIL", "SP_URL", "SP_URL_PARAMS"]:
            valor = config.config.get(key, '')
            print(f"  {key:<15} : {valor if valor else 'No definido'}")
            
        print(f"{VERDE}\n[CREDENCIALES DE SISTEMA]{RESET}")
        for srv in ["DEV", "SALESFORCE", "SHAREPOINT"]:
            datos = config.config["CREDENTIALS"].get(srv, {})
            user_key = "EMAIL" if srv in ["DEV", "SHAREPOINT"] else "USERNAME"
            usuario = datos.get(user_key, "No definido")
            pwd = "******" if datos.get("PASSWORD") else "Vacío"
            print(f"  {srv:<12} : {usuario} (Clave: {pwd})")

        print(f"{VERDE}\n[BASES DE DATOS] ({len(config.config['CREDENTIALS']['DATABASES'])}){RESET}")
        for db, cred in config.config["CREDENTIALS"]["DATABASES"].items():
            print(f"  - {db:<15} (Usuario: {cred.get('USERNAME')})")

        print(f"{VERDE}\n[SITIOS WEB] ({len(config.config['CREDENTIALS']['WEBSITES'])}){RESET}")
        for web, cred in config.config["CREDENTIALS"]["WEBSITES"].items():
            print(f"  - {web:<15} (Usuario: {cred.get('USERNAME')})")

        print(f"{VERDE}\n[REPORTES 'ONLY_PUBLISH'] ({len(config.config['ONLY_PUBLISH'])}){RESET}")
        for rep in config.config["ONLY_PUBLISH"]:
            print(f"  - {rep}")

        print(f"{VERDE}\n[REPORTES 'SKIP'] ({len(config.config['SKIP'])}){RESET}")
        for rep in config.config["SKIP"]:
            print(f"  - {rep}")
        print("\n" + "=" * 36 + "\n")

    elif args.comando == "set":
        config.update_root_setting(args.key, args.value)
        print(f"{VERDE}[SUCCESS]{RESET} Variable {args.key} actualizada correctamente.")

    elif args.comando == "set-cred":
        config.update_fixed_credential(args.servicio, args.usuario, args.password)
        print(f"{VERDE}[SUCCESS]{RESET} Credenciales de {args.servicio} actualizadas.")

    elif args.comando == "add-db":
        config.add_or_update_database(args.host, args.usuario, args.password)
        print(f"{VERDE}[SUCCESS]{RESET} Base de datos '{args.host}' actualizada.")

    elif args.comando == "del-db":
        if config.delete_database(args.host): print(f"{VERDE}[SUCCESS]{RESET} Base de datos eliminada.")
        else: print(f"{ROJO}[ERROR]{RESET} La base de datos no existe en el registro.")

    elif args.comando == "add-web":
        config.add_or_update_website(args.url, args.usuario, args.password)
        print(f"{VERDE}[SUCCESS]{RESET} Sitio Web '{args.url}' actualizado.")

    elif args.comando == "del-web":
        if config.delete_website(args.url): print(f"{VERDE}[SUCCESS]{RESET} Sitio Web eliminado.")
        else: print(f"{ROJO}[ERROR]{RESET} El Sitio Web no existe en el registro.")

    elif args.comando == "add-report":
        try:
            if config.add_report_to_publish(args.reporte): 
                print(f"{VERDE}[SUCCESS]{RESET} Reporte '{args.reporte}' añadido exitosamente.")
            else: 
                print(f"{AMARILLO}[INFO]{RESET} El reporte '{args.reporte}' ya se encontraba en la lista.")
                
        except ValueError as e:
            print(f"{ROJO}[ERROR]{RESET} Configuración incompleta: {e}")
            
        except FileNotFoundError as e:
            print(f"{ROJO}[ERROR]{RESET} Archivo no encontrado.")
            print(f"Detalle: {e}")
            print("Verifique que el nombre sea correcto (incluyendo la extensión '.pbix').")

    elif args.comando == "del-report":
        if config.remove_report_from_publish(args.reporte): print(f"{VERDE}[SUCCESS]{RESET} Reporte eliminado de la lista.")
        else: print(f"{ROJO}[ERROR]{RESET} El reporte no estaba en la lista.")

    elif args.comando == "add-skip-report":
        try:
            if config.add_report_to_skip(args.reporte): 
                print(f"{VERDE}[SUCCESS]{RESET} Reporte '{args.reporte}' añadido exitosamente.")
            else: 
                print(f"{AMARILLO}[INFO]{RESET} El reporte '{args.reporte}' ya se encontraba en la lista.")
                
        except ValueError as e:
            print(f"{ROJO}[ERROR]{RESET} Configuración incompleta: {e}")
            
        except FileNotFoundError as e:
            print(f"{ROJO}[ERROR]{RESET} Archivo no encontrado.")
            print(f"Detalle: {e}")
            print("Verifique que el nombre sea correcto (incluyendo la extensión '.pbix').")

    elif args.comando == "del-skip-report":
        if config.remove_report_from_skip(args.reporte): print(f"{VERDE}[SUCCESS]{RESET} Reporte eliminado de la lista.")
        else: print(f"{ROJO}[ERROR]{RESET} El reporte no estaba en la lista.")

if __name__ == "__main__":
    main()