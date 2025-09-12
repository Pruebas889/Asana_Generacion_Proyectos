import logging
import os
import time
from selenium.webdriver import ActionChains
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from typing import Optional, List, Tuple
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import gspread
from google.oauth2.service_account import Credentials
from selenium.webdriver.common.keys import Keys

# --- Configuración de logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("automatizacion.log", encoding='utf-8')
    ]
)
TEAM_CORREOS = {
    "Team prueba": [
        "santiago.vega.cop@gmail.com",
        "david.forero.cop@gmail.com",
        "santiago.vega.cop@gmail.com"
    ],
    "Team SIICOP v.2": [
        "miguel.lopez@copservir.co",
        "miguel.lopez@copservir.co",
        "jhon.primero@copservir.co"
    ],
    "Team QA": [
        "miguel.lopez@copservir.co",
        "miguel.lopez@copservir.co"
    ],
    "Team INTEGRACIONES": [
        "miguel.lopez@copservir.co",
        "miguel.lopez@copservir.co",
        "andres.velasco@copservir.co"
    ],
    "Team POSWEB": [
        "miguel.lopez@copservir.co",
        "miguel.lopez@copservir.co",
        "juan.walteros@copservir.co"
    ],
    "Team SIICOP": [
        "miguel.lopez@copservir.co",
        "miguel.lopez@copservir.co",
        "diego.palacio@copservir.co"
    ]
    
    # Si necesitas más teams, los agregas aquí...
}

# Lista de equipos válidos
VALID_TEAMS = [
    "Team prueba",
    "Team SIICOP v.2",
    "Team QA",
    "Team INTEGRACIONES",
    "Team POSWEB",
    "Team SIICOP"
]

TEAM_SUFIJOS = {
    "Team prueba": "Prueba",
    "Team SIICOP v.2": "SIICOP V2",
    "Team QA": "QA",
    "Team INTEGRACIONES": "Integraciones - Ecommerce",
    "Team POSWEB": "POSWEB",
    "Team SIICOP": "SIICOP"
}

# --- Configuración de Google Sheets ---
SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
SERVICE_ACCOUNT_FILE = 'sheets.json'

def setup_google_sheets():
    try:
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            raise FileNotFoundError(f"Archivo de credenciales '{SERVICE_ACCOUNT_FILE}' no encontrado.")
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        sheet_id = '1CSoOvveXgLYQ9qSuVvNFOkzn28KHKcKSswqLyR4OgUQ'
        worksheet_name = 'Sprint'
        sheet = client.open_by_key(sheet_id).worksheet(worksheet_name)
        logging.info("Conexión exitosa a Google Sheets.")
        return sheet
    except Exception as e:
        logging.error(f"Error al configurar Google Sheets: {e}")
        return None

def leer_datos_sheet(sheet) -> Tuple[str, List[str]]:
    """
    Lee el nombre del portafolio en la columna A y todos los proyectos en columnas siguientes.
    """
    try:
        data = sheet.get_all_values()
        if len(data) < 2 or not data[0]:
            logging.warning("Hoja vacía o sin headers.")
            return "", []

        portfolio_name = data[1][0].strip() if len(data[1]) > 0 else ""
        projects = []
        
        if len(data[1]) > 1:
            for col_index in range(1, len(data[1])):
                project_name = data[1][col_index].strip()
                if project_name:
                    projects.append(project_name)
                else:
                    break  # Detener si hay una celda vacía
        
        if not portfolio_name:
            logging.warning("No se encontró nombre de portafolio en la fila 2, columna A.")
            return "", []

        logging.info(f"Portafolio leído: {portfolio_name}")
        logging.info(f"Proyectos leídos: {projects}")
        return portfolio_name, projects

    except Exception as e:
        logging.error(f"Error al leer datos de la hoja: {e}")
        return "", []

def borrar_datos_sheet(sheet, start_row=2, end_col='E'):
    try:
        data = sheet.get_all_values()
        if len(data) < 2:
            logging.info("No hay datos suficientes para limpiar (menos de 2 filas).")
            return
        
        # Determinar la última fila y columna con datos
        last_row = len(data)
        last_col_index = max(len(row) for row in data[1:]) if len(data) > 1 else 2
        last_col_letter = chr(65 + min(last_col_index - 1, ord('E') - 65))  # Limitar a columna E
        range_to_clear = f'A{start_row}:{last_col_letter}{last_row}'
        sheet.batch_clear([range_to_clear])
        logging.info(f"Datos de la hoja eliminados exitosamente (rango: {range_to_clear}).")
    except Exception as e:
        logging.error(f"Error al eliminar datos de la hoja: {e}")

def iniciar_driver() -> Optional[webdriver.Chrome]:
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        # options.add_argument("--headless")
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.maximize_window()
        logging.info("Driver de Chrome iniciado exitosamente.")
        return driver
    except WebDriverException as e:
        logging.error(f"Error al iniciar el driver: {e}")
        return None

def login_asana(driver):
    try:
        driver.get("https://app.asana.com/0/portfolio/1205257480867940/1207672212054810")
        logging.info("Navegando a la página de Asana.")
        
        input_correo = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='email' and @name='e']")))
        input_correo.click()
        input_correo.send_keys("javier.perdomo@copservir.co")
        logging.info("Correo ingresado.")
        
        continuar = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='button' and contains(@class,'LoginEmailForm-continueButton') and normalize-space(text())='Continue']")))
        driver.execute_script("arguments[0].click();", continuar)
        logging.info("Botón 'Continue' clickeado.")
                
        input_contrasena = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='password' and @name='p']")))
        input_contrasena.click()
        input_contrasena.send_keys("Clave123+-")
        logging.info("Contraseña ingresada.")
        
        iniciar_sesion = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='button' and contains(normalize-space(text()),'Log in')]")))
        driver.execute_script("arguments[0].click();", iniciar_sesion)
        logging.info("Botón 'Log in' clickeado.")
        
        logging.info("Login completado.")
        return True
    except Exception as e:
        logging.error(f"Error durante el login: {e}")
        return False

def navigate_to_team(driver, team_name):
    try:
        # Buscar y seleccionar el equipo especificado
        team_xpath = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'SpreadsheetPortfolioGridNameAndDetailsCellGroup-title')]//a[.//span[normalize-space()='{team_name}']]")))
        driver.execute_script("arguments[0].click();", team_xpath)
        logging.info(f"Equipo '{team_name}' seleccionado.")
        return True
    except Exception as e:
        logging.error(f"Error al navegar al equipo '{team_name}': {e}")
        return False

def create_portfolio(driver, portfolio_name):
    try:
        click_boton_crear_proyecto = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='button' and contains(@aria-label,'More actions')]")))
        driver.execute_script("arguments[0].click();", click_boton_crear_proyecto)
        logging.info("Menú 'More actions' abierto.")
        
        clic_crear_portafolio = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='menuitem']//span[contains(text(),'Crear un portafolio nuevo')]")))
        driver.execute_script("arguments[0].click();", clic_crear_portafolio)
        logging.info("Opción 'Crear un portafolio nuevo' seleccionada.")
        
        time.sleep(3)
        
        input_nombre_portafolio = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@id='new_portfolio_dialog_content_name_input']")))
        input_nombre_portafolio.click()
        input_nombre_portafolio.clear()
        input_nombre_portafolio.send_keys(portfolio_name)
        logging.info(f"Nombre del portafolio ingresado: {portfolio_name}")
        
        time.sleep(0.5)
        boton_crear_portafolio = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='button' and contains(normalize-space(.), 'Continuar')]")))
        driver.execute_script("arguments[0].click();", boton_crear_portafolio)
        logging.info("Botón 'Continuar' para portafolio clickeado.")
        
        time.sleep(0.5)
        # Seleccionar opción "Comparte con compañeros de equipo"
        opcion_compartir = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@class='CreatePortfolioModalNextStepForm-rowItem']//label[contains(., 'Comparte con compañeros de equipo')]"))
        )
        driver.execute_script("arguments[0].click();", opcion_compartir)
        logging.info("Opción 'Comparte con compañeros de equipo' seleccionada.")

        time.sleep(0.5)
        ve_al_portafolio = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='button' and contains(normalize-space(.),'Ve al portafolio')]")))
        driver.execute_script("arguments[0].click();", ve_al_portafolio)
        logging.info(f"Portafolio '{portfolio_name}' creado y accedido exitosamente.")
        time.sleep(3)        
        return True
    except Exception as e:
        logging.error(f"Error al crear portafolio: {e}")
        return False

def agregar_invitados_team(driver, correos: List[str]):
    """
    Agrega invitados al portafolio ingresando correos y presionando Enter por cada uno.
    """
    try:
        for correo in correos:
            try:
                time.sleep(1)  # pequeña pausa para evitar problemas de carga
                WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[contains(@class,'TokenizerInput-input')]"))
                )                # Reintento en caso de que el campo se refresque
                input_emails = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, "//input[contains(@class,'TokenizerInput-input')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", input_emails)
                input_emails.click()
                driver.execute_script("arguments[0].click();", input_emails)
                time.sleep(2)  # pequeña pausa para dar tiempo a la UI
                input_emails.send_keys(correo)
                input_emails.send_keys(Keys.ARROW_DOWN)
                time.sleep(0.3)
                input_emails.send_keys(Keys.ENTER)
                logging.info(f"Correo '{correo}' ingresado y confirmado.")

                time.sleep(0.5)

            except StaleElementReferenceException:
                logging.warning(f"Elemento para correo '{correo}' se refrescó. Reintentando...")
                time.sleep(1)
                continue  # vuelve a intentar en la siguiente iteración

        time.sleep(10)
        # Hacer clic en botón 'Invitar' al final
        boton_invitar = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'ButtonPrimaryPresentation') and normalize-space()='Invitar']"))
        )
        driver.execute_script("arguments[0].click();", boton_invitar)
        logging.info("Botón 'Invitar' clickeado. Invitaciones enviadas.")

        time.sleep(2)
        return True

    except Exception as e:
        logging.error(f"Error al agregar invitados: {e}")
        return False


def crear_proyecto(driver, project_name, portfolio_name):
    try:
        click_boton_crear_proyecto = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='button' and contains(@aria-label,'More actions')]")))
        driver.execute_script("arguments[0].click();", click_boton_crear_proyecto)
        logging.info("Menú 'More actions' abierto para proyecto.")
        
        clic_crear_proyecto = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='menuitem']//span[contains(text(),'Crear un proyecto nuevo')]")))
        driver.execute_script("arguments[0].click();", clic_crear_proyecto)
        logging.info("Opción 'Crear un proyecto nuevo' seleccionada.")
        
        crear_proyecto_en_blanco = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='button' and contains(text(),'Proyecto en blanco')]")))
        driver.execute_script("arguments[0].click();", crear_proyecto_en_blanco)
        logging.info("Tipo 'Proyecto en blanco' seleccionado.")
        
        input_nombre_proyecto = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@id='new_project_dialog_content_name_input']")))
        input_nombre_proyecto.click()
        input_nombre_proyecto.clear()
        input_nombre_proyecto.send_keys(project_name)
        logging.info(f"Nombre del proyecto ingresado: {project_name}")
        
        continuar_creacion_proyecto = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='button' and contains(normalize-space(.), 'Continuar')]")))
        driver.execute_script("arguments[0].click();", continuar_creacion_proyecto)
        logging.info("Botón 'Continuar' para proyecto clickeado.")
        
        crear_proyecto = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='button' and contains(normalize-space(.), 'Crear proyecto')]")))
        driver.execute_script("arguments[0].click();", crear_proyecto)
        logging.info(f"Proyecto '{project_name}' creado exitosamente.")
        
        time.sleep(5)
        volver_a_portafolio = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, f"//a[contains(@class,'NavigationBreadcrumbContent-portfolioNameAndIcon')]//span[contains(normalize-space(.),'{portfolio_name}')]")))
        driver.execute_script("arguments[0].click();", volver_a_portafolio)
        logging.info(f"Regresando al portafolio '{portfolio_name}'.")
        time.sleep(3)
        return True
    except Exception as e:
        logging.error(f"Error al crear proyecto '{project_name}': {e}")
        return False

def procesar_teams(driver, sheet):
    """
    Recorre todos los teams válidos, crea portafolio y proyectos, 
    y guarda todo en la hoja con columnas: Tipo | Nombre | Link | Team
    """
    # Fila inicial
    row_index = 6  

    try:
        portfolio_name, projects = leer_datos_sheet(sheet)
        if not portfolio_name or not projects:
            logging.error("No hay datos válidos para procesar.")
            return

        for team in VALID_TEAMS:
            logging.info(f"Procesando equipo: {team}")

            # Navegar al team
            if not navigate_to_team(driver, team):
                logging.error(f"No se pudo navegar al team {team}. Se continúa con el siguiente.")
                continue

            # Crear portafolio
            if create_portfolio(driver, portfolio_name):
                correos_team = TEAM_CORREOS.get(team, [])
                if correos_team:
                    agregar_invitados_team(driver, correos_team)
                else:
                    logging.warning(f"No hay correos configurados para el team {team}.")
                portfolio_url = driver.current_url

                # Guardar portafolio en tabla: Tipo | Nombre | Link | Team
                sheet.update_cell(row_index, 1, "Portafolio")
                sheet.update_cell(row_index, 2, portfolio_name)
                sheet.update_cell(row_index, 3, portfolio_url)
                sheet.update_cell(row_index, 4, team)
                logging.info(f"Portafolio '{portfolio_name}' creado. Guardado en fila {row_index}.")

                row_index += 1  # Avanzar a siguiente fila para proyectos
            else:
                logging.error(f"No se pudo crear portafolio {portfolio_name} en team {team}.")
                continue

            # Crear proyectos con sufijo del team
            suffix = TEAM_SUFIJOS.get(team, team)
            for project in projects:
                project_name = f"{project} {suffix}"
                if crear_proyecto(driver, project_name, portfolio_name):
                    project_url = driver.current_url

                    # Guardar proyecto en tabla
                    sheet.update_cell(row_index, 1, "Proyecto")
                    sheet.update_cell(row_index, 2, project_name)
                    sheet.update_cell(row_index, 3, project_url)
                    sheet.update_cell(row_index, 4, team)
                    logging.info(f"Proyecto '{project_name}' creado. Guardado en fila {row_index}.")

                    row_index += 1
                else:
                    logging.warning(f"No se pudo crear proyecto {project_name} en {team}")

            # Fila en blanco para separar cada team
            row_index += 1  

            # Regresar a la página principal de Sprints
            try:
                sprints_link = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class,'NavigationBreadcrumbContent-portfolioNameAndIcon')]//span[contains(normalize-space(.),'Sprints')]"))
                )
                driver.execute_script("arguments[0].click();", sprints_link)
                time.sleep(2)
            except Exception as e:
                logging.error(f"No se pudo regresar a la página de Sprints: {e}")

        logging.info("Procesamiento completado para todos los teams.")

    except Exception as e:
        logging.error(f"Error en procesar_teams: {e}")

def main():
    driver = iniciar_driver()
    if not driver:
        logging.error("No se pudo iniciar el driver. Terminando.")
        return
    
    sheet = setup_google_sheets()
    if not sheet:
        logging.error("No se pudo conectar a Google Sheets. Terminando.")
        return
    
    try:
        # Iniciar sesión en Asana
        if not login_asana(driver):
            logging.error("Fallo en el login. Terminando.")
            return
        
        # Procesar todos los equipos y proyectos
        procesar_teams(driver, sheet)

        logging.info("Automatización completada exitosamente para todos los equipos y proyectos.")
        
    except Exception as e:
        logging.error(f"Error general en la ejecución: {e}")
    
    finally:
        if driver:
            driver.quit()
            logging.info("Driver cerrado.")

if __name__ == "__main__":
    main()