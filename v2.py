import openpyxl
import time
import random
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# Configuración
ARCHIVO_EXCEL = "contactos.xlsx"
MENSAJE_BASE = "Hola {nombre}, este es un mensaje automático enviado a través de un bot de wpp, el mismo fue desarrollado por Mr. Fucking Good Vibes"

# Configuración de Chrome
chrome_options = Options()
chrome_options.add_argument("--user-data-dir=C:\\Temp\\ChromeProfile")
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

def enviar_mensaje(numero, nombre):
    try:
        # Paso 1: Buscar contacto mediante la barra de búsqueda
        search_box = WebDriverWait(driver, 10).until( #ORIGINALMENTE 20
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
        )
        
        # Limpiar y buscar número
        search_box.click()
        search_box.send_keys(Keys.CONTROL + "a")
        search_box.send_keys(Keys.BACKSPACE)
        search_box.send_keys(numero)
        time.sleep(2)  # Esperar resultados de búsqueda

        # Paso 2: Seleccionar chat correcto
        try:
            WebDriverWait(driver, 5).until( #ORIGINALMENTE 10
                EC.presence_of_element_located((By.XPATH, f'//span[@title="{numero}"]'))
            ).click()
        except:
            # Si no existe chat previo, abrir nuevo chat
            driver.get(f"https://web.whatsapp.com/send?phone={numero}")
        
        # Paso 3: Enviar mensaje
        input_box = WebDriverWait(driver, 10).until( #ORIGINALMENTE 15
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        
        # Estrategia de envío combinada
        mensaje = MENSAJE_BASE.format(nombre=nombre)
        input_box.send_keys(mensaje)
        time.sleep(1)
        input_box.send_keys(Keys.SHIFT + Keys.ENTER)  # Envío alternativo
        time.sleep(1)
        driver.find_element(By.XPATH, '//button[@aria-label="Enviar"]').click()
        
        print(f"Mensaje enviado a {nombre}")
        return True
        
    except Exception as e:
        print(f"Error con {numero}: {str(e)}")
        return False

try:
    driver.get("https://web.whatsapp.com")
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//div[@data-tab="3"]'))  # Esperar barra de búsqueda
    )

    # Procesar Excel
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
    hoja = wb.active

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        numero = str(fila[0]).strip().replace(" ", "").replace("-", "")  # Normalizar número
        nombre = str(fila[1]).strip()
        
        if not numero.startswith("+"):
            numero = f"+{numero}"
        
        if enviar_mensaje(numero, nombre):
            time.sleep(random.uniform(25, 40))

finally:
    driver.quit()