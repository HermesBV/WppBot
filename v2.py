import os
import openpyxl
import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from openpyxl.styles import PatternFill

# Configuración
ARCHIVO_EXCEL = "contactos.xlsx"
CARPETA_ERRORES = "errores"
MENSAJE_BASE = """Hola {nombre}!
Te hablo desde el S.O.S por la recorrida para ingresantes en Económicas (Av. Córdoba 2122) a la que te anotaste.
Es este *martes a las 18hrs*. El punto de encuentro va a ser en nuestra mesita, entrando por la puerta principal a la izquierda.
Igualmente va a haber estudiantes en la entrada con la remera verde del S.O.S para indicarte cómo llegar.
Cualquier cosa avisame, saludos!"""
MENSAJE_SIN_NOMBRE = """Hola!
Te hablo desde el S.O.S por la recorrida para ingresantes en Económicas (Av. Córdoba 2122) a la que te anotaste.
Es este *martes a las 18hrs*. El punto de encuentro va a ser en nuestra mesita, entrando por la puerta principal a la izquierda.
Igualmente va a haber estudiantes en la entrada con la remera verde del S.O.S para indicarte cómo llegar.
Cualquier cosa avisame, saludos!"""
COLOR_ERROR = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

# Crear carpeta de errores si no existe
if not os.path.exists(CARPETA_ERRORES):
    os.makedirs(CARPETA_ERRORES)

# Configuración de Chrome
chrome_options = Options()
chrome_options.add_argument("--user-data-dir=C:\\Temp\\ChromeProfile")
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

def marcar_error(fila, hoja):
    hoja.cell(row=fila, column=3, value="X").fill = COLOR_ERROR

def tomar_captura(numero):
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"{CARPETA_ERRORES}/error_{numero}_{timestamp}.png"
    driver.save_screenshot(nombre_archivo)
    return nombre_archivo

def enviar_mensaje(numero, nombre):
    try:
        driver.get(f"https://web.whatsapp.com/send?phone={numero}")
        
        input_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"][@contenteditable="true"][@data-tab="10"]'))
        )
        
        # Mensaje con nombre o genérico (respeta saltos de línea)
        if nombre and nombre.strip():
            mensaje = MENSAJE_BASE.format(nombre=nombre.strip())
        else:
            mensaje = MENSAJE_SIN_NOMBRE
        
        # Envía el mensaje línea por línea (para respetar formatos)
        for line in mensaje.split('\n'):
            input_box.send_keys(line)
            input_box.send_keys(Keys.SHIFT + Keys.ENTER)  # Salto de línea sin enviar
        input_box.send_keys(Keys.ENTER)  # Envía el mensaje completo
        
        WebDriverWait(driver, 10).until(
            lambda d: mensaje.replace('\n', '') in input_box.text.replace('\n', '')
        )
        
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//span[@data-testid="msg-dblcheck" or @data-testid="msg-time"]'))
            )
            print(f"✓✓ Envío confirmado a {numero}")
            return True
            
        except Exception as e:
            print(f"⚠ Envío no confirmado (pero posiblemente exitoso): {str(e)}")
            return True
        
    except Exception as e:
        captura = tomar_captura(numero)
        print(f"✗ Error con {numero}: {str(e)} | Captura: {captura}")
        return False

try:
    driver.get("https://web.whatsapp.com")
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"][@contenteditable="true"]'))
    )
    
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
    hoja = wb.active
    
    if hoja.cell(row=1, column=3).value != "Estado":
        hoja.cell(row=1, column=3, value="Estado")
    
    contador = 0
    
    for fila_idx, fila in enumerate(hoja.iter_rows(min_row=2, values_only=True), start=2):
        numero = str(fila[0]).strip().replace(" ", "").replace("-", "")
        nombre = str(fila[1]).strip() if fila[1] else ""
        
        if not numero:
            marcar_error(fila_idx, hoja)
            continue
            
        if not numero.startswith("+"):
            numero = f"+{numero}"
        
        if enviar_mensaje(numero, nombre):
            hoja.cell(row=fila_idx, column=3, value="✓")
        else:
            marcar_error(fila_idx, hoja)
        
        wb.save(ARCHIVO_EXCEL)
        contador += 1
        
        # Pausa cada 50 mensajes (5 minutos)
        if contador % 50 == 0:
            print(f"⏳ Pausa de 5 minutos después de {contador} mensajes...")
            time.sleep(300)  # 300 segundos = 5 minutos
        else:
            time.sleep(random.uniform(6, 12))  # Delay aleatorio entre 6-12s

finally:
    try:
        wb.save(ARCHIVO_EXCEL)
        wb.close()
    except Exception as e:
        print(f"Error al guardar Excel: {str(e)}")
    driver.quit()
    print("Proceso completado. Navegador cerrado.")