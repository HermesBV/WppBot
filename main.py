import openpyxl
import pywhatkit as kit
import time

# Configuración
ARCHIVO_EXCEL = "contactos.xlsx"  # Nombre del archivo Excel
MENSAJE_BASE = "Hola {nombre}, este es un mensaje automático enviado a través de un bot de wpp desarrollado por Mister Good Vibes." 

# Leer el Excel
wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
hoja = wb.active

for fila in hoja.iter_rows(min_row=2, values_only=True):  # Ignorar la primera fila (encabezados)
    numero_corto = str(fila[0])  # Columna "numero"
    numero = "+54" + numero_corto
    nombre = fila[1]  # Columna "nombre"
    
    if not numero or not nombre:
        continue  # Saltar si falta dato
    
    # Personalizar el mensaje
    mensaje = MENSAJE_BASE.format(nombre=nombre)
    
    
    print(f"Enviando mensaje a {nombre} ({numero})...")
    
    try:
        # Enviar mensaje
        kit.sendwhatmsg_instantly(
            phone_no=numero,
            message=mensaje,
            wait_time=10,  # Tiempo de espera antes de enviar (segundos)
            tab_close=True  # Cerrar pestaña después de enviar
        )
        
        # Esperar entre mensajes para evitar bloqueos
        time.sleep(10)  # Mínimo 20-30 segundos entre mensajes
        
    except Exception as e:
        
        
        print(f"Error al enviar a {numero}: {e}")




print("Proceso completado.")