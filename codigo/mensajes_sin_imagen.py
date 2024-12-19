import pywhatkit
import pandas as pd
import time
import os
from datetime import datetime
import pyautogui  # Importar la librería para capturar pantallas

def corregir_numero_telefono(numero):
    # Asegurarse de que los números tengan el formato correcto (+ y código de país)
    if not numero.startswith('+'):
        return '+57' + numero  # Agregar el "+" si falta
    return numero

def tomar_screenshot(carpeta_screenshots, cedula):
    try:
        # Usamos solo la cédula como nombre del archivo
        ruta_screenshot = os.path.join(carpeta_screenshots, f"{cedula}.png")
        
        # Tomar la captura de pantalla
        screenshot = pyautogui.screenshot()
        screenshot.save(ruta_screenshot)  # Guardar la captura en la carpeta
        print(f"Captura de pantalla guardada en {ruta_screenshot}")
        return ruta_screenshot  # Devolvemos la ruta para usarla si es necesario
    except Exception as e:
        print(f"Error al tomar la captura de pantalla: {e}")
        return None

def cargar_mensaje(archivo_excel, carpeta_screenshots, tiempo_cambio):
    try:
        # Leer el archivo Excel para obtener los números de teléfono y cédulas
        df = pd.read_excel(archivo_excel)
        df['telefono'] = df['telefono'].astype(str)  # Asegurarnos de que los números sean cadenas
        df['cedula'] = df['cedula'].astype(str)  # Asegurarnos de que las cédulas sean cadenas
        numeros_telefono = df['telefono'].tolist()
        cedulas = df['cedula'].tolist()
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return

    # Iterar sobre cada fila de la columna 'cedula' y 'telefono'
    for i, fila in df.iterrows():
        numero_telefono = corregir_numero_telefono(fila['telefono'])  # Asegurarse de que el número es correcto
        cedula = fila['cedula']
        
        try:
            print(f"Enviando mensaje al numero: {numero_telefono}")
            print(f"Mensaje: 'Buen día, se envía notificación'")  # Puedes modificar el mensaje si es necesario
            
            # Enviar el mensaje de texto a través de WhatsApp Web, con un tiempo de espera para cargar
            pywhatkit.sendwhatmsg(numero_telefono, "Buen día, se envía notificación", 10)  # Tiempo de espera de 10 segundos
            
            # Capturar la hora de envío
            hora_envio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"Mensaje enviado a {numero_telefono} a las {hora_envio}")
            
            # Actualizar la columna 'enviarWhatsapp' con la hora de envío
            df.at[i, 'enviarWhatsapp'] = hora_envio
            
            # Espera de 3 segundos antes de tomar la captura de pantalla
            print("Esperando 3 segundos para tomar el pantallazo...")
            time.sleep(3)  # Espera de 3 segundos para capturar la pantalla
            
            # Tomar el pantallazo y guardarlo con el número de cédula como nombre
            tomar_screenshot(carpeta_screenshots, cedula)
            
            # Espera para cambiar al siguiente número
            print(f"Esperando {tiempo_cambio} segundos para cambiar al siguiente numero...")
            time.sleep(tiempo_cambio)  # Espera el tiempo indicado antes de continuar con el siguiente número
        except Exception as e:
            print(f"Error al enviar mensaje a {numero_telefono}: {e}")
    
    # Guardar el archivo Excel con los cambios en la columna 'enviarWhatsapp'
    try:
        df.to_excel(archivo_excel, index=False)
        print("El archivo Excel ha sido actualizado correctamente.")
    except Exception as e:
        print(f"Error al guardar el archivo Excel: {e}")

# Ejemplo de uso:
archivo_excel = "C:/Mensajes/datos.xlsx"  # Ruta al archivo Excel
carpeta_screenshots = "C:/Mensajes/screenshots"  # Carpeta donde se guardarán las capturas de pantalla
tiempo_cambio = 10  # Tiempo en segundos para cambiar al siguiente número (puedes ajustarlo según sea necesario)

# Asegurarse de que la carpeta para capturas de pantalla exista
if not os.path.exists(carpeta_screenshots):
    os.makedirs(carpeta_screenshots)

cargar_mensaje(archivo_excel, carpeta_screenshots, tiempo_cambio)
