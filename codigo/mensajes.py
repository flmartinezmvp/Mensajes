
# import pywhatkit
# import pandas as pd
# import time
# import os

# def corregir_numero_telefono(numero):
#     # Asegurarse de que los números tengan el formato correcto (+ y código de país)
#     if not numero.startswith('+'):
#         return '+57' + numero  # Agregar el "+" si falta
#     return numero

# def cargar_imagen_y_mensaje(archivo_excel, carpeta_imagenes, tiempo_cambio):
#     try:
#         # Leer el archivo Excel para obtener los números de teléfono y cédulas
#         df = pd.read_excel(archivo_excel)
#         df['telefono'] = df['telefono'].astype(str)  # Asegurarnos de que los números sean cadenas
#         df['cedula'] = df['cedula'].astype(str)  # Asegurarnos de que las cédulas sean cadenas
#         numeros_telefono = df['telefono'].tolist()
#         cedulas = df['cedula'].tolist()
#     except Exception as e:
#         print(f"Error al leer el archivo Excel: {e}")
#         return

#     # Iterar sobre cada fila de la columna 'cedula' y 'telefono'
#     for i, fila in df.iterrows():
#         numero_telefono = corregir_numero_telefono(fila['telefono'])  # Asegurarse de que el número es correcto
#         cedula = fila['cedula']
        
#         # Construir la ruta de la imagen usando la cédula
#         ruta_imagen = os.path.join(carpeta_imagenes, f"{cedula}.jpg")
        
#         # Verificar si el archivo de imagen existe
#         if os.path.exists(ruta_imagen):
#             try:
#                 print(f"Se cargaría la imagen para el numero: {numero_telefono}")
#                 print(f"Imagen: {ruta_imagen}")
#                 print(f"Mensaje: 'Hola'")  # Puedes modificar el mensaje si es necesario
#                 # Cargar la imagen y el mensaje en WhatsApp Web, con un tiempo de espera para cargar
#                 pywhatkit.sendwhats_image(numero_telefono, ruta_imagen, "Buen día, se envia notificación", 10)  # Tiempo de espera de 10 segundos
                
#                 # Espera para cambiar al siguiente número
#                 print(f"Esperando {tiempo_cambio} segundos para cambiar al siguiente numero...")
#                 time.sleep(tiempo_cambio)  # Espera el tiempo indicado antes de continuar con el siguiente número
#             except Exception as e:
#                 print(f"Error al preparar el envío de imagen a {numero_telefono}: {e}")
#         else:
#             print(f"Archivo no encontrado para la cédula {cedula}. No se enviará imagen.")
        
# # Ejemplo de uso:
# archivo_excel = "C:/Mensajes/datos.xlsx"  # Ruta al archivo Excel
# carpeta_imagenes = "C:/Mensajes/jpg"  # Carpeta donde están las imágenes
# tiempo_cambio = 10  # Tiempo en segundos para cambiar al siguiente número (puedes ajustarlo según sea necesario)

# cargar_imagen_y_mensaje(archivo_excel, carpeta_imagenes, tiempo_cambio)


#------------------------------------ check 5:10 -> 17-12-2024 --------------------

# import pywhatkit
# import pandas as pd
# import time
# import os
# from datetime import datetime

# def corregir_numero_telefono(numero):
#     # Asegurarse de que los números tengan el formato correcto (+ y código de país)
#     if not numero.startswith('+'):
#         return '+57' + numero  # Agregar el "+" si falta
#     return numero

# def cargar_imagen_y_mensaje(archivo_excel, carpeta_imagenes, tiempo_cambio):
#     try:
#         # Leer el archivo Excel para obtener los números de teléfono y cédulas
#         df = pd.read_excel(archivo_excel)
#         df['telefono'] = df['telefono'].astype(str)  # Asegurarnos de que los números sean cadenas
#         df['cedula'] = df['cedula'].astype(str)  # Asegurarnos de que las cédulas sean cadenas
#         numeros_telefono = df['telefono'].tolist()
#         cedulas = df['cedula'].tolist()
#     except Exception as e:
#         print(f"Error al leer el archivo Excel: {e}")
#         return

#     # Iterar sobre cada fila de la columna 'cedula' y 'telefono'
#     for i, fila in df.iterrows():
#         numero_telefono = corregir_numero_telefono(fila['telefono'])  # Asegurarse de que el número es correcto
#         cedula = fila['cedula']
        
#         # Construir la ruta de la imagen usando la cédula
#         ruta_imagen = os.path.join(carpeta_imagenes, f"{cedula}.jpg")
        
#         # Verificar si el archivo de imagen existe
#         if os.path.exists(ruta_imagen):
#             try:
#                 print(f"Se cargaría la imagen para el numero: {numero_telefono}")
#                 print(f"Imagen: {ruta_imagen}")
#                 print(f"Mensaje: 'Buen día, se envía notificación'")  # Puedes modificar el mensaje si es necesario
                
#                 # Cargar la imagen y el mensaje en WhatsApp Web, con un tiempo de espera para cargar
#                 pywhatkit.sendwhats_image(numero_telefono, ruta_imagen, "Buen día, se envia notificación", 10)  # Tiempo de espera de 10 segundos
                
#                 # Capturar la hora de envío
#                 hora_envio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#                 print(f"Mensaje enviado a {numero_telefono} a las {hora_envio}")
                
#                 # Actualizar la columna 'enviarWhatsapp' con la hora de envío
#                 df.at[i, 'enviarWhatsapp'] = hora_envio
                
#                 # Espera para cambiar al siguiente número
#                 print(f"Esperando {tiempo_cambio} segundos para cambiar al siguiente numero...")
#                 time.sleep(tiempo_cambio)  # Espera el tiempo indicado antes de continuar con el siguiente número
#             except Exception as e:
#                 print(f"Error al preparar el envío de imagen a {numero_telefono}: {e}")
#         else:
#             print(f"Archivo no encontrado para la cédula {cedula}. No se enviará imagen.")
    
#     # Guardar el archivo Excel con los cambios en la columna 'enviarWhatsapp'
#     try:
#         df.to_excel(archivo_excel, index=False)
#         print("El archivo Excel ha sido actualizado correctamente.")
#     except Exception as e:
#         print(f"Error al guardar el archivo Excel: {e}")

# # Ejemplo de uso:
# archivo_excel = "C:/Mensajes/datos.xlsx"  # Ruta al archivo Excel
# carpeta_imagenes = "C:/Mensajes/jpg"  # Carpeta donde están las imágenes
# tiempo_cambio = 10  # Tiempo en segundos para cambiar al siguiente número (puedes ajustarlo según sea necesario)

# cargar_imagen_y_mensaje(archivo_excel, carpeta_imagenes, tiempo_cambio)


#------------------------------creacion de screenshot con nombre siendo la cédula-----------------------
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

def cargar_imagen_y_mensaje(archivo_excel, carpeta_imagenes, tiempo_cambio, carpeta_screenshots):
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
        
        # Construir la ruta de la imagen usando la cédula
        ruta_imagen = os.path.join(carpeta_imagenes, f"{cedula}.jpg")
        
        # Verificar si el archivo de imagen existe
        if os.path.exists(ruta_imagen):
            try:
                print(f"Se cargaría la imagen para el numero: {numero_telefono}")
                print(f"Imagen: {ruta_imagen}")
                print(f"Mensaje: 'Buen día Sr(a) asociado(a) COEDUCADORES BOYACÁ adjunta NOTIFICACIÓN HABEAS DATA'")  # Puedes modificar el mensaje si es necesario
                
                # Cargar la imagen y el mensaje en WhatsApp Web, con un tiempo de espera para cargar
                pywhatkit.sendwhats_image(numero_telefono, ruta_imagen, "Buen día, se envia notificación", 10)  # Tiempo de espera de 10 segundos
                
                # Capturar la hora de envío
                hora_envio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                print(f"Mensaje enviado a {numero_telefono} a las {hora_envio}")
                
                # Actualizar la columna 'enviarWhatsapp' con la hora de envío
                df.at[i, 'enviarWhatsapp'] = hora_envio
                
                # Espera de 3 segundos después de enviar el mensaje antes de tomar el pantallazo
                print("Esperando 3 segundos para tomar el pantallazo...")
                time.sleep(3)  # Espera de 3 segundos para capturar la pantalla
                
                # Tomar el pantallazo y guardarlo con el número de cédula como nombre
                tomar_screenshot(carpeta_screenshots, cedula)
                
                # Espera para cambiar al siguiente número
                print(f"Esperando {tiempo_cambio} segundos para cambiar al siguiente numero...")
                time.sleep(tiempo_cambio)  # Espera el tiempo indicado antes de continuar con el siguiente número
            except Exception as e:
                print(f"Error al preparar el envío de imagen a {numero_telefono}: {e}")
        else:
            print(f"Archivo no encontrado para la cédula {cedula}. No se enviará imagen.")
    
    # Guardar el archivo Excel con los cambios en la columna 'enviarWhatsapp'
    try:
        df.to_excel(archivo_excel, index=False)
        print("El archivo Excel ha sido actualizado correctamente.")
    except Exception as e:
        print(f"Error al guardar el archivo Excel: {e}")

# Ejemplo de uso:
archivo_excel = "C:/Mensajes/datos.xlsx"  # Ruta al archivo Excel
carpeta_imagenes = "C:/Mensajes/jpg"  # Carpeta donde están las imágenes
carpeta_screenshots = "C:/Mensajes/screenshots"  # Carpeta donde se guardarán las capturas de pantalla
tiempo_cambio = 10  # Tiempo en segundos para cambiar al siguiente número (puedes ajustarlo según sea necesario)

# Asegurarse de que la carpeta para capturas de pantalla exista
if not os.path.exists(carpeta_screenshots):
    os.makedirs(carpeta_screenshots)

cargar_imagen_y_mensaje(archivo_excel, carpeta_imagenes, tiempo_cambio, carpeta_screenshots)
