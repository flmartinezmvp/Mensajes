import pywhatkit
import pandas as pd
import time
import os
from datetime import datetime, timedelta
import pyautogui  # Para tomar las capturas de pantalla

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

def cargar_imagen_y_mensaje(archivo_excel, tiempo_cambio, carpeta_screenshots):
    try:
        # Leer el archivo Excel para obtener los números de teléfono y cédulas
        df = pd.read_excel(archivo_excel)
        df['telefono'] = df['telefono'].astype(str)  # Asegurarnos de que los números sean cadenas
        df['cedula'] = df['cedula'].astype(str)  # Asegurarnos de que las cédulas sean cadenas
        df['nombre'] = df['nombre'].astype(str)  # Asegurarnos de que los nombres sean cadenas
        df['credito'] = df['credito'].astype(str)  # Asegurarnos de que el crédito sea cadena
        df['fecha_cuota'] = pd.to_datetime(df['fecha_cuota'], errors='coerce')  # Asegurarnos de que la fecha sea en formato correcto
        numeros_telefono = df['telefono'].tolist()
        cedulas = df['cedula'].tolist()
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return

    # Obtener la hora actual
    hora_actual = datetime.now()
    hora = hora_actual.hour
    minuto = hora_actual.minute + 1  # Sumar 1 minuto a la hora actual

    # Ajustar si el minuto supera 59
    if minuto >= 60:
        minuto -= 60
        hora += 1

    # Diccionarios para traducir los días y meses a español
    dias_semana = {'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles', 
                   'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'}
    meses = {'January': 'Enero', 'February': 'Febrero', 'March': 'Marzo', 'April': 'Abril', 
             'May': 'Mayo', 'June': 'Junio', 'July': 'Julio', 'August': 'Agosto', 
             'September': 'Septiembre', 'October': 'Octubre', 'November': 'Noviembre', 'December': 'Diciembre'}

    # Iterar sobre cada fila de la columna 'cedula', 'telefono', 'nombre', 'credito', 'fecha_cuota'
    for i, fila in df.iterrows():
        numero_telefono = corregir_numero_telefono(fila['telefono'])  # Asegurarse de que el número es correcto
        cedula = fila['cedula']
        nombre = fila['nombre']
        credito = fila['credito']
        fecha_cuota = fila['fecha_cuota']

        # Formatear la fecha al formato deseado: "Lunes 30 de diciembre del 2024"
        # Obtener el nombre del día y mes con strftime
        dia_semana = fecha_cuota.strftime('%A')  # Día de la semana completo (ejemplo: Monday)
        dia_mes = fecha_cuota.strftime('%d')  # Día del mes con dos dígitos (ejemplo: 30)
        mes = fecha_cuota.strftime('%B')  # Nombre completo del mes (ejemplo: December)
        anio = fecha_cuota.strftime('%Y')  # Año en formato 4 dígitos (ejemplo: 2024)

        # Traducir el nombre del día de la semana y el mes al español
        dia_semana_es = dias_semana[dia_semana]
        mes_es = meses[mes]

        # Crear la fecha final con formato: "Lunes 30 de diciembre del 2024"
        fecha_formateada_espanol = f"{dia_semana_es} {dia_mes} de {mes_es} del {anio}"

        try:
            # Construir el mensaje personalizado
            mensaje = (f"Apreciado Asociado (a), {nombre}, Coeducadores Boyacá le recuerda "
                       f"que la cuota de su obligación crediticia No. {credito} vence el día {fecha_formateada_espanol}, "
                       "Si ya realizó el pago haga caso omiso a este mensaje.")
            
            # Enviar mensaje por WhatsApp 1 minuto después de la hora actual
            print(f"Enviando mensaje a: {numero_telefono}")
            pywhatkit.sendwhatmsg(numero_telefono, mensaje, hora, minuto)  # Enviar 1 minuto después

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
            print(f"Error al enviar mensaje a {numero_telefono}: {e}")
    
    # Guardar el archivo Excel con los cambios en la columna 'enviarWhatsapp'
    try:
        df.to_excel(archivo_excel, index=False)
        print("El archivo Excel ha sido actualizado correctamente.")
    except Exception as e:
        print(f"Error al guardar el archivo Excel: {e}")

# Ejemplo de uso:
archivo_excel = "C:/Mensajes/Preventivo/datos_preventivo.xlsx"  # Ruta al archivo Excel
carpeta_screenshots = "C:/Mensajes/Preventivo/screenshots"  # Carpeta donde se guardarán las capturas de pantalla
tiempo_cambio = 10  # Tiempo en segundos para cambiar al siguiente número (puedes ajustarlo según sea necesario)

# Asegurarse de que la carpeta para capturas de pantalla exista
if not os.path.exists(carpeta_screenshots):
    os.makedirs(carpeta_screenshots)

cargar_imagen_y_mensaje(archivo_excel, tiempo_cambio, carpeta_screenshots)
