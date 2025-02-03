import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import os
import shutil
import PyPDF2
import sys
import pdfplumber
import re
import fitz  # PyMuPDF
import pywhatkit
import pandas as pd
import time
from datetime import datetime
import pyautogui  # Importar la librería para capturar pantallas

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
# Establecer la codificación de la consola para UTF-8
sys.stdout.reconfigure(encoding='utf-8')

class habeasData(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.title("Habeas Data")
        self.geometry("400x800")
        self.config(bg="#2C3E50")

        label = tk.Label(self, text="Seleccione un documento para cargar", font=("Helvetica", 14), fg="white", bg="#2C3E50")
        label.pack(pady=30)

        self.boton_cargar = self.create_button("Cargar pdf Notificaciones", self.cargar_documento)
        self.boton_cargar.pack(pady=20)
        
        self.boton_cargar_datos = self.create_button("Cargar Documento de Datos", self.cargar_documento_datos)
        self.boton_cargar_datos.pack(pady=20)
        self.boton_cargar_datos.config(state="disabled")
        
        self.boton_dividir = self.create_button("Dividir Documento", self.dividir_pdf)
        self.boton_dividir.pack(pady=20)
        self.boton_dividir.config(state="disabled")
        
        self.boton_nombrar = self.create_button("Nombrar Archivos", self.renombrar_archivos_en_carpeta)
        self.boton_nombrar.pack(pady=20)
        self.boton_nombrar.config(state="disabled")
        
        self.boton_convertir = self.create_button("Convertir pdf a jpg", self.convertir_pdf_jpg)
        self.boton_convertir.pack(pady=20)
        self.boton_convertir.config(state="disabled")
        
        self.boton_mensaje_whatsapp = self.create_button("Enviar mensajes WhatsApp", self.cargar_imagen_y_mensaje)
        self.boton_mensaje_whatsapp.pack(pady=20)
        self.boton_mensaje_whatsapp.config(state="normal")
        
        self.boton_enviar_correo = self.create_button("Enviar Correos", self.enviar_correo)
        self.boton_enviar_correo.pack(pady=20)
        self.boton_enviar_correo.config(state="normal")
        
        

    def create_button(self, text, command):
        button = tk.Button(self, text=text, width=25, height=2, font=("Helvetica", 12),
                        bg="#16A085", fg="white", relief="solid", bd=2, command=command)
        button.config(activebackground="#1ABC9C", activeforeground="white")
        return button

    # Cargar Documento
    def cargar_documento(self):
        archivo = askopenfilename(title="Seleccionar Documento", filetypes=[("Archivos PDF", "*.pdf"), ("Todos los Archivos", "*.*")])
        carpeta_destino = "C:/mensajes/habeas_data/archivo"
    
        if archivo:
            nombre_archivo = os.path.basename(archivo)

            if nombre_archivo.lower() != "habeas_data.pdf":
                messagebox.showwarning("Nombre de documento inválido", "El archivo debe llamarse 'habeas_data.pdf'. Por favor, cambie el nombre del archivo y vuelva a intentarlo.")
            else:
                # Verificar si la carpeta destino existe, si no, crearla
                if not os.path.exists(carpeta_destino):
                    try:
                        os.makedirs(carpeta_destino)
                    except Exception as e:
                        messagebox.showerror("Error al crear la carpeta", f"No se pudo crear la carpeta de destino. Error: {e}")
                        return  # Salir de la función si no se pudo crear la carpeta

                # Definir la ruta completa del archivo de destino
                archivo_destino = os.path.join(carpeta_destino, nombre_archivo)

                try:
                    # Copiar el archivo a la carpeta destino
                    shutil.copy(archivo, archivo_destino)
                    messagebox.showinfo("Documento Cargado", f"El archivo se ha subido correctamente")

                    # Inhabilitar el botón de cargar archivo después de subirlo
                    self.boton_cargar.config(state="disabled")
                    self.boton_cargar_datos.config(state="normal")
                    
                except Exception as e:
                    messagebox.showerror("Error al copiar el archivo", f"No se pudo copiar el archivo. Error: {e}")
        else:
            messagebox.showwarning("No se seleccionó archivo", "No se ha seleccionado ningún archivo.")

    def cargar_documento_datos(self):
        archivo = askopenfilename(title="Seleccionar Documento", filetypes=[("Archivos PDF", "*.pdf"), ("Todos los Archivos", "*.*")])
        carpeta_destino = "C:/mensajes/habeas_data/datos"
    
        if archivo:
            nombre_archivo = os.path.basename(archivo)

            if nombre_archivo.lower() != "datos.xlsx":
                messagebox.showwarning("Nombre de documento inválido", "El archivo debe llamarse 'datos.xlsx'. Por favor, cambie el nombre del archivo y vuelva a intentarlo.")
            else:
                # Verificar si la carpeta destino existe, si no, crearla
                if not os.path.exists(carpeta_destino):
                    try:
                        os.makedirs(carpeta_destino)
                    except Exception as e:
                        messagebox.showerror("Error al crear la carpeta", f"No se pudo crear la carpeta de destino. Error: {e}")
                        return  # Salir de la función si no se pudo crear la carpeta

                # Definir la ruta completa del archivo de destino
                archivo_destino = os.path.join(carpeta_destino, nombre_archivo)

                try:
                    # Copiar el archivo a la carpeta destino
                    shutil.copy(archivo, archivo_destino)
                    messagebox.showinfo("Documento Cargado", f"El archivo se ha subido correctamente")

                    # Inhabilitar el botón de cargar archivo después de subirlo
                    self.boton_cargar_datos.config(state="disabled")
                    self.boton_dividir.config(state="normal")
                    
                except Exception as e:
                    messagebox.showerror("Error al copiar el archivo", f"No se pudo copiar el archivo. Error: {e}")
        else:
            messagebox.showwarning("No se seleccionó archivo", "No se ha seleccionado ningún archivo.")



    def dividir_pdf(self):
        # La carpeta donde se guardarán los archivos
        ruta_destino = "C:/Mensajes/habeas_data/pdf" 

        # Verificar si la carpeta de destino existe, si no, crearla
        os.makedirs(ruta_destino, exist_ok=True)

        archivo_pdf = "C:/Mensajes/habeas_data/archivo/habeas_data.pdf"  # Ruta del archivo PDF original
        num_paginas_por_archivo = 1  # Dividir en partes de 1 página cada una

        # Abrir el archivo PDF
        with open(archivo_pdf, 'rb') as archivo:
            lector_pdf = PyPDF2.PdfReader(archivo)
            total_paginas = len(lector_pdf.pages)
            archivos_generados = total_paginas
            # Dividir el PDF en partes más pequeñas
            for i in range(0, total_paginas, num_paginas_por_archivo):
                escritor_pdf = PyPDF2.PdfWriter()

                # Añadir las páginas correspondientes al nuevo archivo
                for j in range(i, min(i + num_paginas_por_archivo, total_paginas)):
                    escritor_pdf.add_page(lector_pdf.pages[j])

                # Guardar la nueva parte del PDF en la carpeta de destino
                nombre_archivo = os.path.join(ruta_destino, f"parte_{i//num_paginas_por_archivo + 1}.pdf")
                with open(nombre_archivo, 'wb') as salida:
                    escritor_pdf.write(salida)
        messagebox.showinfo("Proceso Completado", f"Se generaron {archivos_generados} archivos PDF.\nEl proceso se completó exitosamente.")
        self.boton_dividir.config(state="disabled")
        self.boton_nombrar.config(state="normal")
        

    # Función para extraer el texto de todas las páginas de un PDF
    def extraer_todo_texto_pdf(self, archivo_pdf):
        texto = ""

        try:
            # Abrir el archivo PDF
            with pdfplumber.open(archivo_pdf) as pdf:
                # Iterar sobre cada página del PDF
                for pagina in pdf.pages:
                    texto += pagina.extract_text() + "\n"  # Extraer el texto de la página y añadir un salto de línea entre páginas

        except FileNotFoundError:
            print(f"Error: El archivo '{archivo_pdf}' no se encuentra.")
            return ""
        except Exception as e:
            print(f"Ha ocurrido un error: {e}")
            return ""

        return texto

    # Función para buscar una palabra entre dos palabras clave
    def buscar_palabra_entre(self, texto, palabra_antes, palabra_despues):
        # Limpiar el texto eliminando saltos de línea y espacios extras
        texto_limpio = texto.replace("\n", " ")  # Eliminar saltos de línea
        texto_limpio = re.sub(r'\s+', ' ', texto_limpio)  # Reemplazar múltiples espacios por uno solo

        # Expresión regular para buscar la palabra entre las dos palabras clave
        patron = r'(?<=\b' + re.escape(palabra_antes) + r'\b)(.*?)(?=\b' + re.escape(palabra_despues) + r'\b)'

        # Buscar el texto entre las dos palabras clave
        resultado = re.findall(patron, texto_limpio, re.DOTALL)

        return resultado

    # Función para renombrar todos los archivos PDF en una carpeta
    def renombrar_archivos_en_carpeta(self):
        contador  = 0
        #definir la ruta de la carpeta
        carpeta = "C:/Mensajes/habeas_data/pdf" 
        # Obtener la lista de archivos en la carpeta
        archivos = os.listdir(carpeta)

        self.ventana_esperando = tk.Toplevel(self)
        self.ventana_esperando.title("Esperando...")
        self.ventana_esperando.geometry("300x100")
        etiqueta_esperando = tk.Label(self.ventana_esperando, text="Nombrando los archivos...\nEspere por favor.")
        etiqueta_esperando.pack(pady=20)
        self.update()  # Esto asegura que la ventana se actualice
        
        for archivo in archivos:
            # Verificar si el archivo es un archivo PDF
            if archivo.lower().endswith(".pdf"):
                archivo_pdf = os.path.join(carpeta, archivo)  # Ruta completa al archivo PDF
                print(f"\nProcesando el archivo: {archivo_pdf}")
                contador +=1

                # Extraer el texto del archivo PDF
                texto_extraido = self.extraer_todo_texto_pdf(archivo_pdf)

                if texto_extraido:
                    # Limpiar el texto eliminando saltos de línea y espacios adicionales
                    texto_limpio = texto_extraido.replace("\n", " ")  # Eliminar saltos de línea
                    texto_limpio = re.sub(r'\s+', ' ', texto_limpio)  # Reemplazar múltiples espacios por uno solo

                    # Buscar la palabra entre "Seńor(a) " y " Asociado"
                    palabra_antes = "CC"
                    palabra_despues = " Asociado"

                    # Buscar el nombre entre las dos palabras clave
                    resultado = self.buscar_palabra_entre(texto_extraido, palabra_antes, palabra_despues)
                
                    if resultado:
                        nombre_extraido = resultado[0].strip()  # Obtener el nombre extraído
                        print(f"Nombre extraído: {nombre_extraido}")

                        # Crear una nueva ruta para el archivo renombrado
                        nuevo_nombre_pdf = os.path.join(carpeta, f"{nombre_extraido}.pdf")

                        # Renombrar el archivo
                        try:
                            os.rename(archivo_pdf, nuevo_nombre_pdf)
                            
                        except Exception as e:
                            print(f"Error al renombrar el archivo {archivo_pdf}: {e}")
                    else:
                        print("\nNo se encontró ninguna palabra entre las dos palabras clave.")
                else:
                    print("No se ha extraído texto del PDF.")
            else:
                print(f"{archivo} no es un archivo PDF, se omite.")
        messagebox.showinfo("Completado", f"Se han renombrado exitosamente los {contador} archivos")
        self.boton_nombrar.config(state="disabled")
        self.boton_convertir.config(state="normal")

    


    def convertir_pdf_jpg(self):
        # Ruta de la carpeta que contiene los PDFs
        pdf_folder = "C:/Mensajes/habeas_data/pdf" 

        # Ruta de la carpeta donde quieres guardar los JPGs
        jpg_folder = "C:/Mensajes/habeas_data/jpg" 

        # Crear la carpeta JPG si no existe
        if not os.path.exists(jpg_folder):
            os.makedirs(jpg_folder)

        # Crear ventana de "Esperando"
        self.ventana_esperando = tk.Toplevel(self)
        self.ventana_esperando.title("Esperando...")
        self.ventana_esperando.geometry("300x100")
        etiqueta_esperando = tk.Label(self.ventana_esperando, text="Convirtiendo archivos PDF a JPG...\nEspere por favor.")
        etiqueta_esperando.pack(pady=20)
        self.update()  # Esto asegura que la ventana se actualice

        # Definir el factor de escala (aumenta el valor para mejor calidad)
        zoom_factor = 2.0  # Cambia este valor para ajustar la calidad

        # Recorrer todos los archivos PDF en la carpeta
        for pdf_file in os.listdir(pdf_folder):
            if pdf_file.endswith('.pdf'):  # Solo archivos .pdf
                # Ruta completa del archivo PDF
                pdf_path = os.path.join(pdf_folder, pdf_file)

                # Abrir el PDF usando PyMuPDF
                doc = fitz.open(pdf_path)

                # Obtener la primera página del PDF
                page = doc.load_page(0)  # La primera página es la página 0

                # Ajustar la resolución utilizando un factor de zoom (matrix)
                mat = fitz.Matrix(zoom_factor, zoom_factor)  # Establece el zoom en ambos ejes

                # Convertir la página a una imagen (pixmap) con la matriz de zoom aplicada
                pix = page.get_pixmap(matrix=mat)

                # Crear nombre de archivo para la imagen JPG
                jpg_file = f"{os.path.splitext(pdf_file)[0]}.jpg"
                jpg_path = os.path.join(jpg_folder, jpg_file)

                # Guardar la imagen en formato JPG
                pix.save(jpg_path)

                # Puedes imprimir un mensaje en la consola para ver el progreso (opcional)
                print(f"El archivo {pdf_file} ha sido convertido a JPG con alta calidad.")

        # Cerrar la ventana de "Esperando"
        self.ventana_esperando.destroy()

        # Mostrar mensaje de "completado"
        messagebox.showinfo("Completado", "Se han convertido los pdf a jpg.")
        
        # Deshabilitar el botón después de la conversión
        self.boton_convertir.config(state="disabled")
        self.boton_mensaje_whatsapp.config(state="normal")


#---------------Mensajes de whatsapp----------------

    def corregir_numero_telefono(self,numero):
        # Asegurarse de que los números tengan el formato correcto (+ y código de país)
        if not numero.startswith('+'):
            return '+57' + numero  # Agregar el "+" si falta
        return numero

    def tomar_screenshot(self, carpeta_screenshots, cedula):
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

    def cargar_imagen_y_mensaje(self):
    
        archivo_excel = "C:/Mensajes/habeas_data/datos/datos.xlsx" # Ruta al archivo Excel
        carpeta_imagenes = "C:/Mensajes/habeas_data/jpg"  # Carpeta donde están las imágenes
        carpeta_screenshots = "C:/Mensajes/habeas_Data/screenshots"  # Carpeta donde se guardarán las capturas de pantalla
        tiempo_cambio = 10  # Tiempo en segundos para cambiar al siguiente número (puedes ajustarlo según sea necesario)
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
            numero_telefono = self.corregir_numero_telefono(fila['telefono'])  # Asegurarse de que el número es correcto
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
                    pywhatkit.sendwhats_image(numero_telefono, ruta_imagen, "Buen día Sr(a) asociado COEDUCADORES BOYACÁ adjunta NOTIFICACIÓN HABEAS DATA", 10)  # Tiempo de espera de 10 segundos

                    # Capturar la hora de envío
                    hora_envio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    print(f"Mensaje enviado a {numero_telefono} a las {hora_envio}")

                    # Actualizar la columna 'enviarWhatsapp' con la hora de envío
                    df.at[i, 'enviarWhatsapp'] = hora_envio

                    # Espera de 3 segundos después de enviar el mensaje antes de tomar el pantallazo
                    print("Esperando 3 segundos para tomar el pantallazo...")
                    time.sleep(3)  # Espera de 3 segundos para capturar la pantalla

                    # Tomar el pantallazo y guardarlo con el número de cédula como nombre
                    self.tomar_screenshot(carpeta_screenshots, cedula)

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

    


        if not os.path.exists(carpeta_screenshots):
            os.makedirs(carpeta_screenshots)

    def enviar_correo(self):
        
        # Configuración del servidor SMTP de Gmail
        smtp_host = 'smtp.gmail.com'
        smtp_port = 587
        username = 'cobranzascoeducadores@gmail.com'  # Tu correo de Gmail
        password = 'leuz vzyx wqzx ryrg'  # Usa una contraseña de aplicación

        # Cargar el archivo Excel
        archivo_excel = "C:/Mensajes/habeas_data/datos/datos.xlsx"
        df = pd.read_excel(archivo_excel, sheet_name='Datos')

        # Asegurarte de que las columnas tengan un tipo de datos adecuado para aceptar strings
        df['Enviado'] = df['Enviado'].astype(str)
        df['HoraEnvio'] = df['HoraEnvio'].astype(str)

        # Crear ventana de "Esperando"
        self.ventana_esperando = tk.Toplevel(self)
        self.ventana_esperando.title("Esperando...")
        self.ventana_esperando.geometry("300x100")
        etiqueta_esperando = tk.Label(self.ventana_esperando, text="Enviando Correos Electronicos...\nEspere por favor.")
        etiqueta_esperando.pack(pady=20)
        self.update()  # Esto asegura que la ventana se actualice
        
        # Iterar sobre cada fila de la columna 'correo' y 'cedula'
        for i, fila in df.iterrows():
            destinatario = fila['correo']
            cedula = fila['cedula']

            # Ruta del archivo JPG específico para cada cédula
            archivo_path = f'C:/Mensajes/habeas_data/jpg/{cedula}.jpg'

            # Verificar si el archivo existe
            if os.path.exists(archivo_path):
                # Construir el mensaje
                msg = MIMEMultipart('related')
                msg['From'] = username
                msg['To'] = destinatario
                msg['Subject'] = 'NOTIFICACIÓN HABEAS DATA COEDUCADORES'

                cuerpo_mensaje = '''
                <html>
                <body>
                    <p>Estimado,</p>
                    <p>Buen día Sr(a) asociado COEDUCADORES BOYACÁ adjunta NOTIFICACIÓN HABEAS DATA.</p>
                    <img src="cid:comunicado_image" />
                </body>
                </html>
                '''
                msg.attach(MIMEText(cuerpo_mensaje, 'html'))

                # Adjuntar la imagen y asignarle un Content-ID para incrustarla en el HTML
                with open(archivo_path, 'rb') as archivo_adjunto:
                    imagen = MIMEBase('image', 'jpeg')
                    imagen.set_payload(archivo_adjunto.read())
                    encoders.encode_base64(imagen)
                    imagen.add_header('Content-ID', '<comunicado_image>')  # Usamos un ID único
                    imagen.add_header('Content-Disposition', 'inline; filename=' + os.path.basename(archivo_path))
                    msg.attach(imagen)

                # Enviar el correo
                try:
                    with smtplib.SMTP(smtp_host, smtp_port) as servidor:
                        servidor.starttls()  # Inicia la conexión segura
                        servidor.login(username, password)
                        servidor.send_message(msg)
                        print(f"Correo enviado exitosamente a {destinatario}.")

                        # Agregar marca de 'Enviado' y la hora al DataFrame
                        df.loc[i, 'Enviado'] = 'S'
                        df.loc[i, 'HoraEnvio'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                except Exception as e:
                    print(f"Error al enviar el correo a {destinatario}: {e}")
                    df.loc[i, 'Enviado'] = 'Error'
                    df.loc[i, 'HoraEnvio'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            else:
                print(f"Archivo no encontrado para {destinatario} (Cédula: {cedula}).")
                df.loc[i, 'Enviado'] = 'Archivo no encontrado'
                df.loc[i, 'HoraEnvio'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Guardar el DataFrame modificado en el archivo Excel
        df.to_excel(archivo_excel, sheet_name='Datos', index=False)
        print("El archivo Excel ha sido actualizado.")
        messagebox.showinfo("Completado", f"Correos enviados exitosamente")
        
