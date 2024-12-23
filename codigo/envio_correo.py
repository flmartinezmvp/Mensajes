

# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.base import MIMEBase
# from email import encoders
# import os
# import pandas as pd
# from datetime import datetime

# # Configuración del servidor SMTP de Gmail
# # # Configuración del servidor SMTP de Gmail
# smtp_host = 'smtp.gmail.com'
# smtp_port = 587
# username = 'cobranzascoeducadores@gmail.com'  # Tu correo de Gmail
# password = 'leuz vzyx wqzx ryrg'  # Usa una contraseña de aplicación

# # Cargar el archivo Excel
# archivo_excel = 'C:/Mensajes/datos.xlsx'
# df = pd.read_excel(archivo_excel, sheet_name='Datos')

# # Asegurarte de que las columnas tengan un tipo de datos adecuado para aceptar strings
# df['Enviado'] = df['Enviado'].astype(str)
# df['HoraEnvio'] = df['HoraEnvio'].astype(str)


# # Iterar sobre cada fila de la columna 'correo' y 'cedula'
# for i, fila in df.iterrows():
#     destinatario = fila['correo']
#     cedula = fila['cedula']
    
#     # Ruta del archivo PDF específico para cada cédula
#     archivo_path = f'C:/Mensajes/jpg/{cedula}.jpg'
    
#     # Verificar si el archivo existe
#     if os.path.exists(archivo_path):
#         # Construir el mensaje
#         msg = MIMEMultipart()
#         msg['From'] = username
#         msg['To'] = destinatario
#         msg['Subject'] = 'Notificación'
        
#         cuerpo_mensaje = 'Adjunto comunicado.'
#         msg.attach(MIMEText(cuerpo_mensaje, 'plain'))
        
#         # Adjuntar el archivo
#         with open(archivo_path, 'rb') as archivo_adjunto:
#             parte = MIMEBase('application', 'octet-stream')
#             parte.set_payload(archivo_adjunto.read())
#             encoders.encode_base64(parte)
#             parte.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(archivo_path)}')
#             msg.attach(parte)
        
#         # Enviar el correo
#         try:
#             with smtplib.SMTP(smtp_host, smtp_port) as servidor:
#                 servidor.starttls()  # Inicia la conexión segura
#                 servidor.login(username, password)
#                 servidor.send_message(msg)
#                 print(f"Correo enviado exitosamente a {destinatario}.")
                
#                 # Agregar marca de 'Enviado' y la hora al DataFrame
#                 df.loc[i, 'Enviado'] = 'S'
#                 df.loc[i, 'HoraEnvio'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
#         except Exception as e:
#             print(f"Error al enviar el correo a {destinatario}: {e}")
#             df.loc[i, 'Enviado'] = 'Error'
#             df.loc[i, 'HoraEnvio'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#     else:
#         print(f"Archivo no encontrado para {destinatario} (Cédula: {cedula}).")
#         df.loc[i, 'Enviado'] = 'Archivo no encontrado'
#         df.loc[i, 'Hora de Envio'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

# # Guardar el DataFrame modificado en el archivo Excel
# df.to_excel(archivo_excel, sheet_name='Datos', index=False)
# print("El archivo Excel ha sido actualizado.")

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd
from datetime import datetime

# Configuración del servidor SMTP de Gmail
smtp_host = 'smtp.gmail.com'
smtp_port = 587
username = 'cobranzascoeducadores@gmail.com'  # Tu correo de Gmail
password = 'leuz vzyx wqzx ryrg'  # Usa una contraseña de aplicación

# Cargar el archivo Excel
archivo_excel = 'C:/Mensajes/datos.xlsx'
df = pd.read_excel(archivo_excel, sheet_name='Datos')

# Asegurarte de que las columnas tengan un tipo de datos adecuado para aceptar strings
df['Enviado'] = df['Enviado'].astype(str)
df['HoraEnvio'] = df['HoraEnvio'].astype(str)

# Iterar sobre cada fila de la columna 'correo' y 'cedula'
for i, fila in df.iterrows():
    destinatario = fila['correo']
    cedula = fila['cedula']
    
    # Ruta del archivo JPG específico para cada cédula
    archivo_path = f'C:/Mensajes/jpg/{cedula}.jpg'
    
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
            <p>Adjunto comunicado con la imagen.</p>
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
