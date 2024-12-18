# import os
# import fitz  # PyMuPDF

# # Ruta de la carpeta que contiene los PDFs
# pdf_folder = 'C:/Users/flmartinez/Desktop/Mensajes/pdf'

# # Ruta de la carpeta donde quieres guardar los JPGs
# jpg_folder = 'C:/Users/flmartinez/Desktop/Mensajes/jpg'

# # Crear la carpeta JPG si no existe
# if not os.path.exists(jpg_folder):
#     os.makedirs(jpg_folder)

# # Recorrer todos los archivos PDF en la carpeta
# for pdf_file in os.listdir(pdf_folder):
#     if pdf_file.endswith('.pdf'):  # Solo archivos .pdf
#         # Ruta completa del archivo PDF
#         pdf_path = os.path.join(pdf_folder, pdf_file)
        
#         # Abrir el PDF usando PyMuPDF
#         doc = fitz.open(pdf_path)
        
#         # Obtener la primera página del PDF (como solo tienes 1 página por archivo, solo trabajamos con la primera)
#         page = doc.load_page(0)  # La primera página es la página 0
        
#         # Convertir la página a una imagen (pixmap)
#         pix = page.get_pixmap()

#         # Crear nombre de archivo para la imagen JPG
#         jpg_file = f"{os.path.splitext(pdf_file)[0]}.jpg"
#         jpg_path = os.path.join(jpg_folder, jpg_file)
        
#         # Guardar la imagen en formato JPG
#         pix.save(jpg_path)

#         print(f"El archivo {pdf_file} ha sido convertido a JPG.")


import os
import fitz  # PyMuPDF

# Ruta de la carpeta que contiene los PDFs
pdf_folder = 'C:/Mensajes/pdf'

# Ruta de la carpeta donde quieres guardar los JPGs
jpg_folder = 'C:/Mensajes/jpg'

# Crear la carpeta JPG si no existe
if not os.path.exists(jpg_folder):
    os.makedirs(jpg_folder)

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

        print(f"El archivo {pdf_file} ha sido convertido a JPG con alta calidad.")
