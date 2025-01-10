import sys
import pdfplumber
import re
import os

# Establecer la codificación de la consola para UTF-8
sys.stdout.reconfigure(encoding='utf-8')

# Función para extraer el texto de todas las páginas de un PDF
def extraer_todo_texto_pdf(archivo_pdf):
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
def buscar_palabra_entre(texto, palabra_antes, palabra_despues):
    # Limpiar el texto eliminando saltos de línea y espacios extras
    texto_limpio = texto.replace("\n", " ")  # Eliminar saltos de línea
    texto_limpio = re.sub(r'\s+', ' ', texto_limpio)  # Reemplazar múltiples espacios por uno solo
    
    # Expresión regular para buscar la palabra entre las dos palabras clave
    patron = r'(?<=\b' + re.escape(palabra_antes) + r'\b)(.*?)(?=\b' + re.escape(palabra_despues) + r'\b)'
    
    # Buscar el texto entre las dos palabras clave
    resultado = re.findall(patron, texto_limpio, re.DOTALL)
    
    return resultado

# Función para renombrar todos los archivos PDF en una carpeta
def renombrar_archivos_en_carpeta(carpeta):
    # Obtener la lista de archivos en la carpeta
    archivos = os.listdir(carpeta)
    
    for archivo in archivos:
        # Verificar si el archivo es un archivo PDF
        if archivo.lower().endswith(".pdf"):
            archivo_pdf = os.path.join(carpeta, archivo)  # Ruta completa al archivo PDF
            print(f"\nProcesando el archivo: {archivo_pdf}")
            
            # Extraer el texto del archivo PDF
            texto_extraido = extraer_todo_texto_pdf(archivo_pdf)
            
            if texto_extraido:
                # Limpiar el texto eliminando saltos de línea y espacios adicionales
                texto_limpio = texto_extraido.replace("\n", " ")  # Eliminar saltos de línea
                texto_limpio = re.sub(r'\s+', ' ', texto_limpio)  # Reemplazar múltiples espacios por uno solo
                
                # Buscar la palabra entre "Seńor(a) " y " Asociado"
                palabra_antes = "CC"
                palabra_despues = " Asociado"
                
                # Buscar el nombre entre las dos palabras clave
                resultado = buscar_palabra_entre(texto_extraido, palabra_antes, palabra_despues)
                
                if resultado:
                    nombre_extraido = resultado[0].strip()  # Obtener el nombre extraído
                    print(f"Nombre extraído: {nombre_extraido}")
                    
                    # Crear una nueva ruta para el archivo renombrado
                    nuevo_nombre_pdf = os.path.join(carpeta, f"{nombre_extraido}.pdf")
                    
                    # Renombrar el archivo
                    try:
                        os.rename(archivo_pdf, nuevo_nombre_pdf)
                        print(f"Archivo renombrado exitosamente a: {nuevo_nombre_pdf}")
                    except Exception as e:
                        print(f"Error al renombrar el archivo {archivo_pdf}: {e}")
                else:
                    print("\nNo se encontró ninguna palabra entre las dos palabras clave.")
            else:
                print("No se ha extraído texto del PDF.")
        else:
            print(f"{archivo} no es un archivo PDF, se omite.")

# Ruta a la carpeta "pdf"
carpeta_divididos = "C:/Mensajes/pdf"

# Renombrar los archivos en la carpeta
renombrar_archivos_en_carpeta(carpeta_divididos)
