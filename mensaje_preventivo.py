import pdfplumber
import pandas as pd

# Ruta del archivo PDF
pdf_path = 'COBRANZA_PREVENTIVA.pdf'

# Función para extraer los datos después de "Id.Deudor"
def extraer_datos_desde_pdf(pdf_path):
    extracted_data = []  # Lista para almacenar los datos extraídos

    with pdfplumber.open(pdf_path) as pdf:
        # Iterar por cada página
        for page in pdf.pages:
            # Extraer texto de la página
            text = page.extract_text()
            
            # Buscar la posición de "Id.Deudor"
            if "Id.Deudor" in text:
                # Encontrar el índice donde comienza "Id.Deudor"
                start_index = text.index("Id.Deudor")
                
                # Extraer todo el texto después de "Id.Deudor"
                data_after = text[start_index:]
                
                # Dividir el texto extraído en líneas
                lines = data_after.split('\n')
                
                # Iterar sobre las líneas para extraer los datos (asegurarse de que cada valor sea tratado como cadena)
                for line in lines:
                    # Aquí se asume que los datos están separados por espacios; si usan otro delimitador, ajusta esto
                    row = [f"'{str(item).strip()}" if item.isdigit() and len(item) > 1 else str(item).strip() for item in line.split()]
                    
                    if row:  # Asegurarse de que la línea no esté vacía
                        extracted_data.append(row)

    # Convertir los datos extraídos a un DataFrame de pandas
    df = pd.DataFrame(extracted_data)

    # Guardar el DataFrame en un archivo Excel con codificación UTF-8
    df.to_excel('datos_extraidos.xlsx', index=False, engine='openpyxl')

    print("Los datos se han guardado en 'datos_extraidos.xlsx'.")

# Llamar a la función para extraer los datos y exportarlos a Excel
extraer_datos_desde_pdf(pdf_path)
