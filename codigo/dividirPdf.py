import os
import PyPDF2

def dividir_pdf(archivo_pdf, num_paginas_por_archivo, ruta_destino):
    # Verificar si la carpeta de destino existe, si no, crearla
    os.makedirs(ruta_destino, exist_ok=True)
    
    # Abrir el archivo PDF
    with open(archivo_pdf, 'rb') as archivo:
        lector_pdf = PyPDF2.PdfReader(archivo)
        total_paginas = len(lector_pdf.pages)
        
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

# Llamar a la función para dividir el PDF y guardar en una ruta específica
ruta_destino = "C:/Mensajes/pdf/"  # La carpeta donde se guardarán los archivos
dividir_pdf("C:/Mensajes/archivo/archivo.pdf", 1, ruta_destino)  # Divide en partes de 1 página
