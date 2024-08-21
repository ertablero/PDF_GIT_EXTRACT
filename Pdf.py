import fitz  # PyMuPDF
import os
from openpyxl import Workbook

# Ruta al directorio que contiene los archivos PDF
pdf_directory = 'C:/Users/ersal/OneDrive/Desktop/AUT_PYT'
# Ruta del archivo Excel de salida
excel_path = 'C:/Users/ersal/OneDrive/Desktop/AUT_PYT/archivo_salida.xlsx'
# Texto que estás buscando
search_text = 'INSCRIPTO'

# Crear un nuevo libro de Excel
wb = Workbook()
ws = wb.active
ws.title = 'Datos PDF'
ws.append(['Nombre del Archivo', 'Texto Encontrado'])  # Encabezados de las columnas

# Recorrer todos los archivos PDF en el directorio
for filename in os.listdir(pdf_directory):
    if filename.lower().endswith('.pdf'):
        pdf_path = os.path.join(pdf_directory, filename)
        doc = fitz.open(pdf_path)
        found_text = None
        
        # Buscar el texto en cada página del PDF
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text()
            if search_text in text:
                found_text = search_text
                break
        
        # Guardar el resultado en el archivo Excel
        ws.append([filename, found_text if found_text else 'No encontrado'])
        
        # Cerrar el documento PDF
        doc.close()

# Guardar el archivo Excel
wb.save(excel_path)

print('Proceso completado. Datos guardados en:', excel_path)