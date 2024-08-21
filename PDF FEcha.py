import fitz  # PyMuPDF
import os
from openpyxl import Workbook
import re
from datetime import datetime

# Ruta al directorio que contiene los archivos PDF
pdf_directory = 'C:/Users/ersal/OneDrive/Desktop/AUT_PYT'
# Ruta del archivo Excel de salida
excel_path = 'C:/Users/ersal/OneDrive/Desktop/AUT_PYT/archivo_salida_completo.xlsx'
# Texto completo que estás buscando
search_text = 'Contribuyente : NO INSCRIPTO'
certificado_text = 'Certificado de no Retención y no Percepción'

# Expresión regular para encontrar fechas en formato dd-mm-yyyy
date_pattern = r'\b(\d{2}-\d{2}-\d{4})\b'

# Crear un nuevo libro de Excel
wb = Workbook()
ws = wb.active
ws.title = 'Datos PDF'
ws.append(['Nombre del Archivo', 'Texto Encontrado', 'Primera Fecha', 'Segunda Fecha'])  # Encabezados de las columnas

# Función para extraer hasta dos fechas del texto
def extract_dates(text):
    matches = re.findall(date_pattern, text)
    if len(matches) > 1:
        return matches[0], matches[1]
    elif len(matches) == 1:
        return matches[0], ''
    else:
        return '', ''

# Obtener la fecha actual
today = datetime.now()

# Recorrer todos los archivos PDF en el directorio
for filename in os.listdir(pdf_directory):
    if filename.lower().endswith('.pdf'):
        pdf_path = os.path.join(pdf_directory, filename)
        doc = fitz.open(pdf_path)
        
        found_text = 'No encontrado'
        first_date = ''
        second_date = ''
        
        # Buscar el texto en cada página del PDF
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text()
            
            # Extraer las fechas del texto
            date1, date2 = extract_dates(text)
            
            if date1:
                first_date = date1
            if date2:
                second_date = date2
            
            # Verificar el texto completo "Contribuyente : NO INSCRIPTO"
            if search_text in text:
                found_text = search_text
                break
            elif certificado_text in text:
                found_text = 'CORRESPONDE 0.75'
        
        # Verificar la fecha
        if second_date:
            date_format = '%d-%m-%Y'
            try:
                date_object = datetime.strptime(second_date, date_format)
                if date_object > today:
                    found_text = 'CORRESPONDE 0.75'
            except ValueError:
                # Fecha en formato incorrecto
                pass

        # Guardar el resultado en el archivo Excel
        ws.append([filename, found_text, first_date, second_date])
        
        # Cerrar el documento PDF
        doc.close()

# Guardar el archivo Excel
wb.save(excel_path)

print('Proceso completado. Datos guardados en:', excel_path)
