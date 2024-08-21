import fitz  # PyMuPDF
import os
from openpyxl import Workbook
import re
from datetime import datetime

# Ruta al directorio que contiene los archivos PDF
pdf_directory = 'C:/Users/ersal/OneDrive/Desktop/AUT_PYT/Consulta_Alic_P_R_082024'
# Ruta del archivo Excel de salida
excel_path = 'C:/Users/ersal/OneDrive/Desktop/AUT_PYT/archivo_salida_completo.xlsx'

# Texto completo que estás buscando
search_text = 'Contribuyente : NO INSCRIPTO'
certificado_text = 'Certificado de no Retención y no Percepción'
cuit_prefix = 'CUIT : '

# Expresión regular para encontrar fechas en formato dd-mm-yyyy
date_pattern = r'\b(\d{2}-\d{2}-\d{4})\b'

# Crear un nuevo libro de Excel
wb = Workbook()
ws = wb.active
ws.title = 'Datos PDF'
ws.append(['Nombre del Archivo', 'CUIT', 'Valor', 'Texto Encontrado', 'Primera Fecha', 'Segunda Fecha'])  # Encabezados de las columnas

# Función para extraer hasta dos fechas del texto
def extract_dates(text):
    matches = re.findall(date_pattern, text)
    if len(matches) > 1:
        return matches[0], matches[1]
    elif len(matches) == 1:
        return matches[0], ''
    else:
        return '', ''

# Función para extraer el CUIT
def extract_cuit(text):
    cuit_start = text.find(cuit_prefix)
    if cuit_start != -1:
        cuit_start += len(cuit_prefix)
        cuit_end = text.find('\n', cuit_start)
        if cuit_end == -1:
            cuit_end = len(text)
        extracted_cuit = text[cuit_start:cuit_end].strip()
        # Verificar si el CUIT es válido (11 dígitos)
        if re.match(r'^\d{11}$', extracted_cuit):
            return extracted_cuit
    return ''

# Obtener la fecha actual
today = datetime.now()

# Recorrer todos los archivos PDF en el directorio
for filename in os.listdir(pdf_directory):
    if filename.lower().endswith('.pdf'):
        pdf_path = os.path.join(pdf_directory, filename)
        doc = fitz.open(pdf_path)
        
        found_text = 'No encontrado'
        cuit_value = ''
        first_date = ''
        second_date = ''
        value = '0.00%'  # Valor por defecto
        
        # Extraer el CUIT de la primera página
        if len(doc) > 0:
            first_page = doc.load_page(0)
            text = first_page.get_text()
            cuit_value = extract_cuit(text)
        
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
                value = '0.00%'
                break
            elif certificado_text in text:
                found_text = 'NO Aplicar 0.75'
                value = '0.00%'
            elif text.count('NO POSEE') >= 2:
                found_text = 'Aplicar 0.75'
                value = '0.75%'
        
        # Verificar la fecha
        if second_date:
            date_format = '%d-%m-%Y'
            try:
                date_object = datetime.strptime(second_date, date_format)
                if date_object > today:
                    found_text = 'NO Aplicar 0.75'
                    value = '0.00%'
            except ValueError:
                # Fecha en formato incorrecto
                pass

        # Verificar el CUIT extraído
        print(f"Extracted CUIT: '{cuit_value}'")  # Imprimir el CUIT extraído para depuración

        # Guardar el resultado en el archivo Excel
        ws.append([filename, cuit_value, value, found_text, first_date, second_date])
        
        # Cerrar el documento PDF
        doc.close()

# Guardar el archivo Excel
wb.save(excel_path)

print('Proceso completado. Datos guardados en:', excel_path)
