import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pymupdf #PyMuPDF 
import os
from openpyxl import Workbook
import re
from datetime import datetime, timedelta

def select_directory(title):
    """Abre un cuadro de diálogo para seleccionar un directorio y devuelve la ruta seleccionada."""
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    directory = filedialog.askdirectory(title=title)
    if not directory:
        messagebox.showwarning("Advertencia", "No se seleccionó ningún directorio.")
    return directory

def extract_dates(text):
    matches = re.findall(date_pattern, text)
    if len(matches) > 1:
        return matches[0], matches[1]
    elif len(matches) == 1:
        return matches[0], ''
    else:
        return '', ''

def extract_cuit(text, cuit_prefix):
    cuit_start = text.find(cuit_prefix)
    if cuit_start != -1:
        cuit_start += len(cuit_prefix)
        cuit_end = text.find('\n', cuit_start)
        if cuit_end == -1:
            cuit_end = len(text)
        extracted_cuit = text[cuit_start:cuit_end].strip()
        if re.match(r'^\d{11}$', extracted_cuit):
            return extracted_cuit
    return ''

# Función principal
def main():
    global pdf_directory, excel_path, date_pattern, cuit_prefix

    # Selecciona los directorios
    pdf_directory = select_directory("Selecciona el directorio que contiene los archivos PDF")
    if not pdf_directory:
        return
    excel_path = select_directory("Selecciona el directorio para guardar el archivo de salida")
    if not excel_path:
        return
    excel_path = os.path.join(excel_path, 'archivo_salida_completo.xlsx')

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
    ws.append(['Nombre del Archivo', 'CUIT', 'Valor', 'Texto Encontrado', 'Primera Fecha', 'Segunda Fecha', 'Alerta'])  # Encabezados de las columnas

    # Obtener la fecha actual
    today = datetime.now()

    # Recorrer todos los archivos PDF en el directorio
    for filename in os.listdir(pdf_directory):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(pdf_directory, filename)
            doc = pymupdf.open(pdf_path)
            
            found_text = 'No encontrado'
            cuit_value = ''
            first_date = ''
            second_date = ''
            value = '0.00%'  # Valor por defecto
            alerta = ''  # Valor por defecto para la columna "Alerta"
            
            # Extraer el CUIT de la primera página
            if len(doc) > 0:
                first_page = doc.load_page(0)
                text = first_page.get_text()
                cuit_value = extract_cuit(text, cuit_prefix)
            
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
            
            # Verificar la fecha para alerta (si la segunda fecha es en los próximos 5 días)
            if second_date:
                date_format = '%d-%m-%Y'
                try:
                    date_object = datetime.strptime(second_date, date_format)
                    # Compara si la segunda fecha está dentro de los próximos 5 días a partir de hoy
                    if today <= date_object <= (today + timedelta(days=5)):
                        alerta = 'ALERTA'
                    elif date_object > today + timedelta(days=5):
                        found_text = 'NO Aplicar 0.75'
                        value = '0.00%'
                except ValueError:
                    # Fecha en formato incorrecto
                    pass

            # Verificar el CUIT extraído
            print(f"Extracted CUIT: '{cuit_value}'")  # Imprimir el CUIT extraído para depuración

            # Guardar el resultado en el archivo Excel, incluyendo la columna "Alerta"
            ws.append([filename, cuit_value, value, found_text, first_date, second_date, alerta])
            
            # Cerrar el documento PDF
            doc.close()

    # Guardar el archivo Excel
    wb.save(excel_path)

    print('Proceso completado. Datos guardados en:', excel_path)

if __name__ == "__main__":
    main()
