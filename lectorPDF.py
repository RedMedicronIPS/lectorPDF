# pip install pymupdf
# pip install pandas

import fitz  # PyMuPDF
import pandas as pd
import os
import re

# Ruta de la carpeta donde están los PDFs
carpeta_pdf = 'C:/Users/IPS OBRERO/Desktop/Lector pdf/PDF'  # Cambia esto por la ruta de tu carpeta
archivos_pdf = [f for f in os.listdir(carpeta_pdf) if f.endswith('.pdf')]

# Frases delimitadoras globales y específicas por página
frase_inicio = "Manejo integral según guía de:"
frase_final = "Notas auditor:"
frase_inicio_pag_1 = "Manejo integral según guía de:"
frase_final_pag_1 = "SERVICIO"
frase_inicio_pag_2 = "Lazos"
frase_final_pag_2 = "Notas auditor:"

# Crear una lista para almacenar todos los datos
todos_datos = []

# Procesar cada archivo PDF
for archivo in archivos_pdf:
    ruta_pdf = os.path.join(carpeta_pdf, archivo)
    pdf = fitz.open(ruta_pdf)

    # Crear un diccionario para almacenar los datos del archivo actual
    data = {'Nombre Archivo': archivo}  # Agregar el nombre del archivo

    # Concatenar el texto de todas las páginas
    texto_completo = ""
    for pagina in pdf:
        texto_completo += pagina.get_text("text") + "\n"

    # Usar expresión regular para extraer el texto global entre las frases delimitadoras
    patron_global = re.escape(frase_inicio) + r"(.*?)" + re.escape(frase_final)
    coincidencia_global = re.search(patron_global, texto_completo, re.DOTALL)

    # Si no encuentra en el texto global, buscar en las páginas específicas
    if coincidencia_global:
        data['Texto Extraido'] = coincidencia_global.group(1).strip()
    else:
        # Buscar en la primera página usando las frases delimitadoras específicas
        texto_pag_1 = pdf[0].get_text("text")
        inicio_idx_pag_1 = texto_pag_1.find(frase_inicio_pag_1)
        final_idx_pag_1 = texto_pag_1.find(frase_final_pag_1)
        
        # Buscar en la segunda página usando las frases delimitadoras específicas
        texto_pag_2 = pdf[1].get_text("text") if len(pdf) > 1 else ""
        inicio_idx_pag_2 = texto_pag_2.find(frase_inicio_pag_2)
        final_idx_pag_2 = texto_pag_2.find(frase_final_pag_2)
        
        # Extraer el texto entre delimitadores específicos por página
        texto_extraido = ""
        if inicio_idx_pag_1 != -1 and final_idx_pag_1 != -1:
            texto_extraido += texto_pag_1[inicio_idx_pag_1 + len(frase_inicio_pag_1):final_idx_pag_1].strip() + "\n"
        
        if inicio_idx_pag_2 != -1 and final_idx_pag_2 != -1:
            texto_extraido += texto_pag_2[inicio_idx_pag_2 + len(frase_inicio_pag_2):final_idx_pag_2].strip() + "\n"

        data['Texto Extraido'] = texto_extraido if texto_extraido else "Sin contenido entre frases específicas"

    todos_datos.append(data)
    pdf.close()  # Cierra el archivo PDF

# Convertir la lista de diccionarios a un DataFrame
df = pd.DataFrame(todos_datos)

# Guardar el DataFrame en un archivo Excel
df.to_excel('datos_extraidos_filtrados.xlsx', index=False)

print("Datos guardados en 'datos_extraidos_filtrados.xlsx'")
