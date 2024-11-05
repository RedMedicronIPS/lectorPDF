# Instala PyMuPDF si no lo tienes: pip install pymupdf
import fitz  # PyMuPDF
import pandas as pd
import os
import re

# Ruta de la carpeta donde están los PDFs
carpeta_pdf = 'C:/Users/IPS OBRERO/Desktop/Lector pdf/PDF'  # Cambia esto por la ruta de tu carpeta
archivos_pdf = [f for f in os.listdir(carpeta_pdf) if f.endswith('.pdf')]

# Definir las frases delimitadoras
frase_inicio = "Manejo integral según guía de:"
frase_final = "Notas auditor:"

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

    # Usar expresión regular para extraer el texto entre las frases delimitadoras
    patron = re.escape(frase_inicio) + r"(.*?)" + re.escape(frase_final)
    coincidencia = re.search(patron, texto_completo, re.DOTALL)  # re.DOTALL para que "." incluya saltos de línea

    # Si se encuentra coincidencia, agregar el texto extraído; si no, indicar que no hay contenido
    if coincidencia:
        data['Texto Extraido'] = coincidencia.group(1).strip()
    else:
        data['Texto Extraido'] = "Sin contenido entre frases"

    todos_datos.append(data)
    pdf.close()  # Cierra el archivo PDF

# Convertir la lista de diccionarios a un DataFrame
df = pd.DataFrame(todos_datos)

# Guardar el DataFrame en un archivo Excel
df.to_excel('datos_extraidos_filtrados.xlsx', index=False)

print("Datos guardados en 'datos_extraidos_filtrados.xlsx'")
