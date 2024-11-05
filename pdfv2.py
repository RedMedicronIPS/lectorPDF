import fitz  # PyMuPDF
import pandas as pd
import os

# Ruta de la carpeta donde están los PDFs
carpeta_pdf = 'C:/Users/IPS OBRERO/Desktop/Lector pdf/PDF'  # Cambia esto por la ruta de tu carpeta
archivos_pdf = [f for f in os.listdir(carpeta_pdf) if f.endswith('.pdf')]

# Crear una lista para almacenar todos los datos
todos_datos = []

# Procesar cada archivo PDF
for archivo in archivos_pdf:
    ruta_pdf = os.path.join(carpeta_pdf, archivo)
    pdf = fitz.open(ruta_pdf)

    # Crear un diccionario para almacenar los datos del archivo actual
    data = {'Nombre Archivo': archivo}  # Agregar el nombre del archivo

    # Extraer texto de cada página
    texto_completo = ""
    for pagina_num, pagina in enumerate(pdf):
        texto_completo += f"--- Página {pagina_num + 1} ---\n"
        texto_completo += pagina.get_text("text") + "\n"  # Extrae texto como texto plano

    data['Texto Completo'] = texto_completo
    todos_datos.append(data)

    pdf.close()  # Cierra el archivo PDF

# Convertir la lista de diccionarios a un DataFrame
df = pd.DataFrame(todos_datos)

# Guardar el DataFrame en un archivo Excel
df.to_excel('datos_extraidos_completos.xlsx', index=False)

print("Datos guardados en 'datos_extraidos_completos.xlsx'")
