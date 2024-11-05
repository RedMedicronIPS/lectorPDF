import pdfplumber
import os
import pandas as pd
import re

# Ruta de la carpeta donde están los PDFs
carpeta_pdf = 'C:/Users/IPS OBRERO/Desktop/Lector pdf/PDF'
archivos_pdf = [f for f in os.listdir(carpeta_pdf) if f.endswith('.pdf')]

# Definir palabras de inicio y final de la tabla


# Crear una lista para almacenar todos los datos
todos_datos = []

# Procesar cada archivo PDF
for archivo in archivos_pdf:
    ruta_pdf = os.path.join(carpeta_pdf, archivo)

    with pdfplumber.open(ruta_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()

            # Buscar las posiciones de las frases de inicio y final
            inicio = texto.find(frase_inicio)
            final = texto.find(frase_final)

            if inicio != -1 and final != -1 and inicio < final:
                # Extraer el texto entre las frases de inicio y final
                texto_interes = texto[inicio + len(frase_inicio):final].strip()
                
                # Dividir el texto en líneas
                lineas = texto_interes.split('\n')

                # Filtrar líneas que tengan un formato similar a "SERVICIO  CÓDIGO  CANTIDAD"
                for linea in lineas:
                    # Usar una expresión regular para identificar filas con tres columnas
                    match = re.match(r"(.+?)\s+(\S+)\s+(\d+)", linea)
                    if match:
                        servicio = match.group(1).strip()
                        codigo = match.group(2).strip()
                        cantidad = match.group(3).strip()

                        todos_datos.append({
                            'Archivo': archivo,
                            'SERVICIO': servicio,
                            'CÓDIGO': codigo,
                            'CANTIDAD': cantidad
                        })

# Convertir la lista de diccionarios a un DataFrame
df = pd.DataFrame(todos_datos)

# Guardar los datos en un archivo Excel
df.to_excel('datos_tablas_extraidos.xlsx', index=False)
print("Datos guardados en 'datos_tablas_extraidos.xlsx'")
