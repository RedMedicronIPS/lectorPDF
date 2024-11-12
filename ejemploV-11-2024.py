import fitz  # PyMuPDF
import pandas as pd
import os
import re  # Librería para manejar expresiones regulares

# Carpeta y archivos PDF
carpeta_pdf = 'C:/Users/IPS OBRERO/Desktop/Lector pdf/PDF'
archivos_pdf = [f for f in os.listdir(carpeta_pdf) if f.endswith('.pdf')]

todos_datos = []

# Frases para extraer secciones
frases = [
    ("Fecha:", "NÚMERO DE AUTORIZACIÓN:"),
    ("Permiso especial de permanencia", "Número documento de identificación"),
    ("Manejo integral según guía de:", "SERVICIO\nCÓDIGO"),
    ("Lazos", "Notas auditor:")
]

# Función para preprocesar el texto y reemplazar "\n" entre palabras
def preprocesar_texto(texto):
    texto = re.sub(r'([a-zA-Z]+ )\n([a-zAZ]+ )', r'\1 \2', texto)
    return texto

# Función para procesar las secciones en una lista de líneas con reglas específicas
def procesar_seccion(texto):
    texto = preprocesar_texto(texto)
    
    lineas = texto.split('\n')
    resultado = []
    
    for linea in lineas:
        linea = linea.strip()
        
        if not linea:
            continue
        
        if linea.isdigit():
            resultado.append(linea)
        elif linea.isalpha():
            resultado.append(linea)
        elif re.match(r'^[a-zA-Z]+ \d+$', linea):  # Ejemplo: "CALCIO 903810"
            resultado.append(linea)
        elif re.match(r'^\d+ \d+$', linea):  # Ejemplo: "903810 1"
            partes = linea.split()
            resultado.extend(partes)
        else:
            resultado.append(linea)
    
    return resultado

# Recorre los archivos PDF
for archivo in archivos_pdf:
    ruta_pdf = os.path.join(carpeta_pdf, archivo)
    pdf = fitz.open(ruta_pdf)
    data = {'Nombre Archivo': archivo}

    # Extraer texto de cada página
    for pagina_num, pagina in enumerate(pdf):
        texto_pagina = pagina.get_text("text")

        # Extraer secciones basadas en las frases de inicio y final
        for i, (frase_inicio, frase_final) in enumerate(frases):
            inicio = texto_pagina.find(frase_inicio)
            fin = texto_pagina.find(frase_final, inicio + len(frase_inicio))
            
            if inicio != -1 and fin != -1:
                texto_seccion = texto_pagina[inicio + len(frase_inicio):fin].strip()
                
                # Procesar el texto de la sección en una lista de líneas
                lista_seccion = procesar_seccion(texto_seccion)
                
                # Guardar la lista en el diccionario
                data[f"Seccion {i + 1}"] = lista_seccion

        # Ahora procesamos las Sección 3 y Sección 4 (es decir, las secciones extraídas del PDF)
        if 'Seccion 3' in data:
            for item in data['Seccion 3']:  # Recorrer las líneas extraídas de la Sección 3
                # Aquí agregamos los datos de la Sección 3 a cada entrada
                seccion3 = data['Seccion 3']
                entry = {
                    'Nombre Archivo': archivo,
                    'Seccion 1': data.get('Seccion 1', []),
                    'Seccion 2': data.get('Seccion 2', []),
                    'Nombre': item,  # Nombre extraído de la Sección 3
                    'Código': "",  # Se puede extraer también si existe un formato de número
                    'Cantidad': "" # Lo mismo para la cantidad si está presente
                }
                todos_datos.append(entry)
        todos_datos.append(data)
        if 'Seccion 4' in data:
            for item in data['Seccion 4']:  # Recorrer las líneas extraídas de la Sección 4
                # Aquí agregamos los datos de la Sección 4 a cada entrada
                entry = {
                    'Nombre Archivo': archivo,
                    'Seccion 1': data.get('Seccion 1', []),
                    'Seccion 2': data.get('Seccion 2', []),
                    'Nombre': item ,  # Nombre extraído de la Sección 4
                    'Código': "" ,  # Se puede extraer también si existe un formato de número
                    'Cantidad': "" # Lo mismo para la cantidad si está presente
                }
            todos_datos.append(entry)
            

    pdf.close()

# Convertir la lista de diccionarios a un DataFrame
df = pd.DataFrame(todos_datos)

# Guardar el DataFrame en un archivo Excel
df.to_excel('datos_extraidos6.xlsx', index=False, engine='openpyxl')

print("Datos guardados en 'datos_extraidos.xlsx'")
