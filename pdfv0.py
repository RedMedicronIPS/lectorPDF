#edision 12 noviembre 2024
import fitz  # PyMuPDF
import pandas as pd
import os
import re 
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


carpeta_pdf = 'C:/Users/IPS OBRERO/Desktop/Lector pdf/PDF' 
archivos_pdf = [f for f in os.listdir(carpeta_pdf) if f.endswith('.pdf')]

todos_datos = []

frases = [
    ("Fecha:", "NÚMERO DE AUTORIZACIÓN:"),
    ("Permiso especial de permanencia", "Número documento de identificación"),
    ("Manejo integral según guía de:", "SERVICIO\nCÓDIGO"),
    ("Lazos", "Notas auditor:")

]

# Función para preprocesar el texto y reemplazar "\n" entre palabras
def preprocesar_texto(texto):
    # Reemplaza "\n" con un espacio solo si está entre dos segmentos de texto
    texto = re.sub(r'([a-zA-Z]+)\n([a-zA-Z]+)', r'\1 \2', texto)
    return texto

# Función para procesar las secciones en una lista de líneas con reglas específicas
def procesar_seccion(texto):
    # Preprocesar texto para unir casos de "texto\ntexto"
    texto = preprocesar_texto(texto)
    
    lineas = texto.split('\n')
    resultado = []
    
    for linea in lineas:
        linea = linea.strip()
        
        # Ignorar líneas vacías
        if not linea:
            continue
        
        # Caso 1: Separar si una línea es solo texto o solo números
        elif linea.isdigit():
            resultado.append(linea)
        elif linea.isalpha():
            resultado.append(linea)
        
        # Caso 2: Separar si la línea tiene patrón "texto número" o "número texto"
        elif re.match(r'^[a-zA-Z]+ \d+$', linea):  # Ejemplo "CALCIO 903810"
            resultado.append(linea)
        
        # Caso 3: Separar si la línea tiene un patrón mixto "número número"
        elif re.match(r'^\d+ \d+$', linea):  # Ejemplo "903810 1"
            partes = linea.split()
            resultado.extend(partes)  # Agregar cada número como una línea separada
        
        # Agregar línea tal cual si no cumple las condiciones anteriores
        else:
            resultado.append(linea)
    
    return resultado

# Recorre los archivos PDF
for archivo in archivos_pdf:
    ruta_pdf = os.path.join(carpeta_pdf, archivo)
    pdf = fitz.open(ruta_pdf)

    # Crear un diccionario para almacenar los datos del archivo actual
    data = {'Nombre Archivo': archivo} 

    # Extraer texto de cada página
    for pagina_num, pagina in enumerate(pdf):
        texto_pagina = pagina.get_text("text")

        # Extraer secciones basadas en las frases de inicio y final
        for i, (frase_inicio, frase_final) in enumerate(frases):
            inicio = texto_pagina.find(frase_inicio)
            fin = texto_pagina.find(frase_final, inicio + len(frase_inicio))
            
            # Si ambas frases se encuentran en la página, extraer el texto entre ellas
            if inicio != -1 and fin != -1:
                texto_seccion = texto_pagina[inicio + len(frase_inicio):fin].strip()
                
                # Procesar el texto de la sección en una lista de líneas
                lista_seccion = procesar_seccion(texto_seccion)
                
                # Guardar la lista en el diccionario en lugar del texto original
                data[f"Seccion {i + 1}"] = lista_seccion
                # Ahora procesamos las Sección 3 y Sección 4 (es decir, las secciones extraídas del PDF)
            if 'Seccion 3' in data:
                seccion_3 = data['Seccion 3']
                # Recorremos la lista de 'Seccion 3' en grupos de tres
                for i in range(0, len(seccion_3), 3):
                    nombre = seccion_3[i]         
                    codigo = seccion_3[i + 1]      
                    cantidad = seccion_3[i + 2]
                    autorizacion = str(data.get('Seccion 1', '')).replace("'", '').replace('[', '').replace(']', '').strip()
                    numero_documento = str(data.get('Seccion 2', '')).replace("'", '').replace('[', '').replace(']', '').strip()

                    
                    # Crear la entrada con los datos deseados
                    entry = {
                        'Autorizacion': autorizacion,
                        'Numero Documento': numero_documento,
                        'Nombre': nombre,
                        'Codigo': codigo,
                        'Cantidad': cantidad
                    }
                    todos_datos.append(entry)
                
            if 'Seccion 4' in data:
                seccion_4 = data['Seccion 4']
                
                # Asegúrate de que `seccion_4` tiene un múltiplo de 3 elementos para evitar errores de índice
                if len(seccion_4) % 3 == 0:
                    for i in range(0, len(seccion_4), 3):
                        try:
                            nombre = seccion_4[i]
                            codigo = seccion_4[i + 1]
                            cantidad = seccion_4[i + 2]
                            
                            # Extraer los datos de autorización y documento
                            autorizacion = str(data.get('Seccion 1', '')).replace("'", '').replace('[', '').replace(']', '').strip()
                            numero_documento = str(data.get('Seccion 2', '')).replace("'", '').replace('[', '').replace(']', '').strip()

                            # Crear la entrada con los datos deseados
                            entry = {
                                'Autorizacion': autorizacion,
                                'Numero Documento': numero_documento,
                                'Nombre': nombre,
                                'Codigo': codigo,
                                'Cantidad': cantidad
                            }
                            todos_datos.append(entry)

                        except IndexError:
                            print(f"Error al extraer los datos en el archivo {archivo}, sección 4 no tiene suficientes elementos.")
                else:
                    print(f"Advertencia: La Sección 4 del archivo {archivo} no tiene una cantidad de elementos múltiplo de 3.")

    # Agregar el diccionario con las secciones procesadas a la lista general
    pdf.close()


# Convertir la lista de diccionarios a un DataFrame
df = pd.DataFrame(todos_datos)

# Guardar el DataFrame en un archivo CSV
df.to_csv('datos_extraidos_3.csv', index=False)

print("Datos guardados en 'datos_extraidos.txt'")


