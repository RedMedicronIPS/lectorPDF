#version 10 de diciembre 2024 07:54
import fitz  # PyMuPDF
import pandas as pd
import os
import re 
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Lista para almacenar los datos extraídos
todos_datos = []

# Frases de inicio y fin para extraer secciones específicas
frases_pg1 = [
    ("Fecha:", "NÚMERO DE AUTORIZACIÓN:"),
    ("Permiso especial de permanencia", "Número documento de identificación"),
    ("Manejo integral según guía de:", "SERVICIO\nCÓDIGO"),
]
frases_pg2 = [
    ("Lazos", "Notas auditor:")
]

# Variables globales para mantener los valores de Sección 1 y 2
autorizacion_global = ""
numero_documento_global = ""

# Función para preprocesar el texto y reemplazar "\n" entre palabras
def preprocesar_texto(texto):
    texto = re.sub(r'([a-zA-Z]+)\n([a-zA-Z]+)', r'\1 \2', texto)
    # Reemplazar ')' seguido de '\n' por un espacio
    texto = re.sub(r'\)\n', ') ', texto)
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
        elif linea.isdigit() or linea.isalpha():
            resultado.append(linea)
        elif re.match(r'^[a-zA-Z]+ \d+$', linea):  # Ejemplo "CALCIO 903810"
            resultado.append(linea)
        elif re.match(r'^\d+ \d+$', linea):  # Ejemplo "903810 1"
            partes = linea.split()
            resultado.extend(partes)
        else:
            resultado.append(linea)

    return resultado

# Función para extraer datos de los archivos PDF
def extraer_datos():
    if not carpeta_pdf:
        messagebox.showwarning("Error", "Debe seleccionar una carpeta primero.")
        return

    archivos_pdf = [f for f in os.listdir(carpeta_pdf) if f.endswith('.pdf')]
    barra_progreso.start()

    for archivo in archivos_pdf:
        ruta_pdf = os.path.join(carpeta_pdf, archivo)
        try:
            pdf = fitz.open(ruta_pdf)
        except Exception as e:
            print(f"Error al abrir {archivo}: {e}")
            continue

        for pagina_num, pagina in enumerate(pdf):
            try:
                texto_pagina = pagina.get_text("text")
                if not texto_pagina.strip():
                    print(f"No hay texto en {archivo}, página {pagina_num + 1}")
                    continue
                #print(texto_pagina)
                #print(pagina)

                data = {'Nombre Archivo': archivo}
                
                # Extraer secciones usando expresiones regulares
                # Seleccionar frases según la página
                if pagina_num == 0:  # Primera página
                    frases_a_usar = frases_pg1
                elif pagina_num == 1:  # Segunda página
                    frases_a_usar = frases_pg2
                elif pagina_num == 2:  # Segunda página
                    frases_a_usar = frases_pg2
                else:
                    continue  # Omitir otras páginas si es necesario

                # Extraer secciones usando expresiones regulares
                for i, (frase_inicio, frase_final) in enumerate(frases_a_usar):
                    patron = re.escape(frase_inicio) + r'(.*?)' + re.escape(frase_final)
                    coincidencias = re.findall(patron, texto_pagina, re.DOTALL)

                    if coincidencias:
                        texto_seccion = coincidencias[0].strip()
                        lista_seccion = procesar_seccion(texto_seccion)

                        if pagina_num == 1:  
                            data["Seccion 4"] = lista_seccion
                        else:  # Página 1 -> Secciones normales (Seccion 1, 2, 3...)
                            data[f"Seccion {i + 1}"] = lista_seccion

                #print(data)
                # Procesar Sección 3 si existe
                if 'Seccion 3' in data:
                    seccion_3 = data['Seccion 3']
                    autorizacion = str(data.get('Seccion 1', '')).replace("'", '').replace('[', '').replace(']', '').strip()
                    numero_documento = str(data.get('Seccion 2', '')).replace("'", '').replace('[', '').replace(']', '').strip()
                    if len(seccion_3) % 3 == 0:
                        for i in range(0, len(seccion_3), 3):
                            nombre = seccion_3[i]
                            codigo = seccion_3[i + 1]
                            cantidad = seccion_3[i + 2]

                            entry = {
                                'Autorizacion': autorizacion,
                                'Numero Documento': numero_documento,
                                'Nombre': nombre,
                                'Codigo': codigo,
                                'Cantidad': cantidad
                            }
                            todos_datos.append(entry)
                    else:
                        print(f"Advertencia: Sección 3 en {archivo} tiene datos incompletos.")

                # Procesar Sección 4 si existe
                if 'Seccion 4' in data:
                    seccion_4 = data['Seccion 4']

                    if len(seccion_4) % 3 == 0:
                        for i in range(0, len(seccion_4), 3):
                            nombre = seccion_4[i]
                            codigo = seccion_4[i + 1]
                            cantidad = seccion_4[i + 2]

                            entry = {
                                'Autorizacion': autorizacion,
                                'Numero Documento': numero_documento,
                                'Nombre': nombre,
                                'Codigo': codigo,
                                'Cantidad': cantidad
                            }
                            todos_datos.append(entry)
                    else:
                        print(f"Advertencia: Sección 4 en {archivo} tiene datos incompletos.")

            except Exception as e:
                print(f"Error al procesar {archivo}, página {pagina_num + 1}: {e}")

        pdf.close()
    
    df = pd.DataFrame(todos_datos)
    archivo_csv = os.path.join(carpeta_pdf, 'datos_extraidos.csv')
    df.to_csv(archivo_csv, index=False)
    messagebox.showinfo("Proceso Completo", f"Datos guardados en '{archivo_csv}'")
    label_estado.config(text="Extracción completada")
    barra_progreso.stop()

# Función para seleccionar la carpeta
def seleccionar_carpeta():
    global carpeta_pdf
    carpeta_pdf = filedialog.askdirectory()
    if carpeta_pdf:
        label_estado.config(text=f"Carpeta seleccionada: {carpeta_pdf}")

# Configuración de la interfaz gráfica con tkinter
root = tk.Tk()
root.title("Extractor de Datos de PDF")
root.geometry("500x300")
root.config(bg="#f0f0f0")

# Título
label_titulo = tk.Label(root, text="Extractor de Datos de PDF", font=("Arial", 16, "bold"), bg="#f0f0f0")
label_titulo.pack(pady=10)

# Instrucción
label = tk.Label(root, text="Seleccione la carpeta que contiene los archivos PDF:", font=("Arial", 12), bg="#f0f0f0")
label.pack(pady=5)

# Botón para seleccionar carpeta
boton_seleccionar = tk.Button(root, text="Seleccionar Carpeta", command=seleccionar_carpeta, font=("Arial", 12), bg="#4CAF50", fg="white")
boton_seleccionar.pack(pady=10)

# Botón para ejecutar extracción
boton_extraer = tk.Button(root, text="Ejecutar Extracción", command=extraer_datos, font=("Arial", 12), bg="#2196F3", fg="white")
boton_extraer.pack(pady=10)

# Etiqueta de estado
label_estado = tk.Label(root, text="Ninguna carpeta seleccionada", font=("Arial", 10), bg="#f0f0f0", fg="grey")
label_estado.pack(pady=5)

# Barra de progreso
barra_progreso = ttk.Progressbar(root, mode="indeterminate")
barra_progreso.pack(pady=10, fill=tk.X, padx=20)

root.mainloop()