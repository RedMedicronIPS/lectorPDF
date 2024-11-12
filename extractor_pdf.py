import fitz  # PyMuPDF
import pandas as pd
import os
import re 
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Lista para almacenar los datos extraídos
todos_datos = []

# Frases de inicio y fin para extraer secciones específicas
frases = [
    ("Fecha:", "NÚMERO DE AUTORIZACIÓN:"),
    ("Permiso especial de permanencia", "Número documento de identificación"),
    ("Manejo integral según guía de:", "SERVICIO\nCÓDIGO"),
    ("Lazos", "Notas auditor:")
]

# Función para preprocesar el texto y reemplazar "\n" entre palabras
def preprocesar_texto(texto):
    texto = re.sub(r'([a-zA-Z]+)\n([a-zA-Z]+)', r'\1 \2', texto)
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
        pdf = fitz.open(ruta_pdf)
        data = {'Nombre Archivo': archivo}

        for pagina_num, pagina in enumerate(pdf):
            texto_pagina = pagina.get_text("text")

            for i, (frase_inicio, frase_final) in enumerate(frases):
                inicio = texto_pagina.find(frase_inicio)
                fin = texto_pagina.find(frase_final, inicio + len(frase_inicio))

                if inicio != -1 and fin != -1:
                    texto_seccion = texto_pagina[inicio + len(frase_inicio):fin].strip()
                    lista_seccion = procesar_seccion(texto_seccion)
                    data[f"Seccion {i + 1}"] = lista_seccion

            if 'Seccion 3' in data:
                seccion_3 = data['Seccion 3']
                for i in range(0, len(seccion_3), 3):
                    nombre = seccion_3[i]
                    codigo = seccion_3[i + 1]
                    cantidad = seccion_3[i + 2]
                    autorizacion = str(data.get('Seccion 1', '')).replace("'", '').replace('[', '').replace(']', '').strip()
                    numero_documento = str(data.get('Seccion 2', '')).replace("'", '').replace('[', '').replace(']', '').strip()

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
                if len(seccion_4) % 3 == 0:
                    for i in range(0, len(seccion_4), 3):
                        nombre = seccion_4[i]
                        codigo = seccion_4[i + 1]
                        cantidad = seccion_4[i + 2]
                        autorizacion = str(data.get('Seccion 1', '')).replace("'", '').replace('[', '').replace(']', '').strip()
                        numero_documento = str(data.get('Seccion 2', '')).replace("'", '').replace('[', '').replace(']', '').strip()

                        entry = {
                            'Autorizacion': autorizacion,
                            'Numero Documento': numero_documento,
                            'Nombre': nombre,
                            'Codigo': codigo,
                            'Cantidad': cantidad
                        }
                        todos_datos.append(entry)
                else:
                    print(f"Advertencia: La Sección 4 del archivo {archivo} no tiene una cantidad de elementos múltiplo de 3.")

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
