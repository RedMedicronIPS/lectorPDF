import fitz  # PyMuPDF
import pandas as pd
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Lista para almacenar los datos extraídos
todos_datos = []
carpeta_pdf = None

# Frases de inicio y fin para extraer secciones específicas
frases = [
    ("Fecha:", "NÚMERO DE AUTORIZACIÓN:"),
    ("Permiso especial de permanencia", "Número documento de identificación"),
    ("Manejo integral según guía de:", "SERVICIO\nCÓDIGO"),
    ("Lazos", "Notas auditor:")
]

# Función para preprocesar el texto
def preprocesar_texto(texto):
    texto = re.sub(r'([a-zA-Z]+)\n([a-zA-Z]+)', r'\1 \2', texto)
    # Reemplazar ')' seguido de '\n' por un espacio
    texto = re.sub(r'\)\n', ') ', texto)
    # Reemplazar saltos de línea entre palabras, números o símbolos
    #texto = re.sub(r'(\S)\n(\S)', r'\1 \2', texto)
    # Eliminar saltos de línea adicionales
    #texto = re.sub(r'\n+', ' ', texto)
    # Eliminar espacios extra innecesarios
    #texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

#def preprocesar_texto(texto):
#    # Reemplazar saltos de línea entre palabras, números o símbolos
#    texto = re.sub(r'(\S)\n(\S)', r'\1 \2', texto)
#    # Eliminar saltos de línea adicionales
#    texto = re.sub(r'\n+', ' ', texto)
#    # Eliminar espacios extra innecesarios
#    texto = re.sub(r'\s+', ' ', texto).strip()
#    return texto


# Procesar texto en secciones
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
        elif re.match(r'^[a-zA-Z]+ \d+$', linea):
            resultado.append(linea)
        elif re.match(r'^\d+ \d+$', linea):
            partes = linea.split()
            resultado.extend(partes)
        else:
            resultado.append(linea)

    return resultado

# Extraer datos de PDFs
def extraer_datos():
    global carpeta_pdf
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
            texto_pagina = pagina.get_text("text")
            if not texto_pagina.strip():
                print(f"No hay texto en {archivo}, página {pagina_num + 1}")
                continue

            data = {'Nombre Archivo': archivo}

            for i, (frase_inicio, frase_final) in enumerate(frases):
                inicio = texto_pagina.find(frase_inicio)
                fin = texto_pagina.find(frase_final, inicio + len(frase_inicio))

                if inicio == -1 or fin == -1:
                    continue

                texto_seccion = texto_pagina[inicio + len(frase_inicio):fin].strip()
                lista_seccion = procesar_seccion(texto_seccion)
                data[f"Seccion {i + 1}"] = lista_seccion

            # Guardar en todos_datos
            if 'Seccion 3' in data:
                seccion_3 = data['Seccion 3']
                for i in range(0, len(seccion_3), 3):
                    if len(seccion_3) >= i + 3:
                        todos_datos.append({
                            'Autorizacion': data.get('Seccion 1', ''),
                            'Nombre': seccion_3[i],
                            'Codigo': seccion_3[i + 1],
                            'Cantidad': seccion_3[i + 2]
                        })
            if 'Seccion 4' in data:
                seccion_4 = data['Seccion 4']
                for i in range(0, len(seccion_4), 3):
                    if len(seccion_4) >= i + 3:
                        todos_datos.append({
                            'Autorizacion': data.get('Seccion 1', ''),
                            'Nombre': seccion_4[i],
                            'Codigo': seccion_4[i + 1],
                            'Cantidad': seccion_4[i + 2]
                        })

        pdf.close()
    print(todos_datos)
    df = pd.DataFrame(todos_datos)
    archivo_csv = os.path.join(carpeta_pdf, 'datos_extraidos.csv')
    df.to_csv(archivo_csv, index=False)
    messagebox.showinfo("Proceso Completo", f"Datos guardados en '{archivo_csv}'")
    barra_progreso.stop()

# Funciones de interfaz gráfica
def seleccionar_carpeta():
    global carpeta_pdf
    carpeta_pdf = filedialog.askdirectory()
    label_estado.config(text=f"Carpeta seleccionada: {carpeta_pdf}")

# Configuración de la interfaz
root = tk.Tk()
root.title("Extractor PDF")
label_estado = tk.Label(root, text="Ninguna carpeta seleccionada")
label_estado.pack()
boton_seleccionar = tk.Button(root, text="Seleccionar Carpeta", command=seleccionar_carpeta)
boton_seleccionar.pack()
boton_extraer = tk.Button(root, text="Extraer Datos", command=extraer_datos)
boton_extraer.pack()
barra_progreso = ttk.Progressbar(root, mode="indeterminate")
barra_progreso.pack()
root.mainloop()
