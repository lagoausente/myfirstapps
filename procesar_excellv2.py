import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def obtener_columnas_comunes(carpeta):
    archivos = [f for f in os.listdir(carpeta) if f.endswith('.xlsx') or f.endswith('.xls')]
    columnas_comunes = None
    
    for archivo in archivos:
        ruta = os.path.join(carpeta, archivo)
        df = pd.read_excel(ruta, nrows=1)  # Leer solo la primera fila para extraer las columnas
        
        if columnas_comunes is None:
            columnas_comunes = set(df.columns)
        else:
            columnas_comunes &= set(df.columns)  # Intersección de columnas
    
    return list(columnas_comunes) if columnas_comunes else []

def seleccionar_carpeta():
    global carpeta_seleccionada
    carpeta_seleccionada = filedialog.askdirectory()
    
    if carpeta_seleccionada:
        carpeta_label.config(text=f"📂 Carpeta seleccionada:\n{carpeta_seleccionada}")
        columnas = obtener_columnas_comunes(carpeta_seleccionada)
        lista_columnas.delete(0, tk.END)  # Borrar opciones previas
        
        for col in columnas:
            lista_columnas.insert(tk.END, col)  # Añadir columnas comunes

def seleccionar_destino():
    global ruta_destino
    ruta_destino = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if ruta_destino:
        destino_label.config(text=f"📁 Guardar en:\n{ruta_destino}")

def procesar_archivos():
    if not carpeta_seleccionada:
        messagebox.showerror("Error", "Selecciona una carpeta primero")
        return
    
    if not ruta_destino:
        messagebox.showerror("Error", "Selecciona un destino para guardar el archivo")
        return

    filtro_texto = entrada_filtro.get()
    seleccionadas = [lista_columnas.get(i) for i in lista_columnas.curselection()]

    if not seleccionadas:
        messagebox.showerror("Error", "Selecciona al menos una columna")
        return

    archivos = [f for f in os.listdir(carpeta_seleccionada) if f.endswith('.xlsx') or f.endswith('.xls')]
    data_frames = []

    for archivo in archivos:
        ruta = os.path.join(carpeta_seleccionada, archivo)
        df = pd.read_excel(ruta, usecols=seleccionadas)  # Cargar solo las columnas seleccionadas

        if filtro_texto:
            df = df[df.apply(lambda row: row.astype(str).str.contains(filtro_texto, case=False).any(), axis=1)]

        data_frames.append(df)

    if data_frames:
        df_final = pd.concat(data_frames, ignore_index=True)
        df_final.to_excel(ruta_destino, index=False)
        messagebox.showinfo("Éxito", "Archivo procesado y guardado correctamente")

# Interfaz gráfica
root = tk.Tk()
root.title("Procesador de Excel")

tk.Label(root, text="Selecciona una carpeta:").pack()
btn_carpeta = tk.Button(root, text="📂 Seleccionar Carpeta", command=seleccionar_carpeta)
btn_carpeta.pack()
carpeta_label = tk.Label(root, text="📂 Ninguna carpeta seleccionada", fg="gray")
carpeta_label.pack()

tk.Label(root, text="Filtrar por texto:").pack()
entrada_filtro = tk.Entry(root)
entrada_filtro.pack()

tk.Label(root, text="Selecciona columnas:").pack()
lista_columnas = tk.Listbox(root, selectmode=tk.MULTIPLE)
lista_columnas.pack()

tk.Label(root, text="Selecciona el destino del archivo:").pack()
btn_destino = tk.Button(root, text="📁 Seleccionar Destino", command=seleccionar_destino)
btn_destino.pack()
destino_label = tk.Label(root, text="📁 Ningún destino seleccionado", fg="gray")
destino_label.pack()

btn_procesar = tk.Button(root, text="⚙️ Procesar Archivos", command=procesar_archivos)
btn_procesar.pack()

# Variables globales
carpeta_seleccionada = ""
ruta_destino = ""

root.mainloop()
