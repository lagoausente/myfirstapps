import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def seleccionar_carpeta_origen():
    carpeta = filedialog.askdirectory()
    origen_var.set(carpeta)
    cargar_columnas_disponibles()  # Se cargan las columnas al seleccionar carpeta

def seleccionar_carpeta_destino():
    carpeta = filedialog.askdirectory()
    destino_var.set(carpeta)

def cargar_columnas_disponibles():
    """Carga las columnas del primer archivo encontrado en la carpeta de origen"""
    carpeta_origen = origen_var.get()
    if not carpeta_origen:
        return

    archivos = [f for f in os.listdir(carpeta_origen) if f.endswith('.xlsx') or f.endswith('.xls')]
    if not archivos:
        messagebox.showerror("Error", "No se encontraron archivos Excel en la carpeta de origen.")
        return

    archivo_path = os.path.join(carpeta_origen, archivos[0])  # Tomar el primer archivo
    df = pd.read_excel(archivo_path)

    listbox_columnas.delete(0, tk.END)  # Limpiar lista
    for col in df.columns:
        listbox_columnas.insert(tk.END, col)  # Agregar columnas dinámicamente

def obtener_columnas_seleccionadas():
    """Obtiene las columnas seleccionadas en la lista"""
    seleccion = listbox_columnas.curselection()
    return [listbox_columnas.get(i) for i in seleccion]

def procesar_archivos():
    carpeta_origen = origen_var.get()
    carpeta_destino = destino_var.get()
    nombre_archivo = nombre_var.get()
    texto_filtro = filtro_texto_var.get()

    if not carpeta_origen or not carpeta_destino or not nombre_archivo:
        messagebox.showerror("Error", "Por favor, complete todos los campos.")
        return

    archivos = [f for f in os.listdir(carpeta_origen) if f.endswith('.xlsx') or f.endswith('.xls')]
    
    if not archivos:
        messagebox.showerror("Error", "No se encontraron archivos Excel en la carpeta de origen.")
        return

    df_combinado = pd.DataFrame()

    for archivo in archivos:
        archivo_path = os.path.join(carpeta_origen, archivo)
        df = pd.read_excel(archivo_path)

        if filtro_edad.get():
            if 'Edad' in df.columns:
                df = df[df['Edad'] > 30]
        
        if modificar_nombre.get():
            if 'Nombre' in df.columns:
                df['Nombre'] = df['Nombre'].str.upper()
        
        if texto_filtro:
            df = df[df.apply(lambda row: row.astype(str).str.contains(texto_filtro, case=False).any(), axis=1)]

        columnas = obtener_columnas_seleccionadas()
        if columnas:
            df = df[columnas]

        df_combinado = pd.concat([df_combinado, df], ignore_index=True)

    archivo_salida = os.path.join(carpeta_destino, nombre_archivo + '.xlsx')
    df_combinado.to_excel(archivo_salida, index=False)

    messagebox.showinfo("Éxito", f"Archivo combinado guardado como {archivo_salida}")

# Crear la ventana principal
root = tk.Tk()
root.title("Procesador de Archivos Excel")

# Variables de Tkinter
origen_var = tk.StringVar()
destino_var = tk.StringVar()
nombre_var = tk.StringVar()
filtro_texto_var = tk.StringVar()

modificar_nombre = tk.BooleanVar()
columnas_seleccionadas = tk.BooleanVar()
filtro_edad = tk.BooleanVar()

# Interfaz gráfica
tk.Label(root, text="Carpeta de origen:").grid(row=0, column=0, padx=10, pady=10)
entry_origen = tk.Entry(root, textvariable=origen_var, width=50)
entry_origen.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Seleccionar", command=seleccionar_carpeta_origen).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Carpeta de destino:").grid(row=1, column=0, padx=10, pady=10)
entry_destino = tk.Entry(root, textvariable=destino_var, width=50)
entry_destino.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Seleccionar", command=seleccionar_carpeta_destino).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Nombre archivo de salida:").grid(row=2, column=0, padx=10, pady=10)
entry_nombre = tk.Entry(root, textvariable=nombre_var, width=50)
entry_nombre.grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="Filtrar por texto:").grid(row=3, column=0, padx=10, pady=10)
entry_filtro_texto = tk.Entry(root, textvariable=filtro_texto_var, width=50)
entry_filtro_texto.grid(row=3, column=1, padx=10, pady=10)

tk.Checkbutton(root, text="Filtrar Edad > 30", variable=filtro_edad).grid(row=4, column=0, columnspan=3)
tk.Checkbutton(root, text="Modificar Nombre a Mayúsculas", variable=modificar_nombre).grid(row=5, column=0, columnspan=3)

tk.Label(root, text="Seleccionar columnas:").grid(row=6, column=0, padx=10, pady=10)
listbox_columnas = tk.Listbox(root, selectmode=tk.MULTIPLE, height=6, width=50)
listbox_columnas.grid(row=6, column=1, padx=10, pady=10)

tk.Button(root, text="Procesar Archivos", command=procesar_archivos).grid(row=7, column=0, columnspan=3, pady=20)

# Iniciar la aplicación
root.mainloop()
