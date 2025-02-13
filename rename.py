import os
import tkinter as tk
from tkinter import filedialog

def renombrar_archivos():
    # Obtener los directorios seleccionados por el usuario
    origen_dir = origen_var.get()
    destino_dir = destino_var.get()
    
    # Renombrar los archivos dentro del directorio origen y moverlos al destino
    try:
        for archivo in os.listdir(origen_dir):
            if archivo.endswith(".xlsx"):  # Solo renombrar archivos .xlsx
                nuevo_nombre = prefijo_var.get() + archivo + sufijo_var.get()
                os.rename(os.path.join(origen_dir, archivo), os.path.join(destino_dir, nuevo_nombre))
        status_label.config(text="Archivos renombrados correctamente")
    except Exception as e:
        status_label.config(text=f"Error: {str(e)}")

def seleccionar_origen():
    origen_dir = filedialog.askdirectory(title="Selecciona la carpeta de origen")
    origen_var.set(origen_dir)

def seleccionar_destino():
    destino_dir = filedialog.askdirectory(title="Selecciona la carpeta de destino")
    destino_var.set(destino_dir)

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Renombrador de Archivos")

# Variables
origen_var = tk.StringVar()
destino_var = tk.StringVar()
prefijo_var = tk.StringVar()
sufijo_var = tk.StringVar()

# Crear la interfaz gráfica
tk.Label(ventana, text="Carpeta de origen:").grid(row=0, column=0)
tk.Entry(ventana, textvariable=origen_var, width=40).grid(row=0, column=1)
tk.Button(ventana, text="Seleccionar", command=seleccionar_origen).grid(row=0, column=2)

tk.Label(ventana, text="Carpeta de destino:").grid(row=1, column=0)
tk.Entry(ventana, textvariable=destino_var, width=40).grid(row=1, column=1)
tk.Button(ventana, text="Seleccionar", command=seleccionar_destino).grid(row=1, column=2)

tk.Label(ventana, text="Prefijo:").grid(row=2, column=0)
tk.Entry(ventana, textvariable=prefijo_var, width=40).grid(row=2, column=1)

tk.Label(ventana, text="Sufijo:").grid(row=3, column=0)
tk.Entry(ventana, textvariable=sufijo_var, width=40).grid(row=3, column=1)

tk.Button(ventana, text="Renombrar Archivos", command=renombrar_archivos).grid(row=4, column=0, columnspan=3)

status_label = tk.Label(ventana, text="", fg="green")
status_label.grid(row=5, column=0, columnspan=3)

# Ejecutar la ventana
ventana.mainloop()
