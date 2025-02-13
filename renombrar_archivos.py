import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import simpledialog

# Función para renombrar los archivos en el directorio seleccionado
def renombrar_archivos():
    # Obtener el directorio desde el que se quiere renombrar los archivos
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los archivos")
    
    # Comprobar si se seleccionó una carpeta
    if not carpeta:
        return
    
    # Obtener las opciones de renombrado
    prefijo = prefijo_entry.get()
    sufijo = sufijo_entry.get()
    reemplazar_nombre = reemplazar_var.get()
    tipos_archivo = tipos_archivo_entry.get().split(",")  # Separa por comas
    
    try:
        archivos_a_renombrar = []
        # Recorrer todos los archivos de la carpeta
        for archivo in os.listdir(carpeta):
            if any(archivo.endswith(ext.strip()) for ext in tipos_archivo):
                # Si es un archivo compatible
                archivo_antiguo = os.path.join(carpeta, archivo)
                
                if reemplazar_nombre:
                    # Si se reemplaza el nombre completamente
                    nombre_nuevo = prefijo + archivo.split('.')[0] + sufijo + "." + archivo.split('.')[-1]
                else:
                    # Si no se reemplaza, solo se agrega prefijo/sufijo
                    nombre_nuevo = prefijo + archivo + sufijo

                archivos_a_renombrar.append((archivo, nombre_nuevo))

        # Mostrar vista previa de los cambios
        mostrar_vista_previa(archivos_a_renombrar)

    except Exception as e:
        messagebox.showerror("Error", f"Hubo un error al procesar los archivos: {e}")


# Función para mostrar la vista previa de los cambios
def mostrar_vista_previa(archivos):
    # Crear una nueva ventana para mostrar la vista previa
    vista_previa_ventana = tk.Toplevel(ventana)
    vista_previa_ventana.title("Vista Previa de Renombrado")
    
    # Crear un scroll
    scroll = tk.Scrollbar(vista_previa_ventana)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    
    lista_archivos = tk.Listbox(vista_previa_ventana, width=80, height=15, yscrollcommand=scroll.set)
    lista_archivos.pack(pady=10)
    
    # Agregar los archivos a la lista
    for archivo_antiguo, archivo_nuevo in archivos:
        lista_archivos.insert(tk.END, f"{archivo_antiguo} --> {archivo_nuevo}")

    # Botón para confirmar el renombrado
    confirmar_button = tk.Button(vista_previa_ventana, text="Confirmar Renombrado", command=lambda: confirmar_renombrado(archivos))
    confirmar_button.pack(pady=10)

    scroll.config(command=lista_archivos.yview)


# Función para confirmar el renombrado de los archivos
def confirmar_renombrado(archivos):
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los archivos")
    
    try:
        for archivo_antiguo, archivo_nuevo in archivos:
            ruta_antigua = os.path.join(carpeta, archivo_antiguo)
            ruta_nueva = os.path.join(carpeta, archivo_nuevo)
            os.rename(ruta_antigua, ruta_nueva)

        messagebox.showinfo("Renombrado", "Los archivos se han renombrado con éxito.")
    except Exception as e:
        messagebox.showerror("Error", f"Hubo un error al renombrar los archivos: {e}")


# Crear ventana principal de la aplicación
ventana = tk.Tk()
ventana.title("Renombrar Archivos - Versión Avanzada")
ventana.geometry("500x350")

# Etiqueta y campo de entrada para el prefijo
prefijo_label = tk.Label(ventana, text="Prefijo:")
prefijo_label.pack(pady=5)
prefijo_entry = tk.Entry(ventana, width=40)
prefijo_entry.pack(pady=5)

# Etiqueta y campo de entrada para el sufijo
sufijo_label = tk.Label(ventana, text="Sufijo:")
sufijo_label.pack(pady=5)
sufijo_entry = tk.Entry(ventana, width=40)
sufijo_entry.pack(pady=5)

# Opción para reemplazar completamente el nombre del archivo
reemplazar_var = tk.BooleanVar()
reemplazar_check = tk.Checkbutton(ventana, text="Reemplazar nombre completamente", variable=reemplazar_var)
reemplazar_check.pack(pady=5)

# Etiqueta y campo de entrada para los tipos de archivo
tipos_archivo_label = tk.Label(ventana, text="Tipos de archivo (separados por coma):")
tipos_archivo_label.pack(pady=5)
tipos_archivo_entry = tk.Entry(ventana, width=40)
tipos_archivo_entry.pack(pady=5)
tipos_archivo_entry.insert(0, ".xlsx, .xls, .csv")  # valor por defecto

# Botón para iniciar el proceso de renombrado
renombrar_button = tk.Button(ventana, text="Renombrar Archivos", command=renombrar_archivos)
renombrar_button.pack(pady=20)

# Iniciar la ventana
ventana.mainloop()
