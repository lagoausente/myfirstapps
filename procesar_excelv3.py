import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def obtener_columnas_comunes(carpeta):
    try:
        archivos = [f for f in os.listdir(carpeta) if f.endswith('.xlsx') or f.endswith('.xls')]
        print(f"Archivos encontrados en la carpeta: {archivos}")  # Depuración

        if not archivos:
            print("⚠️ No se encontraron archivos Excel en la carpeta seleccionada.")
            return []

        columnas_comunes = None
        
        for archivo in archivos:
            ruta = os.path.join(carpeta, archivo)
            print(f"Intentando leer: {ruta}")  # Ver qué archivos se están leyendo

            try:
                df = pd.read_excel(ruta, header=None)  # No asume nombres de columna

                if df.empty:
                    print(f"⚠️ El archivo {archivo} está vacío o no tiene contenido útil.")
                    continue

                # Buscar la primera fila con contenido
                for i, row in df.iterrows():
                    if not row.isnull().all():  # Si la fila no está completamente vacía
                        df.columns = row  # Usar esta fila como nombres de columna
                        df = df.iloc[i+1:]  # Eliminar las filas superiores
                        break

                # Si todas las filas estaban vacías o los nombres de columna son NaN, asignar nombres automáticos
                if df.columns.isnull().all():
                    df.columns = [f"Columna_{i+1}" for i in range(len(df.columns))]

                print(f"Columnas en {archivo}: {df.columns.tolist()}")  # Ver columnas detectadas

                if columnas_comunes is None:
                    columnas_comunes = set(df.columns)
                else:
                    columnas_comunes &= set(df.columns)  # Intersección de columnas

            except Exception as e:
                print(f"❌ Error al leer {archivo}: {e}")  # Mostrar error si falla la lectura
                continue  # Saltar este archivo y seguir con el siguiente

        print(f"✅ Columnas comunes detectadas: {columnas_comunes}")  # Última verificación
        return list(columnas_comunes) if columnas_comunes else []

    except Exception as e:
        print(f"❌ Error en obtener_columnas_comunes(): {e}")
        return []



def seleccionar_carpeta():
    global carpeta_seleccionada
    carpeta_seleccionada = filedialog.askdirectory()
    
    if carpeta_seleccionada:
        print(f"📂 Carpeta seleccionada: {carpeta_seleccionada}")

        # Verificar que hay archivos en la carpeta
        archivos = os.listdir(carpeta_seleccionada)
        print(f"Archivos detectados en la carpeta: {archivos}")

        carpeta_label.config(text=f"📂 Carpeta seleccionada:\n{carpeta_seleccionada}")

        columnas = obtener_columnas_comunes(carpeta_seleccionada)
        print(f"Columnas a mostrar en Listbox: {columnas}")

        lista_columnas.delete(0, tk.END)  # Borrar opciones previas
        
        if not columnas:
            lista_columnas.insert(tk.END, "⚠️ No se encontraron columnas")

        for col in columnas:
            lista_columnas.insert(tk.END, col)  # Añadir columnas comunes


def seleccionar_destino():
    global ruta_destino
    ruta_destino = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if ruta_destino:
        destino_label.config(text=f"\ud83d\udcc1 Guardar en:\n{ruta_destino}")

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
btn_carpeta = tk.Button(root, text="\📁 Seleccionar Carpeta", command=seleccionar_carpeta)
btn_carpeta.pack()
carpeta_label = tk.Label(root, text="\📁 Ninguna carpeta seleccionada", fg="gray")
carpeta_label.pack()

tk.Label(root, text="Filtrar por texto:").pack()
entrada_filtro = tk.Entry(root)
entrada_filtro.pack()

tk.Label(root, text="Selecciona columnas:").pack()
lista_columnas = tk.Listbox(root, selectmode=tk.MULTIPLE)
lista_columnas.pack()

tk.Label(root, text="Selecciona el destino del archivo:").pack()
btn_destino = tk.Button(root, text="\📁Seleccionar Destino", command=seleccionar_destino)
btn_destino.pack()
destino_label = tk.Label(root, text="\📁 Ningún destino seleccionado", fg="gray")
destino_label.pack()

btn_procesar = tk.Button(root, text="⚙️ Procesar Archivos", command=procesar_archivos)
btn_procesar.pack()

# Variables globales
carpeta_seleccionada = ""
ruta_destino = ""

root.mainloop()
