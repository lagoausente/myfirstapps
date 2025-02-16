import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def obtener_columnas_comunes(carpeta):
    try:
        archivos = [f for f in os.listdir(carpeta) if f.endswith('.xlsx') or f.endswith('.xls')]
        if not archivos:
            return []
        columnas_comunes = None
        for archivo in archivos:
            ruta = os.path.join(carpeta, archivo)
            try:
                df = pd.read_excel(ruta, header=None)
                if df.empty:
                    continue
                for i, row in df.iterrows():
                    if not row.isnull().all():
                        df.columns = [str(col).strip().lower() for col in row]  # Convierte nombres a minúsculas
                        df = df.iloc[i+1:]  # Elimina filas superiores
                        # 🔹 Ahora eliminamos columnas vacías si quedaron
                        df = df.dropna(axis=1, how="all")
                        df = df.loc[:, (df.columns != "")]  # También elimina columnas con nombre vacío
                        break
                if df.columns.isnull().all():
                    df.columns = [f"Columna_{i+1}" for i in range(len(df.columns))]
                if columnas_comunes is None:
                    columnas_comunes = set(df.columns)
                else:
                    columnas_comunes &= set(df.columns)
            except:
                continue
        return list(columnas_comunes) if columnas_comunes else []
    except:
        return []

def seleccionar_carpeta():
    global carpeta_seleccionada
    carpeta_seleccionada = filedialog.askdirectory()
    if carpeta_seleccionada:
        carpeta_label.config(text=f"📂 Carpeta seleccionada:\n{carpeta_seleccionada}")
        columnas = obtener_columnas_comunes(carpeta_seleccionada)
        lista_columnas.delete(0, tk.END)
        if not columnas:
            lista_columnas.insert(tk.END, "⚠️ No se encontraron columnas")
        for col in columnas:
            lista_columnas.insert(tk.END, col)

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
    progress_bar.start()
    for archivo in archivos:
        ruta = os.path.join(carpeta_seleccionada, archivo)
        # Leer el archivo sin filtrar columnas primero
        df = pd.read_excel(ruta)
        columnas_validas = [col for col in seleccionadas if col in df.columns]
        df = df[columnas_validas]  # Aplicar la selección de columnas
        if filtro_texto:
            df = df[df.apply(lambda row: row.astype(str).str.contains(filtro_texto, case=False).any(), axis=1)]
        data_frames.append(df)
    if data_frames:
        df_final = pd.concat(data_frames, ignore_index=True)
        df_final.to_excel(ruta_destino, index=False)
        messagebox.showinfo("Éxito", "Archivo procesado y guardado correctamente")
    progress_bar.stop()

root = tk.Tk()
root.title("Procesador de Excel")
root.geometry("500x400")
style = ttk.Style()
style.theme_use("clam")  # Fuerza un tema compatible que respeta los colores
style.configure("TButton", padding=5, relief="flat", font=("Arial", 10))
style.map("TButton", background=[("active", "#005bb5"), ("!active", "#0078D7")], foreground=[("active", "white"), ("!active", "white")])
frame = ttk.Frame(root, padding=10)
frame.pack(fill=tk.BOTH, expand=True)

carpeta_label = ttk.Label(frame, text="📂 Ninguna carpeta seleccionada", foreground="black")
carpeta_label.pack()
btn_carpeta = ttk.Button(frame, text="📁 Seleccionar Carpeta", command=seleccionar_carpeta)
btn_carpeta.pack()

tk.Label(frame, text="Filtrar por texto:").pack()
entrada_filtro = ttk.Entry(frame)
entrada_filtro.pack()

tk.Label(frame, text="Selecciona columnas:").pack()
lista_columnas = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=6)
lista_columnas.pack()

destino_label = ttk.Label(frame, text="📁 Ningún destino seleccionado", foreground="black")
destino_label.pack()
btn_destino = ttk.Button(frame, text="📁 Seleccionar Destino", command=seleccionar_destino)
btn_destino.pack()

btn_procesar = ttk.Button(frame, text="⚙️ Procesar Archivos", command=procesar_archivos)
btn_procesar.pack()

progress_bar = ttk.Progressbar(frame, mode="indeterminate")
progress_bar.pack(fill=tk.X, pady=5)

root.mainloop()
