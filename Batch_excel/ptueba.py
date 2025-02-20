import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ExcelProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Procesador de Excel")
        self.geometry("1100x700")

        self.folder_path = ""
        self.files = {}  # Diccionario de archivos y sus hojas
        self.selected_files = []
        self.selected_sheets = {}

        self.frame = ctk.CTkFrame(self)
        self.frame.pack(pady=10, padx=10, fill="both", expand=True)

        # Sección de selección de carpeta
        ctk.CTkLabel(self.frame, text="1. Seleccionar Carpeta con Archivos Excel:").grid(row=0, column=0, columnspan=2, pady=5)
        self.btn_select_folder = ctk.CTkButton(self.frame, text="Seleccionar Carpeta", command=self.load_folder)
        self.btn_select_folder.grid(row=1, column=0, columnspan=2, pady=5)

        # Sección de archivos y hojas
        ctk.CTkLabel(self.frame, text="2. Seleccionar Archivos:").grid(row=2, column=0, pady=5, sticky="w")
        self.file_scroll = ctk.CTkScrollableFrame(self.frame, width=350, height=200)
        self.file_scroll.grid(row=3, column=0, padx=10, pady=5, sticky="w")

        self.btn_confirm_files = ctk.CTkButton(self.frame, text="Confirmar Archivos", command=self.confirm_files)
        self.btn_confirm_files.grid(row=4, column=0, pady=5, sticky="w")

        ctk.CTkLabel(self.frame, text="3. Seleccionar Hojas:").grid(row=2, column=1, pady=5, sticky="w")
        self.sheet_scroll = ctk.CTkScrollableFrame(self.frame, width=350, height=200)
        self.sheet_scroll.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        self.btn_confirm_sheets = ctk.CTkButton(self.frame, text="Confirmar Hojas", command=self.confirm_sheets)
        self.btn_confirm_sheets.grid(row=4, column=1, pady=5, sticky="w")

        # Botón para procesar
        self.btn_process = ctk.CTkButton(self.frame, text="Procesar Datos y Exportar", command=self.process_selected_sheets)
        self.btn_process.grid(row=5, column=0, columnspan=2, pady=10)

    def load_folder(self):
        """Carga los archivos Excel de la carpeta seleccionada y los muestra en la lista."""
        folder_selected = filedialog.askdirectory()
        if not folder_selected:
            return

        self.folder_path = folder_selected
        self.files.clear()

        for widget in self.file_scroll.winfo_children():
            widget.destroy()  # Limpiar la lista anterior

        archivos_excel = [f for f in os.listdir(folder_selected) if f.endswith((".xlsx", ".xls"))]

        if not archivos_excel:
            messagebox.showwarning("Sin archivos", "No se encontraron archivos Excel en la carpeta seleccionada.")
            return

        self.file_checkboxes = {}  # Diccionario de Checkboxes para los archivos

        for file in archivos_excel:
            file_path = os.path.join(folder_selected, file)
            try:
                xls = pd.ExcelFile(file_path)
                self.files[file] = xls.sheet_names  # Guardamos las hojas del archivo
                var = ctk.BooleanVar()
                checkbox = ctk.CTkCheckBox(self.file_scroll, text=file, variable=var)
                checkbox.pack(anchor="w", padx=5, pady=2)
                self.file_checkboxes[file] = var  # Asociamos cada archivo con su checkbox
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el archivo {file}.\nError: {e}")

        messagebox.showinfo("Cargados", f"Se encontraron {len(self.files)} archivos Excel.")

    def confirm_files(self):
        """Confirma los archivos seleccionados y muestra sus hojas en la lista de hojas."""
        self.selected_files = [file for file, var in self.file_checkboxes.items() if var.get()]

        if not self.selected_files:
            messagebox.showwarning ("Sin selección", "Seleccion")