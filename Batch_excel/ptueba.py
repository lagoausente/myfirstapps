import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, Listbox, MULTIPLE

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ExcelProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Procesador de Excel")
        self.geometry("1000x700")

        self.folder_path = ""
        self.files = {}
        self.selected_files = []
        self.selected_sheets = {}
        self.columns = []

        self.frame = ctk.CTkFrame(self)
        self.frame.pack(pady=10, padx=10, fill="both", expand=True)

        ctk.CTkLabel(self.frame, text="1. Seleccionar Carpeta con Archivos Excel:").grid(row=0, column=0, columnspan=2, pady=5)
        self.btn_select_folder = ctk.CTkButton(self.frame, text="Seleccionar Carpeta", command=self.load_folder)
        self.btn_select_folder.grid(row=1, column=0, columnspan=2, pady=5)

        ctk.CTkLabel(self.frame, text="2. Seleccionar Archivos:").grid(row=2, column=0, pady=5, sticky="w")
        self.file_listbox = Listbox(self.frame, selectmode=MULTIPLE, height=7, width=50)
        self.file_listbox.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.btn_confirm_files = ctk.CTkButton(self.frame, text="Confirmar Archivos", command=self.confirm_files)
        self.btn_confirm_files.grid(row=4, column=0, pady=5, sticky="w")

        ctk.CTkLabel(self.frame, text="3. Seleccionar Hojas:").grid(row=2, column=1, pady=5, sticky="w")
        self.sheet_listbox = Listbox(self.frame, selectmode=MULTIPLE, height=7, width=50)
        self.sheet_listbox.grid(row=3, column=1, padx=10, pady=5, sticky="w")
        self.btn_confirm_sheets = ctk.CTkButton(self.frame, text="Confirmar Hojas", command=self.confirm_sheets)
        self.btn_confirm_sheets.grid(row=4, column=1, pady=5, sticky="w")

        ctk.CTkLabel(self.frame, text="4. Seleccionar Columnas Comunes:").grid(row=5, column=0, columnspan=2, pady=5)
        self.column_listbox = Listbox(self.frame, selectmode=MULTIPLE, height=7, width=100)
        self.column_listbox.grid(row=6, column=0, columnspan=2, padx=10, pady=5)

        ctk.CTkLabel(self.frame, text="5. Opciones de Procesamiento:").grid(row=7, column=0, columnspan=2, pady=5)
        
        ctk.CTkLabel(self.frame, text="Filtrar filas que contengan:").grid(row=8, column=0, sticky="w", padx=10)
        self.filter_entry = ctk.CTkEntry(self.frame, placeholder_text="Ingrese texto para filtrar...")
        self.filter_entry.grid(row=9, column=0, padx=10, pady=5, sticky="w")
        
        ctk.CTkLabel(self.frame, text="Formato de Texto:").grid(row=8, column=1, sticky="w", padx=10)
        self.transform_var = ctk.StringVar(value="Ninguna")
        self.transform_dropdown = ctk.CTkComboBox(self.frame, values=["Ninguna", "Mayúsculas", "Minúsculas"], variable=self.transform_var)
        self.transform_dropdown.grid(row=9, column=1, padx=10, pady=5, sticky="w")
        
        ctk.CTkLabel(self.frame, text="Formato de Exportación:").grid(row=10, column=0, sticky="w", padx=10)
        self.export_var = ctk.StringVar(value="Excel")
        self.export_dropdown = ctk.CTkComboBox(self.frame, values=["Excel", "CSV", "TSV"], variable=self.export_var)
        self.export_dropdown.grid(row=11, column=0, padx=10, pady=5, sticky="w")

        self.btn_output = ctk.CTkButton(self.frame, text="Seleccionar Carpeta de Exportación", command=self.select_output)
        self.btn_output.grid(row=12, column=0, columnspan=2, pady=5)

        self.btn_process = ctk.CTkButton(self.frame, text="Procesar Datos y Exportar", command=self.process_selected_sheets)
        self.btn_process.grid(row=13, column=0, columnspan=2, pady=10)

    def load_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path = folder_selected
            messagebox.showinfo("Carpeta Seleccionada", f"Carpeta: {self.folder_path}")

    def confirm_files(self):
        pass  # Implementar lógica de selección de archivos

    def confirm_sheets(self):
        pass  # Implementar lógica de selección de hojas

    def select_output(self):
        self.output_path = filedialog.askdirectory()

    def process_selected_sheets(self):
        messagebox.showinfo("Exportación", "Proceso de exportación en construcción.")

if __name__ == "__main__":
    app = ExcelProcessorApp()
    app.mainloop()
