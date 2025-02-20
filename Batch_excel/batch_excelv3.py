import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, Listbox, MULTIPLE
import tkinter.font as tkFont

class ExcelProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Procesador de Excel Batch")
        self.geometry("900x700")
        self.font_style = tkFont.Font(family="Arial", size=12)

        self.folder_path = ""
        self.output_path = ""
        self.dataframes = []
        self.columns = []
        self.history = []

        # Frame principal
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(pady=10, padx=10, fill="x")


        # Botón para seleccionar carpeta
        self.btn_select = ctk.CTkButton(self.frame, text="Seleccionar carpeta", command=self.load_folder)
        self.btn_select.pack(pady=10)

        # Frame que contiene los Listbox
        self.listbox_frame = ctk.CTkFrame(self.frame)
        self.listbox_frame.pack(pady=10, padx=10, fill="x")

        # Usar grid() en lugar de pack() para alinear bien los listbox
        self.listbox_frame.columnconfigure(0, weight=1)  # Columna 1
        self.listbox_frame.columnconfigure(1, weight=1)  # Columna 2
        self.listbox_frame.columnconfigure(2, weight=1)  # Columna 3

        # Listbox 1: Selección de columnas
        self.column_listbox = Listbox(self.listbox_frame, selectmode=MULTIPLE, height=10, width=15, font=self.font_style)
        self.column_listbox.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

         # Listbox 2: Selección de archivos
        self.center_listbox = Listbox(self.listbox_frame, selectmode=MULTIPLE, height=10, width=25, font=self.font_style)
        self.center_listbox.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

         # Listbox 3: Selección de hojas
        self.right_listbox = Listbox(self.listbox_frame, selectmode=MULTIPLE, height=10, width=20, font=self.font_style)
        self.right_listbox.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")


        # Opciones de transformación (sin "Eliminar espacios")
        self.transform_var = ctk.StringVar(value="Ninguna")
        self.transform_dropdown = ctk.CTkComboBox(self.frame, values=["Ninguna", "Mayúsculas", "Minúsculas"], variable=self.transform_var)
        self.transform_dropdown.pack(pady=10)

        # Entrada para filtro de datos
        self.filter_entry = ctk.CTkEntry(self.frame, placeholder_text="Filtrar filas que contengan...")
        self.filter_entry.pack(pady=10)

        # Selección de formato de exportación
        self.export_var = ctk.StringVar(value="Excel")
        self.export_dropdown = ctk.CTkComboBox(self.frame, values=["Excel", "CSV", "TSV"], variable=self.export_var)
        self.export_dropdown.pack(pady=10)

        # Selección de ruta de salida
        self.btn_output = ctk.CTkButton(self.frame, text="Seleccionar ruta de salida", command=self.select_output)
        self.btn_output.pack(pady=10)

        # Botón para exportar
        self.btn_export = ctk.CTkButton(self.frame, text="Exportar archivo", command=self.export_file)
        self.btn_export.pack(pady=10)

        # Barra de progreso
        self.progress = ctk.CTkProgressBar(self.frame)
        self.progress.pack(pady=10, fill="x")
        self.progress.set(0)

        # Historial de acciones
        self.history_textbox = ctk.CTkTextbox(self.frame, height=100, width=600)
        self.history_textbox.pack(pady=10)

        # Vista previa extendida
        self.tree = ctk.CTkTextbox(self.frame, height=300, width=800)
        self.tree.pack(pady=5, padx=5, fill="both", expand=True)

    def load_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path = folder_selected
            files = [f for f in os.listdir(folder_selected) if f.endswith(('.xlsx', '.xls'))]
            if files:
                self.load_files(files)
            else:
                messagebox.showwarning("Sin archivos", "No se encontraron archivos Excel en la carpeta seleccionada.")

    def load_files(self, files):
        self.dataframes = []
        for file in files:
            file_path = os.path.join(self.folder_path, file)
            df = pd.read_excel(file_path)
            df = self.clean_dataframe(df)
            self.dataframes.append(df)

        if self.dataframes:
            self.columns = list(self.dataframes[0].columns)
            self.column_listbox.delete(0, "end")
            for col in self.columns:
                self.column_listbox.insert("end", col)
            messagebox.showinfo("Archivos cargados", f"Se han cargado {len(self.dataframes)} archivos correctamente.")

    def select_output(self):
        self.output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("TSV", "*.tsv")])
        if self.output_path:
            messagebox.showinfo("Ruta seleccionada", f"El archivo se guardará en: {self.output_path}")

    def export_file(self):
        if not self.dataframes:
            messagebox.showwarning("Sin archivos", "Seleccione una carpeta antes de exportar.")
            return

        selected_indices = self.column_listbox.curselection()
        selected_columns = [self.column_listbox.get(i) for i in selected_indices]

        if not selected_columns:
            messagebox.showwarning("Selección inválida", "Seleccione al menos una columna para exportar.")
            return

        filter_text = self.filter_entry.get()
        transform_option = self.transform_var.get()

        self.progress.set(0.2)

        processed_data = []
        for df in self.dataframes:
            df = df[selected_columns]

            # Siempre eliminar espacios extra
            df = df.applymap(lambda x: self.remove_extra_spaces(x) if isinstance(x, str) else x)

            # Aplicar transformación según la opción elegida
            if transform_option == "Mayúsculas":
                df = df.applymap(lambda x: x.upper() if isinstance(x, str) else x)
            elif transform_option == "Minúsculas":
                df = df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

            if filter_text:
                df = df[df.astype(str).apply(lambda row: row.str.contains(filter_text, case=False).any(), axis=1)]

            processed_data.append(df)

        if processed_data:
            final_df = pd.concat(processed_data, ignore_index=True)
            self.display_preview(final_df)
            self.history.append(f"Exportado: {len(processed_data)} archivos con {len(final_df)} filas.")

            if not self.output_path:
                self.select_output()
                if not self.output_path:
                    return

            export_format = self.export_var.get()
            if export_format == "Excel":
                final_df.to_excel(self.output_path, index=False)
            elif export_format == "CSV":
                final_df.to_csv(self.output_path, index=False, sep=",")
            elif export_format == "TSV":
                final_df.to_csv(self.output_path, index=False, sep="\t")

            self.progress.set(1.0)
            messagebox.showinfo("Exportación completada", "El archivo se ha guardado correctamente.")

    def remove_extra_spaces(self, text):
        """Elimina espacios extra dentro del texto (manteniendo separación entre palabras)"""
        while "  " in text:  # Reemplaza espacios dobles hasta que solo quede uno
            text = text.replace("  ", " ")
        return text.strip()  # También elimina espacios al inicio y final

    def display_preview(self, df):
        preview_text = df.head(10).to_string(index=False)
        self.tree.delete("1.0", "end")
        self.tree.insert("1.0", preview_text)

    def clean_dataframe(self, df):
        df = df.dropna(how='all')  
        df = df.loc[:, ~df.columns.duplicated()]  
        df.columns = df.columns.str.strip().str.lower()  
        df = df.loc[:, ~df.columns.str.startswith("unnamed")]  
        return df

if __name__ == "__main__":
    app = ExcelProcessorApp()
    app.mainloop()
