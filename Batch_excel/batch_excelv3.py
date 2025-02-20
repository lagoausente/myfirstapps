import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, Listbox, MULTIPLE

ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue")

class ExcelProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Procesador de Excel Batch")
        self.geometry("900x700")

        self.folder_path = ""
        self.output_path = ""
        self.dataframes = {}  # Almacenar archivos y sus hojas
        self.selected_files = set()  # Archivos confirmados
        self.selected_sheets = {}  # Hojas confirmadas de cada archivo

        # Frame principal
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(pady=10, padx=10, fill="both", expand=True)

        # Sección de archivos y pestañas
        self.files_frame = ctk.CTkFrame(self.frame)
        self.files_frame.pack(pady=5, padx=5, fill="x")

        # Lista de archivos Excel
        self.label_files = ctk.CTkLabel(self.files_frame, text="Archivos Excel:")
        self.label_files.grid(row=0, column=0, padx=5, pady=5)

        self.files_listbox = Listbox(self.files_frame, selectmode=MULTIPLE, height=5, width=40)
        self.files_listbox.grid(row=1, column=0, padx=5, pady=5)

        # Botón para confirmar archivos
        self.btn_confirm_files = ctk.CTkButton(self.files_frame, text="Confirmar Archivos", command=self.confirm_files)
        self.btn_confirm_files.grid(row=2, column=0, padx=5, pady=5)

        # Lista de pestañas (hojas de Excel)
        self.label_sheets = ctk.CTkLabel(self.files_frame, text="Hojas Disponibles:")
        self.label_sheets.grid(row=0, column=1, padx=5, pady=5)

        self.sheets_listbox = Listbox(self.files_frame, selectmode=MULTIPLE, height=5, width=40)
        self.sheets_listbox.grid(row=1, column=1, padx=5, pady=5)

        # Botón para confirmar hojas
        self.btn_confirm_sheets = ctk.CTkButton(self.files_frame, text="Confirmar Hojas", command=self.confirm_sheets)
        self.btn_confirm_sheets.grid(row=2, column=1, padx=5, pady=5)

        # Botón para seleccionar carpeta
        self.btn_select = ctk.CTkButton(self.frame, text="Seleccionar carpeta", command=self.load_folder)
        self.btn_select.pack(pady=10)

        # Botón de exportación
        self.btn_export = ctk.CTkButton(self.frame, text="Exportar archivo", command=self.export_file)
        self.btn_export.pack(pady=10)

        # Vista previa
        self.tree = ctk.CTkTextbox(self.frame, height=200, width=800)
        self.tree.pack(pady=5, padx=5, fill="both", expand=True)

    def load_folder(self):
        """Carga los archivos Excel y permite seleccionarlos sin perder selecciones previas."""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path = folder_selected
            files = [f for f in os.listdir(folder_selected) if f.endswith(('.xlsx', '.xls'))]
            if files:
                self.files_listbox.delete(0, "end")
                for file in files:
                    file_path = os.path.join(self.folder_path, file)
                    try:
                        xls = pd.ExcelFile(file_path)
                        self.files_listbox.insert("end", file)
                        self.dataframes[file] = xls
                        if file not in self.selected_sheets:
                            self.selected_sheets[file] = set()  # Inicializa selección vacía
                    except Exception as e:
                        messagebox.showwarning("Error al cargar", f"No se pudo leer {file}: {e}")
            else:
                messagebox.showwarning("Sin archivos", "No se encontraron archivos Excel en la carpeta seleccionada.")

    def confirm_files(self):
        """Confirma la selección de archivos antes de seleccionar hojas."""
        selected_indices = self.files_listbox.curselection()
        selected_files = {self.files_listbox.get(i) for i in selected_indices}
        
        if not selected_files:
            messagebox.showwarning("Selección inválida", "Seleccione al menos un archivo antes de confirmar.")
            return

        self.selected_files = selected_files  # Fijar archivos seleccionados
        self.sheets_listbox.delete(0, "end")  # Limpiar la lista de hojas

        # Agregar hojas de los archivos seleccionados
        for file in self.selected_files:
            if file in self.dataframes:
                sheets = self.dataframes[file].sheet_names
                for sheet in sheets:
                    self.sheets_listbox.insert("end", f"{file} -> {sheet}")

        messagebox.showinfo("Confirmación", f"Archivos seleccionados: {', '.join(self.selected_files)}")

    def confirm_sheets(self):
        """Confirma la selección de hojas antes de la exportación."""
        selected_indices = self.sheets_listbox.curselection()
        selected_sheets = [self.sheets_listbox.get(i) for i in selected_indices]

        if not selected_sheets:
            messagebox.showwarning("Selección inválida", "Seleccione al menos una hoja antes de confirmar.")
            return

        for sheet_entry in selected_sheets:
            file, sheet = sheet_entry.split(" -> ")
            if file in self.selected_sheets:
                self.selected_sheets[file].add(sheet)

        messagebox.showinfo("Confirmación", f"Hojas seleccionadas:\n" + "\n".join(selected_sheets))

    def export_file(self):
        """Exporta las hojas confirmadas de los archivos seleccionados."""
        if not self.selected_files:
            messagebox.showwarning("Selección inválida", "Seleccione y confirme al menos un archivo.")
            return

        selected_data = []
        for file in self.selected_files:
            if file in self.selected_sheets:
                for sheet in self.selected_sheets[file]:
                    try:
                        df = self.dataframes[file].parse(sheet)
                        selected_data.append(df)
                    except Exception as e:
                        messagebox.showwarning("Error", f"No se pudo leer {file} -> {sheet}: {e}")

        if not selected_data:
            messagebox.showwarning("Error", "No se seleccionó ninguna hoja válida para exportar.")
            return

        final_df = pd.concat(selected_data, ignore_index=True)
        self.display_preview(final_df)

        self.output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("TSV", "*.tsv")])
        if not self.output_path:
            return

        export_format = "Excel"
        if export_format == "Excel":
            final_df.to_excel(self.output_path, index=False)
        elif export_format == "CSV":
            final_df.to_csv(self.output_path, index=False, sep=",")
        elif export_format == "TSV":
            final_df.to_csv(self.output_path, index=False, sep="\t")

        messagebox.showinfo("Exportación completada", "El archivo se ha guardado correctamente.")

    def display_preview(self, df):
        """Muestra una vista previa de los primeros 10 registros."""
        preview_text = df.head(10).to_string(index=False)
        self.tree.delete("1.0", "end")
        self.tree.insert("1.0", preview_text)

if __name__ == "__main__":
    app = ExcelProcessorApp()
    app.mainloop()