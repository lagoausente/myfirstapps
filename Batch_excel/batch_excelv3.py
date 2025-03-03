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
        self.selected_files_sheets = {}  # Diccionario para almacenar archivos y sus hojas seleccionadas
        self.dataframes = []
        self.columns = []
        self.history = []
        self.files = []  # Lista para almacenar los nombres de los archivos Excel
        self.sheets = {}  # Diccionario para almacenar las hojas de cada archivo
        self.files_sheets = {}  # Diccionario para almacenar archivos y sus hojas seleccionadas


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

        # Listbox 1: Selección de archivos
        self.center_listbox = Listbox(self.listbox_frame, selectmode=MULTIPLE, height=10, width=25, font=self.font_style, exportselection=False)
        self.center_listbox.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.center_listbox.bind("<<ListboxSelect>>", self.load_sheets)

        # Listbox 2: Selección de hojas
        self.right_listbox = Listbox(self.listbox_frame, selectmode=MULTIPLE, height=10, width=20, font=self.font_style, exportselection=False)
        self.right_listbox.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.right_listbox.bind("<<ListboxSelect>>", self.load_columns)

        # Listbox 3: Selección de columnas
        self.column_listbox = Listbox(self.listbox_frame, selectmode=MULTIPLE, height=10, width=15, font=self.font_style, exportselection=False)
        self.column_listbox.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")

        # Botón para mover columna arriba
        self.btn_up = ctk.CTkButton(self.listbox_frame, text="Subir", command=self.move_column_up)
        self.btn_up.grid(row=1, column=2, padx=5, pady=2, sticky="ew")  # 📌 Alineado con columnas

        # Botón para mover columna abajo
        self.btn_down = ctk.CTkButton(self.listbox_frame, text="Bajar", command=self.move_column_down)
        self.btn_down.grid(row=2, column=2, padx=5, pady=2, sticky="ew")  # 📌 Debajo del otro botón


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
        
    def move_column_up(self):
        selection = self.column_listbox.curselection()
        if not selection:
            return
        for i in selection:
            if i > 0:  # No mover si está en la primera posición
                text = self.column_listbox.get(i)
                self.column_listbox.delete(i)
                self.column_listbox.insert(i - 1, text)
                self.column_listbox.selection_set(i - 1)

    def move_column_down(self):
        selection = self.column_listbox.curselection()
        if not selection:
            return
        for i in reversed(selection):  # Recorremos en orden inverso para evitar desorden
            if i < self.column_listbox.size() - 1:  # No mover si está en la última posición
                text = self.column_listbox.get(i)
                self.column_listbox.delete(i)
                self.column_listbox.insert(i + 1, text)
                self.column_listbox.selection_set(i + 1)
                



    def load_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path = folder_selected
            self.files = [f for f in os.listdir(folder_selected) if f.endswith(('.xlsx', '.xls'))]

            self.center_listbox.delete(0, "end")  # ✅ Borrar lista de archivos (están en center_listbox)
            self.sheets.clear()
            self.selected_files_sheets.clear()

            if self.files:
                for file in self.files:
                    self.center_listbox.insert("end", file)  # ✅ Insertamos archivos en `center_listbox`
                    file_path = os.path.join(self.folder_path, file)
                    xls = pd.ExcelFile(file_path)
                    self.sheets[file] = xls.sheet_names  # Guardamos las hojas disponibles

                messagebox.showinfo("Archivos cargados", f"Se han cargado {len(self.files)} archivos.")
            else:
                messagebox.showwarning("Sin archivos", "No se encontraron archivos Excel en la carpeta seleccionada.")


    def load_columns(self, event):
        self.dataframes = []  # Limpiar antes de cargar nuevas selecciones
        self.column_listbox.delete(0, "end")  # ✅ Ahora `column_listbox` almacena columnas

        print("Ejecutando load_columns...")  # 🛠 Línea de depuración

        for file_name_clean, selected_sheets in self.selected_files_sheets.items():
            # Recuperar el nombre de archivo original con su extensión
            file_name = next((f for f in self.files if f.startswith(file_name_clean)), None)
            if not file_name:
                print(f"⚠ No se encontró el archivo real para {file_name_clean}")
                continue

            file_path = os.path.join(self.folder_path, file_name)

            for sheet_name in selected_sheets:
                print(f"Cargando hoja '{sheet_name}' de '{file_name}'")  # 🛠 Depuración

                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    df = self.clean_dataframe(df)
                    self.dataframes.append(df)
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo cargar {sheet_name} de {file_name}.\n{str(e)}")

        # Si hay dataframes, actualizar la lista de columnas
        if self.dataframes:
            self.columns = list(self.dataframes[0].columns)
            print("Columnas cargadas:", self.columns)  # 🛠 Línea de depuración

            for col in self.columns:
                self.column_listbox.insert("end", col)  # ✅ Ahora columnas van en `column_listbox`

            messagebox.showinfo("Carga Completa", f"Se han cargado {len(self.dataframes)} hojas correctamente.")
        else:
            print("⚠ No se cargaron hojas, `self.dataframes` está vacío.")  # 🛠 Depuración
            messagebox.showwarning("Sin datos", "No se pudieron cargar datos de las hojas seleccionadas.")




    def load_sheets(self, event):
        selected_indices = self.center_listbox.curselection()
        self.right_listbox.delete(0, 'end')  # ✅ Limpiar lista de hojas

        self.selected_files_sheets.clear()

        for i in selected_indices:
            file_name = self.center_listbox.get(i)
            file_name_clean = file_name.replace(".xlsx", "").replace(".xls", "")  # ✅ Quitar extensión

            if file_name in self.sheets:
                for sheet in self.sheets[file_name]:
                    file_path = os.path.join(self.folder_path, file_name)
                    
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet)
                        if df.dropna(how="all").empty:  # ❌ Si la hoja está vacía, la ignoramos
                            continue
                        
                        sheet_label = f"{file_name_clean} - {sheet}"  # ✅ Usar nombre sin extensión
                        self.right_listbox.insert('end', sheet_label)
                    except Exception as e:
                        print(f"Error al leer {sheet} de {file_name}: {e}")

                self.selected_files_sheets[file_name] = self.sheets[file_name]

        self.right_listbox.bind("<<ListboxSelect>>", self.update_selected_sheets)




    def update_selected_sheets(self, event):
        selected_sheets = self.right_listbox.curselection()
        if not selected_sheets:
            return

        self.selected_files_sheets.clear()  # Limpiar selección previa

        for i in selected_sheets:
            sheet_label = self.right_listbox.get(i)  # Ejemplo: "archivo1 - Hoja1"
            file_name_clean, sheet_name = sheet_label.split(" - ")  # Separar archivo y hoja
            
            # Volver a agregar la extensión para que coincida con los nombres de archivo reales
            for file_name in self.files:
                if file_name.startswith(file_name_clean):  # Comparar sin la extensión
                    if file_name not in self.selected_files_sheets:
                        self.selected_files_sheets[file_name] = []
                    self.selected_files_sheets[file_name].append(sheet_name)
                    break  # Salimos del bucle al encontrar el archivo correcto

        print("Diccionario actualizado:", self.selected_files_sheets)  # 🛠 Depuración

        self.load_columns(None)  # Cargar columnas después de actualizar la selección






    def load_files(self, files):
        self.dataframes = []
        for file in files:
            file_path = os.path.join(self.folder_path, file)
            df = pd.read_excel(file_path)
            df = self.clean_dataframe(df)
            self.dataframes.append(df)
            print(f"Archivo {file} cargado con {len(df)} filas y {len(df.columns)} columnas.")  # Línea de depuración

        if self.dataframes:
            self.columns = list(self.dataframes[0].columns)
            self.column_listbox.delete(0, "end")
            for col in self.columns:
                self.column_listbox.insert("end", col)
            messagebox.showinfo("Archivos cargados", f"Se han cargado {len(self.dataframes)} archivos correctamente.")

    def select_output(self):
        self.output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("TSV", "*.tsv")])
        print(f"Ruta de salida seleccionada: {self.output_path}")  # Línea de depuración
        if self.output_path:
            messagebox.showinfo("Ruta seleccionada", f"El archivo se guardará en: {self.output_path}")

    def export_file(self):
        if not self.dataframes:
            messagebox.showwarning("Sin archivos", "Seleccione una carpeta antes de exportar.")
            return

        # Obtener columnas en el orden actual del Listbox
        ordered_columns = [self.column_listbox.get(i) for i in range(self.column_listbox.size())]

        if not ordered_columns:
            messagebox.showwarning("Selección inválida", "Seleccione al menos una columna para exportar.")
            return

        filter_text = self.filter_entry.get()
        transform_option = self.transform_var.get()

        self.progress.set(0.2)

        processed_data = []
        for df in self.dataframes:
            df = df[ordered_columns]  # 🔹 Aplicar el orden personalizado del usuario

            # Aplicar transformaciones
            df = df.applymap(lambda x: self.remove_extra_spaces(x) if isinstance(x, str) else x)
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
