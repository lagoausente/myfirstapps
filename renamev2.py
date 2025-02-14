import tkinter as tk
from tkinter import filedialog, messagebox
import os

class FileRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Renombrar Archivos")
        self.root.geometry("600x500")

        # Variables
        self.origin_dir = tk.StringVar()
        self.prefix = tk.StringVar()
        self.suffix = tk.StringVar()
        self.new_extension = tk.StringVar()

        # Selección de carpeta
        tk.Label(root, text="Carpeta de origen:").pack()
        tk.Entry(root, textvariable=self.origin_dir, width=50).pack()
        tk.Button(root, text="Seleccionar", command=self.select_folder).pack()

        # Opciones de renombrado
        tk.Label(root, text="Prefijo:").pack()
        tk.Entry(root, textvariable=self.prefix).pack()

        tk.Label(root, text="Sufijo:").pack()
        tk.Entry(root, textvariable=self.suffix).pack()

        tk.Label(root, text="Nueva extensión (opcional, sin punto):").pack()
        tk.Entry(root, textvariable=self.new_extension).pack()

        # Botón Previsualizar
        tk.Button(root, text="Previsualizar", command=self.preview_changes).pack()

        # Lista para mostrar previsualización
        self.preview_list = tk.Listbox(root, width=80, height=10)
        self.preview_list.pack()

        # Botón para aplicar cambios
        tk.Button(root, text="Renombrar", command=self.rename_files).pack()

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        self.origin_dir.set(folder_selected)

    def preview_changes(self):
        """ Muestra una previsualización de los cambios sin aplicar el renombrado. """
        self.preview_list.delete(0, tk.END)  # Limpiar lista

        folder = self.origin_dir.get()
        prefix = self.prefix.get()
        suffix = self.suffix.get()
        new_ext = self.new_extension.get()

        if not folder:
            messagebox.showerror("Error", "Selecciona una carpeta primero.")
            return

        for filename in os.listdir(folder):
            old_path = os.path.join(folder, filename)

            if os.path.isfile(old_path):
                name, ext = os.path.splitext(filename)
                
                # Construir nuevo nombre
                new_name = f"{prefix}{name}{suffix}"
                if new_ext:
                    new_name += f".{new_ext}"
                else:
                    new_name += ext  # Mantener la extensión original

                self.preview_list.insert(tk.END, f"{filename}  →  {new_name}")

    def rename_files(self):
        """ Aplica los cambios de nombre en la carpeta. """
        folder = self.origin_dir.get()
        prefix = self.prefix.get()
        suffix = self.suffix.get()
        new_ext = self.new_extension.get()

        if not folder:
            messagebox.showerror("Error", "Selecciona una carpeta primero.")
            return

        for filename in os.listdir(folder):
            old_path = os.path.join(folder, filename)

            if os.path.isfile(old_path):
                name, ext = os.path.splitext(filename)
                
                new_name = f"{prefix}{name}{suffix}"
                if new_ext:
                    new_name += f".{new_ext}"
                else:
                    new_name += ext

                new_path = os.path.join(folder, new_name)

                try:
                    os.rename(old_path, new_path)
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo renombrar {filename}: {e}")

        messagebox.showinfo("Completado", "Los archivos han sido renombrados.")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileRenamerApp(root)
    root.mainloop()
