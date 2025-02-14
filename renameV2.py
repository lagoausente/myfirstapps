#root.geometry("600x400")  # Ajusta estos valores según lo que necesites
import os
import tkinter as tk
from tkinter import filedialog, messagebox

class FileRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Renombrar Archivos")
        self.root.geometry("600x400")  # Ajustar tamaño
        self.root.resizable(False, False)  # Evitar cambios de tamaño        
        self.origin_dir = tk.StringVar()
        self.dest_dir = tk.StringVar()
        self.prefix = tk.StringVar()
        self.suffix = tk.StringVar()
        self.new_extension = tk.StringVar()
        self.root.update_idletasks()
        
        tk.Label(root, text="Directorio de Origen:").grid(row=0, column=0)
        tk.Entry(root, textvariable=self.origin_dir, width=40).grid(row=0, column=1)
        tk.Button(root, text="Seleccionar", command=self.select_origin_dir).grid(row=0, column=2)
        
        tk.Label(root, text="Directorio de Destino:").grid(row=1, column=0)
        tk.Entry(root, textvariable=self.dest_dir, width=40).grid(row=1, column=1)
        tk.Button(root, text="Seleccionar", command=self.select_dest_dir).grid(row=1, column=2)
        
        tk.Label(root, text="Prefijo:").grid(row=2, column=0)
        tk.Entry(root, textvariable=self.prefix, width=30).grid(row=2, column=1)
        
        tk.Label(root, text="Sufijo:").grid(row=3, column=0)
        tk.Entry(root, textvariable=self.suffix, width=30).grid(row=3, column=1)
        
        tk.Label(root, text="Nueva Extensión (opcional, sin punto):").grid(row=4, column=0)
        tk.Entry(root, textvariable=self.new_extension, width=30).grid(row=4, column=1)
        
        tk.Button(root, text="Renombrar Archivos", command=self.rename_files).grid(row=5, column=1)
    
    def select_origin_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.origin_dir.set(directory)
    
    def select_dest_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.dest_dir.set(directory)
    
    def rename_files(self):
        origin = self.origin_dir.get()
        dest = self.dest_dir.get()
        prefix = self.prefix.get()
        suffix = self.suffix.get()
        new_ext = self.new_extension.get().strip()
        
        if not origin or not dest:
            messagebox.showerror("Error", "Debe seleccionar los directorios de origen y destino.")
            return
        
        if not os.path.exists(dest):
            os.makedirs(dest)
        
        for filename in os.listdir(origin):
            old_path = os.path.join(origin, filename)
            if os.path.isfile(old_path):
                name, ext = os.path.splitext(filename)
                
                if new_ext:
                    ext = f".{new_ext}"
                
                new_name = f"{prefix}{name}{suffix}{ext}"
                new_path = os.path.join(dest, new_name)
                
                os.rename(old_path, new_path)
        
        messagebox.showinfo("Éxito", "Archivos renombrados correctamente.")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileRenamerApp(root)
    root.mainloop()
