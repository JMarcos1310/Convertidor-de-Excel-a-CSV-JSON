import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import re

class ExcelConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertidor de Excel a CSV / JSON")
        self.root.geometry("640x520")
        self.root.configure(bg="#FFFFFF")
        self.excel_data = {}

        self.setup_styles()
        self.create_widgets()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        color_fondo = "#FFFFFF"
        color_texto = "#4B1C1C"
        color_boton = "#8B0000"
        color_hover = "#A52A2A"
        style.configure("TLabel", font=("Segoe UI", 10), background=color_fondo, foreground=color_texto)
        style.configure("TButton", font=("Segoe UI", 10, "bold"), background=color_boton, foreground="white", borderwidth=0, padding=6)
        style.map("TButton", background=[("active", color_hover)])
        style.configure("TCheckbutton", background=color_fondo, foreground=color_texto, font=("Segoe UI", 10))
        style.configure("TFrame", background=color_fondo)

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(expand=True)

        ttk.Label(frame, text="Convertidor de Excel a CSV / JSON", font=("Segoe UI", 14, "bold"), foreground="#8B0000").grid(row=0, column=0, columnspan=3, pady=10)

        ttk.Label(frame, text="Archivo Excel:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.entrada_archivo = ttk.Entry(frame, width=45)
        self.entrada_archivo.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Buscar", command=self.seleccionar_archivo).grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(frame, text="Carpeta de destino:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.entrada_directorio = ttk.Entry(frame, width=45)
        self.entrada_directorio.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Seleccionar", command=self.seleccionar_directorio).grid(row=2, column=2, padx=5, pady=5)

        ttk.Label(frame, text="Nombre del archivo:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.entrada_nombre = ttk.Entry(frame, width=45)
        self.entrada_nombre.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="Formatos de salida:").grid(row=4, column=0, sticky="ne", padx=5, pady=10)
        formato_frame = ttk.Frame(frame)
        formato_frame.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        self.var_csv = tk.BooleanVar()
        self.var_json = tk.BooleanVar()
        ttk.Checkbutton(formato_frame, text="CSV", variable=self.var_csv).grid(row=0, column=0, sticky="w", padx=5)
        ttk.Checkbutton(formato_frame, text="JSON", variable=self.var_json).grid(row=0, column=1, sticky="w", padx=10)

        ttk.Button(frame, text="Convertir archivo", command=self.convertir_excel).grid(row=5, column=0, columnspan=3, pady=30)

    def seleccionar_archivo(self):
        archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx *.xls")])
        if archivo:
            self.entrada_archivo.delete(0, tk.END)
            self.entrada_archivo.insert(0, archivo)
            self.cargar_hojas(archivo)

    def seleccionar_directorio(self):
        carpeta = filedialog.askdirectory()
        if carpeta:
            self.entrada_directorio.delete(0, tk.END)
            self.entrada_directorio.insert(0, carpeta)

    def cargar_hojas(self, ruta_archivo):
        try:
            self.excel_data = pd.read_excel(ruta_archivo, sheet_name=None)
            self.excel_data = {n: h for n, h in self.excel_data.items() if not h.dropna(how='all').empty}
            if not self.excel_data:
                messagebox.showwarning("Advertencia", "El archivo no contiene hojas con datos.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{str(e)}")

    def convertir_excel(self):
        archivo = self.entrada_archivo.get()
        carpeta_destino = self.entrada_directorio.get()
        nombre_personalizado = self.entrada_nombre.get().strip()
        if not archivo or not os.path.isfile(archivo):
            messagebox.showwarning("Advertencia", "Selecciona un archivo Excel válido.")
            return
        if not carpeta_destino or not os.path.isdir(carpeta_destino):
            messagebox.showwarning("Advertencia", "Selecciona una carpeta de destino válida.")
            return
        if not nombre_personalizado:
            messagebox.showwarning("Advertencia", "Debes asignar un nombre al archivo de salida.")
            return
        if not any([self.var_csv.get(), self.var_json.get()]):
            messagebox.showwarning("Advertencia", "Selecciona al menos un formato de salida.")
            return
        if not self.excel_data:
            messagebox.showwarning("Advertencia", "No hay hojas con datos para convertir.")
            return

        nombre_archivo = re.sub(r'[\\/*?:"<>|]', "_", nombre_personalizado)
        try:
            df_combined = pd.concat(self.excel_data.values(), ignore_index=True)
            ruta_base = os.path.join(carpeta_destino, nombre_archivo)
            if self.var_csv.get():
                df_combined.to_csv(ruta_base + ".csv", index=False)
            if self.var_json.get():
                df_combined.to_json(ruta_base + ".json", orient="records", indent=4, force_ascii=False)
            messagebox.showinfo("Éxito", f"Archivo convertido y guardado como:\n{ruta_base}")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelConverterApp(root)
    root.mainloop()