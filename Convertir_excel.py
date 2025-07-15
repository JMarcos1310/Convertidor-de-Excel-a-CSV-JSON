import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import re
import threading
import json
import xml.etree.ElementTree as ET
import configparser

# Arrastrar y soltar
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# --- Internacionalización ---
LANGS = {
    "es": {
        "title": "Convertidor de Excel a CSV / JSON",
        "file": "Archivo Excel:",
        "folder": "Carpeta de destino:",
        "name": "Nombre del archivo:",
        "formats": "Formatos de salida:",
        "csv": "CSV",
        "json": "JSON",
        "xml": "XML",
        "sheets": "Selecciona hojas:",
        "convert": "Convertir archivo",
        "history": "Historial:",
        "help": "Ayuda",
        "success": "Éxito",
        "saved": "Archivo convertido y guardado como:\n{}",
        "warn_file": "Selecciona un archivo Excel válido.",
        "warn_folder": "Selecciona una carpeta de destino válida.",
        "warn_name": "Debes asignar un nombre al archivo de salida.",
        "warn_format": "Selecciona al menos un formato de salida.",
        "warn_data": "No hay hojas con datos para convertir.",
        "warn_sheet": "Selecciona al menos una hoja para convertir.",
        "warn_empty": "El archivo no contiene hojas con datos.",
        "error_load": "No se pudo cargar el archivo:\n{}",
        "error_save": "Ocurrió un error:\n{}",
        "help_text": "1. Selecciona un archivo Excel\n2. Elige carpeta y nombre\n3. Selecciona hojas y formato\n4. ¡Convierte!"
    },
    "en": {
        "title": "Excel to CSV / JSON Converter",
        "file": "Excel file:",
        "folder": "Destination folder:",
        "name": "File name:",
        "formats": "Output formats:",
        "csv": "CSV",
        "json": "JSON",
        "xml": "XML",
        "sheets": "Select sheets:",
        "convert": "Convert file",
        "history": "History:",
        "help": "Help",
        "success": "Success",
        "saved": "File converted and saved as:\n{}",
        "warn_file": "Select a valid Excel file.",
        "warn_folder": "Select a valid destination folder.",
        "warn_name": "You must assign a name to the output file.",
        "warn_format": "Select at least one output format.",
        "warn_data": "No sheets with data to convert.",
        "warn_sheet": "Select at least one sheet to convert.",
        "warn_empty": "The file contains no sheets with data.",
        "error_load": "Could not load file:\n{}",
        "error_save": "An error occurred:\n{}",
        "help_text": "1. Select an Excel file\n2. Choose folder and name\n3. Select sheets and format\n4. Convert!"
    }
}

class ExcelConverterApp:
    def __init__(self, root):
        self.root = root
        self.lang = "es"  # Cambia a "en" para inglés
        self.text = LANGS[self.lang]
        self.root.title(self.text["title"])
        self.root.geometry("700x600")
        self.root.configure(bg="#FFFFFF")
        self.excel_data = {}
        self.historial = []
        self.config_file = "config.ini"
        self.load_config()

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
        style.configure("TEntry", padding=4)

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(expand=True, fill="both")

        # Título
        ttk.Label(frame, text=self.text["title"], font=("Segoe UI", 16, "bold"), foreground="#8B0000").grid(row=0, column=0, columnspan=4, pady=10)

        # Archivo Excel
        ttk.Label(frame, text=self.text["file"]).grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.entrada_archivo = ttk.Entry(frame, width=45)
        self.entrada_archivo.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="🔍", command=self.seleccionar_archivo).grid(row=1, column=2, padx=5, pady=5)
        # Arrastrar y soltar
        if DND_AVAILABLE:
            self.entrada_archivo.drop_target_register(DND_FILES)
            self.entrada_archivo.dnd_bind('<<Drop>>', self.drop_archivo)
        self.add_tooltip(self.entrada_archivo, "Arrastra aquí tu archivo Excel")

        # Carpeta de destino
        ttk.Label(frame, text=self.text["folder"]).grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.entrada_directorio = ttk.Entry(frame, width=45)
        self.entrada_directorio.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(frame, text="📁", command=self.seleccionar_directorio).grid(row=2, column=2, padx=5, pady=5)
        self.add_tooltip(self.entrada_directorio, "Carpeta donde se guardará el archivo convertido")

        # Nombre del archivo
        ttk.Label(frame, text=self.text["name"]).grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.entrada_nombre = ttk.Entry(frame, width=45)
        self.entrada_nombre.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        self.add_tooltip(self.entrada_nombre, "Nombre del archivo de salida (sin extensión)")

        # Formatos de salida
        ttk.Label(frame, text=self.text["formats"]).grid(row=4, column=0, sticky="ne", padx=5, pady=10)
        formato_frame = ttk.Frame(frame)
        formato_frame.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        self.var_csv = tk.BooleanVar()
        self.var_json = tk.BooleanVar()
        self.var_xml = tk.BooleanVar()
        ttk.Checkbutton(formato_frame, text=self.text["csv"], variable=self.var_csv).grid(row=0, column=0, sticky="w", padx=5)
        ttk.Checkbutton(formato_frame, text=self.text["json"], variable=self.var_json).grid(row=0, column=1, sticky="w", padx=10)
        ttk.Checkbutton(formato_frame, text=self.text["xml"], variable=self.var_xml).grid(row=0, column=2, sticky="w", padx=10)

        # Selección de hojas
        ttk.Label(frame, text=self.text["sheets"]).grid(row=5, column=0, sticky="ne", padx=5, pady=5)
        self.listbox_hojas = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=5, exportselection=False)
        self.listbox_hojas.grid(row=5, column=1, columnspan=2, sticky="w", padx=5, pady=5)
        self.add_tooltip(self.listbox_hojas, "Selecciona las hojas que deseas convertir")

        # Botón convertir
        ttk.Button(frame, text=self.text["convert"], command=self.convertir_excel).grid(row=6, column=0, columnspan=3, pady=20)

        # Barra de progreso
        self.progress = ttk.Progressbar(frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky="ew", padx=5, pady=10)

        # Historial de conversiones
        ttk.Label(frame, text=self.text["history"]).grid(row=8, column=0, sticky="ne", padx=5, pady=5)
        self.listbox_historial = tk.Listbox(frame, height=5)
        self.listbox_historial.grid(row=8, column=1, columnspan=2, sticky="w", padx=5, pady=5)
        self.add_tooltip(self.listbox_historial, "Historial de conversiones exitosas")

        # Botón de ayuda
        ttk.Button(frame, text=self.text["help"], command=self.mostrar_ayuda).grid(row=9, column=2, sticky="e", pady=10)

        # Cargar última carpeta usada
        if self.last_dir:
            self.entrada_directorio.insert(0, self.last_dir)

    # --- Tooltips ---
    def add_tooltip(self, widget, text):
        def on_enter(event):
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + 20
            self.tooltip.wm_geometry(f"+{x}+{y}")
            label = tk.Label(self.tooltip, text=text, background="#FFFFE0", relief="solid", borderwidth=1, font=("Segoe UI", 9))
            label.pack()
        def on_leave(event):
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    # --- Arrastrar y soltar ---
    def drop_archivo(self, event):
        archivo = event.data.strip('{}')
        self.entrada_archivo.delete(0, tk.END)
        self.entrada_archivo.insert(0, archivo)
        self.cargar_hojas(archivo)

    # --- Selección de archivo ---
    def seleccionar_archivo(self):
        archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx *.xls")])
        if archivo:
            self.entrada_archivo.delete(0, tk.END)
            self.entrada_archivo.insert(0, archivo)
            self.cargar_hojas(archivo)

    # --- Selección de carpeta ---
    def seleccionar_directorio(self):
        carpeta = filedialog.askdirectory()
        if carpeta:
            self.entrada_directorio.delete(0, tk.END)
            self.entrada_directorio.insert(0, carpeta)
            self.save_config(carpeta)

    # --- Cargar hojas del Excel ---
    def cargar_hojas(self, ruta_archivo):
        try:
            self.excel_data = pd.read_excel(ruta_archivo, sheet_name=None)
            self.excel_data = {n: h for n, h in self.excel_data.items() if not h.dropna(how='all').empty}
            self.listbox_hojas.delete(0, tk.END)
            for hoja in self.excel_data:
                self.listbox_hojas.insert(tk.END, hoja)
            if not self.excel_data:
                messagebox.showwarning(self.text["title"], self.text["warn_empty"])
        except Exception as e:
            messagebox.showerror(self.text["title"], self.text["error_load"].format(str(e)))

    # --- Conversión principal (con barra de progreso e hilos) ---
    def convertir_excel(self):
        archivo = self.entrada_archivo.get()
        carpeta_destino = self.entrada_directorio.get()
        nombre_personalizado = self.entrada_nombre.get().strip()
        if not archivo or not os.path.isfile(archivo):
            messagebox.showwarning(self.text["title"], self.text["warn_file"])
            return
        if not carpeta_destino or not os.path.isdir(carpeta_destino):
            messagebox.showwarning(self.text["title"], self.text["warn_folder"])
            return
        if not nombre_personalizado:
            messagebox.showwarning(self.text["title"], self.text["warn_name"])
            return
        if not any([self.var_csv.get(), self.var_json.get(), self.var_xml.get()]):
            messagebox.showwarning(self.text["title"], self.text["warn_format"])
            return
        if not self.excel_data:
            messagebox.showwarning(self.text["title"], self.text["warn_data"])
            return

        seleccion = self.listbox_hojas.curselection()
        if not seleccion:
            messagebox.showwarning(self.text["title"], self.text["warn_sheet"])
            return
        hojas_seleccionadas = [self.listbox_hojas.get(i) for i in seleccion]
        nombre_archivo = re.sub(r'[\\/*?:"<>|]', "_", nombre_personalizado)
        ruta_base = os.path.join(carpeta_destino, nombre_archivo)

        # Inicia barra de progreso y ejecuta en hilo
        self.progress.start()
        threading.Thread(target=self._convertir_thread, args=(hojas_seleccionadas, ruta_base)).start()

    def _convertir_thread(self, hojas_seleccionadas, ruta_base):
        try:
            df_combined = pd.concat([self.excel_data[h] for h in hojas_seleccionadas], ignore_index=True)
            formatos = []
            if self.var_csv.get():
                df_combined.to_csv(ruta_base + ".csv", index=False)
                formatos.append("CSV")
            if self.var_json.get():
                df_combined.to_json(ruta_base + ".json", orient="records", indent=4, force_ascii=False)
                formatos.append("JSON")
            if self.var_xml.get():
                self.df_to_xml(df_combined, ruta_base + ".xml")
                formatos.append("XML")
            self.root.after(0, lambda: self.listbox_historial.insert(0, f"{os.path.basename(ruta_base)} ({', '.join(formatos)})"))
            self.root.after(0, lambda: messagebox.showinfo(self.text["success"], self.text["saved"].format(ruta_base)))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror(self.text["title"], self.text["error_save"].format(str(e))))
        finally:
            self.root.after(0, self.progress.stop)

    # --- Exportar a XML ---
    def df_to_xml(self, df, filename):
        root = ET.Element("data")
        for _, row in df.iterrows():
            item = ET.SubElement(root, "row")
            for col, val in row.items():
                child = ET.SubElement(item, str(col))
                child.text = str(val)
        tree = ET.ElementTree(root)
        tree.write(filename, encoding="utf-8", xml_declaration=True)

    # --- Ayuda ---
    def mostrar_ayuda(self):
        messagebox.showinfo(self.text["help"], self.text["help_text"])

    # --- Configuración: guardar/cargar última carpeta ---
    def save_config(self, carpeta):
        config = configparser.ConfigParser()
        config["DEFAULT"] = {"last_dir": carpeta}
        with open(self.config_file, "w") as f:
            config.write(f)

    def load_config(self):
        self.last_dir = ""
        if os.path.exists(self.config_file):
            config = configparser.ConfigParser()
            config.read(self.config_file)
            self.last_dir = config["DEFAULT"].get("last_dir", "")

if __name__ == "__main__":
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = ExcelConverterApp(root)
    root.mainloop()