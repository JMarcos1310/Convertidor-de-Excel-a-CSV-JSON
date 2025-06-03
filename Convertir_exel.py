import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import re

# Variable global para guardar las hojas con contenido
excel_data = {}

# Función para seleccionar el archivo Excel
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx *.xls")])
    if archivo:
        entrada_archivo.delete(0, tk.END)
        entrada_archivo.insert(0, archivo)
        cargar_hojas(archivo)

# Función para seleccionar la carpeta donde se guardará el archivo convertido
def seleccionar_directorio():
    carpeta = filedialog.askdirectory()
    if carpeta:
        entrada_directorio.delete(0, tk.END)
        entrada_directorio.insert(0, carpeta)

# Carga el archivo y filtra solo las hojas con datos
def cargar_hojas(ruta_archivo):
    try:
        global excel_data
        excel_data = pd.read_excel(ruta_archivo, sheet_name=None)
        # Se eliminan las hojas que están vacías
        excel_data = {nombre: hoja for nombre, hoja in excel_data.items() if not hoja.dropna(how='all').empty}
        if not excel_data:
            messagebox.showwarning("Advertencia", "El archivo no contiene hojas con datos.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{str(e)}")

# Función que convierte el Excel a JSON y/o CSV según la selección
def convertir_excel():
    archivo = entrada_archivo.get()
    carpeta_destino = entrada_directorio.get()
    nombre_personalizado = entrada_nombre.get().strip()

    # Validaciones
    if not archivo or not os.path.isfile(archivo):
        messagebox.showwarning("Advertencia", "Selecciona un archivo Excel válido.")
        return
    if not carpeta_destino or not os.path.isdir(carpeta_destino):
        messagebox.showwarning("Advertencia", "Selecciona una carpeta de destino válida.")
        return
    if not nombre_personalizado:
        messagebox.showwarning("Advertencia", "Debes asignar un nombre al archivo de salida.")
        return
    if not any([var_csv.get(), var_json.get()]):
        messagebox.showwarning("Advertencia", "Selecciona al menos un formato de salida.")
        return
    if not excel_data:
        messagebox.showwarning("Advertencia", "No hay hojas con datos para convertir.")
        return

    # Eliminar caracteres no válidos para nombres de archivos
    nombre_archivo = re.sub(r'[\\/*?:"<>|]', "_", nombre_personalizado)

    try:
        # Combina todas las hojas con contenido
        df_combined = pd.concat(excel_data.values(), ignore_index=True)

        # Ruta base para guardar los archivos
        ruta_base = os.path.join(carpeta_destino, nombre_archivo)

        if var_csv.get():
            df_combined.to_csv(ruta_base + ".csv", index=False)
        if var_json.get():
            df_combined.to_json(ruta_base + ".json", orient="records", indent=4, force_ascii=False)

        messagebox.showinfo("Éxito", f"Archivo convertido y guardado como:\n{ruta_base}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")

# ---------------- Interfaz Gráfica ---------------- #

ventana = tk.Tk()
ventana.title("Convertidor de Excel a CSV / JSON")
ventana.geometry("640x520")
ventana.configure(bg="#F5F5DC")  # Color beige

# Estilos modernos
style = ttk.Style()
style.theme_use("clam")

color_fondo = "#F5F5DC"     # Beige
color_texto = "#4B1C1C"     # Marrón oscuro
color_boton = "#8B0000"     # Rojo vino
color_hover = "#A52A2A"     # Vino más claro

style.configure("TLabel", font=("Segoe UI", 10), background=color_fondo, foreground=color_texto)
style.configure("TButton", font=("Segoe UI", 10, "bold"), background=color_boton, foreground="white", borderwidth=0, padding=6)
style.map("TButton", background=[("active", color_hover)])
style.configure("TCheckbutton", background=color_fondo, foreground=color_texto, font=("Segoe UI", 10))
style.configure("TFrame", background=color_fondo)

# Frame principal
frame = ttk.Frame(ventana, padding=20)
frame.pack(expand=True)

# Título
ttk.Label(frame, text="Convertidor de Excel a CSV / JSON", font=("Segoe UI", 14, "bold"), foreground=color_boton).grid(row=0, column=0, columnspan=3, pady=10)

# Entrada: Archivo Excel
ttk.Label(frame, text="Archivo Excel:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
entrada_archivo = ttk.Entry(frame, width=45)
entrada_archivo.grid(row=1, column=1, padx=5, pady=5)
ttk.Button(frame, text="Buscar", command=seleccionar_archivo).grid(row=1, column=2, padx=5, pady=5)

# Entrada: Carpeta de destino
ttk.Label(frame, text="Carpeta de destino:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
entrada_directorio = ttk.Entry(frame, width=45)
entrada_directorio.grid(row=2, column=1, padx=5, pady=5)
ttk.Button(frame, text="Seleccionar", command=seleccionar_directorio).grid(row=2, column=2, padx=5, pady=5)

# Entrada: Nombre del archivo
ttk.Label(frame, text="Nombre del archivo:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
entrada_nombre = ttk.Entry(frame, width=45)
entrada_nombre.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="w")

# Selección de formatos
ttk.Label(frame, text="Formatos de salida:").grid(row=4, column=0, sticky="ne", padx=5, pady=10)
formato_frame = ttk.Frame(frame)
formato_frame.grid(row=4, column=1, sticky="w", padx=5, pady=5)

var_csv = tk.BooleanVar()
var_json = tk.BooleanVar()
ttk.Checkbutton(formato_frame, text="CSV", variable=var_csv).grid(row=0, column=0, sticky="w", padx=5)
ttk.Checkbutton(formato_frame, text="JSON", variable=var_json).grid(row=0, column=1, sticky="w", padx=10)

# Botón principal de conversión
ttk.Button(frame, text="Convertir archivo", command=convertir_excel).grid(row=5, column=0, columnspan=3, pady=30)

# Ejecutar la ventana
ventana.mainloop()
