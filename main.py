import sys
import pandas as pd
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import subprocess
import locale
from datetime import datetime

def obtener_ruta_guardado():
    escritorio = Path.home() / "Desktop"
    documentos = Path.home() / "Documents"
    return escritorio if escritorio.exists() else documentos if documentos.exists() else Path.home()

ruta_guardado = obtener_ruta_guardado()
archivo_json = "expse_montos.json"

if os.path.exists(archivo_json):
    with open(archivo_json, "r") as file:
        expse_montos = json.load(file)
else:
    expse_montos = {"EXPSE 1": "", "EXPSE 2": "", "EXPSE 4": "", "EXPSE 6": "", "EXPSE 7": ""}

ruta_archivo_seleccionado = None
ruta_archivo_modificado = None 

def seleccionar_archivo():
    global ruta_archivo_seleccionado
    ruta_archivo_seleccionado = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if ruta_archivo_seleccionado:
        lbl_archivo.config(text=f"Archivo seleccionado: {os.path.basename(ruta_archivo_seleccionado)}")

def actualizar_expse_montos():
    for expse, entry in entries.items():
        expse_montos[expse] = entry.get()
    
    with open(archivo_json, "w") as file:
        json.dump(expse_montos, file, indent=4)

def obtener_nombre_archivo_unico(base_path):
    locale.setlocale(locale.LC_TIME, "es_ES.utf8")

    hoy = datetime.now()

    if hoy.month == 1:
        mes_anterior = 12
        año = hoy.year - 1
    else:
        mes_anterior = hoy.month - 1
        año = hoy.year

    fecha_mes_anterior = datetime(año, mes_anterior, 1)
    nombre_mes = fecha_mes_anterior.strftime("%B").capitalize()

    nombre_base = f"Liquidación {nombre_mes} {año}.xlsx"
    ruta_archivo = base_path / nombre_base
    contador = 1

    while ruta_archivo.exists():
        nombre_modificado = f"Liquidación {nombre_mes} {año} ({contador}).xlsx"
        ruta_archivo = base_path / nombre_modificado
        contador += 1

    return ruta_archivo


def procesar_excel():
    global ruta_archivo_seleccionado, ruta_archivo_modificado

    if not ruta_archivo_seleccionado:
        messagebox.showerror("Error", "Primero selecciona un archivo de Excel")
        return

    try:
        actualizar_expse_montos() 

        df = pd.read_excel(ruta_archivo_seleccionado, header=0)
        columnas_deseadas = ["Fecha", "Profesional", "HC", "Trabajador", "Prestación"]

        if all(col in df.columns for col in columnas_deseadas):
            df = df[columnas_deseadas]
        else:
            columnas_indices = {"Fecha": 1, "Profesional": 2, "HC": 3, "Trabajador": 4, "Prestación": 7}
            df = df.iloc[:, list(columnas_indices.values())]
    
        if not all(col in df.columns for col in columnas_deseadas):
            df.columns = columnas_deseadas

        df["Monto"] = df["Prestación"].astype(str).str.strip().map(expse_montos).fillna(0)
        df["Monto"] = pd.to_numeric(df["Monto"], errors="coerce").fillna(0)
        df = df.drop(index=1).reset_index(drop=True)

        total_monto = df["Monto"].sum()
        df.loc[len(df)] = ["TOTAL", "", "", "", "", total_monto]

        ruta_archivo_modificado = obtener_nombre_archivo_unico(ruta_guardado)
        
        df = df.iloc[1:]
        df.to_excel(ruta_archivo_modificado, index=False, header=True)

        messagebox.showinfo("Éxito", f"Archivo guardado en: {ruta_archivo_modificado}")

        btn_abrir_ubicacion.config(state=tk.NORMAL)

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo procesar el archivo\n{e}")


def abrir_ubicacion():
    if ruta_archivo_modificado and os.path.exists(ruta_archivo_modificado):
        carpeta = os.path.dirname(ruta_archivo_modificado)
        subprocess.run(["explorer", carpeta] if os.name == "nt" else ["xdg-open", carpeta])
    else:
        messagebox.showwarning("Aviso", "No hay archivo generado aún.")

# --- Interfaz Gráfica --- #
root = tk.Tk()
root.title("Cargar y Modificar Excel")
root.geometry("500x400")

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")

icon_path = os.path.join(base_path, "icon.ico")
root.iconbitmap(icon_path)


root.update_idletasks()
ancho = root.winfo_width()
alto = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (ancho // 2)
y = (root.winfo_screenheight() // 2) - (alto // 2)
root.geometry(f"+{x}+{y}")

frame_contenedor = tk.Frame(root)
frame_contenedor.pack(expand=True, fill="both")

frame_interior = tk.Frame(frame_contenedor)
frame_interior.pack(expand=True) 

btn_cargar = tk.Button(frame_interior, text="Seleccionar Planilla", width=25, command=seleccionar_archivo)
btn_cargar.pack(pady=10)

lbl_archivo = tk.Label(frame_interior, text="Ningún archivo seleccionado", fg="blue")
lbl_archivo.pack()

entries = {}

for expse in expse_montos:
    frame = tk.Frame(frame_interior)
    frame.pack(pady=2)
    tk.Label(frame, text=f"{expse}:").pack(side=tk.LEFT, padx=5)
    entry = tk.Entry(frame, width=10)
    entry.insert(0, expse_montos[expse]) 
    entry.pack(side=tk.LEFT)
    entries[expse] = entry

btn_procesar = tk.Button(frame_interior, text="Procesar Planilla", width=25, command=procesar_excel)
btn_procesar.pack(pady=10)

btn_abrir_ubicacion = tk.Button(frame_interior, text="Abrir Ubicación de Planilla Modificada", width=30, command=abrir_ubicacion, state=tk.DISABLED)
btn_abrir_ubicacion.pack(pady=10)

root.mainloop()
