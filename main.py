import sys
import pandas as pd
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import subprocess
from datetime import datetime

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

# Filas de datos del Excel de origen que se descartan antes de liquidar
FILAS_A_DESCARTAR = 2


def obtener_directorio_app():
    """Carpeta del .exe (si está empaquetado) o del script, independiente del directorio de trabajo."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def obtener_ruta_guardado():
    escritorio = Path.home() / "Desktop"
    documentos = Path.home() / "Documents"
    return escritorio if escritorio.exists() else documentos if documentos.exists() else Path.home()


ruta_guardado = obtener_ruta_guardado()
archivo_json = obtener_directorio_app() / "expse_montos.json"
archivo_json_anterior = Path("expse_montos.json")  # ubicación usada por versiones anteriores

expse_montos = {"EXPSE 1": "", "EXPSE 2": "", "EXPSE 4": "", "EXPSE 6": "", "EXPSE 7": ""}

for ruta in (archivo_json, archivo_json_anterior):
    if ruta.exists():
        try:
            with open(ruta, "r", encoding="utf-8") as file:
                expse_montos = json.load(file)
            break
        except (OSError, json.JSONDecodeError):
            pass

ruta_archivo_seleccionado = None
ruta_archivo_modificado = None


def periodo_mes_vencido():
    hoy = datetime.now()
    if hoy.month == 1:
        return 12, hoy.year - 1
    return hoy.month - 1, hoy.year


def seleccionar_archivo():
    global ruta_archivo_seleccionado
    ruta = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if ruta:
        ruta_archivo_seleccionado = ruta
        lbl_archivo.config(text=f"Archivo seleccionado: {os.path.basename(ruta)}")


def parsear_monto(texto):
    """Convierte el texto ingresado a número. Vacío = 0. Acepta coma decimal. Devuelve None si es inválido."""
    texto = texto.strip().replace(" ", "")
    if texto == "":
        return 0.0
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def leer_y_validar_montos():
    """Lee los montos de la interfaz. Devuelve (montos_numericos, errores)."""
    montos = {}
    errores = []
    for expse, entry in entries.items():
        valor = parsear_monto(entry.get())
        if valor is None:
            errores.append(f"{expse}: \"{entry.get()}\" no es un número válido")
        else:
            montos[expse] = valor
    return montos, errores


def guardar_expse_montos():
    for expse, entry in entries.items():
        expse_montos[expse] = entry.get()
    try:
        with open(archivo_json, "w", encoding="utf-8") as file:
            json.dump(expse_montos, file, indent=4, ensure_ascii=False)
    except OSError as e:
        messagebox.showwarning("Aviso", f"No se pudieron guardar los montos en {archivo_json}\n{e}")


def obtener_periodo_seleccionado():
    """Devuelve (numero_mes, año) del período elegido, o lanza ValueError."""
    nombre_mes = cmb_mes.get()
    if nombre_mes not in MESES:
        raise ValueError("Seleccioná un mes válido.")
    try:
        año = int(spn_año.get())
    except ValueError:
        raise ValueError("Ingresá un año válido.")
    if not 2000 <= año <= 2100:
        raise ValueError("Ingresá un año válido.")
    return MESES.index(nombre_mes) + 1, año


def obtener_nombre_archivo_unico(base_path, mes, año):
    nombre_mes = MESES[mes - 1]
    ruta_archivo = base_path / f"Liquidación {nombre_mes} {año}.xlsx"
    contador = 1

    while ruta_archivo.exists():
        ruta_archivo = base_path / f"Liquidación {nombre_mes} {año} ({contador}).xlsx"
        contador += 1

    return ruta_archivo


def procesar_excel():
    global ruta_archivo_modificado

    if not ruta_archivo_seleccionado:
        messagebox.showerror("Error", "Primero selecciona un archivo de Excel")
        return

    montos, errores = leer_y_validar_montos()
    if errores:
        messagebox.showerror("Montos inválidos", "Corregí los siguientes montos:\n\n" + "\n".join(errores))
        return

    try:
        mes, año = obtener_periodo_seleccionado()
    except ValueError as e:
        messagebox.showerror("Período inválido", str(e))
        return

    guardar_expse_montos()

    try:
        df = pd.read_excel(ruta_archivo_seleccionado, header=0)
        columnas_deseadas = ["Fecha", "Profesional", "HC", "Trabajador", "Prestación"]

        if all(col in df.columns for col in columnas_deseadas):
            df = df[columnas_deseadas]
        else:
            df = df.iloc[:, [1, 2, 3, 4, 7]]
            df.columns = columnas_deseadas

        # Descartar filas iniciales ANTES de calcular el total, para que el total
        # coincida exactamente con las filas que quedan en la planilla.
        df = df.iloc[FILAS_A_DESCARTAR:].reset_index(drop=True)

        montos_normalizados = {k.strip().upper(): v for k, v in montos.items()}
        prestaciones = df["Prestación"].fillna("").astype(str).str.strip()
        df["Monto"] = prestaciones.str.upper().map(montos_normalizados)

        # Prestaciones que no corresponden a ningún EXPSE configurado
        mascara_sin_monto = df["Monto"].isna()
        sin_monto = sorted(set(prestaciones[mascara_sin_monto]))
        cantidad_sin_monto = int(mascara_sin_monto.sum())
        df["Monto"] = df["Monto"].fillna(0)

        total_monto = df["Monto"].sum()
        df.loc[len(df)] = ["TOTAL", "", "", "", "", total_monto]

        ruta_archivo_modificado = obtener_nombre_archivo_unico(ruta_guardado, mes, año)
        df.to_excel(ruta_archivo_modificado, index=False, header=True)

        mensaje = f"Archivo guardado en: {ruta_archivo_modificado}"
        if sin_monto:
            lista = "\n".join(f"  • {p if p else '(vacío)'}" for p in sin_monto)
            mensaje += (
                f"\n\nAtención: {cantidad_sin_monto} fila(s) quedaron con monto 0 porque su prestación "
                f"no coincide con ningún EXPSE configurado:\n{lista}"
            )
            messagebox.showwarning("Procesado con advertencias", mensaje)
        else:
            messagebox.showinfo("Éxito", mensaje)

        btn_abrir_ubicacion.config(state=tk.NORMAL)

    except ImportError as e:
        messagebox.showerror("Error", f"Falta una librería para leer este tipo de archivo\n{e}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo procesar el archivo\n{e}")


def abrir_ubicacion():
    if ruta_archivo_modificado and os.path.exists(ruta_archivo_modificado):
        carpeta = os.path.dirname(ruta_archivo_modificado)
        subprocess.run(["explorer", carpeta] if os.name == "nt" else ["xdg-open", carpeta])
    else:
        messagebox.showwarning("Aviso", "No hay archivo generado aún.")


def restablecer_mes_vencido():
    mes, año = periodo_mes_vencido()
    cmb_mes.set(MESES[mes - 1])
    var_año.set(str(año))


# --- Interfaz Gráfica --- #
root = tk.Tk()
root.title("Cargar y Modificar Excel")
root.geometry("500x460")

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

try:
    icon_path = os.path.join(base_path, "icon.ico")
    if os.path.exists(icon_path):
        root.iconbitmap(icon_path)
except Exception as e:
    print(f"No se pudo cargar el ícono: {e}")


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

# Período a liquidar (por defecto: mes vencido)
frame_periodo = tk.LabelFrame(frame_interior, text="Período a liquidar")
frame_periodo.pack(pady=10)

cmb_mes = ttk.Combobox(frame_periodo, values=MESES, state="readonly", width=12)
cmb_mes.pack(side=tk.LEFT, padx=5, pady=5)

var_año = tk.StringVar()
spn_año = tk.Spinbox(frame_periodo, from_=2000, to=2100, width=6, textvariable=var_año)
spn_año.pack(side=tk.LEFT, padx=5, pady=5)

btn_mes_vencido = tk.Button(frame_periodo, text="Mes vencido", command=restablecer_mes_vencido)
btn_mes_vencido.pack(side=tk.LEFT, padx=5, pady=5)

restablecer_mes_vencido()

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
