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

COLUMNAS_SALIDA = ["Fecha", "Profesional", "HC", "Trabajador", "Prestación"]
FILAS_BUSQUEDA_ENCABEZADO = 10
PRESTACIONES_POR_DEFECTO = [
    "EXPSE 1", "EXPSE 2", "EXPSE 4", "EXPSE 6", "EXPSE 7",
    "ATEPSIQ", "PSICOMEY", "CONS", "CUES DES",
]


class FormatoInvalido(Exception):
    pass


def obtener_directorio_app():
    """Carpeta del .exe o del script, independiente del directorio de trabajo."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def obtener_ruta_guardado():
    escritorio = Path.home() / "Desktop"
    documentos = Path.home() / "Documents"
    return escritorio if escritorio.exists() else documentos if documentos.exists() else Path.home()


def normalizar_prestacion(nombre):
    return " ".join(str(nombre).split()).upper()


def leer_json(ruta):
    try:
        with open(ruta, "r", encoding="utf-8") as file:
            datos = json.load(file)
        return {normalizar_prestacion(k): str(v) for k, v in datos.items()}
    except (OSError, json.JSONDecodeError, AttributeError):
        return None


def cargar_montos():
    """Lee prestaciones.json; si no existe, migra los montos de expse_montos.json."""
    if archivo_json.exists():
        montos = leer_json(archivo_json)
        if montos is not None:
            return montos

    montos = {}
    for ruta in archivos_json_anteriores:
        if ruta.exists():
            montos = leer_json(ruta) or {}
            break
    for prestacion in PRESTACIONES_POR_DEFECTO:
        montos.setdefault(prestacion, "")
    return montos


ruta_guardado = obtener_ruta_guardado()
archivo_json = obtener_directorio_app() / "prestaciones.json"
archivos_json_anteriores = (obtener_directorio_app() / "expse_montos.json", Path("expse_montos.json"))

montos_prestaciones = cargar_montos()

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
    """Lee los montos de la interfaz. Devuelve (montos_numericos, errores). Los vacíos se omiten."""
    montos = {}
    errores = []
    for prestacion, entry in entries.items():
        if not entry.get().strip():
            continue
        valor = parsear_monto(entry.get())
        if valor is None:
            errores.append(f"{prestacion}: \"{entry.get()}\" no es un número válido")
        else:
            montos[prestacion] = valor
    return montos, errores


def sincronizar_montos():
    for prestacion, entry in entries.items():
        montos_prestaciones[prestacion] = entry.get()


def guardar_montos():
    sincronizar_montos()
    try:
        with open(archivo_json, "w", encoding="utf-8") as file:
            json.dump(montos_prestaciones, file, indent=4, ensure_ascii=False)
    except OSError as e:
        messagebox.showwarning("Aviso", f"No se pudieron guardar los montos en {archivo_json}\n{e}")


def dibujar_prestaciones():
    for widget in frame_lista.winfo_children():
        widget.destroy()
    entries.clear()
    for i, prestacion in enumerate(montos_prestaciones):
        fila, col = divmod(i, 2)
        col *= 3
        tk.Label(frame_lista, text=f"{prestacion}:").grid(row=fila, column=col, sticky="e", padx=(10, 2), pady=2)
        entry = tk.Entry(frame_lista, width=10)
        entry.insert(0, montos_prestaciones[prestacion])
        entry.grid(row=fila, column=col + 1, pady=2)
        tk.Button(frame_lista, text="✕", width=2, command=lambda p=prestacion: quitar_prestacion(p)).grid(
            row=fila, column=col + 2, padx=(2, 10), pady=2)
        entries[prestacion] = entry


def agregar_prestacion():
    nombre = normalizar_prestacion(ent_nueva_nombre.get())
    monto = ent_nueva_monto.get().strip()
    if not nombre:
        messagebox.showerror("Prestación inválida", "Ingresá el nombre de la prestación.")
        return
    if nombre in entries:
        messagebox.showerror("Prestación repetida", f"{nombre} ya está en la lista.")
        return
    if monto and parsear_monto(monto) is None:
        messagebox.showerror("Monto inválido", f"\"{monto}\" no es un número válido.")
        return
    sincronizar_montos()
    montos_prestaciones[nombre] = monto
    dibujar_prestaciones()
    guardar_montos()
    ent_nueva_nombre.delete(0, tk.END)
    ent_nueva_monto.delete(0, tk.END)


def quitar_prestacion(prestacion):
    if not messagebox.askyesno("Quitar prestación", f"¿Quitar {prestacion} de la lista?"):
        return
    sincronizar_montos()
    del montos_prestaciones[prestacion]
    dibujar_prestaciones()
    guardar_montos()


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


def leer_planilla_origen(ruta):
    vista_previa = pd.read_excel(ruta, header=None, nrows=FILAS_BUSQUEDA_ENCABEZADO)
    for fila_encabezado, fila in vista_previa.iterrows():
        valores = {str(v).strip() for v in fila if pd.notna(v)}
        if {"Fecha", "Prestación"} <= valores:
            break
    else:
        raise FormatoInvalido("No se encontró la fila de encabezados (con las columnas Fecha y Prestación).")

    df = pd.read_excel(ruta, header=fila_encabezado)
    df.columns = [str(c).strip() for c in df.columns]

    faltantes = [c for c in COLUMNAS_SALIDA if c not in df.columns]
    if faltantes:
        raise FormatoInvalido("Faltan las columnas: " + ", ".join(faltantes))

    df = df.dropna(subset=["Fecha", "Prestación"], how="all")

    no_cumplidas = 0
    if "Cumplida" in df.columns:
        cumplida = df["Cumplida"].fillna("").astype(str).str.strip().str.upper().str.replace("Í", "I")
        es_cumplida = cumplida == "SI"
        no_cumplidas = int((~es_cumplida).sum())
        df = df[es_cumplida]

    return df[COLUMNAS_SALIDA].reset_index(drop=True), no_cumplidas


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

    guardar_montos()

    try:
        df, no_cumplidas = leer_planilla_origen(ruta_archivo_seleccionado)

        prestaciones = df["Prestación"].fillna("").astype(str).str.strip()
        df["Monto"] = prestaciones.map(normalizar_prestacion).map(montos)

        mascara_sin_monto = df["Monto"].isna()
        sin_monto = sorted(set(prestaciones[mascara_sin_monto]))
        cantidad_sin_monto = int(mascara_sin_monto.sum())
        df["Monto"] = df["Monto"].fillna(0)

        total_monto = df["Monto"].sum()
        df.loc[len(df)] = ["TOTAL", "", "", "", "", total_monto]

        ruta_archivo_modificado = obtener_nombre_archivo_unico(ruta_guardado, mes, año)
        df.to_excel(ruta_archivo_modificado, index=False, header=True)

        advertencias = []
        if no_cumplidas:
            advertencias.append(f"Se excluyeron {no_cumplidas} orden(es) no cumplida(s).")
        if sin_monto:
            lista = "\n".join(f"  • {p if p else '(vacío)'}" for p in sin_monto)
            advertencias.append(
                f"{cantidad_sin_monto} fila(s) quedaron con monto 0 porque su prestación "
                f"no está configurada o no tiene monto cargado:\n{lista}"
            )

        mensaje = f"Archivo guardado en: {ruta_archivo_modificado}"
        if advertencias:
            mensaje += "\n\nAtención:\n" + "\n\n".join(advertencias)
            messagebox.showwarning("Procesado con advertencias", mensaje)
        else:
            messagebox.showinfo("Éxito", mensaje)

        btn_abrir_ubicacion.config(state=tk.NORMAL)

    except FormatoInvalido as e:
        messagebox.showerror("Formato no reconocido", f"El archivo no tiene el formato esperado.\n{e}")
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
root.minsize(500, 0)

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

frame_contenedor = tk.Frame(root)
frame_contenedor.pack(expand=True, fill="both")

frame_interior = tk.Frame(frame_contenedor)
frame_interior.pack(expand=True, padx=20, pady=10)

btn_cargar = tk.Button(frame_interior, text="Seleccionar Planilla", width=25, command=seleccionar_archivo)
btn_cargar.pack(pady=10)

lbl_archivo = tk.Label(frame_interior, text="Ningún archivo seleccionado", fg="blue")
lbl_archivo.pack()

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

frame_montos = tk.LabelFrame(frame_interior, text="Montos por prestación")
frame_montos.pack(pady=10, fill="x")

entries = {}
frame_lista = tk.Frame(frame_montos)
frame_lista.pack(pady=5)
dibujar_prestaciones()

frame_nueva = tk.Frame(frame_montos)
frame_nueva.pack(pady=(5, 8))
tk.Label(frame_nueva, text="Nueva:").pack(side=tk.LEFT, padx=5)
ent_nueva_nombre = tk.Entry(frame_nueva, width=14)
ent_nueva_nombre.pack(side=tk.LEFT, padx=2)
tk.Label(frame_nueva, text="Monto:").pack(side=tk.LEFT, padx=5)
ent_nueva_monto = tk.Entry(frame_nueva, width=10)
ent_nueva_monto.pack(side=tk.LEFT, padx=2)
tk.Button(frame_nueva, text="Agregar", command=agregar_prestacion).pack(side=tk.LEFT, padx=5)

btn_procesar = tk.Button(frame_interior, text="Procesar Planilla", width=25, command=procesar_excel)
btn_procesar.pack(pady=10)

btn_abrir_ubicacion = tk.Button(frame_interior, text="Abrir Ubicación de Planilla Modificada", width=30, command=abrir_ubicacion, state=tk.DISABLED)
btn_abrir_ubicacion.pack(pady=10)

root.update_idletasks()
ancho = root.winfo_reqwidth()
alto = root.winfo_reqheight()
x = (root.winfo_screenwidth() // 2) - (ancho // 2)
y = (root.winfo_screenheight() // 2) - (alto // 2)
root.geometry(f"+{x}+{max(y, 0)}")

root.mainloop()
