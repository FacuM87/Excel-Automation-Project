import json
import os
import shutil
import sys
from pathlib import Path

from .procesamiento import normalizar_prestacion

VERSION = "2.0"
ARCHIVOS_DE_DATOS = ("prestaciones.json", "expse_montos.json", "ajustes.json", "liquidaciones.db")

PRESTACIONES_POR_DEFECTO = [
    "EXPSE 1", "EXPSE 2", "EXPSE 4", "EXPSE 6", "EXPSE 7",
    "ATEPSIQ", "PSICOMEY", "CONS", "CUES DES",
]


def obtener_directorio_app():
    """Carpeta del .exe o del proyecto, independiente del directorio de trabajo."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def obtener_directorio_recursos():
    """Carpeta de los archivos empaquetados (ícono) dentro del .exe, o la del proyecto."""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return obtener_directorio_app()


def migrar_datos(origen, destino):
    """Copia los datos de origen a destino si destino todavía no tiene ninguno."""
    if any((destino / nombre).exists() for nombre in ARCHIVOS_DE_DATOS):
        return
    for nombre in ARCHIVOS_DE_DATOS:
        if (origen / nombre).is_file():
            shutil.copy2(origen / nombre, destino / nombre)


def obtener_directorio_datos():
    """En el .exe, %APPDATA%\\MisLiquidaciones (con los datos que hubiera al lado del .exe);
    desde el código, la carpeta del proyecto."""
    if not getattr(sys, 'frozen', False):
        return obtener_directorio_app()
    destino = Path(os.environ.get("APPDATA") or Path.home()) / "MisLiquidaciones"
    try:
        destino.mkdir(parents=True, exist_ok=True)
        migrar_datos(obtener_directorio_app(), destino)
    except OSError:
        return obtener_directorio_app()
    return destino


def obtener_ruta_guardado():
    escritorio = Path.home() / "Desktop"
    documentos = Path.home() / "Documents"
    return escritorio if escritorio.exists() else documentos if documentos.exists() else Path.home()


DIRECTORIO_DATOS = obtener_directorio_datos()
ARCHIVO_MONTOS = DIRECTORIO_DATOS / "prestaciones.json"
ARCHIVOS_MONTOS_ANTERIORES = (DIRECTORIO_DATOS / "expse_montos.json", Path("expse_montos.json"))
ARCHIVO_BASE = DIRECTORIO_DATOS / "liquidaciones.db"


def leer_json(ruta):
    try:
        with open(ruta, "r", encoding="utf-8") as file:
            datos = json.load(file)
        return {normalizar_prestacion(k): str(v) for k, v in datos.items()}
    except (OSError, json.JSONDecodeError, AttributeError):
        return None


def cargar_montos(archivo=ARCHIVO_MONTOS, anteriores=ARCHIVOS_MONTOS_ANTERIORES):
    """Lee prestaciones.json; si no existe, migra los montos de expse_montos.json."""
    if archivo.exists():
        montos = leer_json(archivo)
        if montos is not None:
            return montos

    montos = {}
    for ruta in anteriores:
        if ruta.exists():
            montos = leer_json(ruta) or {}
            break
    for prestacion in PRESTACIONES_POR_DEFECTO:
        montos.setdefault(prestacion, "")
    return montos


def guardar_montos(montos, archivo=ARCHIVO_MONTOS):
    with open(archivo, "w", encoding="utf-8") as file:
        json.dump(montos, file, indent=4, ensure_ascii=False)


ARCHIVO_AJUSTES = DIRECTORIO_DATOS / "ajustes.json"
TEMAS = {"claro": "flatly", "oscuro": "darkly"}
AJUSTES_POR_DEFECTO = {"tema": "claro", "carpeta_guardado": ""}


def cargar_ajustes(archivo=ARCHIVO_AJUSTES):
    """Lee ajustes.json; completa con los valores por defecto lo que falte o sea inválido."""
    ajustes = dict(AJUSTES_POR_DEFECTO)
    try:
        with open(archivo, "r", encoding="utf-8") as file:
            datos = json.load(file)
    except (OSError, json.JSONDecodeError):
        return ajustes
    if isinstance(datos, dict):
        if datos.get("tema") in TEMAS:
            ajustes["tema"] = datos["tema"]
        if isinstance(datos.get("carpeta_guardado"), str):
            ajustes["carpeta_guardado"] = datos["carpeta_guardado"]
    return ajustes


def guardar_ajustes(ajustes, archivo=ARCHIVO_AJUSTES):
    with open(archivo, "w", encoding="utf-8") as file:
        json.dump(ajustes, file, indent=4, ensure_ascii=False)


def carpeta_guardado(ajustes):
    """La carpeta elegida si todavía existe; si no, la de siempre (Escritorio)."""
    elegida = ajustes.get("carpeta_guardado")
    if elegida and Path(elegida).is_dir():
        return Path(elegida)
    return obtener_ruta_guardado()
