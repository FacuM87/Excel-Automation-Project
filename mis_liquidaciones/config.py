import json
import sys
from pathlib import Path

from .procesamiento import normalizar_prestacion

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


def obtener_ruta_guardado():
    escritorio = Path.home() / "Desktop"
    documentos = Path.home() / "Documents"
    return escritorio if escritorio.exists() else documentos if documentos.exists() else Path.home()


ARCHIVO_MONTOS = obtener_directorio_app() / "prestaciones.json"
ARCHIVOS_MONTOS_ANTERIORES = (obtener_directorio_app() / "expse_montos.json", Path("expse_montos.json"))
ARCHIVO_BASE = obtener_directorio_app() / "liquidaciones.db"


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


ARCHIVO_AJUSTES = obtener_directorio_app() / "ajustes.json"
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
