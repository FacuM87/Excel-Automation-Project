import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

COLUMNAS_SALIDA = ["Fecha", "Profesional", "HC", "Trabajador", "Prestación"]
FILAS_BUSQUEDA_ENCABEZADO = 10
PRESTACION_OPCIONAL = "PSICOMEY"


class FormatoInvalido(Exception):
    pass


@dataclass
class Liquidacion:
    """Resultado de liquidar una planilla: tabla con montos, total y avisos."""
    tabla: pd.DataFrame
    total: float
    no_cumplidas: int = 0
    sin_monto: list = field(default_factory=list)
    cantidad_sin_monto: int = 0
    excluidas_opcional: int = 0


def normalizar_prestacion(nombre):
    return " ".join(str(nombre).split()).upper()


def parsear_monto(texto):
    """Convierte el texto ingresado a número. Vacío = 0. Acepta coma decimal y punto de miles
    (1.000 = mil). Devuelve None si es inválido."""
    texto = texto.strip().replace(" ", "")
    if texto == "":
        return 0.0
    if "," in texto or re.fullmatch(r"\d{1,3}(\.\d{3})+", texto):
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def periodo_mes_vencido(hoy=None):
    hoy = hoy or datetime.now()
    if hoy.month == 1:
        return 12, hoy.year - 1
    return hoy.month - 1, hoy.year


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


def liquidar(ruta, montos, incluir_opcional=False):
    """Lee la planilla y aplica los montos ({prestación: número}). Las prestaciones sin monto quedan en 0.
    Si incluir_opcional es False, se quitan las filas de PRESTACION_OPCIONAL."""
    df, no_cumplidas = leer_planilla_origen(ruta)

    normalizadas = df["Prestación"].fillna("").astype(str).map(normalizar_prestacion)
    excluidas_opcional = 0
    if not incluir_opcional:
        es_opcional = normalizadas == PRESTACION_OPCIONAL
        excluidas_opcional = int(es_opcional.sum())
        df = df[~es_opcional].reset_index(drop=True)
        normalizadas = normalizadas[~es_opcional].reset_index(drop=True)

    montos_normalizados = {normalizar_prestacion(k): v for k, v in montos.items()}
    df["Monto"] = normalizadas.map(montos_normalizados)

    mascara_sin_monto = df["Monto"].isna()
    df["Monto"] = df["Monto"].fillna(0)

    return Liquidacion(
        tabla=df,
        total=df["Monto"].sum(),
        no_cumplidas=no_cumplidas,
        sin_monto=sorted(set(normalizadas[mascara_sin_monto])),
        cantidad_sin_monto=int(mascara_sin_monto.sum()),
        excluidas_opcional=excluidas_opcional,
    )


def obtener_nombre_archivo_unico(carpeta, mes, año):
    nombre_mes = MESES[mes - 1]
    ruta_archivo = Path(carpeta) / f"Liquidación {nombre_mes} {año}.xlsx"
    contador = 1

    while ruta_archivo.exists():
        ruta_archivo = Path(carpeta) / f"Liquidación {nombre_mes} {año} ({contador}).xlsx"
        contador += 1

    return ruta_archivo


def guardar_liquidacion(liquidacion, carpeta, mes, año):
    """Guarda la liquidación con la fila TOTAL y devuelve la ruta del archivo generado."""
    df = liquidacion.tabla.copy()
    df.loc[len(df)] = ["TOTAL", "", "", "", "", liquidacion.total]
    ruta = obtener_nombre_archivo_unico(carpeta, mes, año)
    df.to_excel(ruta, index=False, header=True)
    return ruta
