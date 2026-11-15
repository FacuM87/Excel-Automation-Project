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
COLUMNA_ORDEN = "Orden"
COLUMNA_SUCESO = "Suceso"
FILAS_BUSQUEDA_ENCABEZADO = 10
PSICOMEY = "PSICOMEY"


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
    psicomey_duplicados: int = 0

    @property
    def tiene_numero_orden(self):
        return COLUMNA_ORDEN in self.tabla.columns


@dataclass
class PeriodoDetectado:
    mes: int
    año: int
    ordenes: int

    @property
    def nombre(self):
        return f"{MESES[self.mes - 1]} {self.año}"


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


def formatear_monto(valor):
    """1234567.5 -> '1.234.567,50'; sin decimales si es entero."""
    texto = f"{valor:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return texto[:-3] if texto.endswith(",00") else texto


def normalizar_orden(valor):
    """Número de orden como entero si es posible, texto si no, None si falta."""
    if pd.isna(valor) or str(valor).strip() == "":
        return None
    try:
        numero = float(valor)
        if numero.is_integer():
            return int(numero)
    except (TypeError, ValueError):
        pass
    return str(valor).strip()


def ordenes_repetidas(df):
    """Números de orden que aparecen más de una vez en la tabla."""
    if COLUMNA_ORDEN not in df.columns:
        return []
    numeros = df[COLUMNA_ORDEN].dropna()
    return sorted(set(numeros[numeros.duplicated()]), key=str)


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

    columnas = list(COLUMNAS_SALIDA)
    for columna in (COLUMNA_ORDEN, COLUMNA_SUCESO):
        if columna in df.columns:
            valores = pd.Series([normalizar_orden(v) for v in df[columna]], index=df.index, dtype=object)
            df = df.assign(**{columna: valores})
    if COLUMNA_ORDEN in df.columns:
        columnas.insert(0, COLUMNA_ORDEN)
    if COLUMNA_SUCESO in df.columns:
        columnas.append(COLUMNA_SUCESO)
    return df[columnas].reset_index(drop=True), no_cumplidas


def convertir_fechas(serie):
    """Fechas de Excel o texto dd/mm/aaaa. Las inválidas quedan como NaT."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    return pd.to_datetime(serie, dayfirst=True, errors="coerce", format="mixed")


def detectar_periodos(df):
    """Meses presentes en las órdenes, en orden cronológico. Lista vacía si no hay fechas válidas."""
    fechas = convertir_fechas(df["Fecha"]).dropna()
    conteo = fechas.groupby([fechas.dt.year.rename("año"), fechas.dt.month.rename("mes")]).size()
    return [PeriodoDetectado(int(mes), int(año), int(n)) for (año, mes), n in conteo.items()]


def quitar_psicomey_duplicados(df):
    """Quita las filas de PSICOMEY que duplican otra orden: otra prestación con la misma fecha y hora, HC y
    Suceso (si la columna existe) y distinto número de orden (R3). Devuelve (tabla, cantidad quitada)."""
    df = df.reset_index(drop=True)
    es_psicomey = df["Prestación"].fillna("").astype(str).map(normalizar_prestacion) == PSICOMEY
    if not es_psicomey.any():
        return df, 0
    fechas = convertir_fechas(df["Fecha"])
    hc = [normalizar_orden(v) for v in df["HC"]]
    suceso = list(df[COLUMNA_SUCESO]) if COLUMNA_SUCESO in df.columns else [None] * len(df)
    orden = list(df[COLUMNA_ORDEN]) if COLUMNA_ORDEN in df.columns else [None] * len(df)

    def clave(i):
        if pd.isna(fechas[i]) or hc[i] is None:
            return None
        return fechas[i], hc[i], suceso[i]

    otras = {}
    for i in df.index[~es_psicomey]:
        if clave(i) is not None:
            otras.setdefault(clave(i), []).append(orden[i])
    duplicada = [
        bool(es_psicomey[i]) and clave(i) in otras and any(o is None or o != orden[i] for o in otras[clave(i)])
        for i in df.index
    ]
    mascara = pd.Series(duplicada, index=df.index)
    return df[~mascara].reset_index(drop=True), int(mascara.sum())


def liquidar(ruta, montos, periodo=None):
    """Lee la planilla y aplica los montos ({prestación: número}). Las prestaciones sin monto quedan en 0.
    Si se indica periodo (mes, año), solo se liquidan las órdenes de ese mes."""
    df, no_cumplidas = leer_planilla_origen(ruta)

    if periodo is not None:
        mes, año = periodo
        fechas = convertir_fechas(df["Fecha"])
        df = df[(fechas.dt.month == mes) & (fechas.dt.year == año)].reset_index(drop=True)

    return aplicar_montos(df, montos, no_cumplidas)


def aplicar_montos(df, montos, no_cumplidas=0):
    """Liquida una tabla ya leída: quita los PSICOMEY duplicados y agrega la columna Monto."""
    df, psicomey_duplicados = quitar_psicomey_duplicados(df)
    normalizadas = df["Prestación"].fillna("").astype(str).map(normalizar_prestacion)

    montos_normalizados = {normalizar_prestacion(k): v for k, v in montos.items()}
    df = df.copy()
    df["Monto"] = normalizadas.map(montos_normalizados)

    mascara_sin_monto = df["Monto"].isna()
    df["Monto"] = df["Monto"].fillna(0)

    return Liquidacion(
        tabla=df,
        total=df["Monto"].sum(),
        no_cumplidas=no_cumplidas,
        sin_monto=sorted(set(normalizadas[mascara_sin_monto])),
        cantidad_sin_monto=int(mascara_sin_monto.sum()),
        psicomey_duplicados=psicomey_duplicados,
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
    df = liquidacion.tabla.drop(columns=[COLUMNA_ORDEN, COLUMNA_SUCESO], errors="ignore")
    df.loc[len(df)] = ["TOTAL", "", "", "", "", liquidacion.total]
    ruta = obtener_nombre_archivo_unico(carpeta, mes, año)
    df.to_excel(ruta, index=False, header=True)
    return ruta
