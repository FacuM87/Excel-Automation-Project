from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from . import base_datos
from .procesamiento import (
    COLUMNA_ORDEN, MESES, Liquidacion, aplicar_montos, convertir_fechas, leer_planilla_origen, ordenes_repetidas,
)


@dataclass
class PeriodoImportable:
    """Un mes a importar, armado con las órdenes de uno o varios archivos."""
    mes: int
    año: int
    archivos: list
    liquidacion: Liquidacion
    duplicadas_descartadas: int = 0
    registrado: base_datos.LiquidacionRegistrada = None
    en_otros_periodos: dict = field(default_factory=dict)

    @property
    def nombre(self):
        return f"{MESES[self.mes - 1]} {self.año}"

    @property
    def repetidas_en_archivo(self):
        return ordenes_repetidas(self.liquidacion.tabla)


@dataclass
class AnalisisImportacion:
    periodos: list
    errores: list
    sin_fecha: int = 0
    no_cumplidas: int = 0


@dataclass
class ResultadoImportacion:
    registrados: list = field(default_factory=list)
    reemplazados: list = field(default_factory=list)
    salteados: list = field(default_factory=list)


def _unir_sin_duplicar(partes):
    """Une las tablas de un mes. Las órdenes cuyo número ya apareció en un archivo anterior se descartan."""
    vistas = set()
    tablas = []
    descartadas = 0
    for df in partes:
        if COLUMNA_ORDEN in df.columns:
            numeros = df[COLUMNA_ORDEN]
            repetida = numeros.notna() & numeros.isin(vistas)
            descartadas += int(repetida.sum())
            vistas.update(numeros.dropna())
            df = df[~repetida]
        tablas.append(df)
    unida = pd.concat(tablas, ignore_index=True)
    if COLUMNA_ORDEN in unida.columns:
        unida[COLUMNA_ORDEN] = unida[COLUMNA_ORDEN].astype(object).where(unida[COLUMNA_ORDEN].notna(), None)
    orden = convertir_fechas(unida["Fecha"]).argsort(kind="stable")
    return unida.iloc[orden].reset_index(drop=True), descartadas


def analizar_archivos(rutas, montos):
    """Lee los archivos y arma un período por mes. Los archivos que no se pueden leer quedan en errores."""
    partes = {}
    no_cumplidas = 0
    errores = []
    sin_fecha = 0
    for ruta in sorted(rutas, key=lambda r: Path(r).name):
        nombre = Path(ruta).name
        try:
            df, excluidas = leer_planilla_origen(ruta)
            no_cumplidas += excluidas
        except Exception as e:
            errores.append((nombre, str(e)))
            continue
        fechas = convertir_fechas(df["Fecha"])
        sin_fecha += int(fechas.isna().sum())
        claves = list(zip(fechas.dt.month, fechas.dt.year))
        for mes, año in sorted({c for c in claves if not pd.isna(c[0])}, key=lambda c: (c[1], c[0])):
            periodo = (int(mes), int(año))
            filas = (fechas.dt.month == mes) & (fechas.dt.year == año)
            partes.setdefault(periodo, []).append((nombre, df[filas]))

    periodos = []
    for (mes, año) in sorted(partes, key=lambda p: (p[1], p[0])):
        archivos = [nombre for nombre, _ in partes[(mes, año)]]
        tabla, descartadas = _unir_sin_duplicar([df for _, df in partes[(mes, año)]])
        liquidacion = aplicar_montos(tabla, montos)
        periodos.append(PeriodoImportable(mes, año, archivos, liquidacion, descartadas))
    return AnalisisImportacion(periodos, errores, sin_fecha, no_cumplidas)


def verificar_en_base(con, periodos):
    """Marca los períodos ya registrados y las órdenes que están en otros períodos de la base."""
    for periodo in periodos:
        periodo.registrado = base_datos.buscar_liquidacion(con, periodo.mes, periodo.año)
        periodo.en_otros_periodos = {}
        if periodo.liquidacion.tiene_numero_orden:
            periodo.en_otros_periodos = base_datos.buscar_ordenes_registradas(
                con, periodo.liquidacion.tabla[COLUMNA_ORDEN], (periodo.mes, periodo.año))


def importar(con, periodos, reemplazar=False):
    """Registra los períodos en la base, sin generar Excel. Los ya registrados se reemplazan o se saltean."""
    resultado = ResultadoImportacion()
    for periodo in periodos:
        if base_datos.buscar_liquidacion(con, periodo.mes, periodo.año) and not reemplazar:
            resultado.salteados.append(periodo.nombre)
            continue
        reemplazado = base_datos.guardar(
            con, periodo.mes, periodo.año, periodo.liquidacion, ", ".join(periodo.archivos))
        (resultado.reemplazados if reemplazado else resultado.registrados).append(periodo.nombre)
    return resultado
