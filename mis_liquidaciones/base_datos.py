import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from .procesamiento import (
    COLUMNA_ORDEN, COLUMNA_SUCESO, MESES, PSICOMEY, convertir_fechas, normalizar_orden, normalizar_prestacion,
)

VERSION_ESQUEMA = 2
LOTE_CONSULTA = 500

ESQUEMA = """
CREATE TABLE IF NOT EXISTS liquidaciones (
    id INTEGER PRIMARY KEY,
    mes INTEGER NOT NULL,
    año INTEGER NOT NULL,
    generada TEXT NOT NULL,
    archivo_origen TEXT,
    archivo_generado TEXT,
    incluye_opcional INTEGER NOT NULL,
    cantidad_ordenes INTEGER NOT NULL,
    total REAL NOT NULL,
    enviada TEXT,
    UNIQUE (mes, año)
);
CREATE TABLE IF NOT EXISTS ordenes (
    id INTEGER PRIMARY KEY,
    liquidacion_id INTEGER NOT NULL REFERENCES liquidaciones(id) ON DELETE CASCADE,
    orden,
    fecha TEXT,
    profesional TEXT,
    hc TEXT,
    trabajador TEXT,
    prestacion TEXT NOT NULL,
    monto REAL NOT NULL,
    suceso
);
CREATE INDEX IF NOT EXISTS idx_ordenes_orden ON ordenes(orden);
CREATE INDEX IF NOT EXISTS idx_ordenes_liquidacion ON ordenes(liquidacion_id);
"""


@dataclass
class LiquidacionRegistrada:
    id: int
    mes: int
    año: int
    generada: datetime
    archivo_origen: str
    archivo_generado: str
    incluye_opcional: bool
    cantidad_ordenes: int
    total: float
    enviada: datetime = None

    @property
    def nombre(self):
        return f"{MESES[self.mes - 1]} {self.año}"


def abrir(ruta):
    """Abre la base (la crea si no existe) con las claves foráneas activas."""
    con = sqlite3.connect(Path(ruta))
    con.execute("PRAGMA foreign_keys = ON")
    con.executescript(ESQUEMA)
    columnas = {fila[1] for fila in con.execute("PRAGMA table_info(ordenes)")}
    if "suceso" not in columnas:
        con.execute("ALTER TABLE ordenes ADD COLUMN suceso")
    columnas = {fila[1] for fila in con.execute("PRAGMA table_info(liquidaciones)")}
    if "enviada" not in columnas:
        con.execute("ALTER TABLE liquidaciones ADD COLUMN enviada TEXT")
    if con.execute("PRAGMA user_version").fetchone()[0] < 2:
        limpiar_psicomey_duplicados(con)
        con.execute(f"PRAGMA user_version = {VERSION_ESQUEMA}")
    return con


def limpiar_psicomey_duplicados(con):
    """Borra los PSICOMEY guardados que duplican otra orden de la misma liquidación (misma fecha y hora, HC y
    suceso si está guardado) y recalcula cantidad y total. Devuelve cuántas órdenes borró."""
    with con:
        borradas = con.execute(
            "DELETE FROM ordenes WHERE prestacion = ? AND fecha IS NOT NULL AND hc IS NOT NULL AND EXISTS ("
            "SELECT 1 FROM ordenes x WHERE x.liquidacion_id = ordenes.liquidacion_id AND x.prestacion <> ? "
            "AND x.fecha = ordenes.fecha AND x.hc = ordenes.hc "
            "AND (x.suceso IS NULL OR ordenes.suceso IS NULL OR x.suceso = ordenes.suceso) "
            "AND (x.orden IS NULL OR ordenes.orden IS NULL OR x.orden <> ordenes.orden))",
            (PSICOMEY, PSICOMEY)).rowcount
        if borradas:
            con.execute(
                "UPDATE liquidaciones SET "
                "cantidad_ordenes = (SELECT COUNT(*) FROM ordenes o WHERE o.liquidacion_id = liquidaciones.id), "
                "total = (SELECT COALESCE(SUM(monto), 0) FROM ordenes o WHERE o.liquidacion_id = liquidaciones.id)")
    return borradas


def _a_registro(fila):
    id_, mes, año, generada, origen, generado, opcional, cantidad, total, enviada = fila
    return LiquidacionRegistrada(
        id_, mes, año, datetime.fromisoformat(generada), origen, generado, bool(opcional), cantidad, total,
        datetime.fromisoformat(enviada) if enviada else None)


_COLUMNAS_LIQUIDACION = (
    "id, mes, año, generada, archivo_origen, archivo_generado, incluye_opcional, cantidad_ordenes, total, enviada")


def buscar_liquidacion(con, mes, año):
    fila = con.execute(
        f"SELECT {_COLUMNAS_LIQUIDACION} FROM liquidaciones WHERE mes = ? AND año = ?", (mes, año)).fetchone()
    return _a_registro(fila) if fila else None


def listar_liquidaciones(con):
    filas = con.execute(f"SELECT {_COLUMNAS_LIQUIDACION} FROM liquidaciones ORDER BY año, mes").fetchall()
    return [_a_registro(f) for f in filas]


def buscar_ordenes_registradas(con, numeros, excluir_periodo=None):
    """Cuáles de esos números de orden ya están registrados. Devuelve {(mes, año): [números]} en orden cronológico."""
    numeros = sorted({n for n in (normalizar_orden(v) for v in numeros) if n is not None}, key=str)
    encontradas = {}
    for i in range(0, len(numeros), LOTE_CONSULTA):
        lote = numeros[i:i + LOTE_CONSULTA]
        marcas = ", ".join("?" * len(lote))
        filas = con.execute(
            f"SELECT l.mes, l.año, o.orden FROM ordenes o JOIN liquidaciones l ON l.id = o.liquidacion_id "
            f"WHERE o.orden IN ({marcas})", lote).fetchall()
        for mes, año, orden in filas:
            if (mes, año) != excluir_periodo:
                encontradas.setdefault((mes, año), set()).add(orden)
    return {p: sorted(encontradas[p], key=str) for p in sorted(encontradas, key=lambda p: (p[1], p[0]))}


def _texto(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip()


def guardar(con, mes, año, liquidacion, archivo_origen=None, archivo_generado=None, generada=None):
    """Registra la liquidación y sus órdenes. Si el período ya existía, lo reemplaza. Devuelve True si reemplazó."""
    tabla = liquidacion.tabla
    fechas = convertir_fechas(tabla["Fecha"])
    numeros = tabla[COLUMNA_ORDEN] if COLUMNA_ORDEN in tabla.columns else [None] * len(tabla)
    sucesos = tabla[COLUMNA_SUCESO] if COLUMNA_SUCESO in tabla.columns else [None] * len(tabla)
    generada = generada or datetime.now()

    with con:
        reemplazada = con.execute("DELETE FROM liquidaciones WHERE mes = ? AND año = ?", (mes, año)).rowcount > 0
        cursor = con.execute(
            "INSERT INTO liquidaciones (mes, año, generada, archivo_origen, archivo_generado, incluye_opcional, "
            "cantidad_ordenes, total) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (mes, año, generada.isoformat(sep=" ", timespec="seconds"),
             Path(archivo_origen).name if archivo_origen else None,
             str(archivo_generado) if archivo_generado else None,
             1, len(tabla), float(liquidacion.total)))
        liquidacion_id = cursor.lastrowid
        con.executemany(
            "INSERT INTO ordenes (liquidacion_id, orden, fecha, profesional, hc, trabajador, prestacion, monto, "
            "suceso) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            [
                (liquidacion_id, normalizar_orden(numero),
                 None if pd.isna(fecha) else fecha.isoformat(sep=" "),
                 _texto(profesional), _texto(hc), _texto(trabajador),
                 normalizar_prestacion("" if pd.isna(prestacion) else prestacion), float(monto),
                 normalizar_orden(suceso))
                for numero, fecha, profesional, hc, trabajador, prestacion, monto, suceso in zip(
                    numeros, fechas, tabla["Profesional"], tabla["HC"], tabla["Trabajador"],
                    tabla["Prestación"], tabla["Monto"], sucesos)
            ])
    return reemplazada


def leer_ordenes(con, liquidacion_id):
    return pd.read_sql_query(
        "SELECT orden, fecha, profesional, hc, trabajador, prestacion, monto, suceso FROM ordenes "
        "WHERE liquidacion_id = ? ORDER BY id", con, params=(liquidacion_id,))


def resumen_por_prestacion(con, liquidacion_id):
    """Lista de (prestación, cantidad, total) de una liquidación, ordenada por prestación."""
    return con.execute(
        "SELECT prestacion, COUNT(*), SUM(monto) FROM ordenes WHERE liquidacion_id = ? "
        "GROUP BY prestacion ORDER BY prestacion", (liquidacion_id,)).fetchall()


def borrar(con, liquidacion_id):
    """Borra la liquidación y sus órdenes. Devuelve True si existía."""
    with con:
        return con.execute("DELETE FROM liquidaciones WHERE id = ?", (liquidacion_id,)).rowcount > 0


def marcar_enviadas(con, periodos, fecha=None):
    """Guarda la fecha de envío de los períodos (mes, año) registrados. Devuelve cuántos marcó."""
    fecha = (fecha or datetime.now()).isoformat(sep=" ", timespec="seconds")
    with con:
        return sum(con.execute("UPDATE liquidaciones SET enviada = ? WHERE mes = ? AND año = ?",
                               (fecha, mes, año)).rowcount for mes, año in periodos)
