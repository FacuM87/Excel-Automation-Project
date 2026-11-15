from datetime import datetime

import pandas as pd
import pytest

from mis_liquidaciones.procesamiento import (
    FormatoInvalido, guardar_liquidacion, leer_planilla_origen, liquidar,
    normalizar_prestacion, obtener_nombre_archivo_unico, parsear_monto, periodo_mes_vencido,
)

from .conftest import ENCABEZADOS, fila


@pytest.mark.parametrize("entrada, esperado", [
    ("expse 1", "EXPSE 1"),
    ("  Expse   7 ", "EXPSE 7"),
    ("cues des", "CUES DES"),
    ("", ""),
])
def test_normalizar_prestacion(entrada, esperado):
    assert normalizar_prestacion(entrada) == esperado


@pytest.mark.parametrize("texto, esperado", [
    ("", 0.0),
    ("   ", 0.0),
    ("1500", 1500.0),
    ("1500.5", 1500.5),
    ("1500,5", 1500.5),
    ("1.000", 1000.0),
    ("12.345.678", 12345678.0),
    ("1.000,50", 1000.5),
    ("1.5", 1.5),
    ("1.50", 1.5),
    (" 2 000 ", 2000.0),
])
def test_parsear_monto_valido(texto, esperado):
    assert parsear_monto(texto) == esperado


@pytest.mark.parametrize("texto", ["abc", "12a", "1,2,3"])
def test_parsear_monto_invalido(texto):
    assert parsear_monto(texto) is None


def test_periodo_mes_vencido():
    assert periodo_mes_vencido(datetime(2026, 5, 10)) == (4, 2026)
    assert periodo_mes_vencido(datetime(2026, 1, 1)) == (12, 2025)


def test_lee_encabezado_en_fila_2_sin_perder_la_primera_orden(crear_planilla):
    ruta = crear_planilla([fila(1, "EXPSE 1"), fila(2, "EXPSE 2"), fila(3, "EXPSE 4")])
    df, no_cumplidas = leer_planilla_origen(ruta)
    assert len(df) == 3
    assert df.loc[0, "Trabajador"] == "Trabajador 1"
    assert no_cumplidas == 0
    assert list(df.columns) == ["Fecha", "Profesional", "HC", "Trabajador", "Prestación"]


def test_lee_planilla_sin_titulo(crear_planilla):
    ruta = crear_planilla([fila(1, "EXPSE 1")], titulo=None)
    df, _ = leer_planilla_origen(ruta)
    assert len(df) == 1


def test_descarta_filas_vacias(crear_planilla):
    ruta = crear_planilla([fila(1, "EXPSE 1"), [None] * len(ENCABEZADOS), fila(2, "EXPSE 2")])
    df, _ = leer_planilla_origen(ruta)
    assert len(df) == 2


def test_excluye_no_cumplidas(crear_planilla):
    ruta = crear_planilla([
        fila(1, "EXPSE 1", "Sí"),
        fila(2, "EXPSE 1", "Si"),
        fila(3, "EXPSE 1", " SI "),
        fila(4, "EXPSE 1", "No"),
        fila(5, "EXPSE 1", None),
    ])
    df, no_cumplidas = leer_planilla_origen(ruta)
    assert len(df) == 3
    assert no_cumplidas == 2
    assert list(df["Trabajador"]) == ["Trabajador 1", "Trabajador 2", "Trabajador 3"]


def test_sin_columna_cumplida_toma_todas(crear_planilla):
    ruta = crear_planilla([fila(1, "EXPSE 1")[:-1], fila(2, "EXPSE 2")[:-1]], encabezados=ENCABEZADOS[:-1])
    df, no_cumplidas = leer_planilla_origen(ruta)
    assert len(df) == 2
    assert no_cumplidas == 0


def test_formato_invalido_sin_encabezados(crear_planilla):
    ruta = crear_planilla([["a", "b"], ["c", "d"]], encabezados=["Columna 1", "Columna 2"])
    with pytest.raises(FormatoInvalido):
        leer_planilla_origen(ruta)


def test_formato_invalido_falta_columna(crear_planilla):
    encabezados = [c for c in ENCABEZADOS if c != "HC"]
    ruta = crear_planilla([], encabezados=encabezados)
    with pytest.raises(FormatoInvalido, match="HC"):
        leer_planilla_origen(ruta)


def test_liquidar_aplica_montos_y_total(crear_planilla):
    ruta = crear_planilla([
        fila(1, "EXPSE 1"),
        fila(2, "expse 1"),
        fila(3, " EXPSE  2 "),
        fila(4, "ATEPSIQ"),
    ])
    liquidacion = liquidar(ruta, {"Expse 1": 100.0, "EXPSE 2": 250.5, "ATEPSIQ": 1000.0})
    assert list(liquidacion.tabla["Monto"]) == [100.0, 100.0, 250.5, 1000.0]
    assert liquidacion.total == 1450.5
    assert liquidacion.sin_monto == []
    assert liquidacion.cantidad_sin_monto == 0


def test_liquidar_prestaciones_sin_monto(crear_planilla):
    ruta = crear_planilla([
        fila(1, "EXPSE 1"),
        fila(2, "PSICOMEY"),
        fila(3, "PSICOMEY"),
        fila(4, "NUEVA"),
        fila(5, None),
        fila(6, " nueva "),
    ])
    liquidacion = liquidar(ruta, {"EXPSE 1": 100.0}, incluir_opcional=True)
    assert liquidacion.total == 100.0
    assert liquidacion.sin_monto == ["", "NUEVA", "PSICOMEY"]
    assert liquidacion.cantidad_sin_monto == 5
    assert list(liquidacion.tabla["Monto"]) == [100.0, 0, 0, 0, 0, 0]


def planilla_con_psicomey(crear_planilla):
    return crear_planilla([
        fila(1, "EXPSE 1"),
        fila(2, "PSICOMEY"),
        fila(3, "EXPSE 2"),
        fila(4, " psicomey "),
        fila(5, "EXPSE 1"),
    ])


def test_sin_incluir_opcional_quita_psicomey(crear_planilla):
    liquidacion = liquidar(planilla_con_psicomey(crear_planilla), {"EXPSE 1": 100.0, "EXPSE 2": 200.0, "PSICOMEY": 500.0})
    assert liquidacion.excluidas_opcional == 2
    assert list(liquidacion.tabla["Trabajador"]) == ["Trabajador 1", "Trabajador 3", "Trabajador 5"]
    assert list(liquidacion.tabla["Monto"]) == [100.0, 200.0, 100.0]
    assert liquidacion.total == 400.0
    assert list(liquidacion.tabla.index) == [0, 1, 2]


def test_sin_incluir_opcional_no_avisa_psicomey_sin_monto(crear_planilla):
    liquidacion = liquidar(planilla_con_psicomey(crear_planilla), {"EXPSE 1": 100.0, "EXPSE 2": 200.0})
    assert liquidacion.sin_monto == []
    assert liquidacion.cantidad_sin_monto == 0


def test_incluir_opcional_liquida_psicomey(crear_planilla):
    liquidacion = liquidar(planilla_con_psicomey(crear_planilla), {"EXPSE 1": 100.0, "EXPSE 2": 200.0, "PSICOMEY": 500.0},
                           incluir_opcional=True)
    assert liquidacion.excluidas_opcional == 0
    assert len(liquidacion.tabla) == 5
    assert liquidacion.total == 1400.0


def test_incluir_opcional_sin_monto_avisa(crear_planilla):
    liquidacion = liquidar(planilla_con_psicomey(crear_planilla), {"EXPSE 1": 100.0, "EXPSE 2": 200.0}, incluir_opcional=True)
    assert liquidacion.sin_monto == ["PSICOMEY"]
    assert liquidacion.cantidad_sin_monto == 2


def test_guardar_sin_psicomey(crear_planilla, tmp_path):
    liquidacion = liquidar(planilla_con_psicomey(crear_planilla), {"EXPSE 1": 100.0, "EXPSE 2": 200.0})
    df = pd.read_excel(guardar_liquidacion(liquidacion, tmp_path, 3, 2026))
    assert len(df) == 4
    assert "PSICOMEY" not in set(df["Prestación"].astype(str).str.upper())
    assert df.iloc[-1]["Monto"] == 400.0


def test_liquidar_informa_no_cumplidas(crear_planilla):
    ruta = crear_planilla([fila(1, "EXPSE 1"), fila(2, "EXPSE 1", "No")])
    liquidacion = liquidar(ruta, {"EXPSE 1": 100.0})
    assert liquidacion.no_cumplidas == 1
    assert liquidacion.total == 100.0


def test_guardar_liquidacion_agrega_fila_total(crear_planilla, tmp_path):
    ruta = crear_planilla([fila(1, "EXPSE 1"), fila(2, "EXPSE 2")])
    liquidacion = liquidar(ruta, {"EXPSE 1": 100.0, "EXPSE 2": 200.0})
    salida = guardar_liquidacion(liquidacion, tmp_path, 3, 2026)

    assert salida.name == "Liquidación Marzo 2026.xlsx"
    df = pd.read_excel(salida)
    assert list(df.columns) == ["Fecha", "Profesional", "HC", "Trabajador", "Prestación", "Monto"]
    assert len(df) == 3
    assert df.iloc[-1]["Fecha"] == "TOTAL"
    assert df.iloc[-1]["Monto"] == 300.0
    assert df.iloc[:-1]["Monto"].sum() == df.iloc[-1]["Monto"]


def test_nombre_archivo_unico(tmp_path):
    primero = obtener_nombre_archivo_unico(tmp_path, 12, 2025)
    assert primero.name == "Liquidación Diciembre 2025.xlsx"
    primero.touch()
    segundo = obtener_nombre_archivo_unico(tmp_path, 12, 2025)
    assert segundo.name == "Liquidación Diciembre 2025 (1).xlsx"
    segundo.touch()
    assert obtener_nombre_archivo_unico(tmp_path, 12, 2025).name == "Liquidación Diciembre 2025 (2).xlsx"
