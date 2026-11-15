import pytest
from openpyxl import load_workbook

from mis_liquidaciones import base_datos, reportes
from mis_liquidaciones.procesamiento import liquidar

from .conftest import fila

MONTOS = {"EXPSE 1": 100.0, "EXPSE 2": 200.0, "PSICOMEY": 50.0}


@pytest.fixture
def con(tmp_path):
    conexion = base_datos.abrir(tmp_path / "liquidaciones.db")
    yield conexion
    conexion.close()


@pytest.fixture
def registrar(con, crear_planilla):
    """Registra un período con las prestaciones indicadas (una orden por prestación)."""
    contador = iter(range(1, 10_000))

    def _registrar(mes, año, prestaciones):
        fecha = f"01/{mes:02d}/{año}"
        filas = [fila(next(contador), p, fecha=fecha) for p in prestaciones]
        liquidacion = liquidar(crear_planilla(filas, nombre=f"{mes}-{año}.xlsx"), MONTOS, incluir_opcional=True)
        base_datos.guardar(con, mes, año, liquidacion, incluye_opcional=True)
    return _registrar


def test_reporte_mensual(con, registrar):
    registrar(3, 2026, ["EXPSE 2", "EXPSE 1", "EXPSE 2", "PSICOMEY"])
    reporte = reportes.reporte_mensual(con, 3, 2026)
    assert reporte.nombre == "Marzo 2026"
    assert reporte.por_prestacion == [("EXPSE 1", 1, 100.0), ("EXPSE 2", 2, 400.0)]
    assert (reporte.cantidad, reporte.total) == (3, 500.0)
    assert reporte.excluidas_opcional == 1
    assert reportes.nota_opcional(reporte) == "No incluye 1 orden(es) de PSICOMEY."


def test_reporte_mensual_inexistente(con, registrar):
    registrar(3, 2026, ["EXPSE 1"])
    assert reportes.reporte_mensual(con, 4, 2026) is None


def test_reporte_mensual_solo_opcional(con, registrar):
    registrar(3, 2026, ["PSICOMEY"])
    reporte = reportes.reporte_mensual(con, 3, 2026)
    assert (reporte.por_prestacion, reporte.cantidad, reporte.total) == ([], 0, 0)


def test_reporte_anual(con, registrar):
    registrar(3, 2026, ["EXPSE 1", "EXPSE 2", "PSICOMEY"])
    registrar(5, 2026, ["EXPSE 1", "EXPSE 1"])
    registrar(5, 2025, ["EXPSE 2"])
    reporte = reportes.reporte_anual(con, 2026)
    assert len(reporte.por_mes) == 12
    assert reporte.por_mes[0] == (1, None, None)
    assert reporte.por_mes[2] == (3, 2, 300.0)
    assert reporte.por_mes[4] == (5, 2, 200.0)
    assert reporte.por_prestacion == [("EXPSE 1", 3, 300.0), ("EXPSE 2", 1, 200.0)]
    assert (reporte.cantidad, reporte.total, reporte.excluidas_opcional) == (4, 500.0, 1)


def test_reporte_anual_mes_solo_opcional_figura_en_cero(con, registrar):
    registrar(3, 2026, ["PSICOMEY"])
    assert reportes.reporte_anual(con, 2026).por_mes[2] == (3, 0, 0.0)


def test_años_registrados(con, registrar):
    assert reportes.años_registrados(con) == []
    registrar(3, 2025, ["EXPSE 1"])
    registrar(3, 2026, ["EXPSE 1"])
    registrar(4, 2026, ["EXPSE 1"])
    assert reportes.años_registrados(con) == [2026, 2025]


def _valores(ruta):
    hoja = load_workbook(ruta).active
    return hoja, [tuple(c.value for c in f) for f in hoja.iter_rows(min_row=5)]


def test_exportar_mensual(con, registrar, tmp_path):
    registrar(3, 2026, ["EXPSE 1", "EXPSE 2", "PSICOMEY"])
    ruta = tmp_path / "mensual.xlsx"
    reportes.exportar_mensual(reportes.reporte_mensual(con, 3, 2026), ruta)
    hoja, filas = _valores(ruta)
    assert hoja["A1"].value == "Reporte mensual — Marzo 2026"
    assert hoja["A3"].value == "No incluye 1 orden(es) de PSICOMEY."
    assert filas == [("Prestación", "Cantidad", "Total"), ("EXPSE 1", 1, 100), ("EXPSE 2", 1, 200),
                     ("TOTAL", 2, 300)]
    assert hoja["C8"].number_format == reportes.FORMATO_MONTO
    assert hoja["A8"].font.bold
    assert len(hoja._charts) == 1


def test_exportar_mensual_sin_datos_no_agrega_grafico(con, registrar, tmp_path):
    registrar(3, 2026, ["PSICOMEY"])
    ruta = tmp_path / "mensual.xlsx"
    reportes.exportar_mensual(reportes.reporte_mensual(con, 3, 2026), ruta)
    assert load_workbook(ruta).active._charts == []


def test_exportar_anual(con, registrar, tmp_path):
    registrar(3, 2026, ["EXPSE 1", "EXPSE 2"])
    ruta = tmp_path / "anual.xlsx"
    reportes.exportar_anual(reportes.reporte_anual(con, 2026), ruta)
    hoja, filas = _valores(ruta)
    assert hoja["A1"].value == "Reporte anual 2026"
    assert hoja["A3"].value is None
    assert filas[0] == ("Por mes", None, None)
    assert filas[2] == ("Enero", None, None)
    assert filas[4] == ("Marzo", 2, 300)
    assert filas[14] == ("TOTAL", 2, 300)
    inicio = filas.index(("Por prestación", None, None))
    assert filas[inicio + 1:inicio + 5] == [("Prestación", "Cantidad", "Total"), ("EXPSE 1", 1, 100),
                                            ("EXPSE 2", 1, 200), ("TOTAL", 2, 300)]
    assert isinstance(hoja.cell(5 + 4, 3).value, (int, float))
    assert len(hoja._charts) == 2
