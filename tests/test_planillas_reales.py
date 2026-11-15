"""Tests con los Excel recibidos en docs/. Solo corren de forma local: si los archivos no están, se saltean."""
import pytest

from mis_liquidaciones.config import PRESTACIONES_POR_DEFECTO
from mis_liquidaciones.procesamiento import leer_planilla_origen, liquidar, normalizar_prestacion

from .conftest import DIRECTORIO_DOCS

ORDENES_POR_ARCHIVO = {
    "dgOrdenes.XLSX": 40,
    "dgOrdenes (1).XLSX": 143,
    "dgOrdenes (2).XLSX": 189,
    "dgOrdenes (3).XLSX": 161,
    "dgOrdenes (4).XLSX": 165,
    "dgOrdenes (5).XLSX": 177,
    "dgOrdenes (6).XLSX": 151,
    "dgOrdenes (7).XLSX": 93,
}

PSICOMEY_POR_ARCHIVO = {
    "dgOrdenes (1).XLSX": 1,
    "dgOrdenes (3).XLSX": 3,
    "dgOrdenes (6).XLSX": 2,
}

MONTOS ={p: float(100 * (i + 1)) for i, p in enumerate(PRESTACIONES_POR_DEFECTO)}


@pytest.fixture(params=sorted(ORDENES_POR_ARCHIVO))
def planilla_real(request):
    ruta = DIRECTORIO_DOCS / request.param
    if not ruta.exists():
        pytest.skip(f"{request.param} no está en docs/ (los Excel reales solo existen de forma local)")
    return ruta


def test_cantidad_de_ordenes(planilla_real):
    df, no_cumplidas = leer_planilla_origen(planilla_real)
    assert len(df) == ORDENES_POR_ARCHIVO[planilla_real.name]
    assert no_cumplidas == 0


def test_sin_filas_sin_fecha(planilla_real):
    df, _ = leer_planilla_origen(planilla_real)
    assert df["Fecha"].notna().all()


def test_prestaciones_conocidas(planilla_real):
    df, _ = leer_planilla_origen(planilla_real)
    assert set(df["Prestación"].map(normalizar_prestacion)) <= set(PRESTACIONES_POR_DEFECTO)


@pytest.mark.parametrize("incluir_opcional", [False, True], ids=["sin_psicomey", "con_psicomey"])
def test_total_es_la_suma(planilla_real, incluir_opcional):
    liquidacion = liquidar(planilla_real, MONTOS, incluir_opcional)
    assert liquidacion.cantidad_sin_monto == 0
    assert liquidacion.total == pytest.approx(liquidacion.tabla["Monto"].sum())
    assert (liquidacion.tabla["Monto"] > 0).all()


def test_sin_incluir_psicomey_quita_sus_ordenes(planilla_real):
    psicomey = PSICOMEY_POR_ARCHIVO.get(planilla_real.name, 0)
    liquidacion = liquidar(planilla_real, MONTOS)
    assert liquidacion.excluidas_opcional == psicomey
    assert len(liquidacion.tabla) == ORDENES_POR_ARCHIVO[planilla_real.name] - psicomey
    assert "PSICOMEY" not in set(liquidacion.tabla["Prestación"].map(normalizar_prestacion))


def test_incluir_psicomey_mantiene_todas_las_ordenes(planilla_real):
    liquidacion = liquidar(planilla_real, MONTOS, incluir_opcional=True)
    assert liquidacion.excluidas_opcional == 0
    assert len(liquidacion.tabla) == ORDENES_POR_ARCHIVO[planilla_real.name]
