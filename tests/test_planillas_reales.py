"""Tests con los Excel recibidos en docs/. Solo corren de forma local: si los archivos no están, se saltean."""
import pytest

from mis_liquidaciones import base_datos, importacion
from mis_liquidaciones.config import PRESTACIONES_POR_DEFECTO
from mis_liquidaciones.procesamiento import (
    detectar_periodos, leer_planilla_origen, liquidar, normalizar_prestacion, ordenes_repetidas,
)

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

PERIODO_POR_ARCHIVO = {
    "dgOrdenes.XLSX": (3, 2026),
    "dgOrdenes (1).XLSX": (3, 2026),
    "dgOrdenes (2).XLSX": (4, 2026),
    "dgOrdenes (3).XLSX": (5, 2026),
    "dgOrdenes (4).XLSX": (6, 2026),
    "dgOrdenes (5).XLSX": (7, 2026),
    "dgOrdenes (6).XLSX": (8, 2026),
    "dgOrdenes (7).XLSX": (9, 2026),
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


def test_total_es_la_suma(planilla_real):
    liquidacion = liquidar(planilla_real, MONTOS)
    assert liquidacion.cantidad_sin_monto == 0
    assert liquidacion.total == pytest.approx(liquidacion.tabla["Monto"].sum())
    assert (liquidacion.tabla["Monto"] > 0).all()


def test_quita_los_psicomey_duplicados(planilla_real):
    """En las muestras todos los PSICOMEY duplican un EXPSE 7 (R3)."""
    psicomey = PSICOMEY_POR_ARCHIVO.get(planilla_real.name, 0)
    liquidacion = liquidar(planilla_real, MONTOS)
    assert liquidacion.psicomey_duplicados == psicomey
    assert len(liquidacion.tabla) == ORDENES_POR_ARCHIVO[planilla_real.name] - psicomey
    assert "PSICOMEY" not in set(liquidacion.tabla["Prestación"].map(normalizar_prestacion))


def test_detecta_el_periodo(planilla_real):
    df, _ = leer_planilla_origen(planilla_real)
    periodos = detectar_periodos(df)
    assert len(periodos) == 1
    assert (periodos[0].mes, periodos[0].año) == PERIODO_POR_ARCHIVO[planilla_real.name]
    assert periodos[0].ordenes == ORDENES_POR_ARCHIVO[planilla_real.name]


def test_liquidar_por_periodo_detectado_da_lo_mismo(planilla_real):
    completa = liquidar(planilla_real, MONTOS)
    por_periodo = liquidar(planilla_real, MONTOS, periodo=PERIODO_POR_ARCHIVO[planilla_real.name])
    assert completa.tabla.equals(por_periodo.tabla)


def _real(nombre):
    ruta = DIRECTORIO_DOCS / nombre
    if not ruta.exists():
        pytest.skip(f"{nombre} no está en docs/")
    return ruta


def test_sin_ordenes_repetidas_en_el_archivo(planilla_real):
    df, _ = leer_planilla_origen(planilla_real)
    assert df["Orden"].notna().all()
    assert ordenes_repetidas(df) == []


def test_guardar_en_la_base(planilla_real, tmp_path):
    liquidacion = liquidar(planilla_real, MONTOS)
    mes, año = PERIODO_POR_ARCHIVO[planilla_real.name]
    with base_datos.abrir(tmp_path / "base.db") as con:
        base_datos.guardar(con, mes, año, liquidacion, planilla_real)
        registrada = base_datos.buscar_liquidacion(con, mes, año)
        ordenes = base_datos.leer_ordenes(con, registrada.id)
    con.close()
    psicomey = PSICOMEY_POR_ARCHIVO.get(planilla_real.name, 0)
    assert registrada.cantidad_ordenes == ORDENES_POR_ARCHIVO[planilla_real.name] - psicomey
    assert ordenes["suceso"].notna().all()
    assert registrada.total == pytest.approx(liquidacion.total)
    assert ordenes["monto"].sum() == pytest.approx(liquidacion.total)
    assert ordenes["fecha"].notna().all()
    assert ordenes["trabajador"].notna().all()


def test_caso_h6_reemplaza_marzo_sin_duplicar(tmp_path):
    parcial = liquidar(_real("dgOrdenes.XLSX"), MONTOS)
    completa = liquidar(_real("dgOrdenes (1).XLSX"), MONTOS)
    con = base_datos.abrir(tmp_path / "base.db")
    base_datos.guardar(con, 3, 2026, parcial)
    assert base_datos.buscar_ordenes_registradas(con, completa.tabla["Orden"], excluir_periodo=(3, 2026)) == {}
    assert len(base_datos.buscar_ordenes_registradas(con, completa.tabla["Orden"])[(3, 2026)]) == 40
    assert base_datos.guardar(con, 3, 2026, completa)
    assert base_datos.buscar_liquidacion(con, 3, 2026).cantidad_ordenes == 142
    assert con.execute("SELECT COUNT(*) FROM ordenes").fetchone()[0] == 142
    con.close()


def test_archivos_numerados_sin_duplicados_entre_si(tmp_path):
    con = base_datos.abrir(tmp_path / "base.db")
    for nombre in sorted(ORDENES_POR_ARCHIVO):
        if nombre == "dgOrdenes.XLSX":
            continue
        liquidacion = liquidar(_real(nombre), MONTOS)
        mes, año = PERIODO_POR_ARCHIVO[nombre]
        assert base_datos.buscar_ordenes_registradas(con, liquidacion.tabla["Orden"], (mes, año)) == {}
        base_datos.guardar(con, mes, año, liquidacion)
    assert len(base_datos.listar_liquidaciones(con)) == 7
    con.close()


def test_importar_todo_el_historial(tmp_path):
    rutas = [_real(nombre) for nombre in ORDENES_POR_ARCHIVO]
    analisis = importacion.analizar_archivos(rutas, MONTOS)
    assert analisis.errores == []
    assert [(p.mes, p.año) for p in analisis.periodos] == [(m, 2026) for m in range(3, 10)]
    assert analisis.periodos[0].duplicadas_descartadas == 40
    con = base_datos.abrir(tmp_path / "base.db")
    assert sum(p.liquidacion.psicomey_duplicados for p in analisis.periodos) == 6
    importacion.importar(con, analisis.periodos)
    assert base_datos.buscar_liquidacion(con, 3, 2026).cantidad_ordenes == 142
    assert con.execute("SELECT COUNT(*) FROM ordenes WHERE prestacion = 'PSICOMEY'").fetchone()[0] == 0
    total, distintas = con.execute("SELECT COUNT(*), COUNT(DISTINCT orden) FROM ordenes").fetchone()
    assert total == distintas
    assert len(base_datos.listar_liquidaciones(con)) == 7
    con.close()
