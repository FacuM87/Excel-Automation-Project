from datetime import datetime

import pandas as pd
import pytest

from mis_liquidaciones import base_datos
from mis_liquidaciones.procesamiento import Liquidacion, liquidar

from .conftest import ENCABEZADOS, fila

MONTOS = {"EXPSE 1": 100.0, "EXPSE 2": 200.0, "PSICOMEY": 50.0}


@pytest.fixture
def con(tmp_path):
    conexion = base_datos.abrir(tmp_path / "liquidaciones.db")
    yield conexion
    conexion.close()


def _liquidar(crear_planilla, numeros, fecha="01/03/2026", nombre="planilla.xlsx", **kwargs):
    filas = [fila(n, "EXPSE 1" if n % 2 else "EXPSE 2", fecha=fecha) for n in numeros]
    return liquidar(crear_planilla(filas, nombre=nombre), MONTOS, **kwargs)


def test_abrir_dos_veces_no_rompe(tmp_path):
    ruta = tmp_path / "liquidaciones.db"
    base_datos.abrir(ruta).close()
    con = base_datos.abrir(ruta)
    assert con.execute("PRAGMA user_version").fetchone()[0] == base_datos.VERSION_ESQUEMA
    assert base_datos.listar_liquidaciones(con) == []
    con.close()


def test_guardar_y_leer(con, crear_planilla, tmp_path):
    liquidacion = _liquidar(crear_planilla, [1, 2, 3])
    generada = datetime(2026, 4, 10, 9, 30)
    reemplazada = base_datos.guardar(
        con, 3, 2026, liquidacion, tmp_path / "dgOrdenes.XLSX", tmp_path / "Liquidación Marzo 2026.xlsx",
        incluye_opcional=True, generada=generada)

    assert not reemplazada
    registrada = base_datos.buscar_liquidacion(con, 3, 2026)
    assert registrada.nombre == "Marzo 2026"
    assert registrada.generada == generada
    assert registrada.archivo_origen == "dgOrdenes.XLSX"
    assert registrada.archivo_generado.endswith("Liquidación Marzo 2026.xlsx")
    assert registrada.incluye_opcional
    assert registrada.cantidad_ordenes == 3
    assert registrada.total == 400.0

    ordenes = base_datos.leer_ordenes(con, registrada.id)
    assert list(ordenes["orden"]) == [1, 2, 3]
    assert list(ordenes["prestacion"]) == ["EXPSE 1", "EXPSE 2", "EXPSE 1"]
    assert list(ordenes["monto"]) == [100.0, 200.0, 100.0]
    assert ordenes.loc[0, "trabajador"] == "Trabajador 1"
    assert ordenes.loc[0, "hc"] == "1001"
    assert ordenes.loc[0, "profesional"] == "Profesional Prueba"
    assert ordenes.loc[0, "fecha"].startswith("2026-03-01")


def test_buscar_periodo_inexistente(con):
    assert base_datos.buscar_liquidacion(con, 3, 2026) is None


def test_reemplazar_borra_las_ordenes_anteriores(con, crear_planilla):
    base_datos.guardar(con, 3, 2026, _liquidar(crear_planilla, [1, 2], nombre="a.xlsx"))
    assert base_datos.guardar(con, 3, 2026, _liquidar(crear_planilla, [1, 2, 3, 4], nombre="b.xlsx"))

    assert len(base_datos.listar_liquidaciones(con)) == 1
    assert base_datos.buscar_liquidacion(con, 3, 2026).cantidad_ordenes == 4
    assert con.execute("SELECT COUNT(*) FROM ordenes").fetchone()[0] == 4
    huerfanas = con.execute(
        "SELECT COUNT(*) FROM ordenes WHERE liquidacion_id NOT IN (SELECT id FROM liquidaciones)").fetchone()[0]
    assert huerfanas == 0


def test_buscar_ordenes_registradas(con, crear_planilla):
    base_datos.guardar(con, 2, 2026, _liquidar(crear_planilla, [1, 2], fecha="01/02/2026", nombre="feb.xlsx"))
    base_datos.guardar(con, 3, 2026, _liquidar(crear_planilla, [3, 4], nombre="mar.xlsx"))

    encontradas = base_datos.buscar_ordenes_registradas(con, [4, 2.0, "1", 99, None])
    assert encontradas == {(2, 2026): [1, 2], (3, 2026): [4]}
    assert list(encontradas) == [(2, 2026), (3, 2026)]
    assert base_datos.buscar_ordenes_registradas(con, [3, 4], excluir_periodo=(3, 2026)) == {}
    assert base_datos.buscar_ordenes_registradas(con, []) == {}


def test_buscar_muchas_ordenes(con):
    tabla = pd.DataFrame({
        "Orden": list(range(1200)), "Fecha": ["01/03/2026"] * 1200, "Profesional": "P", "HC": 1,
        "Trabajador": "T", "Prestación": "EXPSE 1", "Monto": 1.0,
    })
    base_datos.guardar(con, 3, 2026, Liquidacion(tabla=tabla, total=1200.0))
    encontradas = base_datos.buscar_ordenes_registradas(con, range(0, 2000))
    assert len(encontradas[(3, 2026)]) == 1200


def test_guarda_sin_columna_orden(con, crear_planilla):
    encabezados = [c for c in ENCABEZADOS if c != "Orden"]
    ruta = crear_planilla([fila(1, "EXPSE 1")[1:], fila(2, "EXPSE 2")[1:]], encabezados=encabezados)
    base_datos.guardar(con, 3, 2026, liquidar(ruta, MONTOS))
    ordenes = base_datos.leer_ordenes(con, base_datos.buscar_liquidacion(con, 3, 2026).id)
    assert ordenes["orden"].isna().all()
    assert len(ordenes) == 2


def test_psicomey_incluido_se_guarda(con, crear_planilla):
    ruta = crear_planilla([fila(1, "EXPSE 1"), fila(2, "psicomey")])
    base_datos.guardar(con, 3, 2026, liquidar(ruta, MONTOS, incluir_opcional=True), incluye_opcional=True)
    ordenes = base_datos.leer_ordenes(con, base_datos.buscar_liquidacion(con, 3, 2026).id)
    assert list(ordenes["prestacion"]) == ["EXPSE 1", "PSICOMEY"]


def test_resumen_por_prestacion(con, crear_planilla):
    base_datos.guardar(con, 3, 2026, _liquidar(crear_planilla, [1, 2, 3]))
    registrada = base_datos.buscar_liquidacion(con, 3, 2026)
    assert base_datos.resumen_por_prestacion(con, registrada.id) == [("EXPSE 1", 2, 200.0), ("EXPSE 2", 1, 200.0)]
    assert base_datos.resumen_por_prestacion(con, 999) == []


def test_borrar_quita_la_liquidacion_y_sus_ordenes(con, crear_planilla):
    base_datos.guardar(con, 2, 2026, _liquidar(crear_planilla, [1, 2], fecha="01/02/2026", nombre="feb.xlsx"))
    base_datos.guardar(con, 3, 2026, _liquidar(crear_planilla, [3, 4, 5], nombre="mar.xlsx"))
    marzo = base_datos.buscar_liquidacion(con, 3, 2026)

    assert base_datos.borrar(con, marzo.id)
    assert base_datos.buscar_liquidacion(con, 3, 2026) is None
    assert [r.nombre for r in base_datos.listar_liquidaciones(con)] == ["Febrero 2026"]
    assert con.execute("SELECT COUNT(*) FROM ordenes").fetchone()[0] == 2
    assert not base_datos.borrar(con, marzo.id)
