import sqlite3
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
        generada=generada)

    assert not reemplazada
    registrada = base_datos.buscar_liquidacion(con, 3, 2026)
    assert registrada.nombre == "Marzo 2026"
    assert registrada.generada == generada
    assert registrada.archivo_origen == "dgOrdenes.XLSX"
    assert registrada.archivo_generado.endswith("Liquidación Marzo 2026.xlsx")
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


def test_psicomey_unico_se_guarda_con_suceso(con, crear_planilla):
    ruta = crear_planilla([fila(1, "EXPSE 1"), fila(2, "psicomey")])
    base_datos.guardar(con, 3, 2026, liquidar(ruta, MONTOS))
    ordenes = base_datos.leer_ordenes(con, base_datos.buscar_liquidacion(con, 3, 2026).id)
    assert list(ordenes["prestacion"]) == ["EXPSE 1", "PSICOMEY"]
    assert list(con.execute("SELECT suceso FROM ordenes")) == [("Suceso",), ("Suceso",)]


ESQUEMA_V1 = """
CREATE TABLE liquidaciones (
    id INTEGER PRIMARY KEY, mes INTEGER NOT NULL, año INTEGER NOT NULL, generada TEXT NOT NULL,
    archivo_origen TEXT, archivo_generado TEXT, incluye_opcional INTEGER NOT NULL,
    cantidad_ordenes INTEGER NOT NULL, total REAL NOT NULL, UNIQUE (mes, año));
CREATE TABLE ordenes (
    id INTEGER PRIMARY KEY, liquidacion_id INTEGER NOT NULL REFERENCES liquidaciones(id) ON DELETE CASCADE,
    orden, fecha TEXT, profesional TEXT, hc TEXT, trabajador TEXT, prestacion TEXT NOT NULL, monto REAL NOT NULL);
PRAGMA user_version = 1;
"""


def test_base_vieja_se_limpia_al_abrir(tmp_path):
    ruta = tmp_path / "vieja.db"
    vieja = sqlite3.connect(ruta)
    vieja.executescript(ESQUEMA_V1)
    vieja.execute("INSERT INTO liquidaciones VALUES (1, 3, 2026, '2026-04-01 10:00:00', NULL, NULL, 1, 5, 1950)")
    vieja.execute("INSERT INTO liquidaciones VALUES (2, 4, 2026, '2026-05-01 10:00:00', NULL, NULL, 1, 1, 50)")
    hora = "2026-03-05 10:15:30.187000"
    vieja.executemany("INSERT INTO ordenes (liquidacion_id, orden, fecha, hc, prestacion, monto) VALUES (?, ?, ?, ?, ?, ?)", [
        (1, 1, hora, "500", "EXPSE 7", 700), (1, 2, hora, "500", "PSICOMEY", 50),
        (1, 3, hora, "501", "PSICOMEY", 50), (1, 4, "2026-03-06 09:00:00", "502", "EXPSE 7", 700),
        (1, 5, "2026-03-06 09:00:00", "502", "EXPSE 1", 450), (2, 6, hora, "500", "PSICOMEY", 50),
    ])
    vieja.commit()
    vieja.close()

    con = base_datos.abrir(ruta)
    assert con.execute("PRAGMA user_version").fetchone()[0] == base_datos.VERSION_ESQUEMA
    marzo, abril = base_datos.listar_liquidaciones(con)
    assert (marzo.cantidad_ordenes, marzo.total) == (4, 1900)
    assert (abril.cantidad_ordenes, abril.total) == (1, 50)
    assert list(base_datos.leer_ordenes(con, marzo.id)["orden"]) == [1, 3, 4, 5]
    assert base_datos.limpiar_psicomey_duplicados(con) == 0
    con.close()
    base_datos.abrir(ruta).close()


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
