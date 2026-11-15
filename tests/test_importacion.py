import pandas as pd
import pytest

from mis_liquidaciones import base_datos, importacion
from mis_liquidaciones.importacion import analizar_archivos, importar, verificar_en_base

from .conftest import ENCABEZADOS, fila

MONTOS = {"EXPSE 1": 100.0, "EXPSE 2": 200.0, "PSICOMEY": 50.0}


@pytest.fixture
def base(tmp_path):
    con = base_datos.abrir(tmp_path / "base.db")
    yield con
    con.close()


def _ordenes(con):
    return sorted(r[0] for r in con.execute("SELECT orden FROM ordenes"))


def test_un_archivo_un_mes(crear_planilla, base):
    ruta = crear_planilla([fila(1, "EXPSE 1"), fila(2, "EXPSE 2")])
    analisis = analizar_archivos([ruta], MONTOS)
    assert analisis.errores == []
    [periodo] = analisis.periodos
    assert (periodo.mes, periodo.año, periodo.archivos) == (3, 2026, ["planilla.xlsx"])
    assert periodo.liquidacion.total == 300

    resultado = importar(base, analisis.periodos)
    assert resultado.registrados == ["Marzo 2026"]
    registrada = base_datos.buscar_liquidacion(base, 3, 2026)
    assert (registrada.cantidad_ordenes, registrada.total) == (2, 300)
    assert registrada.archivo_origen == "planilla.xlsx"
    assert registrada.archivo_generado is None


def test_mismo_mes_en_dos_archivos_se_une_sin_duplicar(crear_planilla, base):
    parcial = crear_planilla([fila(1, "EXPSE 1"), fila(2, "EXPSE 1")], nombre="a.xlsx")
    completo = crear_planilla([fila(1, "EXPSE 1"), fila(2, "EXPSE 1"), fila(3, "EXPSE 2")], nombre="b.xlsx")
    analisis = analizar_archivos([completo, parcial], MONTOS)
    [periodo] = analisis.periodos
    assert periodo.archivos == ["a.xlsx", "b.xlsx"]
    assert periodo.duplicadas_descartadas == 2
    assert periodo.liquidacion.total == 400

    importar(base, analisis.periodos)
    assert _ordenes(base) == [1, 2, 3]
    assert base_datos.buscar_liquidacion(base, 3, 2026).archivo_origen == "b.xlsx"


def test_origen_nombra_los_archivos_que_aportan_ordenes(crear_planilla, base):
    a = crear_planilla([fila(1, "EXPSE 1"), fila(2, "EXPSE 1")], nombre="a.xlsx")
    b = crear_planilla([fila(2, "EXPSE 1"), fila(3, "EXPSE 2"), fila(4, "EXPSE 2")], nombre="b.xlsx")
    [periodo] = analizar_archivos([a, b], MONTOS).periodos
    assert (periodo.archivos, periodo.origen, periodo.duplicadas_descartadas) == (["a.xlsx", "b.xlsx"], "a.xlsx, b.xlsx", 1)


def test_archivo_con_dos_meses(crear_planilla, base):
    ruta = crear_planilla([
        fila(1, "EXPSE 1", fecha="31/03/2026"), fila(2, "EXPSE 2", fecha="01/04/2026"),
        fila(3, "EXPSE 1", fecha="sin fecha")])
    analisis = analizar_archivos([ruta], MONTOS)
    assert [p.nombre for p in analisis.periodos] == ["Marzo 2026", "Abril 2026"]
    assert analisis.sin_fecha == 1
    importar(base, analisis.periodos)
    assert [(l.mes, l.total) for l in base_datos.listar_liquidaciones(base)] == [(3, 100), (4, 200)]


def test_ordena_los_meses_y_las_fechas(crear_planilla):
    tardio = crear_planilla([fila(5, "EXPSE 1", fecha="20/05/2026"), fila(4, "EXPSE 1", fecha="02/05/2026")],
                            nombre="a.xlsx")
    temprano = crear_planilla([fila(1, "EXPSE 1", fecha="10/01/2026")], nombre="b.xlsx")
    analisis = analizar_archivos([tardio, temprano], MONTOS)
    assert [p.nombre for p in analisis.periodos] == ["Enero 2026", "Mayo 2026"]
    assert list(analisis.periodos[1].liquidacion.tabla["Orden"]) == [4, 5]


def test_archivo_invalido_no_frena_el_resto(crear_planilla, tmp_path, base):
    valido = crear_planilla([fila(1, "EXPSE 1")])
    invalido = tmp_path / "roto.xlsx"
    invalido.write_text("no es un excel")
    analisis = analizar_archivos([valido, invalido], MONTOS)
    assert [nombre for nombre, _ in analisis.errores] == ["roto.xlsx"]
    assert len(analisis.periodos) == 1


def test_formato_invalido_queda_en_errores(crear_planilla):
    ruta = crear_planilla([[1, "x"]], encabezados=["Orden", "Otra"])
    analisis = analizar_archivos([ruta], MONTOS)
    assert analisis.periodos == []
    assert analisis.errores[0][0] == "planilla.xlsx"


def test_periodo_registrado_se_saltea_o_se_reemplaza(crear_planilla, base):
    viejo = analizar_archivos([crear_planilla([fila(1, "EXPSE 1")], nombre="viejo.xlsx")], MONTOS)
    importar(base, viejo.periodos)
    nuevo = analizar_archivos([crear_planilla([fila(1, "EXPSE 1"), fila(2, "EXPSE 2")], nombre="nuevo.xlsx")], MONTOS)
    verificar_en_base(base, nuevo.periodos)
    assert nuevo.periodos[0].registrado.cantidad_ordenes == 1

    resultado = importar(base, nuevo.periodos)
    assert (resultado.registrados, resultado.salteados) == ([], ["Marzo 2026"])
    assert _ordenes(base) == [1]

    resultado = importar(base, nuevo.periodos, reemplazar=True)
    assert resultado.reemplazados == ["Marzo 2026"]
    assert _ordenes(base) == [1, 2]


def test_avisa_ordenes_de_otro_periodo(crear_planilla, base):
    importar(base, analizar_archivos([crear_planilla([fila(1, "EXPSE 1")], nombre="marzo.xlsx")], MONTOS).periodos)
    abril = analizar_archivos([crear_planilla([fila(1, "EXPSE 1", fecha="01/04/2026")], nombre="abril.xlsx")], MONTOS)
    verificar_en_base(base, abril.periodos)
    assert abril.periodos[0].registrado is None
    assert abril.periodos[0].en_otros_periodos == {(3, 2026): [1]}


def test_psicomey_duplicado_entre_archivos_del_mismo_mes(crear_planilla, base):
    expse = [1, "05/03/2026 10:15:30", "Profesional", 500, "Trabajador", 900, "Examen", "EXPSE 1", "Sí"]
    psicomey = [2, "05/03/2026 10:15:30", "Profesional", 500, "Trabajador", 900, "Examen", "PSICOMEY", "Sí"]
    unico = [3, "06/03/2026 11:00:00", "Profesional", 501, "Otro", 901, "Examen", "PSICOMEY", "Sí"]
    rutas = [crear_planilla([expse], nombre="a.xlsx"), crear_planilla([psicomey, unico], nombre="b.xlsx")]
    analisis = analizar_archivos(rutas, MONTOS)
    assert analisis.periodos[0].liquidacion.psicomey_duplicados == 1
    importar(base, analisis.periodos)
    assert base_datos.buscar_liquidacion(base, 3, 2026).cantidad_ordenes == 2


def test_no_cumplidas_se_cuentan_una_vez(crear_planilla):
    ruta = crear_planilla([
        fila(1, "EXPSE 1", fecha="31/03/2026"), fila(2, "EXPSE 1", fecha="01/04/2026"),
        fila(3, "EXPSE 1", cumplida="No")])
    analisis = analizar_archivos([ruta], MONTOS)
    assert analisis.no_cumplidas == 1


def test_archivo_sin_orden(crear_planilla, base):
    encabezados = [e for e in ENCABEZADOS if e != "Orden"]
    ruta = crear_planilla([fila(1, "EXPSE 1")[1:], fila(2, "EXPSE 2")[1:]], encabezados=encabezados)
    analisis = analizar_archivos([ruta, ruta], MONTOS)
    [periodo] = analisis.periodos
    assert not periodo.liquidacion.tiene_numero_orden
    assert len(periodo.liquidacion.tabla) == 4
    importar(base, analisis.periodos)
    assert base_datos.buscar_liquidacion(base, 3, 2026).cantidad_ordenes == 4


def test_no_genera_excel(crear_planilla, tmp_path, base):
    ruta = crear_planilla([fila(1, "EXPSE 1")])
    antes = set(tmp_path.iterdir())
    importar(base, analizar_archivos([ruta], MONTOS).periodos)
    assert set(tmp_path.iterdir()) == antes


def _tres_meses(crear_planilla):
    expse = [1, "05/03/2026 10:15:30", "Profesional", 500, "Trabajador", 900, "Examen", "EXPSE 1", "Sí"]
    psicomey = [2, "05/03/2026 10:15:30", "Profesional", 500, "Trabajador", 900, "Examen", "PSICOMEY", "Sí"]
    ruta = crear_planilla([expse, psicomey, fila(3, "EXPSE 2", fecha="10/03/2026"),
                           fila(4, "EXPSE 1", fecha="01/04/2026"), fila(5, "EXPSE 2", fecha="01/05/2026")])
    return analizar_archivos([ruta], MONTOS).periodos


def test_genera_excel_de_los_meses_elegidos(crear_planilla, tmp_path, base):
    salida = tmp_path / "salida"
    salida.mkdir()
    periodos = _tres_meses(crear_planilla)
    resultado = importar(base, periodos, carpeta=salida, generar={(3, 2026), (5, 2026)})
    assert [r.name for r in resultado.excel_generados] == ["Liquidación Marzo 2026.xlsx", "Liquidación Mayo 2026.xlsx"]
    assert sorted(p.name for p in salida.iterdir()) == ["Liquidación Marzo 2026.xlsx", "Liquidación Mayo 2026.xlsx"]

    df = pd.read_excel(salida / "Liquidación Marzo 2026.xlsx")
    assert list(df["Prestación"].fillna("")) == ["EXPSE 1", "EXPSE 2", ""]
    assert list(df["Fecha"])[-1] == "TOTAL"
    assert df["Monto"].iloc[-1] == 300

    assert base_datos.buscar_liquidacion(base, 3, 2026).archivo_generado == str(salida / "Liquidación Marzo 2026.xlsx")
    assert base_datos.buscar_liquidacion(base, 4, 2026).archivo_generado is None


def test_genera_todos_y_no_pisa_existentes(crear_planilla, tmp_path, base):
    salida = tmp_path / "salida"
    salida.mkdir()
    (salida / "Liquidación Marzo 2026.xlsx").write_text("anterior")
    periodos = _tres_meses(crear_planilla)
    resultado = importar(base, periodos, carpeta=salida, generar={(p.mes, p.año) for p in periodos})
    assert [r.name for r in resultado.excel_generados] == [
        "Liquidación Marzo 2026 (1).xlsx", "Liquidación Abril 2026.xlsx", "Liquidación Mayo 2026.xlsx"]
    assert (salida / "Liquidación Marzo 2026.xlsx").read_text() == "anterior"


def test_mes_salteado_genera_excel_sin_tocar_el_historial(crear_planilla, tmp_path, base):
    salida = tmp_path / "salida"
    salida.mkdir()
    importar(base, analizar_archivos([crear_planilla([fila(1, "EXPSE 1")], nombre="viejo.xlsx")], MONTOS).periodos)
    periodos = _tres_meses(crear_planilla)
    resultado = importar(base, periodos, carpeta=salida, generar={(3, 2026)})
    assert resultado.salteados == ["Marzo 2026"]
    assert [r.name for r in resultado.excel_generados] == ["Liquidación Marzo 2026.xlsx"]
    registrada = base_datos.buscar_liquidacion(base, 3, 2026)
    assert (registrada.cantidad_ordenes, registrada.archivo_generado) == (1, None)


def test_error_al_guardar_no_frena_los_demas(crear_planilla, tmp_path, base, monkeypatch):
    salida = tmp_path / "salida"
    salida.mkdir()
    original = importacion.guardar_liquidacion

    def guardar(liquidacion, carpeta, mes, año):
        if mes == 4:
            raise PermissionError("el archivo está abierto")
        return original(liquidacion, carpeta, mes, año)

    monkeypatch.setattr(importacion, "guardar_liquidacion", guardar)
    periodos = _tres_meses(crear_planilla)
    resultado = importar(base, periodos, carpeta=salida, generar={(p.mes, p.año) for p in periodos})
    assert resultado.errores_excel == [("Abril 2026", "el archivo está abierto")]
    assert len(resultado.excel_generados) == 2
    assert len(resultado.registrados) == 3
    assert base_datos.buscar_liquidacion(base, 4, 2026).archivo_generado is None
