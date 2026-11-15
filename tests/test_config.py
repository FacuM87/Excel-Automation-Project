import json

from mis_liquidaciones.config import PRESTACIONES_POR_DEFECTO, cargar_montos, guardar_montos


def escribir_json(ruta, datos):
    ruta.write_text(json.dumps(datos), encoding="utf-8")


def test_sin_archivos_carga_prestaciones_por_defecto(tmp_path):
    montos = cargar_montos(tmp_path / "prestaciones.json", (tmp_path / "expse_montos.json",))
    assert list(montos) == PRESTACIONES_POR_DEFECTO
    assert all(v == "" for v in montos.values())


def test_lee_prestaciones_json(tmp_path):
    archivo = tmp_path / "prestaciones.json"
    escribir_json(archivo, {"expse 1": 100, "NUEVA": "250,5"})
    assert cargar_montos(archivo, ()) == {"EXPSE 1": "100", "NUEVA": "250,5"}


def test_migra_desde_expse_montos(tmp_path):
    anterior = tmp_path / "expse_montos.json"
    escribir_json(anterior, {"EXPSE 1": "100", "EXPSE 7": "700"})
    montos = cargar_montos(tmp_path / "prestaciones.json", (anterior,))
    assert montos["EXPSE 1"] == "100"
    assert montos["EXPSE 7"] == "700"
    assert set(montos) == set(PRESTACIONES_POR_DEFECTO)


def test_prestaciones_json_tiene_prioridad_sobre_el_anterior(tmp_path):
    archivo = tmp_path / "prestaciones.json"
    anterior = tmp_path / "expse_montos.json"
    escribir_json(archivo, {"EXPSE 1": "999"})
    escribir_json(anterior, {"EXPSE 1": "100"})
    assert cargar_montos(archivo, (anterior,)) == {"EXPSE 1": "999"}


def test_json_danado_carga_los_valores_por_defecto(tmp_path):
    archivo = tmp_path / "prestaciones.json"
    archivo.write_text("{no es json", encoding="utf-8")
    montos = cargar_montos(archivo, (tmp_path / "no_existe.json",))
    assert list(montos) == PRESTACIONES_POR_DEFECTO


def test_guardar_y_volver_a_cargar(tmp_path):
    archivo = tmp_path / "prestaciones.json"
    montos = {"EXPSE 1": "100", "CUES DES": "", "ATEPSIQ": "1000,5"}
    guardar_montos(montos, archivo)
    assert cargar_montos(archivo, ()) == montos
