from pathlib import Path

import pytest
from openpyxl import Workbook

DIRECTORIO_DOCS = Path(__file__).resolve().parent.parent / "docs"

ENCABEZADOS = ["Orden", "Fecha", "Profesional", "HC", "Trabajador", "Suceso", "Tipo", "Prestación", "Cumplida"]


def fila(orden, prestacion, cumplida="Sí", fecha="01/03/2026"):
    """Fila de una orden con datos inventados."""
    return [orden, fecha, "Profesional Prueba", 1000 + orden, f"Trabajador {orden}", "Suceso", "Examen", prestacion, cumplida]


@pytest.fixture
def crear_planilla(tmp_path):
    """Arma un Excel con el formato de dgOrdenes: título en la fila 1 y encabezados en la fila 2."""
    def _crear(filas, encabezados=ENCABEZADOS, titulo="Examenes", nombre="planilla.xlsx"):
        wb = Workbook()
        ws = wb.active
        if titulo is not None:
            ws.append([titulo])
        ws.append(encabezados)
        for f in filas:
            ws.append(f)
        ruta = tmp_path / nombre
        wb.save(ruta)
        return ruta
    return _crear
