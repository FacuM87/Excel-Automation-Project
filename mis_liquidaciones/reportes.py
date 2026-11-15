from dataclasses import dataclass, field
from datetime import datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.label import DataLabelList
from openpyxl.descriptors import Typed
from openpyxl.styles import Border, Font, PatternFill, Side

from .procesamiento import MESES, PRESTACION_OPCIONAL

FORMATO_MONTO = '"$" #,##0.00'
FORMATO_ABREVIADO = '"$" #,##0, "mil"'
COLOR_BARRAS = "3498DB"


@dataclass
class ReporteMensual:
    mes: int
    año: int
    por_prestacion: list = field(default_factory=list)
    cantidad: int = 0
    total: float = 0.0
    excluidas_opcional: int = 0

    @property
    def nombre(self):
        return f"{MESES[self.mes - 1]} {self.año}"


@dataclass
class ReporteAnual:
    """por_mes tiene los 12 meses: (mes, órdenes, total), con None en los meses sin liquidación."""
    año: int
    por_mes: list = field(default_factory=list)
    por_prestacion: list = field(default_factory=list)
    cantidad: int = 0
    total: float = 0.0
    excluidas_opcional: int = 0


def años_registrados(con):
    return [a for (a,) in con.execute("SELECT DISTINCT año FROM liquidaciones ORDER BY año DESC")]


def _por_prestacion(con, condicion, parametros):
    return [(p, c, t) for p, c, t in con.execute(
        f"SELECT o.prestacion, COUNT(*), SUM(o.monto) FROM ordenes o JOIN liquidaciones l ON l.id = o.liquidacion_id "
        f"WHERE {condicion} AND o.prestacion <> ? GROUP BY o.prestacion ORDER BY o.prestacion",
        (*parametros, PRESTACION_OPCIONAL))]


def _excluidas(con, condicion, parametros):
    return con.execute(
        f"SELECT COUNT(*) FROM ordenes o JOIN liquidaciones l ON l.id = o.liquidacion_id "
        f"WHERE {condicion} AND o.prestacion = ?", (*parametros, PRESTACION_OPCIONAL)).fetchone()[0]


def reporte_mensual(con, mes, año):
    """Cantidad y total por prestación del período, sin PSICOMEY (R2). None si no está registrado."""
    if not con.execute("SELECT 1 FROM liquidaciones WHERE mes = ? AND año = ?", (mes, año)).fetchone():
        return None
    condicion, parametros = "l.mes = ? AND l.año = ?", (mes, año)
    filas = _por_prestacion(con, condicion, parametros)
    return ReporteMensual(mes, año, filas, sum(f[1] for f in filas), sum(f[2] for f in filas),
                          _excluidas(con, condicion, parametros))


def reporte_anual(con, año):
    """Totales por mes y por prestación del año, sin PSICOMEY (R2)."""
    registrados = {mes: (0, 0.0) for (mes,) in con.execute("SELECT mes FROM liquidaciones WHERE año = ?", (año,))}
    for mes, cantidad, total in con.execute(
            "SELECT l.mes, COUNT(*), SUM(o.monto) FROM ordenes o JOIN liquidaciones l ON l.id = o.liquidacion_id "
            "WHERE l.año = ? AND o.prestacion <> ? GROUP BY l.mes", (año, PRESTACION_OPCIONAL)):
        registrados[mes] = (cantidad, total)
    por_mes = [(mes, *registrados.get(mes, (None, None))) for mes in range(1, 13)]
    filas = _por_prestacion(con, "l.año = ?", (año,))
    return ReporteAnual(año, por_mes, filas, sum(f[1] for f in filas), sum(f[2] for f in filas),
                        _excluidas(con, "l.año = ?", (año,)))


def nota_opcional(reporte):
    if not reporte.excluidas_opcional:
        return ""
    return f"No incluye {reporte.excluidas_opcional} orden(es) de {PRESTACION_OPCIONAL}."


# --- Exportación a Excel --- #


class _Etiquetas(DataLabelList):
    """openpyxl escribe el formato de las etiquetas sin sourceLinked=0 y Excel lo ignora."""
    numFmt = Typed(expected_type=NumFmt, allow_none=True)
    __attrs__ = DataLabelList.__attrs__
    __nested__ = tuple(n for n in DataLabelList.__nested__ if n != "numFmt")
    __elements__ = DataLabelList.__elements__


_NEGRITA = Font(bold=True)
_RELLENO_ENCABEZADO = PatternFill("solid", fgColor="DCE6F1")
_LINEA = Side(style="thin", color="808080")


def _escribir_tabla(hoja, fila, subtitulo, encabezados, filas, fila_total):
    """Escribe subtítulo, encabezados, filas y TOTAL desde la fila indicada. Devuelve (fila de encabezados, siguiente libre)."""
    if subtitulo:
        hoja.cell(fila, 1, subtitulo).font = Font(bold=True, size=12)
        fila += 1
    encabezado = fila
    for col, texto in enumerate(encabezados, 1):
        celda = hoja.cell(fila, col, texto)
        celda.font = _NEGRITA
        celda.fill = _RELLENO_ENCABEZADO
        celda.border = Border(bottom=_LINEA)
    for valores in [*filas, fila_total]:
        fila += 1
        for col, valor in enumerate(valores, 1):
            celda = hoja.cell(fila, col, valor)
            if col == len(valores):
                celda.number_format = FORMATO_MONTO
    for col in range(1, len(fila_total) + 1):
        celda = hoja.cell(fila, col)
        celda.font = _NEGRITA
        celda.border = Border(top=_LINEA)
    return encabezado, fila + 2


def _nuevo_libro(titulo, reporte):
    libro = Workbook()
    hoja = libro.active
    hoja.title = "Reporte"
    hoja.cell(1, 1, titulo).font = Font(bold=True, size=14)
    hoja.cell(2, 1, f"Generado el {datetime.now():%d/%m/%Y}")
    if nota_opcional(reporte):
        hoja.cell(3, 1, nota_opcional(reporte)).font = Font(italic=True)
    for letra, ancho in (("A", 22), ("B", 12), ("C", 18)):
        hoja.column_dimensions[letra].width = ancho
    return libro, hoja


def _grafico(hoja, encabezado, cantidad_filas, titulo, horizontal, alto):
    """Gráfico de barras nativo de Excel con los totales (columna C) de la tabla, ubicado en la columna E."""
    if not cantidad_filas:
        return
    grafico = BarChart()
    grafico.type = "bar" if horizontal else "col"
    grafico.title = titulo
    grafico.title.overlay = False
    grafico.legend = None
    grafico.height, grafico.width = alto, 16 if horizontal else 22
    grafico.gapWidth = 60
    grafico.add_data(Reference(hoja, min_col=3, min_row=encabezado, max_row=encabezado + cantidad_filas),
                     titles_from_data=True)
    grafico.set_categories(Reference(hoja, min_col=1, min_row=encabezado + 1, max_row=encabezado + cantidad_filas))
    serie = grafico.series[0]
    serie.graphicalProperties.solidFill = COLOR_BARRAS
    serie.graphicalProperties.line.noFill = True
    formato = '"$" #,##0' if horizontal else FORMATO_ABREVIADO
    grafico.dataLabels = _Etiquetas(numFmt=NumFmt(formatCode=formato, sourceLinked=False), showVal=True,
                                    showSerName=False, showCatName=False, showLegendKey=False,
                                    showPercent=False, showLeaderLines=False)
    grafico.y_axis.majorGridlines = None
    grafico.x_axis.delete = False
    grafico.y_axis.delete = True
    if horizontal:
        grafico.x_axis.scaling.orientation = "maxMin"
    hoja.add_chart(grafico, f"E{encabezado}")


def exportar_mensual(reporte, ruta):
    libro, hoja = _nuevo_libro(f"Reporte mensual — {reporte.nombre}", reporte)
    encabezado, _ = _escribir_tabla(hoja, 5, None, ("Prestación", "Cantidad", "Total"), reporte.por_prestacion,
                                    ("TOTAL", reporte.cantidad, reporte.total))
    filas = len(reporte.por_prestacion)
    _grafico(hoja, encabezado, filas, f"Total por prestación — {reporte.nombre}", True, max(6, 1 + filas * 0.9))
    libro.save(ruta)


def exportar_anual(reporte, ruta):
    libro, hoja = _nuevo_libro(f"Reporte anual {reporte.año}", reporte)
    filas_mes = [(MESES[mes - 1], cantidad, total) for mes, cantidad, total in reporte.por_mes]
    encabezado_mes, fila = _escribir_tabla(hoja, 5, "Por mes", ("Mes", "Órdenes", "Total"), filas_mes,
                                           ("TOTAL", reporte.cantidad, reporte.total))
    encabezado_prestacion, _ = _escribir_tabla(
        hoja, fila, "Por prestación", ("Prestación", "Cantidad", "Total"), reporte.por_prestacion,
        ("TOTAL", reporte.cantidad, reporte.total))
    _grafico(hoja, encabezado_mes, 12, f"Total por mes — {reporte.año}", False, 7)
    filas = len(reporte.por_prestacion)
    _grafico(hoja, encabezado_prestacion, filas, f"Total por prestación — {reporte.año}", True,
             max(6, 1 + filas * 0.9))
    libro.save(ruta)
