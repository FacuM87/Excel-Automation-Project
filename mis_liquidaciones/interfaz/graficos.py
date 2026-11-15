import tkinter as tk
import tkinter.font as tkfont

import ttkbootstrap as tb

from ..procesamiento import formatear_monto

DURACION_MS, PASOS = 250, 12


def monto_abreviado(valor):
    """705368 -> '$ 705 mil'; 1234567 -> '$ 1,2 M'."""
    if valor >= 1_000_000:
        return f"$ {valor / 1_000_000:.1f} M".replace(".", ",")
    if valor >= 1_000:
        return f"$ {valor / 1_000:.0f} mil"
    return f"$ {formatear_monto(valor)}"


class _Grafico(tk.Canvas):
    """Canvas que se redibuja al cambiar de tamaño o de tema; las barras crecen al cambiar los datos."""

    def __init__(self, padre):
        super().__init__(padre, highlightthickness=0, bd=0, height=200)
        self.datos = []
        self.avance = 1.0
        self._animacion = None
        self.fuente = tkfont.nametofont("TkDefaultFont")
        self.fuente_chica = self.fuente.copy()
        self.fuente_chica.configure(size=max(self.fuente.cget("size") - 1, 8))
        self.bind("<Configure>", lambda _e: self.dibujar())
        self.bind("<<ThemeChanged>>", lambda _e: self.dibujar())

    def mostrar(self, datos):
        self.datos = list(datos)
        if self._animacion:
            self.after_cancel(self._animacion)
        self._animar(1)

    def _animar(self, paso):
        t = paso / PASOS
        self.avance = 1 - (1 - t) ** 3
        self.dibujar()
        self._animacion = self.after(DURACION_MS // PASOS, self._animar, paso + 1) if paso < PASOS else None

    def dibujar(self):
        colores = tb.Style().colors
        self.configure(bg=colores.bg)
        self.delete("all")
        ancho, alto = self.winfo_width(), self.winfo_height()
        if ancho < 50 or alto < 50:
            return
        if not any(d[1] for d in self.datos):
            self.create_text(ancho / 2, alto / 2, text="Sin datos", fill=colores.fg, font=self.fuente)
            return
        self._dibujar(colores, ancho, alto)


class GraficoBarras(_Grafico):
    """Barras verticales. datos: (etiqueta, total o None)."""

    def _dibujar(self, colores, ancho, alto):
        """Montos sobre cada barra; si no entran derechos van girados, siempre fuera de la barra."""
        margen, alto_texto = 10, self.fuente.metrics("linespace")
        hueco = (ancho - 2 * margen) / len(self.datos)
        textos = [None if total is None else monto_abreviado(total) for _e, total in self.datos]
        largo_maximo = max(self.fuente_chica.measure(t) for t in textos if t)
        girados = largo_maximo > hueco - 4
        base = alto - margen - alto_texto - 4
        techo = margen + (largo_maximo + 6 if girados else alto_texto + 4)
        maximo = max(d[1] or 0 for d in self.datos)
        ancho_barra = hueco * 0.62
        self.create_line(margen, base, ancho - margen, base, fill=colores.border)
        for i, ((etiqueta, total), texto) in enumerate(zip(self.datos, textos)):
            centro = margen + hueco * (i + 0.5)
            self.create_text(centro, base + 4, text=etiqueta, anchor="n", fill=colores.fg, font=self.fuente)
            if total is None:
                self.create_text(centro, base - 4, text="—", anchor="s", fill=colores.fg, font=self.fuente_chica)
                continue
            y = base - (base - techo) * (total / maximo if maximo else 0) * self.avance
            self.create_rectangle(centro - ancho_barra / 2, y, centro + ancho_barra / 2, base,
                                  fill=colores.info, outline="")
            if girados:
                self.create_text(centro, y - 4, text=texto, anchor="w", angle=90, fill=colores.fg,
                                 font=self.fuente_chica)
            else:
                self.create_text(centro, y - 3, text=texto, anchor="s", fill=colores.fg, font=self.fuente_chica)


class GraficoBarrasHorizontales(_Grafico):
    """Barras horizontales de mayor a menor. datos: (etiqueta, total, cantidad)."""

    def _dibujar(self, colores, ancho, alto):
        datos = sorted(self.datos, key=lambda d: d[1], reverse=True)
        margen = 10
        textos = [f"$ {formatear_monto(total)}  ({cantidad})" for _e, total, cantidad in datos]
        ancho_etiqueta = max(self.fuente.measure(d[0]) for d in datos) + 12
        ancho_texto = max(self.fuente_chica.measure(t) for t in textos) + 10
        izquierda = margen + ancho_etiqueta
        largo = max(ancho - izquierda - ancho_texto - margen, 10)
        fila = min((alto - 2 * margen) / len(datos), 34)
        maximo = max(d[1] for d in datos)
        self.create_line(izquierda, margen, izquierda, margen + fila * len(datos), fill=colores.border)
        for i, ((etiqueta, total, _c), texto) in enumerate(zip(datos, textos)):
            centro = margen + fila * (i + 0.5)
            grosor = fila * 0.6
            self.create_text(izquierda - 8, centro, text=etiqueta, anchor="e", fill=colores.fg, font=self.fuente)
            x = izquierda + largo * (total / maximo if maximo else 0) * self.avance
            self.create_rectangle(izquierda, centro - grosor / 2, max(x, izquierda + 1), centro + grosor / 2,
                                  fill=colores.info, outline="")
            self.create_text(x + 6, centro, text=texto, anchor="w", fill=colores.fg, font=self.fuente_chica)
