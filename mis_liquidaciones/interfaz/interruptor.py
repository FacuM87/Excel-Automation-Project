import tkinter as tk

import ttkbootstrap as tb

ANCHO, ALTO, MARGEN = 36, 20, 3
DURACION_MS, PASOS = 140, 10


def _mezclar(color_a, color_b, t):
    a = [int(color_a[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(color_b[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02x}" for x, y in zip(a, b))


class Interruptor(tb.Frame):
    """Interruptor con texto que se desliza con una animación corta. Usa una BooleanVar."""

    def __init__(self, padre, text, variable, command=None):
        super().__init__(padre)
        self.variable = variable
        self.command = command
        self.posicion = 1.0 if variable.get() else 0.0
        self._animacion = None

        self.canvas = tk.Canvas(self, width=ANCHO, height=ALTO, highlightthickness=0, bd=0, cursor="hand2",
                                takefocus=True)
        self.canvas.pack(side=tk.LEFT)
        self.etiqueta = tb.Label(self, text=text, cursor="hand2")
        self.etiqueta.pack(side=tk.LEFT, padx=(6, 0))

        for widget in (self.canvas, self.etiqueta):
            widget.bind("<Button-1>", lambda _e: self.alternar())
        self.canvas.bind("<space>", lambda _e: self.alternar())
        self.canvas.bind("<<ThemeChanged>>", lambda _e: self.dibujar())
        self._traza = variable.trace_add("write", lambda *_: self._sincronizar())
        self.bind("<Destroy>", self._al_destruir, add="+")
        self.dibujar()

    def _al_destruir(self, evento):
        if evento.widget is self:
            self.variable.trace_remove("write", self._traza)

    def alternar(self):
        self.canvas.focus_set()
        self.variable.set(not self.variable.get())
        if self.command:
            self.command()

    def _sincronizar(self):
        destino = 1.0 if self.variable.get() else 0.0
        if self._animacion:
            self.after_cancel(self._animacion)
        self._animar(self.posicion, destino, 1)

    def _animar(self, inicio, destino, paso):
        t = paso / PASOS
        suavizado = 1 - (1 - t) ** 3
        self.posicion = inicio + (destino - inicio) * suavizado
        self.dibujar()
        if paso < PASOS:
            self._animacion = self.after(DURACION_MS // PASOS, self._animar, inicio, destino, paso + 1)
        else:
            self._animacion = None

    def dibujar(self):
        colores = tb.Style().colors
        self.canvas.configure(bg=colores.bg)
        self.canvas.delete("all")
        radio = ALTO / 2
        pista = _mezclar(colores.secondary, colores.primary, self.posicion)
        self.canvas.create_line(radio, radio, ANCHO - radio, radio, width=ALTO, capstyle=tk.ROUND, fill=pista)
        x = radio + (ANCHO - ALTO) * self.posicion
        r = radio - MARGEN
        self.canvas.create_oval(x - r, radio - r, x + r, radio + r, fill="#ffffff", outline="")
