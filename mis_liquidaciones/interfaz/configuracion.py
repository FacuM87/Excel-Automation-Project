import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as tb

from .. import config
from ..procesamiento import normalizar_prestacion, parsear_monto


class PestanaConfiguracion:
    def __init__(self, padre, app):
        self.app = app
        self.entries = {}
        self.marco = tb.Frame(padre, padding=15)

        frame_montos = tb.Labelframe(self.marco, text="Montos por prestación", padding=10)
        frame_montos.pack(fill="x")

        self.frame_lista = tb.Frame(frame_montos)
        self.frame_lista.pack(fill="x")
        self.dibujar_prestaciones()

        tb.Separator(frame_montos).pack(fill="x", pady=10)
        frame_nueva = tb.Frame(frame_montos)
        frame_nueva.pack(anchor="w")
        tb.Label(frame_nueva, text="Nueva prestación:").pack(side=tk.LEFT, padx=(0, 5))
        self.ent_nueva_nombre = tb.Entry(frame_nueva, width=14)
        self.ent_nueva_nombre.pack(side=tk.LEFT)
        tb.Label(frame_nueva, text="Monto:").pack(side=tk.LEFT, padx=(15, 5))
        self.ent_nueva_monto = tb.Entry(frame_nueva, width=12)
        self.ent_nueva_monto.pack(side=tk.LEFT)
        tb.Button(frame_nueva, text="Agregar", bootstyle="success-outline",
                  command=self.agregar_prestacion).pack(side=tk.LEFT, padx=10)

        frame_carpeta = tb.Labelframe(self.marco, text="Carpeta donde se guardan las liquidaciones", padding=10)
        frame_carpeta.pack(fill="x", pady=15)
        self.lbl_carpeta = tb.Label(frame_carpeta, text=str(app.carpeta_guardado))
        self.lbl_carpeta.pack(side=tk.LEFT, fill="x", expand=True)
        tb.Button(frame_carpeta, text="Cambiar…", bootstyle="info-outline",
                  command=self.elegir_carpeta).pack(side=tk.RIGHT)

        frame_tema = tb.Labelframe(self.marco, text="Tema", padding=10)
        frame_tema.pack(fill="x")
        self.var_tema = tk.StringVar(value=app.ajustes["tema"])
        for tema in config.TEMAS:
            tb.Radiobutton(frame_tema, text=tema.capitalize(), value=tema, variable=self.var_tema,
                           command=lambda: app.cambiar_tema(self.var_tema.get())).pack(side=tk.LEFT, padx=(0, 20))

    # --- Prestaciones --- #

    def dibujar_prestaciones(self):
        for widget in self.frame_lista.winfo_children():
            widget.destroy()
        self.entries.clear()
        for i, prestacion in enumerate(self.app.montos_prestaciones):
            fila, col = divmod(i, 2)
            col *= 3
            tb.Label(self.frame_lista, text=prestacion, width=12, anchor="e").grid(
                row=fila, column=col, padx=(10, 5), pady=3)
            entry = tb.Entry(self.frame_lista, width=12)
            entry.insert(0, self.app.montos_prestaciones[prestacion])
            entry.grid(row=fila, column=col + 1, pady=3)
            entry.bind("<KeyRelease>", lambda _e: self._montos_cambiados())
            entry.bind("<FocusOut>", lambda _e: self.guardar_montos(avisar=False))
            tb.Button(self.frame_lista, text="✕", bootstyle="danger-link", width=2,
                      command=lambda p=prestacion: self.quitar_prestacion(p)).grid(
                row=fila, column=col + 2, padx=(2, 25), pady=3)
            self.entries[prestacion] = entry
        self._montos_cambiados()

    def _montos_cambiados(self):
        liquidar = getattr(self.app, "liquidar", None)
        if liquidar:
            liquidar.actualizar_aviso_montos()

    def sincronizar_montos(self):
        for prestacion, entry in self.entries.items():
            self.app.montos_prestaciones[prestacion] = entry.get()

    def guardar_montos(self, avisar=True):
        self.sincronizar_montos()
        try:
            config.guardar_montos(self.app.montos_prestaciones, self.app.archivo_montos)
        except OSError as e:
            if avisar:
                messagebox.showwarning("Aviso", f"No se pudieron guardar los montos en {self.app.archivo_montos}\n{e}")

    def agregar_prestacion(self):
        nombre = normalizar_prestacion(self.ent_nueva_nombre.get())
        monto = self.ent_nueva_monto.get().strip()
        if not nombre:
            messagebox.showerror("Prestación inválida", "Ingresá el nombre de la prestación.")
            return
        if nombre in self.entries:
            messagebox.showerror("Prestación repetida", f"{nombre} ya está en la lista.")
            return
        if monto and parsear_monto(monto) is None:
            messagebox.showerror("Monto inválido", f"\"{monto}\" no es un número válido.")
            return
        self.sincronizar_montos()
        self.app.montos_prestaciones[nombre] = monto
        self.dibujar_prestaciones()
        self.guardar_montos()
        self.ent_nueva_nombre.delete(0, tk.END)
        self.ent_nueva_monto.delete(0, tk.END)

    def quitar_prestacion(self, prestacion):
        if not messagebox.askyesno("Quitar prestación", f"¿Quitar {prestacion} de la lista?"):
            return
        self.sincronizar_montos()
        del self.app.montos_prestaciones[prestacion]
        self.dibujar_prestaciones()
        self.guardar_montos()

    def leer_y_validar_montos(self):
        """Lee los montos de la interfaz. Devuelve (montos_numericos, errores). Los vacíos se omiten."""
        montos = {}
        errores = []
        for prestacion, entry in self.entries.items():
            if not entry.get().strip():
                continue
            valor = parsear_monto(entry.get())
            if valor is None:
                errores.append(f"{prestacion}: \"{entry.get()}\" no es un número válido")
            else:
                montos[prestacion] = valor
        return montos, errores

    # --- Carpeta --- #

    def elegir_carpeta(self, carpeta=None):
        if carpeta is None:
            carpeta = filedialog.askdirectory(title="Elegí dónde guardar las liquidaciones",
                                              initialdir=self.app.carpeta_guardado)
        if carpeta:
            self.app.cambiar_carpeta(carpeta)
            self.lbl_carpeta.config(text=str(self.app.carpeta_guardado))
