import os
import tkinter as tk
import tkinter.font as tkfont
from contextlib import closing
from tkinter import filedialog, messagebox

import ttkbootstrap as tb

from .. import base_datos, reportes
from ..procesamiento import MESES, formatear_monto
from .graficos import GraficoBarras, GraficoBarrasHorizontales
from .utiles import crear_tabla

COLUMNAS_PRESTACION = (("prestacion", "Prestación", 160, "w"), ("cantidad", "Cantidad", 80, "e"),
                       ("total", "Total", 130, "e"))


def _monto(valor):
    return "—" if valor is None else f"$ {formatear_monto(valor)}"


class PestanaReportes:
    def __init__(self, padre, app):
        self.app = app
        self.periodos = {}
        self.reporte = None

        self.marco = tb.Frame(padre, padding=15)
        self.fuente_total = tkfont.nametofont("TkDefaultFont").copy()
        self.fuente_total.configure(weight="bold")

        barra = tb.Frame(self.marco)
        barra.pack(fill="x", pady=(0, 10))
        self.var_tipo = tk.StringVar(value="mensual")
        for valor, texto in (("mensual", "Mensual"), ("anual", "Anual")):
            tb.Radiobutton(barra, text=texto, value=valor, variable=self.var_tipo, bootstyle="info-outline-toolbutton",
                           cursor="hand2", width=9, command=self.actualizar).pack(side=tk.LEFT)
        self.var_vista = tk.StringVar(value="tabla")
        marco_vista = tb.Frame(barra)
        marco_vista.pack(side=tk.LEFT, padx=(15, 0))
        for valor, texto in (("tabla", "Tabla"), ("grafico", "Gráfico")):
            tb.Radiobutton(marco_vista, text=texto, value=valor, variable=self.var_vista,
                           bootstyle="info-outline-toolbutton", cursor="hand2", width=9,
                           command=self.mostrar).pack(side=tk.LEFT)
        self.cmb_periodo = tb.Combobox(barra, state="readonly", width=16)
        self.cmb_periodo.pack(side=tk.LEFT, padx=15)
        self.cmb_periodo.bind("<<ComboboxSelected>>", lambda _e: self.mostrar())
        self.btn_exportar = tb.Button(barra, text="Exportar a Excel…", bootstyle="primary", command=self.exportar,
                                      state=tk.DISABLED)
        self.btn_exportar.pack(side=tk.RIGHT)

        self.lbl_estado = tb.Label(self.marco, text="")
        self.lbl_estado.pack(anchor="w", pady=(0, 10))

        self.contenido = tb.Frame(self.marco)
        self.contenido.pack(fill="both", expand=True)

        self.vista_mensual = tb.Labelframe(self.contenido, text="Por prestación", padding=5)
        marco, self.tabla_mensual = crear_tabla(self.vista_mensual, COLUMNAS_PRESTACION, altura=12)
        marco.pack(fill="both", expand=True)

        self.vista_anual = tb.Panedwindow(self.contenido, orient=tk.HORIZONTAL)
        izquierda = tb.Labelframe(self.vista_anual, text="Por mes", padding=5)
        self.vista_anual.add(izquierda, weight=1)
        marco, self.tabla_meses = crear_tabla(izquierda, (
            ("mes", "Mes", 120, "w"), ("ordenes", "Órdenes", 80, "e"), ("total", "Total", 130, "e")), altura=14)
        marco.pack(fill="both", expand=True)
        derecha = tb.Labelframe(self.vista_anual, text="Por prestación", padding=5)
        self.vista_anual.add(derecha, weight=1)
        marco, self.tabla_anual = crear_tabla(derecha, COLUMNAS_PRESTACION, altura=14)
        marco.pack(fill="both", expand=True)

        for tabla in (self.tabla_mensual, self.tabla_meses, self.tabla_anual):
            tabla.tag_configure("total", font=self.fuente_total)

        self.grafico_mensual_marco = tb.Labelframe(self.contenido, text="Total por prestación", padding=5)
        self.grafico_mensual = GraficoBarrasHorizontales(self.grafico_mensual_marco)
        self.grafico_mensual.pack(fill="both", expand=True)

        self.grafico_anual_divisor = tb.Panedwindow(self.contenido, orient=tk.HORIZONTAL)
        izquierda = tb.Labelframe(self.grafico_anual_divisor, text="Total por mes", padding=5)
        self.grafico_anual_divisor.add(izquierda, weight=3)
        self.grafico_meses = GraficoBarras(izquierda)
        self.grafico_meses.pack(fill="both", expand=True)
        derecha = tb.Labelframe(self.grafico_anual_divisor, text="Por prestación", padding=5)
        self.grafico_anual_divisor.add(derecha, weight=2)
        self.grafico_anual = GraficoBarrasHorizontales(derecha)
        self.grafico_anual.pack(fill="both", expand=True)

        self.vistas = {("mensual", "tabla"): self.vista_mensual, ("anual", "tabla"): self.vista_anual,
                       ("mensual", "grafico"): self.grafico_mensual_marco,
                       ("anual", "grafico"): self.grafico_anual_divisor}

        self.actualizar()

    def _abrir_base(self):
        return closing(base_datos.abrir(self.app.archivo_base))

    def actualizar(self):
        """Recarga los períodos o años registrados y muestra el reporte elegido (si sigue existiendo)."""
        anterior = self.cmb_periodo.get()
        try:
            with self._abrir_base() as con:
                if self.var_tipo.get() == "mensual":
                    registradas = reversed(base_datos.listar_liquidaciones(con))
                    self.periodos = {r.nombre: (r.mes, r.año) for r in registradas}
                else:
                    self.periodos = {str(a): a for a in reportes.años_registrados(con)}
        except Exception as e:
            self.periodos = {}
            self.lbl_estado.config(text=f"No se pudo abrir el historial: {e}")
            return self.mostrar(error=True)

        nombres = list(self.periodos)
        self.cmb_periodo.config(values=nombres)
        self.cmb_periodo.set(anterior if anterior in self.periodos else (nombres[0] if nombres else ""))
        self.mostrar()

    def mostrar(self, error=False):
        mensual = self.var_tipo.get() == "mensual"
        elegida = self.vistas[(self.var_tipo.get(), self.var_vista.get())]
        for vista in self.vistas.values():
            if vista is not elegida:
                vista.pack_forget()
        elegida.pack(fill="both", expand=True)
        for tabla in (self.tabla_mensual, self.tabla_meses, self.tabla_anual):
            tabla.delete(*tabla.get_children())

        clave = self.periodos.get(self.cmb_periodo.get())
        self.reporte = None
        if clave is not None:
            with self._abrir_base() as con:
                self.reporte = reportes.reporte_mensual(con, *clave) if mensual else reportes.reporte_anual(con, clave)
        self.btn_exportar.config(state=tk.NORMAL if self.reporte else tk.DISABLED)
        if not self.reporte:
            for grafico in (self.grafico_mensual, self.grafico_meses, self.grafico_anual):
                grafico.mostrar([])
            if not error:
                self.lbl_estado.config(text="Todavía no hay liquidaciones registradas.")
            return

        r = self.reporte
        resumen = f"{r.cantidad} orden(es) · total $ {formatear_monto(r.total)}"
        self.lbl_estado.config(text="   ".join(t for t in (resumen, reportes.nota_opcional(r)) if t))
        if mensual:
            self._llenar_prestaciones(self.tabla_mensual, r)
            self.grafico_mensual.mostrar((p or "(vacío)", t, c) for p, c, t in r.por_prestacion)
            return
        for mes, cantidad, total in r.por_mes:
            self.tabla_meses.insert("", tk.END, values=(MESES[mes - 1], "—" if cantidad is None else cantidad,
                                                        _monto(total)))
        self.tabla_meses.insert("", tk.END, values=("TOTAL", r.cantidad, _monto(r.total)), tags=("total",))
        self._llenar_prestaciones(self.tabla_anual, r)
        self.grafico_meses.mostrar((MESES[mes - 1][:3], total) for mes, _c, total in r.por_mes)
        self.grafico_anual.mostrar((p or "(vacío)", t, c) for p, c, t in r.por_prestacion)

    def _llenar_prestaciones(self, tabla, reporte):
        for prestacion, cantidad, total in reporte.por_prestacion:
            tabla.insert("", tk.END, values=(prestacion or "(vacío)", cantidad, _monto(total)))
        tabla.insert("", tk.END, values=("TOTAL", reporte.cantidad, _monto(reporte.total)), tags=("total",))

    # --- Exportar --- #

    def exportar(self, ruta=None):
        r = self.reporte
        if not r:
            return
        mensual = isinstance(r, reportes.ReporteMensual)
        if ruta is None:
            ruta = filedialog.asksaveasfilename(
                title="Guardar reporte", initialdir=self.app.carpeta_guardado, defaultextension=".xlsx",
                initialfile=f"Reporte {r.nombre}.xlsx" if mensual else f"Reporte anual {r.año}.xlsx",
                filetypes=[("Excel Files", "*.xlsx")])
        if not ruta:
            return
        try:
            (reportes.exportar_mensual if mensual else reportes.exportar_anual)(r, ruta)
        except PermissionError:
            messagebox.showerror("Error", f"No se pudo guardar el reporte:\n{ruta}\n\n"
                                          f"Si el archivo está abierto en Excel, cerralo y probá de nuevo.")
            return
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el reporte.\n{e}")
            return
        if messagebox.askyesno("Reporte exportado", f"Se guardó el reporte en:\n{ruta}\n\n¿Querés abrirlo?"):
            os.startfile(ruta)
