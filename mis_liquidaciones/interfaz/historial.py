import os
import tkinter as tk
import tkinter.font as tkfont
from contextlib import closing
from tkinter import filedialog, messagebox

import pandas as pd
import ttkbootstrap as tb

from .. import base_datos, importacion
from ..procesamiento import formatear_monto
from .dialogo_importacion import DialogoImportacion
from .utiles import crear_tabla


class PestanaHistorial:
    def __init__(self, padre, app):
        self.app = app
        self.liquidaciones = {}
        self.dialogo_importacion = None

        self.marco = tb.Frame(padre, padding=15)
        self.fuente_total = tkfont.nametofont("TkDefaultFont").copy()
        self.fuente_total.configure(weight="bold")

        barra = tb.Frame(self.marco)
        barra.pack(fill="x", pady=(0, 10))
        tb.Button(barra, text="Importar historial…", bootstyle="primary",
                  command=self.importar_historial).pack(side=tk.LEFT)
        self.btn_borrar = tb.Button(barra, text="Borrar", bootstyle="danger-outline", command=self.borrar,
                                    state=tk.DISABLED)
        self.btn_borrar.pack(side=tk.RIGHT)
        self.btn_abrir_excel = tb.Button(barra, text="Abrir Excel generado", bootstyle="info-outline",
                                         command=self.abrir_excel, state=tk.DISABLED)
        self.btn_abrir_excel.pack(side=tk.RIGHT, padx=10)
        self.lbl_estado = tb.Label(barra, text="")
        self.lbl_estado.pack(side=tk.LEFT, padx=15)

        marco_lista, self.tabla = crear_tabla(self.marco, (
            ("periodo", "Período", 130, "w"), ("ordenes", "Órdenes", 70, "e"), ("total", "Total", 110, "e"),
            ("registrada", "Registrada", 120, "center"), ("origen", "Origen", 260, "w"),
            ("opcional", "PSICOMEY", 80, "center")), altura=7)
        marco_lista.pack(fill="both", expand=True)
        self.tabla.bind("<<TreeviewSelect>>", lambda _e: self.mostrar_detalle())

        self.divisor = tb.Panedwindow(self.marco, orient=tk.HORIZONTAL)
        self.divisor.pack(fill="both", expand=True, pady=(15, 0))

        izquierda = tb.Labelframe(self.divisor, text="Órdenes", padding=5)
        self.divisor.add(izquierda, weight=3)
        marco_ordenes, self.tabla_ordenes = crear_tabla(izquierda, (
            ("orden", "Orden", 90, "e"), ("fecha", "Fecha", 85, "center"), ("trabajador", "Trabajador", 200, "w"),
            ("prestacion", "Prestación", 90, "w"), ("monto", "Monto", 80, "e")))
        marco_ordenes.pack(fill="both", expand=True)

        derecha = tb.Labelframe(self.divisor, text="Por prestación", padding=5)
        self.divisor.add(derecha, weight=1)
        marco_resumen, self.tabla_resumen = crear_tabla(derecha, (
            ("prestacion", "Prestación", 90, "w"), ("cantidad", "Cant.", 50, "e"), ("total", "Total", 100, "e")))
        marco_resumen.pack(fill="both", expand=True)

        self.actualizar()

    def _abrir_base(self):
        return closing(base_datos.abrir(self.app.archivo_base))

    def seleccionada(self):
        seleccion = self.tabla.selection()
        return self.liquidaciones.get(int(seleccion[0])) if seleccion else None

    def actualizar(self, seleccionar=None):
        """Recarga la lista desde la base. Mantiene la selección si sigue existiendo."""
        anterior = self.seleccionada()
        seleccionar = seleccionar or (anterior and (anterior.mes, anterior.año))
        try:
            with self._abrir_base() as con:
                registradas = base_datos.listar_liquidaciones(con)
        except Exception as e:
            registradas = []
            self.lbl_estado.config(text=f"No se pudo abrir el historial: {e}")
        else:
            total = sum(r.total for r in registradas)
            self.lbl_estado.config(
                text=f"{len(registradas)} liquidación(es) · total $ {formatear_monto(total)}" if registradas
                else "Todavía no hay liquidaciones registradas.")

        self.tabla.delete(*self.tabla.get_children())
        self.liquidaciones = {r.id: r for r in registradas}
        for r in reversed(registradas):
            self.tabla.insert("", tk.END, iid=str(r.id), values=(
                r.nombre, r.cantidad_ordenes, f"$ {formatear_monto(r.total)}", f"{r.generada:%d/%m/%Y}",
                r.archivo_origen or "", "Sí" if r.incluye_opcional else "No"))
            if seleccionar == (r.mes, r.año):
                self.tabla.selection_set(str(r.id))
                self.tabla.see(str(r.id))
        self.mostrar_detalle()

    def mostrar_detalle(self):
        self.tabla_ordenes.delete(*self.tabla_ordenes.get_children())
        self.tabla_resumen.delete(*self.tabla_resumen.get_children())
        liquidacion = self.seleccionada()
        estado = tk.NORMAL if liquidacion else tk.DISABLED
        self.btn_borrar.config(state=estado)
        self.btn_abrir_excel.config(state=tk.NORMAL if liquidacion and liquidacion.archivo_generado else tk.DISABLED)
        if not liquidacion:
            return

        with self._abrir_base() as con:
            ordenes = base_datos.leer_ordenes(con, liquidacion.id)
            resumen = base_datos.resumen_por_prestacion(con, liquidacion.id)
        fechas = pd.to_datetime(ordenes["fecha"], errors="coerce")
        for fila, fecha in zip(ordenes.itertuples(), fechas):
            self.tabla_ordenes.insert("", tk.END, values=(
                "" if fila.orden is None or pd.isna(fila.orden) else fila.orden,
                "" if pd.isna(fecha) else f"{fecha:%d/%m/%Y}",
                fila.trabajador or "", fila.prestacion, f"$ {formatear_monto(fila.monto)}"))
        for prestacion, cantidad, total in resumen:
            self.tabla_resumen.insert("", tk.END, values=(prestacion or "(vacío)", cantidad,
                                                          f"$ {formatear_monto(total)}"))
        self.tabla_resumen.insert("", tk.END, values=(
            "TOTAL", liquidacion.cantidad_ordenes, f"$ {formatear_monto(liquidacion.total)}"), tags=("total",))
        self.tabla_resumen.tag_configure("total", font=self.fuente_total)

    # --- Acciones --- #

    def abrir_excel(self):
        liquidacion = self.seleccionada()
        if not liquidacion:
            return
        if not liquidacion.archivo_generado:
            messagebox.showinfo("Sin Excel", f"{liquidacion.nombre} se importó del historial; no tiene Excel generado.")
        elif not os.path.exists(liquidacion.archivo_generado):
            messagebox.showwarning(
                "Archivo no encontrado",
                f"No se encontró el Excel de {liquidacion.nombre}:\n{liquidacion.archivo_generado}\n\n"
                f"Puede haberse movido o borrado.")
        else:
            os.startfile(liquidacion.archivo_generado)

    def borrar(self):
        liquidacion = self.seleccionada()
        if not liquidacion:
            return
        if not messagebox.askyesno(
                "Borrar liquidación",
                f"¿Borrar la liquidación de {liquidacion.nombre} del historial?\n\n"
                f"({liquidacion.cantidad_ordenes} órdenes, total $ {formatear_monto(liquidacion.total)})\n\n"
                f"El Excel generado no se borra.", icon="warning"):
            return
        try:
            with self._abrir_base() as con:
                base_datos.borrar(con, liquidacion.id)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo borrar la liquidación.\n{e}")
            return
        self.tabla.selection_remove(self.tabla.selection())
        self.actualizar()

    def importar_historial(self, rutas=None):
        """Carga en la base los archivos ya recibidos, sin generar Excel."""
        if rutas is None:
            rutas = filedialog.askopenfilenames(
                title="Elegí las planillas a importar", filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not rutas:
            return

        configuracion = self.app.configuracion
        montos, errores = configuracion.leer_y_validar_montos()
        if errores:
            messagebox.showerror("Montos inválidos",
                                 "Corregí los siguientes montos en Configuración:\n\n" + "\n".join(errores))
            return
        configuracion.guardar_montos()
        incluir_opcional = self.app.liquidar.var_incluir_opcional.get()

        analisis = importacion.analizar_archivos(rutas, montos, incluir_opcional)
        if not analisis.periodos:
            detalle = "\n".join(f"  • {nombre}: {error}" for nombre, error in analisis.errores)
            messagebox.showerror("Nada para importar", "No se encontraron órdenes para importar.\n\n" + detalle)
            return

        try:
            with self._abrir_base() as con:
                importacion.verificar_en_base(con, analisis.periodos)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el historial ({self.app.archivo_base}).\n{e}")
            return

        self.dialogo_importacion = DialogoImportacion(
            self.app.root, analisis, incluir_opcional, self._confirmar_importacion)

    def _confirmar_importacion(self, analisis, reemplazar, incluir_opcional):
        try:
            with self._abrir_base() as con:
                resultado = importacion.importar(con, analisis.periodos, reemplazar, incluir_opcional)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo importar el historial.\n{e}")
            return False

        self.actualizar()
        importados = len(resultado.registrados) + len(resultado.reemplazados)
        lineas = [f"Se registraron {importados} período(s) en el historial."]
        if resultado.reemplazados:
            lineas.append("Reemplazados: " + ", ".join(resultado.reemplazados) + ".")
        if resultado.salteados:
            lineas.append("Salteados porque ya estaban registrados: " + ", ".join(resultado.salteados) + ".")
        messagebox.showinfo("Historial importado", "\n".join(lineas))
        return True
