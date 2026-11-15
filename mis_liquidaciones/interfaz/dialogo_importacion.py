import tkinter as tk

import ttkbootstrap as tb
from ttkbootstrap.scrolled import ScrolledText

from ..procesamiento import MESES, PSICOMEY, formatear_monto
from .interruptor import Interruptor
from .utiles import crear_tabla, listar

MARCADO, DESMARCADO = "☑", "☐"


class DialogoImportacion:
    """Vista previa de la importación: un renglón por período, avisos y confirmación."""

    def __init__(self, padre, analisis, al_confirmar, carpeta=""):
        self.analisis = analisis
        self.al_confirmar = al_confirmar
        self.marcados = {(p.mes, p.año) for p in analisis.periodos}

        self.ventana = tb.Toplevel(title="Importar historial")
        self.ventana.transient(padre)
        self.ventana.protocol("WM_DELETE_WINDOW", self.cancelar)

        marco = tb.Frame(self.ventana, padding=15)
        marco.pack(expand=True, fill="both")

        marco_tabla, self.tabla = crear_tabla(marco, (
            ("excel", "Excel", 50, "center"), ("periodo", "Período", 120, "w"), ("archivos", "Archivos", 320, "w"),
            ("ordenes", "Órdenes", 150, "e"), ("total", "Total", 110, "e"), ("estado", "Estado", 110, "w")),
            altura=min(len(analisis.periodos), 12))
        for periodo in analisis.periodos:
            ordenes = str(len(periodo.liquidacion.tabla))
            if periodo.duplicadas_descartadas:
                ordenes += f" ({periodo.duplicadas_descartadas} repetidas)"
            self.tabla.insert("", tk.END, iid=f"{periodo.mes}-{periodo.año}", values=(
                MARCADO, periodo.nombre, ", ".join(periodo.archivos), ordenes,
                f"$ {formatear_monto(periodo.liquidacion.total)}",
                "Ya registrado" if periodo.registrado else "Nuevo"))
        self.tabla.configure(selectmode="none")
        self.tabla.bind("<Button-1>", self._al_hacer_clic)
        marco_tabla.pack(fill="x")

        avisos = self._avisos()
        if avisos:
            tb.Label(marco, text="Atención:", bootstyle="warning").pack(anchor="w", pady=(10, 2))
            self.txt_avisos = ScrolledText(marco, height=min(len(avisos) + 2, 8), wrap="word", autohide=True)
            self.txt_avisos.text.insert("1.0", "\n".join(avisos))
            self.txt_avisos.text.config(state=tk.DISABLED)
            self.txt_avisos.pack(fill="x")

        self.var_reemplazar = tk.BooleanVar(value=False)
        registrados = [p for p in analisis.periodos if p.registrado]
        if registrados:
            Interruptor(marco, f"Reemplazar los períodos ya registrados ({len(registrados)}); si no, se saltean",
                        self.var_reemplazar).pack(anchor="w", pady=(10, 0))

        self.var_generar = tk.BooleanVar(value=False)
        self.var_generar.trace_add("write", lambda *_: self._mostrar_opciones_excel())
        Interruptor(marco, "Generar también los Excel de liquidación", self.var_generar).pack(
            anchor="w", pady=(10, 0))
        ancla = tb.Frame(marco)
        ancla.pack(fill="x")
        self.opciones_excel = tb.Frame(ancla)
        tb.Button(self.opciones_excel, text="Todos", bootstyle="info-outline", width=8,
                  command=lambda: self._marcar_todos(True)).pack(side=tk.LEFT)
        tb.Button(self.opciones_excel, text="Ninguno", bootstyle="info-outline", width=8,
                  command=lambda: self._marcar_todos(False)).pack(side=tk.LEFT, padx=(5, 0))
        tb.Label(self.opciones_excel, text=f"Se guardan en: {carpeta}").pack(side=tk.LEFT, padx=(15, 0))
        self._mostrar_opciones_excel()

        botones = tb.Frame(marco)
        botones.pack(pady=(15, 0))
        tb.Button(botones, text="Importar", bootstyle="success", width=12, command=self.confirmar).pack(
            side=tk.LEFT, padx=5)
        tb.Button(botones, text="Cancelar", bootstyle="info-outline", width=12, command=self.cancelar).pack(
            side=tk.LEFT, padx=5)

        self.ventana.place_window_center()
        self.ventana.grab_set()

    def _mostrar_opciones_excel(self):
        self.tabla.configure(cursor="hand2" if self.var_generar.get() else "")
        if self.var_generar.get():
            self.tabla.configure(displaycolumns="#all")
            self.opciones_excel.pack(anchor="w", pady=(8, 0))
        else:
            self.tabla.configure(displaycolumns=[c for c in self.tabla["columns"] if c != "excel"])
            self.opciones_excel.pack_forget()

    def _al_hacer_clic(self, evento):
        fila = self.tabla.identify_row(evento.y)
        if fila and self.var_generar.get():
            mes, año = map(int, fila.split("-"))
            self.marcar((mes, año), (mes, año) not in self.marcados)

    def marcar(self, periodo, marcado):
        (self.marcados.add if marcado else self.marcados.discard)(periodo)
        self.tabla.set(f"{periodo[0]}-{periodo[1]}", "excel", MARCADO if marcado else DESMARCADO)

    def _marcar_todos(self, marcado):
        for p in self.analisis.periodos:
            self.marcar((p.mes, p.año), marcado)

    def periodos_a_generar(self):
        return set(self.marcados) if self.var_generar.get() else set()

    def _avisos(self):
        avisos = [f"• No se pudo leer {nombre}: {error}" for nombre, error in self.analisis.errores]
        if self.analisis.sin_fecha:
            avisos.append(f"• {self.analisis.sin_fecha} orden(es) sin fecha válida no se importan.")
        sin_monto = sorted({p for per in self.analisis.periodos for p in per.liquidacion.sin_monto})
        if sin_monto:
            cantidad = sum(per.liquidacion.cantidad_sin_monto for per in self.analisis.periodos)
            avisos.append(f"• {cantidad} orden(es) quedan con monto 0 (prestación sin monto cargado): "
                          + ", ".join(p if p else "(vacío)" for p in sin_monto))
        duplicados = sum(per.liquidacion.psicomey_duplicados for per in self.analisis.periodos)
        if duplicados:
            avisos.append(f"• Se quitaron {duplicados} fila(s) de {PSICOMEY} que duplicaban otra orden.")
        if self.analisis.no_cumplidas:
            avisos.append(f"• Se excluyeron {self.analisis.no_cumplidas} orden(es) no cumplida(s).")
        for periodo in self.analisis.periodos:
            if not periodo.liquidacion.tiene_numero_orden:
                avisos.append(f"• {periodo.nombre}: sin columna Orden, no se pudo verificar si hay órdenes repetidas.")
            repetidas = periodo.repetidas_en_archivo
            if repetidas:
                avisos.append(f"• {periodo.nombre}: {len(repetidas)} orden(es) aparecen más de una vez en el "
                              f"mismo archivo: {listar(repetidas)}")
            for (mes, año), numeros in periodo.en_otros_periodos.items():
                avisos.append(f"• {len(numeros)} orden(es) de {periodo.nombre} ya están en la liquidación de "
                              f"{MESES[mes - 1]} {año}: {listar(numeros)}")
        return avisos

    def confirmar(self):
        if self.al_confirmar(self.analisis, self.var_reemplazar.get(), self.periodos_a_generar()):
            self.ventana.destroy()

    def cancelar(self):
        self.ventana.destroy()
