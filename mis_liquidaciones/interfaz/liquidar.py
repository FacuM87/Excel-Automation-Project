import os
import subprocess
import tkinter as tk
from contextlib import closing, nullcontext
from tkinter import filedialog, messagebox

import ttkbootstrap as tb

from .. import base_datos
from ..procesamiento import (
    MESES, PRESTACION_OPCIONAL, FormatoInvalido, detectar_periodos, formatear_monto, guardar_liquidacion,
    leer_planilla_origen, liquidar, normalizar_prestacion, ordenes_repetidas, periodo_mes_vencido,
)
from .interruptor import Interruptor
from .utiles import listar


class PestanaLiquidar:
    def __init__(self, padre, app):
        self.app = app
        self.ruta_archivo_seleccionado = None
        self.ruta_archivo_modificado = None
        self.periodos_detectados = []
        self.prestaciones_archivo = set()
        self.ordenes_sin_fecha = 0

        self.marco = tb.Frame(padre, padding=20)
        interior = tb.Frame(self.marco)
        interior.pack(expand=True, fill="x", padx=60)

        frame_archivo = tb.Labelframe(interior, text="Planilla recibida", padding=15)
        frame_archivo.pack(fill="x")
        tb.Button(frame_archivo, text="Seleccionar planilla…", bootstyle="primary", width=22,
                  command=self.seleccionar_archivo).pack(side=tk.LEFT)
        self.lbl_archivo = tb.Label(frame_archivo, text="Ningún archivo seleccionado")
        self.lbl_archivo.pack(side=tk.LEFT, padx=15)

        self.var_incluir_opcional = tk.BooleanVar(value=False)
        self.interruptor_opcional = Interruptor(interior, f"Incluir {PRESTACION_OPCIONAL}", self.var_incluir_opcional,
                                                command=self.actualizar_aviso_montos)
        self.interruptor_opcional.pack(anchor="w", pady=15)

        frame_periodo = tb.Labelframe(interior, text="Período a liquidar", padding=15)
        frame_periodo.pack(fill="x")
        frame_selector = tb.Frame(frame_periodo)
        frame_selector.pack(anchor="w")
        self.cmb_mes = tb.Combobox(frame_selector, values=MESES, state="readonly", width=12)
        self.cmb_mes.pack(side=tk.LEFT)
        self.var_año = tk.StringVar()
        self.spn_año = tb.Spinbox(frame_selector, from_=2000, to=2100, width=6, textvariable=self.var_año)
        self.spn_año.pack(side=tk.LEFT, padx=10)
        self.lbl_periodo = tb.Label(frame_periodo, text="", wraplength=560, justify="left")
        self.lbl_periodo.pack(anchor="w", pady=(8, 0))
        self.restablecer_mes_vencido()

        self.lbl_aviso_montos = tb.Label(interior, text="", bootstyle="warning", wraplength=600, justify="left")
        self.lbl_aviso_montos.pack(anchor="w", pady=(15, 0))

        frame_botones = tb.Frame(interior)
        frame_botones.pack(pady=20)
        tb.Button(frame_botones, text="Procesar planilla", bootstyle="success", width=22,
                  command=self.procesar).pack(side=tk.LEFT, padx=5, ipady=4)
        self.btn_abrir_ubicacion = tb.Button(frame_botones, text="Abrir carpeta", bootstyle="info-outline",
                                             width=16, command=self.abrir_ubicacion, state=tk.DISABLED)
        self.btn_abrir_ubicacion.pack(side=tk.LEFT, padx=5, ipady=4)

        self.actualizar_aviso_montos()

    # --- Período --- #

    def restablecer_mes_vencido(self):
        mes, año = periodo_mes_vencido()
        self.cmb_mes.set(MESES[mes - 1])
        self.var_año.set(str(año))

    def obtener_periodo_seleccionado(self):
        """Devuelve (numero_mes, año) del período elegido, o lanza ValueError."""
        nombre_mes = self.cmb_mes.get()
        if nombre_mes not in MESES:
            raise ValueError("Seleccioná un mes válido.")
        try:
            año = int(self.spn_año.get())
        except ValueError:
            raise ValueError("Ingresá un año válido.")
        if not 2000 <= año <= 2100:
            raise ValueError("Ingresá un año válido.")
        return MESES.index(nombre_mes) + 1, año

    def _habilitar_selector(self, habilitado):
        self.cmb_mes.config(state="readonly" if habilitado else tk.DISABLED)
        self.spn_año.config(state=tk.NORMAL if habilitado else tk.DISABLED)

    def detectar_periodo(self):
        """Sugiere el período según las fechas del archivo seleccionado."""
        try:
            df, _ = leer_planilla_origen(self.ruta_archivo_seleccionado)
            self.periodos_detectados = detectar_periodos(df)
            self.ordenes_sin_fecha = len(df) - sum(p.ordenes for p in self.periodos_detectados)
            self.prestaciones_archivo = set(df["Prestación"].fillna("").astype(str).map(normalizar_prestacion))
        except Exception:
            self.periodos_detectados = []
            self.ordenes_sin_fecha = 0
            self.prestaciones_archivo = set()

        periodos = self.periodos_detectados
        self._habilitar_selector(len(periodos) <= 1)
        if not periodos:
            self.lbl_periodo.config(text="No se pudo detectar el período; elegilo a mano.")
        elif len(periodos) == 1:
            self.cmb_mes.set(MESES[periodos[0].mes - 1])
            self.var_año.set(str(periodos[0].año))
            self.lbl_periodo.config(text=f"Detectado en el archivo: {periodos[0].nombre}")
        else:
            detalle = ", ".join(f"{p.nombre} ({p.ordenes})" for p in periodos)
            self.lbl_periodo.config(
                text=f"El archivo tiene órdenes de varios meses: {detalle}. Se va a generar una liquidación por mes.")
        self.actualizar_aviso_montos()

    # --- Montos --- #

    def prestaciones_sin_monto(self):
        """Prestaciones sin monto válido: las del archivo elegido o, si no hay archivo, las configuradas."""
        montos, _ = self.app.configuracion.leer_y_validar_montos()
        candidatas = self.prestaciones_archivo if self.ruta_archivo_seleccionado else set(self.app.configuracion.entries)
        if not self.var_incluir_opcional.get():
            candidatas = candidatas - {PRESTACION_OPCIONAL}
        return sorted(p for p in candidatas if p not in montos)

    def actualizar_aviso_montos(self):
        sin_monto = self.prestaciones_sin_monto()
        texto = ""
        if sin_monto:
            lista = ", ".join(p if p else "(vacío)" for p in sin_monto)
            origen = "del archivo " if self.ruta_archivo_seleccionado else ""
            texto = f"⚠ Prestaciones {origen}sin monto (se liquidan en 0): {lista}. Se cargan en Configuración."
        self.lbl_aviso_montos.config(text=texto)

    # --- Acciones --- #

    def seleccionar_archivo(self, ruta=None):
        if ruta is None:
            ruta = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if ruta:
            self.ruta_archivo_seleccionado = ruta
            self.lbl_archivo.config(text=os.path.basename(ruta))
            self.detectar_periodo()

    def procesar(self):
        if not self.ruta_archivo_seleccionado:
            messagebox.showerror("Error", "Primero selecciona un archivo de Excel")
            return

        montos, errores = self.app.configuracion.leer_y_validar_montos()
        if errores:
            messagebox.showerror("Montos inválidos",
                                 "Corregí los siguientes montos en Configuración:\n\n" + "\n".join(errores))
            return

        varios_meses = len(self.periodos_detectados) > 1
        if varios_meses:
            detalle = "\n".join(f"  • {p.nombre} ({p.ordenes} orden(es))" for p in self.periodos_detectados)
            cantidad = len(self.periodos_detectados)
            if not messagebox.askyesno(
                    "Varios meses",
                    f"El archivo tiene órdenes de {cantidad} meses. Se van a generar {cantidad} liquidaciones, "
                    f"una por mes:\n\n{detalle}\n\n¿Continuar?"):
                return
        else:
            try:
                mes, año = self.obtener_periodo_seleccionado()
            except ValueError as e:
                messagebox.showerror("Período inválido", str(e))
                return
            if len(self.periodos_detectados) == 1:
                detectado = self.periodos_detectados[0]
                if (mes, año) != (detectado.mes, detectado.año) and not messagebox.askyesno(
                        "Período distinto",
                        f"El período elegido ({MESES[mes - 1]} {año}) no coincide con las fechas del archivo "
                        f"({detectado.nombre}). ¿Generar la liquidación igual?"):
                    return

        self.app.configuracion.guardar_montos()
        incluir_opcional = self.var_incluir_opcional.get()
        periodos = [(p.mes, p.año) for p in self.periodos_detectados] if varios_meses else [(mes, año)]

        try:
            if varios_meses:
                liquidaciones = [liquidar(self.ruta_archivo_seleccionado, montos, incluir_opcional, p) for p in periodos]
            else:
                liquidaciones = [liquidar(self.ruta_archivo_seleccionado, montos, incluir_opcional)]
        except FormatoInvalido as e:
            messagebox.showerror("Formato no reconocido", f"El archivo no tiene el formato esperado.\n{e}")
            return
        except ImportError as e:
            messagebox.showerror("Error", f"Falta una librería para leer este tipo de archivo\n{e}")
            return
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar el archivo\n{e}")
            return

        advertencias = []
        try:
            con = base_datos.abrir(self.app.archivo_base)
        except Exception as e:
            con = None
            advertencias.append(f"No se pudo abrir el historial ({self.app.archivo_base}); no se registró nada.\n{e}")

        with closing(con) if con else nullcontext():
            registrar = self._confirmar_registro(con, liquidaciones, periodos, advertencias)
            if registrar is None:
                return

            resultados = []
            for liquidacion, (mes, año), registrar_periodo in zip(liquidaciones, periodos, registrar):
                try:
                    ruta = guardar_liquidacion(liquidacion, self.app.carpeta_guardado, mes, año)
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo guardar la liquidación de {MESES[mes - 1]} {año}\n{e}")
                    return
                estado = "no se registró en el historial"
                if con and registrar_periodo:
                    try:
                        reemplazada = base_datos.guardar(
                            con, mes, año, liquidacion, self.ruta_archivo_seleccionado, ruta, incluir_opcional)
                        estado = "reemplazó la registrada en el historial" if reemplazada else "registrada en el historial"
                    except Exception as e:
                        advertencias.append(
                            f"La liquidación de {MESES[mes - 1]} {año} se generó, pero no se pudo registrar "
                            f"en el historial.\n{e}")
                resultados.append((liquidacion, ruta, estado))

        self.ruta_archivo_modificado = resultados[-1][1]
        self.app.historial.actualizar()
        self._mostrar_resultado(resultados, varios_meses, advertencias)
        self.btn_abrir_ubicacion.config(state=tk.NORMAL)

    def _confirmar_registro(self, con, liquidaciones, periodos, advertencias):
        """Pregunta por períodos ya registrados y órdenes repetidas. Devuelve, por período, si se registra
        en la base, o None si se cancela."""
        registrar = []
        repetidas = []
        for liquidacion, (mes, año) in zip(liquidaciones, periodos):
            nombre = f"{MESES[mes - 1]} {año}"
            if not liquidacion.tiene_numero_orden:
                advertencias.append(
                    f"El archivo no tiene la columna Orden: no se pudo verificar si hay órdenes repetidas ({nombre}).")
            else:
                en_archivo = ordenes_repetidas(liquidacion.tabla)
                if en_archivo:
                    repetidas.append(f"  • {len(en_archivo)} orden(es) aparecen más de una vez en el archivo "
                                     f"({nombre}): {listar(en_archivo)}")
            if con is None:
                registrar.append(False)
                continue

            existente = base_datos.buscar_liquidacion(con, mes, año)
            if existente:
                respuesta = messagebox.askyesnocancel(
                    "Período ya registrado",
                    f"Ya hay una liquidación de {nombre} registrada el {existente.generada:%d/%m/%Y} "
                    f"({existente.cantidad_ordenes} órdenes, total $ {formatear_monto(existente.total)}).\n\n"
                    f"¿Reemplazarla?\n\n"
                    f"Sí: reemplaza la registrada y genera el Excel.\n"
                    f"No: genera solo el Excel, sin cambiar el historial.\n"
                    f"Cancelar: no hace nada.")
                if respuesta is None:
                    return None
                registrar.append(respuesta)
            else:
                registrar.append(True)

            if liquidacion.tiene_numero_orden:
                otras = base_datos.buscar_ordenes_registradas(con, liquidacion.tabla["Orden"], (mes, año))
                for (mes_otro, año_otro), numeros in otras.items():
                    repetidas.append(f"  • {len(numeros)} orden(es) de {nombre} ya están en la liquidación de "
                                     f"{MESES[mes_otro - 1]} {año_otro}: {listar(numeros)}")

        if repetidas and not messagebox.askyesno(
                "Órdenes repetidas",
                "Se encontraron órdenes repetidas:\n\n" + "\n".join(repetidas) +
                "\n\nSi continuás, se liquidan igual. ¿Continuar?"):
            return None
        return registrar

    def _mostrar_resultado(self, resultados, varios_meses=False, otras_advertencias=()):
        """resultados: lista de (liquidacion, ruta, estado). Junta los avisos de todas las liquidaciones."""
        liquidaciones = [liq for liq, _, _ in resultados]
        no_cumplidas = liquidaciones[0].no_cumplidas
        sin_monto = sorted({p for liq in liquidaciones for p in liq.sin_monto})
        cantidad_sin_monto = sum(liq.cantidad_sin_monto for liq in liquidaciones)
        excluidas_opcional = sum(liq.excluidas_opcional for liq in liquidaciones)

        advertencias = []
        if no_cumplidas:
            advertencias.append(f"Se excluyeron {no_cumplidas} orden(es) no cumplida(s).")
        if sin_monto:
            lista = "\n".join(f"  • {p if p else '(vacío)'}" for p in sin_monto)
            advertencias.append(
                f"{cantidad_sin_monto} fila(s) quedaron con monto 0 porque su prestación "
                f"no está configurada o no tiene monto cargado:\n{lista}"
            )
        if varios_meses and self.ordenes_sin_fecha:
            advertencias.append(
                f"{self.ordenes_sin_fecha} orden(es) sin fecha válida no se incluyeron en ninguna liquidación.")
        advertencias.extend(otras_advertencias)

        if len(resultados) == 1:
            mensaje = f"Archivo guardado en: {resultados[0][1]}\nLiquidación {resultados[0][2]}."
        else:
            archivos = "\n".join(
                f"  • {ruta.name} ({len(liq.tabla)} orden(es)): {estado}" for liq, ruta, estado in resultados)
            mensaje = f"Se generaron {len(resultados)} liquidaciones en {resultados[0][1].parent}:\n{archivos}"
        if excluidas_opcional:
            mensaje += (
                f"\n\nNo se incluyeron {excluidas_opcional} orden(es) de {PRESTACION_OPCIONAL} "
                f"(casilla \"Incluir {PRESTACION_OPCIONAL}\" desmarcada)."
            )
        if advertencias:
            mensaje += "\n\nAtención:\n" + "\n\n".join(advertencias)
            messagebox.showwarning("Procesado con advertencias", mensaje)
        else:
            messagebox.showinfo("Éxito", mensaje)

    def abrir_ubicacion(self):
        if self.ruta_archivo_modificado and os.path.exists(self.ruta_archivo_modificado):
            carpeta = os.path.dirname(self.ruta_archivo_modificado)
            subprocess.run(["explorer", carpeta] if os.name == "nt" else ["xdg-open", carpeta])
        else:
            messagebox.showwarning("Aviso", "No hay archivo generado aún.")
