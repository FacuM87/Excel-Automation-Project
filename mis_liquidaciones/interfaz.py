import os
import subprocess
import tkinter as tk
from contextlib import closing, nullcontext
from tkinter import filedialog, messagebox, ttk

from . import base_datos, config, importacion
from .procesamiento import (
    MESES, PRESTACION_OPCIONAL, FormatoInvalido, detectar_periodos, formatear_monto, guardar_liquidacion,
    leer_planilla_origen, liquidar, normalizar_prestacion, ordenes_repetidas, parsear_monto, periodo_mes_vencido,
)


class VentanaPrincipal:
    def __init__(self, archivo_montos=config.ARCHIVO_MONTOS, archivos_anteriores=config.ARCHIVOS_MONTOS_ANTERIORES,
                 archivo_base=config.ARCHIVO_BASE):
        self.archivo_montos = archivo_montos
        self.archivo_base = archivo_base
        self.montos_prestaciones = config.cargar_montos(archivo_montos, archivos_anteriores)
        self.carpeta_guardado = config.obtener_ruta_guardado()
        self.ruta_archivo_seleccionado = None
        self.ruta_archivo_modificado = None
        self.periodos_detectados = []
        self.ordenes_sin_fecha = 0
        self.entries = {}
        self.dialogo_importacion = None

        self.root = tk.Tk()
        self.root.title("Cargar y Modificar Excel")
        self.root.minsize(500, 0)
        self._cargar_icono()
        self._construir()
        self._centrar()

    def ejecutar(self):
        self.root.mainloop()

    # --- Construcción de la ventana --- #

    def _cargar_icono(self):
        try:
            icon_path = config.obtener_directorio_recursos() / "icon.ico"
            if icon_path.exists():
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"No se pudo cargar el ícono: {e}")

    def _construir(self):
        frame_contenedor = tk.Frame(self.root)
        frame_contenedor.pack(expand=True, fill="both")

        frame_interior = tk.Frame(frame_contenedor)
        frame_interior.pack(expand=True, padx=20, pady=10)

        tk.Button(frame_interior, text="Seleccionar Planilla", width=25, command=self.seleccionar_archivo).pack(pady=10)

        self.lbl_archivo = tk.Label(frame_interior, text="Ningún archivo seleccionado", fg="blue")
        self.lbl_archivo.pack()

        self.var_incluir_opcional = tk.BooleanVar(value=False)
        tk.Checkbutton(frame_interior, text=f"Incluir {PRESTACION_OPCIONAL}", variable=self.var_incluir_opcional).pack(pady=(10, 0))

        frame_periodo = tk.LabelFrame(frame_interior, text="Período a liquidar")
        frame_periodo.pack(pady=10)

        frame_selector = tk.Frame(frame_periodo)
        frame_selector.pack()

        self.cmb_mes = ttk.Combobox(frame_selector, values=MESES, state="readonly", width=12)
        self.cmb_mes.pack(side=tk.LEFT, padx=5, pady=5)

        self.var_año = tk.StringVar()
        self.spn_año = tk.Spinbox(frame_selector, from_=2000, to=2100, width=6, textvariable=self.var_año)
        self.spn_año.pack(side=tk.LEFT, padx=5, pady=5)

        self.lbl_periodo = tk.Label(frame_periodo, text="", fg="gray25", wraplength=420, justify="center")
        self.lbl_periodo.pack(padx=5, pady=(0, 5))

        self.restablecer_mes_vencido()

        frame_montos = tk.LabelFrame(frame_interior, text="Montos por prestación")
        frame_montos.pack(pady=10, fill="x")

        self.frame_lista = tk.Frame(frame_montos)
        self.frame_lista.pack(pady=5)
        self.dibujar_prestaciones()

        frame_nueva = tk.Frame(frame_montos)
        frame_nueva.pack(pady=(5, 8))
        tk.Label(frame_nueva, text="Nueva:").pack(side=tk.LEFT, padx=5)
        self.ent_nueva_nombre = tk.Entry(frame_nueva, width=14)
        self.ent_nueva_nombre.pack(side=tk.LEFT, padx=2)
        tk.Label(frame_nueva, text="Monto:").pack(side=tk.LEFT, padx=5)
        self.ent_nueva_monto = tk.Entry(frame_nueva, width=10)
        self.ent_nueva_monto.pack(side=tk.LEFT, padx=2)
        tk.Button(frame_nueva, text="Agregar", command=self.agregar_prestacion).pack(side=tk.LEFT, padx=5)

        tk.Button(frame_interior, text="Procesar Planilla", width=25, command=self.procesar).pack(pady=10)

        self.btn_abrir_ubicacion = tk.Button(
            frame_interior, text="Abrir Ubicación de Planilla Modificada", width=30,
            command=self.abrir_ubicacion, state=tk.DISABLED)
        self.btn_abrir_ubicacion.pack(pady=10)

        tk.Button(frame_interior, text="Importar historial…", width=25, command=self.importar_historial).pack(pady=(0, 10))

    def _centrar(self):
        self.root.update_idletasks()
        ancho = self.root.winfo_reqwidth()
        alto = self.root.winfo_reqheight()
        x = (self.root.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.root.winfo_screenheight() // 2) - (alto // 2)
        self.root.geometry(f"+{x}+{max(y, 0)}")

    # --- Prestaciones --- #

    def dibujar_prestaciones(self):
        for widget in self.frame_lista.winfo_children():
            widget.destroy()
        self.entries.clear()
        for i, prestacion in enumerate(self.montos_prestaciones):
            fila, col = divmod(i, 2)
            col *= 3
            tk.Label(self.frame_lista, text=f"{prestacion}:").grid(row=fila, column=col, sticky="e", padx=(10, 2), pady=2)
            entry = tk.Entry(self.frame_lista, width=10)
            entry.insert(0, self.montos_prestaciones[prestacion])
            entry.grid(row=fila, column=col + 1, pady=2)
            tk.Button(self.frame_lista, text="✕", width=2, command=lambda p=prestacion: self.quitar_prestacion(p)).grid(
                row=fila, column=col + 2, padx=(2, 10), pady=2)
            self.entries[prestacion] = entry

    def sincronizar_montos(self):
        for prestacion, entry in self.entries.items():
            self.montos_prestaciones[prestacion] = entry.get()

    def guardar_montos(self):
        self.sincronizar_montos()
        try:
            config.guardar_montos(self.montos_prestaciones, self.archivo_montos)
        except OSError as e:
            messagebox.showwarning("Aviso", f"No se pudieron guardar los montos en {self.archivo_montos}\n{e}")

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
        self.montos_prestaciones[nombre] = monto
        self.dibujar_prestaciones()
        self.guardar_montos()
        self.ent_nueva_nombre.delete(0, tk.END)
        self.ent_nueva_monto.delete(0, tk.END)

    def quitar_prestacion(self, prestacion):
        if not messagebox.askyesno("Quitar prestación", f"¿Quitar {prestacion} de la lista?"):
            return
        self.sincronizar_montos()
        del self.montos_prestaciones[prestacion]
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
        except Exception:
            self.periodos_detectados = []
            self.ordenes_sin_fecha = 0

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

    # --- Acciones --- #

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if ruta:
            self.ruta_archivo_seleccionado = ruta
            self.lbl_archivo.config(text=f"Archivo seleccionado: {os.path.basename(ruta)}")
            self.detectar_periodo()

    def procesar(self):
        if not self.ruta_archivo_seleccionado:
            messagebox.showerror("Error", "Primero selecciona un archivo de Excel")
            return

        montos, errores = self.leer_y_validar_montos()
        if errores:
            messagebox.showerror("Montos inválidos", "Corregí los siguientes montos:\n\n" + "\n".join(errores))
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

        self.guardar_montos()
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
            con = base_datos.abrir(self.archivo_base)
        except Exception as e:
            con = None
            advertencias.append(f"No se pudo abrir el historial ({self.archivo_base}); no se registró nada.\n{e}")

        with closing(con) if con else nullcontext():
            registrar = self._confirmar_registro(con, liquidaciones, periodos, advertencias)
            if registrar is None:
                return

            resultados = []
            for liquidacion, (mes, año), registrar_periodo in zip(liquidaciones, periodos, registrar):
                try:
                    ruta = guardar_liquidacion(liquidacion, self.carpeta_guardado, mes, año)
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
                                     f"({nombre}): {_listar(en_archivo)}")
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
                                     f"{MESES[mes_otro - 1]} {año_otro}: {_listar(numeros)}")

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

    def importar_historial(self, rutas=None):
        """Carga en la base los archivos ya recibidos, sin generar Excel."""
        if rutas is None:
            rutas = filedialog.askopenfilenames(
                title="Elegí las planillas a importar", filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not rutas:
            return

        montos, errores = self.leer_y_validar_montos()
        if errores:
            messagebox.showerror("Montos inválidos", "Corregí los siguientes montos:\n\n" + "\n".join(errores))
            return
        self.guardar_montos()
        incluir_opcional = self.var_incluir_opcional.get()

        analisis = importacion.analizar_archivos(rutas, montos, incluir_opcional)
        if not analisis.periodos:
            detalle = "\n".join(f"  • {nombre}: {error}" for nombre, error in analisis.errores)
            messagebox.showerror("Nada para importar", "No se encontraron órdenes para importar.\n\n" + detalle)
            return

        try:
            with closing(base_datos.abrir(self.archivo_base)) as con:
                importacion.verificar_en_base(con, analisis.periodos)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el historial ({self.archivo_base}).\n{e}")
            return

        self.dialogo_importacion = DialogoImportacion(self.root, analisis, incluir_opcional, self._confirmar_importacion)

    def _confirmar_importacion(self, analisis, reemplazar, incluir_opcional):
        try:
            with closing(base_datos.abrir(self.archivo_base)) as con:
                resultado = importacion.importar(con, analisis.periodos, reemplazar, incluir_opcional)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo importar el historial.\n{e}")
            return False

        importados = len(resultado.registrados) + len(resultado.reemplazados)
        lineas = [f"Se registraron {importados} período(s) en el historial."]
        if resultado.reemplazados:
            lineas.append("Reemplazados: " + ", ".join(resultado.reemplazados) + ".")
        if resultado.salteados:
            lineas.append("Salteados porque ya estaban registrados: " + ", ".join(resultado.salteados) + ".")
        messagebox.showinfo("Historial importado", "\n".join(lineas))
        return True

    def abrir_ubicacion(self):
        if self.ruta_archivo_modificado and os.path.exists(self.ruta_archivo_modificado):
            carpeta = os.path.dirname(self.ruta_archivo_modificado)
            subprocess.run(["explorer", carpeta] if os.name == "nt" else ["xdg-open", carpeta])
        else:
            messagebox.showwarning("Aviso", "No hay archivo generado aún.")


class DialogoImportacion:
    """Vista previa de la importación: un renglón por período, avisos y confirmación."""

    def __init__(self, padre, analisis, incluir_opcional, al_confirmar):
        self.analisis = analisis
        self.incluir_opcional = incluir_opcional
        self.al_confirmar = al_confirmar

        self.ventana = tk.Toplevel(padre)
        self.ventana.title("Importar historial")
        self.ventana.transient(padre)
        self.ventana.protocol("WM_DELETE_WINDOW", self.cancelar)

        marco = tk.Frame(self.ventana)
        marco.pack(expand=True, fill="both", padx=15, pady=10)

        columnas = ("periodo", "archivos", "ordenes", "total", "estado")
        self.tabla = ttk.Treeview(marco, columns=columnas, show="headings", height=min(len(analisis.periodos), 12))
        for col, titulo, ancho, alineacion in (
                ("periodo", "Período", 120, "w"), ("archivos", "Archivos", 230, "w"),
                ("ordenes", "Órdenes", 150, "e"), ("total", "Total", 100, "e"), ("estado", "Estado", 110, "w")):
            self.tabla.heading(col, text=titulo)
            self.tabla.column(col, width=ancho, anchor=alineacion)
        for periodo in analisis.periodos:
            ordenes = str(len(periodo.liquidacion.tabla))
            if periodo.duplicadas_descartadas:
                ordenes += f" ({periodo.duplicadas_descartadas} repetidas)"
            self.tabla.insert("", tk.END, values=(
                periodo.nombre, ", ".join(periodo.archivos), ordenes,
                f"$ {formatear_monto(periodo.liquidacion.total)}",
                "Ya registrado" if periodo.registrado else "Nuevo"))
        self.tabla.pack(fill="x")

        avisos = self._avisos()
        if avisos:
            tk.Label(marco, text="Atención:", anchor="w").pack(fill="x", pady=(10, 0))
            self.txt_avisos = tk.Text(marco, height=min(len(avisos) + 2, 8), wrap="word")
            self.txt_avisos.insert("1.0", "\n".join(avisos))
            self.txt_avisos.config(state=tk.DISABLED)
            self.txt_avisos.pack(fill="x")

        self.var_reemplazar = tk.BooleanVar(value=False)
        registrados = [p for p in analisis.periodos if p.registrado]
        if registrados:
            tk.Checkbutton(
                marco, variable=self.var_reemplazar,
                text=f"Reemplazar los períodos ya registrados ({len(registrados)}); si no, se saltean").pack(
                anchor="w", pady=(10, 0))

        botones = tk.Frame(marco)
        botones.pack(pady=(10, 0))
        tk.Button(botones, text="Importar", width=12, command=self.confirmar).pack(side=tk.LEFT, padx=5)
        tk.Button(botones, text="Cancelar", width=12, command=self.cancelar).pack(side=tk.LEFT, padx=5)

        self.ventana.grab_set()

    def _avisos(self):
        avisos = [f"• No se pudo leer {nombre}: {error}" for nombre, error in self.analisis.errores]
        if self.analisis.sin_fecha:
            avisos.append(f"• {self.analisis.sin_fecha} orden(es) sin fecha válida no se importan.")
        sin_monto = sorted({p for per in self.analisis.periodos for p in per.liquidacion.sin_monto})
        if sin_monto:
            cantidad = sum(per.liquidacion.cantidad_sin_monto for per in self.analisis.periodos)
            avisos.append(f"• {cantidad} orden(es) quedan con monto 0 (prestación sin monto cargado): "
                          + ", ".join(p if p else "(vacío)" for p in sin_monto))
        excluidas = sum(per.liquidacion.excluidas_opcional for per in self.analisis.periodos)
        if excluidas:
            avisos.append(f"• No se importan {excluidas} orden(es) de {PRESTACION_OPCIONAL} "
                          f"(casilla \"Incluir {PRESTACION_OPCIONAL}\" desmarcada).")
        if self.analisis.no_cumplidas:
            avisos.append(f"• Se excluyeron {self.analisis.no_cumplidas} orden(es) no cumplida(s).")
        for periodo in self.analisis.periodos:
            if not periodo.liquidacion.tiene_numero_orden:
                avisos.append(f"• {periodo.nombre}: sin columna Orden, no se pudo verificar si hay órdenes repetidas.")
            repetidas = periodo.repetidas_en_archivo
            if repetidas:
                avisos.append(f"• {periodo.nombre}: {len(repetidas)} orden(es) aparecen más de una vez en el "
                              f"mismo archivo: {_listar(repetidas)}")
            for (mes, año), numeros in periodo.en_otros_periodos.items():
                avisos.append(f"• {len(numeros)} orden(es) de {periodo.nombre} ya están en la liquidación de "
                              f"{MESES[mes - 1]} {año}: {_listar(numeros)}")
        return avisos

    def confirmar(self):
        if self.al_confirmar(self.analisis, self.var_reemplazar.get(), self.incluir_opcional):
            self.ventana.destroy()

    def cancelar(self):
        self.ventana.destroy()


def _listar(numeros, maximo=10):
    texto = ", ".join(str(n) for n in numeros[:maximo])
    return texto + (f" y {len(numeros) - maximo} más" if len(numeros) > maximo else "")
