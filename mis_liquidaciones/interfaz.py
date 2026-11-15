import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import config
from .procesamiento import (
    MESES, PRESTACION_OPCIONAL, FormatoInvalido, detectar_periodos, guardar_liquidacion,
    leer_planilla_origen, liquidar, normalizar_prestacion, parsear_monto, periodo_mes_vencido,
)


class VentanaPrincipal:
    def __init__(self, archivo_montos=config.ARCHIVO_MONTOS, archivos_anteriores=config.ARCHIVOS_MONTOS_ANTERIORES):
        self.archivo_montos = archivo_montos
        self.montos_prestaciones = config.cargar_montos(archivo_montos, archivos_anteriores)
        self.carpeta_guardado = config.obtener_ruta_guardado()
        self.ruta_archivo_seleccionado = None
        self.ruta_archivo_modificado = None
        self.periodos_detectados = []
        self.ordenes_sin_fecha = 0
        self.entries = {}

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

        try:
            if varios_meses:
                resultados = []
                for p in self.periodos_detectados:
                    liquidacion = liquidar(self.ruta_archivo_seleccionado, montos, incluir_opcional, (p.mes, p.año))
                    resultados.append((liquidacion, guardar_liquidacion(liquidacion, self.carpeta_guardado, p.mes, p.año)))
            else:
                liquidacion = liquidar(self.ruta_archivo_seleccionado, montos, incluir_opcional)
                resultados = [(liquidacion, guardar_liquidacion(liquidacion, self.carpeta_guardado, mes, año))]
        except FormatoInvalido as e:
            messagebox.showerror("Formato no reconocido", f"El archivo no tiene el formato esperado.\n{e}")
            return
        except ImportError as e:
            messagebox.showerror("Error", f"Falta una librería para leer este tipo de archivo\n{e}")
            return
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar el archivo\n{e}")
            return

        self.ruta_archivo_modificado = resultados[-1][1]
        self._mostrar_resultado(resultados, varios_meses)
        self.btn_abrir_ubicacion.config(state=tk.NORMAL)

    def _mostrar_resultado(self, resultados, varios_meses=False):
        """resultados: lista de (liquidacion, ruta). Junta los avisos de todas las liquidaciones."""
        liquidaciones = [liq for liq, _ in resultados]
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

        if len(resultados) == 1:
            mensaje = f"Archivo guardado en: {resultados[0][1]}"
        else:
            archivos = "\n".join(f"  • {ruta.name} ({len(liq.tabla)} orden(es))" for liq, ruta in resultados)
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
