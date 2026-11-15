import tkinter as tk
from contextlib import closing
from pathlib import Path
from tkinter import messagebox

import ttkbootstrap as tb
from ttkbootstrap.scrolled import ScrolledText

from .. import base_datos, correo
from .utiles import en_segundo_plano


def _tamaño(ruta):
    kb = Path(ruta).stat().st_size / 1024
    return f"{kb / 1024:.1f} MB" if kb >= 1024 else f"{max(kb, 1):.0f} KB"


def iniciar_envio(app, periodos):
    """Revisa la configuración, los Excel y los envíos anteriores; si todo está bien abre la vista previa."""
    datos = app.ajustes["correo"]
    contraseña = correo.descifrar(datos["contraseña_cifrada"])
    faltan = [texto for texto, falta in (("la cuenta de Gmail", not datos["cuenta"]),
                                         ("la contraseña de aplicación", not contraseña)) if falta]
    if faltan:
        if messagebox.askyesno("Correo sin configurar",
                               f"Para enviar la liquidación falta configurar {correo.unir(faltan)}.\n\n"
                               f"¿Configurarlo ahora?"):
            app.configuracion.configurar_correo(mostrar_pestaña=True)
        return None

    con_excel = [p for p in periodos if p.archivo and Path(p.archivo).is_file()]
    if len(con_excel) < len(periodos):
        detalle = "\n".join(
            f"  • {p.nombre}: " + (f"no se encontró {p.archivo}" if p.archivo else "no tiene Excel generado")
            for p in periodos if p not in con_excel)
        if not con_excel:
            messagebox.showwarning("Nada para enviar", "No hay Excel para adjuntar:\n\n" + detalle)
            return None
        if not messagebox.askyesno("Excel faltantes",
                                   f"Faltan estos Excel:\n\n{detalle}\n\n"
                                   f"¿Enviar solo {correo.unir(p.nombre for p in con_excel)}?"):
            return None
        periodos = con_excel

    try:
        with closing(base_datos.abrir(app.archivo_base)) as con:
            registradas = [base_datos.buscar_liquidacion(con, p.mes, p.año) for p in periodos]
    except Exception:
        registradas = []
    enviadas = [r for r in registradas if r and r.enviada]
    if enviadas and not messagebox.askyesno(
            "Ya enviada",
            "\n".join(f"{r.nombre} ya se envió el {r.enviada:%d/%m/%Y} a las {r.enviada:%H:%M}." for r in enviadas)
            + "\n\n¿Enviar igual?"):
        return None

    def al_enviar(destinatarios):
        aviso = ""
        try:
            with closing(base_datos.abrir(app.archivo_base)) as con:
                base_datos.marcar_enviadas(con, [(p.mes, p.año) for p in periodos if p.registrado])
        except Exception as e:
            aviso = f"\n\nNo se pudo guardar la fecha de envío en el historial.\n{e}"
        app.historial.actualizar()
        mensaje = (f"Se envió la liquidación de {correo.valores_marcadores(periodos)['periodo']} a "
                   f"{', '.join(destinatarios)}.{aviso}")
        (messagebox.showwarning if aviso else messagebox.showinfo)("Mail enviado", mensaje)

    return DialogoEnvio(app.root, periodos, datos, contraseña, al_enviar)


class DialogoEnvio:
    """Vista previa del mail con la plantilla completada; los cambios valen solo para este envío."""

    def __init__(self, padre, periodos, datos, contraseña, al_enviar):
        self.periodos = sorted(periodos, key=lambda p: (p.año, p.mes))
        self.cuenta = datos["cuenta"]
        self.contraseña = contraseña
        self.al_enviar = al_enviar
        self.enviando = False

        self.ventana = tb.Toplevel(title="Enviar liquidación")
        self.ventana.transient(padre)
        self.ventana.protocol("WM_DELETE_WINDOW", self.cancelar)
        marco = tb.Frame(self.ventana, padding=15)
        marco.pack(expand=True, fill="both")
        marco.columnconfigure(1, weight=1)
        marco.rowconfigure(4, weight=1)

        tb.Label(marco, text="De:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=3)
        tb.Label(marco, text=self.cuenta).grid(row=0, column=1, sticky="w", pady=3)
        self.ent_para = self._fila(marco, 1, "Para:", datos["para"])
        self.ent_cc = self._fila(marco, 2, "CC:", datos["cc"])
        self.ent_asunto = self._fila(marco, 3, "Asunto:", correo.completar_plantilla(datos["asunto"], self.periodos))
        tb.Label(marco, text="Cuerpo:").grid(row=4, column=0, sticky="ne", padx=(0, 8), pady=3)
        self.txt_cuerpo = ScrolledText(marco, height=9, width=64, wrap="word", autohide=True)
        self.txt_cuerpo.text.insert("1.0", correo.completar_plantilla(datos["cuerpo"], self.periodos))
        self.txt_cuerpo.grid(row=4, column=1, sticky="nsew", pady=3)

        tb.Label(marco, text="Adjuntos:").grid(row=5, column=0, sticky="ne", padx=(0, 8), pady=(8, 3))
        tb.Label(marco, justify="left", text="\n".join(
            f"• {Path(p.archivo).name} ({_tamaño(p.archivo)})" for p in self.periodos)).grid(
            row=5, column=1, sticky="w", pady=(8, 3))
        tb.Label(marco, text="Los cambios valen solo para este envío; la plantilla se edita en Configuración.",
                 bootstyle="info").grid(row=6, column=1, sticky="w", pady=(8, 0))

        botones = tb.Frame(marco)
        botones.grid(row=7, column=0, columnspan=2, sticky="we", pady=(15, 0))
        self.barra = tb.Progressbar(botones, mode="indeterminate", length=120, bootstyle="info")
        self.lbl_estado = tb.Label(botones, text="")
        self.lbl_estado.pack(side=tk.LEFT)
        self.btn_cancelar = tb.Button(botones, text="Cancelar", bootstyle="info-outline", width=12,
                                      command=self.cancelar)
        self.btn_cancelar.pack(side=tk.RIGHT)
        self.btn_enviar = tb.Button(botones, text="Enviar", bootstyle="success", width=12, command=self.enviar)
        self.btn_enviar.pack(side=tk.RIGHT, padx=5)

        self.ventana.place_window_center()
        self.ventana.grab_set()

    @staticmethod
    def _fila(padre, fila, etiqueta, valor):
        tb.Label(padre, text=etiqueta).grid(row=fila, column=0, sticky="e", padx=(0, 8), pady=3)
        entrada = tb.Entry(padre)
        entrada.insert(0, valor)
        entrada.grid(row=fila, column=1, sticky="we", pady=3)
        return entrada

    def _error(self, titulo, mensaje):
        messagebox.showerror(titulo, mensaje, parent=self.ventana)

    def enviar(self):
        para = correo.separar_direcciones(self.ent_para.get())
        cc = correo.separar_direcciones(self.ent_cc.get())
        if not para:
            self._error("Sin destinatarios", "Ingresá al menos una dirección en Para.")
            return
        invalidas = correo.direcciones_invalidas(para + cc)
        if invalidas:
            self._error("Direcciones inválidas",
                        "Estas direcciones no son válidas:\n\n" + "\n".join(f"  • {d}" for d in invalidas))
            return
        asunto = self.ent_asunto.get().strip()
        if not asunto:
            self._error("Sin asunto", "Ingresá el asunto del mail.")
            return
        try:
            mensaje = correo.armar_mensaje(self.cuenta, para, cc, asunto, self.txt_cuerpo.text.get("1.0", "end-1c"),
                                           [p.archivo for p in self.periodos])
        except correo.ErrorCorreo as e:
            self._error("No se puede enviar", str(e))
            return

        self._mostrar_enviando(True)
        en_segundo_plano(self.ventana, lambda: correo.enviar(mensaje, self.cuenta, self.contraseña),
                         lambda error: self._terminado(error, para + cc))

    def _mostrar_enviando(self, enviando):
        self.enviando = enviando
        estado = tk.DISABLED if enviando else tk.NORMAL
        self.btn_enviar.config(state=estado)
        self.btn_cancelar.config(state=estado)
        self.lbl_estado.config(text="Enviando…" if enviando else "")
        if enviando:
            self.barra.pack(side=tk.LEFT, padx=10)
            self.barra.start(15)
        else:
            self.barra.stop()
            self.barra.pack_forget()

    def _terminado(self, error, destinatarios):
        self._mostrar_enviando(False)
        if error:
            self._error("No se pudo enviar", str(error))
            return
        self.ventana.destroy()
        self.al_enviar(destinatarios)

    def cancelar(self):
        if not self.enviando:
            self.ventana.destroy()
