import tkinter as tk
from tkinter import messagebox

import ttkbootstrap as tb
from ttkbootstrap.scrolled import ScrolledText

from .. import correo
from .interruptor import Interruptor
from .utiles import crear_enlace, en_segundo_plano

TEXTO_MARCADORES = "Marcadores: " + "  ".join("{" + m + "}" for m in correo.MARCADORES)


class DialogoCorreo:
    """Cuenta de Gmail, destinatarios y plantilla del mail; permite mandar un mail de prueba."""

    def __init__(self, padre, datos, al_guardar):
        self.al_guardar = al_guardar
        self.enviando = False
        contraseña = correo.descifrar(datos["contraseña_cifrada"])

        self.ventana = tb.Toplevel(title="Configurar correo")
        self.ventana.transient(padre)
        self.ventana.protocol("WM_DELETE_WINDOW", self.cancelar)
        marco = tb.Frame(self.ventana, padding=15)
        marco.pack(expand=True, fill="both")

        cuenta = tb.Labelframe(marco, text="Cuenta de Gmail que envía", padding=10)
        cuenta.pack(fill="x")
        cuenta.columnconfigure(1, weight=1)
        tb.Label(cuenta, text="Cuenta:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=3)
        self.ent_cuenta = tb.Entry(cuenta, width=45)
        self.ent_cuenta.insert(0, datos["cuenta"])
        self.ent_cuenta.grid(row=0, column=1, sticky="we", pady=3)
        tb.Label(cuenta, text="Contraseña de aplicación:").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=3)
        self.ent_contraseña = tb.Entry(cuenta, width=45, show="•")
        self.ent_contraseña.insert(0, contraseña or "")
        self.ent_contraseña.grid(row=1, column=1, sticky="we", pady=3)
        self.var_mostrar = tk.BooleanVar(value=False)
        self.var_mostrar.trace_add("write", lambda *_: self.ent_contraseña.config(
            show="" if self.var_mostrar.get() else "•"))
        Interruptor(cuenta, "Mostrar", self.var_mostrar).grid(row=1, column=2, padx=(10, 0))
        if contraseña is None:
            tb.Label(cuenta, text="⚠ La contraseña guardada no se pudo leer (¿otro usuario de Windows u otra PC?). "
                                  "Cargala de nuevo.", bootstyle="warning", wraplength=520, justify="left").grid(
                row=2, column=0, columnspan=3, sticky="w", pady=(5, 0))
        tb.Label(cuenta, wraplength=560, justify="left", text=(
            "No es la contraseña normal de Gmail: son 16 letras que da Google y requieren tener activada la "
            "verificación en 2 pasos. Se guarda cifrada con tu usuario de Windows.")).grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(8, 2))
        crear_enlace(cuenta, "Crear una contraseña de aplicación en la cuenta de Google",
                     correo.URL_CONTRASEÑA_APLICACION).grid(row=4, column=0, columnspan=3, sticky="w")

        destinatarios = tb.Labelframe(marco, text="Destinatarios", padding=10)
        destinatarios.pack(fill="x", pady=(15, 0))
        destinatarios.columnconfigure(1, weight=1)
        self.ent_para = self._fila(destinatarios, 0, "Para:", datos["para"])
        self.ent_cc = self._fila(destinatarios, 1, "CC:", datos["cc"])
        tb.Label(destinatarios, text="Separá varias direcciones con coma.").grid(
            row=2, column=1, sticky="w", pady=(2, 0))

        plantilla = tb.Labelframe(marco, text="Plantilla del mail", padding=10)
        plantilla.pack(fill="both", expand=True, pady=(15, 0))
        plantilla.columnconfigure(1, weight=1)
        plantilla.rowconfigure(1, weight=1)
        self.ent_asunto = self._fila(plantilla, 0, "Asunto:", datos["asunto"])
        tb.Label(plantilla, text="Cuerpo:").grid(row=1, column=0, sticky="ne", padx=(0, 8), pady=3)
        self.txt_cuerpo = ScrolledText(plantilla, height=7, width=60, wrap="word", autohide=True)
        self.txt_cuerpo.text.insert("1.0", datos["cuerpo"])
        self.txt_cuerpo.grid(row=1, column=1, sticky="nsew", pady=3)
        tb.Label(plantilla, text=TEXTO_MARCADORES, bootstyle="info").grid(row=2, column=1, sticky="w", pady=(4, 0))
        tb.Label(plantilla, wraplength=560, justify="left", text=(
            "Al enviar se reemplazan por el período (Marzo 2026, o Marzo y Abril 2026), los meses, el año, la "
            "cantidad de órdenes y el total.")).grid(row=3, column=1, sticky="w")

        botones = tb.Frame(marco)
        botones.pack(fill="x", pady=(15, 0))
        self.btn_prueba = tb.Button(botones, text="Enviar prueba", bootstyle="info-outline", width=14,
                                    command=self.enviar_prueba)
        self.btn_prueba.pack(side=tk.LEFT)
        self.lbl_estado = tb.Label(botones, text="")
        self.lbl_estado.pack(side=tk.LEFT, padx=10)
        self.btn_cancelar = tb.Button(botones, text="Cancelar", bootstyle="info-outline", width=12,
                                      command=self.cancelar)
        self.btn_cancelar.pack(side=tk.RIGHT)
        self.btn_guardar = tb.Button(botones, text="Guardar", bootstyle="success", width=12, command=self.guardar)
        self.btn_guardar.pack(side=tk.RIGHT, padx=5)

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

    def _validar_cuenta(self):
        cuenta = self.ent_cuenta.get().strip()
        if cuenta and correo.direcciones_invalidas([cuenta]):
            self._error("Cuenta inválida", f"\"{cuenta}\" no es una dirección de correo válida.")
            return None
        return cuenta

    def enviar_prueba(self):
        cuenta = self._validar_cuenta()
        if cuenta is None:
            return
        contraseña = self.ent_contraseña.get().strip()
        if not cuenta or not contraseña:
            self._error("Faltan datos", "Cargá la cuenta de Gmail y la contraseña de aplicación.")
            return
        self.enviando = True
        for boton in (self.btn_prueba, self.btn_guardar, self.btn_cancelar):
            boton.config(state=tk.DISABLED)
        self.lbl_estado.config(text="Enviando prueba…")
        en_segundo_plano(self.ventana, lambda: correo.enviar(correo.mensaje_de_prueba(cuenta), cuenta, contraseña),
                         lambda error: self._prueba_terminada(cuenta, error))

    def _prueba_terminada(self, cuenta, error):
        self.enviando = False
        for boton in (self.btn_prueba, self.btn_guardar, self.btn_cancelar):
            boton.config(state=tk.NORMAL)
        self.lbl_estado.config(text="")
        if error:
            self._error("No se pudo enviar la prueba", str(error))
        else:
            messagebox.showinfo("Prueba enviada", f"Se envió un mail de prueba a {cuenta}.\n\n"
                                                  f"Revisá la bandeja de entrada.", parent=self.ventana)

    def guardar(self):
        cuenta = self._validar_cuenta()
        if cuenta is None:
            return
        direcciones = {}
        for nombre, entrada in (("para", self.ent_para), ("cc", self.ent_cc)):
            direcciones[nombre] = correo.separar_direcciones(entrada.get())
            invalidas = correo.direcciones_invalidas(direcciones[nombre])
            if invalidas:
                self._error("Direcciones inválidas", "Estas direcciones no son válidas:\n\n"
                            + "\n".join(f"  • {d}" for d in invalidas))
                return
        try:
            cifrada = correo.cifrar(self.ent_contraseña.get().strip())
        except OSError as e:
            self._error("Error", f"No se pudo cifrar la contraseña.\n{e}")
            return
        self.al_guardar({
            "cuenta": cuenta, "contraseña_cifrada": cifrada,
            "para": ", ".join(direcciones["para"]), "cc": ", ".join(direcciones["cc"]),
            "asunto": self.ent_asunto.get().strip(), "cuerpo": self.txt_cuerpo.text.get("1.0", "end-1c"),
        })
        self.ventana.destroy()

    def cancelar(self):
        if not self.enviando:
            self.ventana.destroy()
