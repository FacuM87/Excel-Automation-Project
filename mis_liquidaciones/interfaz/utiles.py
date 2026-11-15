import threading
import tkinter as tk
import tkinter.font as tkfont
import webbrowser

import ttkbootstrap as tb


def listar(numeros, maximo=10):
    texto = ", ".join(str(n) for n in numeros[:maximo])
    return texto + (f" y {len(numeros) - maximo} más" if len(numeros) > maximo else "")


def crear_tabla(padre, columnas, altura=8, selectmode="browse"):
    """Treeview con barra de desplazamiento. columnas: (clave, título, ancho, alineación)."""
    marco = tb.Frame(padre)
    tabla = tb.Treeview(marco, columns=[c[0] for c in columnas], show="headings", height=altura,
                        selectmode=selectmode)
    for clave, titulo, ancho, alineacion in columnas:
        tabla.heading(clave, text=titulo, anchor=alineacion)
        tabla.column(clave, width=ancho, minwidth=40, anchor=alineacion, stretch=True)
    barra = tb.Scrollbar(marco, orient=tk.VERTICAL, command=tabla.yview)
    tabla.configure(yscrollcommand=barra.set)
    tabla.pack(side=tk.LEFT, expand=True, fill="both")
    barra.pack(side=tk.RIGHT, fill="y")
    return marco, tabla


def en_segundo_plano(widget, funcion, al_terminar, intervalo=100):
    """Ejecuta funcion() en un hilo y llama a al_terminar(error) en el hilo de la interfaz (error None si salió bien)."""
    estado = {}

    def trabajar():
        try:
            funcion()
        except Exception as e:
            estado["error"] = e
        estado["listo"] = True

    def revisar():
        if estado.get("listo"):
            al_terminar(estado.get("error"))
        else:
            widget.after(intervalo, revisar)

    threading.Thread(target=trabajar, daemon=True).start()
    widget.after(intervalo, revisar)


def crear_enlace(padre, texto, url):
    """Etiqueta subrayada que abre la URL en el navegador."""
    fuente = tkfont.nametofont("TkDefaultFont").copy()
    fuente.configure(underline=True)
    enlace = tb.Label(padre, text=texto, bootstyle="info", cursor="hand2", font=fuente)
    enlace.fuente = fuente
    enlace.bind("<Button-1>", lambda _e: webbrowser.open(url))
    return enlace
