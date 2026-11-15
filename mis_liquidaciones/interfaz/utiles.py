import tkinter as tk

import ttkbootstrap as tb


def listar(numeros, maximo=10):
    texto = ", ".join(str(n) for n in numeros[:maximo])
    return texto + (f" y {len(numeros) - maximo} más" if len(numeros) > maximo else "")


def crear_tabla(padre, columnas, altura=8):
    """Treeview con barra de desplazamiento. columnas: (clave, título, ancho, alineación)."""
    marco = tb.Frame(padre)
    tabla = tb.Treeview(marco, columns=[c[0] for c in columnas], show="headings", height=altura,
                        selectmode="browse")
    for clave, titulo, ancho, alineacion in columnas:
        tabla.heading(clave, text=titulo, anchor=alineacion)
        tabla.column(clave, width=ancho, minwidth=40, anchor=alineacion, stretch=True)
    barra = tb.Scrollbar(marco, orient=tk.VERTICAL, command=tabla.yview)
    tabla.configure(yscrollcommand=barra.set)
    tabla.pack(side=tk.LEFT, expand=True, fill="both")
    barra.pack(side=tk.RIGHT, fill="y")
    return marco, tabla
