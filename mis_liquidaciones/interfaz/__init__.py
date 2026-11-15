from tkinter import messagebox

import ttkbootstrap as tb

from .. import config
from .configuracion import PestanaConfiguracion
from .historial import PestanaHistorial
from .liquidar import PestanaLiquidar
from .reportes import PestanaReportes


class VentanaPrincipal:
    def __init__(self, archivo_montos=config.ARCHIVO_MONTOS, archivos_anteriores=config.ARCHIVOS_MONTOS_ANTERIORES,
                 archivo_base=config.ARCHIVO_BASE, archivo_ajustes=config.ARCHIVO_AJUSTES):
        self.archivo_montos = archivo_montos
        self.archivo_base = archivo_base
        self.archivo_ajustes = archivo_ajustes
        self.montos_prestaciones = config.cargar_montos(archivo_montos, archivos_anteriores)
        self.ajustes = config.cargar_ajustes(archivo_ajustes)
        self.carpeta_guardado = config.carpeta_guardado(self.ajustes)

        self.root = tb.Window(title="MisLiquidaciones", themename=config.TEMAS[self.ajustes["tema"]])
        self.root.minsize(760, 560)
        self.root.option_add("*TButton.cursor", "hand2")
        self._cargar_icono()

        self.cuaderno = tb.Notebook(self.root)
        self.cuaderno.pack(expand=True, fill="both", padx=10, pady=10)
        self.configuracion = PestanaConfiguracion(self.cuaderno, self)
        self.liquidar = PestanaLiquidar(self.cuaderno, self)
        self.historial = PestanaHistorial(self.cuaderno, self)
        self.reportes = PestanaReportes(self.cuaderno, self)
        for pestaña, titulo in ((self.liquidar, "Liquidar"), (self.configuracion, "Configuración"),
                                (self.historial, "Historial"), (self.reportes, "Reportes")):
            self.cuaderno.add(pestaña.marco, text=f"  {titulo}  ")

        self.root.protocol("WM_DELETE_WINDOW", self.cerrar)
        self._centrar()

    def ejecutar(self):
        self.root.mainloop()

    def cerrar(self):
        self.configuracion.guardar_montos(avisar=False)
        self.root.destroy()

    def _cargar_icono(self):
        try:
            icon_path = config.obtener_directorio_recursos() / "icon.ico"
            if icon_path.exists():
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"No se pudo cargar el ícono: {e}")

    def _centrar(self):
        self.root.update_idletasks()
        ancho = max(self.root.winfo_reqwidth(), 860)
        alto = max(self.root.winfo_reqheight(), 640)
        x = (self.root.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.root.winfo_screenheight() // 2) - (alto // 2)
        self.root.geometry(f"{ancho}x{alto}+{max(x, 0)}+{max(y, 0)}")

    # --- Ajustes compartidos --- #

    def guardar_ajustes(self):
        try:
            config.guardar_ajustes(self.ajustes, self.archivo_ajustes)
        except OSError as e:
            messagebox.showwarning("Aviso", f"No se pudieron guardar los ajustes en {self.archivo_ajustes}\n{e}")

    def cambiar_tema(self, tema):
        self.ajustes["tema"] = tema
        self.root.style.theme_use(config.TEMAS[tema])
        self.guardar_ajustes()

    def cambiar_carpeta(self, carpeta):
        self.ajustes["carpeta_guardado"] = str(carpeta)
        self.carpeta_guardado = config.carpeta_guardado(self.ajustes)
        self.guardar_ajustes()
