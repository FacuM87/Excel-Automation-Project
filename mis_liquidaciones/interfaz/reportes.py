import ttkbootstrap as tb


class PestanaReportes:
    def __init__(self, padre, app):
        self.app = app
        self.marco = tb.Frame(padre, padding=30)
        tb.Label(self.marco, text="Reportes", font="-size 14 -weight bold").pack(pady=(40, 10))
        tb.Label(self.marco, text="Los reportes mensual y anual se agregan en la próxima versión.",
                 ).pack()
