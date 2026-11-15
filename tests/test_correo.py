import smtplib
import sys
from email import policy
from email.parser import BytesParser

import pytest

from mis_liquidaciones import correo
from mis_liquidaciones.correo import (
    TIPO_EXCEL, ErrorCorreo, PeriodoEnviado, armar_mensaje, completar_plantilla, direcciones_invalidas, enviar,
    separar_direcciones, valores_marcadores,
)


def periodo(mes, año=2026, cantidad=10, total=1000.0, archivo=None):
    return PeriodoEnviado(mes, año, cantidad, total, archivo)


# --- Plantilla --- #

def test_marcadores_de_un_mes():
    valores = valores_marcadores([periodo(3, cantidad=38, total=152000)])
    assert valores == {"periodo": "Marzo 2026", "meses": "Marzo", "año": "2026",
                       "cantidad_ordenes": "38", "total": "$ 152.000"}


def test_marcadores_de_varios_meses_del_mismo_año_en_orden():
    valores = valores_marcadores([periodo(5), periodo(3), periodo(4, total=500.5)])
    assert valores["periodo"] == "Marzo, Abril y Mayo 2026"
    assert valores["meses"] == "Marzo, Abril y Mayo"
    assert valores["año"] == "2026"
    assert valores["cantidad_ordenes"] == "30"
    assert valores["total"] == "$ 2.500,50"


def test_marcadores_de_años_distintos():
    valores = valores_marcadores([periodo(1, 2026), periodo(12, 2025)])
    assert valores["periodo"] == "Diciembre 2025 y Enero 2026"
    assert valores["meses"] == "Diciembre y Enero"
    assert valores["año"] == "2025-2026"


def test_completar_plantilla_deja_el_resto_igual():
    texto = "Liquidación {periodo} ({cantidad_ordenes} órdenes, {total}) {desconocido} {"
    assert completar_plantilla(texto, [periodo(3), periodo(4)]) == \
        "Liquidación Marzo y Abril 2026 (20 órdenes, $ 2.000) {desconocido} {"


# --- Direcciones --- #

def test_separar_y_validar_direcciones():
    direcciones = separar_direcciones(" a@x.com, b@y.com.ar ;c@z.org,, ")
    assert direcciones == ["a@x.com", "b@y.com.ar", "c@z.org"]
    assert direcciones_invalidas(direcciones) == []
    assert direcciones_invalidas(["sin-arroba.com", "a@b", "a b@c.com", "ok@x.com"]) == \
        ["sin-arroba.com", "a@b", "a b@c.com"]
    assert separar_direcciones("") == []


# --- Mensaje --- #

def test_armar_mensaje_con_adjuntos(tmp_path):
    marzo = tmp_path / "Liquidación Marzo 2026.xlsx"
    abril = tmp_path / "Liquidación Abril 2026.xlsx"
    marzo.write_bytes(b"excel de marzo")
    abril.write_bytes(b"excel de abril")
    mensaje = armar_mensaje("yo@gmail.com", ["a@x.com", "b@y.com"], ["c@z.com"], "Liquidación Marzo y Abril 2026",
                            "Hola,\nAdjunto.", [marzo, abril])
    leido = BytesParser(policy=policy.default).parsebytes(mensaje.as_bytes())
    assert leido["From"] == "yo@gmail.com"
    assert leido["To"] == "a@x.com, b@y.com"
    assert leido["Cc"] == "c@z.com"
    assert leido["Subject"] == "Liquidación Marzo y Abril 2026"
    assert leido.get_body(("plain",)).get_content().strip() == "Hola,\nAdjunto."
    adjuntos = list(leido.iter_attachments())
    assert [a.get_filename() for a in adjuntos] == [marzo.name, abril.name]
    assert all(a.get_content_type() == TIPO_EXCEL for a in adjuntos)
    assert adjuntos[0].get_content() == b"excel de marzo"


def test_armar_mensaje_sin_cc_ni_adjuntos():
    mensaje = armar_mensaje("yo@gmail.com", ["a@x.com"], [], "Asunto", "Cuerpo")
    assert mensaje["Cc"] is None
    assert list(mensaje.iter_attachments()) == []


def test_armar_mensaje_con_adjunto_faltante(tmp_path):
    with pytest.raises(ErrorCorreo, match="no-existe.xlsx"):
        armar_mensaje("yo@gmail.com", ["a@x.com"], [], "A", "B", [tmp_path / "no-existe.xlsx"])


def test_armar_mensaje_supera_el_limite(tmp_path, monkeypatch):
    archivo = tmp_path / "grande.xlsx"
    archivo.write_bytes(b"x" * 200)
    monkeypatch.setattr(correo, "LIMITE_ADJUNTOS", 100)
    with pytest.raises(ErrorCorreo, match="25 MB"):
        armar_mensaje("yo@gmail.com", ["a@x.com"], [], "A", "B", [archivo])


# --- Envío --- #

class ServidorFalso:
    """Reemplaza a smtplib.SMTP_SSL: registra lo que se hizo o lanza el error pedido."""
    instancias = []

    def __init__(self, servidor, puerto, timeout=None, error_login=None, error_envio=None):
        self.datos = {"servidor": servidor, "puerto": puerto, "timeout": timeout}
        self.error_login = error_login
        self.error_envio = error_envio
        ServidorFalso.instancias.append(self)

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.datos["cerrado"] = True

    def login(self, usuario, contraseña):
        if self.error_login:
            raise self.error_login
        self.datos["login"] = (usuario, contraseña)

    def send_message(self, mensaje):
        if self.error_envio:
            raise self.error_envio
        self.datos["mensaje"] = mensaje


def fabrica(**errores):
    ServidorFalso.instancias = []
    return lambda servidor, puerto, timeout=None: ServidorFalso(servidor, puerto, timeout, **errores)


def test_enviar_usa_gmail_y_la_cuenta():
    mensaje = armar_mensaje("yo@gmail.com", ["a@x.com"], [], "A", "B")
    enviar(mensaje, "yo@gmail.com", "abcd efgh ijkl mnop", fabrica())
    datos = ServidorFalso.instancias[0].datos
    assert (datos["servidor"], datos["puerto"]) == ("smtp.gmail.com", 465)
    assert datos["timeout"] == correo.TIEMPO_ESPERA
    assert datos["login"] == ("yo@gmail.com", "abcdefghijklmnop")
    assert datos["mensaje"] is mensaje
    assert datos["cerrado"]


@pytest.mark.parametrize("errores, texto", [
    ({"error_login": smtplib.SMTPAuthenticationError(535, b"BadCredentials")}, "contraseña de aplicación"),
    ({"error_envio": smtplib.SMTPRecipientsRefused({"malo@x.com": (550, b"no")})}, "malo@x.com"),
    ({"error_envio": smtplib.SMTPDataError(552, b"too big")}, "no aceptó"),
    ({"error_login": ConnectionRefusedError("sin red")}, "conexión a internet"),
])
def test_enviar_traduce_los_errores(errores, texto):
    mensaje = armar_mensaje("yo@gmail.com", ["a@x.com"], [], "A", "B")
    with pytest.raises(ErrorCorreo, match=texto):
        enviar(mensaje, "yo@gmail.com", "clave", fabrica(**errores))


def test_mensaje_de_prueba_va_a_la_misma_cuenta():
    mensaje = correo.mensaje_de_prueba("yo@gmail.com")
    assert mensaje["From"] == mensaje["To"] == "yo@gmail.com"


# --- Contraseña --- #

@pytest.mark.skipif(sys.platform != "win32", reason="DPAPI solo existe en Windows")
def test_cifrar_y_descifrar():
    cifrada = correo.cifrar("abcd efgh ijkl mnop")
    assert cifrada and "abcd" not in cifrada
    assert correo.descifrar(cifrada) == "abcd efgh ijkl mnop"
    assert correo.cifrar("") == "" and correo.descifrar("") == ""


@pytest.mark.skipif(sys.platform != "win32", reason="DPAPI solo existe en Windows")
def test_descifrar_algo_invalido_devuelve_none():
    assert correo.descifrar("no es base64 válido!") is None
    assert correo.descifrar("aG9sYQ==") is None
