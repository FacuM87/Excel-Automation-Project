import base64
import ctypes
import mimetypes
import re
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from pathlib import Path

from .procesamiento import MESES, formatear_monto

SERVIDOR = "smtp.gmail.com"
PUERTO = 465
TIEMPO_ESPERA = 30
LIMITE_ADJUNTOS = 25 * 1024 * 1024
TIPO_EXCEL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
URL_CONTRASEÑA_APLICACION = "https://myaccount.google.com/apppasswords"

MARCADORES = ("periodo", "meses", "año", "cantidad_ordenes", "total")
ASUNTO_POR_DEFECTO = "Liquidación {periodo}"
CUERPO_POR_DEFECTO = "Hola,\n\nAdjunto la liquidación correspondiente a {periodo}.\n\nSaludos."
CORREO_POR_DEFECTO = {
    "cuenta": "", "contraseña_cifrada": "", "para": "", "cc": "",
    "asunto": ASUNTO_POR_DEFECTO, "cuerpo": CUERPO_POR_DEFECTO,
}

_PATRON_DIRECCION = re.compile(r"^[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+$")


class ErrorCorreo(Exception):
    pass


@dataclass
class PeriodoEnviado:
    """Un mes que va en el mail: su Excel y los datos para la plantilla."""
    mes: int
    año: int
    cantidad_ordenes: int
    total: float
    archivo: Path = None
    registrado: bool = True

    @property
    def nombre(self):
        return f"{MESES[self.mes - 1]} {self.año}"


# --- Plantilla --- #

def unir(partes):
    """['A', 'B', 'C'] -> 'A, B y C'."""
    partes = list(partes)
    if len(partes) <= 1:
        return "".join(partes)
    return ", ".join(partes[:-1]) + " y " + partes[-1]


def valores_marcadores(periodos):
    periodos = sorted(periodos, key=lambda p: (p.año, p.mes))
    años = sorted({p.año for p in periodos})
    meses = [MESES[p.mes - 1] for p in periodos]
    if len(años) == 1:
        periodo = f"{unir(meses)} {años[0]}"
        año = str(años[0])
    else:
        periodo = unir(p.nombre for p in periodos)
        año = f"{años[0]}-{años[-1]}"
    return {
        "periodo": periodo,
        "meses": unir(meses),
        "año": año,
        "cantidad_ordenes": str(sum(p.cantidad_ordenes for p in periodos)),
        "total": f"$ {formatear_monto(sum(p.total for p in periodos))}",
    }


def completar_plantilla(plantilla, periodos):
    """Reemplaza los marcadores conocidos; el resto del texto (y otras llaves) queda igual."""
    valores = valores_marcadores(periodos)
    for marcador, valor in valores.items():
        plantilla = plantilla.replace("{" + marcador + "}", valor)
    return plantilla


# --- Direcciones --- #

def separar_direcciones(texto):
    return [d.strip() for d in re.split(r"[,;]", texto or "") if d.strip()]


def direcciones_invalidas(direcciones):
    return [d for d in direcciones if not _PATRON_DIRECCION.match(d)]


# --- Mensaje y envío --- #

def armar_mensaje(remitente, para, cc, asunto, cuerpo, adjuntos=()):
    """Arma el mail. Lanza ErrorCorreo si falta un adjunto o se pasa del límite de Gmail."""
    adjuntos = [Path(a) for a in adjuntos]
    faltantes = [a.name for a in adjuntos if not a.is_file()]
    if faltantes:
        raise ErrorCorreo("No se encontraron estos archivos:\n" + "\n".join(f"  • {n}" for n in faltantes))
    tamaño = sum(a.stat().st_size for a in adjuntos)
    if tamaño > LIMITE_ADJUNTOS:
        raise ErrorCorreo(f"Los adjuntos ocupan {tamaño / 1024 / 1024:.1f} MB y Gmail acepta hasta 25 MB.")

    mensaje = EmailMessage()
    mensaje["From"] = remitente
    mensaje["To"] = ", ".join(para)
    if cc:
        mensaje["Cc"] = ", ".join(cc)
    mensaje["Subject"] = asunto
    mensaje["Date"] = formatdate(localtime=True)
    mensaje["Message-ID"] = make_msgid()
    mensaje.set_content(cuerpo)
    for adjunto in adjuntos:
        tipo = TIPO_EXCEL if adjunto.suffix.lower() == ".xlsx" else (
            mimetypes.guess_type(adjunto.name)[0] or "application/octet-stream")
        principal, secundario = tipo.split("/", 1)
        mensaje.add_attachment(adjunto.read_bytes(), maintype=principal, subtype=secundario, filename=adjunto.name)
    return mensaje


def enviar(mensaje, usuario, contraseña, conexion=smtplib.SMTP_SSL):
    """Envía el mail por el servidor de Gmail. Lanza ErrorCorreo con un mensaje claro si falla."""
    try:
        with conexion(SERVIDOR, PUERTO, timeout=TIEMPO_ESPERA) as servidor:
            servidor.login(usuario, contraseña.replace(" ", ""))
            servidor.send_message(mensaje)
    except smtplib.SMTPAuthenticationError:
        raise ErrorCorreo(
            "Gmail rechazó la cuenta o la contraseña.\n\n"
            "Hay que usar una contraseña de aplicación (16 letras), no la contraseña normal de Gmail. "
            "Se crea en la cuenta de Google y requiere tener activada la verificación en 2 pasos.")
    except smtplib.SMTPRecipientsRefused as e:
        raise ErrorCorreo("Gmail rechazó estos destinatarios:\n" + "\n".join(f"  • {d}" for d in e.recipients))
    except smtplib.SMTPException as e:
        raise ErrorCorreo(f"Gmail no aceptó el envío.\n{e}")
    except OSError as e:
        raise ErrorCorreo(f"No se pudo conectar con Gmail. Revisá la conexión a internet.\n{e}")


def mensaje_de_prueba(cuenta):
    return armar_mensaje(
        cuenta, [cuenta], [], "Prueba de MisLiquidaciones",
        "Este es un mail de prueba enviado desde MisLiquidaciones.\n\n"
        "Si lo recibiste, la cuenta y la contraseña de aplicación funcionan.")


# --- Contraseña (DPAPI de Windows) --- #

class _Blob(ctypes.Structure):
    _fields_ = [("cbData", ctypes.c_uint32), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(proteger, datos):
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    buffer = ctypes.create_string_buffer(datos, len(datos))
    entrada = _Blob(len(datos), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char)))
    salida = _Blob()
    funcion = crypt32.CryptProtectData if proteger else crypt32.CryptUnprotectData
    if not funcion(ctypes.byref(entrada), None, None, None, None, 0x1, ctypes.byref(salida)):
        raise ctypes.WinError()
    try:
        return ctypes.string_at(salida.pbData, salida.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(salida.pbData, ctypes.c_void_p))


def cifrar(texto):
    """Cifra con la cuenta de Windows del usuario; devuelve texto en base64 (vacío si no hay contraseña)."""
    if not texto:
        return ""
    return base64.b64encode(_dpapi(True, texto.encode("utf-8"))).decode("ascii")


def descifrar(texto):
    """Devuelve la contraseña, "" si no hay, o None si no se puede descifrar (otro usuario u otra PC)."""
    if not texto:
        return ""
    try:
        return _dpapi(False, base64.b64decode(texto)).decode("utf-8")
    except (OSError, ValueError, AttributeError):
        return None
