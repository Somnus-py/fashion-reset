import json
import os
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from tkinter import Tk, Label, messagebox


REPO_API_LATEST = "https://api.github.com/repos/Somnus-py/fashion-reset/releases/latest"
APP_EXE_NAME = "Fashion Reset.exe"
VERSION_FILE_NAME = "app_version.txt"
REQUEST_TIMEOUT = 30


def obtener_carpeta_base():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CARPETA_BASE = obtener_carpeta_base()
APP_EXE_PATH = os.path.join(CARPETA_BASE, APP_EXE_NAME)
VERSION_FILE_PATH = os.path.join(CARPETA_BASE, VERSION_FILE_NAME)


def leer_texto(ruta):
    try:
        with open(ruta, "r", encoding="utf-8") as archivo:
            return archivo.read().strip()
    except FileNotFoundError:
        return ""


def guardar_texto(ruta, contenido):
    with open(ruta, "w", encoding="utf-8") as archivo:
        archivo.write(contenido)


def pedir_json(url):
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "FashionResetLauncher",
        },
    )
    with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT) as response:
        return json.loads(response.read().decode("utf-8"))


def descargar_archivo(url, destino):
    request = urllib.request.Request(url, headers={"User-Agent": "FashionResetLauncher"})
    with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT) as response:
        with open(destino, "wb") as archivo:
            archivo.write(response.read())


def obtener_release_latest():
    release = pedir_json(REPO_API_LATEST)
    version = release.get("tag_name", "").strip()
    assets = release.get("assets", [])

    for asset in assets:
        if asset.get("name") == APP_EXE_NAME:
            return version, asset.get("browser_download_url", "")

    return version, ""


def crear_ventana_estado(texto):
    ventana = Tk()
    ventana.title("Fashion Reset")
    ventana.geometry("420x120")
    ventana.resizable(False, False)

    label = Label(ventana, text=texto, font=("Arial", 11), padx=20, pady=30)
    label.pack(fill="both", expand=True)
    ventana.update()
    return ventana, label


def actualizar_estado(ventana, label, texto):
    label.configure(text=texto)
    ventana.update()


def abrir_app():
    subprocess.Popen([APP_EXE_PATH], cwd=CARPETA_BASE)


def mostrar_error(mensaje):
    root = Tk()
    root.withdraw()
    messagebox.showerror("Fashion Reset", mensaje)
    root.destroy()


def main():
    ventana = None
    label = None

    try:
        version_local = leer_texto(VERSION_FILE_PATH)
        version_remota, url_descarga = obtener_release_latest()

        necesita_descarga = (
            not os.path.exists(APP_EXE_PATH)
            or not version_local
            or version_local != version_remota
        )

        if necesita_descarga:
            if not version_remota or not url_descarga:
                raise RuntimeError("No se encontro el ejecutable en la ultima release.")

            ventana, label = crear_ventana_estado("Descargando actualizacion...")

            fd, ruta_temporal = tempfile.mkstemp(suffix=".exe", prefix="fashion_reset_")
            os.close(fd)

            try:
                descargar_archivo(url_descarga, ruta_temporal)
                actualizar_estado(ventana, label, "Instalando actualizacion...")
                os.replace(ruta_temporal, APP_EXE_PATH)
                guardar_texto(VERSION_FILE_PATH, version_remota)
            finally:
                if os.path.exists(ruta_temporal):
                    os.remove(ruta_temporal)

        if ventana is not None:
            actualizar_estado(ventana, label, "Abriendo Fashion Reset...")
            ventana.destroy()

        abrir_app()

    except (urllib.error.URLError, TimeoutError, RuntimeError, OSError) as error:
        if ventana is not None:
            ventana.destroy()

        if os.path.exists(APP_EXE_PATH):
            abrir_app()
            return

        mostrar_error(
            "No se pudo descargar Fashion Reset.\n\n"
            "Revisa la conexion a internet e intenta de nuevo.\n\n"
            f"Detalle: {error}"
        )


if __name__ == "__main__":
    main()
