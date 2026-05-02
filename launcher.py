import json
import os
import ssl
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from tkinter import Tk, Label, messagebox

import certifi


REPO_API_LATEST = "https://api.github.com/repos/Somnus-py/fashion-reset/releases/latest"
APP_EXE_NAME = "Fashion Reset.exe"
APP_ASSET_NAMES = {
    "Fashion Reset.exe",
    "fashion.reset.exe",
    "fashion-reset.exe",
    "fashion_reset.exe",
}
VERSION_FILE_NAME = "app_version.txt"
REQUEST_TIMEOUT = 30
RELEASE_REINTENTOS = 6
RELEASE_ESPERA_SEGUNDOS = 5


def obtener_carpeta_base():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CARPETA_BASE = obtener_carpeta_base()
APP_EXE_PATH = os.path.join(CARPETA_BASE, APP_EXE_NAME)
VERSION_FILE_PATH = os.path.join(CARPETA_BASE, VERSION_FILE_NAME)
SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where())


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
    with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT, context=SSL_CONTEXT) as response:
        return json.loads(response.read().decode("utf-8"))


def descargar_archivo(url, destino):
    request = urllib.request.Request(url, headers={"User-Agent": "FashionResetLauncher"})
    with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT, context=SSL_CONTEXT) as response:
        with open(destino, "wb") as archivo:
            archivo.write(response.read())


def normalizar_nombre_asset(nombre):
    return "".join(
        caracter.lower()
        for caracter in str(nombre or "")
        if caracter.isalnum()
    )


def es_asset_app(nombre_asset):
    nombre_normalizado = normalizar_nombre_asset(nombre_asset)
    nombres_app = {normalizar_nombre_asset(nombre) for nombre in APP_ASSET_NAMES}
    nombres_launcher = {
        normalizar_nombre_asset("Fashion Reset Launcher.exe"),
        normalizar_nombre_asset("fashion.reset.launcher.exe"),
        normalizar_nombre_asset("fashion-reset-launcher.exe"),
        normalizar_nombre_asset("fashion_reset_launcher.exe"),
    }

    return nombre_normalizado in nombres_app and nombre_normalizado not in nombres_launcher


def obtener_release_latest():
    release = pedir_json(REPO_API_LATEST)
    version = release.get("tag_name", "").strip()
    assets = release.get("assets", [])

    for asset in assets:
        nombre_asset = str(asset.get("name") or "").strip()
        if es_asset_app(nombre_asset):
            nombres_assets = [str(item.get("name") or "").strip() for item in assets]
            return version, asset.get("browser_download_url", ""), nombres_assets

    nombres_assets = [str(asset.get("name") or "").strip() for asset in assets]
    return version, "", nombres_assets


def obtener_release_latest_con_reintentos(actualizar=None):
    ultima_version = ""
    ultimos_assets = []

    for intento in range(1, RELEASE_REINTENTOS + 1):
        version, url_descarga, assets = obtener_release_latest()
        ultima_version = version
        ultimos_assets = assets

        if url_descarga:
            return version, url_descarga, assets

        if actualizar is not None:
            actualizar(f"Buscando ejecutable en GitHub... ({intento}/{RELEASE_REINTENTOS})")

        if intento < RELEASE_REINTENTOS:
            time.sleep(RELEASE_ESPERA_SEGUNDOS)

    return ultima_version, "", ultimos_assets


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
        ventana, label = crear_ventana_estado("Buscando actualizaciones...")
        version_remota, url_descarga, assets = obtener_release_latest_con_reintentos(
            lambda texto: actualizar_estado(ventana, label, texto)
        )

        necesita_descarga = (
            not os.path.exists(APP_EXE_PATH)
            or not version_local
            or version_local != version_remota
        )

        if necesita_descarga:
            if not version_remota or not url_descarga:
                assets_texto = ", ".join(assets) if assets else "sin archivos adjuntos"
                raise RuntimeError(
                    f"No se encontro {APP_EXE_NAME} en la ultima release ({version_remota or 'sin version'}). "
                    f"Assets encontrados: {assets_texto}."
                )

            actualizar_estado(ventana, label, "Descargando actualizacion...")
            fd, ruta_temporal = tempfile.mkstemp(
                suffix=".exe",
                prefix="fashion_reset_",
                dir=CARPETA_BASE
            )
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
