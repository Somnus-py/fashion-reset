# README.md

## Fashion Reset

Fashion Reset es una aplicación de escritorio desarrollada en Python para gestionar ingresos, ventas, devoluciones, proveedoras, rendiciones y control de prendas de una tienda de ropa de segunda mano.

La aplicación utiliza Excel como base de datos principal y busca reemplazar procesos manuales con una interfaz simple, rápida y fácil de usar.

---

## Tecnologías utilizadas

* Python
* CustomTkinter
* OpenPyXL
* Excel

---

## Objetivos del proyecto

* Facilitar la carga de ingresos y ventas.
* Mantener control de prendas y proveedoras.
* Automatizar rendiciones.
* Tener una interfaz visual más cómoda que Excel.
* Mantener una estructura simple y fácil de editar.

---

## Estructura general

La aplicación funciona dentro de una sola ventana principal.

Desde el menú principal se accede a:

* Cargar ingreso
* Cargar venta
* Consultas y rendición
* Editar prenda
* Remarque
* Devolución
* Proveedoras
* Eliminar ingreso
* Reversar venta

Cada pantalla secundaria se abre dentro de la misma ventana y tiene un botón de volver.

---

## Diseño de pantallas

* Fondo blanco
* Texto negro
* Estilo minimalista
* Títulos compactos
* Botón de volver arriba a la izquierda
* Tablas tipo Excel para cargas masivas
* Agregar filas en bloques de 10
* Ignorar filas vacías al guardar

---

## Cómo ejecutar la app

Instalar dependencias:

```bash
pip install customtkinter openpyxl
```

Ejecutar:

```bash
python app_visual.py
```

---

## Archivos principales

* `app_visual.py` → interfaz visual
* `main.py` → lógica principal original
* `fashion_reset.xlsx` → base de datos
* `AGENTS.md` → reglas y preferencias del proyecto
