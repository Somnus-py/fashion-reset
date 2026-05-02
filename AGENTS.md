# AGENTS.md

## Objetivo del proyecto

Fashion Reset es una app de escritorio en Python para gestionar ingresos, ventas, devoluciones, proveedoras y rendiciones de una tienda de ropa de segunda mano.

La aplicación debe ser simple, rápida, visualmente limpia y pensada para uso real diario.

---

## Tecnologías

* Python
* CustomTkinter
* OpenPyXL
* Excel como base de datos principal

---

## Filosofía de trabajo

* Trabajar siempre paso a paso.
* No reescribir archivos completos si no es necesario.
* Hacer cambios pequeños y localizados.
* Debatir primero la lógica y el diseño antes de programar.
* Mantener el código ordenado y fácil de editar.
* Evitar soluciones excesivamente complejas.
* Priorizar siempre la practicidad y velocidad de uso.

---

## Diseño visual

* Una sola ventana principal.
* Fondo blanco.
* Texto negro.
* Estilo minimalista.
* Mantener misma tipografía, colores y botones en toda la app.
* Evitar pantallas recargadas.
* Aprovechar el espacio vertical.
* Títulos compactos.
* Flecha de volver siempre arriba a la izquierda.
* Evitar espacios excesivos entre encabezados y contenido.
* Mantener coherencia visual entre todas las pantallas.

---

## Regla para pantallas secundarias

Las pantallas como:

* Cargar ingreso
* Cargar venta
* Devolución
* Editar prenda
* Remarque
* Proveedoras

Deben abrirse dentro de la misma ventana principal, ocultando el menú principal y mostrando un frame nuevo.

---

## Regla para formularios

Cuando haya cargas múltiples:

* No usar formularios largos verticales.
* Usar tablas tipo Excel.
* Fecha y proveedora arriba como campos fijos del lote.
* Debajo, una tabla editable.
* Cada fila representa una prenda.
* Permitir agregar filas dinámicamente.
* Agregar filas en bloques de 10.
* Ignorar filas vacías al guardar.
* Mantener mismo diseño en ingresos, ventas y otras cargas masivas.

---

## Botones inferiores

En pantallas de carga:

* Botón "Agregar 10 filas más"
* Botón "Guardar lote"

El botón de volver queda arriba.

---

## Criterios de programación

* Reutilizar funciones cuando sea posible.
* Evitar repetir bloques largos de código.
* Crear funciones auxiliares para agregar filas, limpiar campos, guardar datos, etc.
* Usar nombres de variables claros.
* Mantener indentación limpia.
* No mezclar lógica visual con lógica de Excel más de lo necesario.
* Separar la construcción visual de la lógica de guardado.

---

## Prioridades del proyecto

1. Que funcione.
2. Que sea fácil de usar.
3. Que sea fácil de editar.
4. Que se vea prolijo.
5. Que soporte crecimiento futuro.
