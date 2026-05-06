# Protocolo De Actualizacion

## Objetivo

Actualizar la app sin poner en riesgo la base de datos de uso real.

La computadora de desarrollo es esta.

La computadora de uso es la de tu esposa.

## Regla principal

Nunca hacer cambios de codigo directamente en la computadora de uso.

Todo cambio se hace primero en la computadora de desarrollo.

## Flujo correcto de actualizacion

1. Hacer los cambios en el proyecto.
2. Probar que la app siga funcionando.
3. Generar un nuevo `Fashion Reset.exe`.
4. Preparar una carpeta de entrega nueva.
5. En la computadora de uso, cerrar la app.
6. Hacer copia de seguridad de la carpeta completa antes de actualizar.
7. Reemplazar el `Fashion Reset.exe` por la nueva version.
8. Mantener el `fashion_reset.xlsx` existente, salvo que haya una migracion planificada.
9. Abrir la app y probar una operacion simple.

## Flujo con launcher y GitHub

Cuando la actualizacion se entrega por GitHub:

1. Hacer commit de los cambios.
2. Hacer push a `main`.
3. Crear un tag nuevo `vX.X.X`.
4. Hacer push del tag.
5. Esperar a que GitHub Actions genere la release.
6. Ejecutar el launcher en la computadora de uso para descargar la version nueva.

Un push normal a `main` no genera ejecutables descargables. La release se genera con tags `v*`.

Si cambia `launcher.py`, hay que pasar manualmente el launcher nuevo al menos una vez. Un launcher viejo no puede corregir errores propios si falla antes de descargar o instalar la actualizacion.

El launcher debe descargar el archivo temporal en la misma carpeta donde esta la app. Esto evita errores de Windows al actualizar desde pendrive u otra unidad, por ejemplo cuando intenta mover un archivo desde `C:` hacia `F:`.

## Cuando solo cambia la interfaz o logica visual

Si los cambios son de estetica, orden, botones, textos o comportamiento de pantalla:

- reemplazar solo `Fashion Reset.exe`
- no reemplazar `fashion_reset.xlsx`

## Cuando cambia la logica de datos

Si los cambios afectan:

- columnas de Excel
- nombres de hojas
- formato de fechas
- guardado de ventas, ingresos, devoluciones o rendiciones
- calculos que dependen de datos historicos

entonces no actualizar de forma directa.

Primero hay que revisar compatibilidad con el Excel actual.

## Antes de cada actualizacion

Confirmar estas 4 cosas:

1. La app abre.
2. Se puede cargar un ingreso.
3. Se puede cargar una venta.
4. Se puede leer el Excel existente sin error.

## Copia de seguridad recomendada

Antes de actualizar en la computadora de uso:

1. Copiar la carpeta completa actual.
2. Guardarla con fecha.

Ejemplo:

`Fashion Reset - backup 2026-04-22`

## Si la actualizacion falla

1. Cerrar la app nueva.
2. Restaurar la carpeta de respaldo.
3. Volver a usar la version anterior.
4. Revisar el problema en la computadora de desarrollo.

## Regla sobre el Excel

El archivo mas sensible no es el `.exe`.

Es `fashion_reset.xlsx`.

El ejecutable se puede reemplazar.

La base de datos no se debe sobrescribir sin motivo.

## Versiones futuras

Cada vez que prepares una actualizacion:

- guardar una copia del `.exe` anterior
- anotar que cambio se hizo
- indicar si la actualizacion requiere o no cambios en el Excel

## Formato sugerido de registro

`Version`

`Fecha`

`Cambios`

`Requiere cambio en Excel: SI/NO`

`Observaciones`
