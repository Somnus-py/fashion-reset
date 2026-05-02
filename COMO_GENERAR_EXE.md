# Como Generar El EXE

## Requisitos

Instalar Python en la computadora donde vayas a generar el ejecutable.

## Pasos

1. Abrir la carpeta del proyecto.
2. Ejecutar `build_exe.bat`.

Ese script hace esto:

1. Actualiza `pip`.
2. Instala dependencias desde `requirements.txt`.
3. Genera `Fashion Reset.exe` con PyInstaller.
4. Copia `fashion_reset.xlsx` al lado del ejecutable.

## Resultado

La salida queda en:

```text
dist/
  Fashion Reset.exe
  fashion_reset.xlsx
```

## Como llevarlo a otra computadora

Copiar la carpeta `dist` completa a la computadora de destino.

La app va a usar:

* el ejecutable `Fashion Reset.exe`
* el Excel local `fashion_reset.xlsx`

Ambos deben quedar juntos en la misma carpeta.

## Importante

La app ya fue adaptada para buscar:

* `fashion_reset.xlsx`
* `Rendiciones de Cuenta`
* `Remarques Enviados`

siempre relativos a la carpeta del ejecutable.

O sea: no importa si esa carpeta está en otro escritorio, en otra cuenta de Windows o en otra ruta.
