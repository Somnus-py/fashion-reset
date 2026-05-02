@echo off
setlocal

cd /d "%~dp0"

echo [1/3] Actualizando dependencias...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo [2/3] Generando ejecutable...
python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --icon "Icono Fashion Reset.ico" ^
  --name "Fashion Reset" ^
  app_visual.py

echo [3/3] Copiando base de datos junto al ejecutable...
if exist "dist\Fashion Reset.exe" (
  copy /Y "fashion_reset.xlsx" "dist\fashion_reset.xlsx" >nul
  copy /Y "Icono Fashion Reset.ico" "dist\Icono Fashion Reset.ico" >nul
)

echo.
echo Listo. El ejecutable se genero en:
echo %~dp0dist
echo.
pause
