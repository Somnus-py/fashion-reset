@echo off
setlocal

cd /d "%~dp0"

echo [1/2] Actualizando dependencias...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo [2/2] Generando launcher...
python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --icon "Icono Fashion Reset.ico" ^
  --name "Fashion Reset Launcher" ^
  launcher.py

echo.
echo Listo. El launcher se genero en:
echo %~dp0dist
echo.
pause
