@echo off
cd /d "%~dp0"
echo Verification de Python...
python --version
if errorlevel 1 (
  echo.
  echo Python n'est pas installe ou n'est pas dans le PATH.
  pause
  exit /b 1
)

echo.
echo Verification des modules...
python -c "import flask, pdfplumber, openpyxl; print('Modules OK')"
if errorlevel 1 (
  echo.
  echo Il manque des dependances. Lancez Installer_dependances.bat puis reessayez.
  pause
  exit /b 1
)

echo.
echo Demarrage du site sur http://127.0.0.1:5000
echo Gardez cette fenetre ouverte pendant l'utilisation.
python web_app.py
pause
