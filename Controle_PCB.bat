@echo off
cd /d "%~dp0"
title Controle Qualite PCB

echo Demarrage du controle qualite PCB...
echo.

python --version >nul 2>&1
if errorlevel 1 (
  echo Python n'est pas installe ou n'est pas accessible.
  echo Installez Python depuis https://www.python.org/downloads/
  echo Pendant l'installation, cochez "Add python.exe to PATH".
  pause
  exit /b 1
)

python -c "import flask, pdfplumber, openpyxl, werkzeug" >nul 2>&1
if errorlevel 1 (
  echo Installation automatique des modules necessaires...
  python -m pip install --upgrade pip
  python -m pip install pdfplumber openpyxl flask werkzeug
  if errorlevel 1 (
    echo.
    echo Installation impossible. Verifiez la connexion internet ou les droits du poste.
    pause
    exit /b 1
  )
)

echo.
echo Le site va s'ouvrir automatiquement.
echo Gardez cette fenetre ouverte pendant l'utilisation.
echo.
python web_app.py
pause
