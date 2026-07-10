@echo off
cd /d "%~dp0"
python -m pip install pyinstaller pdfplumber openpyxl flask werkzeug pymupdf
python -m PyInstaller --onefile --name Controle_Qualite_PCB_Web --add-data "templates;templates" --add-data "static;static" web_app.py
echo.
echo Application creee dans le dossier dist:
echo dist\Controle_Qualite_PCB_Web.exe
pause
