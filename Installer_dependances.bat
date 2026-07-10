@echo off
cd /d "%~dp0"
python -m pip install --upgrade pip
python -m pip install pdfplumber openpyxl flask werkzeug pymupdf
pause
