@echo off
cd /d "%~dp0"
start "Serveur Controle PCB" cmd /k "cd /d "%~dp0" && python web_app.py"
timeout /t 3 /nobreak > nul
start "" "http://127.0.0.1:5000"
