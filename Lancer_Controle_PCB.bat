@echo off
cd /d "%~dp0"
python pcb.py
if errorlevel 1 pause
