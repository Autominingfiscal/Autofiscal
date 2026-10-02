@echo off
chcp 65001 >nul
cd /d "%~dp0"
python "rodar_tudo.py"
if errorlevel 1 pause
