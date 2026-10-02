@echo off
chcp 65001 >nul
cd /d "%~dp0"
python "preencher_chaves.py"
if errorlevel 1 pause
