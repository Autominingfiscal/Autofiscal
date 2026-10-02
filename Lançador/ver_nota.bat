@echo off
title Ver nota fiscal
cd /d "%~dp0"
if "%~1"=="" (echo Arraste um PDF de nota fiscal em cima deste arquivo. & pause & exit /b)
where python >nul 2>&1
if %errorlevel%==0 (python etapa3_lancador.py --ver-nota "%~1") else (py etapa3_lancador.py --ver-nota "%~1")
pause
