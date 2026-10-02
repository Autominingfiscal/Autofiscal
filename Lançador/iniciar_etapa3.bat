@echo off
title Lancador - etapas 1, 2 e 3
cd /d "%~dp0"
where python >nul 2>&1
if %errorlevel%==0 (python etapa3_lancador.py) else (py etapa3_lancador.py)
pause
