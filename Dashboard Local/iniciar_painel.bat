@echo off
title Painel da expedicao (local)
cd /d "%~dp0"
where python >nul 2>&1
if %errorlevel%==0 (python painel_local.py) else (py painel_local.py)
pause
