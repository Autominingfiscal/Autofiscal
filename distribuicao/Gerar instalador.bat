@echo off
rem Gera o Autofiscal.exe e o instalador (veja gerar_instalador.py).
cd /d "%~dp0"
python gerar_instalador.py %*
pause
