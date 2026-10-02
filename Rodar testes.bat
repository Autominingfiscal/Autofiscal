@echo off
rem Roda os testes automaticos de todas as ferramentas.
cd /d "%~dp0"
python -m unittest discover -s tests -t tests -v
pause
