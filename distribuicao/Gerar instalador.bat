@echo off
rem Gera o Autofiscal.exe e o instalador (veja gerar_instalador.py).
rem So o computador que GERA o instalador precisa de Python. Quem INSTALA, nao.
cd /d "%~dp0"
set PY=
python --version >nul 2>&1 && set PY=python
if not defined PY py --version >nul 2>&1 && set PY=py
if not defined PY goto sem_python
%PY% gerar_instalador.py %*
pause
exit /b

:sem_python
echo.
echo  Este computador nao tem Python, e o Python so e necessario para GERAR o instalador.
echo.
echo  Para INSTALAR o Autofiscal aqui, nao precisa de Python: copie para ca o arquivo
echo  Autofiscal-^<versao^>-instalador.exe, gerado no computador de desenvolvimento
echo  (pasta distribuicao\saida), e de dois cliques nele.
echo.
echo  Para gerar o instalador NESTE computador, instale antes:
echo    1. Python 3 (https://www.python.org/downloads/ - marque "Add python.exe to PATH")
echo    2. No Prompt: pip install pyinstaller openpyxl pypdf pillow pywin32 pywinauto
echo    3. Inno Setup 6 (https://jrsoftware.org/isinfo.php)
echo.
pause
exit /b 1
