@echo off
cd /d "%~dp0"
where pythonw >nul 2>&1
if %errorlevel%==0 (start "" pythonw "Autofiscal.pyw" & exit /b)
where pyw >nul 2>&1
if %errorlevel%==0 (start "" pyw "Autofiscal.pyw" & exit /b)
python "Autofiscal.pyw"
if errorlevel 1 pause
