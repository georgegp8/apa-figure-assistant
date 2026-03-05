@echo off
:: Activa el entorno virtual e inicia la aplicacion
call "%~dp0.venv\Scripts\activate.bat"
python "%~dp0main.py"
