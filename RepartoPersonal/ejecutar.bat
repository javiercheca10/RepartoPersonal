@echo off
title RepartoPersonal
cd /d "%~dp0"
python iniciar.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] No se pudo iniciar la aplicacion.
    echo Asegurate de tener instaladas las dependencias ejecutando:
    echo   instalar_dependencias.bat
    echo.
    pause
)
