@echo off
title Compilador RepartoPersonal
cd /d "%~dp0"
echo ====================================================
echo   COMPILANDO REPARTO PERSONAL PARA WINDOWS (.EXE)
echo ====================================================
echo.

python compilar_exe.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ====================================================
    echo [ERROR] La compilacion ha fallado.
    echo Si te falta alguna libreria, ejecuta primero:
    echo   instalar_dependencias.bat
    echo ====================================================
) else (
    echo.
    echo ====================================================
    echo [OK] Compilacion completada con exito.
    echo Tu aplicacion portable esta en: dist\RepartoPersonal\
    echo ====================================================
)

echo.
pause
