@echo off
title Instalar Dependencias - RepartoPersonal
cd /d "%~dp0"
echo ====================================================
echo   INSTALANDO DEPENDENCIAS DE PYTHON (WINDOWS)
echo ====================================================
echo.

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] No se pudieron instalar las dependencias.
    echo Asegurate de tener Python instalado y anadido al PATH de Windows.
) else (
    echo.
    echo [OK] Dependencias instaladas correctamente.
    echo Ya puedes compilar ejecutando: compilar.bat
    echo o abrir la aplicacion con: ejecutar.bat
)

echo.
pause
