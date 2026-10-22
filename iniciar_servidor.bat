@echo off
title Servidor de Correos Masivos Corporativos
cd /d "%~dp0"

echo ======================================================================
echo  Iniciando Servidor de Correos Masivos Corporativos
echo ======================================================================
echo.

if exist .venv\Scripts\python.exe (
    set "PYTHON_CMD=.venv\Scripts\python.exe"
    echo [OK] Usando entorno virtual: .venv
) else (
    set "PYTHON_CMD=python"
    echo [AVISO] Usando Python del sistema
)

echo [OK] Abriendo navegador en http://localhost:5000...
start "" http://localhost:5000

echo.
"%PYTHON_CMD%" server.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] El servidor se detuvo con codigo %ERRORLEVEL%.
    pause
)
