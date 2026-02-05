@echo off
REM =====================================================
REM SCRIPT DE COMPILACION - CONTROL DE PAGOS GCO v2.0
REM =====================================================
echo.
echo ========================================
echo   COMPILADOR DE CONTROL DE PAGOS GCO
echo   VERSION 2.0  
echo ========================================
echo.

REM Verificar que PyInstaller esté instalado
echo [1/5] Verificando PyInstaller...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo ERROR: PyInstaller no esta instalado.
    echo Instalando PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: No se pudo instalar PyInstaller.
        pause
        exit /b 1
    )
)
echo OK: PyInstaller instalado.
echo.

REM Limpiar compilaciones anteriores
echo [2/5] Limpiando archivos temporales...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ControlPagosGCO.spec del /q ControlPagosGCO.spec
echo OK: Archivos temporales eliminados.
echo.

REM Opción de compilación
echo [3/5] Seleccione el modo de compilacion:
echo.
echo 1. CON CONSOLA (recomendado para primera vez / debugging)
echo 2. SIN CONSOLA (version final, solo GUI)
echo.
set /p opcion="Ingrese opcion (1 o 2): "

if "%opcion%"=="1" (
    set modo=--console
    set nombre_modo=CON CONSOLA
) else (
    set modo=--noconsole
    set nombre_modo=SIN CONSOLA
)

echo.
echo [4/5] Compilando en modo: %nombre_modo%...
echo.

REM Compilar con PyInstaller
pyinstaller --onefile %modo% ^
    --icon=icon.ico ^
    --name="ControlPagosGCO" ^
    --add-data "icon.ico;." ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    --hidden-import=openpyxl ^
    --hidden-import=pandas ^
    --hidden-import=tkcalendar ^
    control_pagos_v1.py

if errorlevel 1 (
    echo.
    echo ERROR: La compilacion fallo.
    echo Revise los mensajes de error arriba.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   COMPILACION EXITOSA - VERSION 2.0
echo ========================================
echo.
echo El ejecutable se encuentra en:
echo %cd%\dist\ControlPagosGCO.exe
echo.
echo CAMBIOS EN VERSION 2.0:
echo - Selector de proceso (3 opciones)
echo - Validacion de archivo existente
echo - Mensajes personalizados
echo - Interfaz mejorada
echo.
echo [5/5] Abriendo carpeta de destino...
explorer dist

echo.
echo PROCESO COMPLETADO.
echo.
pause