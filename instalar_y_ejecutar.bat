@echo off
echo ============================================
echo    Mail Blaster Institucional
echo ============================================
echo.

REM Buscar Python en el PATH primero
set PYTHON_CMD=
where python >nul 2>&1 && set PYTHON_CMD=python
if "%PYTHON_CMD%"=="" (
    where python3 >nul 2>&1 && set PYTHON_CMD=python3
)
if "%PYTHON_CMD%"=="" (
    where py >nul 2>&1 && set PYTHON_CMD=py
)

REM Buscar en rutas tipicas de instalacion local (sin PATH)
if "%PYTHON_CMD%"=="" (
    if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
        set PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python313\python.exe
    )
)
if "%PYTHON_CMD%"=="" (
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
        set PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python312\python.exe
    )
)
if "%PYTHON_CMD%"=="" (
    if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
        set PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
    )
)

if "%PYTHON_CMD%"=="" (
    echo ERROR: Python no encontrado.
    echo Instala Python 3.12+ desde https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH" durante la instalacion.
    pause
    exit /b 1
)

echo Python encontrado: %PYTHON_CMD%
echo.
echo Instalando/verificando dependencias...
"%PYTHON_CMD%" -m pip install -r requirements.txt --quiet
echo.
echo ============================================
echo    Ejecutando Mail Blaster...
echo ============================================
"%PYTHON_CMD%" main.py
pause

