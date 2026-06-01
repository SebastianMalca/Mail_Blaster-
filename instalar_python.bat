@echo off
echo ============================================
echo    Mail Blaster - Instalar Python 3.12
echo ============================================
echo.
echo Este script instalara Python 3.12 de forma silenciosa
echo con la opcion "Add to PATH" activada.
echo.
echo Buscando el instalador en Descargas...

set INSTALLER=%USERPROFILE%\Downloads\python-3.12.9-amd64.exe

if not exist "%INSTALLER%" (
    echo.
    echo ERROR: No se encontro el instalador en:
    echo   %INSTALLER%
    echo.
    echo Descargalo manualmente desde:
    echo   https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe
    echo.
    pause
    exit /b 1
)

echo Instalador encontrado. Instalando Python 3.12...
echo (Esto puede tardar unos minutos)
echo.

"%INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0

echo.
echo ============================================
echo   Python 3.12 instalado correctamente.
echo   Cierra y vuelve a abrir esta ventana
echo   para que el PATH se actualice.
echo ============================================
pause
