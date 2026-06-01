@echo off
echo ============================================
echo   Mail Blaster - Generando ejecutable .exe
echo ============================================
echo.

REM Instalar PyInstaller si no está instalado
pip install pyinstaller

echo.
echo Compilando la aplicacion...
echo.

pyinstaller ^
    --onefile ^
    --windowed ^
    --name "MailBlaster" ^
    --add-data "app;app" ^
    --hidden-import customtkinter ^
    --hidden-import PIL._tkinter_finder ^
    --hidden-import openpyxl ^
    --hidden-import pandas ^
    --collect-all customtkinter ^
    main.py

echo.
echo ============================================
echo   Proceso completado.
echo   El ejecutable se encuentra en: dist\MailBlaster.exe
echo ============================================
pause
