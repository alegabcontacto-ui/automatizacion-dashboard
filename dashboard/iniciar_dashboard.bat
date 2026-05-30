@echo off
cd /d "%~dp0"
echo Iniciando Monitor de Cotizaciones IMSS...
echo.
echo Abre tu navegador en: http://localhost:8501
echo Presiona Ctrl+C para detener.
echo.
C:\Users\ALEJAN~1\AppData\Local\Python\PYTHON~1.14-\Scripts\STREAM~1.EXE run dashboard.py --server.headless false
pause
