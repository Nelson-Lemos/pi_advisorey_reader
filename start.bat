@echo off
title PDF Tradutor Pro
echo.
echo  ========================================
echo   PDF Tradutor Pro v7.0 - A arrancar...
echo  ========================================
echo.
echo  Abrir no browser: http://127.0.0.1:8000
echo  Para parar: CTRL+C
echo.
py -m uvicorn main:app --host 0.0.0.0 --port 8000
pause