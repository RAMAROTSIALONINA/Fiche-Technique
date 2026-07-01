@echo off
title Fiche Technique BOGOTA
echo.
echo  ===================================
echo   Fiche Technique BOGOTA - Demarrage
echo  ===================================
echo.
echo  Lancement du serveur...
echo  Ouverture dans le navigateur...
echo.

set PYTHON=C:\Users\User\AppData\Local\Programs\Python\Python314\python.exe
set PYTHONPATH=D:\DSI\BOURBON MORELLI\Lib\site-packages

"%PYTHON%" "%~dp0main.py"
echo.
echo  Le serveur s'est arrete (code: %errorlevel%)
pause
