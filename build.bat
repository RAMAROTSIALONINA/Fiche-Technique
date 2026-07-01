@echo off
title Build FicheTechnique BOGOTA - .exe
echo.
echo  ============================================
echo   Build FicheTechnique BOGOTA en .exe
echo  ============================================
echo.

set PYTHON=C:\Users\User\AppData\Local\Programs\Python\Python314\python.exe
set PYTHONPATH=D:\DSI\BOURBON MORELLI\Lib\site-packages
set SITE=%PYTHONPATH%

cd /d "%~dp0"

echo [1/3] Installation de PyInstaller...
"%PYTHON%" -m pip install pyinstaller --target="%SITE%" -q
if errorlevel 1 (
    echo ERREUR : impossible d'installer PyInstaller.
    pause & exit /b 1
)

echo [2/3] Nettoyage des anciens builds...
if exist build  rmdir /s /q build
if exist dist   rmdir /s /q dist
if exist "FicheTechnique-BOGOTA.spec" del /q "FicheTechnique-BOGOTA.spec"

echo [3/3] Construction du .exe...
"%PYTHON%" -m PyInstaller -y ^
  --onedir ^
  --noconsole ^
  --name "FicheTechnique-BOGOTA" ^
  --icon "static\logo.ico" ^
  --add-data "static;static" ^
  --hidden-import uvicorn.logging ^
  --hidden-import uvicorn.loops ^
  --hidden-import uvicorn.loops.auto ^
  --hidden-import uvicorn.protocols ^
  --hidden-import uvicorn.protocols.http ^
  --hidden-import uvicorn.protocols.http.auto ^
  --hidden-import uvicorn.protocols.websockets ^
  --hidden-import uvicorn.protocols.websockets.auto ^
  --hidden-import uvicorn.lifespan ^
  --hidden-import uvicorn.lifespan.on ^
  --hidden-import sqlalchemy.dialects.sqlite ^
  --hidden-import multipart ^
  --collect-all fitz ^
  main.py

if errorlevel 1 (
    echo.
    echo  ERREUR lors de la construction.
    pause & exit /b 1
)

echo.
echo  ============================================
echo   SUCCES !
echo  ============================================
echo.
echo  Le dossier de l'application se trouve dans :
echo    dist\FicheTechnique-BOGOTA\
echo.
echo  Pour lancer l'application :
echo    Double-cliquez sur FicheTechnique-BOGOTA.exe
echo    dans le dossier dist\FicheTechnique-BOGOTA\
echo.
echo  La base de donnees (fiches_bogota.db) sera
echo  creee automatiquement a cote du .exe.
echo.
pause
