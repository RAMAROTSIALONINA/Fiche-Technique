@echo off
title Création de l'installeur — Fiche Technique BOGOTA
echo.
echo  ============================================
echo   Création de l'installeur Windows (.exe)
echo  ============================================
echo.

cd /d "%~dp0"

:: Chercher Inno Setup
set ISCC=
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe
if exist "C:\Program Files\Inno Setup 6\ISCC.exe"       set ISCC=C:\Program Files\Inno Setup 6\ISCC.exe

if "%ISCC%"=="" (
    echo  ERREUR : Inno Setup n'est pas installe.
    echo.
    echo  Telechargez-le gratuitement ici :
    echo  https://jrsoftware.org/isdl.php
    echo.
    echo  Apres installation, relancez ce fichier.
    pause
    exit /b 1
)

:: Vérifier que le .exe existe
if not exist "dist\FicheTechnique-BOGOTA\FicheTechnique-BOGOTA.exe" (
    echo  ERREUR : Le .exe n'existe pas encore.
    echo  Lancez d'abord build.bat pour construire le .exe.
    pause
    exit /b 1
)

:: Créer le dossier de sortie
if not exist "installer-output" mkdir "installer-output"

echo  Compilation de l'installeur...
"%ISCC%" installer.iss

if errorlevel 1 (
    echo.
    echo  ERREUR lors de la compilation.
    pause
    exit /b 1
)

echo.
echo  ============================================
echo   SUCCES !
echo  ============================================
echo.
echo  L'installeur se trouve dans :
echo    installer-output\Setup-FicheTechnique-BOGOTA.exe
echo.
echo  Ce fichier peut etre copie et execute
echo  sur n'importe quel PC Windows.
echo.
start "" "installer-output"
pause
