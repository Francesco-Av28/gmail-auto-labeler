@echo off
title RIORDINO EMAIL
cd /d "%~dp0"
:menu
echo.
echo  RIORDINO EMAIL - scegliere la fase (vedi LEGGIMI.txt)
echo   1  Login a Gmail (una volta sola)
echo   2  Estrai i metadati delle email (sola lettura)
echo   3  Calcola il piano - PROVA A SECCO, nessuna modifica
echo   4  Apri il piano in Excel
echo   5  ESEGUI il piano (modifica la casella, chiede conferma)
echo   6  Collaudo offline con casella finta
echo   0  Esci (anche solo INVIO)
set "scelta="
set /p scelta="Scelta: "
if "%scelta%"=="" exit /b 0
if "%scelta%"=="1" python -X utf8 riordino_email.py --login
if "%scelta%"=="2" python -X utf8 riordino_email.py --estrai
if "%scelta%"=="3" python -X utf8 riordino_email.py --piano
if "%scelta%"=="4" if exist piano_riordino.xlsx (start "" piano_riordino.xlsx) else (echo Prima calcolare il piano ^(3^).)
if "%scelta%"=="5" python -X utf8 riordino_email.py --esegui
if "%scelta%"=="6" python -X utf8 prova_offline.py
if "%scelta%"=="0" exit /b 0
goto menu
