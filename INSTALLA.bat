@echo off
title RIORDINO EMAIL - installazione
cd /d "%~dp0"
where python >nul 2>&1
if errorlevel 1 (
  echo Python non trovato. Installarlo da https://www.python.org spuntando "Add python.exe to PATH", poi rilanciare.
  pause & exit /b 1
)
echo Installo le librerie...
python -m pip install --upgrade -r requisiti.txt
if errorlevel 1 (
  echo ERRORE durante l'installazione delle librerie.
  pause & exit /b 1
)
echo.
echo Collaudo offline (nessun accesso a Gmail)...
python -X utf8 prova_offline.py
echo.
echo Fatto. Prossimo passo: creare credentials.json (vedi LEGGIMI.txt) e usare AVVIA.bat.
pause
