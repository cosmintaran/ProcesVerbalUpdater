@echo off
setlocal

REM Seteaza datele de input

REM directorul sursa
set "sourceDirectory=c:\Users\cosmi\source\repos\PvUpdater"
REM Directorul unde se vor salva noile documente
set "destinationDirectory=c:\Users\cosmi\source\repos\PvUpdater\Output"
REM Numarul de start
set "startNumber=1"
REM Data procesului verval
set "date=02/02/2024"

REM Setează calea către fișierul executabil
set "executable=.\ProcesVerbalUpdater\ProcesVerbal.exe"

REM Verifică dacă fișierul executabil există
if not exist "%executable%" (
    echo Fișierul executabil nu există: %executable%
    exit /b 1
)

REM Rulează fișierul executabil cu parametrii
"%executable%" -n %startNumber% -d %date% -p "%sourceDirectory%" -o "%destinationDirectory%"

exit /b 0


