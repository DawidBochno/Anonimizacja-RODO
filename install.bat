@echo off
title Instalacja - Anonimizacja RODO
cd /d "%~dp0"

set "PY="
py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY (
    python --version >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo BLAD: nie znaleziono Pythona.
    echo Zainstaluj Python 3.9+ z https://www.python.org/downloads/
    echo Podczas instalacji zaznacz "Add python.exe to PATH".
    pause
    exit /b 1
)
echo === Python ===
%PY% --version

echo.
echo === Instalacja bibliotek (PyMuPDF, python-docx) ===
%PY% -m pip install --upgrade pip
%PY% -m pip install --upgrade -r requirements.txt
if errorlevel 1 (
    echo BLAD instalacji bibliotek.
    pause
    exit /b 1
)

echo.
echo === Silnik OCR (Tesseract) - potrzebny tylko do skanow i zdjec ===
call :findtess
if not defined TESS (
    echo Nie znaleziono - instaluje przez winget ^(moze pojawic sie okno UAC^)...
    winget install --id UB-Mannheim.TesseractOCR -e --accept-package-agreements --accept-source-agreements
    call :findtess
)
if defined TESS (
    echo OK - %TESS%
) else (
    echo UWAGA: brak Tesseracta. DOCX i PDF z tekstem dzialaja, skany i zdjecia nie.
    echo Instalacja reczna: https://github.com/UB-Mannheim/tesseract/wiki
)

echo.
echo === Test ===
%PY% -c "import pymupdf, docx, tkinter; print('Biblioteki OK')"
if errorlevel 1 (
    echo BLAD: brak modulu - jesli chodzi o tkinter, zainstaluj Pythona
    echo ponownie z zaznaczona opcja "tcl/tk and IDLE".
    pause
    exit /b 1
)
%PY% anonimizacja.py --selftest
if errorlevel 1 (
    echo BLAD testu programu.
    pause
    exit /b 1
)

echo.
echo Gotowe. Program uruchamiasz plikiem uruchom.bat
pause
exit /b 0

:findtess
rem Instalacja bez praw administratora laduje w folderze uzytkownika, nie w Program Files.
set TESS=
for /f "delims=" %%p in ('where tesseract 2^>nul') do if not defined TESS set TESS=%%p
for %%p in (
    "C:\Program Files\Tesseract-OCR\tesseract.exe"
    "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
    "%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"
    "%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"
) do if not defined TESS if exist %%p set TESS=%%~p
exit /b 0
