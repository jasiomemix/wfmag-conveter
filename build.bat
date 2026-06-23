@echo off
chcp 65001 >nul
echo ============================================
echo   Budowanie WF-Mag Konwerter (.exe)
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [BLAD] Python nie zostal znaleziony w PATH.
    echo Zainstaluj Python 3.10+ z https://www.python.org/downloads/
    echo Zaznacz "Add Python to PATH" podczas instalacji!
    pause
    exit /b 1
)

echo [1/3] Instalowanie zaleznosci...
pip install --upgrade flask openpyxl pywebview pyinstaller
if errorlevel 1 (
    echo [BLAD] Nie udalo sie zainstalowac zaleznosci.
    pause
    exit /b 1
)

echo.
echo [2/3] Budowanie pliku EXE...
pyinstaller --noconfirm wfmag-converter.spec
if errorlevel 1 (
    echo [BLAD] Budowanie nie powiodlo sie.
    pause
    exit /b 1
)

echo.
echo [3/3] Gotowe!
echo.
echo Plik EXE znajduje sie w: dist\WFMag-Konwerter.exe
echo Mozesz go skopiowac i uruchomic na dowolnym komputerze.
echo.
pause
