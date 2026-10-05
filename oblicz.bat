@echo off
chcp 65001 >nul
rem Przeciagnij plik ZESTAWIENIE (.xls / .xlsx) na ten plik, aby wykonac obliczenia.
if "%~1"=="" (
    echo Uzycie: przeciagnij plik zestawienia na oblicz.bat
    echo     lub: oblicz.bat "ZESTAWIENIE PRZYKLAD 1 obliczenia.xls"
    pause
    exit /b 1
)
python "%~dp0obliczenia.py" %*
pause
