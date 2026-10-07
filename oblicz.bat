@echo off
chcp 65001 >nul
rem Uruchom dwuklikiem - otworzy sie okno wyboru pliku ZESTAWIENIE.
rem Mozna tez przeciagnac plik zestawienia (.xls / .xlsx) na ten plik.
python "%~dp0obliczenia.py" %*
if errorlevel 1 pause
if not "%~1"=="" pause
