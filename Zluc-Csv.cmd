@echo off
REM ====================================================================
REM  Spustac skriptu Zluc-Csv.ps1 dvojklikom.
REM  Musi lezat v rovnakom priecinku ako .ps1 subor.
REM ====================================================================

chcp 65001 >nul
title Zlucenie dodacich listov

set "SKRIPT=%~dp0Zluc-Csv.ps1"

if not exist "%SKRIPT%" (
    echo.
    echo   CHYBA: subor Zluc-Csv.ps1 sa nenasiel.
    echo   Ocakavany na: %SKRIPT%
    echo.
    pause
    exit /b 1
)

REM Ak je nainstalovany PowerShell 7, pouzije sa on, inak Windows PowerShell 5.1
where pwsh.exe >nul 2>&1
if %errorlevel%==0 (
    pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "%SKRIPT%"
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SKRIPT%"
)

echo.
echo ====================================================================
pause
