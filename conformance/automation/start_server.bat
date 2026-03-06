@echo off
setlocal enabledelayedexpansion

set "CONF=%~dp0server.conf"
if not exist "%CONF%" (
    echo Error: server.conf not found.
    echo Create conformance\automation\server.conf with:
    echo   OTSOFT_EXE=C:\path\to\OTSoft.exe
    pause
    exit /b 1
)

for /f "usebackq tokens=1,* delims==" %%a in ("%CONF%") do (
    if "%%a"=="OTSOFT_EXE" set "OTSOFT_EXE=%%b"
)

if not defined OTSOFT_EXE (
    echo Error: OTSOFT_EXE not set in server.conf
    pause
    exit /b 1
)

echo Starting conformance test server...
echo OTSoft: %OTSOFT_EXE%
echo.
python "%~dp0server.py" --otsoft-path "%OTSOFT_EXE%"
pause
