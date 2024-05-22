@echo off
setlocal

REM Define the PowerShell script file
set PowerShellScript=Scanner.ps1

REM Check if the PowerShell script exists
if not exist "%PowerShellScript%" (
    echo PowerShell script "%PowerShellScript%" not found.
    exit /b 1
)

REM Execute the PowerShell script
powershell -NoProfile -ExecutionPolicy Bypass -File "%PowerShellScript%"

REM Check the exit code of the PowerShell script
if %errorlevel% neq 0 (
    echo Failed to create Excel spreadsheet.
    exit /b %errorlevel%
)

echo Excel spreadsheet created successfully.

REM Add data input from USB port
echo Adding Machinery Connection from USB port...

endlocal
