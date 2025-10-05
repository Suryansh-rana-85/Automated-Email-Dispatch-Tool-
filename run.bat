@echo off
REM This batch file executes the PowerShell script using its full path.

echo Starting the email automation process...

REM The line below uses the full path to ensure PowerShell is found
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File "%~dp0send_email.ps1"

echo.
echo Process finished. Press any key to exit.
pause >nul