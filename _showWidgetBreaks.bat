@echo off

@setlocal enableextensions
@cd /d "%~dp0"

REM show widget
IF [%1]==[] (
    SET mode=widget
) else (
    SET mode=%1
)
REM calling the script this way is needed to hide the console window
pwsh.exe -c "Start-Process -FilePath 'powershell.exe' -ArgumentList '-EP Bypass -NoProfile -NoLogo -file .\goodTimes.ps1 %mode% -l 1 -h 8 -b1 .25 -b2 .50 -p 60 -j 0 -m 10 -i 1' -WindowStyle Hidden"