@echo off
REM show widget
IF [%1]==[] (
    SET mode=widget
) else (
    SET mode=%1
)
powershell -EP Bypass -file %~dp0goodTimes.ps1 %mode% -l 1 -h 8 -b1 .25 -b2 .50 -p 60 -j 1 -m 10