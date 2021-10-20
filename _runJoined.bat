@echo off
REM show worktimes without breaks
powershell -EP Bypass -file %~dp0goodTimes.ps1 %1 -l 60 -h 8 -b1 .25 -b2 .50 -p 60 -j 1 -m 10