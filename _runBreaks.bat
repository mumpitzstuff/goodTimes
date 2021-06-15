@echo off
REM show worktimes with breaks
powershell -file .\goodTimes.ps1 %1 -l 60 -h 8 -b1 .25 -b2 .50 -p 60 -j 0 -m 10 -i 1