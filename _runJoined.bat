@echo off
REM show worktimes without breaks
powershell -file .\goodTimes.ps1 %1 -l 60 -h 8 -b .75 -p 60 -j 1 -m 10