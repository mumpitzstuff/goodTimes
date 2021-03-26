@echo off

:menu
cls
echo.
echo Menu
echo ====
echo.
echo [1] Install warning message (possible breaks are part of worktime (check worktime of runJoined.bat))
ping -n 1 localhost>nul
echo [2] Install warning message (breaks are not part of worktime (check worktime of runBreaks.bat))
ping -n 1 localhost>nul
echo [3] Uninstall
ping -n 1 localhost>nul
echo [4] Quit
ping -n 1 localhost>nul
echo.
set asw=0
set /p asw="Selection: "

if %asw%==1 goto Install1
if %asw%==2 goto Install2
if %asw%==3 goto Uninstall
if %asw%==4 goto END
goto END

:Install1
cls
echo.
runJoined.bat install
echo.
pause
goto END

:Install2
cls
echo.
runBreaks.bat install
echo.
pause
goto END

:Uninstall
cls
echo.
powershell -file .\goodTimes.ps1 uninstall
echo.
pause
goto END

:END
