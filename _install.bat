@echo off

:menu
cls
echo.
echo Menu
echo ====
echo.
echo [1] Install warning message (possible breaks are part of worktime (check worktime of runJoined.bat))
ping -n 1 localhost>nul
echo [2] Install widget (possible breaks are part of worktime (check worktime of runJoined.bat)) - admin rights needed
ping -n 1 localhost>nul
echo [3] Install warning message (breaks are NOT part of worktime (check worktime of runBreaks.bat))
ping -n 1 localhost>nul
echo [4] Install widget (breaks are NOT part of worktime (check worktime of runBreaks.bat)) - admin rights needed
ping -n 1 localhost>nul
echo [5] Uninstall warning message
ping -n 1 localhost>nul
echo [6] Uninstall widget - admin rights needed
ping -n 1 localhost>nul
echo [7] Quit
ping -n 1 localhost>nul
echo.
set asw=0
set /p asw="Selection: "

if %asw%==1 goto Install1
if %asw%==2 goto Install2
if %asw%==3 goto Install3
if %asw%==4 goto Install4
if %asw%==5 goto Uninstall1
if %asw%==6 goto Uninstall2
if %asw%==7 goto END
goto END

:Install1
cls
echo.
%~dp0_runJoined.bat install
echo.
pause
goto menu

:Install2
cls
echo.
%~dp0_showWidgetJoined.bat install_widget
echo.
pause
goto menu

:Install3
cls
echo.
%~dp0_runBreaks.bat install
echo.
pause
goto menu

:Install4
cls
echo.
%~dp0_showWidgetBreaks.bat install_widget
echo.
pause
goto menu

:Uninstall1
cls
echo.
powershell -EP Bypass -file %~dp0goodTimes.ps1 uninstall
echo.
pause
goto menu

:Uninstall2
cls
echo.
powershell -EP Bypass -file %~dp0goodTimes.ps1 uninstall_widget
echo.
pause
goto menu

:END
