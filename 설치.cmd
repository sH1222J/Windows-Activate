@echo off
title Windows 최적화 업데이트
echo Please Wait ...
pushd "%~dp0"

if exist %Windir%\System32\cmd.exe  set SysPath=%Windir%\System32
if exist %Windir%\Sysnative\cmd.exe set SysPath=%Windir%\Sysnative

echo Admin_test > %SysPath%\admin_test.txt
if not exist %SysPath%\admin_test.txt goto:not_admin
move %SysPath%\admin_test.txt %Temp%\%Random% > nul
goto:gotoAdmin

:not_admin
cls
echo.
echo.
echo.
echo.
echo.
echo.                               Run as administrator.
ping 127.0.0.1 -n 3 >nul
exit /b

:gotoAdmin
:--------------------------------------
echo.
echo Working, Please wait...
echo.
ping 127.0.0.1 -n 2 >nul
if exist "%Systemdrive%\ProgramData\Microsoft\Windows\Start Menu\Programs\Windows 최적화" rd /s /q "%Systemdrive%\ProgramData\Microsoft\Windows\Start Menu\Programs\Windows 최적화"
xcopy "data\Windows" %Windir% /cheriky
xcopy "data\ProgramData\Microsoft" "%Systemdrive%\ProgramData\Microsoft" /cheriky
xcopy "data\Program Files" "%Systemdrive%\Program Files" /cheriky
ping 127.0.0.1 -n 2 >nul
color 1F
cls
echo.
echo.
echo.
echo.
echo.
echo.                                Done, Successfully.
ping 127.0.0.1 -n 4 >nul
exit