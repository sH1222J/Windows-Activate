@echo off
title Windows 최적화 업데이트 삭제
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
if exist "%Windir%\Optimization.cmd" del /s /q "%Windir%\Optimization.cmd"
if exist "%Windir%\Optimization-Extreme.cmd" attrib "%Windir%\Optimization-Extreme.cmd" -h -r -s
if exist "%Windir%\Optimization-Extreme.cmd" del /s /q "%Windir%\Optimization-Extreme.cmd"
if exist "%Windir%\Optimization-Ultimate.cmd" attrib "%Windir%\Optimization-Ultimate.cmd" -h -r -s
if exist "%Windir%\Optimization-Ultimate.cmd" del /s /q "%Windir%\Optimization-Ultimate.cmd"
if exist "%Windir%\Optimization-Extreme.vbs" attrib "%Windir%\Optimization-Extreme.vbs" -h -r -s
if exist "%Windir%\Optimization-Extreme.vbs" del /s /q "%Windir%\Optimization-Extreme.vbs"
if exist "%Windir%\Optimization-Ultimate.vbs" attrib "%Windir%\Optimization-Ultimate.vbs" -h -r -s
if exist "%Windir%\Optimization-Ultimate.vbs" del /s /q "%Windir%\Optimization-Ultimate.vbs"
if exist "%Windir%\RestoreHealth.cmd" attrib "%Windir%\RestoreHealth.cmd" -h -r -s
if exist "%Windir%\RestoreHealth.cmd" del /s /q "%Windir%\RestoreHealth.cmd"
if exist "%Windir%\RestoreHealth.vbs" attrib "%Windir%\RestoreHealth.vbs" -h -r -s
if exist "%Windir%\RestoreHealth.vbs" del /s /q "%Windir%\RestoreHealth.vbs"
if exist "%Systemdrive%\ProgramData\Microsoft\Windows\Start Menu\Programs\Windows 최적화" rd /s /q "%Systemdrive%\ProgramData\Microsoft\Windows\Start Menu\Programs\Windows 최적화"
if exist %Windir%\임시폴더삭제.exe attrib %Windir%\임시폴더삭제.exe -h -r -s
if exist %Windir%\임시폴더삭제.exe del /s /q %Windir%\임시폴더삭제.exe
if exist "%Homedrive%\Program Files\DefenderControl" rd /s /q "%Homedrive%\Program Files\DefenderControl"
if exist "%Homedrive%\Program Files\Defender Control" rd /s /q "%Homedrive%\Program Files\Defender Control"
if exist "%Homedrive%\Program Files\ReduceMemory" rd /s /q "%Homedrive%\Program Files\ReduceMemory"
if exist "%Homedrive%\Program Files\TempCleaner" rd /s /q "%Homedrive%\Program Files\TempCleaner"
if exist "%Homedrive%\Program Files\Wub" attrib "%Homedrive%\Program Files\Wub" -h -r -s
if exist "%Homedrive%\Program Files\Wub" rd /s /q "%Homedrive%\Program Files\Wub"
if exist "%Windir%\System32\kms" attrib "%Windir%\System32\kms" -h -r -s
if exist "%Windir%\System32\kms" rd /s /q "%Windir%\System32\kms"
if exist "%Windir%\System32\kms.cmd" attrib "%Windir%\System32\kms.cmd" -h -r -s
if exist "%Windir%\System32\kms.cmd" del /s /q "%Windir%\System32\kms.cmd"
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