@echo off
title Windows ÃÖÀûÈ­ Ultimate
set /a a=20
set i=5

:_loop
for %%a in (¡è ¢Ö ¡æ ¢Ù ¡é ¢× ¡ç ¢Ø) do (
echo [2;3H[96m%%a
ping 127.0.0.1 -n 1 > nul)
if %a% == 0 goto _out
set /a a=a-1
goto _loop

:_out
echo.[0m
echo.
goto:main

:main
cls
echo.
echo Working, Please wait ...[%i%]
set /A i=i-1
if /i %i% == -1 goto:Run
ping 127.0.0.1 -n 2 >nul
goto:main

:Run
cls
echo.
echo Optimization in progress. [1/2]
echo Working, Please wait...
echo.
1>nul 2>nul net stop wuauserv
1>nul 2>nul net stop WaaSMedicSvc
1>nul 2>nul takeown /f "%Windir%\WinSxS\Backup" /r /d y & 1>nul 2>nul icacls "%Windir%\WinSxS\Backup" /t /grant *S-1-5-32-544:F
1>nul 2>nul takeown /f "%Windir%\WinSxS\InstallTemp" /r /d y & 1>nul 2>nul icacls "%Windir%\WinSxS\InstallTemp" /t /grant *S-1-5-32-544:F
1>nul 2>nul takeown /f "%Windir%\WinSxS\Temp" /r /d y & 1>nul 2>nul icacls "%Windir%\WinSxS\Temp" /t /grant *S-1-5-32-544:F
1>nul 2>nul takeown /f "%Windir%\Prefetch" /r /d y & 1>nul 2>nul icacls "%Windir%\Prefetch" /t /grant *S-1-5-32-544:F
1>nul 2>nul takeown /f "%Windir%\SoftwareDistribution\Download" /r /d y & 1>nul 2>nul icacls "%Windir%\SoftwareDistribution\Download" /t /grant *S-1-5-32-544:F
1>nul 2>nul takeown /f "%Windir%\Logs" /r /d y & 1>nul 2>nul icacls "%Windir%\Logs" /t /grant *S-1-5-32-544:F
1>nul 2>nul takeown /f "%Windir%\System32\LogFiles" /r /d y & 1>nul 2>nul icacls "%Windir%\System32\LogFiles" /t /grant *S-1-5-32-544:F
1>nul 2>nul takeown /f "%Temp%" /r /d y & 1>nul 2>nul icacls "%Temp%" /t /grant *S-1-5-32-544:F
1>nul 2>nul takeown /f "%Tmp%" /r /d y & 1>nul 2>nul icacls "%Tmp%" /t /grant *S-1-5-32-544:F
1>nul 2>nul takeown /f "%Windir%\Temp" /r /d y & 1>nul 2>nul icacls "%Windir%\Temp" /t /grant *S-1-5-32-544:F
1>nul 2>nul del /s /q "%Windir%\WinSxS\Backup\*"
1>nul 2>nul del /s /q "%Windir%\WinSxS\InstallTemp\*"
1>nul 2>nul del /s /q "%Windir%\WinSxS\Temp\*"
1>nul 2>nul del /s /q "%Windir%\Prefetch\*"
1>nul 2>nul del /s /q "%Windir%\LiveKernelReports\*"
1>nul 2>nul del /s /q "%Temp%\*"
1>nul 2>nul del /s /q "%Windir%\Temp\*"
1>nul 2>nul del /s /q "%Windir%\Logs\*"
1>nul 2>nul del /s /q "%Windir%\System32\LogFiles\*"
1>nul 2>nul del /s /q "%Windir%\SoftwareDistribution\Download\*"
echo OK
echo.
echo Optimization in progress. [3/3]
echo Working, Please wait...
echo.
compact /CompactOs:never
ping 127.0.0.1 -n 3 >nul
goto:end

:end
color 1F
cls
echo.
echo.
echo.
echo.
echo.
echo.                                Done, Successfully.
ping 127.0.0.1 -n 3 >nul
start %Windir%\Optimization-Ultimate.vbs
exit

