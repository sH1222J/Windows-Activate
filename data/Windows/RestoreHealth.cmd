@echo off
title Windows ½Ã½ºÅÛ ÆÄÀÏ º¹±¸
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
Dism /Online /Cleanup-Image /RestoreHealth
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
start %Windir%\RestoreHealth.vbs
exit