@echo off
title KMS Activator
color 1F

IF EXIST %Windir%\System32\cmd.exe  set SysPath=%Windir%\System32
IF EXIST %Windir%\Sysnative\cmd.exe set SysPath=%Windir%\Sysnative
SET RC=%SysPath%\reg.exe

:--------------------------------------
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
ping 127.0.0.1 -n 5 >nul
exit /b

:gotoAdmin
:--------------------------------------
echo.
echo Working, Please wait...
echo.
echo %date% %time%
echo.

pushd "%~dp0"
cd /d "%~dp0kms\"

setlocal EnableExtensions EnableDelayedExpansion
::============================================================================================================================================================
:: Can be 0 (Online Mode - Used for External KMS Server) | 1 (Offline Mode - Used for Internal KMS Server)
set /a _OfflineMode=1
:: Can be 0 (Delete Auto-Renewal Task OR Manual Mode) | 1 (Create Auto-Renewal Task)
set /a _Task=1
::============================================================================================================================================================
:: Set Parameters for KMS Server
:: Can be Custom ePID; Change '_RandomLevel' value to 0 to enable Custom value
set "_WindowsEPID=06401-00206-271-298329-03-1033-9600.0000-0452015"
:: Can be Custom ePID; Change '_RandomLevel' value to 0 to enable Custom value
set "_Office2010EPID=06401-00096-199-198322-03-1033-9600.0000-0452015"
:: Can be Custom ePID; Change '_RandomLevel' value to 0 to enable Custom value
set "_Office2013EPID=06401-00206-234-398213-03-1033-9600.0000-0452015"
:: Can be Custom HardwareID obtained from a Real KMS Server Host
set "_HardwareID=364F463A8863D35F"
:: Can be 0 (UserDefined-Custom ePIDs) | 1 (Randomized ePIDs for every Session) | 2 (Randomized ePIDs for every Request)
set /a _RandomLevel=1
:: Can be (15 to 43200) minutes; Default - 2 hours, Maximum - 30 days
set /a _KMSActivationInterval=120
:: Can be (15 to 43200) minutes; Default - 7 days, Maximum - 30 days
set /a _KMSRenewalInterval=10080
::============================================================================================================================================================
:: Set Parameters for KMS Client
:: Can be (0-255.0-255.0-255.0-255)[Offline Mode], but NOT 127.x.x.x/Localhost IPs; Strongly Recommend to leave it as Default
:: KMS Server Name/IP [Online Mode]
set "_KMSHost=10.11.12.13"
:: Can be (1 to 65535); Strongly Recommend to leave it as Default
set /a _KMSPort=1688
::============================================================================================================================================================
:: Set Registry Key for DLL Hook
set _regKey=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options
::============================================================================================================================================================
:: Registry Keys for SPP and OSPP
set _hkSPP="HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set _huSPP="HKEY_USERS\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set _hkOSPP="HKLM\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
::============================================================================================================================================================
:: Set KMS Genuine Ticket Validation Parameters
:: Can be 0 (Enable Genuine Ticket) | 1 (Disable Genuine Ticket)
set /a _KMSNoGenTicket=1
:: Registry Keys for setting value
set _KMSGenuineKey1=HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform
set _KMSGenuineKey2=HKLM\SOFTWARE\Classes\AppID\slui.exe
::============================================================================================================================================================

:: Check if [Office 2010 on Windows XP SP3 or Later] OR [Office 2013 or Later on Windows 7 / Server 2008 R2] is Installed
wmic path OfficeSoftwareProtectionService get Version >nul 2>&1 && (
	set /a _OSPS=1
) || (
	set /a _OSPS=0
)

:: Check if Office 2016 products are ACTUALLY installed
set /a _Office16=0
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do (
	set "_msi16=%%b"
)
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do (
	set "_msi16wow=%%b"
)
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do (
	set "_ctr16=%%b\Office16"
)
if exist "%_msi16%\OSPP.VBS" (
	set /a _Office16=1
) else if exist "%_msi16wow%\OSPP.VBS" (
	set /a _Office16=1
) else if exist "%_ctr16%\OSPP.VBS" (
	set /a _Office16=1
) else if exist "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" (
	set /a _Office16=1
) else if exist "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" (
	set /a _Office16=1
)
:: Check if Office 2013 products are ACTUALLY installed
set /a _Office15=0
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do (
	set "_msi15=%%b"
)
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do (
	set "_msi15wow=%%b"
)
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do (
	set "_ctr15=%%b\Office15"
)
if exist "%_msi15%\OSPP.VBS" (
	set /a _Office15=1
) else if exist "%_msi15wow%\OSPP.VBS" (
	set /a _Office15=1
) else if exist "%_ctr15%\OSPP.VBS" (
	set /a _Office15=1
) else if exist "C:\Program Files\Microsoft Office\Office15\OSPP.VBS" (
	set /a _Office15=1
) else if exist "C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS" (
	set /a _Office15=1
)
:: Check if Office 2010 products are ACTUALLY installed
set /a _Office14=0
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do (
	set "_msi14=%%b"
)
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do (
	set "_msi14wow=%%b"
)
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do (
	set "_ctr14=%%b\Office14"
)
if exist "%_msi14%\OSPP.VBS" (
	set /a _Office14=1
) else if exist "%_msi14wow%\OSPP.VBS" (
	set /a _Office14=1
) else if exist "%_ctr14%\OSPP.VBS" (
	set /a _Office14=1
) else if exist "C:\Program Files\Microsoft Office\Office14\OSPP.VBS" (
	set /a _Office14=1
) else if exist "C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS" (
	set /a _Office14=1
)

:: Get Architecture of the OS installed; OS Locale/Language Independent from Windows XP / Server 2003 and Later
for /f "tokens=2 delims==" %%a in ('wmic path Win32_Processor get AddressWidth /value') do (
	set _OSarch=%%a-bit
)

:: Get Windows OS version information and goto Respective Function
for /f "tokens=2 delims==" %%a in ('wmic path Win32_OperatingSystem get BuildNumber /value') do (
	set /a _WinBuild=%%a
)
if %_WinBuild% GEQ 9600 (
	echo Program Started...Please Wait...
	goto :Win8.1AndLater
) else if %_WinBuild% GEQ 2600 (
	echo Program Started...Please Wait...
	goto :Win8AndBelow
) else (
	echo KMS_VL_ALL is NOT supported on this OS.
	echo Press any key to exit...
	pause >nul
	exit /b
)
::============================================================================================================================================================
:Close
echo.
echo %date% %time%
echo ______________________________________________________________________________
echo.
setlocal enableextensions
setLocal EnableDelayedExpansion
echo Windows Status
cscript //nologo %SysPath%\slmgr.vbs /dli
echo ______________________________________________________________________________
FOR /F "tokens=2*" %%a IN ('"reg query HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do (SET Office=%%b)
if exist "%Office%\OSPP.VBS" (
echo.
if exist %systemroot%\SysWOW64\cmd.exe (echo Office 2016 64-bit Status) else (echo Office 2016 Status)
echo.
cd /d "%Office%"
cscript //nologo ospp.vbs /dstatus
echo ______________________________________________________________________________
cd /d "%~dp0"
)

if not exist %systemroot%\SysWOW64\cmd.exe goto :Office2013
SET Office=
FOR /F "tokens=2*" %%a IN ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do (SET Office=%%b)
if exist "%Office%\OSPP.VBS" (
echo.
echo Office 2016 32-bit Status
echo.
cd /d "%Office%"
cscript //nologo ospp.vbs /dstatus
echo ______________________________________________________________________________
cd /d "%~dp0"
)

:Office2013
SET Office=
FOR /F "tokens=2*" %%a IN ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do (SET Office=%%b)
if exist "%Office%\OSPP.VBS" (
echo.
if exist %systemroot%\SysWOW64\cmd.exe (echo Office 2013 64-bit Status) else (echo Office 2013 Status)
echo.
cd /d "%Office%"
cscript //nologo ospp.vbs /dstatus
echo ______________________________________________________________________________
cd /d "%~dp0"
)

if not exist %systemroot%\SysWOW64\cmd.exe goto :end
SET Office=
FOR /F "tokens=2*" %%a IN ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do (SET Office=%%b)
if exist "%Office%\OSPP.VBS" (
echo.
echo Office 2013 32-bit Status
echo.
cd /d "%Office%"
cscript //nologo ospp.vbs /dstatus
echo ______________________________________________________________________________
cd /d "%~dp0"
)

:Office2010
SET Office=
FOR /F "tokens=2*" %%a IN ('"reg query HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do (SET Office=%%b)
if exist "%Office%\OSPP.VBS" (
echo.
if exist %systemroot%\SysWOW64\cmd.exe (echo Office 2010 64-bit Status) else (echo Office 2010 Status)
echo.
cd /d "%Office%"
cscript //nologo ospp.vbs /dstatus
echo ______________________________________________________________________________
cd /d "%~dp0"
)

if not exist %systemroot%\SysWOW64\cmd.exe goto :end
SET Office=
FOR /F "tokens=2*" %%a IN ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do (SET Office=%%b)
if exist "%Office%\OSPP.VBS" (
echo.
echo Office 2010 32-bit Status
echo.
cd /d "%Office%"
cscript //nologo ospp.vbs /dstatus
echo ______________________________________________________________________________
cd /d "%~dp0"
)

:end
echo+
echo+
echo Press any key to Exit
pause>nul
exit
::============================================================================================================================================================
:Win8.1AndLater

set ServiceName=KMS Connection System
set port=1688
set kmsip=10.11.12.13
set RandomPID=1
set ActivationInterval=120
set RenewalInterval=10080
set Verbose=0
set PIDWindows=Random
set PIDOffice2010=Random
set PIDOffice2013=Random
set PIDOffice2016=Random
set ServiceKey=HKLM\SYSTEM\CurrentControlSet\Services\%ServiceName%\Parameters

if exist "%Windir%\KMS Connection System\KMS Connection System.exe" goto:Finale

reg.exe query "hklm\software\microsoft\Windows NT\currentversion" /v buildlabex | find /i "amd64" >nul 2>&1
if %errorlevel% equ 0 set xOS=x64
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set xOS=x86
xcopy "%xOS%" "%Windir%\KMS Connection System" /cheriky  >nul

echo Creating KMS Connection System
sc create "%ServiceName%" DisplayName= "KMS Connection System" binPath= "%%Windir%%\KMS Connection System\KMS Connection System.exe" obj= "LocalSystem" type= "own" error= "normal" start= "auto" depend= "RpcSs/tcpip" >nul
echo Adding Registry settings for Service
reg add "%ServiceKey%" /f /v "Port" /d "%Port%" /t "REG_DWORD" >NUL 2>&1
reg add "%ServiceKey%" /f /v "ActivationInterval" /d "%ActivationInterval%" /t "REG_DWORD" >NUL 2>&1
reg add "%ServiceKey%" /f /v "RenewalInterval" /d "%RenewalInterval%" /t "REG_DWORD" >NUL 2>&1
reg add "%ServiceKey%" /f /v "RandomPID" /d "%RandomPID%" /t "REG_DWORD" >NUL 2>&1
reg add "%ServiceKey%" /f /v "Verbose" /d "%Verbose%" /t "REG_DWORD" >NUL 2>&1
reg add "%ServiceKey%" /f /v "WinDivert" /d "1" /t "REG_DWORD" >NUL 2>&1
reg add "%ServiceKey%" /f /v "WinDivertIP" /d "10.11.12.13" /t "REG_SZ" >NUL 2>&1
reg add "%ServiceKey%\PIDs" /f /v "Windows" /d "%PIDWindows%" /t "REG_SZ" >NUL 2>&1
reg add "%ServiceKey%\PIDs" /f /v "Office2010" /d "%PIDOffice2010%" /t "REG_SZ" >NUL 2>&1
reg add "%ServiceKey%\PIDs" /f /v "Office2013" /d "%PIDOffice2013%" /t "REG_SZ" >NUL 2>&1
reg add "%ServiceKey%\PIDs" /f /v "Office2016" /d "%PIDOffice2016%" /t "REG_SZ" >NUL 2>&1

netsh advfirewall firewall show rule name="KMSCS" >nul 2>&1
if %errorlevel% neq 0 (netsh advfirewall firewall add rule name="KMSCS" protocol=any dir=out remoteip=65.52.98.231,65.52.98.232,65.52.98.233 action=block >nul 2>&1)
netsh advfirewall firewall add rule name="KMS_Connection_System" dir=in  program="%Windir%\KMS Connection System\KMS Connection System.exe" localport=%port% protocol=TCP action=allow remoteip=any >nul 2>&1
netsh advfirewall firewall add rule name="KMS_Connection_System" dir=out  program="%Windir%\KMS Connection System\KMS Connection System.exe" localport=%port% protocol=TCP action=allow remoteip=any >nul 2>&1

:Finale
route add 10.11.12.13 0.0.0.0 IF 1 -p >nul 2>&1
echo Starting KMS Connection System
net start "%ServiceName%" >NUL 2>&1

:: Call Windows and Office Activation Main Functions
call :SLSActivation
if %_OSPS% NEQ 0 (
	call :OSPSActivation
)

goto :Close
::============================================================================================================================================================
:Win8AndBelow
:: Exit if No Office 2010 product is installed on Windows XP SP3/Server 2003 R2
if %_OSPS% EQU 0 (
	if %_WinBuild% LSS 6000 (
		echo.
		echo No Office 2010 Product Detected...
		goto :Close
	)
)

if %_OfflineMode% EQU 1 (
	REM Localhost IP can be used for Windows 8 and Below
	set "_KMSHost=127.0.0.2"

	REM Add Firewall Exceptions for VLMCSD and Start KMS Server
	call :AddFirewallRule
	call :StartKMS
)

:: Call Windows and Office Activation Main Functions
if %_WinBuild% GEQ 6000 (
	call :SLSActivation
)
if %_OSPS% NEQ 0 (
	call :OSPSActivation
)

if %_OfflineMode% EQU 1 (
	REM Stop KMS Server and Remove Firewall Exceptions for VLMCSD
	call :StopKMS
	call :RemoveFirewallRule
)
goto :Close
::============================================================================================================================================================
:AddFirewallRule
:: Add VLMCSD KMS Exception to Windows Firewall; Windows XP SP3 or Later Compatible
netsh firewall delete allowedprogram "%~dp0\kms\vlmcsd.exe" >nul 2>&1
netsh firewall add allowedprogram "%~dp0\kms\vlmcsd.exe" "vlmcsd" >nul 2>&1
exit /b
::============================================================================================================================================================
:RemoveFirewallRule
:: Remove VLMCSD KMS Exception from Windows Firewall
netsh firewall delete allowedprogram "%~dp0\kms\vlmcsd.exe" >nul 2>&1
exit /b
::============================================================================================================================================================
:StartKMS
:: Start VLMCSD KMS Server
if %_RandomLevel% EQU 0 (
	start "" /b "%~dp0\kms\vlmcsd.exe" -P %_KMSPort% -0 %_Office2010EPID% -3 %_Office2013EPID% -w %_WindowsEPID% -H %_HardwareID% -R %_KMSRenewalInterval% -A %_KMSActivationInterval% -e >nul 2>&1
) else (
	start "" /b "%~dp0\kms\vlmcsd.exe" -r %_RandomLevel% -P %_KMSPort% -H %_HardwareID% -R %_KMSRenewalInterval% -A %_KMSActivationInterval% -e >nul 2>&1
)
:: Mind boggling BUG Fix; Windows Vista or below takes some time to start KMS Server which prevents Activation; So add delay for it to successfully start
if %_WinBuild% LSS 7600 (
	ping 127.0.0.1 -n 12 >nul 2>&1
)
exit /b
::============================================================================================================================================================
:StopKMS
:: Stop VLMCSD KMS Server
taskkill /im "vlmcsd.exe" /t /f >nul 2>&1
exit /b
::============================================================================================================================================================
:KMSGenuineTicket
:: Enable/Disable KMS Genuine Ticket Validation registry keys based on user parameter
%RC% add "%_KMSGenuineKey1%" /v NoGenTicket /t REG_DWORD /d %_KMSNoGenTicket% /f >nul 2>&1
%RC% add "%_KMSGenuineKey2%" /v NoGenTicket /t REG_DWORD /d %_KMSNoGenTicket% /f >nul 2>&1
exit /b
::============================================================================================================================================================
:StopService
:: Stop service based on parameter
sc query "%1" | findstr /i "STOPPED" >nul 2>&1 || (
	net stop "%1" /y >nul 2>&1
)
sc query "%1" | findstr /i "STOPPED" >nul 2>&1 || (
	sc stop "%1" >nul 2>&1
)
exit /b
::============================================================================================================================================================
:SLSActivation
%RC% delete %_hkSPP%\55c92734-d682-4d71-983e-d6ec3f16059f /f >nul 2>&1
%RC% delete %_hkSPP%\0ff1ce15-a989-479d-af46-f275c6370663 /f >nul 2>&1
set _spp=SoftwareLicensingProduct
set _sps=SoftwareLicensingService

:: Detect if Office 2013 [Volume Licensed] or Later is Installed
wmic path %_spp% where (Description like '%%KMSCLIENT%%') get Name /value 2>nul | findstr /i "Office" >nul 2>&1 && (
	set /a _OfficeVL=1
) || (
	set /a _OfficeVL=0
	if %_WinBuild% GEQ 9200 (
		echo.
		echo No Office 2013 or Later VL Product Detected; Retail Versions need to be converted to VL first.
	)
)

:: Detect if installed Windows supports KMS Activation
wmic path %_spp% where (Description like '%%KMSCLIENT%%') get Name /value 2>nul | findstr /i "Windows" >nul 2>&1 || (
	echo.
	echo No Supported KMS Client Windows Detected...
	if %_OfficeVL% EQU 0 (
		exit /b
	)
)
:: Call Common Core Activation Routines
call :CommonSLSandOSPS
%RC% delete %_huSPP%\55c92734-d682-4d71-983e-d6ec3f16059f /f >nul 2>&1
%RC% delete %_huSPP%\0ff1ce15-a989-479d-af46-f275c6370663 /f >nul 2>&1
exit /b
::============================================================================================================================================================
:OSPSActivation
%RC% delete %_hkOSPP%\59a52881-a989-479d-af46-f275c6370663 /f >nul 2>&1
%RC% delete %_hkOSPP%\0ff1ce15-a989-479d-af46-f275c6370663 /f >nul 2>&1
set _spp=OfficeSoftwareProtectionProduct
set _sps=OfficeSoftwareProtectionService

:: Determine if installed Office product is Retail or VL version; Exit if no VolumeLicensed Office is detected
wmic path %_spp% where (Description like '%%KMSCLIENT%%') get Name >nul 2>&1 || (
	if %_WinBuild% LSS 9200 (
		echo.
		echo No Office 2010 or Later VL Product Detected; Retail Versions need to be converted to VL first.
		exit /b
	) else (
		echo.
		echo No Office 2010 VL Product Detected; Retail Versions need to be converted to VL first.
		exit /b
	)
)
:: Call Common Core Activation Routines
call :CommonSLSandOSPS
exit /b
::============================================================================================================================================================
:CommonSLSandOSPS
:: Get SoftwareLicensingService/OfficeSoftwareProtectionService version to set 'KMSHost' and 'KMSPort' values
for /f "tokens=2 delims==" %%a in ('"wmic path %_sps% get Version /value"') do (
	set _ver=%%a
)
wmic path %_sps% where version='%_ver%' call SetKeyManagementServiceMachine MachineName="%_KMSHost%" >nul 2>&1
wmic path %_sps% where version='%_ver%' call SetKeyManagementServicePort %_KMSPort% >nul 2>&1
:: This is available only on SoftwareLicensingService version 6.2 and later; Not available for OfficeSoftwareProtectionService
wmic path %_sps% where version='%_ver%' call SetVLActivationTypeEnabled 2 >nul 2>&1

:: For all the supported KMS Clients in SoftwareLicensingProduct/OfficeSoftwareProtectionProduct call 'CheckProduct'
for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Description like '%%KMSCLIENT%%') get ID /value"') do (
	set _ActivationID=%%a
	call :CheckProduct !_ActivationID!
)

exit /b
::============================================================================================================================================================
:CheckProduct
:: Ugly hack for Unknown Windows 10 Enterprise LTSB SKU-ID
if '%1' EQU 'b71515d9-89a2-4c60-88c8-656fbcca7f3a' (
	exit /b
)

:: If Detected KMS Client already has GVLK, call Activate function
for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where ID='%1' get PartialProductKey /value"') do (
	set _PartialKey=%%a
)
:: Note that the below check ONLY works when variable don't contain quotes
if [%_PartialKey%] NEQ [] (
	call :Activate %1
	exit /b
)

:: If Detected KMS Client don't have GVLK, do checks for permanent activation, then install GVLK and activate it
for /f "tokens=3 delims==, " %%a in ('"wmic path %_spp% where ID='%1' get Name /value"') do (
	set _ProductName=%%a
)
if '%_ProductName%' EQU '16' (
	if %_Office16% EQU 0 (
		exit /b
	)
	call :CheckOffice16 %1
	exit /b
) else if '%_ProductName%' EQU '15' (
	if %_Office15% EQU 0 (
		exit /b
	)
	call :CheckOffice15 %1
	exit /b
) else if '%_ProductName%' EQU '14' (
	if %_Office14% EQU 0 (
		exit /b
	)
	call :CheckOffice14 %1
	exit /b
) else (
	call :CheckWindows %1
	exit /b
)
::============================================================================================================================================================
:CheckWindows
wmic path %_spp% where (LicenseStatus='1' and GracePeriodRemaining='0') get Name 2>nul | findstr /i "Windows" >nul 2>&1 && (
	echo.
	echo Detected Windows is permanently activated.
	exit /b
)
:: If Windows is not permanently activated, Install GVLK and Activate
call :SelectKey %1
exit /b
::============================================================================================================================================================
:CheckOffice16
set /a _ls=0
if '%1' EQU 'd450596f-894d-49e0-966a-fd39ed4c4c64' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%Office16ProPlusVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Office 2016 ProPlus is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU '6bf301c1-b94a-43e9-ba31-d494598c47fb' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%Office16VisioProVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Visio 2016 Pro is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU '4f414197-0fc2-4c01-b68a-86cbb9ac254c' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%Office16ProjectProVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Project 2016 Pro is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU 'dedfa23d-6ed1-45a6-85dc-63cae0546de6' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%Office16StandardVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Office 2016 Standard is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU 'aa2a7821-1827-4c2c-8f1d-4513a34dda97' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%Office16VisioStdVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Visio 2016 Standard is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU 'da7ddabc-3fbe-4447-9e01-6ab7440b4cd4' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%Office16ProjectStdVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Project 2016 Standard is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
:: If Office 2016 product is not permanently activated, Install GVLK and Activate
call :SelectKey %1
exit /b
::============================================================================================================================================================
:CheckOffice15
set /a _ls=0
if '%1' EQU 'b322da9c-a2e2-4058-9e4e-f59a6970bd69' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeProPlusVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Office 2013 ProPlus is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU 'e13ac10e-75d0-4aff-a0cd-764982cf541c' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeVisioProVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Visio 2013 Pro is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU '4a5d124a-e620-44ba-b6ff-658961b33b9a' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeProjectProVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Project 2013 Pro is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU 'b13afb38-cd79-4ae5-9f7f-eed058d750ca' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeStandardVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Office 2013 Standard is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU 'ac4efaf0-f81f-4f61-bdf7-ea32b02ab117' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeVisioStdVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Visio 2013 Standard is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU '427a28d1-d17c-4abf-b717-32c780ba6f07' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeProjectStdVL_MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Project 2013 Standard is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
:: If Office 2013 product is not permanently activated, Install GVLK and Activate
call :SelectKey %1
exit /b
::============================================================================================================================================================
:CheckOffice14
set /a _ls=0
set /a _ls2=0
for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeVisioPrem-MAK%%') get LicenseStatus /value" 2^>nul') do set /a _vPrem=%%a
for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeVisioPro-MAK%%') get LicenseStatus /value" 2^>nul') do set /a _vPro=%%a
if '%1' EQU '6f327760-8c5c-417c-9b61-836a98287e0c' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeProPlus-MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeProPlusAcad-MAK%%') get LicenseStatus /value"') do set /a _ls2=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Office 2010 ProPlus is permanently MAK activated.
		exit /b
	)
	if !_ls2! EQU 1 (
		echo.
		echo Detected Office 2010 ProPlus Academic is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU 'df133ff7-bf14-4f95-afe3-7b48e7e331ef' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeProjectPro-MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Project 2010 Pro is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU '5dc7bf61-5ec9-4996-9ccb-df806a2d0efe' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeProjectStd-MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Project 2010 Standard is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU '9da2a678-fb6b-4e67-ab84-60dd6a9c819a' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeStandard-MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Office 2010 Standard is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU 'ea509e87-07a1-4a45-9edc-eba5a39f36af' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeSmallBusBasics-MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Office 2010 Small Business is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if '%1' EQU '92236105-bb67-494f-94c7-7f7a607929bd' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeVisioPrem-MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeVisioPro-MAK%%') get LicenseStatus /value"') do set /a _ls2=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Visio 2010 Premium is permanently MAK activated.
		exit /b
	)
	if !_ls2! EQU 1 (
		echo.
		echo Detected Visio 2010 Pro is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if defined _vPrem exit /b
if '%1' EQU 'e558389c-83c3-4b29-adfe-5e4d7f46c358' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeVisioPro-MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeVisioStd-MAK%%') get LicenseStatus /value"') do set /a _ls2=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Visio 2010 Pro is permanently MAK activated.
		exit /b
	)
	if !_ls2! EQU 1 (
		echo.
		echo Detected Visio 2010 Standard is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
if defined _vPro exit /b
if '%1' EQU '9ed833ff-4f92-4f36-b370-8683a4f13275' (
	for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where (Name like '%%OfficeVisioStd-MAK%%') get LicenseStatus /value"') do set /a _ls=%%a
	if !_ls! EQU 1 (
		echo.
		echo Detected Visio 2010 Standard is permanently MAK activated.
		exit /b
	) else (
		call :SelectKey %1
		exit /b
	)
)
:: If Office 2010 product is not permanently activated, Install GVLK and Activate
call :SelectKey %1
exit /b
::============================================================================================================================================================
:Activate
:: Call Activate method of the corresponding KMS Client
for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where ID='%1' get Name /value"') do (
	echo.
	echo Attempting to Activate %%a
)
wmic path %_spp% where ID='%1' call Activate >nul 2>&1

:: Get Remaining Grace Period of the KMS Client
for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where ID='%1' get GracePeriodRemaining /value"') do (
	set /a _gprMinutes=%%a
)
if %_gprMinutes% EQU 43200 (
	echo Windows Core/ProfessionalWMC Activation Successful
	echo Remaining Period: 30 days ^(%_gprMinutes% minutes^)
	exit /b
)
if %_gprMinutes% EQU 64800 (
	echo Windows Core/ProfessionalWMC Activation Successful
	echo Remaining Period: 45 days ^(%_gprMinutes% minutes^)
	exit /b
)
if %_gprMinutes% EQU 259200 (
	echo Product Activation Successful
) else (
	echo Product Activation Failed
)
set /a _gprDays=%_gprMinutes%/1440
echo Remaining Period: %_gprDays% days ^(%_gprMinutes% minutes^)
exit /b
::============================================================================================================================================================
:SelectKey
:: Select GenericVolumeLicenseKey based on Activation ID (SKU-ID) and Install it, if found
for /f "tokens=2 delims==" %%a in ('"wmic path %_spp% where ID='%1' get Name /value"') do (
	set _Name=%%a
	echo.
	echo Searching GenericVolumeLicenseKey for %%a
	goto :%1 2>nul || goto :KeyNotFound
)
::============================================================================================================================================================
:: Office 2016 Professional Plus
:d450596f-894d-49e0-966a-fd39ed4c4c64
set _key=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
goto :InstallKey
:: Office 2016 Standard
:dedfa23d-6ed1-45a6-85dc-63cae0546de6
set _key=JNRGM-WHDWX-FJJG3-K47QV-DRTFM
goto :InstallKey
:: Project 2016 Professional
:4f414197-0fc2-4c01-b68a-86cbb9ac254c
set _key=YG9NW-3K39V-2T3HJ-93F3Q-G83KT
goto :InstallKey
:: Project 2016 Standard
:da7ddabc-3fbe-4447-9e01-6ab7440b4cd4
set _key=GNFHQ-F6YQM-KQDGJ-327XX-KQBVC
goto :InstallKey
:: Visio 2016 Professional
:6bf301c1-b94a-43e9-ba31-d494598c47fb
set _key=PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
goto :InstallKey
:: Visio 2016 Standard
:aa2a7821-1827-4c2c-8f1d-4513a34dda97
set _key=7WHWN-4T7MP-G96JF-G33KR-W8GF4
goto :InstallKey
:: Access 2016
:67c0fc0c-deba-401b-bf8b-9c8ad8395804
set _key=GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW
goto :InstallKey
:: Excel 2016
:c3e65d36-141f-4d2f-a303-a842ee756a29
set _key=9C2PK-NWTVB-JMPW8-BFT28-7FTBF
goto :InstallKey
:: OneNote 2016
:d8cace59-33d2-4ac7-9b1b-9b72339c51c8
set _key=DR92N-9HTF2-97XKM-XW2WJ-XW3J6
goto :InstallKey
:: Outlook 2016
:ec9d9265-9d1e-4ed0-838a-cdc20f2551a1
set _key=R69KK-NTPKF-7M3Q4-QYBHW-6MT9B
goto :InstallKey
:: PowerPoint 2016
:d70b1bba-b893-4544-96e2-b7a318091c33
set _key=J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6
goto :InstallKey
:: Publisher 2016
:041a06cb-c5b8-4772-809f-416d03d16654
set _key=F47MM-N3XJP-TQXJ9-BP99D-8K837
goto :InstallKey
:: Skype for Business 2016
:83e04ee1-fa8d-436d-8994-d31a862cab77
set _key=869NQ-FJ69K-466HW-QYCP2-DDBV6
goto :InstallKey
:: Word 2016
:bb11badf-d8aa-470e-9311-20eaf80fe5cc
set _key=WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6
goto :InstallKey
::============================================================================================================================================================
:: Office 2013 Professional Plus
:b322da9c-a2e2-4058-9e4e-f59a6970bd69
set _key=YC7DK-G2NP3-2QQC3-J6H88-GVGXT
goto :InstallKey
:: Office 2013 Standard
:b13afb38-cd79-4ae5-9f7f-eed058d750ca
set _key=KBKQT-2NMXY-JJWGP-M62JB-92CD4
goto :InstallKey
:: Project 2013 Professional
:4a5d124a-e620-44ba-b6ff-658961b33b9a
set _key=FN8TT-7WMH6-2D4X9-M337T-2342K
goto :InstallKey
:: Project 2013 Standard
:427a28d1-d17c-4abf-b717-32c780ba6f07
set _key=6NTH3-CW976-3G3Y2-JK3TX-8QHTT
goto :InstallKey
:: Visio 2013 Professional
:e13ac10e-75d0-4aff-a0cd-764982cf541c
set _key=C2FG9-N6J68-H8BTJ-BW3QX-RM3B3
goto :InstallKey
:: Visio 2013 Standard
:ac4efaf0-f81f-4f61-bdf7-ea32b02ab117
set _key=J484Y-4NKBF-W2HMG-DBMJC-PGWR7
goto :InstallKey
:: Access 2013
:6ee7622c-18d8-4005-9fb7-92db644a279b
set _key=NG2JY-H4JBT-HQXYP-78QH9-4JM2D
goto :InstallKey
:: Excel 2013
:f7461d52-7c2b-43b2-8744-ea958e0bd09a
set _key=VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB
goto :InstallKey
:: InfoPath 2013
:a30b8040-d68a-423f-b0b5-9ce292ea5a8f
set _key=DKT8B-N7VXH-D963P-Q4PHY-F8894
goto :InstallKey
:: Lync 2013
:1b9f11e3-c85c-4e1b-bb29-879ad2c909e3
set _key=2MG3G-3BNTT-3MFW9-KDQW3-TCK7R
goto :InstallKey
:: OneNote 2013
:efe1f3e6-aea2-4144-a208-32aa872b6545
set _key=TGN6P-8MMBC-37P2F-XHXXK-P34VW
goto :InstallKey
:: Outlook 2013
:771c3afa-50c5-443f-b151-ff2546d863a0
set _key=QPN8Q-BJBTJ-334K3-93TGY-2PMBT
goto :InstallKey
:: PowerPoint 2013
:8c762649-97d1-4953-ad27-b7e2c25b972e
set _key=4NT99-8RJFH-Q2VDH-KYG2C-4RD4F
goto :InstallKey
:: Publisher 2013
:00c79ff1-6850-443d-bf61-71cde0de305f
set _key=PN2WF-29XG2-T9HJ7-JQPJR-FCXK4
goto :InstallKey
:: Word 2013
:d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3
set _key=6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7
goto :InstallKey
:: SharePoint Designer 2013 Retail
:ba3e3833-6a7e-445a-89d0-7802a9a68588
set _key=GYJRG-NMYMF-VGBM4-T3QD4-842DW
goto :InstallKey
:: Mondo 2013
:dc981c6b-fc8e-420f-aa43-f8f33e5c0923
set _key=42QTK-RN8M7-J3C4G-BBGYM-88CYV
goto :InstallKey
::============================================================================================================================================================
:: Office 2010 Professional Plus
:6f327760-8c5c-417c-9b61-836a98287e0c
set _key=VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
goto :InstallKey
:: Office 2010 Standard
:9da2a678-fb6b-4e67-ab84-60dd6a9c819a
set _key=V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
goto :InstallKey
:: Office 2010 Starter Retail
:2745e581-565a-4670-ae90-6bf7c57ffe43
set _key=VXHHB-W7HBD-7M342-RJ7P8-CHBD6
goto :InstallKey
:: Access 2010
:8ce7e872-188c-4b98-9d90-f8f90b7aad02
set _key=V7Y44-9T38C-R2VJK-666HK-T7DDX
goto :InstallKey
:: Excel 2010
:cee5d470-6e3b-4fcc-8c2b-d17428568a9f
set _key=H62QG-HXVKF-PP4HP-66KMR-CW9BM
goto :InstallKey
:: SharePoint Workspace 2010 (Groove)
:8947d0b8-c33b-43e1-8c56-9b674c052832
set _key=QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4
goto :InstallKey
:: SharePoint Designer 2010 Retail
:b78df69e-0966-40b1-ae85-30a5134dedd0
set _key=H48K6-FB4Y6-P83GH-9J7XG-HDKKX
goto :InstallKey
:: InfoPath 2010
:ca6b6639-4ad6-40ae-a575-14dee07f6430
set _key=K96W8-67RPQ-62T9Y-J8FQJ-BT37T
goto :InstallKey
:: OneNote 2010
:ab586f5c-5256-4632-962f-fefd8b49e6f4
set _key=Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX
goto :InstallKey
:: Outlook 2010
:ecb7c192-73ab-4ded-acf4-2399b095d0cc
set _key=7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ
goto :InstallKey
:: PowerPoint 2010
:45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a
set _key=RC8FX-88JRY-3PF7C-X8P67-P4VTT
goto :InstallKey
:: Project 2010 Professional
:df133ff7-bf14-4f95-afe3-7b48e7e331ef
set _key=YGX6F-PGV49-PGW3J-9BTGG-VHKC6
goto :InstallKey
:: Project 2010 Standard
:5dc7bf61-5ec9-4996-9ccb-df806a2d0efe
set _key=4HP3K-88W3F-W2K3D-6677X-F9PGB
goto :InstallKey
:: Publisher 2010
:b50c4f75-599b-43e8-8dcd-1081a7967241
set _key=BFK7F-9MYHM-V68C7-DRQ66-83YTP
goto :InstallKey
:: Word 2010
:2d0882e7-a4e7-423b-8ccc-70d91e0158b1
set _key=HVHB3-C6FV7-KQX9W-YQG79-CRY7T
goto :InstallKey
:: Visio 2010 Premium
:92236105-bb67-494f-94c7-7f7a607929bd
set _key=D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ
goto :InstallKey
:: Visio 2010 Professional
:e558389c-83c3-4b29-adfe-5e4d7f46c358
set _key=7MCW8-VRQVK-G677T-PDJCM-Q8TCP
goto :InstallKey
:: Visio 2010 Standard
:9ed833ff-4f92-4f36-b370-8683a4f13275
set _key=767HD-QGMWX-8QTDB-9G3R2-KHFGJ
goto :InstallKey
:: Office 2010 Home and Business
:ea509e87-07a1-4a45-9edc-eba5a39f36af
set _key=D6QFG-VBYP2-XQHM7-J97RH-VVRCK
goto :InstallKey
:: Office 2010 Mondo
:09ed9640-f020-400a-acd8-d7d867dfd9c2
set _key=YBJTT-JG6MD-V9Q7P-DBKXJ-38W9R
goto :InstallKey
:: Office 2010 Mondo
:ef3d4e49-a53d-4d81-a2b1-2ca6c2556b2c
set _key=7TC2V-WXF6P-TD7RT-BQRXR-B8K32
goto :InstallKey
::============================================================================================================================================================
:: Windows 10 Professional
:2de67392-b7a7-462a-b1ca-108dd189f588
set _key=W269N-WFGWX-YVC9B-4J6C9-T83GX
goto :InstallKey
::============================================================================================================================================================
:: Windows 8.1 Professional
:c06b6981-d7fd-4a35-b7b4-054742b7af67
set _key=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
goto :InstallKey
:: Windows 8.1 Professional N
:7476d79f-8e48-49b4-ab63-4d0b813a16e4
set _key=HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
goto :InstallKey
:: Windows 8.1 Enterprise
:81671aaf-79d1-4eb1-b004-8cbbe173afea
set _key=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
goto :InstallKey
:: Windows 8.1 Enterprise N
:113e705c-fa49-48a4-beea-7dd879b46b14
set _key=TT4HM-HN7YT-62K67-RGRQJ-JFFXW
goto :InstallKey
:: Windows Server 2012 R2 Server Standard
:b3ca044e-a358-4d68-9883-aaa2941aca99
set _key=D2N9P-3P6X9-2R39C-7RTCD-MDVJX
goto :InstallKey
:: Windows Server 2012 R2 Datacenter
:00091344-1ea4-4f37-b789-01750ba6988c
set _key=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
goto :InstallKey
:: Windows Server 2012 R2 Essentials
:21db6ba4-9a7b-4a14-9e29-64a60c59301d
set _key=KNC87-3J2TX-XB4WP-VCPJV-M4FWM
goto :InstallKey
:: Windows 8.1 Professional WMC
:096ce63d-4fac-48a9-82a9-61ae9e800e5f
set _key=789NJ-TQK6T-6XTH8-J39CJ-J8D3P
goto :InstallKey
:: Windows 8.1 Core
:fe1c3238-432a-43a1-8e25-97e7d1ef10f3
set _key=M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK
goto :InstallKey
:: Windows 8.1 Core N
:78558a64-dc19-43fe-a0d0-8075b2a370a3
set _key=7B9N3-D94CG-YTVHR-QBPX3-RJP64
goto :InstallKey
:: Windows 8.1 Core ARM
:ffee456a-cd87-4390-8e07-16146c672fd0
set _key=XYTND-K6QKT-K2MRH-66RTM-43JKP
goto :InstallKey
:: Windows 8.1 Core Single Language
:c72c6a1d-f252-4e7e-bdd1-3fca342acb35
set _key=BB6NG-PQ82V-VRDPW-8XVD2-V8P66
goto :InstallKey
:: Windows 8.1 Core Country Specific
:db78b74f-ef1c-4892-abfe-1e66b8231df6
set _key=NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3
goto :InstallKey
:: Windows Server 2012 R2 Cloud Storage
:b743a2be-68d4-4dd3-af32-92425b7bb623
set _key=3NPTF-33KPT-GGBPR-YX76B-39KDD
goto :InstallKey
:: Windows 8.1 Embedded Industry
:0ab82d54-47f4-4acb-818c-cc5bf0ecb649
set _key=NMMPB-38DD4-R2823-62W8D-VXKJB
goto :InstallKey
:: Windows 8.1 Embedded Industry Enterprise
:cd4e2d9f-5059-4a50-a92d-05d5bb1267c7
set _key=FNFKF-PWTVT-9RC8H-32HB2-JB34X
goto :InstallKey
:: Windows 8.1 Embedded Industry Automotive
:f7e88590-dfc7-4c78-bccb-6f3865b99d1a
set _key=VHXM3-NR6FT-RY6RT-CK882-KW2CJ
goto :InstallKey
:: Windows 8.1 Core Connected (with Bing)
:e9942b32-2e55-4197-b0bd-5ff58cba8860
set _key=3PY8R-QHNP9-W7XQD-G6DPH-3J2C9
goto :InstallKey
:: Windows 8.1 Core Connected N (with Bing)
:c6ddecd6-2354-4c19-909b-306a3058484e
set _key=Q6HTR-N24GM-PMJFP-69CD8-2GXKR
goto :InstallKey
:: Windows 8.1 Core Connected Single Language (with Bing)
:b8f5e3a3-ed33-4608-81e1-37d6c9dcfd9c
set _key=KF37N-VDV38-GRRTV-XH8X6-6F3BB
goto :InstallKey
:: Windows 8.1 Core Connected Country Specific (with Bing)
:ba998212-460a-44db-bfb5-71bf09d1c68b
set _key=R962J-37N87-9VVK2-WJ74P-XTMHR
goto :InstallKey
:: Windows 8.1 Professional Student
:e58d87b5-8126-4580-80fb-861b22f79296
set _key=MX3RK-9HNGX-K3QKC-6PJ3F-W8D7B
goto :InstallKey
:: Windows 8.1 Professional Student N
:cab491c7-a918-4f60-b502-dab75e334f40
set _key=TNFGH-2R6PB-8XM3K-QYHX2-J4296
goto :InstallKey
::============================================================================================================================================================
:: Windows 8 Professional
:a98bcd6d-5343-4603-8afe-5908e4611112
set _key=NG4HW-VH26C-733KW-K6F98-J8CK4
goto :InstallKey
:: Windows 8 Professional N
:ebf245c1-29a8-4daf-9cb1-38dfc608a8c8
set _key=XCVCF-2NXM9-723PB-MHCB7-2RYQQ
goto :InstallKey
:: Windows 8 Enterprise
:458e1bec-837a-45f6-b9d5-925ed5d299de
set _key=32JNW-9KQ84-P47T8-D8GGY-CWCK7
goto :InstallKey
:: Windows 8 Enterprise N
:e14997e7-800a-4cf7-ad10-de4b45b578db
set _key=JMNMF-RHW7P-DMY6X-RF3DR-X2BQT
goto :InstallKey
:: Windows Server 2012 / Windows 8 Core
:c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60
set _key=BN3D2-R7TKB-3YPBD-8DRP2-27GG4
goto :InstallKey
:: Windows Server 2012 N / Windows 8 Core N
:197390a0-65f6-4a95-bdc4-55d58a3b0253
set _key=8N2M2-HWPGY-7PGT9-HGDD8-GVGGY
goto :InstallKey
:: Windows 8 Core ARM
:af35d7b7-5035-4b63-8972-f0b747b9f4dc
set _key=DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV
goto :InstallKey
:: Windows Server 2012 Single Language / Windows 8 Core Single Language
:8860fcd4-a77b-4a20-9045-a150ff11d609
set _key=2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
goto :InstallKey
:: Windows Server 2012 Country Specific / Windows 8 Core Country Specific
:9d5584a2-2d85-419a-982c-a00888bb9ddf
set _key=4K36P-JN4VD-GDC6V-KDT89-DYFKP
goto :InstallKey
:: Windows Server 2012 Standard
:f0f5ec41-0d55-4732-af02-440a44a3cf0f
set _key=XC9B7-NBPP2-83J2H-RHMBY-92BT4
goto :InstallKey
:: Windows Server 2012 MultiPoint Standard
:7d5486c7-e120-4771-b7f1-7b56c6d3170c
set _key=HM7DN-YVMH3-46JC3-XYTG7-CYQJJ
goto :InstallKey
:: Windows Server 2012 MultiPoint Premium
:95fd1c83-7df5-494a-be8b-1300e1c9d1cd
set _key=XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G
goto :InstallKey
:: Windows Server 2012 Datacenter
:d3643d60-0c42-412d-a7d6-52e6635327f6
set _key=48HP8-DN98B-MYWDG-T2DCC-8W83P
goto :InstallKey
:: Windows 8 Professional WMC
:a00018a3-f20f-4632-bf7c-8daa5351c914
set _key=GNBB8-YVD74-QJHX6-27H4K-8QHDG
goto :InstallKey
:: Windows 8 Embedded Industry Professional
:10018baf-ce21-4060-80bd-47fe74ed4dab
set _key=RYXVT-BNQG7-VD29F-DBMRY-HT73M
goto :InstallKey
:: Windows 8 Embedded Industry Enterprise
:18db1848-12e0-4167-b9d7-da7fcda507db
set _key=NKB3R-R2F8T-3XCDP-7Q2KW-XWYQ2
goto :InstallKey
::============================================================================================================================================================
:: Windows 7 Professional
:b92e9980-b9d5-4821-9c94-140f632f6312
set _key=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
goto :InstallKey
:: Windows 7 Professional N
:54a09a0d-d57b-4c10-8b69-a842d6590ad5
set _key=MRPKT-YTG23-K7D7T-X2JMM-QY7MG
goto :InstallKey
:: Windows 7 Professional E
:5a041529-fef8-4d07-b06f-b59b573b32d2
set _key=W82YF-2Q76Y-63HXB-FGJG9-GF7QX
goto :InstallKey
:: Windows 7 Enterprise
:ae2ee509-1b34-41c0-acb7-6d4650168915
set _key=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
goto :InstallKey
:: Windows 7 Enterprise N
:1cb6d605-11b3-4e14-bb30-da91c8e3983a
set _key=YDRBP-3D83W-TY26F-D46B2-XCKRJ
goto :InstallKey
:: Windows 7 Enterprise E
:46bbed08-9c7b-48fc-a614-95250573f4ea
set _key=C29WB-22CC8-VJ326-GHFJW-H9DH4
goto :InstallKey
:: Windows Server 2008 R2 Web
:a78b8bd9-8017-4df5-b86a-09f756affa7c
set _key=6TPJF-RBVHG-WBW2R-86QPH-6RTM4
goto :InstallKey
:: Windows Server 2008 R2 HPC edition
:cda18cf3-c196-46ad-b289-60c072869994
set _key=TT8MH-CG224-D3D7Q-498W2-9QCTX
goto :InstallKey
:: Windows Server 2008 R2 Standard
:68531fb9-5511-4989-97be-d11a0f55633f
set _key=YC6KT-GKW9T-YTKYR-T4X34-R7VHC
goto :InstallKey
:: Windows Server 2008 R2 Enterprise
:620e2b3d-09e7-42fd-802a-17a13652fe7a
set _key=489J6-VHDMP-X63PK-3K798-CPX3Y
goto :InstallKey
:: Windows Server 2008 R2 Datacenter
:7482e61b-c589-4b7f-8ecc-46d455ac3b87
set _key=74YFP-3QFB3-KQT8W-PMXWJ-7M648
goto :InstallKey
:: Windows Server 2008 R2 for Itanium-based Systems
:8a26851c-1c7e-48d3-a687-fbca9b9ac16b
set _key=GT63C-RJFQ3-4GMB6-BRFB9-CB83V
goto :InstallKey
:: Windows 7 Embedded POS Ready
:db537896-376f-48ae-a492-53d0547773d0
set _key=YBYF6-BHCR3-JPKRB-CDW7B-F9BK4
goto :InstallKey
:: Windows 7 Embedded ThinPC
:aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395
set _key=73KQT-CD9G6-K7TQG-66MRP-CQ22C
goto :InstallKey
:: Windows 7 Embedded Standard OEM
:e1a8296a-db37-44d1-8cce-7bc961d59c54
set _key=XGY72-BRBBT-FF8MH-2GG8H-W7KCW
goto :InstallKey
:: Windows MultiPoint Server 2010
:f772515c-0e87-48d5-a676-e6962c3e1195
set _key=736RG-XDKJK-V34PF-BHK87-J6X3K
goto :InstallKey
::============================================================================================================================================================
:: Windows Vista Business
:4f3d1606-3fea-4c01-be3c-8d671c401e3b
set _key=YFKBB-PQJJV-G996G-VWGXY-2V3X8
goto :InstallKey
:: Windows Vista Business N
:2c682dc2-8b68-4f63-a165-ae291d4cf138
set _key=HMBQG-8H2RH-C77VX-27R82-VMQBT
goto :InstallKey
:: Windows Vista Enterprise
:cfd8ff08-c0d7-452b-9f60-ef5c70c32094
set _key=VKK3X-68KWM-X2YGT-QR4M6-4BWMV
goto :InstallKey
:: Windows Vista Enterprise N
:d4f54950-26f2-4fb4-ba21-ffab16afcade
set _key=VTC42-BM838-43QHV-84HX6-XJXKV
goto :InstallKey
:: Windows Web Server 2008
:ddfa9f7c-f09e-40b9-8c1a-be877a9a7f4b
set _key=WYR28-R7TFJ-3X2YQ-YCY4H-M249D
goto :InstallKey
:: Windows Server 2008 Standard
:ad2542d4-9154-4c6d-8a44-30f11ee96989
set _key=TM24T-X9RMF-VWXK6-X8JC9-BFGM2
goto :InstallKey
:: Windows Server 2008 Standard without Hyper-V
:2401e3d0-c50a-4b58-87b2-7e794b7d2607
set _key=W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ
goto :InstallKey
:: Windows Server 2008 Enterprise
:c1af4d90-d1bc-44ca-85d4-003ba33db3b9
set _key=YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
goto :InstallKey
:: Windows Server 2008 Enterprise without Hyper-V
:8198490a-add0-47b2-b3ba-316b12d647b4
set _key=39BXF-X8Q23-P2WWT-38T2F-G3FPG
goto :InstallKey
:: Windows Server 2008 HPC (Compute Cluster)
:7afb1156-2c1d-40fc-b260-aab7442b62fe
set _key=RCTX3-KWVHP-BR6TB-RB6DM-6X7HP
goto :InstallKey
:: Windows Server 2008 Datacenter
:68b6e220-cf09-466b-92d3-45cd964b9509
set _key=7M67G-PC374-GR742-YH8V4-TCBY3
goto :InstallKey
:: Windows Server 2008 Datacenter without Hyper-V
:fd09ef77-5647-4eff-809c-af2b64659a45
set _key=22XQ2-VRXRG-P8D42-K34TD-G3QQC
goto :InstallKey
:: Windows Server 2008 for Itanium-Based Systems
:01ef176b-3e0d-422a-b4f8-4ea880035e8f
set _key=4DWFP-JF3DJ-B7DTH-78FJB-PDRHK
goto :InstallKey
::============================================================================================================================================================
:KeyNotFound
:: If GVLK is not found for the current SKU-ID, attempt Activation as GVLK might be present in OS by default
echo.
echo GVLK for %_Name%
echo with SKU-ID %1 Not Found
echo.
echo If Activation Fails Now, Please enter GVLK for this Product manually and re-run KMS_VL_ALL
call :Activate %1
exit /b
::============================================================================================================================================================
:InstallKey
:: Call InstallProductKey method of SLS/OSPS to install GVLK
echo Installing Key...
wmic path %_sps% where version='%_ver%' call InstallProductKey ProductKey="%_key%" >nul 2>&1
call :Activate %1
::============================================================================================================================================================