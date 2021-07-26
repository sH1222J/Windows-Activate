'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  WScript.Quit
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_DWORD = 4
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem",,48)
For Each objItem in colItems
TimeZone = objItem.CurrentTimeZone
Next
strComputer = "."
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
    strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
strValueName = "InstallDate"
dwValue = UnixTime
oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue

WScript.Echo "Windows 시스템 파일 복구가 완료 되었습니다." +vbCrLf+ "" +vbCrLf+ "" +vbTab+ "" &_
      DateAdd("s", dwValue + TimeZone * 60, "1970-01-01")

Function UnixTime()
UnixTime = DateDiff("s", "1970-01-01 00:00:00", Now())
UnixTime = UnixTime - TimeZone * 60
End Function
