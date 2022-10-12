' script:       userlogging.vbs
' author:       Davy Yan (hA SOE)
' date:         21/03/2016
' description:  logging user logon/logoff activities
'               accepts "logon" or "logoff" as compulsory parameter

' 25/11/2016 - calling the NRDHBUserPrinters.vbs to collect printer info
' 14/09/2017 - calling the Delete_Health.nz_cookie.vbs to remove this IE cookie from user profile

Option Explicit

On Error Resume Next

Const HKLM                = &H80000002
Const adCmdStoredProc     = &H0004
Const adExecuteNoRecords  = &H00000080
Const adInteger           = 3
Const adBoolean           = 11
Const adDate              = 7
Const adVarChar           = 200
Const adParamInput        = &H0001

Dim sUser, sComputer, sDomain, sIP, sModel, sType, sSite
Dim sLastBoot, sUpTime, sArchitec, sOSversion, sOSCaption, sInstallDate
Dim network, wmi, reg, objADSysInfo
Dim sALKeyPath, sALKeyValue, sALKeyData, bAutoLogon, sSOEKeyPath, sSOEKeyValue, sSOEKeyData, sSOEVersion
Dim objItem, colItems
Dim strConn, objConn, objCmd
Dim session, strConnAhsl20, strConnVMMH1CSQL040
Dim wshShell

' === 25/11/2016 ===
Dim fso, shell, envproc, sLogonServer
' === 25/11/2016 ===

sALKeyPath    = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
sALKeyValue   = "AutoAdminLogon"
sSOEKeyPath   = "SOFTWARE\Microsoft\Deployment 4"
sSOEKeyValue  = "Task Sequence Version" 
strConnVMMH1CSQL040       = "PROVIDER=SQLOLEDB;DATA SOURCE=VMMH1CSQL040.healthcare.huarahi.health.govt.nz;Initial Catalog=NRDHB_Machine_Logging;User Id=MachineLogging;Password=pCloGging!;"
strConnAhsl20             = "PROVIDER=SQLOLEDB;DATA SOURCE=ahsl20.adhb.govt.nz;Database=MachineLogging;User Id=MachineLoggingUser;Password=L0gg!ng;"


' Check if user in Citrix session. If SESSIONNAME = ICA-Tcp#nnn then Citrix

set wshShell = CreateObject("Wscript.shell")
session = wshShell.expandEnvironmentStrings("%SESSIONNAME%")


If Left(session, 3) = "ICA" Then
  'WScript.Echo "Remote Desktop / Citrix Session detected"
  WScript.Quit
End If


If WScript.Arguments.Count = 1 Then
  If LCase(WScript.Arguments(0)) = "logon" Or LCase(WScript.Arguments(0)) = "logoff" Then
    sType = UCase(WScript.Arguments(0))
  Else
    WScript.Quit
  End If
Else
  WScript.Quit
End If

' Domain, User, Computer
Set network = CreateObject("WScript.Network")
sDomain     = UCase(network.UserDomain)
sUser       = UCase(network.UserName)
sComputer   = UCase(network.ComputerName)

' Computer Site
Set objADSysInfo = CreateObject("ADSystemInfo")
sSite = objADSysInfo.SiteName

' IP
Set wmi     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colItems = wmi.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration Where IPEnabled = True And DNSDomain LIKE '%.nz'")
For Each objItem In colItems
 If Not IsNull(objItem.IPAddress) Then sIP = Trim(objItem.IPAddress(0))
Next

If LCase(sType) = "logon" Then
  bAutoLogon    = False
  sSOEVersion   = "Unknown"

  ' Model
  Set colItems = wmi.ExecQuery("SELECT * FROM Win32_ComputerSystem")
  For Each objItem In colItems
   sModel = objItem.Model
  Next
  
  ' OS name, version, architec, install date, last reboot
  Set colItems = wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem")
  For Each objItem In colItems
    sOSCaption    = objItem.Caption
    sOSversion    = objItem.Version
    sArchitec     = objItem.OSArchitecture
    sInstallDate  = objItem.InstallDate
    sLastBoot     = objItem.LastBootUpTime
  Next
  
  ' Uptime
  sUpTime = DateDiff("s", CDate(FormatSystemDate(sLastBoot)), Now())
  
  ' Auto-Logon
  Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
  reg.GetStringValue HKLM,sALKeyPath,sALKeyValue,sALKeyData
  If Not IsNull(sALKeyData) Then
    If sALKeyData = 1 Then
      bAutoLogon = True
    End If
  End If
  
  ' SOE Version
  reg.GetStringValue HKLM,sSOEKeyPath,sSOEKeyValue,sSOEKeyData
  If Not IsNull(sSOEKeyData) Then
    If Trim(sSOEKeyData) <> "" Then sSOEVersion = sSOEKeyData
  End If
End If

'WScript.Echo "Logon Domain: " & vbTab & vbTab & sDomain
'WScript.Echo "Logon User: " & vbTab & vbTab & sUser
'WScript.Echo "Computer: " & vbTab & vbTab & sComputer
'WScript.Echo "Site: " & vbTab & vbTab & vbTab & sSite
'WScript.Echo "Computer IP: " & vbTab & vbTab & sIP
'WScript.Echo "Model: " & vbTab & vbTab & vbTab & sModel
'WScript.Echo "OS Caption: " & vbTab & vbTab & sOSCaption
'WScript.Echo "OS Version: " & vbTab & vbTab & sOSversion
'WScript.Echo "OS Architec: " & vbTab & vbTab & sArchitec
'WScript.Echo "Install Date: " & vbTab & vbTab & FormatSystemDate(sInstallDate)
'WScript.Echo "Last Reboot: " & vbTab & vbTab & FormatSystemDate(sLastBoot)
'WScript.Echo "System UpTime: " & vbTab & vbTab & sUpTime & " Seconds"
'WScript.Echo "System UpTime: " & vbTab & vbTab & FormatUpTime(sUpTime)
'WScript.Echo "Auto-Logon: " & vbTab & vbTab & bAutoLogon
'WScript.Echo "SOE Version: " & vbTab & vbTab & sSOEVersion

'WScript.Echo ""
'WScript.Echo "Connecting to database..."



'=====================================
Set objConn = CreateObject("ADODB.Connection")
Set objCmd  = CreateObject("ADODB.Command")


objConn.Open strConnVMMH1CSQL040
'WScript.Echo "Add new record..."
With objCmd   
  .ActiveConnection = objConn
  .CommandText      = "NewEntry"
  .CommandType      = adCmdStoredProc
  .Parameters.Append .CreateParameter("@type", adVarChar, adParamInput, 50, sType)
  .Parameters.Append .CreateParameter("@domain", adVarChar, adParamInput, 50, sDomain)
  .Parameters.Append .CreateParameter("@user", adVarChar, adParamInput, 50, sUser)
  .Parameters.Append .CreateParameter("@computer", adVarChar, adParamInput, 50, sComputer)
  .Parameters.Append .CreateParameter("@site", adVarChar, adParamInput, 50, sSite)
  .Parameters.Append .CreateParameter("@ip", adVarChar, adParamInput, 50, sIP)
  .Parameters.Append .CreateParameter("@model", adVarChar, adParamInput, 100, sModel)
  .Parameters.Append .CreateParameter("@os", adVarChar, adParamInput, 100, sOSCaption)
  .Parameters.Append .CreateParameter("@version", adVarChar, adParamInput, 50, sOSversion)
  .Parameters.Append .CreateParameter("@architec", adVarChar, adParamInput, 50, sArchitec)
  .Parameters.Append .CreateParameter("@install", adDate, adParamInput, 50, FormatSystemDate(sInstallDate))
  .Parameters.Append .CreateParameter("@reboot", adDate, adParamInput, 50, FormatSystemDate(sLastBoot))
  .Parameters.Append .CreateParameter("@autologon", adBoolean, adParamInput, , bAutoLogon)
  .Parameters.Append .CreateParameter("@soeversion", adVarChar, adParamInput, 50, sSOEVersion)
  .Execute , , adExecuteNoRecords
End With
'WScript.Echo "Close connection to DB..."

objConn.Close

'=====================================
' === 25/11/2016 ===
Set fso     = CreateObject("Scripting.FileSystemObject")
Set shell   = CreateObject("WScript.Shell")
Set envproc = shell.Environment("Process")
sLogonServer = envproc("LOGONSERVER")

If fso.FileExists(sLogonServer & "\netlogon\scripts\NRDHBUserPrinters.vbs") Then
  shell.Run "CSCRIPT.EXE " & sLogonServer & "\netlogon\scripts\NRDHBUserPrinters.vbs", 0, False
End If

' === 14/09/2017 ===
If fso.FileExists(sLogonServer & "\netlogon\scripts\Delete_health.nz_cookie.vbs") Then
  shell.Run "CSCRIPT.EXE " & sLogonServer & "\netlogon\scripts\Delete_health.nz_cookie.vbs", 0, False
End If

Set fso     = Nothing
Set shell   = Nothing
Set envproc = Nothing
' === 25/11/2016 ===

Set objConn = Nothing
Set objCmd  = Nothing
Set network = Nothing
Set wmi     = Nothing
Set reg     = Nothing
Set objADSysInfo = Nothing

'=============================================================================================
Function FormatSystemDate(strDate)
'=============================================================================================
  'accept the system return time formate 20160212071456.0000+780
  If strDate <> "" Then
    Dim sY, sM, sD, sH, sN, sS
    sY = Mid(strDate, 1, 4)
    sM = Mid(strDate, 5, 2)
    sD = Mid(strDate, 7, 2)
    sH = Mid(strDate, 9, 2)
    sN = Mid(strDate, 11, 2)
    sS = Mid(strDate, 13, 2)
    FormatSystemDate = sD & "/" & sM & "/" & sY & " " & sH & ":" & sN & ":" & sS
  End If
End Function

'=============================================================================================
Function FormatUpTime(intTimeInSec)
'=============================================================================================
 'accept integer value in seconds
 If intTimeInsec <> 0 Then
   Dim sD, sH, sM, sS, sDD, sHH, sMM, sSS
   sD = CStr(intTimeInSec /(60*60*24))
   'MsgBox "sD=" & sD
   If InStr(sD, ".") = 0 Then
    sH = 0
    sM = 0
    sS = 0
   Else
    sD = Int(sD)
    sH = CStr((intTimeInSec Mod (60*60*24))/(60*60))
    'MsgBox "sH=" & sH
    If InStr(sH, ".") = 0 Then
      sM = 0
      sS = 0
    Else
      sH = Int(sH)
      sM = Cstr(((intTimeInSec Mod (60*60*24)) Mod (60*60))/60)
      If InStr(sM, ".") = 0 Then
        sS = 0
      Else
        sM = Int(sM)
        sS = ((intTimeInSec Mod (60*60*24)) Mod (60*60)) Mod 60
      End If
    End If
   End If
   
   If sD > 1 Then
    sDD = "Days"
   Else
    sDD = "Day"
   End If
   If sH > 1 Then
    sHH = "Hours"
   Else
    sHH = "Hour"
   End If
   If sM > 1 Then
    sMM = "Minutes"
   Else
    sMM = "Minute"
   End If
   If sS > 1 Then
    sSS = "Seconds"
   Else
    sSS = "Second"
   End If
    
   FormatUpTime = sD & " " & sDD & " " & sH & " " & sHH & " " & sM & " " & sMM & " " & sS & " " & sSS
 End If
End Function