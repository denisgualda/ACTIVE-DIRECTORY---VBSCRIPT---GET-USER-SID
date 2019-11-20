Option Explicit
Dim strUser
Dim siduser


strUser = CreateObject("WScript.Network").UserName
'CRIDEM LA FUNCIO strUser
siduser = GetSIDFromUser(strUser)
wscript.echo "SIDUSER:   " & siduser




'FUNCIO OBTÉ SIDUSER USUARI LOGUEJAT
'******************************************************************
Function GetSIDFromUser(UserName)
  Dim DomainName, Result, WMIUser
  If InStr(UserName, "\") > 0 Then
    DomainName = Mid(UserName, 1, InStr(UserName, "\") - 1)
    UserName = Mid(UserName, InStr(UserName, "\") + 1)
  Else
    DomainName = CreateObject("WScript.Network").UserDomain
    wscript.echo "domini: " & DomainName
  End If
  On Error Resume Next
  Set WMIUser = GetObject("winmgmts:{impersonationlevel=impersonate}!" _
    & "/root/cimv2:Win32_UserAccount.Domain='" & DomainName & "'" _
    & ",Name='" & UserName & "'")
  If Err.Number = 0 Then
    Result = WMIUser.SID
  Else
    Result = ""
  End If
  On Error GoTo 0
  GetSIDFromUser = Result
End Function
'******************************************************************