' Copyright (c) Microsoft Corporation 2004
' File:       BlockXPSP2.vbs
' Contents:   Remotely blocks or unblocks the delivery of
' Windows XP SP2 from Windows Update web site or via Automatic
' Updates. 
' History:    8/20/2004   Peter Costantini   Created
' Version:    1.0

On Error Resume Next

' Define constants and global variables.
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "." ' Can be changed to name of remote computer.
strKeyPath = "Software\Policies\Microsoft\Windows\WindowsUpdate"
strEntryName = "DoNotAllowXPSP2"
dwValue = 1

' Handle command-line arguments.
Set colArgs = WScript.Arguments
If colArgs.Count = 0 Then
  ShowUsage
Else
  If colArgs.Count = 2 Then
    strComputer = colArgs(1)
  End If
' Connect with WMI service and StdRegProv class.
  Set objReg = GetObject _
   ("winmgmts:{impersonationLevel=impersonate}!\\" & _
     strComputer & "\root\default:StdRegProv")
  If Err = 0 Then
    If (LCase(colArgs(0)) = "/b") Or _
     (LCase(colArgs(0)) = "-b" ) Then
      AddBlock
    ElseIf (LCase(colArgs(0)) = "/u") Or _
     (LCase(colArgs(0)) = "-u") Then
      RemoveBlock
    Else
      ShowUsage
    End If
  Else
    WScript.Echo "Unable to connect to WMI service on " _
     & strComputer & "."
  End If
  Err.Clear
End If

'*************************************************************

Sub AddBlock

'Check whether WindowsUpdate subkey exists.
strParentPath = "SOFTWARE\Policies\Microsoft\Windows"
strTargetSubKey = "WindowsUpdate"
intCount = 0
intReturn1 = objReg.EnumKey(HKEY_LOCAL_MACHINE, _
 strParentPath, arrSubKeys)
If intReturn1 = 0 Then
  For Each strSubKey In arrSubKeys
    If strSubKey = strTargetSubKey Then
      intCount = 1
    End If
  Next
  If intCount = 1 Then
    SetValue
  Else
    WScript.Echo "Unable to find registry subkey " & _
     strTargetSubKey & ". Creating ..."
    intReturn2 = objReg.CreateKey(HKEY_LOCAL_MACHINE, _
     strKeyPath)
    If intReturn2 = 0 Then
      SetValue
    Else
      WScript.Echo "ERROR: Unable to create registry " & _
       "subkey " & strTargetSubKey & "."
    End If
  End If
Else
  WScript.Echo "ERROR: Unable to find registry path " & _
   strParentPath & "."
End If

End Sub

'*************************************************************

Sub SetValue

intReturn = objReg.SetDWORDValue(HKEY_LOCAL_MACHINE, _
 strKeyPath, strEntryName, dwValue)
If intReturn = 0 Then
  WScript.Echo "Added registry entry to block Windows XP " & _
   "SP2 deployment via Windows Update or Automatic Update."
Else
  WScript.Echo "ERROR: Unable to add registry entry to " & _
   "block Windows XP SP2 deployment via Windows Update " & _
   "or Automatic Update."
End If

End Sub

'*************************************************************

Sub RemoveBlock

intReturn = objReg.DeleteValue(HKEY_LOCAL_MACHINE, _
 strKeyPath, strEntryName)
If intReturn = 0 Then
  WScript.Echo "Deleted registry entry " & strEntryName & _
   ". Unblocked Windows XP SP2 deployment via Windows " & _
   "Update or Automatic Update."
Else
  WScript.Echo "Unable to delete registry entry " & _
   strEntryName & ". Windows XP SP2 deployment via " & _
   "Windows Update or Automatic Update is not blocked."
End If

End Sub

'*************************************************************

Sub ShowUsage

WScript.Echo "Usage:" & VbCrLf & _
 "  BlockXPSP2.vbs { /b | /u | /? } [hostname]" & VbCrLf & _
 "    /b = Block (deny) Windows XP Service Pack 2 " & _
 "deployment" & VbCrLf & _
 "    /u = Unblock (allow) Windows XP Service Pack 2 " & _
 "deployment" & VbCrLf & _
 "    /? = Show usage" & VbCrLf & _
 "    hostname = Optional. Name of remote computer. " & _
 "Default is local computer" & VbCrLf & _
 "Example:" & VbCrLf & _
 "  BlockXPSP2.vbs /b client1"

End Sub

