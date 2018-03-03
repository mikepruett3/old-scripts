'RegSrch.vbs - Search Registry for input string and display results.
'© Bill James - wgjames@mvps.org
' revised 20 Apr 2001 (parses regfile ~3X faster)
' revised 13 Dec 2001 (added Regedit command line switch for Win2K/WindXP)

Option Explicit
Dim oWS : Set oWS = CreateObject("WScript.Shell")
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")

Dim sSearchFor
sSearchFor = InputBox("This script will search your Registry and find all " & _
             "instances of the search string you input."  & vbcrlf & vbcrlf & _
             "This search could take several minutes, so please be patient." & _
             vbcrlf & vbcrlf & "Enter search string (case insensitive) and " & _
             "click OK...", WScript.ScriptName & " " & Chr(169) & " Bill James")

If sSearchFor = "" Then Cleanup()

Dim StartTime : StartTime = Timer

Dim sRegTmp, sOutTmp, eRegLine, iCnt, sRegKey, aRegFileLines

sRegTmp = oWS.Environment("Process")("Temp") & "\RegTmp.tmp "
sOutTmp = oWS.Environment("Process")("Temp") & "\sOutTmp" & _
          Hour(Now) & Minute(Now) & Second(Now) & ".tmp "

oWS.Run "regedit /e /a " & sRegTmp, , True '/a enables export as Ansi for WinXP

With oFSO.OpenTextFile(sOutTmp, 8, True)
  .WriteLine("REGEDIT4" & vbcrlf & "; " & WScript.ScriptName & " " & _
    Chr(169) & " Bill James" & vbcrlf & vbcrlf & "; Registry search " & _
    "results for string " & Chr(34) & sSearchFor & Chr(34) & " " & Now & _
    vbcrlf & vbcrlf & "; NOTE: This file will be deleted when you close " & _
    "WordPad." & vbcrlf & "; You must manually save this file to a new " & _
    "location if you want to refer to it again later." & vbcrlf & "; (If " & _
    "you save the file with a .reg extension, you can use it to restore " & _
    "any Registry changes you make to these values.)" & vbcrlf)

  With oFSO.GetFile(sRegTmp)
    aRegFileLines = Split(.OpenAsTextStream(1, 0).Read(.Size), vbcrlf)
  End With

  oFSO.DeleteFile(sRegTmp)

  For Each eRegLine in aRegFileLines
    If InStr(1, eRegLine, "[", 1) > 0 Then sRegKey = eRegLine
    If InStr(1, eRegLine, sSearchFor, 1) >  0 Then
      If sRegKey <> eRegLine Then
        .WriteLine(vbcrlf & sRegKey) & vbcrlf & eRegLine
      Else
        .WriteLine(vbcrlf & sRegKey)
      End If
      iCnt = iCnt + 1
    End If
  Next

  Erase aRegFileLines

  If iCnt < 1 Then
    oWS.Popup "Search completed in " & FormatNumber(Timer - StartTime, 0) & " seconds." & _
              vbcrlf & vbcrlf & "No instances of " & chr(34) & sSearchFor & chr(34) & _
              " found.",, WScript.ScriptName & " " & Chr(169) & " Bill James", 4096
    .Close
    oFSO.DeleteFile(sOutTmp)
    Cleanup()
  End If
  .Close

End With

oWS.Popup "Search completed in " & FormatNumber(Timer - StartTime, 0) & " seconds." & _
          vbcrlf & vbcrlf & iCnt & " instances of " & chr(34) & sSearchFor & chr(34) & _
          " found." & vbcrlf & vbcrlf & "Click OK to open Results in WordPad.",, _
          WScript.ScriptName & " " & Chr(169) & " Bill James", 4096

oWS.Run "WordPad " & sOutTmp, 3, True

oFSO.DeleteFile(sOutTmp)

Cleanup()

Sub Cleanup()
  Set oWS = Nothing
  Set oFSO = Nothing
  WScript.Quit
End Sub
