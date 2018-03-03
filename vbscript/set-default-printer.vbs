'*****************************************************
' This WSH script allows the user to quickly set
' the default printer.
'*****************************************************
Option Explicit

Dim Text, Title, i, j, tmp, printer, PrinterText
Dim WshNetwork, oDevices

Title = "SWYNK.com"

Set WshNetwork = WScript.CreateObject("WScript.Network")

Set oDevices = WshNetwork.EnumPrinterConnections

Text = "Listed below are the printers currently available to you.  Please enter the number of the printer you want set as the default." & vbCrLf & vbCrLf
j = oDevices.Count
For i = 0 To j - 1 Step 2
    Text = Text & (i/2) & vbTab
    Text = Text & oDevices(i) & vbTab & oDevices(i+1) & vbCrLf
Next


tmp = InputBox(Text, "Select default printer", 0)
If tmp = "" Then
    WScript.Echo "No user input, aborted"
    WScript.Quit
End If

tmp = CInt(tmp)
If (tmp < 0) Or (tmp > (j/2 - 1)) Then
    WScript.Echo "Wrong value, aborted"
    WScript.Quit
End If

printer = oDevices(tmp*2 + 1)

WshNetwork.SetDefaultPrinter printer

MsgBox "Your default printer has been successfully set to " & printer, _
        vbOKOnly + vbInformation, Title

'*** End