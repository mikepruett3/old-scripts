Dim NameSpaceObj
Dim DomObj

Set NameSpaceObj  = GetObject("WinNT:")
NameSpaceObj.Filter = Array("domains")

For Each DomObj in NameSpaceObj
    WScript.Echo DomObj.Name & "," & DomObj.Class
Next
