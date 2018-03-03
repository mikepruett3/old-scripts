Dim NameSpaceObj
Dim DomObj

Set NameSpaceObj  = GetObject("WinNT://ISTAADS")
NameSpaceObj.Filter = Array("user")

For Each DomObj in NameSpaceObj
    WScript.Echo DomObj.Name & "," & DomObj.Class
Next
