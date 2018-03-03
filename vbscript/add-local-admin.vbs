Dim DomainName
Dim UserAccount
Set net = WScript.CreateObject("WScript.Network")
local = net.ComputerName

DomainName = InputBox("Enter Domain Name:")
UserAccount = InputBox("Enter User Name to be added as Local Admin:")

set group = GetObject("WinNT://"& local &"/Administrators")

on error resume next
group.Add "WinNT://"& DomainName &"/"& UserAccount &""
CheckError

sub CheckError
if not err.number=0 then
set ole = CreateObject("ole.err")
MsgBox ole.oleError(err.Number), vbCritical
err.clear
else
MsgBox "Done."
end if
end sub