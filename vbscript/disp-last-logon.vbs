On Error Resume Next
Dim User
Dim UserName
Dim UserDomain
UserDomain = InputBox("Enter the name of the domain:")
UserName = InputBox("Enter the name of the user:")
Set User = GetObject("WinNT://" & UserDomain & "/" & UserName & ",user")
MsgBox "The last time " & UserName & " logged on was: " & vbCRLf & vbCRLf & User.LastLogin