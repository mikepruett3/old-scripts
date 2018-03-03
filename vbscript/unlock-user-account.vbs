UserName = InputBox("Enter the user's login name that you want to unlock:")

DomainName = InputBox("Enter the domain name in which the user account exists:")

Set UserObj = GetObject("WinNT://"& DomainName &"/"& UserName &"")
If UserObj.IsAccountLocked = -1 then UserObj.IsAccountLocked = 0
UserObj.SetInfo

If err.number = 0 Then
    Wscript.Echo "The Account Unlock Failed.  Check that the account is, in fact, locked-out."
Else
    Wscript.Echo "The Account Unlock was Successful"
End if