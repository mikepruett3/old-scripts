UserName = InputBox("Enter the user's login name that you want to unlock:")

Set objUser = GetObject _
  ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
If objUser.IsAccountLocked = True then objUser.IsAccountLocked = False
objUser.SetInfo

If err.number = 0 Then
    Wscript.Echo "The Account Unlock Failed.  Check that the account is, in fact, locked-out."
Else
    Wscript.Echo "The Account Unlock was Successful"
End if