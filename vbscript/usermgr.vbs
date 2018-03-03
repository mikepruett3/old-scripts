Option Explicit         ' Force variable declarations for safety!

' Defination & Declaration
Dim numPass,myLength,trgDomain,trgUsername
Const minLength = 8	' minimum password length
Const maxLength = 20	' maxmum password length
Const UF_LOCKOUT = &H0010 'hex value for account lockout
numPass = 5		' number of passwords required
myLength = 8		' required password length
trgDomain = "ISTAADS"	' Default DomainName

' Main Program
UsrName(trgUsername)
WScript.Echo trgUsername
LockStatus(trgUsername)

'Do Until Count = numPass
'	Count = Count + 1
'	RandomPW(newPW)
'	sFirstChar = Left(newPW,1)
'	sfsf = IsNumeric(sFirstChar)
'	If sfsf = -1 Then
'		Count = Count - 1
'	else
'		WScript.Echo newPW
'	End If
'Loop

' Asks user to enter the Target Username
Function UsrName(dumb) 'Placeholder variable, lack of better name.
	Dim trgUN,cusrdom,trgUsername
	trgUN = Inputbox("Type in the target username:")
	cusrdom=Msgbox("Does this look correct?" & vbcrlf & vbcrlf & vbcrlf & trgDomain & "\" & trgUN,vbOKCancel)
		If cusrdom = 1 then ' The returing value of 1 equals yes.
			msgbox("Correct!")
			trgUsername = trgUN
		else
			UsrName(dumb)
		End If
End Function

' Random Password Generator Function
' Original Author: Carl Mercier (info@carl-mercier.com)
Function RandomPW(dumb) 'Placeholder variable, lack of better name.
	Dim X, Y, strPW
	If myLength = 0 Then
		Randomize
		myLength = Int((maxLength * Rnd) + minLength)
	End If
	For X = 1 To myLength
		'Randomize the type of this character
		Y = Int((3 * Rnd) + 1) '(1) Numeric, (2) Uppercase, (3) Lowercase
		Select Case Y
			Case 1
				'Numeric character
				Randomize
				strPW = strPW & CHR(Int((9 * Rnd) + 48))
			Case 2
				'Uppercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 65))
			Case 3
				'Lowercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 97))
		End Select
	Next
	NewPW = strPW
End Function

Function LockStatus(strUsername)
	Dim objUser,cUserFlags
	WScript.Echo strUsername
	set objUser = GetObject("WinNT://" & "TRUNKS" & "/" & trgUsername & ",User")
	cUserFlags = objUser.Get("UserFlags")
	WScript.Echo cUserFlags
End Function
