' - Title: Random Password Generator v1.1
' - Original Author: Carl Mercier (info@carl-mercier.com)

' Defination & Declaration
numPass = 5		' number of passwords required
myLength = 8		' required password length
Const minLength = 8	' minimum password length
Const maxLength = 20	' maxmum password length

' Main Program
Do Until Count = numPass
	Count = Count + 1
	RandomPW(newPW)
	sFirstChar = Left(newPW,1)
	sfsf = IsNumeric(sFirstChar)
	If sfsf = -1 Then
		Count = Count - 1
	else
		WScript.Echo newPW
	End If
Loop

' Random Password Generator Function
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