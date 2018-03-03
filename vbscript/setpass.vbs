'~~Author~~.        Adrian Grigorof
'~~Email_Address~~. adrian@altairtech.net
'~~Script_Type~~.   vbscript
'~~Sub_Type~~.      DomainAdministration
'~~Keywords~~.      local user, password, change, adsi 

'~~Comment~~.
' This script can be used to remotely change the password for local users on several NT workstations or NT servers. It requires that the the account used to submit this script has the rights to  change local passwords (typically, a member of the Domain Admins group will have this right) The script requires Windows Scripting Host and ADSI installed on the computer used to run it.

'~~Script~~.

' This script can be used to remotely change the password for local users on several NT workstations or
' NT servers. It requires that the the account used to submit this script has the rights to 
' change local passwords (typically, a member of the Domain Admins group will have this right)
' The script requires Windows Scripting Host and ADSI installed on the computer used to run it.
' Both can be downloaded for free from Microsoft or other scripting sites.
' The script is using a list of computers as input file. The list of computers has to look something
' like:
'
' computer1
' computer2
' ...
'
' If for some reason, one of the computers could not be accessed by the script, it will be saved in
' another file for later submission
' 
' The script will output a confirmation of the password change to the screen or a warning 
' message if the computer could not be contacted. DO NOT double click on the script or use wscript
' to lunch it unless you want a confirmation of each password change to pop-up on the screen. 
' Instead, use "cscript setpass.vbs" command.

' The names of the files, user id and new password have to be initialized (see below)

On Error Resume Next
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fso, fsox, fx
Dim inputFile, outputFile, computerIndex, myComputer,myUser, usr, mDSPath

' set inputFile to match the name of the text file with the list of computers to be changed
' set outputFile to match the name of the text file that will contain the names of the 
' computers that could not be accessed by the script (like powered off)
' later, this file can be used as inputFile to reissue the password change to the computers that
' were not available initially
' set myUser to the name of the local account that needs to have the password changed
' set newPassword to the new password for the user
inputFile = "computers.txt"
outputFile = "computers_na.txt"
myUser = "Administrator"
newPassword = "1sta1td3pt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("computers.txt", ForReading, True)
Set fsox = CreateObject("Scripting.FileSystemObject")
Set fx = fsox.OpenTextFile("Xcomputers.txt", ForWriting, True)
computerIndex = 1

Do While f.AtEndOfLine <> True
   myComputer = f.ReadLine
     mDSPath = "WinNT://" & myComputer & "/" & myUser & ",user"
   Set usr = GetObject(mDSPath)
     If Err Then 
          fx.WriteLine(myComputer)
       ' Comment out the next line if no output on the screen is required
          WScript.Echo CStr(computerIndex) & ". " & myComputer & " could not be contacted"
          Err.Clear
     Else
         usr.SetPassword newPassword
         ' Comment out the next line if no output on the screen is required
          WScript.Echo CStr(computerIndex) & ".User: " & myComputer & "\" & myUser & ": password changed"
     End If
     computerIndex = computerIndex + 1     
Loop
f.Close
fx.Close 