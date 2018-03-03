

Set wshShell = CreateObject("WScript.Shell")
Dim strArchitecture, strProgramFiles, strReaderVersion
Dim strAdobeAnnotationsPath, strAdobeJavaScriptsPath
strArchitecture = wshShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
strArchitecture = LCase(strArchitecture)
Select Case strArchitecture
	Case "amd64"
			strProgramFiles = wshShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
	Case "x86"
			strProgramFiles = wshShell.ExpandEnvironmentStrings("%ProgramFiles%")
	Case Else
			MsgBox "Unknown Processor Architecture Type...",vbExclamation
			WScript.Quit
End Select

' Nice Way to pull Acrobat Version from Registry
' Borrwed from http://us.generation-nt.com/answer/get-version-installed-acrobat-reader-help-96282462.html
With CreateObject("WScript.Shell")
strReaderVersion = .RegRead("HKCR\" & .RegRead("HKCR\.pdf\") & "\AcrobatVersion\")
End With
strAdobePath = strProgramFiles & "\Adobe\Acrobat " & strReaderVersion

Select Case PathExists(strAdobePath)
	Case True
			strAdobeAnnotationsPath = strAdobePath & "\Acrobat\plug_ins\Annotations\Stamps\ENU"
			strAdobeJavaScriptsPath = strAdobePath & "\Acrobat\Javascripts"
	Case Else
			MsgBox "Adobe Acrobat Path cannot be located...",vbExclamation
			WScript.Quit
End Select

Select Case PathExists(strAdobeAnnotationsPath)
	Case True
			strPlugin1Copy = Copy2Path("SignHere.pdf","\\dispense\Deploy\ZincStamp",strAdobeAnnotationsPath)
		Case Else
			MsgBox "Acrobat Annotations path cannot be located...",vbExclamation
			WScript.Quit
End Select

Select Case PathExists(strAdobeJavaScriptsPath)
	Case True
			strPlugin2Copy = Copy2Path("ZincMAPS10.js","\\dispense\Deploy\ZincStamp",strAdobeJavaScriptsPath)
	Case Else
			MsgBox "Acrobat JavaScripts path cannot be located...",vbExclamation
			WScript.Quit
End Select

'If (strPlugin1Copy = True) AND (strPlugin2Copy = True) Then
'	MsgBox "The ZincStamp plugins for Adobe Acrobat have been installed sucessfully!",vbInformation
'Else
'	MsgBox "Not all files have been installed, please correct these errors and re-run this script.",vbCritical
'End If

'Script End
Set wshShell = Nothing
Set strArchitecture = Nothing
Set strProgramFiles = Nothing

'Function to check for Path Existence
Function PathExists(strWorkingPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(strWorkingPath) Then
		' The path does not exist
		Set fso = Nothing
		PathExists = False
	Else
		' The path does exist
		Set fso = Nothing
		PathExists = True
	End If
End Function

'Function to copy file into target path
Function Copy2Path(strFileName,strFileLocation,strTargetPath)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not fso.FileExists(strTargetPath & "\" & strFileName) Then
		'The file does not already exist in the target path
		fso.CopyFile strFileLocation & "\" & strFileName, strTargetPath & "\" & strFileName
		Set fso = Nothing
		Copy2Path = True
	Else
		'The file allready exists in the target path
		'We are still going to copy the file, we are just going to overwrite it...
		fso.CopyFile strFileLocation & "\" & strFileName, strTargetPath & "\" & strFileName, True
		Set fso = Nothing
		MsgBox "The file " & strFileName & " already exists in that location...",vbExclamation
		Copy2Path = False
	End If
End Function