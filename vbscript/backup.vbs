' ---------------------------------------------------------------------------
' Title:		Client Network Backup Script
'
' Script: 		Backup.vbs
' Author: 		Mike Pruett		||		amanoj <at> gmail <dot> com
' Version:		v2.1			||		July 28th, 2006 @ 6:34PM
' Desc.:		This script is an evolution of a script was originally made
'				in DOS/Bash script. Script will backup any files from the
'				Source directory, as long as the file matches an extension in
'				the inclusion list. These files are then copied over to the
'				desired network share. NOTE!! In order for the script to work
'				This share must be accessable from the local machine. (Must
'				show up in My Computer as a	drive letter.) The script is not
'				too picky, as it does not require Block Level access to	the 
'				Drive. (NAS and SMB file sharing OK!!)
'
' Insructions:	Just modify the needed var's and your off.
' ---------------------------------------------------------------------------
Option Explicit

Dim arrayPathElement,arrayIncElement
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim oShell : Set oShell = CreateObject("wscript.shell")

Dim year : year = DatePart("yyyy" , Now)
Dim month : month = DatePart("m" , Now)
Dim day : day = DatePart("d" , Now)
Dim hour : hour = DatePart("h" , Now)
Dim min : min = DatePart("n" , Now)
Dim logdt : logdt = month & day & year & "-" & hour & min
Dim strUserName : strUserName = oShell.ExpandEnvironmentStrings("%USERNAME%")

Dim oLog : Set oLog = oFSO.CreateTextFile("I:\backup_" & logdt & ".txt",TRUE) ' Change the name and location of Backup Log.
Dim sDestinationPath : sDestinationPath = "I:\" ' Target Directory. (Must be a drive letter!!)
Dim sSourcePath : sSourcePath = "C:\Users\" & strUsername & "\"
Dim strStartPath : strStartPath = "C:\Users\" & strUserName ' Where to start searching for files.

' array for extensions to search for. Modify to suite application.
Dim arrayInclude : arrayInclude = array(".pdf",".doc",".dot",".xls",".xlt",".ppt",".mpp",".vsd",".jpg",".bmp",".gif",".txt",".url",".html",".htm",".zip",".mht",".max")

' array for folders an paths. You can replace with your own.
Dim arrayPath : arrayPath = array(strStartPath & "\My Documents\",strStartPath & "\Favorites\",strStartPath & "\Desktop\")

oLog.WriteLine "---------------------- Backup Log ----------------------"
oLog.WriteLine vbcrlf
oLog.WriteLine "Backup Date: " & day & "-" & month & "-" & year & " @ " & hour & ":" & min
oLog.WriteLine vbcrlf
oLog.WriteLine "--------------------------------------------------------"

For Each arrayPathElement in arrayPath
	FindFile(arrayPathElement)
	CheckFolder(arrayPathElement)
Next

oLog.WriteLine vbcrlf
oLog.WriteLine "--------------------------------------------------------"

Set oFSO = Nothing
Set oLog = Nothing
Set oShell = Nothing
WScript.Quit

Sub CheckFolder(objCurrentFolder)
	Dim colSelFold, oFolder, strFolder
	strFolder = objCurrentFolder
	Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set colSelFold = oFSO.GetFolder(strFolder)
    For Each oFolder in colSelFold.SubFolders
    	FindFile(oFolder)
    	CheckFolder(oFolder)
    Next
    Set oFSO = Nothing
End Sub

Sub FindFile(strCurrentFolder)
	Dim oTrgFolder,strFile
	Set oTrgFolder = oFSO.GetFolder(strCurrentFolder)
	For Each strFile in oTrgFolder.Files
		For Each arrayIncElement in arrayInclude
			If InStr(LCase(strFile.Path),arrayIncElement) <> 0 Then
				oLog.WriteLine "File to be processed: " & strFile.Path
				Call CopyFiles(strFile.Path, sSourcePath, sDestinationPath)
			End If
		Next
	Next
End Sub

Function CopyFiles(sSourceFile, sSourcePath, sDestinationPath)
	Dim oSourceFile
	Dim oRegExp, colMatches, oMatch
	Dim sTreePath, sFName
	Dim sSourcePattern
	Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
	
	If (Right(sDestinationPath, 1) = "\") Then sDestinationPath = Left(sDestinationPath, Len(sDestinationPath) - 1)
	If (Right(sSourcePath, 1) = "\") Then sSourcePath = Left(sSourcePath, Len(sSourcePath) - 1)
	
	sSourcePattern = EscapeRegex(sSourcePath)
	Set oRegExp = New RegExp
	oRegExp.IgnoreCase = True
	oRegExp.Pattern = "^" & sSourcePattern & "\\?(.*)\\([^\r\n]+)$"
	
	If (Not oFSO.FolderExists(sDestinationPath)) Then oFSO.CreateFolder(sDestinationPath)
	sTreePath = ""
	sFName = ""
	Set colMatches = oRegExp.Execute(sSourceFile)
	For Each oMatch In colMatches
		sTreePath = oMatch.SubMatches(0)
		sFName = oMatch.SubMatches(1)
	Next
	If (sFName = "") Then ' This file does not reside in sSourcePath at all, or other parsing error.
			oLog.WriteLine "File '" & sSourceFile & "' not copied." & vbCrLf & "  Error parsing file name, or does not reside in source path '" & sSourcePath & "'"
		Else
			If (sTreePath <> "") Then ' This file resides in a subfolder of sSourcePath.
				If (Not VerifyFolder(sDestinationPath, sTreePath)) Then
					oLog.WriteLine "File '" & sSourceFile & "' not copied." & VbCrLf & "  Could not create / verify folder: " & sDestinationPath & "\" & sTreePath
				Else
					On Error Resume Next
					oFSO.CopyFile sSourceFile, sDestinationPath & "\" & sTreePath & "\"
					If (Err.Number) Then
						oLog.WriteLine "File '" & sSourceFile & "' not copied." & vbCrLf & "  Error: " & Err.Number & ", Description: " & Err.Description
					End If
					On Error Goto 0
				End If
			Else ' This file resides directly in sSourcePath
				On Error Resume Next
				oFSO.CopyFile sSourceFile, sDestinationPath & "\"
				If (Err.Number) Then
					oLog.WriteLine "File '" & sSourceFile & "' not copied." & vbCrLf & "  Error: " & Err.Number & ", Description: " & Err.Description
				End If
				On Error Goto 0
			End If
	End If
	Set oSourceFile = Nothing
	Set oFSO = Nothing
  	Set oRegExp = Nothing
End Function

Function EscapeRegex(s)
	EscapeRegex = ""
	If (IsNull(s)) Then s = "" Else s = CStr(s)
	Dim oRegEx : Set oRegEx = New RegExp
	oRegEx.Global = True
	oRegEx.Pattern = "([[\\^$.|?*+()])"
	EscapeRegex = oRegEx.Replace(s, "\$1")
	Set oRegEx = Nothing
End Function

Function VerifyFolder(sRoot, sPath)
	Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
	Dim nCurIndex : nCurIndex = Len(sRoot) + 2
	Dim sCompletePath : sCompletePath = sRoot & "\" & sPath
	Dim nPathLen : nPathLen = Len(sCompletePath) 
	
	Do While (nCurIndex <= nPathLen)
		Dim sTemp
		Dim nFoundIndex : nFoundIndex = InStr(nCurIndex, sCompletePath, "\")
		If (nFoundIndex <> 0) Then
			sTemp = Left(sCompletePath, nFoundIndex)
			nCurIndex = nFoundIndex + 1
		Else
			sTemp = sCompletePath
			nCurIndex = nPathLen + 1
		End If
		
		If (Not oFSO.FolderExists(sTemp)) Then
			On Error Resume Next
			oFSO.CreateFolder(sTemp)
			
			If (Err.Number) Then
				oLog.WriteLine "Error creating folder: " & sTemp
				VerifyFolder = False
				Exit Function
			End If
			On Error Goto 0
		End If
	Loop
	Set oFSO = Nothing
	VerifyFolder = True
End Function
