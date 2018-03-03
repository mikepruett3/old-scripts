Dim fso
Dim strCurrentDate
Dim strNewInfo

strNewInfo = InputBox("Enter Date *NO DASHES*: ")
Set fso = CreateObject("scripting.filesystemobject")
Call BuildCurrentDate
Call CopyFilesWithDateStamp ("c:\New-SPS", "c:\SPS")
Set fso = Nothing

Sub CopyFilesWithDateStamp(strSourceFolder, strDestinationFolder)

    Dim fsoFile
    Dim fsoFolder
    Dim fsoSubFolder
    Dim strFileNameFront
    Dim strFileExt
    Dim strNewFileName
   
    Set fsoFolder = fso.GetFolder(strSourceFolder)
    For Each fsoSubFolder In fsoFolder.SubFolders
        fso.CreateFolder strDestinationFolder & "\" & fsoSubFolder.Name
        CopyFilesWithDateStamp fsoSubFolder.Path, strDestinationFolder & "\" & fsoSubFolder.Name
    Next
    

    For Each fsoFile In fsoFolder.Files
        BreakFileName fsoFile.Name, strFileNameFront, strFileExt
        strNewFileName = strFileNameFront & "_" & strNewInfo & strFileExt
        fsoFile.Copy strDestinationFolder & "\" & strNewFileName
    Next
    
    Set fsoFile = Nothing
    Set fsoFolder = Nothing
    Set fsoSubFolder = Nothing
End Sub

Sub BreakFileName(strFullName, strFront, strExtension)

    Dim intPos
    
    strFront = strFullName
    strExtension = ""
    
    intPos = InStrRev(strFullName, ".")
    If intPos > 0 Then
        strFront = Left(strFullName, intPos - 1)
        strExtension = Mid(strFullName, intPos)
    End If

End Sub

Sub BuildCurrentDate

    strCurrentDate = Year(date()) 
    If (Month(date()) < 10) Then
        strCurrentDate = strCurrentDate & "0" & Month(date())
    Else
        strCurrentDate = strCurrentDate & Month(date())
    End If
    If (Day(date()) < 10) Then
        strCurrentDate = strCurrentDate & "0" & Day(Date())
    Else
        strCurrentDate = strCurrentDate & Day(Date())
    End If

End Sub 