Dim dirname,folder,filez,FileCount,message1,message2
Dim strPath,objFSO,objStream,objFolder,objItem,Header1,Header2

message1 = "Type in the name of the directory to scan:"&vbcrlf&vbcrlf&vbtab&"eg. C:\mp3"
message2 = "Type in the full path and report filename:"&vbcrlf&vbcrlf&vbtab&"eg. C:\DirList.txt"

dirname = InputBox(message1)
OutputFile = InputBox(message2)

strPath = dirname

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objStream = objFSO.Createtextfile(OutputFile,True)
Set objFolder = objFSO.GetFolder(strPath)

For Each objItem In objFolder.SubFolders
	Set folder = objFSO.GetFolder(objItem)
	FileCount = folder.Files.Count
	Header1 = "Directory Listing of "&strPath&"\"&objItem.Name
	objStream.writeline Header1
	Call GetFiles
	objStream.writeline vbcrlf
Next

Set objItem = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
MsgBox "Report Complete."&vbcrlf&vbcrlf&"Report located at: "&OutputFile

Sub GetFiles
For Each objItem In objFolder.Files
	Header2 = vbtab & "- " & objItem.Name
	objStream.writeline Header2
Next
End Sub