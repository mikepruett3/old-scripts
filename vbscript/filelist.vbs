' **************************************
' Script: filelist.vbs
' Created by: Doug Cranston
' Modified by: Mike Pruett
'
' Action: Script Creates a file called
' "DirList.txt' in the specified directory,
' that list's the files & folders of the
' specified path
'
' **************************************

Option Explicit

Dim oFileSys,fh1,Dir,IntDoIt,OutputFile
Dim L_Welcome_MsgBox_Message_Text
Dim L_Welcome_MsgBox_Title_Text

dir = InputBox("Please type in the target directory.")
OutputFile = InputBox("Please type in the full path & filename for the report.")

Set oFileSys = CreateObject("Scripting.FileSystemObject")
L_Welcome_MsgBox_Message_Text ="Directory Listing Complete"
L_Welcome_MsgBox_Title_Text = "Directory Lister"

GetDir dir

intDoIt = MsgBox(L_Welcome_MsgBox_Message_Text, _
                      vbOKOnly + vbInformation,    _
                      L_Welcome_MsgBox_Title_Text )


sub GetDir(dir)
dim fh2,fh3,oFolder,oFolders,oFiles,item,Item2

set oFolder=oFileSys.GetFolder(dir)
set oFolders=oFolder.SubFolders
set oFiles=oFolder.Files

' get all sub-folders in this folder
For each item in oFolders
     GetDir(item)
Next
     item2=0
     For each item2 in oFiles
          set fh3=oFileSys.openTextFile(OutputFile,8,True)
          fh3.WriteLine(Dir & "\" & item2.Name)
          fh3.close     
     next

set fh2=oFileSys.openTextFile(OutputFile,8,True)
fh2.WriteLine(dir)
fh2.close

end sub
