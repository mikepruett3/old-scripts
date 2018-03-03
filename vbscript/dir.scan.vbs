Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder("K:\")
For Each file In f.Files
MsgBox file.Name
Next