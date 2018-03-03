' http://www.tek-tips.com/viewthread.cfm?qid=1060083&page=1
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_bios.asp
Const ForAppending = 8
Const F_RDON  = 1
On Error Resume Next

set fsi = CreateObject("Scripting.FileSystemObject")
Set fd = fsi.OpenTextFile("comp.txt", F_RDON)

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("info.txt", ForAppending, True)

objTextFile.WriteLine("Computer" & vbTab & "Manufacturer" & vbTab & "Name" & vbTab & "SerialNumber" & vbTab & "Version")

err.clear
Do while Not fd.AtEndOfStream
strComputer = fd.ReadLine
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    if err.number <>0 then
        objTextFile.WriteLine strComputer + err.description
        err.clear
    else
	Set colBIOS = objWMIService.ExecQuery("Select * from Win32_BIOS")
		For each objBIOS in colBIOS
			objTextFile.WriteLine(strComputer & vbTab & objBIOS.Manufacturer & vbTab & objBIOS.Name & vbTab & objBIOS.SerialNumber & vbTab & objBIOS.Version)
		Next
	set colBIOS=nothing
    end if
loop

objtextfile.close
objtextfile.close