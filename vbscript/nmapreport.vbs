Dim NmOpts,FSO,NmReportPathEnd,NmReportPathBeg,objFile,NmReport,strNmRep,wshShell,NmExec,CMD
Dim year , month , day , hour , min , logdt

NmTarget = WScript.Arguments.Item(0)
year  = DatePart("yyyy" , Now)
month = DatePart("m" , Now)
day   = DatePart("d" , Now)
hour  = DatePart("h" , Now)
min   = DatePart("n" , Now)
logdt = month & day & year & hour & min
NmReportPathBeg = "NMap"
NmReportPathEnd = ".txt"
NmExec = "C:\bin\nmap\nmap.exe"

strNmRep = "C:\"&NmReportPathBeg&"-"&NmTarget&"-"&logdt&NmReportPathEnd
NmOpts = "-A -T4 -sS"


On Error Resume Next
Set wshShell = CreateObject("wscript.shell")
CMD = NmExec & " " & NmOpts & " -oN "& strNmRep & " " & NmTarget
ExitCode = wshShell.Run(CMD, 1, TRUE)



