set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "calc"    
WScript.Sleep 100
WshShell.AppActivate "Calculator"
WScript.Sleep 200
WshShell.SendKeys "1{+}2~"
set WshShell = Nothing  'always deallocate after use...