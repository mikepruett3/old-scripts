set objShell = WScript.CreateObject("WScript.Shell")
set colSystemEnvVars = objShell.Environment("Process")
WScript.Echo colSystemEnvVars("COMPUTERNAME")