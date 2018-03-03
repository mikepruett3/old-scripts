' --------------------------------------------------------------------
' - Script: mbfix1a.vbs
' -
' - Usage: cscript c:\somedirectory\mbfix1a.vbs
' -
' - Description: Script is a upgrade from the 1.0 script that modified 2 
' - registry settings at the local machine, to disable "Master Browser"
' - issues on a LAN with a PDC. Script will scan the services, looking
' - for "Associated" services of the "Computer Browser" service. If any
' - are located, they will be stopped. Next portion stops & disables the
' - Computer Browser service itself. Finally, a new set of registry keys
' - are updated.
' -
' - Background: The original script was created to eliminate the "Master
' - browser issues on our LAN.
' -
' - Changelog:
' -
' - 1.0a - Included stopping of "Browser" service & Associates.
' - 	 - Corrected "MaintainServerList" setting.
' -
' - 1.0	 - Inital Script
' ---------------------------------------------------------------------

strComputer = "."

Set colComputers = objWMIService.ExecQuery ("SELECT * FROM Win32_ComputerSystem")

For Each objComputer in colComputers
	Select Case objComputer.DomainRole 'Determine what "Domain" role target workstation has.
		Case 0
		    strComputerRole = "Standalone Workstation"
		    ApplyKeys
		    ServStop
		Case 1
		    strComputerRole = "Member Workstation"
		    ApplyKeys
		    ServStop
		Case 2
		    strComputerRole = "Standalone Server"
		    ApplyKeys
		    ServStop
		Case 3
		    strComputerRole = "Member Server"
		    ApplyKeys
		    ServStop
		Case 4
		    strComputerRole = "Backup Domain Controller"
 	    	    WScript.Echo "Cannot Apply to Backup Domain Controllers."
		Case 5
		    strComputerRole = "Primary Domain Controller"
		    WScript.Echo "Cannot Apply to Primary Domain Controllers."
	End Select
Next

Sub ServStop 'Stops the "Computer Browser" service & its Associates, then disables "CB" service.
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colServiceList = objWMIService.ExecQuery("Associators of " & "{Win32_Service.Name='Browser'} Where " & "AssocClass=Win32_DependentService " & "Role=Antecedent" )
For each objService in colServiceList
    objService.StopService()
Next
Wscript.Sleep 5000 'Pause required to let service shutdown gracefully.
Set colServiceList = objWMIService.ExecQuery ("Select * from Win32_Service where Name='Browser'")
For each objService in colServiceList
    errReturn = objService.StopService()
    errReturnCode = objService.Change( , , , , "Disabled")
Next
End Sub

Sub ApplyKeys 'Updates 2 Registry Keys to remove "Browser" issues.
KeyOne  = "HKLM\SYSTEM\CurrentControlSet\Services\Browser\Parameters\IsDomainMaster"
KeyTwo  = "HKLM\SYSTEM\CurrentControlSet\Services\Browser\Parameters\MaintainServerList"
ISTAKey = "HKEY_LOCAL_MACHINE\SOFTWARE\ISTAPHARMA\MBFIX1a"
Set objShell = WScript.CreateObject ("WScript.Shell")
objShell.RegWrite KeyOne, "FALSE", "REG_SZ"
objShell.RegWrite KeyTwo, "No", "REG_SZ"
objShell.RegWrite ISTAKey, 1, "REG_SZ"
End Sub
