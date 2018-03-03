' --------------------------------------------------------
' - Local_Administrators_Removal
' - Script: localadm-rem.vbs
' - Usage: wscript localadm-rem.vbs
' - Author: Mike Pruett
' - 		amanoj <at> gmail <dot> com
' - Created: March 9th, 2006
' - Updated: March 17th, 2006
' - Revision: 1.0
' - Desc: This script was created to remove all users from
' - the Local Administrators group from each Workstation
' - on the selected domain. This will not remove the "Administrator" 
' - account, as it will fail.
' --------------------------------------------------------
Dim sComputer,oMember,oComputer ' As String
Dim oWshNet,oGroup,oWMIService ' AS Object
Set oWshNet = CreateObject("WScript.Network")
sComputer = oWshNet.ComputerName
Set oWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")
Set colComputers = oWMIService.ExecQuery ("SELECT * FROM Win32_ComputerSystem")
For Each oComputer in colComputers
	Select Case oComputer.DomainRole
		Case 0
		    strComputerRole = "Standalone Workstation"
		    RemoveLocal
		Case 1
		    strComputerRole = "Member Workstation"
		    RemoveLocal
		Case 2
		    strComputerRole = "Standalone Server"
		    'RemoveLocal
		Case 3
		    strComputerRole = "Member Server"
		Case 4
		    strComputerRole = "Backup Domain Controller"
 	    	    WScript.Echo "Cannot Apply to Backup Domain Controllers."
		Case 5
		    strComputerRole = "Primary Domain Controller"
		    WScript.Echo "Cannot Apply to Primary Domain Controllers."
	End Select
Next

Sub RemoveLocal
    On Error Resume Next
    Set oGroup = GetObject("WinNT://" & sComputer & "/Administrators")
    For Each oMember In oGroup.Members
	If oMember.Class = "User" Then ' remove the user from Administrators group
		oGroup.Remove oMember.ADsPath
	End If
    Next

    oGroup = ""
    Set oGroup = GetObject("WinNT://" & sComputer & "/Network Configuration Operators")
    oGroup.Add("WinNT://ISTAADS/Domain Users")
End Sub