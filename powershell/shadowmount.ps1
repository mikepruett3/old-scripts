# Script: ShadowMount.ps1
#
# Usage: .\ShadowMount.ps1 -link <Link\Path>
#
# Notes: Calling this script with the Link switch, will create a Soft-Link
#        of a new "Client Accessable" System Restore Point in the 
#        specified directory.
#
#        Once done, just rmdir the directory. This will remove the Soft-Link

param (
	[string]$link
)

Function Usage {
	$ScriptName = $MyInvocation.InvocationName
	Echo ""
	Echo "$ScriptName -link <Link\Path>"
	Echo ""
  Echo "**NOTE** Directory must not exist **NOTE**"
	Echo ""
	Break
}

If ($link -eq "") {
	Usage
}
If ((Test-Path $link) -eq $True) {
	Usage
}
$Drive = $Env:SystemDrive + "\"
$CreateShadow = ( Get-WMIObject -List Win32_ShadowCopy ).Create( "$Drive", "ClientAccessible" )
$ShadowCopyVolume = Get-WMIObject Win32_ShadowCopy | ? { $_.ID -eq $CreateShadow.ShadowID }
$Directory = $ShadowCopyVolume.DeviceObject + "\"
cmd /c mklink /d "$link" "$Directory"
#$Test = Read-Host "You Ready"
#$LowerShadow = $CreateShadow.ShadowID.ToLower()
#cmd /c rmdir "$link"
#cmd /c vssadmin delete shadows /shadow=$LowerShadow