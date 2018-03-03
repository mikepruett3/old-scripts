# Script: UserBackup.ps1
#
# Usage: .\UserBackup.ps1
#
# Notes: Calling this script will create a Soft-Link of a new "Client Accessable"
#        System Restore Point, and then proceed to backup newer files from the
#        specified directory.
#
#        Once done, the script will remove the Soft-Link and the ShadowCopy

$UP = $Env:UserProfile
$Profile = $UP.SubString( $UP.IndexOf(":") + 1 )
$7Zip = "C:\Users\Torafuma\Bin\7za.exe"
$BackupLocation = "$UP\Backups"

#Create Temporary Directory
$TimeStamp = get-date -f yyyy-MM-dd_hhmm
$TMP = "$Env:Temp\$TimeStamp\"
If ((Test-Path $TMP) -eq $True) {
	Echo ""
	Echo "Directory Allready Exists"
	Break
} else {
	Echo ""
	Echo "Creating Temporary Directory"
	New-Item -ItemType Directory -Path $TMP | Out-Null
}

# Read information about last backup
[XML]$Backup = Get-Content $PWD\backup.xml
$LastBackup = $Backup.Archive.LastDate
$LastLocation = $Backup.Archive.LastLocation

# Create Soft-Link
$Drive = $Env:SystemDrive + "\"
$Link = $Env:SystemDrive + "\TempLink"
If ((Test-Path $Link) -eq $True) {
	Echo ""
	Echo "Directory Allready Exists"
	Break
} else {
	Echo ""
	$CreateShadow = ( Get-WMIObject -List Win32_ShadowCopy ).Create( "$Drive", "ClientAccessible" )
	$ShadowCopyVolume = Get-WMIObject Win32_ShadowCopy | ? { $_.ID -eq $CreateShadow.ShadowID }
	$Directory = $ShadowCopyVolume.DeviceObject + "\"
	cmd /c mklink /d "$Link" "$Directory"
}

# Find Files and Copy them
Echo ""
Echo "Locating and Coping Files"
Get-ChildItem "$Link$Profile\Documents" -Recurse | Where-Object {$_.Mode -notmatch "d"} | Where-Object {$_.LastWriteTime -ge $LastBackup} | Copy-Item -Destination $TMP
Get-ChildItem "$Link$Profile\Desktop" -Recurse | Where-Object {$_.Mode -notmatch "d"} | Where-Object {$_.LastWriteTime -ge $LastBackup} | Copy-Item -Destination $TMP
Get-ChildItem "$Link$Profile\Pictures" -Recurse | Where-Object {$_.Mode -notmatch "d"} | Where-Object {$_.LastWriteTime -ge $LastBackup} | Copy-Item -Destination $TMP
Get-ChildItem "$Link$Profile\Scripts" -Recurse | Where-Object {$_.Mode -notmatch "d"} | Where-Object {$_.LastWriteTime -ge $LastBackup} | Copy-Item -Destination $TMP
cmd /c $7Zip a "$BackupLocation\$TimeStamp.zip" "$TMP\*.*"
Remove-Item -Force -Recurse "$TMP"

# Remove Soft-Link and Delete ShadowCopy
Echo ""
Echo "Removing Soft-Link and ShadowCopy"
$LowerShadow = $CreateShadow.ShadowID.ToLower()
cmd /c rmdir "$link"
cmd /c vssadmin delete shadows /shadow=$LowerShadow /quiet