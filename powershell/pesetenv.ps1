
$Architecture = $Env:Processor_Architecture
$ProgramFiles = $Env:ProgramFiles

$WAIKInstallation = "$ProgramFiles\Windows AIK"
If (!(Test-Path -Path $WAIKInstallation)) {
	Write-Host ""
	Write-Host ("Folder does not exist!!") -Fore 'Red'
	Write-Host ""
	Break
}

Write-Host ""
Write-Host("Updating path to include dism, oscdimg, imagex...") -Fore 'Blue'
Write-Host ""

#Ensure WAIK PETools are in the Path
#This does the same as "Windows AIK\Tools\PETools\pesetenv.cmd"

Write-Host ""
Write-Host("Checking and updating path if needed...") -Fore 'Yellow'
Write-Host ""

$Path += $WAIKInstallation + "\Tools\PETools" + ";"

if ($Architecture -ne "x86"){
	$Path += $WAIKInstallation + "\Tools\x86;"
	$Path += $WAIKInstallation + "\Tools\x86\Servicing;"
}

$Path += $WAIKInstallation + "\Tools\" + $Architecture + ";"
$Path += $WAIKInstallation + "\Tools\" + $Architecture + "\Servicing;"

If (!$Env:Path.ToLower().Contains($Path.ToLower())){
	[System.Environment]::SetEnvironmentVariable("PATH", $Path + $Env:Path, "process")
}
