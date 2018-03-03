# PowerShell Software Installer Script
#
# Download and Install using PowerShell Scripting
# Changed by using XML Parsing

$TempDir = $Env:TEMP
$SysRoot = $Env:SystemRoot

Function Stage ( [string]$URL, [string]$FileName ) {
	$Download = New-Object System.Net.WebClient
	$NewFile = Join-Path $TempDir $FileName
	Echo "Downloading $FileName from $URL"
	$Download.DownloadFile($URL,$NewFile)
}

#TODO: Fix This!!

Function Install ( [string]$FileName, [string]$iType, [string]$iArguments, [string]$iTransform  ) {
	$Installer = Join-Path $TempDir $FileName
	Echo "Installing $FileName"
	Switch ($iType) {
		EXE {
			#Echo "$Installer $iArguments"
			$Command = "$SysRoot\System32\cmd.exe /c $Installer $iArguments"
			#$Command = "$SysRoot\System32\cmd.exe /c `"psexec.exe -i -d -s $Installer $iArguments`""
			$Process = [WMICLASS]"\\.\ROOT\CIMV2:Win32_Process"
			$Process.Create($Command)
		}
		MSI {
			Echo "MSIEXEC /I $Installer $iArguments TRANSFORM=$iTransform"
		}
	}
}

[XML]$Install = Get-Content $PWD\vim.xml

$Ammount = $Install.Items.Item.Count
$Number = 0
While ($Number -lt $Ammount ) {
	$DownloadURL = $Install.Items.Item[$Number].DownloadURL
	$InstallerFile = $Install.Items.Item[$Number].InstallerFile
	$InstallerType = $Install.Items.Item[$Number].InstallerType
	$InstallArguments = $Install.Items.Item[$Number].InstallArguments
	$Transform = $Install.Items.Item[$Number].Transform
	#Stage $DownloadURL $InstallerFile
	Install $InstallerFile $InstallerType $InstallArguments $Transform
	$Number = $Number + 1
}



