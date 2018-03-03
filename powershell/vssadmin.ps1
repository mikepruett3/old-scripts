
param (
	[string]$path,
	[string]$file,
	[switch]$follow = $false
)

$ScriptName = $MyInvocation.InvocationName
$TargetDirectory = "C:\Test\"

Function Usage {
	Echo ""
	Echo "$ScriptName -path <Path\To\Be\Restored> -file <FileName>"
	Echo "or"
	Echo "$ScriptName -path <Path\To\Be\Restored> -follow"
	Echo ""
	Break
}

If ( $PSBoundParameters.Count -eq 0 ) {
	Usage
}

If ( $path -eq "" ) {
	Usage
}

If ( $follow -eq $true ) {
	If ( $file -eq "" ) {
		Echo ""
		Echo "Assuming the entire directory"
		Echo ""
	} else {
		Echo ""
		Echo "Cannot have the -file switch populated and the -follow switch"
		Echo "Choose one or the other..."
		Echo ""
		Break
	}
} else {
	If ( $file -eq "" ) {
			Usage
	}
}

$ShadowsReport = ShadowsList

Echo ""
Echo "Here is a list of current VSS Shadow Copy dates"
Echo ""
$x1=@{label="Date";Expression={$_.Key};alignment="left"}
$x2=@{label="Shadow Copy";Expression={$_.Value};FormatString="N3"}
Echo $ShadowsReport | Format-Table $x1,$x2 -autosize
Echo ""
Echo "Select Shadow Copy to Mount"
Echo ""
$Selection = Read-Host "Select a Shadow Copy Date?"
$ShadowCopyVolume = $ShadowsReport.Get_Item($Selection)
Echo ""
Echo "Linking Shadow to $Directory .."
Echo ""
$s1 = (Get-WMIObject -List Win32_ShadowCopy).Create("C:\", "ClientAccessible")
$s2 = Get-WMIObject Win32_ShadowCopy | ? { $_.ID -eq $s1.ShadowID }
$d  = $s2.DeviceObject + "\"
cmd /c mklink /d C:\shadowcopy "$d"
	
Echo $ShadowCopyVolume$Path
