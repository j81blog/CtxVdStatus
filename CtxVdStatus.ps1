<#
.NAME
	CtxVdStatus
.SYNOPSIS  
	Get info about the current running Virtual Desktops
.DESCRIPTION  
	Get info about the current running Virtual Desktops
.NOTES  
	File Name  : CtxVdStatus.ps1
	Author     : John Billekens - john@j81.nl
	Requires   : Citrix Delivery Controller 7.x
	Version    : 20150907.1108
.SYNTAX
	.\CtxVdStatus.ps1 [-CsvOutput "<path\file.csv>"][-GridView][-numThreads <5>]
.EXAMPLE
	CtxVdStatus -CsvOutput "C:\Temp\CtxVdStatus.csv" -GridView
	Executes the script with specified CSV output and gives a total overview of all gathered data at the end.
.EXAMPLE
	CtxVdStatus -GridView
	Executes the script and gives a total overview of all gathered data at the end.
.EXAMPLE
	CtxVdStatus -CsvOutput "C:\Temp\CtxVdStatus.csv"
	Executes the script with specified CSV output.
.PARAMETER CsvOutput
	Specify the output path and csv file.
	E.g. -CsvOutput "C:\Temp\CtxVdStatus.csv"
	Default: <current script path>\CtxVdStatus.csv
.PARAMETER GridView
	Specify this parameter to view the data at the end.
	Default: not shown
.PARAMETER numThreads
	Specify the number of threads used.
	E.g. -numThreads 5
	Default: 10 Threads will be used
#>

[cmdletbinding()]
Param(
	[Parameter(Mandatory=$false)][string]$CsvOutput = $null,
	[Parameter(Mandatory=$false)][switch]$GridView = $false,
	[Parameter(Mandatory=$false)][int]$numThreads = 10
)

$Host.UI.RawUI.WindowTitle = "VDI Status"
$newsize = $Host.UI.RawUI.buffersize
$newsize.height = 3000
$newsize.width = 200
$Host.UI.RawUI.buffersize = $newsize
$newsize = $Host.UI.RawUI.WindowSize
$newsize.height = 50
$newsize.width = 200
$Host.UI.RawUI.WindowSize = $newsize
clear-host

IF (-not (Get-Module -Name Citrix.*)) {
	Import-Module Citrix.*
}
IF (-not (Get-PSSnapin -Name Citrix.* -ErrorAction SilentlyContinue)) {
	Add-PSSnapin Citrix.*
}
$ScriptBlock = {
	Param (
		$DCObject
	)
	IF (-not (Get-Module -Name Citrix.*)) {
		Import-Module Citrix.*
	}
	IF (-not (Get-PSSnapin -Name Citrix.* -ErrorAction SilentlyContinue)) {
		Add-PSSnapin Citrix.*
	}
	$success = $true
	$HashInfoObj = @()
	$RegistryPath  = "SYSTEM\CurrentControlSet\Services\bnistack\PvsAgent"
	$computer = $DCObject.HostedMachineName
	If ((test-connection -ComputerName $computer -Count 1 -Quiet) -eq "True") {
		If ($DCObject.ProvisioningType -eq "PVS") {
			$disk = Get-WmiObject Win32_LogicalDisk -ComputerName $computer -Filter "DeviceID='D:'" | Select-Object Size,FreeSpace
			$OutDiskDSize = "{0:N2}" -f ($disk.Size / 1GB)
			$OutDiskDFree = "{0:N2}" -f ($disk.FreeSpace / 1GB)
			if (($disk.FreeSpace -ge 0) -and ($disk.Size -ge 0)) {
				$OutDiskDFreePercent = "{0:N2}" -f ($disk.FreeSpace / ($disk.Size / 100))
			} else {
				$OutDiskDFreePercent = "{0:N2}" -f (100)
			}
			
			$OutvDiskUsed = (([Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer)).OpenSubKey("$RegistryPath")).GetValue("DiskName")
		} else {
			$OutDiskDSize = "N/A"
			$OutDiskDFree = "N/A"
			$OutDiskDFreePercent = "{0:N2}" -f 100
			$OutvDiskUsed = "N/A"
		}
		$lastboottime = (Get-WmiObject -Class Win32_OperatingSystem -computername $computer).LastBootUpTime            
		$PCupt = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($lastboottime)
		try {
			$CurrStateHours = "{0:N2}" -f ((NEW-TIMESPAN -Start (get-date $DCObject.SessionStateChangeTime -format G) -End (get-date -format G)).TotalHours)
		} catch {
			$CurrStateHours = "{0:N2}" -f 0
		} finally {
			if (!($CurrStateHours -ge 0)) {
				$CurrStateHours = "{0:N2}" -f 0
			}
		}
		
		$HashInfoObj = [PSCustomObject]@{
			"Computer" = $computer ;
			"RegistrationState" = $DCObject.RegistrationState ;
			"Uptime" = ([string]$PCupt.days + ":" + [string]$PCupt.hours + ":" + [string]$PCupt.minutes + ":" + [string]$PCupt.seconds) ;
			"DiskFreePercent" = $OutDiskDFreePercent ;
			"DiskSizeGB" = $OutDiskDSize ;
			"DiskFreeGB" = $OutDiskDFree ;
			"vdisk" = $OutvDiskUsed ;
			"User" = $DCObject.SessionUserName ;
			"SessionState" = $DCObject.SessionState ;
			"SessionStateChangeTime" = $DCObject.SessionStateChangeTime ;
			"HoursInCurrentState" = $CurrStateHours
			"ProvisioningType" = $DCObject.ProvisioningType ;
			"InMaintenanceMode" = $DCObject.InMaintenanceMode ;
			"Succesfull" = $success
		}
	} else {
		$HashInfoObj = [PSCustomObject]@{
			"Computer" = $computer ;
			"Succesfull" = $false
		}
	}
	return $HashInfoObj
}

function PickList(){
	[cmdletbinding()]
	param(
		[string[]]$xPickListArray,
		[String]$xPickListPromptText,
		[Boolean]$xPickListRequired
	)
	$counter=0
	$xPickListSelection=$null
	foreach($entry in $xPickListArray){
		$array += (,($counter,$entry))
		$counter++
	}
	Write-Host "`n$xPickListPromptText`n"
	foreach ($arrayentry in $array){
		Write-Host $("`t"+$arrayentry[0]+".`t"+$arrayentry[1])
	}
	if ($xPickListRequired) {
		while (!$Answer){
			Write-Host "`nRequired " -Fore Red -NoNewline
			$xPickListSelection = Read-Host "Enter Option Number"
			$Answer = $null
			try {
				if ([int]$xPickListSelection -is [int]) {
					if (([int]$xPickListSelection -ge 0) -and ([int]$xPickListSelection -lt $counter)) {
						$Answer = 1
					}
				}
			} catch {
				$Answer = $null
			}
		}
		return $array[$xPickListSelection][1]
	} else {
		Write-Host "`nNot Required " -Fore White -NoNewline
		$xPickListSelection = Read-Host "Enter Option Number: "
		if($xPickListSelection){
			return $array[$xPickListSelection][1]
		}
	}
}

function Ask-YesNo {
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$true)][string]$title,
		[Parameter(Mandatory=$true)][string]$message,
		[Parameter(Mandatory=$false)][string]$yesinfo="Yes, please.",
		[Parameter(Mandatory=$false)][string]$noinfo="No. Thank you.",
		[Parameter(Mandatory=$false)][int]$default=0
	)
	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", $yesinfo
	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", $noinfo
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
	return $host.ui.PromptForChoice($title, $message, $options, 1) 
}

$RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $numThreads)
$RunspacePool.Open()
$Jobs = @()
$progress = 0
$numSteps = 4
$numUsers = 0
$numUnregistered = 0
$numUsersDisconnected = 0
Clear-Host
$delgrps = (Get-BrokerDesktopGroup).Name
$MenuChoiceText = "Please select a Delivery Group"
$Selected = PickList $delgrps $MenuChoiceText $true
Clear-Host
$sw = [system.diagnostics.stopwatch]::startNew()
Write-Progress -Activity 'Gathering VDI status info' -CurrentOperation "Collecting desktops..." -PercentComplete ((1 / $numSteps) * 100) -Id 1
$desktops = get-brokerdesktop -DesktopGroupName $Selected -MaxRecordCount 3000
[int]$numUnregistered = (get-brokerdesktop -RegistrationState Unregistered -MaxRecordCount 3000).Count
[int]$numPoweredOn = (get-brokerdesktop -PowerState On -MaxRecordCount 3000).Count
Write-Progress -Activity 'Gathering VDI status info' -CurrentOperation "Collecting desktop info..." -PercentComplete ((2 / $numSteps) * 100) -Id 1
foreach ($desktop in $desktops) {
	$computer = $desktop.HostedMachineName
	$progress++
	if (!($desktop.SessionUserName -eq $null)) {
		$numUsers++
		if ($desktop.SessionState -eq "Disconnected") {
			$numUsersDisconnected++
		}
	} 
	Write-Progress -Activity 'Processing VMs' -CurrentOperation $computer -PercentComplete (($progress / $desktops.count) * 100) -Id 2 -ParentId 1
	$DCObject = get-brokerdesktop -HostedMachineName $computer -PowerState On
	if ( $DCObject -ne $NULL ) {
		$Job = [powershell]::Create().AddScript($ScriptBlock).AddParameter("DCObject", $DCObject)
		$Job.RunspacePool = $RunspacePool
		$Jobs += New-Object PSObject -Property @{
			RunNum = $_
			Job = $Job
			Result = $Job.BeginInvoke()
		}
	}
}
Write-Progress -Activity 'Processing VMs' -Completed -Id 2 -ParentId 1

$progress = 0
$HashInfoCol = @()
$output = @()
$returned = @()
Write-Progress -Activity 'Gathering VDI status info' -CurrentOperation "Processing gathered info..." -PercentComplete ((3 / $numSteps) * 100) -Id 1
ForEach ($Job in $Jobs) {
	$progress++
	Write-Progress -Activity 'Processing Jobs' -PercentComplete (($progress / $Jobs.count) * 100) -Id 2 -ParentId 1
	$returned = $Job.Job.EndInvoke($Job.Result)
	If ($returned.Succesfull) {
		$output = [PSCustomObject]@{
			"Computer" = $returned.Computer ;
			"RegistrationState" = $returned.RegistrationState ;
			"Uptime" = $returned.Uptime ;
			"DiskFreePercent" = $returned.DiskFreePercent ;
			"DiskSizeGB" = $returned.DiskSizeGB ;
			"DiskFreeGB" = $returned.DiskFreeGB ;
			"vdisk" = $returned.vdisk ;
			"User" = $returned.User ;
			"SessionState" = $returned.SessionState ;
			"SessionStateChangeTime" = $returned.SessionStateChangeTime ;
			"HoursInCurrentState" = $returned.HoursInCurrentState ;
			"ProvisioningType" = $returned.ProvisioningType ;
			"InMaintenanceMode" = $returned.InMaintenanceMode
		}
		$HashInfoCol += $output
	}
}

Write-Progress -Activity 'Processing Jobs' -Completed -Id 2 -ParentId 1
Write-Progress -Activity 'Gathering VDI status info' -CurrentOperation "Preparing output..." -PercentComplete ((4 / $numSteps) * 100) -Id 1
Write-Host -Fore Yellow ("`n`tTop Uptime:")
$HashInfoCol | Sort-Object Uptime -Descending| select -First 10 | ft -AutoSize
Write-Host -Fore Yellow ("`n`tTop least free diskspace%:")
$HashInfoCol | Sort-Object DiskFreePercent | select -First 10 | ft -AutoSize
Write-Progress -Activity 'Gathering VDI status info' -Completed -Id 1
$sw.Stop()
Write-Host ("`n`tIt took " + ("{0:N2}" -f $sw.Elapsed.TotalSeconds) + " seconds to complete (" + ("{0:N2}" -f $sw.Elapsed.TotalMinutes) + " minutes)`n`n")
Write-Host -NoNewLine -ForegroundColor Yellow "INFO: "
Write-Host -ForegroundColor White ([string]$numPoweredOn + " VM's are powerend on (total)")
Write-Host -NoNewLine -ForegroundColor Yellow "INFO: "
Write-Host -ForegroundColor White ([string]$numUnregistered + " VM's are Unregistered (total)")
Write-Host -NoNewLine -ForegroundColor Yellow "INFO: "
if ($numUsersDisconnected -eq 1) { $isare = " is disconnected." } else { $isare = " are disconnected." }
if ($numUsers -eq 1) { $hashave = " user has a desktop of which " } else { $hashave = " users have a desktop of which " }
Write-Host -ForegroundColor White ([string]$Selected + ": " + [string]$numUsers + $hashave + [string]$numUsersDisconnected + $isare)
Write-Host ""

If ($CsvOutput) {
	try {
		$HashInfoCol | ConvertTo-Csv -Delimiter ";" -NoTypeInformation | Out-File -FilePath $CsvOutput
		Write-Host -Fore Yellow ("`n`tData written to " + $CsvOutput)
	} catch {
		Write-Host -Fore Red ('`n`tSomething went wrong writing the data to: "' + $CsvOutput + '"')
	}
} else {
	Write-Host -NoNewLine -ForegroundColor Yellow "INFO: "
	Write-Host -ForegroundColor White 'Specify the -CsvOutput "<path\csvfile.csv>" to save the data to a file.'
	$title = "Save CSV"
	$message = "Would you like to save the data to a CSV-File"
	switch (Ask-YesNo -title $title -message $message -default 1){
		0 {#Yes
			try {
				$CsvOutput = "CtxVdStatus.csv"
				$HashInfoCol | ConvertTo-Csv -Delimiter ";" -NoTypeInformation | Out-File -FilePath $CsvOutput
				Write-Host -Fore Yellow ("`n`tData written to " + $CsvOutput)
			} catch {
				Write-Host -Fore Red ('`n`tSomething went wrong writing the data to: "' + $CsvOutput + '"')
			}
		
		}
		1 {#No
		
		}
	}

}
If ($GridView){
	$HashInfoCol | Out-GridView -Title ('Virtual Desktops from the "' + $Selected + '" pool')
} else {
	Write-Host -NoNewLine -ForegroundColor Yellow "INFO: "
	Write-Host -ForegroundColor White 'Specify the -GridView parameter to view data in a seperate window'
	$title = "GridView"
	$message = "Would you like to save the data to view the data in a seperate window"
	switch (Ask-YesNo -title $title -message $message -default 1){
		0 {#Yes
			$HashInfoCol | Out-GridView -Title ('Virtual Desktops from the "' + $Selected + '" pool')
		}
		1 {#No
		
		}
	}


}
Write-Host ""
