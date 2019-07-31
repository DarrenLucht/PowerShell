<#
.SYNOPSIS
AvastRMM.ps1

.DESCRIPTION 
This will check and alert for Avast Antivirus service not running and
Out of Date Signatures.

This is for Datto RMM integration to show server device Antivirus Product
and Antivirus Status useful for reporting.

A site monitoring policy must be set within RMM and enabled.

.NOTES
Written by: Darren Lucht

License:

The MIT License (MIT)

Copyright (c) 2019 Darren Lucht

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Change Log:
V1.00, 10/31/2018 - Initial version
#>


param ([Int32]$uptodate=7)

if ([Environment]::GetEnvironmentVariable("uptodate", "Process"))
{
    $uptodate = [Environment]::GetEnvironmentVariable("uptodate", "Process")
}

function Write-Alert {

    param([string]$Alert)
	Write-Host "<-Start Result->"
    Write-Host "CSMon_Result="$Alert
    Write-Host "<-End Result->"
    exit 1
}

# Avast installation path
$Bits_OS = (Get-WmiObject Win32_OperatingSystem).OSArchitecture.Substring(0,2)
if ($Bits_OS -eq 32) {
    $AvastInstallPath = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\AVAST Software\Avast' -Name ProgramFolder).ProgramFolder
} 
elseif ($Bits_OS -eq 64) {
    $AvastInstallPath = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\AVAST Software\Avast' -Name ProgramFolder).ProgramFolder
}

# Compare Avast definition file to current date
$AvastDef = $AvastInstallPath + '\Defs\aswdefs.ini'
$AvastUpdate = Get-Item $AvastDef | select -Property LastWriteTime

$today = Get-Date -format yyyy-M-dd
If ($today -lt $AvastUpdate) {
    Write-Alert "Avast Antivirus Definitions Not Up To Date"
    $DefUpToDate = $false
}else{
    $DefUpToDate = $true
}

# Check to see if the Avast service is running
$ServiceName = "avast! Antivirus"
if (Get-Service $ServiceName -ErrorAction SilentlyContinue)
	{
		if ((Get-Service -Name $ServiceName).Status -eq 'Running')
			{
				$ServiceStatus = (Get-Service -Name $ServiceName).Status
                $ServiceRunning = $true
			}
		elseif ((Get-Service -Name $ServiceName).Status -eq 'Stopped')
			{
				$ServiceStatus = (Get-Service -Name $ServiceName).Status
                $ServiceRunning = $false
				Write-Alert $ServiceName "-" $ServiceStatus
			}	
		else
			{
				$ServiceStatus = (Get-Service -Name $ServiceName).Status
				Write-Alert $ServiceName "-" $ServiceStatus
			}
	}
else
	{
        $ServiceRunning = $false
		Write-Alert "$ServiceName not found"
	}

# Write Avast status to JSON file
    @{product="Avast Antivirus";running=$ServiceRunning;upToDate=$DefUpToDate} | ConvertTo-Json -Compress -depth 100 | Out-File "$env:ProgramData\CentraStage\AEMAgent\antivirus.json"
