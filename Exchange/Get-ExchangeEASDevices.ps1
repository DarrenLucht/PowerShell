<#
.SYNOPSIS
Get-ExchangeEASDevices.ps1 - Exchange Server ActiveSync device report

.DESCRIPTION 
Produces a report of ActiveSync device associations in the organization.

.OUTPUTS
Results are output to screen and CSV file, as well as optional HTML report, and email

.PARAMETER Html
Produces a HTML report containing stats for all ActiveSync devices.

.PARAMETER SendEmail
Sends the HTML report via email using the SMTP configuration within the script.

.PARAMETER Age
Limits the report to devices that have not attempted synced in more than XX days.

.EXAMPLE
./Get-ExchangeEASDevices.ps1
Produces a CSV file containing stats for all ActiveSync devices.

.EXAMPLE
./Get-ExchangeEASDevices.ps1 -Html
Produces a HTML report containing stats for all ActiveSync devices.

.EXAMPLE
./Get-ExchangeEASDevices.ps1 -SendEmail -MailFrom:exchange@example.com -MailTo:admin@example.com -MailServer:smtp.example.com
Sends an email HTML report with CSV file attached for all ActiveSync devices.

.EXAMPLE
./Get-ExchangeEASDevices.ps1 -Age 30
Limits the report to devices that have not attempted synced in more than 30 days.

.NOTES
Modified by: Darren Lucht
Original script obtained from exchangeserverpro.net

* Github: https://github.com/DarrenLucht/PowerShell/tree/master/Exchange

Change Log:
V1.01, 04/01/2019 - Added Html parameter
V1.00, 03/01/2019 - Initial version

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
#>

[CmdletBinding()]
param (
	[Parameter( Mandatory=$false)]
	[switch]$Html,
	[Parameter( Mandatory=$false)]
	[switch]$SendEmail,
    [Parameter( Mandatory=$false)]
	[string]$MailFrom,
	[Parameter( Mandatory=$false)]
	[string]$MailTo,
	[Parameter( Mandatory=$false)]
	[string]$MailServer,
    [Parameter( Mandatory=$false)]
	[int]$Age = 0
	)

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportfile = "$myDir\Get-ExchangeEASDevices.csv"
$htmlfile = "$myDir\Get-ExchangeEASDevices.html"

# Variables
$now = Get-Date
$date = $now.ToShortDateString()
$report = @()
$stats = @("DeviceID",
           "DeviceAccessState",
           "DeviceModel"
           "DeviceType",
           "DeviceFriendlyName",
           "DeviceOS",
           "LastSuccessSync"
)

# Email Settings
$reportemailsubject = "Exchange ActiveSync Devices - $date"
$smtpsettings = @{
	To =  $MailTo
	From = $MailFrom
    Subject = $reportemailsubject
	SmtpServer = $MailServer
}
	
#Add Exchange snapin if not already loaded in the PowerShell session
if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
{
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}
else
{
    Write-Warning "Exchange Server management tools are not installed on this computer."
    EXIT
}

Write-Host "Fetching list of Mailboxes with EAS Device Partnerships"
$MailboxesWithEASDevices = @(Get-CASMailbox -Resultsize Unlimited | Where-Object {$_.HasActiveSyncDevicePartnership})
Write-Host "$($MailboxesWithEASDevices.count) Mailboxes with EAS device partnerships"

Foreach ($Mailbox in $MailboxesWithEASDevices)
{
    $EASDeviceStats = @(Get-ActiveSyncDeviceStatistics -Mailbox $Mailbox.Identity -WarningAction SilentlyContinue)    
    Write-Host "$($Mailbox.Identity) has $($EASDeviceStats.Count) device(s)"
    $MailboxInfo = Get-Mailbox $Mailbox.Identity | Select-Object DisplayName,PrimarySMTPAddress,Office    
    
    Foreach ($EASDevice in $EASDeviceStats)
    {
        Write-Host -ForegroundColor Green "Processing $($EASDevice.DeviceID)"       
        $lastsyncattempt = ($EASDevice.LastSyncAttemptTime)
        if ($null -eq $lastsyncattempt)
        {
            $syncAge = "Never"
        }
        else
        {
            $syncAge = ($now - $lastsyncattempt).Days
        }
		
        #Add to report if last sync attempt greater than Age specified
        if ($syncAge -ge $Age -or $syncAge -eq "Never")
        {
            $reportObj = New-Object PSObject
            $reportObj | Add-Member NoteProperty -Name "Display Name" -Value $MailboxInfo.DisplayName
            $reportObj | Add-Member NoteProperty -Name "Office" -Value $MailboxInfo.Office
            $reportObj | Add-Member NoteProperty -Name "Email Address" -Value $MailboxInfo.PrimarySMTPAddress
            $reportObj | Add-Member NoteProperty -Name "Sync Age" -Value $syncAge
                
            Foreach ($stat in $stats)
            {
                $reportObj | Add-Member NoteProperty -Name $stat -Value $EASDevice.$stat
            }
            $report += $reportObj
        }
    }
}
Write-Host -ForegroundColor White "Saving Report to $reportfile"
$report | Export-Csv -NoTypeInformation $reportfile -Encoding UTF8

if ($Html)
{
    $reporthtml = $report | ConvertTo-Html -Fragment
	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 22px; font-weight: bold; font-family: Arial;}
				H2{font-size: 18px; font-weight: bold; font-family: Arial;}
				H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt; white-space: nowrap;}
				TR:nth-child(even){background: #d3ffd3;}
				TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid #969595; padding: 5px;}
				td.pass{background: #B7EB83;}
				td.warn{background: #FFF275;}
				td.fail{background: #FF2626; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
                <H2>Exchange ActiveSync Devices Associations with Greater Than $age Days Since Last Sync ($date)</H2>"		
	$htmltail = "</body></html>"
	Write-Host -ForegroundColor White "Saving Report to $htmlfile"
	$htmlreport = $htmlhead + $reporthtml + $htmltail | Out-File -FilePath $htmlfile -Encoding utf8
	Invoke-Item $htmlfile
}

if ($SendEmail)
{
    $reporthtml = $report | ConvertTo-Html -Fragment
	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid #969595; padding: 5px; }
				td.pass{background: #B7EB83;}
				td.warn{background: #FFF275;}
				td.fail{background: #FF2626; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
                <p>Report of Exchange ActiveSync device associations with greater than $age days since last sync attempt as of $date. CSV version of report attached to this email.</p>"		
	$htmltail = "</body></html>"	
	$htmlreport = $htmlhead + $reporthtml + $htmltail
	Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -Attachments $reportfile
}
