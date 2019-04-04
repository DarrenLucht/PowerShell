<#
.SYNOPSIS
.\Get-ExchangeMailboxSize.ps1
1. Display Mailbox Sizes with Item Count, Database, Server Details to Screen
2. Export Mailbox Sizes with Item Count, Database, Server Details to Csv file
3. Output Mailbox Sizes with Item Count, Database, Server Details to Html file

.DESCRIPTION 
Use Get-ExchangeMailboxSize to view Exchange Server Mailbox Sizes and Statistics.

.EXAMPLE
./Get-ExchangeMailboxSize.ps1

.NOTES
Written by: Darren Lucht

* Github:	https://github.com/DarrenLucht/PowerShell/tree/master/Exchange

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

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
V1.00, 04/01/2019 - Initial version
#>

#Add Exchange Snap-In if not already loaded in the PowerShell session
If (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
{
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}
Else
{
    Write-Warning "Exchange Server management tools are not installed on this computer."
    Exit
}

Set-ADServerSettings -ViewEntireForest $True
$MyDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TimestampCsv = Get-Date -UFormat %Y-%m-%d
$TimestampHtml = Get-Date -UFormat %m/%d/%Y
$CsvFile = "$MyDir\Get-ExchangeMailboxSize_$TimestampCsv.csv"
$HtmlFile = "$MyDir\Get-ExchangeMailboxSize.html"

Clear-Host
Write-Host ""
Write-host "
  -----------------------------------
  |      Mailbox Size Report        |
  -----------------------------------
   
       1. Display in Console
       2. Export to CSV File
       3. Output to HTML File
       4. Exit"-ForeGround "Green"

Write-Host ""
Write-Host ""
$Number = Read-Host "       Choose The Task"
Write-Host ""
Write-Host ""
Clear-Host

$Output = @()
Switch ($Number) 
{
    # Display to Screen
    1 {
        $AllMailbox = Get-mailbox -resultsize unlimited -IgnoreDefaultScope -Filter {Name -notlike 'Discovery*'}
        ForEach($Mbx In $AllMailbox)
        {
            $Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue
            $userObj = New-Object PSObject
            $userObj | Add-Member NoteProperty -Name "Name" -Value $Mbx.displayname
            $userObj | Add-Member NoteProperty -Name "Email" -Value $Mbx.PrimarySmtpAddress
            $userObj | Add-Member NoteProperty -Name "Mbx-Size" -Value ($Stats.TotalItemSize.Value.ToMB() + $Stats.TotalDeletedItemSize.Value.ToMB())
            $userObj | Add-Member NoteProperty -Name "Mbx-Items" -Value $Stats.ItemCount
			$userObj | Add-Member NoteProperty -Name "Database" -Value $Mbx.Database
            $userObj | Add-Member NoteProperty -Name "Server" -Value $Mbx.ServerName			
            Write-Output $UserObj
        }
        ;Break
    }

    # Export to Csv file
    2 {
        $i = 0
        $AllMailbox = Get-mailbox -resultsize unlimited -Filter {Name -notlike 'Discovery*'}
        ForEach($Mbx In $AllMailbox)
        {
            $Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue
			$User = Get-User $Mbx
			$Mobile = Get-CASMailbox -Identity $Mbx.Name						
            $userObj = New-Object PSObject
            $userObj | Add-Member NoteProperty -Name "Name" -Value $Mbx.DisplayName			
            $userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
			$userObj | Add-Member NoteProperty -Name "Account" -Value $Mbx.SamAccountName
			#$userObj | Add-Member NoteProperty -Name "Title" -Value $User.Title
			$userObj | Add-Member NoteProperty -Name "Office" -Value $User.Office
			$userObj | Add-Member NoteProperty -Name "Email" -Value $Mbx.PrimarySmtpAddress          
            #$userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.EmailAddresses.smtpaddress -join ";")
			If ($Mobile)
			{
				$userObj | Add-Member NoteProperty -Name "Mobile" -Value $Mobile.HasActiveSyncDevicePartnership
			}
            If($Stats)
            {
                $userObj | Add-Member NoteProperty -Name "Mbx-Size" -Value ($Stats.TotalItemSize.Value.ToMB() + $Stats.TotalDeletedItemSize.Value.ToMB())
                $userObj | Add-Member NoteProperty -Name "Mbx-Items" -Value $Stats.ItemCount
                $userObj | Add-Member NoteProperty -Name "Del-Size" -Value $Stats.TotalDeletedItemSize.Value.ToMB()
				$userObj | Add-Member NoteProperty -Name "Del-Items" -Value $Stats.DeletedItemCount
            }
			$userObj | Add-Member NoteProperty -Name "Hidden" -Value $Mbx.HiddenFromAddressListsEnabled			
            $userObj | Add-Member NoteProperty -Name "OU" -Value $Mbx.OrganizationalUnit 			
            #$userObj | Add-Member NoteProperty -Name "ProhibitQuota" -Value $ProhibitSendReceiveQuota
            #$userObj | Add-Member NoteProperty -Name "QuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
            $userObj | Add-Member NoteProperty -Name "LastLogon" -Value $Stats.LastLogonTime
			$userObj | Add-Member NoteProperty -Name "Type" -Value $Mbx.RecipientTypeDetails
			$userObj | Add-Member NoteProperty -Name "Database" -Value $Mbx.Database
            $userObj | Add-Member NoteProperty -Name "Server" -Value $Mbx.ServerName			
            $Output += $UserObj
            
            # Update Counters and Write Progress
            $i++
            Write-Progress -Activity "Scanning Mailboxes . . ." -Status "Scanned: $i of $($AllMailbox.Count)" -PercentComplete ($i/$AllMailbox.Count*100)
        }
		
		Clear-Host
		Write-Host ""
		Write-host "                                                                              " -BackgroundColor White -ForegroundColor DarkBlue
		Write-host "                                                                              " -BackgroundColor White -ForegroundColor DarkBlue
		Write-Host "   Exported Csv Report to $CsvFile   " -BackgroundColor White -ForegroundColor DarkBlue
		Write-host "                                                                              " -BackgroundColor White -ForegroundColor DarkBlue
		Write-host "                                                                              " -BackgroundColor White -ForegroundColor DarkBlue
		Write-Host ""
		
        $Output | Sort-Object Name | Export-Csv -Path $CsvFile -NoTypeInformation
        ;Break
    }

    # Output to Html file
	3 {
        $i = 0
        $AllMailbox = Get-mailbox -resultsize unlimited -Filter {Name -notlike 'Discovery*'}
        ForEach($Mbx In $AllMailbox)
        {
            $Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue
			$User = Get-User $Mbx
			$Mobile = Get-CASMailbox -Identity $Mbx.Name						
            $userObj = New-Object PSObject
            $userObj | Add-Member NoteProperty -Name "Name" -Value $Mbx.DisplayName			
            $userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
			$userObj | Add-Member NoteProperty -Name "Account" -Value $Mbx.SamAccountName
			#$userObj | Add-Member NoteProperty -Name "Title" -Value $User.Title
			$userObj | Add-Member NoteProperty -Name "Office" -Value $User.Office
			$userObj | Add-Member NoteProperty -Name "Email" -Value $Mbx.PrimarySmtpAddress          
            #$userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.EmailAddresses.smtpaddress -join ";")
			If ($Mobile)
			{
				$userObj | Add-Member NoteProperty -Name "Mobile" -Value $Mobile.HasActiveSyncDevicePartnership
			}
            If($Stats)
            {
                $userObj | Add-Member NoteProperty -Name "Mbx-Size" -Value ($Stats.TotalItemSize.Value.ToMB() + $Stats.TotalDeletedItemSize.Value.ToMB())
                $userObj | Add-Member NoteProperty -Name "Mbx-Items" -Value $Stats.ItemCount
                $userObj | Add-Member NoteProperty -Name "Del-Size" -Value $Stats.TotalDeletedItemSize.Value.ToMB()
				$userObj | Add-Member NoteProperty -Name "Del-Items" -Value $Stats.DeletedItemCount
            }
			$userObj | Add-Member NoteProperty -Name "Hidden" -Value $Mbx.HiddenFromAddressListsEnabled			
            $userObj | Add-Member NoteProperty -Name "OU" -Value $Mbx.OrganizationalUnit 			
            #$userObj | Add-Member NoteProperty -Name "ProhibitQuota" -Value $ProhibitSendReceiveQuota
            #$userObj | Add-Member NoteProperty -Name "QuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
            $userObj | Add-Member NoteProperty -Name "LastLogon" -Value $Stats.LastLogonTime
			$userObj | Add-Member NoteProperty -Name "Type" -Value $Mbx.RecipientTypeDetails
			$userObj | Add-Member NoteProperty -Name "Database" -Value $Mbx.Database
            $userObj | Add-Member NoteProperty -Name "Server" -Value $Mbx.ServerName			
            $Output += $UserObj
            
            # Update Counters and Write Progress
            $i++
            Write-Progress -Activity "Scanning Mailboxes . . ." -Status "Scanned: $i of $($AllMailbox.Count)" -PercentComplete ($i/$AllMailbox.Count*100)
        }

        $reporthtml = ($Output | Sort-Object Name | ConvertTo-Html -Fragment)
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
                    <H2>Exchange Mailbox Size Report ($TimestampHtml)</H2>"		
        $htmltail = "</body></html>"
		
		Clear-Host
		Write-Host ""
		Write-host "                                                                     " -BackgroundColor White -ForegroundColor DarkBlue
		Write-host "                                                                     " -BackgroundColor White -ForegroundColor DarkBlue
		Write-Host "   Exported Html Report to $HtmlFile   " -BackgroundColor White -ForegroundColor DarkBlue
		Write-host "                                                                     " -BackgroundColor White -ForegroundColor DarkBlue
		Write-host "                                                                     " -BackgroundColor White -ForegroundColor DarkBlue
		Write-Host ""

        $Output = $htmlhead + $reporthtml + $htmltail | Out-File -FilePath $HtmlFile -Encoding utf8
        Invoke-Item $HtmlFile
        ;Break
    }
    
    # Exit
	4 {
		;Break
	}
}
