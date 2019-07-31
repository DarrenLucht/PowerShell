<#
.SYNOPSIS
Set-ExchangeOfflineAddressBook.ps1

.DESCRIPTION 
This script modifies the OAB named Default Offline Address Book to allow any 
virtual directory in the organization to accept requests to download the OAB.

.EXAMPLE
./Set-ExchangeOfflineAddressBook.ps1

.NOTES
Written by: Darren Lucht

* Github:  https://github.com/DarrenLucht/PowerShell/tree/master/Exchange

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
V1.00, 03/27/2019 - Initial version
#>


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

Clear-Host

Start-Transcript Set-ExchangeOfflineAddressBook.txt
Get-OfflineAddressBook | Format-List Name, AddressLists, GeneratingMailbox, IsDefault, VirtualDirectories, GlobalWebDistributionEnabled
Get-OfflineAddressBook | Where-Object {$_.ExchangeVersion.ExchangeBuild.Major -Eq 15} | Set-OfflineAddressBook -Identity "Default Offline Address List (Ex2013)" -GlobalWebDistributionEnabled $true -VirtualDirectories $null 
Get-OfflineAddressBook | Format-List Name, AddressLists, GeneratingMailbox, IsDefault, VirtualDirectories, GlobalWebDistributionEnabled
Stop-Transcript
