<#
.SYNOPSIS
Get-ExchangeReceiveConnector.ps1

.DESCRIPTION 
This cmdlet is available only in on-premises Exchange.

Use the Get-ExchangeReceiveConnector to view Receive connectors on Mailbox servers. 
Receive connectors listen for inbound SMTP connections on the Exchange server.

.PARAMETER Server
The Server parameter filters the results by the specified Exchange server.

.EXAMPLE
./Get-ExchangeReceiveConnector.ps1 -Server Exchange01
This example returns detailed information for the Receive connectors on the server named Exchange01.

.EXAMPLE
./Get-ExchangeReceiveConnector.ps1
This example returns detailed information for the Receive connectors on all Exchange servers.

.NOTES
Written by: Darren Lucht

* Github:	https://github.com/DarrenLucht/PowerShell/tree/master/Exchange

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

Param(	
    [Parameter( Mandatory=$false)]
	[string]$Server
)

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


If ($Server) {
	Start-Transcript Get-ExchangeReceiveConnector-$Server.txt
	$ERC = @(Get-ReceiveConnector -Server $Server | Select-Object Server, Fqdn, Enabled, Name, Identity, TransportRole, Bindings, AuthMechanism, RemoteIPRanges, PermissionGroups, WhenCreated, WhenChanged, OriginatingServer | Format-List)
	$ERC
	Stop-Transcript
}


Else {
	Start-Transcript Get-ExchangeReceiveConnector.txt
	$ERC = @(Get-ReceiveConnector | Select-Object Server, Fqdn, Enabled, Name, Identity, TransportRole, Bindings, AuthMechanism, RemoteIPRanges, PermissionGroups, WhenCreated, WhenChanged, OriginatingServer | Format-List)
	$ERC
	Stop-Transcript
}
