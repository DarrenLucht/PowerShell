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
	$ERC = @(Get-ReceiveConnector -Server $Server | Select Server, Fqdn, Enabled, Name, Identity, TransportRole, Bindings, AuthMechanism, RemoteIPRanges, PermissionGroups, WhenCreated, WhenChanged, OriginatingServer | Format-List)
	$ERC
	Stop-Transcript
}


Else {
	Start-Transcript Get-ExchangeReceiveConnector.txt
	$ERC = @(Get-ReceiveConnector | Select Server, Fqdn, Enabled, Name, Identity, TransportRole, Bindings, AuthMechanism, RemoteIPRanges, PermissionGroups, WhenCreated, WhenChanged, OriginatingServer | Format-List)
	$ERC
	Stop-Transcript
}

