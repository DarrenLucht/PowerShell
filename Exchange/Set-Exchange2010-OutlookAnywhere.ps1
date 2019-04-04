<#
.SYNOPSIS
Set-Exchange2010-OutlookAnywhere.ps1

.DESCRIPTION 
Configure Outlook Anywhere on your Exchange 2010 servers for coexistence with Exchange 2016 servers.

Execute On Exchange 2010 server

.EXAMPLE
./Set-Exchange2010-OutlookAnywhere.ps1

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
V1.00, 03/28/2019 - Initial version
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

Start-Transcript Set-Exchange2010-OutlookAnywhere.txt

# External host name of your Exchange 2016 Server 
$Exchange2016 = "mail.example.com"

# Configure Exchange 2010 servers to accept connections from Exchange 2016 servers.
Get-ExchangeServer | Where-Object {($_.AdminDisplayVersion -Like "Version 14*") -And ($_.ServerRole -Like "*ClientAccess*")} | Get-ClientAccessServer | Where-Object {$_.OutlookAnywhereEnabled -Eq $True} | ForEach-Object {Set-OutlookAnywhere "$_\RPC (Default Web Site)" -ClientAuthenticationMethod Basic -SSLOffloading $False -ExternalHostName $Exchange2016 -IISAuthenticationMethods NTLM, Basic}

# Enable Outlook Anywhere and configure Exchange 2010 to accept connections from Exchange 2016 servers.
Get-ExchangeServer | Where-Object {($_.AdminDisplayVersion -Like "Version 14*") -And ($_.ServerRole -Like "*ClientAccess*")} | Get-ClientAccessServer | Where-Object {$_.OutlookAnywhereEnabled -Eq $False} | Enable-OutlookAnywhere -ClientAuthenticationMethod Basic -SSLOffloading $False -ExternalHostName $Exchange2016 -IISAuthenticationMethods NTLM, Basic

# Verify that Outlook Anywhere on your Exchange 2010 servers accept connections redirected from Exchange 2016
Get-ExchangeServer | Where-Object {($_.AdminDisplayVersion -Like "Version 14*") -And ($_.ServerRole -Like "*ClientAccess*")} | Get-OutlookAnywhere | Format-Table Server, ClientAuthenticationMethod, IISAuthenticationMethods, SSLOffloading, ExternalHostname -Auto

Stop-Transcript
