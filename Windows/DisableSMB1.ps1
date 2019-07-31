
# Get current SMB Server Configuration
Get-SmbServerConfiguration
Read-Host -Prompt "Press Enter key to continue"

# Disable SMB1.0 Server
Set-SmbServerConfiguration -EnableSMB1Protocol $false -Force

# Get current SMB Server Configuration
Get-SmbServerConfiguration
Read-Host -Prompt "Press Enter key to continue"

# Disable SMB1.0 Client
sc.exe config lanmanworkstation depend= bowser/mrxsmb20/nsi
sc.exe config mrxsmb10 start= disabled
