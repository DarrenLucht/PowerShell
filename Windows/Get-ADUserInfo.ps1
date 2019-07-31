
$MyDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TimestampCsv = Get-Date -UFormat %Y-%m-%d
$TimestampHtml = Get-Date -UFormat %m/%d/%Y
$CsvFile = "$MyDir\Get-ADUserInfo_$TimestampCsv.csv"
$HtmlFile = "$MyDir\Get-ADUserInfo.html"


Get-ADUser -Filter {Name -Notlike 'Discovery*' -AND Name -Notlike 'Health*' -AND Name -NotLike 'System*'} `
-properties Name, EmailAddress, Enabled, Office, ScriptPath, HomeDrive, HomeDirectory, LastLogonDate | 
Sort-Object Name | 
Select-Object Name, EmailAddress, Enabled, Office, ScriptPath, HomeDrive, HomeDirectory, LastLogonDate | 
Export-Csv -path $CsvFile -NoTypeInformation
