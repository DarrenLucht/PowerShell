<# 
Clear Exchange 2013/2016/2019 Log & ETL Files
Original Script: https://gallery.technet.microsoft.com/office/Clear-Exchange-2013-Log-71abba44
NOTE: This will not Clean the Mail Logs

Add a scheduled task to run on the Exchange server
Action: Start a program
Settings
  > Program/script: Powershell.exe
  > Add arguments (optional): Specify path to powershell script

Example to run manually: ./Exchange-CleanLogs.ps1
#>

Set-Executionpolicy -Scope CurrentUser RemoteSigned

#Number of days to retain logs
$Days=7

# Modify paths if needed
$IISLogPath = "C:\inetpub\logs\LogFiles\"
$ExchangeLoggingPath = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\"
$ETLLoggingPath = "C:\Program Files\Microsoft\Exchange Server\V15\Bin\Search\Ceres\Diagnostics\ETLTraces\"
$ETLLoggingPath2 = "C:\Program Files\Microsoft\Exchange Server\V15\Bin\Search\Ceres\Diagnostics\Logs\"

Function CleanLogfiles($TargetFolder)
{
  Write-Host -debug -ForegroundColor Yellow -BackgroundColor DarkBlue $TargetFolder

    If (Test-Path $TargetFolder) {
        $Now = Get-Date
        $LastWrite = $Now.AddDays(-$Days)
        $Files = Get-ChildItem $TargetFolder -Recurse | Where-Object {$_.Name -like "*.log" -or $_.Name -like "*.blg" -or $_.Name -like "*.etl"}  | Where {$_.lastWriteTime -le "$LastWrite"} | Select-Object FullName  
        ForEach ($File in $Files)
            {
				$FullFileName = $File.FullName  
				Write-Host "Deleting File: $FullFileName" -ForegroundColor "Yellow"; 
				Remove-Item $FullFileName -ErrorAction SilentlyContinue | Out-Null
            }
    }
Else {
		Write-Host "The folder $TargetFolder doesn't exist! Check the folder path!" -ForegroundColor "Red"
    }
}

# Execute function on target folders
CleanLogfiles($IISLogPath)
CleanLogfiles($ExchangeLoggingPath)
CleanLogfiles($ETLLoggingPath)
CleanLogfiles($ETLLoggingPath2)
