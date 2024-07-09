# omoda_itsec.ps1
# This script will be updated weekly

$logFilePath = "C:\Omoda\itsec_log.txt"
$date = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
"Log entry: $date" | Out-File -Append -FilePath $logFilePath
Write-Host "Log entry created at $date"