# omoda_itsec.ps1
# This script will be updated weekly

#check 7zip version
function Get-7ZipVersion {
    $7zipPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    $keys = Get-ChildItem -Path $7zipPath
    foreach ($key in $keys) {
        $displayName = (Get-ItemProperty -Path $key.PSPath).DisplayName
        if ($displayName -like "7-Zip*") {
            return (Get-ItemProperty -Path $key.PSPath).DisplayVersion
        }
    }
    return $null
}

# Check installed version
$installedVersion = Get-7ZipVersion
$latestVersion = [Version]"24.08"

if ($installedVersion) {
    $installedVersionObj = [Version]$installedVersion
    
    if ($installedVersionObj -lt $latestVersion) {
        Write-Output "7-Zip version is older than $latestVersion. Updating now..."
        # Define URL and installer path
        $url = "https://www.7-zip.org/a/7z2408-x64.exe"
        $installerPath = "$env:TEMP\7z2408-x64.exe"

        # Download the installer
        Invoke-WebRequest -Uri $url -OutFile $installerPath

        # Run the installer silently
        Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait

        Write-Output "7-Zip updated successfully."
    } else {
        #Write-Output "7-Zip is already up to date."
    }
} else {
    #Write-Output "7-Zip is not installed."
}


$logFilePath = "C:\Omoda\itsec_log.txt"
$date = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
"Log entry: $date" | Out-File -Append -FilePath $logFilePath
Write-Host "Log entry created at $date"
# Define the webhook URL for posting the data
$webhookUrl = "https://www.larksuite.com/flow/api/trigger-webhook/47a6143be9f49e03cdf178644bcc65bd"

# List all user profiles in C:\Users\
$userProfiles = Get-ChildItem -Path "C:\Users\" | Where-Object { $_.PSIsContainer -eq $true }

foreach ($profile in $userProfiles) {
    $regHivePath = "$($profile.FullName)\NTUSER.DAT"
    $tempRegPath = "HKU\TempUser_$($profile.Name)"

    if (Test-Path $regHivePath) {
        try {
            # Load the registry hive for this user profile
            reg load $tempRegPath $regHivePath
            Write-Host "Loaded user hive for $($profile.Name) from $regHivePath"

            # Initialize variables to store account information
            $m365Accounts = @()
            $oneDriveAccount = $null

            # Define registry paths for the loaded hive
            $oneDriveRegistryPath = "$tempRegPath\Software\Microsoft\OneDrive\Accounts\Business1"
            $languageResourcesPath = "$tempRegPath\Software\Microsoft\Office\16.0\Common\LanguageResources\LocalCache"

            # Check the OneDrive account email
            if (Test-Path $oneDriveRegistryPath) {
                $oneDriveAccount = Get-ItemProperty -Path $oneDriveRegistryPath -Name "UserEmail" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty UserEmail
                
                if ($oneDriveAccount) {
                    Write-Host "Found OneDrive account email for $($profile.Name): $oneDriveAccount"
                } else {
                    Write-Host "No OneDrive account email found in the registry for $($profile.Name)."
                }
            } else {
                Write-Host "OneDrive account registry key not found for $($profile.Name)."
            }

            # Check the language resources for M365 accounts
            if (Test-Path $languageResourcesPath) {
                $subKeys = Get-ChildItem -Path $languageResourcesPath

                foreach ($subKey in $subKeys) {
                    $email = $subKey.PSChildName
                    if ($email -match '^[^@]+@[^@]+\.[^@]+$') {
                        Write-Host "Found email for $($profile.Name): $email"
                        $m365Accounts += $email
                    }
                }
            }

            # Combine all found M365 accounts into a single string separated by semicolons
            $m365AccountString = $m365Accounts -join ";"

            # Get the computer's serial number
            $serialNumber = (Get-WmiObject -Class Win32_BIOS).SerialNumber

            # Prepare the JSON payload
            $payload = @{
                "m365" = @(
                    @{
                        "SN"              = $serialNumber
                        "m365account"     = $m365AccountString
                        "onedriveaccount" = $oneDriveAccount
                    }
                )
            } | ConvertTo-Json

            # Display the JSON payload for verification
            Write-Host "JSON Payload for $($profile.Name):"
            Write-Host $payload

            # Upload the JSON payload to the webhook
            $response = Invoke-RestMethod -Uri $webhookUrl -Method Post -ContentType "application/json" -Body $payload
            Write-Host "Successfully uploaded the data to the webhook for $($profile.Name)."

        } finally {
            # Unload the registry hive
            reg unload $tempRegPath
            Write-Host "Unloaded user hive for $($profile.Name)."
        }
    } else {
        Write-Host "NTUSER.DAT not found for profile: $($profile.Name)"
    }
}

Write-Host "All profiles have been processed."
