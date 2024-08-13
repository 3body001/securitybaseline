# omoda_itsec.ps1
# This script will be updated weekly

$logFilePath = "C:\Omoda\itsec_log.txt"
$date = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
"Log entry: $date" | Out-File -Append -FilePath $logFilePath
Write-Host "Log entry created at $date"
# Initialize variables to store account information
$m365Accounts = @()
$oneDriveAccount = $null

# Define the registry path for the OneDrive account
$oneDriveRegistryPath = "HKCU:\Software\Microsoft\OneDrive\Accounts\Business1"

# Check if the OneDrive registry path exists
if (Test-Path $oneDriveRegistryPath) {
    # Retrieve the UserEmail value from the OneDrive registry
    $oneDriveAccount = Get-ItemProperty -Path $oneDriveRegistryPath -Name "UserEmail" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty UserEmail
    
    if ($oneDriveAccount) {
        Write-Host "Found OneDrive account email: $oneDriveAccount"
    } else {
        Write-Host "No OneDrive account email found in the registry."
    }
} else {
    Write-Host "OneDrive account registry key not found."
}

# Define another registry path (e.g., LanguageResources) and retrieve emails as needed
$languageResourcesPath = "HKCU:\Software\Microsoft\Office\16.0\Common\LanguageResources\LocalCache"

# Check if the registry path exists
if (Test-Path $languageResourcesPath) {
    # Get all subkeys under the LocalCache path
    $subKeys = Get-ChildItem -Path $languageResourcesPath

    # Collect all email-like subkeys into a single string
    foreach ($subKey in $subKeys) {
        # Assuming the subkey names are valid email addresses
        $email = $subKey.PSChildName
        # Validate that the subkey resembles an email format
        if ($email -match '^[^@]+@[^@]+\.[^@]+$') {
            Write-Host "Found email: $email"
            # Store each email in the $m365Accounts array
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
#Write-Host "JSON Payload:"
#Write-Host $payload

# Define the webhook URL
$webhookUrl = "https://www.larksuite.com/flow/api/trigger-webhook/47a6143be9f49e03cdf178644bcc65bd"

# Upload the JSON payload to the webhook
$response = Invoke-RestMethod -Uri $webhookUrl -Method Post -ContentType "application/json" -Body $payload
