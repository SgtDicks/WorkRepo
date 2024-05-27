# Path to the log file
$logFile = "C:\Path\To\Your\LogFile.txt"

# Webhook URL
$webhookUrl = "https://your-webhook-url.com"

# Function to send webhook notification
function Send-Webhook {
    param (
        [string]$message
    )
    
    $body = @{
        text = $message
    } | ConvertTo-Json
    
    try {
        Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $body -ContentType "application/json"
    } catch {
        Write-Output "Failed to send webhook: $_"
    }
}

# Function to log device change
function Log-DeviceChange {
    param (
        [string]$eventType,
        [string]$deviceName
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $eventType: $deviceName"
    
    Add-Content -Path $logFile -Value $logEntry
    
    Send-Webhook -message $logEntry
}

# Register event for device change
Register-WmiEvent -Class Win32_DeviceChangeEvent -SourceIdentifier USBChangeEvent -Action {
    $eventType = "Device Change"
    $deviceName = "USB Device"
    
    Log-DeviceChange -eventType $eventType -deviceName $deviceName
}

# Keep the script running to listen for events
while ($true) {
    Start-Sleep -Seconds 10
}
