# ================================
# Script: Check-ExchangeQueueAndSendToNewRelic.ps1
# Purpose: Check Exchange mail queues, log results, and send to New Relic
# Version: 1.2
# ================================

# --- Configuration ---
$threshold = 1
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogFilePath = Join-Path $ScriptDir "ExchangeQueueData.json"
$NewRelicInsertKey = "c3692725ba5d04e07bb02423175871f9FFFFNRAL"   # Replace with your key
$NewRelicAccountId = "4438265"                                     # Replace with your New Relic account ID
$EventType = "ExchangeQueueData"

# --- Function to send data to New Relic ---
function Send-ToNewRelic {
    param(
        [Parameter(Mandatory = $true)][object]$Payload
    )

    $uri = "https://insights-collector.newrelic.com/v1/accounts/$NewRelicAccountId/events"

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers @{
            "Content-Type"  = "application/json"
            "X-Insert-Key"  = $NewRelicInsertKey
        } -Body ($Payload | ConvertTo-Json -Depth 6)

        Write-Host "Data successfully sent to New Relic."
    }
    catch {
        Write-Warning "Failed to send data to New Relic: $_"
    }
}

# --- Run Exchange Queue Command ---
try {
    $QueueData = Get-Queue | Where-Object { $_.MessageCount -ge $threshold -and $_.DeliveryType -ne "ShadowRedundancy" } |
        Select-Object Identity, DeliveryType, Status, MessageCount, NextHopDomain
}
catch {
    $ErrorMessage = "Exchange shell command failed: $($_.Exception.Message)"
    Write-Warning $ErrorMessage
    $QueueData = @()
}

# --- Determine Alert Info ---
if (-not $QueueData -or $QueueData.Count -eq 0) {
    $AlertStatus = "No messages found or command returned no value."
}
elseif ($QueueData | Where-Object { $_.Status -match "Retry" }) {
    $AlertStatus = "Warning: One or more queues in retry state."
}
else {
    $AlertStatus = "OK"
}

# --- Additional Derived Metrics ---
$totalMessageCount = ($QueueData | Measure-Object -Property MessageCount -Sum).Sum
$retryQueueCount   = ($QueueData | Where-Object { $_.Status -match "Retry" } | Measure-Object).Count
$largestQueue      = $QueueData | Sort-Object -Property MessageCount -Descending | Select-Object -First 1

# --- Host & System Info ---
try {
    $OSInfo = Get-CimInstance Win32_OperatingSystem
    $OSVersion = $OSInfo.Caption
    $Uptime = ((Get-Date) - $OSInfo.LastBootUpTime).TotalHours
}
catch {
    $OSVersion = "Unknown"
    $Uptime = 0
}

try {
    $ExchangeServer = Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue
    $ExchangeVersion = if ($ExchangeServer) { $ExchangeServer.AdminDisplayVersion } else { "Unknown" }
}
catch {
    $ExchangeVersion = "Unknown"
}

# --- Build enriched payload ---
if (-not $QueueData) { $QueueData = @() }  # ensure it's an array even if empty

$Payload = [PSCustomObject]@{
    eventType         = $EventType
    timestamp         = (Get-Date).ToUniversalTime().ToString("o")
    host              = $env:COMPUTERNAME
    alertStatus       = $AlertStatus
    queueCount        = ($QueueData | Measure-Object).Count
    totalMessageCount = $totalMessageCount
    retryQueueCount   = $retryQueueCount
    largestQueue      = if ($largestQueue) {
                            [PSCustomObject]@{
                                Identity     = $largestQueue.Identity
                                MessageCount = $largestQueue.MessageCount
                                Status       = $largestQueue.Status
                                NextHop      = $largestQueue.NextHopDomain
                            }
                        } else { $null }
    OSVersion         = $OSVersion
    uptimeHours       = [math]::Round($Uptime, 2)
    exchangeVersion   = "$ExchangeVersion"
    scriptVersion     = "1.2"
    queues            = @($QueueData)  # keep detailed queue info
}

# --- Write JSON to local file (append mode) ---
try {
    if (-not (Test-Path $LogFilePath)) {
        # If the file doesn't exist, create it as an array with the first entry
        @($Payload) | ConvertTo-Json -Depth 6 | Out-File -FilePath $LogFilePath -Encoding utf8
    }
    else {
        # Read existing JSON
        $Existing = Get-Content $LogFilePath -Raw | ConvertFrom-Json

        # Ensure it's an array
        if ($Existing -isnot [System.Collections.IEnumerable] -or $Existing -is [string]) {
            $Existing = @($Existing)
        }

        # Append new record
        $Existing += $Payload

        # Write back to file
        $Existing | ConvertTo-Json -Depth 6 | Out-File -FilePath $LogFilePath -Encoding utf8
    }

    Write-Host "Appended data to $LogFilePath"
}
catch {
    Write-Warning "Failed to write JSON to file: $_"
}

# --- Send data to New Relic ---
Send-ToNewRelic -Payload $Payload
