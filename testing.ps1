# --- Load Exchange Management Shell if not already loaded ---
try {
    if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
    }
}
catch {
    try {
        # For newer Exchange versions (Exchange 2016/2019+), connect to local PowerShell session
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://localhost/powershell/ -Authentication Kerberos
        Import-PSSession $Session -DisableNameChecking | Out-Null
    }
    catch {
        Write-Error "Failed to load Exchange Management Shell environment: $($_.Exception.Message)"
        exit 1
    }
}

# --- Configuration ---
$EventType = "ExchangeCheckData"

# --- Run Exchange Queue Command ---
try {
    # Remove threshold filter so all queues are collected (even zero messages)
    $QueueData = Get-Queue | Where-Object { $_.DeliveryType -ne "ShadowRedundancy" } |
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

# --- Build payload for Flex ---
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
    scriptVersion     = "1.5"
    queues            = @($QueueData)
}

# --- Output JSON to stdout (for Flex) ---
$Payload | ConvertTo-Json -Depth 6
