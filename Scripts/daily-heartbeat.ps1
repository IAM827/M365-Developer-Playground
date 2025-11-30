<#
.SYNOPSIS
  Daily M365 Developer Heartbeat - optimized, robust, OneDrive-safe.

.PARAMETER TenantId
  Azure AD tenant ID (GUID) or tenant domain.

.PARAMETER ClientId
  App (client) ID for the Azure AD app.

.PARAMETER ClientSecret
  Client secret - store as GitHub secret and pass into the workflow.

.EXAMPLE
  .\daily-heartbeat.ps1 -TenantId $env:TENANT_ID -ClientId $env:CLIENT_ID -ClientSecret $env:CLIENT_SECRET
#>

param(
    [Parameter(Mandatory = $true)] [string]$TenantId,
    [Parameter(Mandatory = $true)] [string]$ClientId,
    [Parameter(Mandatory = $true)] [string]$ClientSecret
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-GraphAccessToken {
    param($TenantId, $ClientId, $ClientSecret)

    $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ErrorAction Stop
        return $resp.access_token
    } catch {
        throw "Failed to acquire token: $($_.Exception.Message)"
    }
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory=$true)][string]$Method,
        [Parameter(Mandatory=$true)][string]$Uri,
        [hashtable]$Headers,
        $Body = $null,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 2
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if ($Body) {
                $jsonBody = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
                return Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method -Body $jsonBody -ContentType "application/json" -ErrorAction Stop
            } else {
                return Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method -ErrorAction Stop
            }
        } catch {
            $status = $_.Exception.Response.StatusCode.value__ 2>$null
            Write-Host "Graph request failed (attempt $attempt/$MaxAttempts): $($_.Exception.Message)" -ForegroundColor Yellow
            if ($attempt -lt $MaxAttempts) {
                Start-Sleep -Seconds ($DelaySeconds * $attempt)
                continue
            } else {
                throw $_
            }
        }
    }
}

function Log {
    param([string]$Text, [ConsoleColor]$Color = "White")
    Write-Host $Text -ForegroundColor $Color
}

try {
    Log "=== M365 Developer Heartbeat Started ===" "Cyan"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Log "Timestamp: $timestamp" "Gray"

    # Acquire token
    $token = Get-GraphAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    $headers = @{ Authorization = "Bearer $token" }

    # 1) Get list of active users (sample top 50 for performance)
    Log "`n1. Enumerating candidate users..." "Yellow"
    $usersUri = "https://graph.microsoft.com/v1.0/users?`$filter=accountEnabled eq true&`$select=id,displayName,userPrincipalName,mail&`$top=50"
    $usersResp = Invoke-Graph -Method Get -Uri $usersUri -Headers $headers
    if (-not $usersResp.value -or $usersResp.value.Count -eq 0) {
        throw "No users returned from Graph. Ensure your app has User.Read.All (application) and admin consent."
    }

    # Pick a random user (seed if you want deterministic)
    $rnd = New-Object System.Random
    $targetUser = $usersResp.value[$rnd.Next(0, $usersResp.value.Count)]
    Log "Random User Selected: $($targetUser.displayName) ($($targetUser.userPrincipalName))" "Green"

    # 2) Randomizer - choose which activities to run (example)
    $diceRoll = $rnd.Next(1,11)  # 1..10
    Log "Daily Randomizer Roll: $diceRoll / 10" "Gray"

    # 3) Optional: SharePoint activity (only if roll <= 6)
    if ($diceRoll -le 6) {
        Log "`n2. Accessing SharePoint sites (sample)..." "Yellow"
        try {
            $sites = Invoke-Graph -Method Get -Uri "https://graph.microsoft.com/v1.0/sites?`$top=10&`$select=id,displayName,webUrl" -Headers $headers
            if ($sites.value) {
                foreach ($s in $sites.value) { Log "   - Site: $($s.displayName) ($($s.webUrl))" "DarkGray" }
            }
        } catch {
            Log "   - Warning: Could not fetch SharePoint sites: $($_.Exception.Message)" "Yellow"
        }
    } else {
        Log "`n2. Skipping SharePoint activity today (Randomizer)" "DarkGray"
    }

    # 4) Exchange: read top 3 messages from the user's mailbox (app-only requires Mail.Read application permission)
    Log "`n3. Checking recent emails (if permitted by app permissions)..." "Yellow"
    try {
        $mailUri = "https://graph.microsoft.com/v1.0/users/$($targetUser.id)/messages?`$top=3&`$select=subject,receivedDateTime,from"
        $messages = Invoke-Graph -Method Get -Uri $mailUri -Headers $headers
        if ($messages.value) {
            foreach ($m in $messages.value) {
                Log "   - [$($m.receivedDateTime)] $($m.subject) (from: $($m.from.emailAddress.address))" "DarkGray"
            }
        } else {
            Log "   - No recent messages or no Mail.Read application permission." "DarkGray"
        }
    } catch {
        Log "   - Mail read skipped or failed: $($_.Exception.Message)" "Yellow"
    }

    # 5) OneDrive write/delete activity (runs when diceRoll >= 4)
    if ($diceRoll -ge 4) {
        Log "`n4. Performing OneDrive Write/Delete Activity..." "Yellow"

        # 5.1 Check if user's OneDrive (drive) exists/provisioned
        $driveUri = "https://graph.microsoft.com/v1.0/users/$($targetUser.id)/drive"
        try {
            $drive = Invoke-Graph -Method Get -Uri $driveUri -Headers $headers
        } catch {
            # Catch 403/404 or other errors
            $err = $_.Exception.Message
            Log "   - OneDrive check failed: $err" "Yellow"

            # If 403 or 404, skip OneDrive activity for this user
            Log "   - Skipping OneDrive activity for this user (no drive/provisioning or no permission)." "DarkGray"
            $drive = $null
        }

        if ($drive) {
            # Use a consistent heartbeat folder
            $heartbeatFolder = "heartbeat"
            $fileName = "heartbeat-report-$((Get-Date).ToString('yyyyMMdd-HHmmss')).json"
            $folderPath = "/$heartbeatFolder"
            $folderCreateUri = "https://graph.microsoft.com/v1.0/users/$($targetUser.id)/drive/root:$folderPath"

            # 5.2 Ensure folder exists (create if 404)
            try {
                # Try to get the folder metadata
                $folderMeta = Invoke-Graph -Method Get -Uri "$folderCreateUri" -Headers $headers
                Log "   - Found existing heartbeat folder." "DarkGray"
            } catch {
                # If folder not found, create it
                $msg = $_.Exception.Message
                if ($msg -match "404" -or $msg -match "Not Found") {
                    Log "   - Heartbeat folder not found. Creating..." "Gray"
                    $createFolderUri = "https://graph.microsoft.com/v1.0/users/$($targetUser.id)/drive/root/children"
                    $folderPayload = @{ name = $heartbeatFolder; folder = @{}; "@microsoft.graph.conflictBehavior" = "replace" }
                    try {
                        $createRes = Invoke-Graph -Method Post -Uri $createFolderUri -Headers $headers -Body $folderPayload
                        if ($createRes.id) { Log "   - Heartbeat folder created." "DarkGray" }
                    } catch {
                        Log "   - Failed to create heartbeat folder: $($_.Exception.Message)" "Yellow"
                        # if creation fails, skip oneDrive activity
                        $drive = $null
                    }
                } else {
                    Log "   - Unexpected error while checking folder: $msg" "Yellow"
                    $drive = $null
                }
            }

            if ($drive) {
                # 5.3 Upload a small JSON report directly into heartbeat folder using content endpoint
                $uploadUri = "https://graph.microsoft.com/v1.0/users/$($targetUser.id)/drive/root:/$heartbeatFolder/${fileName}:/content"
                $smallReport = @{
                    Timestamp = (Get-Date).ToString("o")
                    TargetUser = $targetUser.userPrincipalName
                    Dice = $diceRoll
                    Note = "Automated heartbeat"
                } | ConvertTo-Json -Depth 5

                try {
                    # Use PUT to create new file with content
                    $response = Invoke-RestMethod -Method Put -Uri $uploadUri -Headers @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" } -Body $smallReport -ErrorAction Stop
                    if ($response.id) {
                        Log "   - Uploaded heartbeat file: ${fileName}" "Green"
                        # Optionally delete the file (simulate write+delete) - keep this optional
                        # To remove, uncomment:
                        # Invoke-Graph -Method Delete -Uri "https://graph.microsoft.com/v1.0/users/$($targetUser.id)/drive/items/$($response.id)" -Headers $headers
                        # Log "   - Deleted heartbeat file: $($response.id)" "DarkGray"
                    } else {
                        Log "   - Upload did not return an item id." "Yellow"
                    }
                } catch {
                    Log "   - Failed to upload heartbeat file: $($_.Exception.Message)" "Yellow"
                }
            } else {
                Log "   - Skipping upload: drive/folder unavailable." "DarkGray"
            }
        }
    } else {
        Log "`n4. Skipping OneDrive activity today (Randomizer)" "DarkGray"
    }

    # 6) Save local report for audit (runner workspace)
    $localTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $localFile = "heartbeat-report-$localTimestamp.json"
    $reportObj = @{
        Status = "Success"
        Timestamp = $timestamp
        RandomDiceRoll = $diceRoll
        TargetUser = if ($targetUser) { $targetUser.userPrincipalName } else { "None" }
        ActivityType = "Daily Heartbeat (Randomized)"
    }
    $reportObj | ConvertTo-Json -Depth 6 | Out-File -FilePath $localFile -Encoding UTF8
    Log "`nReport saved to local file: $localFile" "White"

    Log "`n=== Heartbeat Completed Successfully ===" "Cyan"
    exit 0
}
catch {
    Write-Error "Heartbeat failed: $($_.Exception.Message)"
    exit 1
}
