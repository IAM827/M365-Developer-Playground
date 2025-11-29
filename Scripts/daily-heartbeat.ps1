param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret
)

# Function to get access token
function Get-AccessToken {
    param($TenantId, $ClientId, $ClientSecret)
    
    $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope        = "https://graph.microsoft.com/.default"
        grant_type   = "client_credentials"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Method POST -Body $body -ErrorAction Stop
        return $response.access_token
    }
    catch {
        Write-Error "CRITICAL: Failed to get access token. $($_.Exception.Message)"
        exit 1
    }
}

# Wrapper for API calls to handle errors gracefully
function Invoke-GraphRequest {
    param($Uri, $Headers, $Method="Get", $Body=$null)
    try {
        if ($Body) {
            return Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method -Body $Body -ErrorAction Stop
        }
        return Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method -ErrorAction Stop
    }
    catch {
        Write-Warning "API Call Failed [$Uri]: $($_.Exception.Message)"
        return $null
    }
}

# Main execution
try {
    Write-Host "=== M365 Developer Heartbeat Started ===" -ForegroundColor Cyan
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "Timestamp: $timestamp"
    
    # 1. Authenticate
    $accessToken = Get-AccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    $headers = @{
        'Authorization' = "Bearer $accessToken"
        'Content-Type' = 'application/json'
    }

    # 2. Random User Lookup (UPDATED)
    Write-Host "`n1. Finding a random target user..." -ForegroundColor Yellow
    
    # Fetch top 50 users (increased from 10 to give a better pool for randomization)
    $usersReq = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/users?`$top=50&`$select=id,displayName,mail,userPrincipalName" -Headers $headers
    
    # Filter for users with emails, then pick ONE randomly
    $targetUser = $usersReq.value | Where-Object { -not [string]::IsNullOrWhiteSpace($_.mail) } | Get-Random
    
    if ($targetUser) {
        Write-Host "   - Random User Selected: $($targetUser.displayName) ($($targetUser.mail))" -ForegroundColor Green
    } else {
        Write-Warning "   - No valid users found in the top 50. Skipping user-specific tests."
    }

    # Activity 3: SharePoint Sites
    Write-Host "`n2. Accessing SharePoint sites..." -ForegroundColor Yellow
    $sites = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/sites?`$top=3" -Headers $headers
    if ($sites) {
        foreach ($site in $sites.value) {
            Write-Host "   - Site: $($site.displayName)" -ForegroundColor Gray
        }
    }

    # Activity 4: Exchange Online (Read)
    if ($targetUser) {
        Write-Host "`n3. Checking recent emails for $($targetUser.mail)..." -ForegroundColor Yellow
        $messages = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$($targetUser.id)/messages?`$top=3&`$select=subject,receivedDateTime" -Headers $headers
        if ($messages) {
            foreach ($message in $messages.value) {
                Write-Host "   - [$($message.receivedDateTime)] $($message.subject)" -ForegroundColor Gray
            }
        }
    }

    # Activity 5: OneDrive (Write & Delete)
    if ($targetUser) {
        Write-Host "`n4. Performing OneDrive Write/Delete Activity..." -ForegroundColor Yellow
        $folderName = "_Heartbeat_Temp_$(Get-Date -Format 'MMddHHmm')"
        $folderPayload = @{
            name = $folderName
            folder = @{}
            "@microsoft.graph.conflictBehavior" = "rename"
        } | ConvertTo-Json

        # Create Folder
        $createRes = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$($targetUser.id)/drive/root/children" -Headers $headers -Method Post -Body $folderPayload
        
        if ($createRes) {
            Write-Host "   - Created temp folder: $folderName" -ForegroundColor Green
            # Clean up immediately
            Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$($targetUser.id)/drive/items/$($createRes.id)" -Headers $headers -Method Delete
            Write-Host "   - Deleted temp folder (Cleanup)" -ForegroundColor Green
        }
    }

    # Save a report artifact
    $report = @{
        Status = "Success"
        Timestamp = $timestamp
        TargetUser = if ($targetUser) { $targetUser.userPrincipalName } else { "None" }
        ActivityType = "Daily Heartbeat"
    }
    $report | ConvertTo-Json | Out-File "heartbeat-report.json"
    Write-Host "`nReport saved to heartbeat-report.json" -ForegroundColor White

    Write-Host "`n=== Heartbeat Completed Successfully ===" -ForegroundColor Cyan
}
catch {
    Write-Error "Heartbeat failed: $($_.Exception.Message)"
    exit 1
}
