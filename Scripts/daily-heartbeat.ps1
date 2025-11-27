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
        $response = Invoke-RestMethod -Uri $uri -Method POST -Body $body
        return $response.access_token
    }
    catch {
        Write-Error "Failed to get access token: $($_.Exception.Message)"
        exit 1
    }
}

# Main execution
try {
    Write-Host "=== M365 Developer Heartbeat Started ==="
    Write-Host "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    
    # Get access token
    $accessToken = Get-AccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    $headers = @{
        'Authorization' = "Bearer $accessToken"
        'Content-Type' = 'application/json'
    }

    # Activity 1: Get tenant users count
    Write-Host "`n1. Getting tenant users information..."
    $users = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users?`$top=5&`$select=displayName,userPrincipalName" -Headers $headers
    Write-Host "   - Found $($users.value.count) users"
    foreach ($user in $users.value[0..2]) {
        Write-Host "   - User: $($user.displayName)"
    }

    # Activity 2: Get SharePoint sites
    Write-Host "`n2. Accessing SharePoint sites..."
    $sites = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites?`$top=3" -Headers $headers
    foreach ($site in $sites.value) {
        Write-Host "   - Site: $($site.displayName)"
    }

    # Activity 3: Get recent emails (metadata only)
    Write-Host "`n3. Checking recent emails..."
    $messages = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/admin@thelhub.co.za/messages?`$top=3&`$select=subject,receivedDateTime" -Headers $headers
    foreach ($message in $messages.value) {
        Write-Host "   - [$($message.receivedDateTime)] $($message.subject)"
    }

    # Activity 4: Get Teams presence (if available)
    Write-Host "`n4. Checking Microsoft 365 services..."
    $me = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/admin@thelhub.co.za" -Headers $headers
    Write-Host "   - User: $($me.displayName)"
    Write-Host "   - Mail: $($me.mail)"
    Write-Host "   - Department: $($me.department)"

    Write-Host "`n=== M365 Developer Heartbeat Completed Successfully ==="
    Write-Host "Activities logged at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}
catch {
    Write-Error "Heartbeat failed: $($_.Exception.Message)"
    exit 1
}
