[CmdletBinding()]
param(
    [string]$TenantId = "4744e69f-9ae0-413e-8cf2-0deea4bf6952",
    [string]$ClientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46" # Microsoft Graph PowerShell public client
)

$scopes = @("User.Read.All")

if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
    Write-Error "The Microsoft Graph PowerShell SDK is required. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

Write-Host "Signing in to tenant $TenantId with delegated permissions using Microsoft Graph PowerShell app (clientId=$ClientId)..." -ForegroundColor Cyan
Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Scopes $scopes -NoWelcome -ErrorAction Stop | Out-Null

try {
    # $users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Mail, AccountEnabled |
    #          Select-Object DisplayName, UserPrincipalName, Mail, AccountEnabled, Id |
    #          Sort-Object DisplayName
    #
    # if (-not $users) {
    #     Write-Host "No users returned for tenant $TenantId." -ForegroundColor Yellow
    #     return
    # }
    #
    # $users | Format-Table -AutoSize
}
finally {
    # Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

$context = Get-MgContext
$context | Format-List

$accessToken = $context.AccessToken
if (-not $accessToken -and $context.MsalToken) {
    # Some module versions leave AccessToken empty but keep the MSAL authentication result.
    $accessToken = $context.MsalToken.AccessToken
}

if ($accessToken) {
    Write-Host "`nAccess token:`n$accessToken"
} else {
    Write-Host "`nAccess token not available in current context." -ForegroundColor Yellow
}
