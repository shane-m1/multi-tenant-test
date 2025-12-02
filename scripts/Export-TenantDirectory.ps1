[CmdletBinding()]
param(
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$TenantCsvPath,
    [string]$OutputDirectory = "./tenant-exports",
    [string]$TenantIdColumn = "TenantId",
    [string[]]$Retrieve = @("Users", "Groups", "ServicePrincipals", "AppPermissions", "Roles", "Applications", "UserAuthMethods")
)

$scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
if (-not $scriptDirectory) {
    $scriptDirectory = Get-Location
}
$envFilePath = Join-Path $scriptDirectory ".env"
$knownOptionKeys = @("ClientId", "ClientSecret", "TenantCsvPath", "OutputDirectory", "TenantIdColumn", "Retrieve")

$envData = @{}
if (Test-Path -Path $envFilePath) {
    foreach ($line in Get-Content -Path $envFilePath) {
        if ($line -match "^\s*(#|$)") {
            continue
        }

        $pair = $line -split "=", 2
        if ($pair.Count -eq 2) {
            $key = $pair[0].Trim()
            $value = $pair[1].Trim()

            if ($key) {
                $envData[$key] = $value
            }
        }
    }
}

if (-not $ClientId -and $envData.ContainsKey("ClientId")) {
    $ClientId = $envData["ClientId"]
}

if (-not $ClientSecret -and $envData.ContainsKey("ClientSecret")) {
    $ClientSecret = $envData["ClientSecret"]
}

if (-not $TenantCsvPath -and $envData.ContainsKey("TenantCsvPath")) {
    $TenantCsvPath = $envData["TenantCsvPath"]
}

if (-not $PSBoundParameters.ContainsKey("OutputDirectory") -and $envData.ContainsKey("OutputDirectory")) {
    $OutputDirectory = $envData["OutputDirectory"]
}

if (-not $PSBoundParameters.ContainsKey("TenantIdColumn") -and $envData.ContainsKey("TenantIdColumn")) {
    $TenantIdColumn = $envData["TenantIdColumn"]
}

if (-not $PSBoundParameters.ContainsKey("Retrieve") -and $envData.ContainsKey("Retrieve")) {
    $Retrieve = $envData["Retrieve"] -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

$missing = @()
if (-not $ClientId) {
    $missing += "ClientId"
}
if (-not $ClientSecret) {
    $missing += "ClientSecret"
}
if (-not $TenantCsvPath) {
    $missing += "TenantCsvPath"
}

if ($missing.Count -gt 0) {
    throw "Missing required settings: $($missing -join ', '). Provide them as parameters or via the .env file at $envFilePath."
}

$shouldPersistEnv = $false
foreach ($key in $knownOptionKeys) {
    if ($PSBoundParameters.ContainsKey($key)) {
        $shouldPersistEnv = $true
        break
    }
}

if ($shouldPersistEnv) {
    $persistData = @{}

    foreach ($entry in $envData.GetEnumerator()) {
        $persistData[$entry.Key] = $entry.Value
    }

    foreach ($key in $knownOptionKeys) {
        if ($PSBoundParameters.ContainsKey($key)) {
            $persistData[$key] = (Get-Variable -Name $key -ValueOnly)
        }
    }

    $envContent = @()
    foreach ($key in $knownOptionKeys) {
        if ($persistData.ContainsKey($key)) {
            $valueToPersist = $persistData[$key]
            if ($valueToPersist -is [System.Array]) {
                $valueToPersist = $valueToPersist -join ","
            }

            $envContent += "$key=$valueToPersist"
        }
    }

    Set-Content -Path $envFilePath -Value $envContent -Encoding UTF8
}

if (-not (Test-Path -Path $TenantCsvPath)) {
    throw "CSV file not found at $TenantCsvPath"
}

if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
    Write-Error "The Microsoft Graph PowerShell SDK is required. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Groups -ErrorAction Stop
Import-Module Microsoft.Graph.DirectoryObjects -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
Import-Module Microsoft.Graph.Applications -ErrorAction Stop

# Normalize retrieval selection (default is all).
$retrieveSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($item in $Retrieve) {
    if ([string]::IsNullOrWhiteSpace($item)) { continue }
    $trimmed = $item.Trim()
    if ($trimmed -ieq "AppRegistrations" -or $trimmed -ieq "Apps") {
        $trimmed = "Applications"
    }
    elseif ($trimmed -ieq "AuthMethods" -or $trimmed -ieq "UserAuthenticationMethods") {
        $trimmed = "UserAuthMethods"
    }

    [void]$retrieveSet.Add($trimmed)
}

$retrieveUsers = $retrieveSet.Count -eq 0 -or $retrieveSet.Contains("Users")
$retrieveGroups = $retrieveSet.Count -eq 0 -or $retrieveSet.Contains("Groups")
$retrieveAppPermissions = $retrieveSet.Count -eq 0 -or $retrieveSet.Contains("AppPermissions")
$retrieveServicePrincipals = $retrieveSet.Count -eq 0 -or $retrieveSet.Contains("ServicePrincipals") -or $retrieveAppPermissions
$retrieveRoles = $retrieveSet.Count -eq 0 -or $retrieveSet.Contains("Roles")
$retrieveApplications = $retrieveSet.Count -eq 0 -or $retrieveSet.Contains("Applications")
$retrieveUserAuthMethods = $retrieveSet.Count -eq 0 -or $retrieveSet.Contains("UserAuthMethods")

$tenants = Import-Csv -Path $TenantCsvPath

if (-not $tenants) {
    throw "No tenant entries found in CSV."
}

if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
}

$secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$appCredential = New-Object System.Management.Automation.PSCredential ($ClientId, $secureSecret)


function Connect-Tenant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    try {
        Connect-MgGraph -TenantId $TenantId `
                        -ClientSecretCredential $appCredential `
                        -NoWelcome `
                        -ErrorAction Stop | Out-Null
        return
    }
    catch {
        if ($_.Exception.Message -notmatch "ClientSecretCredential") {
            throw
        }
    }

    if (-not ("Azure.Identity.ClientSecretCredential" -as [Type])) {
        try {
            Add-Type -AssemblyName "Azure.Identity" -ErrorAction Stop
        }
        catch {
            throw "Failed to load Azure.Identity assembly required for ClientSecretCredential. $($_.Exception.Message)"
        }
    }


    Connect-MgGraph -ClientSecretCredential $appCredential `
                    -NoWelcome `
                    -ErrorAction Stop | Out-Null
}

function Resolve-Principal {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Cache,

        [Parameter(Mandatory = $true)]
        [string]$PrincipalId,

        [string]$FallbackDisplayName,

        [string]$FallbackType
    )

    if ($Cache.ContainsKey($PrincipalId)) {
        return $Cache[$PrincipalId]
    }

    $displayName = $FallbackDisplayName
    $principalType = $FallbackType

    if (-not $displayName -or -not $principalType) {
        $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $PrincipalId -ErrorAction SilentlyContinue

        if ($directoryObject) {
            $principalType = $directoryObject.AdditionalProperties['@odata.type']
            $displayName = $directoryObject.AdditionalProperties['displayName']

            if (-not $displayName -and $directoryObject.AdditionalProperties['userPrincipalName']) {
                $displayName = $directoryObject.AdditionalProperties['userPrincipalName']
            }
        }
    }

    if (-not $principalType) {
        $principalType = "unknown"
    }

    $info = [pscustomobject]@{
        DisplayName   = $displayName
        ObjectType    = $principalType
    }

    $Cache[$PrincipalId] = $info
    return $info
}

function Get-GraphPagedResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $results = @()
    $nextUri = $Uri

    while ($nextUri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop

        if ($response.value) {
            $results += $response.value
        }

        $nextUri = $response.'@odata.nextLink'
    }

    return $results
}

$usersResults = @()
$groupsResults = @()
$roleResults = @()
$servicePrincipalResults = @()
$appPermissionResults = @()
$applicationOwnerResults = @()
$userAuthMethodResults = @()

foreach ($tenant in $tenants) {
    if (-not $tenant.PSObject.Properties.Match($TenantIdColumn)) {
        throw "Column '$TenantIdColumn' not present in CSV."
    }

    $tenantId = $tenant.$TenantIdColumn

    if ([string]::IsNullOrWhiteSpace($tenantId)) {
        Write-Warning "Skipping a row with empty tenant identifier."
        continue
    }

    Write-Host "Processing tenant $tenantId..." -ForegroundColor Cyan

    try {
        Connect-Tenant -TenantId $tenantId

        $tenantUsers = @()
        if ($retrieveUsers -or $retrieveUserAuthMethods) {
            Write-Host "Retrieving users..." -ForegroundColor Yellow
            $tenantUsers = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Mail, AccountEnabled, UserType, ExternalUserState |
                Select-Object @{Name = "TenantId"; Expression = { $tenantId }},
                              Id,
                              DisplayName,
                              UserPrincipalName,
                              Mail,
                              AccountEnabled,
                              UserType,
                              ExternalUserState,
                              @{Name = "GuestInviteAccepted"; Expression = {
                                  if ($_.UserType -eq "Guest" -and $_.ExternalUserState) {
                                      return $_.ExternalUserState -eq "Accepted"
                                  }
                                  return $null
                              }}

            if ($retrieveUsers) {
                $usersResults += $tenantUsers
            }
        }

        if ($retrieveUserAuthMethods -and $tenantUsers.Count -gt 0) {
            Write-Host "Retrieving user authentication methods..." -ForegroundColor Yellow

            foreach ($user in $tenantUsers) {
                $userIdEncoded = [uri]::EscapeDataString($user.Id)
                $methodsUri = "https://graph.microsoft.com/v1.0/users/$userIdEncoded/authentication/methods"
                try {
                    $userMethods = Get-GraphPagedResult -Uri $methodsUri
                }
                catch {
                    Write-Warning "Failed to retrieve authentication methods for user $($user.UserPrincipalName) in tenant $tenantId. $($_.Exception.Message)"
                    continue
                }

                foreach ($method in $userMethods) {
                    $additional = $method.AdditionalProperties
                    $methodType = $method.'@odata.type'
                    if (-not $methodType -and $additional) { $methodType = $additional['@odata.type'] }

                    $phoneNumber = $method.phoneNumber
                    if (-not $phoneNumber -and $additional) { $phoneNumber = $additional['phoneNumber'] }

                    $phoneType = $method.phoneType
                    if (-not $phoneType -and $additional) { $phoneType = $additional['phoneType'] }

                    $emailAddress = $method.emailAddress
                    if (-not $emailAddress -and $additional) { $emailAddress = $additional['emailAddress'] }

                    $displayName = $method.displayName
                    if (-not $displayName -and $additional) { $displayName = $additional['displayName'] }

                    $deviceTag = $method.deviceTag
                    if (-not $deviceTag -and $additional) { $deviceTag = $additional['deviceTag'] }

                    $keyId = $method.keyId
                    if (-not $keyId -and $additional) { $keyId = $additional['keyId'] }

                    $state = $method.state
                    if (-not $state -and $additional) { $state = $additional['state'] }

                    $creationDateTime = $method.creationDateTime
                    if (-not $creationDateTime -and $additional) { $creationDateTime = $additional['creationDateTime'] }

                    $userAuthMethodResults += [pscustomobject]@{
                        TenantId           = $tenantId
                        UserId             = $user.Id
                        UserPrincipalName  = $user.UserPrincipalName
                        UserDisplayName    = $user.DisplayName
                        MethodId           = $method.Id
                        MethodType         = $methodType
                        PhoneNumber        = $phoneNumber
                        PhoneType          = $phoneType
                        EmailAddress       = $emailAddress
                        DeviceTag          = $deviceTag
                        DisplayName        = $displayName
                        KeyId              = $keyId
                        State              = $state
                        CreationDateTime   = $creationDateTime
                    }
                }
            }
        }

        if ($retrieveGroups) {
            Write-Host "Retrieving groups..." -ForegroundColor Yellow
            $tenantGroups = Get-MgGroup -All -Property Id, DisplayName, Mail, MailEnabled, SecurityEnabled, GroupTypes |
                Select-Object @{Name = "TenantId"; Expression = { $tenantId }},
                              Id,
                              DisplayName,
                              Mail,
                              MailEnabled,
                              SecurityEnabled,
                              @{Name = "GroupTypes"; Expression = { $_.GroupTypes -join ";" }}

            $groupsResults += $tenantGroups
        }

        if ($retrieveApplications) {
            Write-Host "Retrieving application registrations..." -ForegroundColor Yellow
            $tenantApplications = Get-MgApplication -All -Property Id, AppId, DisplayName, SignInAudience, VerifiedPublisher

            foreach ($app in $tenantApplications) {
                $owners = Get-MgApplicationOwner -ApplicationId $app.Id -All -ErrorAction SilentlyContinue

                foreach ($owner in $owners) {
                    $ownerType = $owner.AdditionalProperties['@odata.type']
                    $ownerDisplayName = $owner.AdditionalProperties['displayName']
                    $ownerUpn = $owner.AdditionalProperties['userPrincipalName']
                    $ownerMail = $owner.AdditionalProperties['mail']
                    $ownerId = $owner.Id

                    if (-not $ownerDisplayName -and $owner.AdditionalProperties['appDisplayName']) {
                        $ownerDisplayName = $owner.AdditionalProperties['appDisplayName']
                    }

                    $verifiedPublisherId = $app.VerifiedPublisher.VerifiedPublisherId
                    $verifiedPublisherDisplayName = $app.VerifiedPublisher.DisplayName
                    $applicationOwnerResults += [pscustomobject]@{
                        TenantId                 = $tenantId
                        ApplicationObjectId      = $app.Id
                        ApplicationAppId         = $app.AppId
                        ApplicationDisplayName   = $app.DisplayName
                        ApplicationSignInAudience = $app.SignInAudience
                        ApplicationVerifiedPublisherId = $verifiedPublisherId
                        ApplicationVerifiedPublisherDisplayName = $verifiedPublisherDisplayName
                        ApplicationHasVerifiedPublisher = [bool]$verifiedPublisherId
                        OwnerId                  = $ownerId
                        OwnerDisplayName         = $ownerDisplayName
                        OwnerUserPrincipalName   = $ownerUpn
                        OwnerMail                = $ownerMail
                        OwnerType                = $ownerType
                        MembershipSource         = "DirectOwner"
                        OwningGroupId            = $null
                        OwningGroupDisplayName   = $null
                    }

                    if ($ownerType -and $ownerType -like "*group*") {
                        $groupMembers = Get-MgGroupTransitiveMember -GroupId $ownerId -All -Property Id, DisplayName, UserPrincipalName, Mail, "@odata.type"

                        foreach ($member in $groupMembers) {
                            $memberType = $member.AdditionalProperties['@odata.type']
                            $memberDisplayName = $member.AdditionalProperties['displayName']
                            $memberUpn = $member.AdditionalProperties['userPrincipalName']
                            $memberMail = $member.AdditionalProperties['mail']

                            $applicationOwnerResults += [pscustomobject]@{
                                TenantId                 = $tenantId
                                ApplicationObjectId      = $app.Id
                                ApplicationAppId         = $app.AppId
                                ApplicationDisplayName   = $app.DisplayName
                                ApplicationSignInAudience = $app.SignInAudience
                                ApplicationVerifiedPublisherId = $verifiedPublisherId
                                ApplicationVerifiedPublisherDisplayName = $verifiedPublisherDisplayName
                                ApplicationHasVerifiedPublisher = [bool]$verifiedPublisherId
                                OwnerId                  = $member.Id
                                OwnerDisplayName         = $memberDisplayName
                                OwnerUserPrincipalName   = $memberUpn
                                OwnerMail                = $memberMail
                                OwnerType                = $memberType
                                MembershipSource         = "GroupMemberOfOwnerGroup"
                                OwningGroupId            = $ownerId
                                OwningGroupDisplayName   = $ownerDisplayName
                            }
                        }
                    }
                }
            }
        }

        $tenantServicePrincipalsRaw = @()
        if ($retrieveServicePrincipals) {
            Write-Host "Retrieving service principals..." -ForegroundColor Yellow
            $tenantServicePrincipalsRaw = Get-MgServicePrincipal -All -Property Id, AppId, DisplayName, ServicePrincipalType, AccountEnabled, AppOwnerOrganizationId, Tags

            if ($retrieveSet.Contains("ServicePrincipals") -or $retrieveSet.Count -eq 0) {
                $servicePrincipalResults += $tenantServicePrincipalsRaw | Select-Object @{Name = "TenantId"; Expression = { $tenantId }},
                                                                                       Id,
                                                                                       AppId,
                                                                                       DisplayName,
                                                                                       ServicePrincipalType,
                                                                                       AccountEnabled,
                                                                                       AppOwnerOrganizationId,
                                                                                       @{Name = "Tags"; Expression = { $_.Tags -join ";" }}
            }
        }

        if ($retrieveAppPermissions -and $tenantServicePrincipalsRaw.Count -gt 0) {
            Write-Host "Retrieving application permissions for service principals..." -ForegroundColor Yellow
            $resourceAppRoleCache = @{}
            $resourceDisplayNameCache = @{}
            foreach ($sp in $tenantServicePrincipalsRaw) {
                $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction SilentlyContinue

                foreach ($assignment in $appRoleAssignments) {
                    $resourceDisplayName = $assignment.ResourceDisplayName
                    $appRoleDisplayName = $null
                    $appRoleValue = $null

                    if ($assignment.ResourceId) {
                        $resourceIdKey = $assignment.ResourceId.ToString()

                        if (-not $resourceAppRoleCache.ContainsKey($resourceIdKey)) {
                            $resourceAppRoleCache[$resourceIdKey] = @{}
                            $resourceSp = Get-MgServicePrincipal -ServicePrincipalId $assignment.ResourceId -Property AppRoles, DisplayName -ErrorAction SilentlyContinue

                            if ($resourceSp) {
                                $resourceDisplayNameCache[$resourceIdKey] = $resourceSp.DisplayName

                                if ($resourceSp.AppRoles) {
                                    foreach ($role in $resourceSp.AppRoles) {
                                        if ($role.Id) {
                                            $resourceAppRoleCache[$resourceIdKey][$role.Id.ToString()] = @{
                                                DisplayName = $role.DisplayName
                                                Value       = $role.Value
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (-not $resourceDisplayName -and $resourceDisplayNameCache.ContainsKey($resourceIdKey)) {
                            $resourceDisplayName = $resourceDisplayNameCache[$resourceIdKey]
                        }

                        if ($assignment.AppRoleId) {
                            $roleKey = $assignment.AppRoleId.ToString()

                            if ($resourceAppRoleCache.ContainsKey($resourceIdKey) -and
                                $resourceAppRoleCache[$resourceIdKey].ContainsKey($roleKey)) {
                                $roleInfo = $resourceAppRoleCache[$resourceIdKey][$roleKey]
                                $appRoleDisplayName = $roleInfo.DisplayName
                                $appRoleValue = $roleInfo.Value
                            }
                        }
                    }

                    $appPermissionResults += [pscustomobject]@{
                        TenantId                      = $tenantId
                        ServicePrincipalId            = $sp.Id
                        ServicePrincipalAppId         = $sp.AppId
                        ServicePrincipalDisplayName   = $sp.DisplayName
                        ResourceId                    = $assignment.ResourceId
                        ResourceDisplayName           = $resourceDisplayName
                        AppRoleDisplayName            = $appRoleDisplayName
                        AppRoleValue                  = $appRoleValue
                        PrincipalId                   = $assignment.PrincipalId
                        PrincipalType                 = $assignment.PrincipalType
                    }
                }
            }
        }

        if ($retrieveRoles) {
            Write-Host "Retrieving role assignments..." -ForegroundColor Yellow

            $principalCache = @{}

            $directoryRoles = Get-MgDirectoryRole -All -ErrorAction SilentlyContinue
            if ($directoryRoles) {
                foreach ($role in $directoryRoles) {
                    $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue

                    foreach ($member in $members) {
                        $principalInfo = Resolve-Principal -Cache $principalCache `
                                                          -PrincipalId $member.Id `
                                                          -FallbackDisplayName $member.AdditionalProperties['displayName'] `
                                                          -FallbackType $member.AdditionalProperties['@odata.type']

                        $roleResults += [pscustomobject]@{
                            TenantId            = $tenantId
                            AssignmentType      = "Active"
                            RoleId              = $role.Id
                            RoleDisplayName     = $role.DisplayName
                            PrincipalId         = $member.Id
                            PrincipalType       = $principalInfo.ObjectType
                            PrincipalDisplayName = $principalInfo.DisplayName
                        }
                    }
                }
            }

            try {
                $roleDefinitions = @{}
                $roleDefinitionData = Get-MgDirectoryRole -All -ErrorAction SilentlyContinue

                foreach ($definition in $roleDefinitionData) {
                    if ($definition.Id -and -not $roleDefinitions.ContainsKey($definition.Id)) {
                        $roleDefinitions[$definition.Id] = $definition.DisplayName
                    }
                    if ($definition.RoleTemplateId -and -not $roleDefinitions.ContainsKey($definition.RoleTemplateId)) {
                        $roleDefinitions[$definition.RoleTemplateId] = $definition.DisplayName
                    }
                }

                $eligibilitySchedules = Get-GraphPagedResult -Uri "https://graph.microsoft.com/beta/roleManagement/directory/roleEligibilitySchedules`?\$select=id,roleDefinitionId,principalId,principalType,memberType"

                foreach ($schedule in $eligibilitySchedules) {
                    $principalInfo = Resolve-Principal -Cache $principalCache `
                                                      -PrincipalId $schedule.principalId `
                                                      -FallbackDisplayName $null `
                                                      -FallbackType $schedule.principalType

                    $roleDisplayName = $null
                    if ($schedule.roleDefinitionId -and $roleDefinitions.ContainsKey($schedule.roleDefinitionId)) {
                        $roleDisplayName = $roleDefinitions[$schedule.roleDefinitionId]
                    }

                    $roleResults += [pscustomobject]@{
                        TenantId             = $tenantId
                        AssignmentType       = "Eligible"
                        RoleId               = $schedule.roleDefinitionId
                        RoleDisplayName      = $roleDisplayName
                        PrincipalId          = $schedule.principalId
                        PrincipalType        = $principalInfo.ObjectType
                        PrincipalDisplayName = $principalInfo.DisplayName
                        MemberType           = $schedule.memberType
                    }
                }
            }
            catch {
                Write-Warning "Failed to retrieve eligibility schedules for tenant $tenantId. $($_.Exception.Message)"
            }
        }

        Write-Host "Tenant $tenantId complete." -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to process tenant $tenantId. $($_.Exception.Message)"
    }
    finally {
        # Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

if ($retrieveUsers) {
    if ($usersResults.Count -gt 0) {
        $userExportPath = Join-Path $OutputDirectory "AllTenants-Users.csv"
        $usersResults | Export-Csv -Path $userExportPath -NoTypeInformation
        Write-Host "User export written to $userExportPath" -ForegroundColor Green
    }
    else {
        Write-Host "No user data collected." -ForegroundColor Yellow
    }
}
else {
    Write-Host "User export skipped (not requested)." -ForegroundColor Yellow
}

if ($retrieveUserAuthMethods) {
    if ($userAuthMethodResults.Count -gt 0) {
        $authMethodsExportPath = Join-Path $OutputDirectory "AllTenants-UserAuthenticationMethods.csv"
        $userAuthMethodResults | Export-Csv -Path $authMethodsExportPath -NoTypeInformation
        Write-Host "User authentication methods export written to $authMethodsExportPath" -ForegroundColor Green
    }
    else {
        Write-Host "No user authentication method data collected." -ForegroundColor Yellow
    }
}
else {
    Write-Host "User authentication methods export skipped (not requested)." -ForegroundColor Yellow
}

if ($retrieveGroups) {
    if ($groupsResults.Count -gt 0) {
        $groupExportPath = Join-Path $OutputDirectory "AllTenants-Groups.csv"
        $groupsResults | Export-Csv -Path $groupExportPath -NoTypeInformation
        Write-Host "Group export written to $groupExportPath" -ForegroundColor Green
    }
    else {
        Write-Host "No group data collected." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Group export skipped (not requested)." -ForegroundColor Yellow
}

if ($retrieveSet.Contains("ServicePrincipals") -or $retrieveSet.Count -eq 0) {
    if ($servicePrincipalResults.Count -gt 0) {
        $servicePrincipalExportPath = Join-Path $OutputDirectory "AllTenants-ServicePrincipals.csv"
        $servicePrincipalResults | Export-Csv -Path $servicePrincipalExportPath -NoTypeInformation
        Write-Host "Service principal export written to $servicePrincipalExportPath" -ForegroundColor Green
    }
    else {
        Write-Host "No service principal data collected." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Service principal export skipped (not requested)." -ForegroundColor Yellow
}

if ($retrieveAppPermissions) {
    if ($appPermissionResults.Count -gt 0) {
        $appPermissionsExportPath = Join-Path $OutputDirectory "AllTenants-ServicePrincipalAppPermissions.csv"
        $appPermissionResults | Export-Csv -Path $appPermissionsExportPath -NoTypeInformation
        Write-Host "Service principal application permissions export written to $appPermissionsExportPath" -ForegroundColor Green
    }
    else {
        Write-Host "No service principal application permissions collected." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Service principal application permissions export skipped (not requested)." -ForegroundColor Yellow
}

if ($retrieveRoles) {
    if ($roleResults.Count -gt 0) {
        $roleExportPath = Join-Path $OutputDirectory "AllTenants-Roles.csv"
        $roleResults | Export-Csv -Path $roleExportPath -NoTypeInformation
        Write-Host "Role export written to $roleExportPath" -ForegroundColor Green
    }
    else {
        Write-Host "No role assignment data collected." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Role export skipped (not requested)." -ForegroundColor Yellow
}

if ($retrieveApplications) {
    if ($applicationOwnerResults.Count -gt 0) {
        $applicationExportPath = Join-Path $OutputDirectory "AllTenants-AppRegistrations.csv"
        $applicationOwnerResults | Export-Csv -Path $applicationExportPath -NoTypeInformation
        Write-Host "Application registration export written to $applicationExportPath" -ForegroundColor Green
    }
    else {
        Write-Host "No application registration data collected." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Application registration export skipped (not requested)." -ForegroundColor Yellow
}
