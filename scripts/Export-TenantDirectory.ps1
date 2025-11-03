[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $true)]
    [string]$TenantCsvPath,

    [string]$OutputDirectory = "./tenant-exports",

    [string]$TenantIdColumn = "TenantId"
)

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

        Write-Host "Retrieving users..." -ForegroundColor Yellow
        $tenantUsers = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Mail, AccountEnabled |
            Select-Object @{Name = "TenantId"; Expression = { $tenantId }},
                          Id,
                          DisplayName,
                          UserPrincipalName,
                          Mail,
                          AccountEnabled

        $usersResults += $tenantUsers

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
            $roleDefinitionData = Get-GraphPagedResult -Uri "https://graph.microsoft.com/beta/roleManagement/directory/roleDefinitions`?\$select=id,displayName"

            foreach ($definition in $roleDefinitionData) {
                if ($definition.id -and -not $roleDefinitions.ContainsKey($definition.id)) {
                    $roleDefinitions[$definition.id] = $definition.displayName
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

        Write-Host "Tenant $tenantId complete." -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to process tenant $tenantId. $($_.Exception.Message)"
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

if ($usersResults.Count -gt 0) {
    $userExportPath = Join-Path $OutputDirectory "AllTenants-Users.csv"
    $usersResults | Export-Csv -Path $userExportPath -NoTypeInformation
    Write-Host "User export written to $userExportPath" -ForegroundColor Green
}
else {
    Write-Host "No user data collected." -ForegroundColor Yellow
}

if ($groupsResults.Count -gt 0) {
    $groupExportPath = Join-Path $OutputDirectory "AllTenants-Groups.csv"
    $groupsResults | Export-Csv -Path $groupExportPath -NoTypeInformation
    Write-Host "Group export written to $groupExportPath" -ForegroundColor Green
}
else {
    Write-Host "No group data collected." -ForegroundColor Yellow
}

if ($roleResults.Count -gt 0) {
    $roleExportPath = Join-Path $OutputDirectory "AllTenants-Roles.csv"
    $roleResults | Export-Csv -Path $roleExportPath -NoTypeInformation
    Write-Host "Role export written to $roleExportPath" -ForegroundColor Green
}
else {
    Write-Host "No role assignment data collected." -ForegroundColor Yellow
}
