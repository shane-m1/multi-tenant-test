[CmdletBinding()]
param(
    [string[]]$TenantIds,
    [string]$TenantCsvPath,
    [string]$TenantIdColumn = "TenantId",
    [string]$OutputPath
)

$scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
if (-not $scriptDirectory) {
    $scriptDirectory = Get-Location
}
$envFilePath = Join-Path $scriptDirectory ".env"
$knownOptionKeys = @("TenantCsvPath", "TenantIdColumn", "OutputPath")

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

if (-not $TenantCsvPath -and $envData.ContainsKey("TenantCsvPath")) {
    $TenantCsvPath = $envData["TenantCsvPath"]
}

if (-not $PSBoundParameters.ContainsKey("TenantIdColumn") -and $envData.ContainsKey("TenantIdColumn")) {
    $TenantIdColumn = $envData["TenantIdColumn"]
}

if (-not $PSBoundParameters.ContainsKey("OutputPath") -and $envData.ContainsKey("OutputPath")) {
    $OutputPath = $envData["OutputPath"]
}

$missing = @()
if (-not $TenantIds -and -not $TenantCsvPath) {
    $missing += "TenantIds or TenantCsvPath"
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
            $envContent += "$key=$($persistData[$key])"
        }
    }

    Set-Content -Path $envFilePath -Value $envContent -Encoding UTF8
}

$tenantIdList = @()
if ($TenantIds) {
    foreach ($tenant in $TenantIds) {
        if (-not [string]::IsNullOrWhiteSpace($tenant)) {
            $value = $tenant.Trim()
            if ($tenantIdList -notcontains $value) {
                $tenantIdList += $value
            }
        }
    }
}

if (-not $tenantIdList -and $TenantCsvPath) {
    if (-not (Test-Path -Path $TenantCsvPath)) {
        throw "CSV file not found at $TenantCsvPath"
    }

    $csvTenants = Import-Csv -Path $TenantCsvPath
    if (-not $csvTenants) {
        throw "No tenant entries found in CSV."
    }

    if (-not $csvTenants[0].PSObject.Properties.Match($TenantIdColumn)) {
        throw "Column '$TenantIdColumn' not present in CSV."
    }

    foreach ($row in $csvTenants) {
        $value = $row.$TenantIdColumn
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $trimmed = $value.Trim()
            if ($tenantIdList -notcontains $trimmed) {
                $tenantIdList += $trimmed
            }
        }
    }
}

if (-not $tenantIdList) {
    throw "No tenant identifiers were supplied."
}

if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
    Write-Error "The Microsoft Graph PowerShell SDK is required. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

$connectParams = @{
    Scopes      = @("CrossTenantInformation.ReadBasic.All")
    NoWelcome   = $true
    ErrorAction = "Stop"
}

Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Yellow
Connect-MgGraph @connectParams | Out-Null

function Find-TenantInformationById {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    return Find-MgTenantRelationshipTenantInformationByTenantId -TenantId $TenantId -ErrorAction Stop
}

$results = @()
foreach ($tenantId in $tenantIdList) {
    Write-Host "Resolving tenant $tenantId..." -ForegroundColor Cyan

    try {
        $tenantInfo = Find-TenantInformationById -TenantId $tenantId

        $result = [pscustomobject]@{
            TenantId        = $tenantId
            DisplayName     = $tenantInfo.DisplayName
            DefaultDomain   = $tenantInfo.DefaultDomainName
            InitialDomain   = $null
            VerifiedDomains = $null
            FederationBrand = $tenantInfo.FederationBrandName
            TenantType      = $tenantInfo.TenantType
        }

        $results += $result
        Write-Host "Resolved $tenantId -> $($result.DefaultDomain) ($($result.DisplayName))" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to resolve tenant $tenantId. $($_.Exception.Message)"
        $results += [pscustomobject]@{
            TenantId        = $tenantId
            DisplayName     = $null
            DefaultDomain   = $null
            InitialDomain   = $null
            VerifiedDomains = $null
            FederationBrand = $null
            TenantType      = $null
            Error           = $_.Exception.Message
        }
    }
}

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

if ($OutputPath) {
    $resolvedOutput = Resolve-Path -LiteralPath $OutputPath -ErrorAction SilentlyContinue
    if ($resolvedOutput) {
        $destinationPath = $resolvedOutput.Path
        if (Test-Path -Path $destinationPath -PathType Container) {
            $destinationPath = Join-Path -Path $destinationPath -ChildPath "TenantIdentity.csv"
        }
    }
    else {
        $destinationPath = $OutputPath
        $outputDirectory = Split-Path -Path $destinationPath -Parent
        if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }
    }

    $exportDirectory = Split-Path -Path $destinationPath -Parent
    if ($exportDirectory -and -not (Test-Path -Path $exportDirectory)) {
        New-Item -ItemType Directory -Path $exportDirectory -Force | Out-Null
    }

    $results | Export-Csv -Path $destinationPath -NoTypeInformation -Encoding UTF8
    Write-Host "Saved results to $destinationPath" -ForegroundColor Yellow
}

return $results
