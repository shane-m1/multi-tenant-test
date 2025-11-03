[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $true)]
    [string[]]$TenantIds,

    [string]$TestSuite = "Baseline",

[string]$MgProfile = "beta"
)

if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
    Write-Error "The Microsoft Graph PowerShell SDK is required. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
    exit 1
}

if (-not (Get-Command -Name Invoke-Maester -ErrorAction SilentlyContinue)) {
    Write-Error "The Maester module is required. Install it with: Install-Module Maester -Scope CurrentUser"
    exit 1
}

$secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$appCredential = New-Object System.Management.Automation.PSCredential ($ClientId, $secureSecret)

foreach ($tenantId in $TenantIds) {
    Write-Host "Processing tenant $tenantId..." -ForegroundColor Cyan

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

        Connect-MgGraph -TenantId $tenantId `
                        -ClientSecretCredential $appCredential `
                        -NoWelcome `
                        -ErrorAction Stop | Out-Null

        if ($MgProfile) {
            # Select-MgProfile -Name $MgProfile -ErrorAction Stop
        }

        Write-Host "Running Maester test suite '$TestSuite'..." -ForegroundColor Yellow

        $result = Invoke-Maester #-TestSuite $TestSuite -ErrorAction Stop

        if ($result -and $result.Summary) {
            Write-Host ("Pass: {0} | Fail: {1} | Inconclusive: {2}" -f `
                $result.Summary.Passed, `
                $result.Summary.Failed, `
                $result.Summary.Inconclusive) -ForegroundColor Green
        }
        else {
            Write-Host "Maester test completed. Review output above for details." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Failed to process tenant $tenantId. $($_.Exception.Message)"
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
