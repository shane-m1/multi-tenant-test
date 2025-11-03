# Multi-Tenant Testing Project Guide

This document provides essential guidance for AI agents working with this codebase.

## Project Overview

This project provides tools for managing and testing multi-tenant Azure AD (Microsoft Entra ID) environments. It focuses on:
- Exporting directory data across multiple tenants
- Running automated compliance tests using Maester
- Managing tenant configurations through CSV-based inputs

## Key Components

### Directory Export (`scripts/Export-TenantDirectory.ps1`)
- Exports users, groups, and role assignments across multiple tenants
- Uses Microsoft Graph PowerShell SDK for tenant access
- Generates CSV outputs in the `exports` directory
- Handles both active and eligible role assignments

### Test Runner (`scripts/Run-MaesterTests.ps1`)
- Executes Maester test suites across multiple tenants
- Validates tenant configurations and compliance
- Reports test results with pass/fail summaries

### Configuration
- Tenant information is managed in `tenants.csv`
- Required format: CSV with `TenantId` column (accepts both GUIDs and domain names)
- Exports are stored in `./tenant-exports` by default

## Development Workflows

### Prerequisites
```powershell
# Install required PowerShell modules
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module Maester -Scope CurrentUser
```

### Running Directory Exports
```powershell
./scripts/Export-TenantDirectory.ps1 `
    -ClientId "<app-id>" `
    -ClientSecret "<secret>" `
    -TenantCsvPath "./tenants.csv"
```

### Running Compliance Tests
```powershell
./scripts/Run-MaesterTests.ps1 `
    -ClientId "<app-id>" `
    -ClientSecret "<secret>" `
    -TenantIds "tenant-id-1","tenant-id-2" `
    -TestSuite "Baseline"
```

## Key Patterns

1. **Authentication**
   - Uses client credentials flow (service principal)
   - Supports both standard and Azure.Identity authentication methods
   - Always disconnects from Microsoft Graph between tenant switches

2. **Error Handling**
   - Continues processing remaining tenants if one fails
   - Preserves collected data even if some operations fail
   - Uses warning messages for non-fatal errors

3. **Data Collection**
   - Implements paging for large data sets via `Get-GraphPagedResult`
   - Caches principal information to minimize API calls
   - Aggregates data across all tenants before export

## Integration Points

1. **Microsoft Graph API**
   - Uses beta API profile for advanced features
   - Required permissions: Directory.Read.All, RoleManagement.Read.All
   - Accessed via Microsoft.Graph PowerShell modules

2. **Maester Integration**
   - Compliance testing framework integration
   - Test suites defined externally to this project
   - Results captured and reported per tenant