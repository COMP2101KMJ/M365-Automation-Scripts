<#
.SYNOPSIS
    Microsoft 365 User and License Report

.DESCRIPTION
    Connects to Microsoft Graph and exports Microsoft 365 user account
    and license information to a CSV report.

    This script demonstrates:
      - Microsoft Graph PowerShell authentication
      - Entra ID / Microsoft 365 user administration
      - License reporting
      - PowerShell automation
      - Error handling
      - CSV reporting

.NOTES
    Author: Kevin Mathews Jose
    Portfolio: GitHub
    Version: 1.0

.REQUIREMENTS
    Microsoft.Graph PowerShell module

    Install with:
    Install-Module Microsoft.Graph -Scope CurrentUser

.SECURITY
    This sample contains no tenant IDs, passwords, API secrets,
    company names, or other confidential information.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\M365-User-License-Report.csv"
)

$ErrorActionPreference = "Stop"

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message"
}

try {
    Write-Log "Checking Microsoft Graph PowerShell module..."

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft Graph PowerShell is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    Write-Log "Connecting to Microsoft Graph..."

    Connect-MgGraph `
        -Scopes "User.Read.All", "Directory.Read.All" `
        -NoWelcome

    Write-Log "Connected successfully." -Level "SUCCESS"

    Write-Log "Retrieving Microsoft 365 users..."

    $users = Get-MgUser -All -Property `
        "Id,DisplayName,UserPrincipalName,AccountEnabled,Department,JobTitle,CreatedDateTime,AssignedLicenses"

    if (-not $users) {
        Write-Log "No users were returned from Microsoft Graph." -Level "WARNING"
        return
    }

    Write-Log "Retrieving license SKU information..."

    $skus = Get-MgSubscribedSku -All

    $skuLookup = @{}

    foreach ($sku in $skus) {
        $skuLookup[$sku.SkuId.ToString()] = $sku.SkuPartNumber
    }

    Write-Log "Building report for $($users.Count) users..."

    $report = foreach ($user in $users) {

        $licenseNames = foreach ($license in $user.AssignedLicenses) {

            $skuId = $license.SkuId.ToString()

            if ($skuLookup.ContainsKey($skuId)) {
                $skuLookup[$skuId]
            }
            else {
                $skuId
            }
        }

        if (-not $licenseNames) {
            $licenseNames = "Unlicensed"
        }

        [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            Department        = $user.Department
            JobTitle          = $user.JobTitle
            License           = ($licenseNames -join "; ")
            CreatedDate       = $user.CreatedDateTime
        }
    }

    $report = $report | Sort-Object DisplayName

    $report | Export-Csv `
        -Path $OutputPath `
        -NoTypeInformation `
        -Encoding UTF8

    Write-Log "Report generated successfully." -Level "SUCCESS"
    Write-Log "Users processed: $($report.Count)"
    Write-Log "Report saved to: $OutputPath"

    Write-Host ""
    Write-Host "Report Summary"
    Write-Host "--------------"
    Write-Host "Total Users : $($report.Count)"
    Write-Host "Enabled     : $(($report | Where-Object AccountEnabled -eq $true).Count)"
    Write-Host "Disabled    : $(($report | Where-Object AccountEnabled -eq $false).Count)"
    Write-Host "Unlicensed  : $(($report | Where-Object License -eq 'Unlicensed').Count)"
}
catch {
    Write-Log $_.Exception.Message -Level "ERROR"
    exit 1
}
finally {
    try {
        Disconnect-MgGraph | Out-Null
        Write-Log "Disconnected from Microsoft Graph."
    }
    catch {
        # No active Graph session.
    }
}
