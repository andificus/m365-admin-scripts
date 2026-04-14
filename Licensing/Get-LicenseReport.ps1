<#
.SYNOPSIS
    Reports on all assigned Microsoft 365 licenses across the tenant.

.DESCRIPTION
    Connects to Microsoft Graph and pulls a full license assignment report
    showing which licenses are assigned, how many are consumed, and how
    many are remaining. Useful for license audits and cost management.

.PARAMETER ExportPath
    Optional path to export results as a CSV file.

.EXAMPLE
    .\Get-LicenseReport.ps1
    .\Get-LicenseReport.ps1 -ExportPath "C:\Reports\licenses.csv"

.NOTES
    Requires Microsoft Graph PowerShell SDK.
    Requires Organization.Read.All permission.
    Install module: Install-Module Microsoft.Graph -Scope CurrentUser
#>

param (
    [string]$ExportPath = ""
)

# --- Connect ---
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "Organization.Read.All" -ErrorAction Stop
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect — $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --- Get License Data ---
Write-Host "Retrieving license information..." -ForegroundColor Cyan

try {
    $subscribedSkus = Get-MgSubscribedSku -ErrorAction Stop
} catch {
    Write-Host "Failed to retrieve licenses — $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

$results = @()

foreach ($sku in $subscribedSkus) {
    $total     = $sku.PrepaidUnits.Enabled
    $consumed  = $sku.ConsumedUnits
    $remaining = $total - $consumed
    $status    = if ($remaining -le 0) { "❌ None Remaining" } elseif ($remaining -le 5) { "⚠️ Low" } else { "✅ OK" }

    $results += [PSCustomObject]@{
        LicenseName   = $sku.SkuPartNumber
        Total         = $total
        Consumed      = $consumed
        Remaining     = $remaining
        Status        = $status
    }
}

# --- Summary ---
Write-Host "`n--- License Report ---" -ForegroundColor Cyan
$results | Sort-Object Remaining | Format-Table -AutoSize

$lowLicenses = $results | Where-Object { $_.Status -notlike "*OK*" }
if ($lowLicenses) {
    Write-Host "`n⚠️  The following licenses need attention:" -ForegroundColor Yellow
    $lowLicenses | Format-Table -AutoSize
} else {
    Write-Host "`n✅ All licenses have sufficient availability." -ForegroundColor Green
}

# --- Export if path provided ---
if ($ExportPath -ne "") {
    try {
        $results | Export-Csv -Path $ExportPath -NoTypeInformation
        Write-Host "`n✅ Report exported to: $ExportPath" -ForegroundColor Green
    } catch {
        Write-Host "`n❌ Failed to export — $($_.Exception.Message)" -ForegroundColor Red
    }
}

Disconnect-MgGraph