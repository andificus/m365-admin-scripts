<#
.SYNOPSIS
    Finds users who have not signed in within a specified number of days.

.DESCRIPTION
    Connects to Microsoft Graph and reports on users whose last sign-in
    is older than the defined threshold. Useful for identifying stale
    accounts that should be reviewed or disabled.

.PARAMETER DaysThreshold
    Number of days since last sign-in before flagging a user. Default is 30.

.PARAMETER ExportPath
    Optional path to export results as a CSV file.

.EXAMPLE
    .\Get-InactiveUsers.ps1
    .\Get-InactiveUsers.ps1 -DaysThreshold 60
    .\Get-InactiveUsers.ps1 -DaysThreshold 30 -ExportPath "C:\Reports\inactive-users.csv"

.NOTES
    Requires Microsoft Graph PowerShell SDK.
    Requires AuditLog.Read.All and User.Read.All permissions.
    Install module: Install-Module Microsoft.Graph -Scope CurrentUser
#>

param (
    [int]$DaysThreshold = 30,
    [string]$ExportPath = ""
)

# --- Connect ---
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "AuditLog.Read.All", "User.Read.All" -ErrorAction Stop
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect — $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --- Get Users ---
Write-Host "Retrieving users..." -ForegroundColor Cyan
$users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, SignInActivity

$cutoffDate = (Get-Date).AddDays(-$DaysThreshold)
$results = @()

foreach ($user in $users) {
    $lastSignIn = $user.SignInActivity.LastSignInDateTime

    if ($null -eq $lastSignIn) {
        $status   = "⚠️ Never Signed In"
        $daysAgo  = "N/A"
    } elseif ($lastSignIn -lt $cutoffDate) {
        $daysAgo  = (New-TimeSpan -Start $lastSignIn -End (Get-Date)).Days
        $status   = "❌ Inactive"
    } else {
        $daysAgo  = (New-TimeSpan -Start $lastSignIn -End (Get-Date)).Days
        $status   = "✅ Active"
    }

    $results += [PSCustomObject]@{
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        AccountEnabled    = $user.AccountEnabled
        LastSignIn        = if ($lastSignIn) { $lastSignIn.ToString("yyyy-MM-dd") } else { "Never" }
        DaysSinceSignIn   = $daysAgo
        Status            = $status
    }
}

# --- Summary ---
$inactive     = ($results | Where-Object { $_.Status -like "*Inactive*" }).Count
$neverSignedIn = ($results | Where-Object { $_.Status -like "*Never*" }).Count
$active       = ($results | Where-Object { $_.Status -like "*Active*" }).Count

Write-Host "`n--- Inactive User Summary (Threshold: $DaysThreshold days) ---" -ForegroundColor Cyan
Write-Host "✅ Active:          $active" -ForegroundColor Green
Write-Host "❌ Inactive:        $inactive" -ForegroundColor Red
Write-Host "⚠️  Never Signed In: $neverSignedIn" -ForegroundColor Yellow

$results | Where-Object { $_.Status -notlike "*Active*" } | Sort-Object DaysSinceSignIn -Descending | Format-Table -AutoSize

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