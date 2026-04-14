<#
.SYNOPSIS
    Reports on mailbox sizes across the Microsoft 365 tenant.

.DESCRIPTION
    Connects to Exchange Online and pulls mailbox size information for all
    user mailboxes. Useful for storage audits, identifying large mailboxes,
    and planning migrations.

.PARAMETER ExportPath
    Optional path to export results as a CSV file.

.PARAMETER TopCount
    Show the top X largest mailboxes in the summary. Default is 10.

.EXAMPLE
    .\Get-MailboxSizeReport.ps1
    .\Get-MailboxSizeReport.ps1 -TopCount 20
    .\Get-MailboxSizeReport.ps1 -ExportPath "C:\Reports\mailbox-sizes.csv"

.NOTES
    Requires Exchange Online PowerShell module.
    Install module: Install-Module ExchangeOnlineManagement -Scope CurrentUser
#>

param (
    [string]$ExportPath = "",
    [int]$TopCount = 10
)

# --- Connect ---
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Connected to Exchange Online" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect — $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --- Get Mailboxes ---
Write-Host "Retrieving mailboxes..." -ForegroundColor Cyan
$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox

$results = @()
$i = 0

foreach ($mailbox in $mailboxes) {
    $i++
    Write-Progress -Activity "Retrieving mailbox sizes" -Status "$i of $($mailboxes.Count) — $($mailbox.UserPrincipalName)" -PercentComplete (($i / $mailboxes.Count) * 100)

    try {
        $stats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName -ErrorAction Stop

        # Parse size string to MB
        $sizeString = $stats.TotalItemSize.ToString()
        $sizeMB = if ($sizeString -match "([\d,]+)\s*bytes") {
            [math]::Round([int64]($matches[1] -replace ",", "") / 1MB, 2)
        } else {
            0
        }

        $results += [PSCustomObject]@{
            DisplayName       = $mailbox.DisplayName
            UserPrincipalName = $mailbox.UserPrincipalName
            ItemCount         = $stats.ItemCount
            "Size (MB)"       = $sizeMB
            LastLogon         = if ($stats.LastLogonTime) { $stats.LastLogonTime.ToString("yyyy-MM-dd") } else { "Never" }
        }
    } catch {
        $results += [PSCustomObject]@{
            DisplayName       = $mailbox.DisplayName
            UserPrincipalName = $mailbox.UserPrincipalName
            ItemCount         = "N/A"
            "Size (MB)"       = "N/A"
            LastLogon         = "N/A"
        }
    }
}

Write-Progress -Completed -Activity "Retrieving mailbox sizes"

# --- Summary ---
$totalSizeMB = ($results | Where-Object { $_."Size (MB)" -ne "N/A" } | Measure-Object -Property "Size (MB)" -Sum).Sum
$totalSizeGB = [math]::Round($totalSizeMB / 1024, 2)

Write-Host "`n--- Mailbox Size Summary ---" -ForegroundColor Cyan
Write-Host "Total Mailboxes : $($results.Count)"
Write-Host "Total Size      : $totalSizeGB GB"

Write-Host "`n--- Top $TopCount Largest Mailboxes ---" -ForegroundColor Cyan
$results | Sort-Object "Size (MB)" -Descending | Select-Object -First $TopCount | Format-Table -AutoSize

# --- Export if path provided ---
if ($ExportPath -ne "") {
    try {
        $results | Sort-Object "Size (MB)" -Descending | Export-Csv -Path $ExportPath -NoTypeInformation
        Write-Host "`n✅ Report exported to: $ExportPath" -ForegroundColor Green
    } catch {
        Write-Host "`n❌ Failed to export — $($_.Exception.Message)" -ForegroundColor Red
    }
}

Disconnect-ExchangeOnline -Confirm:$false