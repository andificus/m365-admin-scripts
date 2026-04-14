<#
.SYNOPSIS
    Reports on MFA registration status for all users in the M365 tenant.

.DESCRIPTION
    Connects to Microsoft Graph and pulls MFA registration status for all
    users. Flags accounts with MFA not registered. Useful for security
    audits and compliance reporting.

.PARAMETER ExportPath
    Optional path to export results as a CSV file.

.EXAMPLE
    .\Get-MFAStatusReport.ps1
    .\Get-MFAStatusReport.ps1 -ExportPath "C:\Reports\mfa-status.csv"

.NOTES
    Requires Microsoft Graph PowerShell SDK.
    Requires UserAuthenticationMethod.Read.All permission.
    Install module: Install-Module Microsoft.Graph -Scope CurrentUser
#>

param (
    [string]$ExportPath = ""
)

# --- Connect to Graph ---
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "UserAuthenticationMethod.Read.All", "User.Read.All" -ErrorAction Stop
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Microsoft Graph — $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --- Get All Users ---
Write-Host "Retrieving users..." -ForegroundColor Cyan
$users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled

$results = @()
$i = 0

foreach ($user in $users) {
    $i++
    Write-Progress -Activity "Checking MFA status" -Status "$i of $($users.Count) — $($user.UserPrincipalName)" -PercentComplete (($i / $users.Count) * 100)

    try {
        $methods = Get-MgUserAuthenticationMethod -UserId $user.Id

        # Filter out default password method to find real MFA methods
        $mfaMethods = $methods | Where-Object {
            $_.AdditionalProperties["@odata.type"] -ne "#microsoft.graph.passwordAuthenticationMethod"
        }

        $mfaStatus = if ($mfaMethods.Count -gt 0) { "✅ Registered" } else { "❌ Not Registered" }
        $methodList = $mfaMethods | ForEach-Object { $_.AdditionalProperties["@odata.type"] -replace "#microsoft.graph.", "" }

        $results += [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            MFAStatus         = $mfaStatus
            Methods           = ($methodList -join ", ")
        }
    } catch {
        $results += [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            MFAStatus         = "⚠️ Error retrieving"
            Methods           = $_.Exception.Message
        }
    }
}

Write-Progress -Completed -Activity "Checking MFA status"

# --- Summary ---
$registered    = ($results | Where-Object { $_.MFAStatus -like "*Registered*" }).Count
$notRegistered = ($results | Where-Object { $_.MFAStatus -like "*Not Registered*" }).Count

Write-Host "`n--- MFA Status Summary ---" -ForegroundColor Cyan
Write-Host "✅ Registered:     $registered" -ForegroundColor Green
Write-Host "❌ Not Registered: $notRegistered" -ForegroundColor Red

$results | Format-Table -AutoSize

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