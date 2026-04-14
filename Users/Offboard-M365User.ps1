<#
.SYNOPSIS
    Offboards a user from Microsoft 365.

.DESCRIPTION
    Performs all standard M365 offboarding steps for a departing user.
    Revokes active sessions, removes licenses, sets mailbox auto-reply,
    blocks sign-in, and optionally grants manager access to the mailbox.

.PARAMETER UserPrincipalName
    The UPN of the user to offboard. Example: jsmith@domain.com

.PARAMETER ManagerUPN
    Optional UPN of the manager to grant mailbox access to.

.PARAMETER AutoReplyMessage
    Optional custom auto-reply message. A default message is used if not provided.

.PARAMETER ExportPath
    Optional path to export the offboarding report as a CSV file.

.EXAMPLE
    .\Offboard-M365User.ps1 -UserPrincipalName "jsmith@domain.com"
    .\Offboard-M365User.ps1 -UserPrincipalName "jsmith@domain.com" -ManagerUPN "mjones@domain.com"

.NOTES
    Requires Microsoft Graph PowerShell SDK and Exchange Online module.
    Install modules:
        Install-Module Microsoft.Graph -Scope CurrentUser
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [string]$ManagerUPN = "",
    [string]$AutoReplyMessage = "",
    [string]$ExportPath = ""
)

$results = @()

function Write-ActionLog {
    param([string]$Action, [string]$Status, [string]$Notes = "")
    $results += [PSCustomObject]@{
        Action  = $Action
        Status  = $Status
        Notes   = $Notes
    }
    $color = if ($Status -like "*✅*") { "Green" } elseif ($Status -like "*❌*") { "Red" } else { "Yellow" }
    Write-Host "$Status — $Action" -ForegroundColor $color
}

# --- Connect ---
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All" -ErrorAction Stop
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect — $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Connected to Exchange Online" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Exchange Online — $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nStarting offboarding for: $UserPrincipalName`n" -ForegroundColor Cyan

# --- Get User ---
try {
    $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
} catch {
    Write-Host "User not found: $UserPrincipalName" -ForegroundColor Red
    exit 1
}

# --- Step 1: Revoke Sessions ---
try {
    Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop
    Write-ActionLog "Revoke active sessions" "✅ Done"
} catch {
    Write-ActionLog "Revoke active sessions" "❌ Failed" $_.Exception.Message
}

# --- Step 2: Block Sign-In ---
try {
    Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop
    Write-ActionLog "Block sign-in" "✅ Done"
} catch {
    Write-ActionLog "Block sign-in" "❌ Failed" $_.Exception.Message
}

# --- Step 3: Remove Licenses ---
try {
    $licenses = Get-MgUserLicenseDetail -UserId $user.Id
    if ($licenses) {
        $skuIds = $licenses | ForEach-Object { $_.SkuId }
        Set-MgUserLicense -UserId $user.Id -RemoveLicenses $skuIds -AddLicenses @() -ErrorAction Stop
        Write-ActionLog "Remove M365 licenses" "✅ Done" "$($licenses.Count) license(s) removed"
    } else {
        Write-ActionLog "Remove M365 licenses" "⚠️ Skipped" "No licenses found"
    }
} catch {
    Write-ActionLog "Remove M365 licenses" "❌ Failed" $_.Exception.Message
}

# --- Step 4: Set Auto Reply ---
try {
    $message = if ($AutoReplyMessage -ne "") {
        $AutoReplyMessage
    } else {
        "$($user.DisplayName) is no longer with the organization. Please contact your account manager for assistance."
    }

    Set-MailboxAutoReplyConfiguration `
        -Identity $UserPrincipalName `
        -AutoReplyState Enabled `
        -InternalMessage $message `
        -ExternalMessage $message `
        -ErrorAction Stop

    Write-ActionLog "Set mailbox auto-reply" "✅ Done"
} catch {
    Write-ActionLog "Set mailbox auto-reply" "❌ Failed" $_.Exception.Message
}

# --- Step 5: Grant Manager Mailbox Access ---
if ($ManagerUPN -ne "") {
    try {
        Add-MailboxPermission `
            -Identity $UserPrincipalName `
            -User $ManagerUPN `
            -AccessRights FullAccess `
            -InheritanceType All `
            -ErrorAction Stop

        Write-ActionLog "Grant manager mailbox access" "✅ Done" "Access granted to $ManagerUPN"
    } catch {
        Write-ActionLog "Grant manager mailbox access" "❌ Failed" $_.Exception.Message
    }
} else {
    Write-ActionLog "Grant manager mailbox access" "⚠️ Skipped" "No manager UPN provided"
}

# --- Summary ---
Write-Host "`n--- Offboarding Summary for $UserPrincipalName ---" -ForegroundColor Cyan
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
Disconnect-ExchangeOnline -Confirm:$false