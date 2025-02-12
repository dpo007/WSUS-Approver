<#
.SYNOPSIS
    Automates the approval and declination of updates in a WSUS (Windows Server Update Services) server.

.DESCRIPTION
    This script connects to a WSUS server and processes updates based on specified parameters.
    It can approve or decline updates, handle license agreements, and log actions to a file.

.PARAMETER WsusServer
    The WSUS server to connect to (default: 'localhost').

.PARAMETER Port
    The port to use for the WSUS server (default: 8530).

.PARAMETER UseSSL
    Use SSL to connect to the WSUS server.

.PARAMETER NoSync
    Do not synchronize updates before processing.

.PARAMETER Reset
    Reset the update list before processing.

.PARAMETER DryRun
    Perform a dry run without making any changes.

.PARAMETER DeclineOnly
    Only decline updates, do not approve any.

.PARAMETER IncludeUpgrades
    Include upgrade classifications in the approval process.

.PARAMETER DeclineIA64
    Decline IA64 updates (default: $true).

.PARAMETER DeclineARM64
    Decline ARM64 updates (default: $true).

.PARAMETER DeclineX86
    Decline x86 updates (default: $true).

.PARAMETER DeclineX64
    Decline x64 updates (default: $false).

.PARAMETER DeclinePreview
    Decline preview updates (default: $true).

.PARAMETER DeclineBeta
    Decline beta updates (default: $true).

.PARAMETER RestrictToLanguages
    Only keep updates for specified languages/locales (default: @('en-us', 'en-gb')).

.NOTES
    The script logs its actions to a log file located in the 'logs' directory under the script's root directory.
    Superseded and expired updates are declined after new updates are approved.

.EXAMPLE
    .\WSUS-Approver.ps1 -WsusServer 'wsus-server' -Port 8531 -UseSSL -DryRun

    Connects to the specified WSUS server using SSL on port 8531 and performs a dry run of the update approval process.
#>
param (
    [string]$WsusServer = 'localhost',
    [int]$Port = 8530,
    [switch]$UseSSL,
    [switch]$NoSync,
    [switch]$Reset,
    [switch]$DryRun,
    [switch]$DeclineOnly,
    [switch]$IncludeUpgrades,
    [bool]$DeclineIA64 = $true,
    [bool]$DeclineARM64 = $true,
    [bool]$DeclineX86 = $true,
    [bool]$DeclineX64 = $false,
    [bool]$DeclinePreview = $true,
    [bool]$DeclineBeta = $true,
    [string[]]$RestrictToLanguages = @('en-us', 'en-gb') # Only these languages/locales will be kept
)

# Debug
$DeclineOnly = $true
$DryRun = $true
$NoSync = $true

# Set default error action to Stop
$ErrorActionPreference = 'Stop'

#region Initialize values
$logFile = ('{0}\logs\wsus-approver_{1}.Log' -f $PSScriptRoot, (Get-Date).ToString('dddd_htt'))

# Do not add upgrades here. They are currently handled manually for more control
$approve_classifications = @(
    'Critical Updates',
    'Definition Updates',
    'Drivers',
    'Feature Packs',
    'Security Updates',
    'Service Packs',
    'Tools',
    'Update Rollups',
    'Updates'
)

if ($IncludeUpgrades) {
    $approve_classifications += 'Upgrades'
}

# All Computers group (default group for all computers in WSUS)
$approve_group = 'All Computers'

# Get all locales (languages) from the system
[string[]]$allLocales = [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::AllCultures) | Where-Object { $_.Name -match "-" } | Select-Object -ExpandProperty Name
#endregion Initialize values

#region Function Definitions
function Log ($text) {
    Write-Output "$(Get-Date -Format s): $text" | Tee-Object -Append $logFile
}

function is_selected ($update) {
    #
    # Is there any way to check update against language list in here?
    #
    if ($update.UpdateClassificationTitle -in $update_classifications.Title) {
        return $true
    }

    foreach ($product in $update.ProductTitles) {
        if ($product -in $update_categories.Title) {
            return $true
        }
    }

    return $false
}

function TestUpdateTitleLanguageMatch {
    param (
        [string]$Title,
        [string[]]$AllLocales,
        [string[]]$RestrictToLanguages
    )

    # Check if title contains any entry from $AllLocales
    $matchesLocale = $Title -imatch ($AllLocales -join "|")

    # Check if title also contains any entry from $RestrictToLanguages
    $matchesRestricted = $Title -imatch ($RestrictToLanguages -join "|")

    # Return true if the title matches any locale but does not match any restricted language
    return $matchesLocale -and -not $matchesRestricted
}
#endregion Function Definitions

#################
# Main Entry Point
###################

# Ensure Log folder exists
if (-not (Test-Path ('{0}\logs' -f $PSScriptRoot))) {
    New-Item -ItemType Directory -Path ('{0}\logs' -f $PSScriptRoot) | Out-Null
}

# Reset Log file
if (Test-Path $logFile) {
    Remove-Item $logFile
}

# Load the WSUS assembly
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WsusServer, $UseSSL, $Port)

# Get the groups, subscription, categories and classifications
$group = $wsus.GetComputerTargetGroups() | Where-Object { $_.Name -eq $approve_group }
$subscription = $wsus.GetSubscription()
$update_categories = $subscription.GetUpdateCategories()
$update_classifications = $subscription.GetUpdateClassifications()

# Start synchronization (if not already running)
if (-not $NoSync) {
    if ($subscription.GetSynchronizationStatus() -eq "NotProcessing") {
        Log "Starting synchronization..."
        $subscription.StartSynchronization()
    }
}

# Wait for any currently running synchronization jobs to finish before continuing
while ($subscription.GetSynchronizationStatus() -ne 'NotProcessing') {
    Log "Waiting for synchronization to finish..."
    Start-Sleep -s 10
}

# Start by removing deselected updates as there is no need to do further processing on them
Log "Checking for deselected updates"
$wsus.GetUpdates() | ForEach-Object {
    if (-not (is_selected $_)) {
        Log "Deleting deselected update: $($_.Title)"
        if (-not $DryRun) { $wsus.DeleteUpdate($_.Id.UpdateId) }
    }
}

if ($Reset) {
    $updates = $wsus.GetUpdates()
}
else {
    $updates = $wsus.GetUpdates() | Where-Object { -not $_.IsDeclined }
}

foreach ($update in $updates) {
    switch ($true) {
        { $DeclineIA64 -and ($update.Title -match 'ia64|itanium' -or $update.LegacyName -match 'ia64|itanium') } {
            Log "Declining $($update.Title) [ia64]"
            if (-not $DryRun) { $update.Decline() }
            break
        }

        { $DeclineARM64 -and $update.Title -match 'arm64' } {
            Log "Declining $($update.Title) [arm64]"
            if (-not $DryRun) { $update.Decline() }
            break
        }

        { $DeclineX86 -and $update.Title -match 'x86' } {
            Log "Declining $($update.Title) [x86]"
            if (-not $DryRun) { $update.Decline() }
            break
        }

        { $DeclineX64 -and $update.Title -match 'x64' } {
            Log "Declining $($update.Title) [x64]"
            if (-not $DryRun) { $update.Decline() }
            break
        }

        { $DeclinePreview -and $update.Title -match 'preview' } {
            Log "Declining $($update.Title) [preview]"
            if (-not $DryRun) { $update.Decline() }
            break
        }

        { $DeclineBeta -and ($update.IsBeta -or $update.Title -match 'beta') } {
            Log "Declining $($update.Title) [beta]"
            if (-not $DryRun) { $update.Decline() }
            break
        }

        { $RestrictToLanguages.Count -gt 0 -and (TestUpdateTitleLanguageMatch -Title $update.Title -AllLocales $allLocales -RestrictToLanguages $RestrictToLanguages) } {
            Log "Declining $($update.Title) [language]"
            if (-not $DryRun) { $update.Decline() }
            break
        }

        { $update.IsSuperseded -or $update.PublicationState -eq "Expired" } {
            # Skip superseded and expired updates, they will be declined later
            continue
        }

        { -not $update.IsApproved -and -not $DeclineOnly } {
            if ($update.IsWsusInfrastructureUpdate -or $approve_classifications.Contains($update.UpdateClassificationTitle)) {
                if ($update.RequiresLicenseAgreementAcceptance) {
                    Log "Accepting license agreement for $($update.Title)"
                    if (-not $DryRun) { $update.AcceptLicenseAgreement() }
                }

                Log "Approving $($update.Title)"
                if (-not $DryRun) { $update.Approve("Install", $group) }
            }
        }
    }
}

# After any new superseding updates have been approved above, superseded and expired updates
# can be declined. We need to handle both here as it seems like superseded updates are also
# marked expired, but some updates are just expired without being superseded.
$updates = $wsus.GetUpdates() | Where-Object { -not $_.IsDeclined }
$updates | ForEach-Object {
    if ($_.IsSuperseded) {
        Log "Declining $($_.Title) [superseded]"
        if (-not $DryRun) { $_.Decline() }
    }
    elseif ($_.IsSuperseded -or $_.PublicationState -eq "Expired") {
        Log "Declining $($_.Title) [expired]"
        if (-not $DryRun) { $_.Decline() }
    }
}