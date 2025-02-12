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

#region Debug Values
$DeclineOnly = $true
$DryRun = $true
$NoSync = $true
#endregion Debug Values

# Set default error action to Stop
$ErrorActionPreference = 'Stop'

#region Initialize values
# Log file path
$logFolder = ('{0}\logs' -f $PSScriptRoot)
$script:logFile = ('{0}\wsus-approver_{1}.Log' -f $logFolder, (Get-Date).ToString('dddd_htt'))

# Classifications to process for approval
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

# Include Upgrades classifications if specified
if ($IncludeUpgrades) {
    $approve_classifications += 'Upgrades'
}

# Approval group to target
$approve_group = 'All Computers'

# Get all locales (languages) from the system
[string[]]$allLocales = [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::AllCultures) | Where-Object { $_.Name -match "-" } | Select-Object -ExpandProperty Name
#endregion Initialize values

#region Function Definitions
function Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [System.ConsoleColor]$ForegroundColor = [System.ConsoleColor]::LightGray
    )

    # Format the log entry with a timestamp
    $timestamp = (Get-Date -Format 's')
    $logEntry = '{0} :: {1}' -f $timestamp, $Message

    # Write the log entry to the host (console) with the specified foreground color
    Write-Host ('{0} :: ' -f $timestamp) -NoNewline
    Write-Host $Message -ForegroundColor $ForegroundColor

    # Append the log entry to the file
    Add-Content -Path $script:logFile -Value $logEntry
}

function IsSelected {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.UpdateServices.Administration.IUpdate]$Update
    )

    # Check if the update's classification title is in the list of update classifications
    if ($Update.UpdateClassificationTitle -in $update_classifications.Title) {
        return $true
    }

    # Check if any of the update's product titles are in the list of update categories
    foreach ($product in $Update.ProductTitles) {
        if ($product -in $update_categories.Title) {
            return $true
        }
    }

    # Return false if neither the classification nor the product title matches
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
if (-not (Test-Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder | Out-Null
}

# Reset Log file
if (Test-Path $script:logFile) {
    Remove-Item $script:logFile
}

#region Output script parameter choices
if ($DryRun) {
    Log ('{0} DryRun flag set, no changes will be made.' -f [char]0x2713) -ForegroundColor Yellow
}

if ($NoSync) {
    Log ('{0} NoSync flag set, no synchronization will be performed.' -f [char]0x2713) -ForegroundColor Yellow
}

if ($IncludeUpgrades) {
    Log ('{0} IncludeUpgrades flag set, "Upgrades" classification will be included in the processing.' -f [char]0x2713) -ForegroundColor Yellow
}

if ($DeclineOnly) {
    Log ('{0} DeclineOnly flag set, only declining updates, no approvals will be made.' -f [char]0x2713) -ForegroundColor Yellow
}

if ($Reset) {
    Log ('{0} Reset flag set, the update list will be reset before processing.' -f [char]0x2713) -ForegroundColor Yellow
}
#endregion Output script parameter choices

# Load the WSUS assembly
[reflection.assembly]::LoadWithPartialName('Microsoft.UpdateServices.Administration') | Out-Null
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WsusServer, $UseSSL, $Port)

# Get the groups, subscription, categories and classifications
$group = $wsus.GetComputerTargetGroups() | Where-Object { $_.Name -eq $approve_group }
$subscription = $wsus.GetSubscription()
$update_categories = $subscription.GetUpdateCategories()
$update_classifications = $subscription.GetUpdateClassifications()

# Start synchronization (if not already running)
if (-not $NoSync) {
    if ($subscription.GetSynchronizationStatus() -eq 'NotProcessing') {
        Log 'Starting synchronization...'
        $subscription.StartSynchronization()
    }
}

# Wait for any currently running synchronization jobs to finish before continuing
while ($subscription.GetSynchronizationStatus() -ne 'NotProcessing') {
    Log ('{0} Waiting for synchronization to finish...' -f [char]0x2514) -ForegroundColor DarkGray
    Start-Sleep -Seconds 10
}

# Start by removing deselected updates as there is no need to do further processing on
Log 'Checking for deselected updates...'
$wsus.GetUpdates() | ForEach-Object {
    if (-not (IsSelected $_)) {
        Log ('Deleting deselected update: {0}' -f $_.Title) -ForegroundColor DarkYellow
        if (-not $DryRun) { $wsus.DeleteUpdate($_.Id.UpdateId) }
    }
}

# Refresh the list of updates
if ($Reset) {
    Log 'Resetting update list...'
    $updates = $wsus.GetUpdates()
}
else {
    Log 'Refetching updates...'
    $updates = $wsus.GetUpdates() | Where-Object { -not $_.IsDeclined }
}

Log 'Processing updates...'
foreach ($update in $updates) {
    switch ($true) {
        { $DeclineIA64 -and ($update.Title -match 'ia64|itanium' -or $update.LegacyName -match 'ia64|itanium') } {
            Log ('Declining {0} [ia64]' -f $update.Title)
            if (-not $DryRun) { $update.Decline() } -ForegroundColor DarkRed
            break
        }

        { $DeclineARM64 -and $update.Title -match 'arm64' } {
            Log ('Declining {0} [arm64]' -f $update.Title)
            if (-not $DryRun) { $update.Decline() } -ForegroundColor DarkRed
            break
        }

        { $DeclineX86 -and $update.Title -match 'x86' } {
            Log ('Declining {0} [x86]' -f $update.Title)
            if (-not $DryRun) { $update.Decline() } -ForegroundColor DarkRed
            break
        }

        { $DeclineX64 -and $update.Title -match 'x64' } {
            Log ('Declining {0} [x64]' -f $update.Title)
            if (-not $DryRun) { $update.Decline() } -ForegroundColor DarkRed
            break
        }

        { $DeclinePreview -and $update.Title -match 'preview' } {
            Log ('Declining {0} [preview]' -f $update.Title)
            if (-not $DryRun) { $update.Decline() } -ForegroundColor DarkRed
            break
        }

        { $DeclineBeta -and ($update.IsBeta -or $update.Title -match 'beta') } {
            Log ('Declining {0} [beta]' -f $update.Title)
            if (-not $DryRun) { $update.Decline() } -ForegroundColor DarkRed
            break
        }

        { $RestrictToLanguages.Count -gt 0 -and (TestUpdateTitleLanguageMatch -Title $update.Title -AllLocales $allLocales -RestrictToLanguages $RestrictToLanguages) } {
            Log ('Declining {0} [language]' -f $update.Title)
            if (-not $DryRun) { $update.Decline() } -ForegroundColor DarkRed
            break
        }

        { $update.IsSuperseded -or $update.PublicationState -eq 'Expired' } {
            # Skip superseded and expired updates, they will be declined later
            continue
        }

        { -not $update.IsApproved -and -not $DeclineOnly } {
            if ($update.IsWsusInfrastructureUpdate -or $approve_classifications.Contains($update.UpdateClassificationTitle)) {
                if ($update.RequiresLicenseAgreementAcceptance) {
                    Log ('Accepting license agreement for {0}' -f $update.Title) -ForegroundColor DarkCyan
                    if (-not $DryRun) { $update.AcceptLicenseAgreement() }
                }

                Log ('Approving {0}' -f $update.Title) -ForegroundColor Green
                if (-not $DryRun) { $update.Approve('Install', $group) }
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
        Log ('Declining {0} [superseded]' -f $_.Title) -ForegroundColor DarkRed
        if (-not $DryRun) { $_.Decline() }
    }
    elseif ($_.IsSuperseded -or $_.PublicationState -eq 'Expired') {
        Log ('Declining {0} [expired]' -f $_.Title) -ForegroundColor DarkRed
        if (-not $DryRun) { $_.Decline() }
    }
}