# Use instead of approval rules as the approval rule system is too limited

param (
    [string]$WsusServer = 'localhost',
    [int]$Port = 8530,
    [switch]$UseSSL,
    [switch]$NoSync,
    [switch]$Reset,
    [switch]$DryRun,
    [switch]$DeclineOnly,
    [bool]$DeclineIA64 = $true,
    [bool]$DeclineARM64 = $true,
    [bool]$DeclineX86 = $true,
    [bool]$DeclineX64 = $false,
    [bool]$DeclinePreview = $true,
    [bool]$DeclineBeta = $true,
    [string[]]$RestrictToLanguages = @('en-us', 'en-gb') # Only these languages/locales will be kept
)

# Initialize values
$logFile = ('{0}\logs\wsus-approver.Log' -f $PSScriptRoot)

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
$approve_group = 'All Computers'

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
#endregion Function Definitions

#################
# Main Entry Point
###################

# Reset log file
if (Test-Path $logFile) {
    Remove-Item $logFile
}

# Load the WSUS assembly
[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WsusServer, $UseSSL, $Port)
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
while ($subscription.GetSynchronizationStatus() -ne "NotProcessing") {
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

$updates | ForEach-Object {
    switch -Regex ($_.Title) {
        'ia64|itanium' {
            if ($DeclineIA64 -or $_.LegacyName -Match 'ia64|itanium') {
                Log "Declining $($_.Title) [ia64]"
                if (-not $DryRun) { $_.Decline() }
            }
            break
        }
        'arm64' {
            if ($DeclineARM64) {
                Log "Declining $($_.Title) [arm64]"
                if (-not $DryRun) { $_.Decline() }
            }
            break
        }
        'x86' {
            if ($DeclineX86) {
                Log "Declining $($_.Title) [x86]"
                if (-not $DryRun) { $_.Decline() }
            }
            break
        }
        'x64' {
            if ($DeclineX64) {
                Log "Declining $($_.Title) [x64]"
                if (-not $DryRun) { $_.Decline() }
            }
            break
        }
        'preview' {
            if ($DeclinePreview) {
                Log "Declining $($_.Title) [preview]"
                if (-not $DryRun) { $_.Decline() }
            }
            break
        }
        'beta' {
            if ($DeclineBeta -and ($_.IsBeta -or $_.Title -Match 'beta')) {
                Log "Declining $($_.Title) [beta]"
                if (-not $DryRun) { $_.Decline() }
            }
            break
        }
        default {
            if ($RestrictToLanguages.Count -gt 0) {
                if ($_.LocalizedProperties.ContainsKey('Locale')) {
                    $locale = $_.LocalizedProperties['Locale'].Value
                    if ($locale -notin $RestrictToLanguages) {
                        Log "Declining $($_.Title) [language: $locale]"
                        if (-not $DryRun) { $_.Decline() }
                    }
                }
            }
            elseif ($_.IsSuperseded -or $_.PublicationState -eq 'Expired') {
                # Handle superseded and expired packages after any new updates have been approved
                return
            }
            elseif (-not $_.IsApproved -and -not $DeclineOnly) {
                # Add this condition
                if ($_.IsWsusInfrastructureUpdate -or $approve_classifications.Contains($_.UpdateClassificationTitle)) {
                    if ($_.RequiresLicenseAgreementAcceptance) {
                        Log "Accepting license agreement for $($_.Title)"
                        if (-not $DryRun) { $_.AcceptLicenseAgreement() }
                    }

                    Log "Approving $($_.Title)"
                    if (-not $DryRun) { $_.Approve('Install', $group) }
                }
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