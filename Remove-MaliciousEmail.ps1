#####################################
## Author: James Tarran // Techary ##
#####################################

<#
.SYNOPSIS
    Searches and removes malicious/phishing emails from all Microsoft 365 mailboxes.

.DESCRIPTION
    A PowerShell tool for IT administrators to search and purge emails across an entire
    Microsoft 365 organization using Compliance Search. Supports wildcard matching,
    date filtering, CSV export, and detailed logging.

.PARAMETER Subject
    The email subject to search for. Use * for wildcards. Separate multiple with commas.

.PARAMETER Sender
    The sender email address to search for. Use * for wildcards. Separate multiple with commas.

.PARAMETER Recipient
    Filter by recipient email address. Optional.

.PARAMETER AttachmentName
    Filter by attachment filename. Use * for wildcards. Optional.

.PARAMETER StartDate
    Search for emails received on or after this date. Format: yyyy-MM-dd

.PARAMETER EndDate
    Search for emails received on or before this date. Format: yyyy-MM-dd

.PARAMETER Last24Hours
    Search only emails received in the last 24 hours.

.PARAMETER Last7Days
    Search only emails received in the last 7 days.

.PARAMETER ExcludeMailboxes
    Comma-separated list of mailboxes to exclude from search.

.PARAMETER ExportPath
    Path to export results CSV. Defaults to script directory.

.PARAMETER LogPath
    Path for log file. Defaults to script directory.

.PARAMETER HardDelete
    Permanently delete emails instead of soft delete.

.PARAMETER PreviewOnly
    Search and preview results without deleting.

.PARAMETER NonInteractive
    Run without prompts (requires all parameters to be specified).

.EXAMPLE
    .\Remove-MaliciousEmail.ps1
    Runs in interactive mode with menus.

.EXAMPLE
    .\Remove-MaliciousEmail.ps1 -Subject "*invoice*" -Sender "*@malicious.com" -Last24Hours -PreviewOnly
    Preview emails matching criteria from last 24 hours.

.EXAMPLE
    .\Remove-MaliciousEmail.ps1 -Subject "Urgent Payment" -Sender "attacker@bad.com" -StartDate "2024-01-01" -EndDate "2024-01-31" -NonInteractive
    Non-interactive search within date range.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Subject,

    [Parameter(Mandatory = $false)]
    [string]$Sender,

    [Parameter(Mandatory = $false)]
    [string]$Recipient,

    [Parameter(Mandatory = $false)]
    [string]$AttachmentName,

    [Parameter(Mandatory = $false)]
    [Nullable[datetime]]$StartDate,

    [Parameter(Mandatory = $false)]
    [Nullable[datetime]]$EndDate,

    [Parameter(Mandatory = $false)]
    [switch]$Last24Hours,

    [Parameter(Mandatory = $false)]
    [switch]$Last7Days,

    [Parameter(Mandatory = $false)]
    [string]$ExcludeMailboxes,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath,

    [Parameter(Mandatory = $false)]
    [switch]$HardDelete,

    [Parameter(Mandatory = $false)]
    [switch]$PreviewOnly,

    [Parameter(Mandatory = $false)]
    [switch]$NonInteractive
)

# ===== Initialize Paths =====
$ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrEmpty($ScriptDirectory)) { $ScriptDirectory = Get-Location }

if ([string]::IsNullOrEmpty($ExportPath)) { $ExportPath = $ScriptDirectory }
if ([string]::IsNullOrEmpty($LogPath)) { $LogPath = $ScriptDirectory }

$Script:LogFile = Join-Path $LogPath "Remove-MaliciousEmail_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# ===== Logging Function =====
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    # Write to log file
    Add-Content -Path $Script:LogFile -Value $logEntry -ErrorAction SilentlyContinue

    # Write to console with color
    switch ($Level) {
        "INFO"    { Write-Host $Message -ForegroundColor White }
        "WARN"    { Write-Host $Message -ForegroundColor Yellow }
        "ERROR"   { Write-Host $Message -ForegroundColor Red }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
    }
}

# ===== Module Check =====
$requiredVersion = [Version]"3.9.0"
$module = Get-Module -Name ExchangeOnlineManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1

if ($null -eq $module) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
} elseif ($module.Version -lt $requiredVersion) {
    Write-Host "Updating ExchangeOnlineManagement module to v3.9.0+..." -ForegroundColor Yellow
    Update-Module -Name ExchangeOnlineManagement -Force
}
Import-Module ExchangeOnlineManagement -MinimumVersion $requiredVersion -ErrorAction Stop

# ===== ASCII Logo =====
function Print-TecharyLogo {
    $logo = @"
     _______        _
    |__   __|      | |
       | | ___  ___| |__   __ _ _ __ _   _
       | |/ _ \/ __| '_ \ / _`` | '__| | | |
       | |  __/ (__| | | | (_| | |  | |_| |
       |_|\___|\___|_| |_|\__,_|_|   \__, |
                                      __/ |
                                     |___/
"@
    Write-Host -ForegroundColor Green $logo
}

# ===== Progress Spinner =====
function Show-Progress {
    param(
        [string]$Activity,
        [int]$SecondsToWait = 5
    )

    $spinner = @('|', '/', '-', '\')
    $spinIndex = 0

    for ($i = 0; $i -lt ($SecondsToWait * 4); $i++) {
        Write-Host "`r$Activity $($spinner[$spinIndex])" -NoNewline -ForegroundColor Cyan
        Start-Sleep -Milliseconds 250
        $spinIndex = ($spinIndex + 1) % 4
    }
    Write-Host "`r$Activity   " -NoNewline
}

# ===== Menu Functions =====
function Get-SearchType {
    Write-Host "`n===== Search Type Selection =====" -ForegroundColor Cyan
    Write-Host "1. Exact Match   - Search for exact subject/sender"
    Write-Host "2. Wildcard      - Use * for partial matches (e.g., *invoice*, *@domain.com)"
    Write-Host "=================================" -ForegroundColor Cyan

    do {
        $choice = Read-Host "`nSelect search type (1 or 2)"
        switch ($choice) {
            "1" { $Script:SearchType = "Exact"; return }
            "2" { $Script:SearchType = "Wildcard"; return }
            default { Write-Host "Please enter 1 or 2" -ForegroundColor Yellow }
        }
    } until ($choice -eq "1" -or $choice -eq "2")
}

function Get-DateFilter {
    Write-Host "`n===== Date Range Filter =====" -ForegroundColor Cyan
    Write-Host "1. All time        - Search all emails (no date filter)"
    Write-Host "2. Last 24 hours   - Emails received in the last day"
    Write-Host "3. Last 7 days     - Emails received in the last week"
    Write-Host "4. Last 30 days    - Emails received in the last month"
    Write-Host "5. Custom range    - Specify start and end dates"
    Write-Host "==============================" -ForegroundColor Cyan

    do {
        $choice = Read-Host "`nSelect date range (1-5)"
        switch ($choice) {
            "1" {
                $Script:StartDate = $null
                $Script:EndDate = $null
                return
            }
            "2" {
                $Script:StartDate = (Get-Date).AddDays(-1)
                $Script:EndDate = Get-Date
                return
            }
            "3" {
                $Script:StartDate = (Get-Date).AddDays(-7)
                $Script:EndDate = Get-Date
                return
            }
            "4" {
                $Script:StartDate = (Get-Date).AddDays(-30)
                $Script:EndDate = Get-Date
                return
            }
            "5" {
                $startInput = Read-Host "Enter start date (yyyy-MM-dd)"
                $endInput = Read-Host "Enter end date (yyyy-MM-dd)"
                try {
                    $Script:StartDate = [datetime]::ParseExact($startInput, "yyyy-MM-dd", $null)
                    $Script:EndDate = [datetime]::ParseExact($endInput, "yyyy-MM-dd", $null)
                    return
                } catch {
                    Write-Host "Invalid date format. Please use yyyy-MM-dd" -ForegroundColor Red
                }
            }
            default { Write-Host "Please enter 1-5" -ForegroundColor Yellow }
        }
    } until ($false)
}

function Get-EmailSubject {
    Write-Host ""
    if ($Script:SearchType -eq "Wildcard") {
        Write-Host "Wildcard Tips: Use * for partial matches" -ForegroundColor Yellow
        Write-Host "  Examples: *invoice*  |  *urgent payment*  |  Your account*" -ForegroundColor Gray
        Write-Host "  Multiple: invoice,payment,urgent (comma-separated)" -ForegroundColor Gray
    }
    $Script:Subject = Read-Host "`nEnter the subject of the email (comma-separate for multiple)"

    if ([string]::IsNullOrWhiteSpace($Script:Subject)) {
        Write-Host "Subject cannot be empty. Please try again." -ForegroundColor Red
        Get-EmailSubject
        return
    }
    Get-EmailSender
}

function Get-EmailSender {
    Write-Host ""
    if ($Script:SearchType -eq "Wildcard") {
        Write-Host "Wildcard Tips: Use * for partial matches" -ForegroundColor Yellow
        Write-Host "  Examples: *@malicious.com  |  attacker*  |  *phish*" -ForegroundColor Gray
        Write-Host "  Multiple: bad@evil.com,*@phish.com (comma-separated)" -ForegroundColor Gray
    }
    $Script:SenderAddress = Read-Host "`nEnter the sender email address (comma-separate for multiple)"

    if ([string]::IsNullOrWhiteSpace($Script:SenderAddress)) {
        Write-Host "Sender address cannot be empty. Please try again." -ForegroundColor Red
        Get-EmailSender
        return
    }
    Get-OptionalFilters
}

function Get-OptionalFilters {
    Write-Host "`n===== Optional Filters =====" -ForegroundColor Cyan
    Write-Host "Press Enter to skip any optional filter"
    Write-Host "=============================" -ForegroundColor Cyan

    $Script:Recipient = Read-Host "`nFilter by recipient (optional)"
    $Script:AttachmentName = Read-Host "Filter by attachment name (optional, e.g., *.exe, invoice.pdf)"
    $Script:ExcludeMailboxes = Read-Host "Exclude mailboxes (optional, comma-separated)"

    Get-InfoConfirmation
}

function Get-InfoConfirmation {
    $dateRange = if ($Script:StartDate -and $Script:EndDate) {
        "$($Script:StartDate.ToString('yyyy-MM-dd')) to $($Script:EndDate.ToString('yyyy-MM-dd'))"
    } else { "All time" }

    Write-Host "`n===== Search Summary =====" -ForegroundColor Cyan
    Write-Host "Search Type:    $Script:SearchType"
    Write-Host "Subject:        $Script:Subject"
    Write-Host "Sender:         $Script:SenderAddress"
    Write-Host "Date Range:     $dateRange"
    if ($Script:Recipient) { Write-Host "Recipient:      $Script:Recipient" }
    if ($Script:AttachmentName) { Write-Host "Attachment:     $Script:AttachmentName" }
    if ($Script:ExcludeMailboxes) { Write-Host "Excluded:       $Script:ExcludeMailboxes" }
    Write-Host "==========================" -ForegroundColor Cyan

    do {
        $confirm = Read-Host "`nIs this correct? (Y/N)"
        switch ($confirm.ToUpper()) {
            "Y" { return }
            "N" { Get-EmailSubject; return }
            default { Write-Host "Please enter Y or N" -ForegroundColor Yellow }
        }
    } until ($confirm.ToUpper() -eq "Y")
}

function Get-DeleteType {
    Write-Host "`n===== Deletion Type =====" -ForegroundColor Cyan
    Write-Host "1. Soft Delete (Recommended) - Moves to Recoverable Items, can be restored"
    Write-Host "2. Hard Delete               - Permanently deletes, cannot be recovered"
    Write-Host "3. Cancel                    - Do not delete"
    Write-Host "==========================" -ForegroundColor Cyan

    do {
        $choice = Read-Host "`nSelect deletion type (1, 2, or 3)"
        switch ($choice) {
            "1" { return "SoftDelete" }
            "2" {
                $confirm = Read-Host "Are you sure? Hard delete cannot be undone (YES to confirm)"
                if ($confirm -eq "YES") { return "HardDelete" }
                else { Write-Host "Hard delete cancelled. Using soft delete." -ForegroundColor Yellow; return "SoftDelete" }
            }
            "3" { return "Cancel" }
            default { Write-Host "Please enter 1, 2, or 3" -ForegroundColor Yellow }
        }
    } until ($false)
}

# ===== Query Building =====
function Build-ContentMatchQuery {
    $queryParts = @()

    # Build subject query (support multiple with OR)
    $subjects = $Script:Subject -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $subjectQueries = @()
    foreach ($subj in $subjects) {
        if ($Script:SearchType -eq "Exact") {
            $subjectQueries += "Subject:`"$subj`""
        } else {
            $subjectQueries += "Subject:$subj"
        }
    }
    if ($subjectQueries.Count -gt 1) {
        $queryParts += "(" + ($subjectQueries -join " OR ") + ")"
    } else {
        $queryParts += $subjectQueries[0]
    }

    # Build sender query (support multiple with OR)
    $senders = $Script:SenderAddress -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $senderQueries = @()
    foreach ($sndr in $senders) {
        if ($Script:SearchType -eq "Exact") {
            $senderQueries += "From:`"$sndr`""
        } else {
            $senderQueries += "From:$sndr"
        }
    }
    if ($senderQueries.Count -gt 1) {
        $queryParts += "(" + ($senderQueries -join " OR ") + ")"
    } else {
        $queryParts += $senderQueries[0]
    }

    # Add recipient filter
    if ($Script:Recipient) {
        $queryParts += "To:$($Script:Recipient)"
    }

    # Add attachment filter
    if ($Script:AttachmentName) {
        $queryParts += "Attachment:$($Script:AttachmentName)"
    }

    # Add date range filter
    if ($Script:StartDate) {
        $queryParts += "Received>=$($Script:StartDate.ToString('yyyy-MM-dd'))"
    }
    if ($Script:EndDate) {
        $queryParts += "Received<=$($Script:EndDate.ToString('yyyy-MM-dd'))"
    }

    return ($queryParts -join " AND ")
}

# ===== Export Function =====
function Export-SearchResults {
    param([string]$SearchName)

    try {
        Write-Log "Exporting search results to CSV..." -Level INFO

        # Get search results directly from the compliance search
        $search = Get-ComplianceSearch -Identity $SearchName
        $successResults = $search.SuccessResults

        if ($successResults) {
            $csvPath = Join-Path $ExportPath "EmailSearch_${SearchName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

            # Parse and export results
            $parsedResults = @()
            $resultLines = $successResults -split ';'

            foreach ($line in $resultLines) {
                if ($line -match "Location:\s*(.+?),\s*Item count:\s*(\d+),\s*Total size:\s*(\d+)") {
                    $parsedResults += [PSCustomObject]@{
                        Mailbox    = $matches[1].Trim()
                        ItemCount  = [int]$matches[2]
                        TotalSize  = [int]$matches[3]
                        SearchName = $SearchName
                        SearchDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        Subject    = $Script:Subject
                        Sender     = $Script:SenderAddress
                    }
                }
            }

            if ($parsedResults.Count -gt 0) {
                $parsedResults | Export-Csv -Path $csvPath -NoTypeInformation
                Write-Log "Results exported to: $csvPath" -Level SUCCESS
                return $csvPath
            }
        }

        Write-Log "No detailed results available for export" -Level WARN
        return $null
    } catch {
        Write-Log "Error exporting results: $_" -Level ERROR
        return $null
    }
}

# ===== Preview Function =====
function Show-SearchPreview {
    param([string]$SearchName)

    Write-Host "`n===== Affected Mailboxes =====" -ForegroundColor Cyan

    try {
        # Get search results directly from the compliance search
        $search = Get-ComplianceSearch -Identity $SearchName
        $successResults = $search.SuccessResults

        if ($successResults) {
            Write-Host ""
            Write-Host "Mailboxes with matching emails:" -ForegroundColor Yellow
            Write-Host "-" * 60

            $resultLines = $successResults -split ';'
            $displayCount = 0

            foreach ($line in $resultLines) {
                if ($line -match "Location:\s*(.+?),\s*Item count:\s*(\d+),\s*Total size:\s*(\d+)" -and $displayCount -lt 20) {
                    $mailbox = $matches[1].Trim()
                    $count = $matches[2]
                    Write-Host "  $mailbox - $count item(s)" -ForegroundColor White
                    $displayCount++
                }
            }

            if ($resultLines.Count -gt 20) {
                Write-Host "  ... and more (see exported CSV for full list)" -ForegroundColor Gray
            }
            Write-Host "-" * 60
        } else {
            Write-Host "No mailbox details available" -ForegroundColor Gray
        }
    } catch {
        Write-Log "Error showing preview: $_" -Level ERROR
    }
}

# ===== Search Functions =====
function Get-ContentSearchStatus {
    Write-Log "Waiting for search to complete..." -Level INFO

    try {
        $maxAttempts = 120
        $attempts = 0

        while ((Get-ComplianceSearch -Identity $Script:RandomIdentity).Status -ne "Completed") {
            Show-Progress -Activity "Searching mailboxes" -SecondsToWait 5
            $attempts++

            # Show progress percentage if available
            $search = Get-ComplianceSearch -Identity $Script:RandomIdentity
            if ($search.JobProgress -gt 0) {
                Write-Host "`rSearch progress: $($search.JobProgress)%   " -NoNewline -ForegroundColor Cyan
            }

            if ($attempts -ge $maxAttempts) {
                Write-Log "Search timed out after 10 minutes. Check compliance center manually." -Level ERROR
                return
            }
        }

        Write-Host ""
        $Script:Items = (Get-ComplianceSearch -Identity $Script:RandomIdentity).Items
        $searchQuery = (Get-ComplianceSearch -Identity $Script:RandomIdentity).ContentMatchQuery
        $searchSize = (Get-ComplianceSearch -Identity $Script:RandomIdentity).Size

        # Format size
        $sizeFormatted = if ($searchSize -gt 1GB) { "{0:N2} GB" -f ($searchSize / 1GB) }
                         elseif ($searchSize -gt 1MB) { "{0:N2} MB" -f ($searchSize / 1MB) }
                         elseif ($searchSize -gt 1KB) { "{0:N2} KB" -f ($searchSize / 1KB) }
                         else { "$searchSize bytes" }

        # Success sound
        [Console]::Beep(659, 125); [Console]::Beep(659, 125); [Console]::Beep(784, 375)

        Write-Host "`n===== Search Results =====" -ForegroundColor Green
        Write-Host "Search ID:     $Script:RandomIdentity"
        Write-Host "Query:         $searchQuery"
        Write-Host "Items Found:   $Script:Items email(s)"
        Write-Host "Total Size:    $sizeFormatted"
        Write-Host "==========================" -ForegroundColor Green

        Write-Log "Search completed. Found $Script:Items items ($sizeFormatted)" -Level SUCCESS

        if ($Script:Items -ne 0) {
            # Show preview
            Show-SearchPreview -SearchName $Script:RandomIdentity

            # Export results
            $exportedFile = Export-SearchResults -SearchName $Script:RandomIdentity

            if ($PreviewOnly) {
                Write-Log "Preview only mode - no deletion performed" -Level INFO
                return
            }

            Remove-ContentSearchResults
        } else {
            [Console]::Beep(440, 500); [Console]::Beep(349, 350)

            Write-Log "No emails found with the specified criteria." -Level WARN
            Write-Host "`nTips:" -ForegroundColor Cyan
            Write-Host "  - Check spelling of subject and sender"
            Write-Host "  - Try using wildcard search with * for partial matches"
            Write-Host "  - Adjust the date range"
            Write-Host "  - Verify the email hasn't already been deleted"

            if (-not $NonInteractive) {
                $retry = Read-Host "`nWould you like to search again? (Y/N)"
                if ($retry.ToUpper() -eq "Y") {
                    Get-SearchType
                    Get-DateFilter
                    Get-EmailSubject
                    Start-NewSearch
                }
            }
        }
    } catch {
        Write-Log "Error during search: $_" -Level ERROR
    }
}

function Remove-ContentSearchResults {
    Write-Host ""

    # Get deletion type
    if ($HardDelete) {
        $purgeType = "HardDelete"
    } elseif ($NonInteractive) {
        $purgeType = "SoftDelete"
    } else {
        $purgeType = Get-DeleteType
    }

    if ($purgeType -eq "Cancel") {
        Write-Log "Deletion cancelled by user." -Level INFO
        return
    }

    Write-Log "Initiating $purgeType for $Script:Items email(s)..." -Level INFO

    try {
        New-ComplianceSearchAction -SearchName $Script:RandomIdentity -Purge -PurgeType $purgeType -Confirm:$false | Out-Null

        $maxAttempts = 120
        $attempts = 0

        while ((Get-ComplianceSearchAction -Identity "$Script:RandomIdentity`_Purge").Status -ne "Completed") {
            Show-Progress -Activity "Deleting emails ($purgeType)" -SecondsToWait 5
            $attempts++
            if ($attempts -ge $maxAttempts) {
                Write-Log "Purge timed out. Check compliance center manually." -Level ERROR
                return
            }
        }

        Write-Host ""
        Write-Host "`n===== Deletion Complete =====" -ForegroundColor Green
        Write-Host "Emails Deleted: $Script:Items"
        Write-Host "Delete Type:    $purgeType"
        Write-Host "Search ID:      $Script:RandomIdentity"
        Write-Host "Search Type:    $Script:SearchType"
        Write-Host "Subject:        $Script:Subject"
        Write-Host "Sender:         $Script:SenderAddress"
        if ($Script:StartDate) { Write-Host "Date Range:     $($Script:StartDate.ToString('yyyy-MM-dd')) to $($Script:EndDate.ToString('yyyy-MM-dd'))" }
        Write-Host "Log File:       $Script:LogFile"
        Write-Host "==============================" -ForegroundColor Green

        Write-Log "Deletion completed. $Script:Items emails removed using $purgeType." -Level SUCCESS
        Write-Host "`nScreenshot this message for your records." -ForegroundColor Cyan

        if (-not $NonInteractive) { Pause }
    } catch {
        Write-Log "Error during deletion: $_" -Level ERROR
    }
}

function Start-NewSearch {
    $Script:RandomIdentity = "MalEmail_" + (Get-Date -Format "yyyyMMdd_HHmmss") + "_" + (Get-Random -Maximum 9999)
    $query = Build-ContentMatchQuery

    Write-Log "Starting compliance search..." -Level INFO
    Write-Host "Search ID: $Script:RandomIdentity" -ForegroundColor Gray
    Write-Host "Query: $query" -ForegroundColor Gray

    try {
        # Determine exchange locations
        if ($Script:ExcludeMailboxes) {
            $excludeList = $Script:ExcludeMailboxes -split ',' | ForEach-Object { $_.Trim() }
            New-ComplianceSearch -Name $Script:RandomIdentity -ExchangeLocation All -ExchangeLocationExclusion $excludeList -ContentMatchQuery $query | Out-Null
        } else {
            New-ComplianceSearch -Name $Script:RandomIdentity -ExchangeLocation All -ContentMatchQuery $query | Out-Null
        }

        Start-ComplianceSearch -Identity $Script:RandomIdentity
        Get-ContentSearchStatus
    } catch {
        Write-Log "Error creating search: $_" -Level ERROR
    }
}

# ===== Main Script Execution =====
Print-TecharyLogo

Write-Host "Remove Malicious Email Tool" -ForegroundColor Cyan
Write-Host "===========================`n" -ForegroundColor Cyan

Write-Log "Script started. Log file: $Script:LogFile" -Level INFO

# Disconnect any existing sessions to avoid assembly conflicts
try {
    $existingSessions = Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*" -or $_.Name -like "*ExchangeOnline*" }
    if ($existingSessions) {
        Write-Log "Closing existing Exchange sessions..." -Level WARN
        $existingSessions | Remove-PSSession -ErrorAction SilentlyContinue
    }
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 2>$null
} catch {
    # Ignore disconnection errors
}

# Connect to Exchange Online Security & Compliance Center
Write-Log "Connecting to Exchange Online Security & Compliance Center..." -Level INFO
try {
    Connect-IPPSSession -EnableSearchOnlySession -ErrorAction Stop
    Write-Log "Connected successfully." -Level SUCCESS
} catch {
    Write-Log "Failed to connect to Exchange Online: $_" -Level ERROR
    Write-Log "Ensure you have ExchangeOnlineManagement v3.9.0+ installed" -Level ERROR
    Write-Log "If issue persists, close PowerShell completely and reopen." -Level ERROR
    exit 1
}

# Handle command-line parameters (non-interactive mode)
if ($Subject -and $Sender) {
    Write-Log "Running in parameter mode" -Level INFO

    $Script:Subject = $Subject
    $Script:SenderAddress = $Sender
    $Script:Recipient = $Recipient
    $Script:AttachmentName = $AttachmentName
    $Script:ExcludeMailboxes = $ExcludeMailboxes
    $Script:SearchType = if ($Subject -match '\*' -or $Sender -match '\*') { "Wildcard" } else { "Exact" }

    # Handle date parameters
    if ($Last24Hours) {
        $Script:StartDate = (Get-Date).AddDays(-1)
        $Script:EndDate = Get-Date
    } elseif ($Last7Days) {
        $Script:StartDate = (Get-Date).AddDays(-7)
        $Script:EndDate = Get-Date
    } else {
        $Script:StartDate = $StartDate
        $Script:EndDate = $EndDate
    }

    Start-NewSearch
} else {
    # Interactive mode
    Write-Warning "Ensure you have the subject and sender of the email.`nThis information can be obtained from the email headers."

    Get-SearchType
    Get-DateFilter
    Get-EmailSubject
    Start-NewSearch
}

# Cleanup
Write-Host ""
Write-Log "Disconnecting from Exchange Online..." -Level INFO
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
Write-Log "Session complete." -Level SUCCESS
