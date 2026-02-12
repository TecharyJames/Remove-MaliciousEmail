#####################################
## Author: James Tarran // Techary ##
#####################################

# Check and import Exchange Online Management module
$module = Get-Module -Name ExchangeOnlineManagement -ListAvailable
if ($null -eq $module) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Prints 'Techary' in ASCII
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

function CountDown {
    param($timeSpan)
    $spinner = @('|', '/', '-', '\')
    $colours = @("Red", "DarkRed", "Magenta", "DarkMagenta", "Blue", "DarkBlue", "Cyan", "DarkCyan", "Green", "DarkGreen", "Yellow", "DarkYellow", "White", "Gray", "DarkGray")
    $colourIndex = 0
    while ($timeSpan -gt 0) {
        foreach ($spin in $spinner) {
            Write-Host "`r$spin" -NoNewline -ForegroundColor $colours[$colourIndex]
            Start-Sleep -Milliseconds 90
        }
        $colourIndex++
        if ($colourIndex -ge $colours.Length) {
            $colourIndex = 0
        }
        $timeSpan = $timeSpan - 1
    }
}

function Get-SearchType {
    Write-Host "`n===== Search Type Selection =====" -ForegroundColor Cyan
    Write-Host "1. Exact Match   - Search for exact subject/sender (recommended for specific emails)"
    Write-Host "2. Wildcard      - Use * for partial matches (e.g., *invoice*, *@malicious.com)"
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

function Get-EmailSubject {
    Write-Host ""
    if ($Script:SearchType -eq "Wildcard") {
        Write-Host "Wildcard Tips: Use * for partial matches" -ForegroundColor Yellow
        Write-Host "  Examples: *invoice*  |  *urgent payment*  |  Your account*" -ForegroundColor Gray
    }
    $Script:Subject = Read-Host "`nEnter the subject of the email (paste recommended for accuracy)"

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
    }
    $Script:SenderAddress = Read-Host "`nEnter the sender email address (paste recommended for accuracy)"

    if ([string]::IsNullOrWhiteSpace($Script:SenderAddress)) {
        Write-Host "Sender address cannot be empty. Please try again." -ForegroundColor Red
        Get-EmailSender
        return
    }
    Get-InfoConfirmation
}

function Get-InfoConfirmation {
    Write-Host "`n===== Search Summary =====" -ForegroundColor Cyan
    Write-Host "Search Type: $Script:SearchType"
    Write-Host "Subject:     $Script:Subject"
    Write-Host "Sender:      $Script:SenderAddress"
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

function Build-ContentMatchQuery {
    $subjectQuery = $Script:Subject
    $senderQuery = $Script:SenderAddress

    # For wildcard searches, ensure proper KQL syntax
    if ($Script:SearchType -eq "Wildcard") {
        # KQL uses * for wildcards - user should include them in their input
        # We just need to ensure proper quoting for phrases with spaces
        if ($subjectQuery -match '\s' -and $subjectQuery -notmatch '^\*.*\*$') {
            # Contains spaces but isn't wrapped in wildcards on both sides
            $subjectQuery = "`"$subjectQuery`""
        }
    } else {
        # Exact match - quote the values to search as phrases
        $subjectQuery = "`"$subjectQuery`""
        $senderQuery = "`"$senderQuery`""
    }

    return "(From:$senderQuery) AND (Subject:$subjectQuery)"
}

function Get-ContentSearchStatus {
    Write-Host "`nWaiting for search to complete..." -ForegroundColor Yellow

    try {
        $maxAttempts = 120  # 10 minute timeout (120 * 5 seconds)
        $attempts = 0

        while ((Get-ComplianceSearch -Identity $Script:RandomIdentity).Status -ne "Completed") {
            CountDown -timeSpan 5
            $attempts++
            if ($attempts -ge $maxAttempts) {
                Write-Host "`nSearch timed out. Please check the compliance center manually." -ForegroundColor Red
                return
            }
        }

        $Script:Items = (Get-ComplianceSearch -Identity $Script:RandomIdentity).Items
        $searchQuery = (Get-ComplianceSearch -Identity $Script:RandomIdentity).ContentMatchQuery

        # Success sound
        [Console]::Beep(659, 125); [Console]::Beep(659, 125); [Console]::Beep(659, 125)
        [Console]::Beep(523, 125); [Console]::Beep(659, 125); [Console]::Beep(784, 375); [Console]::Beep(392, 375)

        Write-Host "`n===== Search Results =====" -ForegroundColor Green
        Write-Host "Query Used:    $searchQuery"
        Write-Host "Items Found:   $Script:Items email(s)"
        Write-Host "==========================" -ForegroundColor Green

        if ($Script:Items -ne 0) {
            New-ComplianceSearchAction -SearchName $Script:RandomIdentity -Purge -PurgeType SoftDelete -Confirm:$false | Out-Null
            Remove-ContentSearchResults
        } else {
            # No results sound
            [Console]::Beep(440, 500); [Console]::Beep(349, 350); [Console]::Beep(440, 500)

            Write-Host "`nNo emails found with the specified criteria." -ForegroundColor Yellow
            Write-Host "Tips:" -ForegroundColor Cyan
            Write-Host "  - Check spelling of subject and sender"
            Write-Host "  - Try using wildcard search with * for partial matches"
            Write-Host "  - Verify the email hasn't already been deleted"

            $retry = Read-Host "`nWould you like to search again? (Y/N)"
            if ($retry.ToUpper() -eq "Y") {
                Get-SearchType
                Get-EmailSubject
                Start-NewSearch
            }
        }
    } catch {
        Write-Host "Error during search: $_" -ForegroundColor Red
    }
}

function Remove-ContentSearchResults {
    Write-Host ""
    $delete = Read-Host "Do you want to delete all $Script:Items found email(s)? (Y/N)"

    switch ($delete.ToUpper()) {
        "Y" {
            Write-Host "`nDeleting $Script:Items email(s), please wait..." -ForegroundColor Yellow

            try {
                $maxAttempts = 120
                $attempts = 0

                while ((Get-ComplianceSearchAction -Identity "$Script:RandomIdentity`_Purge").Status -ne "Completed") {
                    CountDown -timeSpan 5
                    $attempts++
                    if ($attempts -ge $maxAttempts) {
                        Write-Host "`nPurge timed out. Please check the compliance center manually." -ForegroundColor Red
                        return
                    }
                }

                Write-Host "`n===== Deletion Complete =====" -ForegroundColor Green
                Write-Host "Emails Deleted: $Script:Items"
                Write-Host "Search ID:      $Script:RandomIdentity"
                Write-Host "Search Type:    $Script:SearchType"
                Write-Host "Subject:        $Script:Subject"
                Write-Host "Sender:         $Script:SenderAddress"
                Write-Host "==============================" -ForegroundColor Green
                Write-Host "`nScreenshot this message for your records." -ForegroundColor Cyan
                Pause
            } catch {
                Write-Host "Error during deletion: $_" -ForegroundColor Red
            }
        }
        "N" {
            Write-Host "Deletion cancelled. Search results preserved." -ForegroundColor Yellow
        }
        default {
            Write-Host "Please enter Y or N" -ForegroundColor Yellow
            Remove-ContentSearchResults
        }
    }
}

function Start-NewSearch {
    $Script:RandomIdentity = Get-Random -Maximum 999999
    $query = Build-ContentMatchQuery

    Write-Host "`nStarting compliance search..." -ForegroundColor Yellow
    Write-Host "Search ID: $Script:RandomIdentity" -ForegroundColor Gray
    Write-Host "Query: $query" -ForegroundColor Gray

    try {
        New-ComplianceSearch -Name $Script:RandomIdentity -ExchangeLocation All -ContentMatchQuery $query | Start-ComplianceSearch
        Get-ContentSearchStatus
    } catch {
        Write-Host "Error creating search: $_" -ForegroundColor Red
    }
}

# ===== Main Script Execution =====
Print-TecharyLogo

Write-Host "Connecting to Exchange Online Security & Compliance Center..." -ForegroundColor Yellow
try {
    Connect-IPPSSession -ErrorAction Stop
    Write-Host "Connected successfully.`n" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Exchange Online: $_" -ForegroundColor Red
    exit 1
}

Write-Host "===== Remove Malicious Email Tool =====" -ForegroundColor Cyan
Write-Host "This tool searches and removes emails from all mailboxes."
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Warning "Ensure you have the subject and sender of the email.`nThis information can be obtained from the email headers."

# Get search type (exact or wildcard)
Get-SearchType

# Get email details
Get-EmailSubject

# Start the search
Start-NewSearch

# Cleanup
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Yellow
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
Write-Host "Done." -ForegroundColor Green