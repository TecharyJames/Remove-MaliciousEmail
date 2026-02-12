# Remove-MaliciousEmail

A PowerShell tool for searching and removing malicious or phishing emails from all mailboxes in Microsoft 365 Exchange Online using Compliance Search.

## Requirements

- PowerShell 5.1 or later
- Microsoft 365 account with **Compliance Administrator** or **eDiscovery Manager** role
- Internet connection

The script will automatically install the `ExchangeOnlineManagement` module if not present.

## Usage

### Interactive Mode

```powershell
.\Remove-MaliciousEmail.ps1
```

The script will guide you through:
1. Search type selection (Exact or Wildcard)
2. Date range filter
3. Email subject and sender
4. Optional filters (recipient, attachment, exclusions)
5. Preview results and export to CSV
6. Choose deletion type (Soft or Hard delete)

### Command-Line Mode

```powershell
# Preview emails from last 24 hours (no deletion)
.\Remove-MaliciousEmail.ps1 -Subject "*invoice*" -Sender "*@malicious.com" -Last24Hours -PreviewOnly

# Delete emails within a date range
.\Remove-MaliciousEmail.ps1 -Subject "Urgent Payment" -Sender "attacker@bad.com" -StartDate "2024-01-01" -EndDate "2024-01-31" -NonInteractive

# Search with all filters
.\Remove-MaliciousEmail.ps1 -Subject "*invoice*,*payment*" -Sender "*@evil.com" -Recipient "finance@company.com" -AttachmentName "*.exe" -Last7Days -ExcludeMailboxes "ceo@company.com,legal@company.com"

# Hard delete (permanent)
.\Remove-MaliciousEmail.ps1 -Subject "Phishing" -Sender "bad@actor.com" -HardDelete -NonInteractive
```

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-Subject` | Email subject to search. Use `*` for wildcards. Comma-separate for multiple. |
| `-Sender` | Sender email address. Use `*` for wildcards. Comma-separate for multiple. |
| `-Recipient` | Filter by recipient email address (optional) |
| `-AttachmentName` | Filter by attachment filename, e.g., `*.exe`, `invoice.pdf` (optional) |
| `-StartDate` | Search emails from this date. Format: `yyyy-MM-dd` |
| `-EndDate` | Search emails until this date. Format: `yyyy-MM-dd` |
| `-Last24Hours` | Search only emails from the last 24 hours |
| `-Last7Days` | Search only emails from the last 7 days |
| `-ExcludeMailboxes` | Comma-separated mailboxes to skip |
| `-ExportPath` | Directory for CSV export (default: script directory) |
| `-LogPath` | Directory for log files (default: script directory) |
| `-HardDelete` | Permanently delete instead of soft delete |
| `-PreviewOnly` | Search and preview without deleting |
| `-NonInteractive` | Run without prompts (for automation) |

## Search Types

### Exact Match
Searches for emails with the exact subject and sender you specify. Best for removing a specific known email.

### Wildcard Search
Use `*` as a wildcard character to match partial text. Useful when:
- You only know part of the subject line
- Emails come from multiple addresses at the same domain
- The attacker uses variations of the same subject

### Multiple Values
Comma-separate values to search for multiple subjects or senders (uses OR logic):
```powershell
-Subject "*invoice*,*payment*,*urgent*"
-Sender "*@malicious.com,*@phishing.net"
```

## Wildcard Examples

| Field | Example | Matches |
|-------|---------|---------|
| Subject | `*invoice*` | "Your invoice is ready", "Invoice #12345" |
| Subject | `Your account*` | "Your account has been suspended" |
| Sender | `*@malicious.com` | Any sender from malicious.com domain |
| Sender | `support*` | support@example.com, support-team@company.com |
| Attachment | `*.exe` | Any .exe attachment |
| Attachment | `invoice*.pdf` | invoice.pdf, invoice_2024.pdf |

## Features

| Feature | Description |
|---------|-------------|
| **Wildcard Search** | Use `*` for partial matching on subject, sender, attachments |
| **Multiple Patterns** | Search for multiple subjects/senders in one query |
| **Date Filtering** | Filter by last 24h, 7 days, 30 days, or custom range |
| **Recipient Filter** | Target emails sent to specific users |
| **Attachment Filter** | Find emails with specific attachment names |
| **Mailbox Exclusions** | Skip specific mailboxes from search |
| **Preview Mode** | View results without deleting |
| **CSV Export** | Automatically exports results for documentation |
| **Logging** | All actions logged to timestamped log file |
| **Hard/Soft Delete** | Choose between recoverable or permanent deletion |
| **Progress Display** | Shows search progress percentage |
| **Non-Interactive** | Command-line mode for scripting/automation |

## Output Files

The script generates two files in the script directory (or specified paths):

### Log File
`Remove-MaliciousEmail_20240115_143022.log`
```
[2024-01-15 14:30:22] [INFO] Script started
[2024-01-15 14:30:25] [SUCCESS] Connected successfully
[2024-01-15 14:31:45] [SUCCESS] Search completed. Found 47 items (2.5 MB)
[2024-01-15 14:32:10] [SUCCESS] Deletion completed. 47 emails removed using SoftDelete
```

### CSV Export
`EmailSearch_MalEmail_20240115_143022_1234_20240115_143145.csv`
```csv
Mailbox,ItemCount,TotalSize,SearchName,SearchDate,Subject,Sender
user1@company.com,5,102400,MalEmail_20240115,2024-01-15 14:31:45,*invoice*,*@malicious.com
user2@company.com,12,256000,MalEmail_20240115,2024-01-15 14:31:45,*invoice*,*@malicious.com
```

## Deletion Types

| Type | Description | Recovery |
|------|-------------|----------|
| **Soft Delete** | Moves to Recoverable Items folder | Users can restore for 14 days |
| **Hard Delete** | Permanently removes from mailbox | Cannot be recovered |

## Permissions Required

Your Microsoft 365 account needs one of these roles:

- **Compliance Administrator**
- **eDiscovery Manager**
- **Organization Management** (with Search and Purge role)

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Failed to connect" | Verify your account has compliance permissions |
| No results found | Try wildcard search, adjust date range, check spelling |
| Search timeout | Check Microsoft 365 Service Health, try again later |
| Module install fails | Run PowerShell as Administrator |
| Preview not loading | Large result sets take longer, wait for completion |

## Notes

- Searches cover all mailboxes in the organization (unless exclusions specified)
- Soft-deleted emails can be recovered by users for 14 days by default
- Large searches may take several minutes to complete
- All actions are logged for audit purposes
- CSV exports provide documentation for incident response

## Author

James Tarran // [Techary](https://techary.com)

## License

MIT License - Feel free to use and modify for your organization's needs.
