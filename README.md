# Remove-MaliciousEmail

A PowerShell tool for searching and removing malicious or phishing emails from all mailboxes in Microsoft 365 Exchange Online using Compliance Search.

## Requirements

- PowerShell 5.1 or later
- Microsoft 365 account with **Compliance Administrator** or **eDiscovery Manager** role
- Internet connection

The script will automatically install the `ExchangeOnlineManagement` module if not present.

## Usage

```powershell
.\Remove-MaliciousEmail.ps1
```

The script will:
1. Connect to Exchange Online Security & Compliance Center
2. Prompt you to select a search type (Exact or Wildcard)
3. Ask for the email subject and sender address
4. Search all mailboxes for matching emails
5. Display results and offer to delete found emails

## Search Types

### Exact Match
Searches for emails with the exact subject and sender you specify. Best for removing a specific known email.

### Wildcard Search
Use `*` as a wildcard character to match partial text. Useful when:
- You only know part of the subject line
- Emails come from multiple addresses at the same domain
- The attacker uses variations of the same subject

#### Wildcard Examples

| Field | Example | Matches |
|-------|---------|---------|
| Subject | `*invoice*` | "Your invoice is ready", "Invoice #12345", "Unpaid invoice reminder" |
| Subject | `Your account*` | "Your account has been suspended", "Your account verification" |
| Sender | `*@malicious.com` | Any sender from malicious.com domain |
| Sender | `support*` | support@example.com, support-team@company.com |
| Sender | `*phish*` | phishing@bad.com, nophish@test.com |

## Features

- **Wildcard support** - Partial matching for subject and sender
- **Input validation** - Prevents empty searches
- **Timeout protection** - 10-minute maximum wait for searches
- **Error handling** - Graceful handling of connection and search failures
- **Audit trail** - Displays search ID and query for documentation
- **Soft delete** - Emails moved to Recoverable Items (can be restored if needed)
- **Retry option** - Search again if no results found

## Output

After deletion, the script displays a summary suitable for screenshots:

```
===== Deletion Complete =====
Emails Deleted: 47
Search ID:      123456
Search Type:    Wildcard
Subject:        *invoice*
Sender:         *@malicious.com
==============================
```

## Permissions Required

To run compliance searches and purge emails, your account needs one of these roles in Microsoft 365:

- **Compliance Administrator**
- **eDiscovery Manager**
- **Organization Management** (with Search and Purge role)

## Notes

- Searches cover all mailboxes in the organization
- Deleted emails are soft-deleted (moved to Recoverable Items folder)
- Users can recover soft-deleted items for 14 days by default
- Large searches may take several minutes to complete
- Search results are limited by Microsoft's compliance search limits

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Failed to connect" | Verify your account has compliance permissions |
| No results found | Try wildcard search, check spelling, verify email exists |
| Search timeout | Check Microsoft 365 Service Health, try again later |
| Module install fails | Run PowerShell as Administrator |

## Author

James Tarran // [Techary](https://techary.com)

## License

MIT License - Feel free to use and modify for your organization's needs.
