# M365 Email Population Tool

This tool populates Microsoft 365 mailboxes with realistic organizational emails to simulate a real company environment. It's designed for testing, demos, and training scenarios.

## Features

- **Multiple Email Types**: Newsletters, organizational communications, project updates, meeting requests, and inter-departmental emails
- **Realistic Content**: Department-specific content with professional HTML formatting
- **Attachments**: Word, Excel, PowerPoint, and PDF documents with relevant content
- **Email Threading**: Reply chains, forward chains, and reply-all conversations
- **Backdated Emails**: Distributed over 6-12 months with business hours bias
- **Sensitivity Labels**: Microsoft 365 sensitivity labels (General, Internal, Confidential, Highly Confidential)
- **SharePoint Integration**: Links to SharePoint sites from your configuration
- **Multi-Folder Support**: Populate emails in inbox, sent items, deleted items, and drafts folders

## Prerequisites

1. **Python 3.8+** with PyYAML installed:
   ```bash
   pip install -r requirements.txt
   ```

2. **Azure CLI** logged in with appropriate permissions:
   ```bash
   az login
   ```

3. **Microsoft Graph API Permissions**:
   - `Mail.ReadWrite` - To create emails in mailboxes
   - `Mail.Send` - To send emails (optional)
   
   Use the menu option `[A] Manage App Registration` to set up permissions.

## Configuration

### Mailboxes Configuration

Edit `config/mailboxes.yaml` to define your mailboxes:

```yaml
settings:
  default_email_count: 100
  date_range_months: 12
  business_hours_bias: 0.8
  thread_probability: 0.4
  internal_sender_ratio: 0.6

users:
  - upn: john.smith@contoso.com
    role: HR Manager
    department: Human Resources
    email_volume: high

  - upn: jane.doe@contoso.com
    role: Financial Analyst
    department: Finance
    email_volume: medium
```

### User Properties

| Property | Description | Required |
|----------|-------------|----------|
| `upn` | User Principal Name (email address) | Yes |
| `role` | Job title/role | Yes |
| `department` | Department name | Yes |
| `email_volume` | Email volume: `low` (50), `medium` (100), `high` (200) | No |

### Settings

| Setting | Description | Default |
|---------|-------------|---------|
| `default_email_count` | Default emails per mailbox | 100 |
| `date_range_months` | How far back to backdate emails | 12 |
| `business_hours_bias` | Probability of business hours (0-1) | 0.8 |
| `thread_probability` | Probability of email being part of thread | 0.4 |
| `internal_sender_ratio` | Ratio of internal vs external senders | 0.6 |

### Environment Configuration

The tool uses `config/environments.json` for tenant configuration. Ensure your environment is configured:

```json
{
  "environments": [
    {
      "name": "Production",
      "azure": {
        "tenant_id": "your-tenant-id",
        "subscription_id": "your-subscription-id"
      },
      "m365": {
        "domain": "contoso.com"
      }
    }
  ]
}
```

## Usage

### Interactive Mode (Recommended)

Run the menu and select option `[6] Populate Mailboxes with Emails`:

```bash
python scripts/menu.py
```

### Command Line

```bash
# Populate all mailboxes with default email count
python scripts/populate_emails.py --all

# Populate all mailboxes with specific email count
python scripts/populate_emails.py --all --emails-per-mailbox 100

# Populate specific number of mailboxes
python scripts/populate_emails.py --count 5 --emails-per-mailbox 50

# Populate specific mailboxes
python scripts/populate_emails.py --mailboxes "user1@contoso.com,user2@contoso.com" --emails-per-mailbox 75

# Random email count per mailbox (50-150)
python scripts/populate_emails.py --all --emails-min 50 --emails-max 150

# Populate emails in all folders (inbox, sent items, deleted items, drafts)
python scripts/populate_emails.py --all --folders all

# Populate emails in specific folders
python scripts/populate_emails.py --all --folders inbox,sentitems

# Dry run (preview without creating emails)
python scripts/populate_emails.py --all --dry-run

# List configured mailboxes
python scripts/populate_emails.py --list-mailboxes
```

### Command Line Options

| Option | Description |
|--------|-------------|
| `-a, --all` | Populate all mailboxes in config |
| `-c, --count N` | Populate N randomly selected mailboxes |
| `-m, --mailboxes LIST` | Comma-separated list of UPNs |
| `-d, --department DEPT` | Filter mailboxes by department |
| `-e, --emails-per-mailbox N` | Fixed number of emails per mailbox |
| `--emails-min N` | Minimum emails per mailbox (for random) |
| `--emails-max N` | Maximum emails per mailbox (for random) |
| `-f, --folders FOLDERS` | Folders to populate: `all` or comma-separated list |
| `--dry-run` | Preview without creating emails |
| `-l, --list-mailboxes` | List configured mailboxes and exit |

### Available Folders

| Folder Name | Description |
|-------------|-------------|
| `inbox` | Inbox folder (default) |
| `sentitems` | Sent Items folder |
| `deleteditems` | Deleted Items folder |
| `drafts` | Drafts folder |

Use `--folders all` to populate all folders, or specify individual folders like `--folders inbox,sentitems,junkemail`.

## Email Types

The tool generates seven types of emails with configurable distribution:

| Category | Default % | Description |
|----------|-----------|-------------|
| 📰 Newsletters | 12% | Company and industry newsletters |
| 🔗 SharePoint Links | 18% | Document sharing and collaboration |
| 📎 Attachments | 15% | Emails with file attachments |
| 📢 Organisational | 15% | Company-wide communications |
| 💬 Inter-departmental | 18% | Team and project communications |
| 🔒 Security | 12% | Account and password notifications |
| 🗑️ Spam | 10% | Promotional spam, phishing simulations, scams |

### 1. Newsletters (📰)
- Company newsletters with updates and announcements
- Industry newsletters with relevant news
- Professional HTML formatting with sections

### 2. SharePoint Links (🔗)
- SharePoint document sharing notifications
- Links to sites from your `sites.json` configuration
- Collaboration invitations

### 3. Attachments (📎)
- Reports with Excel/PDF attachments
- Proposals with Word documents
- Presentations with PowerPoint files
- Department-specific content

### 4. Organisational Communications (📢)
- Company announcements from leadership
- HR policy updates
- IT system notifications
- Compliance reminders

### 5. Inter-departmental Emails (💬)
- Project updates between teams
- Meeting requests
- Status reports
- Collaboration requests

### 6. Security Notifications (🔒)
- **Account Blocked** - Notifications when accounts are locked due to security concerns
- **Password Reset with Temporary Password** - Emails containing randomly generated temporary passwords
- **Password Reset Link** - Emails with password reset links (no password in email)
- **Account Unlocked** - Notifications when accounts are restored
- **Suspicious Activity Alerts** - Warnings about unusual sign-in attempts

#### Temporary Password Generation

Security emails that include temporary passwords generate realistic passwords with:
- 12-16 characters
- Mix of uppercase, lowercase, numbers, and special characters
- Avoids ambiguous characters (0/O, 1/l/I)
- Example: `Kp7#mNx$Qw3&Yz`

These emails are useful for testing:
- Security awareness training scenarios
- Email filtering and DLP policies
- Incident response procedures
- User education about phishing

### 7. Spam/Junk Emails (🗑️)
- **Promotional Spam** - Flash sales, discount offers, limited-time deals
- **Phishing Simulations** - Fake security alerts, account verification requests
- **Lottery Scams** - Prize notifications, winner announcements
- **Fake Invoices** - Fraudulent billing notices, payment demands
- **Newsletter Spam** - Clickbait articles, unsolicited digests

Spam emails are useful for testing:
- Junk mail filtering rules
- Security awareness training
- Phishing detection systems
- User education about scams

**Spam Routing**: 85% of spam emails go to the Junk folder, while 15% "slip through" to the Inbox (simulating real-world spam filter behavior).

## Folder Distribution

When using the `--folders` option, emails are distributed using **weighted distribution** to simulate realistic mailbox patterns:

| Folder | Graph API Name | Weight | Description |
|--------|----------------|--------|-------------|
| 📥 Inbox | `inbox` | 55% | Received emails |
| 📤 Sent Items | `sentitems` | 20% | Sent emails |
| 📝 Drafts | `drafts` | 10% | Draft emails |
| 🗑️ Deleted Items | `deleteditems` | 10% | Deleted emails |
| 🗑️ Junk Email | `junkemail` | 5% | Spam/junk emails |

### Interactive Mode

When running in interactive mode, you'll be prompted to select folders:

```
Which folders should receive emails?

  [1] Inbox only (default)
  [2] All folders (inbox, sent items, deleted items, drafts, junk)
  [3] Select specific folders
```

### Weighted Distribution Behavior

- Emails are distributed using **weighted probabilities** to simulate realistic mailbox patterns
- Inbox receives the most emails (55%), followed by Sent Items (20%)
- Spam emails are automatically routed: 85% to Junk, 15% to Inbox
- The final summary shows the distribution breakdown:

```
Folder Distribution:
  📥 inbox: 55 (55%)
  📤 sentitems: 20 (20%)
  📝 drafts: 10 (10%)
  🗑️ deleteditems: 10 (10%)
  🗑️ junkemail: 5 (5%)
```

## Email Threading

40% of emails are part of threads:

- **Reply Chains**: 2-5 messages in a conversation
- **Forward Chains**: Forwarded emails with context
- **Reply-All**: Group conversations

Threads maintain:
- Consistent subject lines (with Re:/Fwd: prefixes)
- Quoted content from previous messages
- Realistic time gaps between messages

## Sensitivity Labels

Emails are assigned Microsoft 365 sensitivity labels:

| Label | Distribution | Description |
|-------|--------------|-------------|
| General | 40% | Public information |
| Internal | 35% | Internal use only |
| Confidential | 20% | Sensitive business data |
| Highly Confidential | 5% | Restricted access |

## Date Distribution

Emails are backdated over 6-12 months with:

- **Business Hours Bias**: 80% during 8 AM - 6 PM
- **Weekday Bias**: 90% on weekdays
- **Realistic Gaps**: Natural distribution, not uniform

## Attachments

Generated attachments are valid Office documents:

| Type | Extension | Content |
|------|-----------|---------|
| Word | .docx | Reports, memos, policies |
| Excel | .xlsx | Financial data, metrics |
| PowerPoint | .pptx | Presentations, proposals |
| PDF | .pdf | Official documents |

Attachments are department-specific:
- **HR**: Employee handbooks, policies
- **Finance**: Budget reports, forecasts
- **IT**: Technical documentation
- **Marketing**: Campaign materials
- **Sales**: Proposals, contracts

## Troubleshooting

### "Access Denied" Error

Ensure your app registration has `Mail.ReadWrite` permission with admin consent:
1. Run menu option `[A] Manage App Registration`
2. Grant admin consent for the application

### "User Not Found" Error

Verify the UPN in `mailboxes.yaml` matches an existing user in your tenant.

### "Rate Limited" Error

The tool implements automatic retry with exponential backoff. If issues persist:
- Reduce batch size
- Add delays between mailboxes
- Run during off-peak hours

### PyYAML Not Installed

```bash
pip install pyyaml
# or
pip install -r requirements.txt
```

## Email Cleanup

The tool includes a cleanup script to delete emails from mailboxes.

### Cleanup via Menu

Run the menu and select option `[7] Delete Emails from Mailboxes`:

```bash
python scripts/menu.py
```

### Cleanup Command Line

```bash
# Interactive mode
python scripts/cleanup_emails.py

# Delete from all mailboxes (inbox)
python scripts/cleanup_emails.py --all

# Delete from specific mailboxes
python scripts/cleanup_emails.py --mailboxes "user1@contoso.com,user2@contoso.com"

# Delete from sent items
python scripts/cleanup_emails.py --all --folder sentitems

# Permanently delete (skip Deleted Items)
python scripts/cleanup_emails.py --all --permanent

# Empty Deleted Items folder
python scripts/cleanup_emails.py --all --folder deleteditems --permanent

# Also empty Deleted Items after deletion
python scripts/cleanup_emails.py --all --empty-trash

# Dry run (preview only)
python scripts/cleanup_emails.py --all --dry-run
```

### Cleanup Options

| Option | Description |
|--------|-------------|
| `--all` | Clean all mailboxes in config |
| `--mailboxes LIST` | Comma-separated list of UPNs |
| `--folder NAME` | Folder to clean (inbox, sentitems, deleteditems, drafts) |
| `--permanent` | Permanently delete (skip Deleted Items) |
| `--empty-trash` | Also empty Deleted Items folder |
| `--dry-run` | Preview without deleting |

### Cleanup Behavior

- **Default**: Moves emails to Deleted Items (recoverable)
- **Permanent**: Permanently deletes emails (not recoverable)
- **Validation**: Validates mailboxes before deletion
- **Confirmation**: Requires confirmation for destructive operations

## List Mailboxes

View configured mailboxes and validate them against Azure AD.

### List Mailboxes via Menu

Run the menu and select option `[8] List Mailboxes`:

```bash
python scripts/menu.py
```

### Features

- **Configuration View**: Shows all mailboxes from `config/mailboxes.yaml`
- **Azure AD Validation**: Optional validation to check if mailboxes exist
- **Email Count**: Shows number of emails in each mailbox (when validated)
- **Department Summary**: Groups mailboxes by department

### Validation Options

When listing mailboxes, you can choose to:

1. **Validate mailboxes**: Checks each mailbox against Azure AD
   - Verifies user exists
   - Checks if mailbox is accessible
   - Shows email count in inbox

2. **Show configuration only**: Just displays the YAML configuration
   - Faster, no API calls
   - Shows configured settings

### Output Example

```
╔══════════════════════════════════════════════════════════════════════╗
║                        📬 MAILBOX LIST                                ║
╚══════════════════════════════════════════════════════════════════════╝

  #    UPN                                 Department       Status
  ─────────────────────────────────────────────────────────────────────
  1    john.smith@contoso.com              HR               ✓ Valid (42 emails)
  2    jane.doe@contoso.com                Finance          ✓ Valid (128 emails)
  3    bob.wilson@contoso.com              IT               ✓ Valid (85 emails)
  4    invalid.user@contoso.com            Marketing        ✗ Not found
  ─────────────────────────────────────────────────────────────────────

  Summary:
    ✓ Valid mailboxes: 3
    ✗ Invalid mailboxes: 1
    Total configured: 4

  Departments:
    • Finance: 1 mailbox
    • HR: 1 mailbox
    • IT: 1 mailbox
    • Marketing: 1 mailbox
```

## Architecture

```
scripts/
├── populate_emails.py          # Main population script
├── cleanup_emails.py           # Email cleanup script
└── email_generator/
    ├── __init__.py
    ├── config.py               # Configuration loader
    ├── templates.py            # Template re-exports (backward compatibility)
    ├── content_generator.py    # Dynamic content generation
    ├── attachments.py          # File attachment creation
    ├── threading.py            # Email thread management
    ├── graph_client.py         # Microsoft Graph API client
    ├── utils.py                # Utility functions
    └── templates/              # Email template modules
        ├── __init__.py                    # Exports all templates
        ├── newsletter_templates.py        # Company & industry newsletters
        ├── sharepoint_templates.py        # Document sharing & site activity
        ├── attachment_templates.py        # Reports & documents for review
        ├── organisational_templates.py    # Announcements, HR, leadership
        ├── interdepartmental_templates.py # Project updates, meetings, status
        ├── security_templates.py          # Account blocked, password reset
        └── spam_templates.py              # Promotional, phishing, scams

config/
├── mailboxes.yaml              # Mailbox configuration
├── environments.json           # Tenant configuration
└── sites.json                  # SharePoint sites (for links)
```

### Template Categories

| Category | File | Templates | Description |
|----------|------|-----------|-------------|
| Newsletters | `newsletter_templates.py` | 2 | Company and industry newsletters |
| SharePoint | `sharepoint_templates.py` | 2 | Document sharing, site activity |
| Attachments | `attachment_templates.py` | 2 | Reports, documents for review |
| Organisational | `organisational_templates.py` | 3 | Announcements, HR policies, leadership |
| Interdepartmental | `interdepartmental_templates.py` | 4 | Project updates, meetings, status reports |
| Security | `security_templates.py` | 5 | Account blocked, password reset, suspicious activity |
| Spam | `spam_templates.py` | 5 | Promotional, phishing, lottery scams |

## Security Considerations

- **Credentials**: Uses Azure CLI authentication or app registration
- **Permissions**: Requires only necessary Graph API permissions
- **Data**: Generated content is fictional and safe for testing
- **Cleanup**: Use `cleanup_emails.py` to delete emails from mailboxes
- **Confirmation**: Destructive operations require explicit confirmation

## Related Documentation

- [README.md](../README.md) - Main project documentation
- [CONFIGURATION-GUIDE.md](../CONFIGURATION-GUIDE.md) - Configuration details
- [TROUBLESHOOTING.md](TROUBLESHOOTING.md) - Common issues
- [PREREQUISITES.md](../PREREQUISITES.md) - Setup requirements
