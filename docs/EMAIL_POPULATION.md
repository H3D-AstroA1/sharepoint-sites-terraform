# M365 Email Population Tool

This tool populates Microsoft 365 mailboxes with realistic organizational emails to simulate a real company environment. It's designed for testing, demos, and training scenarios.

## Features

- **Multiple Email Types**: Newsletters, organizational communications, project updates, meeting requests, inter-departmental emails, security alerts, external business communications, and spam
- **Realistic Content**: Department-specific content with professional HTML formatting
- **Comprehensive Variation System**: Extensive content pools for greetings, closings, project names, document names, meeting topics, and more - ensuring each email is unique
- **Azure AD Auto-Discovery**: Automatically discover users and groups from Azure AD for realistic CC/BCC recipients
- **CC/BCC Support**: Emails include realistic CC and BCC recipients from your organization
- **External Business Emails**: Legitimate external communications from vendors, partners, and clients
- **Attachments**: Word, Excel, PowerPoint, and PDF documents with relevant content
- **Email Threading**: Reply chains, forward chains, and reply-all conversations
- **Backdated Emails**: Distributed over 6-12 months with business hours bias using MIME import
- **Sensitivity Labels**: Microsoft 365 sensitivity labels (General, Internal, Confidential, Highly Confidential)
- **SharePoint Integration**: Links to SharePoint sites from your configuration
- **Multi-Folder Support**: Populate emails in inbox, sent items, deleted items, drafts, and junk folders with weighted distribution
- **Realistic Email Properties**: Read/unread status, importance levels, flags, and color categories
- **Rate Limiting**: Configurable delays to prevent API throttling with automatic retry

## Prerequisites

1. **Python 3.8+** with required packages:
   ```bash
   pip install -r requirements.txt
   ```

2. **Azure CLI** logged in with appropriate permissions:
   ```bash
   az login
   ```

3. **App Registration with Permissions**:
   
   Use the menu option `[A] Manage App Registration` to create an app with:
   
   **Microsoft Graph API Permissions**:
   - `Mail.ReadWrite` - To create emails in mailboxes
   - `Mail.Send` - To send emails (optional)
   - `User.Read.All` - For Azure AD user discovery
   
   **Exchange Online API Permissions** (for EWS):
   - `full_access_as_app` - Full mailbox access via EWS
   
   The app registration also creates a **client secret** which is required for EWS authentication.

4. **exchangelib** (Recommended):
   ```bash
   pip install exchangelib
   ```
   
   The `exchangelib` package enables Exchange Web Services (EWS) support, which provides:
   - **Proper email timestamps** - Emails show backdated dates instead of creation time
   - **No "[Draft]" prefix** - Emails appear as received messages, not drafts
   - **Full control over email properties** - Read status, importance, categories
   
   **EWS Authentication**: EWS uses OAuth2 client credentials flow with the app's client_id
   and client_secret from `.app_config.json` (created by the app registration process).
   
   Without `exchangelib`, the tool falls back to Graph API which has limitations:
   - Emails show current timestamp (not backdated)
   - Emails may show "[Draft]" prefix
   
   Run menu option `[0] Check & Install Prerequisites` to automatically install.

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

The tool generates eight types of emails with configurable distribution:

| Category | Default % | Description |
|----------|-----------|-------------|
| 📰 Newsletters | 10% | Company and industry newsletters |
| 🔗 SharePoint Links | 15% | Document sharing and collaboration |
| 📎 Attachments | 15% | Emails with file attachments |
| 📢 Organisational | 15% | Company-wide communications |
| 💬 Inter-departmental | 15% | Team and project communications |
| 🔒 Security | 8% | Account and password notifications |
| 🏢 External Business | 15% | Legitimate external communications (vendors, partners, clients) |
| 🗑️ Spam | 7% | Promotional spam, phishing simulations, scams |

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

### 7. External Business Emails (🏢)

Legitimate external communications from vendors, partners, and clients:

- **Follow-up Emails** - Post-meeting follow-ups, action items
- **Proposals** - Business proposals, partnership opportunities
- **Meeting Requests** - External meeting scheduling
- **Project Updates** - Status updates from external partners
- **Invoices** - Legitimate billing and payment communications
- **Contract Discussions** - Contract reviews, negotiations
- **Introductions** - Business introductions, networking
- **Thank You Notes** - Appreciation for business relationships
- **Support Tickets** - Customer support communications
- **Event Invitations** - Conference invites, webinar registrations
- **Product Updates** - Vendor product announcements
- **Feedback Requests** - Survey and feedback requests

External business emails come from realistic external domains like:
- `acme-consulting.com`, `globaltech-solutions.com`
- `premier-services.net`, `innovate-partners.com`
- `enterprise-systems.io`, `strategic-advisors.com`

### 8. Spam/Junk Emails (🗑️)
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

**Important**: Spam emails ALWAYS use external spam senders. Internal users will never appear as spam senders.

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

## Azure AD Auto-Discovery

The tool can automatically discover users and groups from Azure AD to populate realistic CC/BCC recipients.

### Enabling Azure AD Discovery

Add the following to your `mailboxes.yaml`:

```yaml
azure_ad:
  enabled: true
  discover_users: true
  discover_groups: true
  cache:
    enabled: true
    ttl_minutes: 60
    path: ".azure_ad_cache.json"
  user_filter:
    include_guests: false
    require_mail: true
  group_filter:
    types:
      - unified  # Microsoft 365 Groups
      - security
    max_members: 50
```

### CC/BCC Configuration

Configure how CC and BCC recipients are selected:

```yaml
cc_bcc:
  enabled: true
  cc_probability: 0.3      # 30% of emails have CC
  bcc_probability: 0.1     # 10% of emails have BCC
  max_cc_recipients: 3
  max_bcc_recipients: 2
  prefer_same_department: true
  include_managers: true
```

### Using Azure AD Discovery

Run the menu and select option `[9] Azure AD Discovery` to:
- Discover users and groups from your tenant
- View discovery statistics
- Clear the cache

Or use the command line:
```bash
python scripts/populate_emails.py --all --auto-discover
```

## Rate Limiting

The tool includes configurable rate limiting to prevent API throttling.

### Configuration

Add to your `mailboxes.yaml`:

```yaml
rate_limiting:
  request_delay_ms: 100    # Delay between individual requests (ms)
  batch_delay_ms: 500      # Delay between batches (ms)
  max_retries: 5           # Maximum retry attempts with exponential backoff
```

### Automatic Retry

When rate limited (HTTP 429) or experiencing timeouts, the tool automatically:
1. Reads the `Retry-After` header (if present)
2. Waits the specified time
3. Retries the request with exponential backoff
4. EWS: 3 retries with 2s, 4s, 8s delays
5. Graph API: 5 retries with 1s, 2s, 4s, 8s, 16s delays

### EWS Performance Optimizations

The EWS client includes several performance optimizations:

| Setting | Value | Description |
|---------|-------|-------------|
| Timeout | 90 seconds | Extended timeout for slow Exchange responses |
| Max Retries | 3 | Retry attempts for failed requests |
| Backoff | Exponential | 2s, 4s, 8s between retries |
| Request Delay | 50ms | Faster than Graph API (100ms) |

### Recommendations

| Scenario | request_delay_ms | batch_delay_ms | max_retries |
|----------|------------------|----------------|-------------|
| Small batches (<50 emails) | 50 | 200 | 3 |
| Medium batches (50-200) | 100 | 500 | 5 |
| Large batches (200+) | 200 | 1000 | 7 |

## Date Distribution

Emails are backdated with a weighted distribution to simulate realistic mailbox activity:

| Time Period | Weight | Description |
|-------------|--------|-------------|
| Last 30 days | 40% | Recent emails |
| 1-3 months ago | 30% | Moderately recent |
| 3-6 months ago | 20% | Older emails |
| 6-12 months ago | 10% | Historical emails |

Additional features:
- **Business Hours Bias**: 80% during 8 AM - 6 PM
- **Weekday Bias**: 90% on weekdays
- **Realistic Gaps**: Natural distribution, not uniform
- **MIME Import**: Uses MIME message format for proper date handling via EWS

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

### Emails Show "[Draft]" Prefix or Current Timestamp

This is a known limitation of the Microsoft Graph API. To fix this:

1. **Install exchangelib**:
   ```bash
   pip install exchangelib
   ```
   Or run menu option `[0] Check & Install Prerequisites` to install automatically.

2. **Set up App Registration with EWS permissions**:
   - Run menu option `[A] Manage App Registration`
   - This creates an app with both Graph API and EWS permissions
   - The app includes `full_access_as_app` permission for Exchange Online
   - A client secret is created and stored in `.app_config.json`

3. **Grant Admin Consent**:
   - Admin consent is required for the EWS `full_access_as_app` permission
   - Use the admin consent option in the app registration menu

With `exchangelib` installed and proper permissions, the tool uses Exchange Web Services (EWS) which provides full control over email timestamps and draft status.

### EWS Authentication Failed

If you see "EWS authentication failed, falling back to Graph API":

1. **Check app registration**: Ensure you've created an app via `[A] Manage App Registration`

2. **Verify client secret**: Check that `.app_config.json` exists and contains a valid `client_secret`

3. **Grant admin consent**: EWS requires admin consent for `full_access_as_app` permission
   - Run `[A] Manage App Registration` → `[3] Grant Admin Consent`

4. **Check EWS permission**: Ensure the app has the Exchange Online `full_access_as_app` permission
   - If missing, delete and recreate the app registration

5. **Verify tenant ID**: The `.app_config.json` must have the correct `tenant_id`

### exchangelib Installation Failed

If `exchangelib` installation fails:

1. **Try manual installation**:
   ```bash
   pip install exchangelib>=5.0
   ```

2. **Check Python version**: `exchangelib` requires Python 3.8+

3. **Install dependencies**: On some systems, you may need to install additional packages:
   ```bash
   # Windows
   pip install pywin32
   
   # Linux
   sudo apt-get install libxml2-dev libxslt1-dev
   ```

4. **Use Graph API fallback**: If EWS doesn't work, the tool will automatically fall back to Graph API (with limitations)

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

### Cleanup Modes (Interactive)

When running in interactive mode, you can choose from three deletion modes:

| Mode | Description |
|------|-------------|
| `[1] Move to Deleted Items` | Emails can be recovered from Deleted Items |
| `[2] Permanently delete` | Emails are deleted but may be in Recoverable Items |
| `[3] 🔥 FULL PURGE` | Truly unrecoverable - also purges Recoverable Items folder |

### Full Purge

The **Full Purge** option (mode 3) performs a complete cleanup:

1. Permanently deletes emails from selected folders
2. Purges items from the Recoverable Items folder:
   - `recoverableitemsdeletions` - Soft-deleted items
   - `recoverableitemsversions` - Previous versions
   - `recoverableitemspurges` - Items pending purge

⚠️ **Warning**: Items purged from Recoverable Items cannot be recovered by any means.

### Cleanup Behavior

- **Default**: Moves emails to Deleted Items (recoverable)
- **Permanent**: Permanently deletes emails (may still be in Recoverable Items)
- **Full Purge**: Truly unrecoverable deletion including Recoverable Items
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
├── cleanup_emails.py           # Email cleanup script (with Full Purge support)
└── email_generator/
    ├── __init__.py
    ├── config.py               # Configuration loader
    ├── templates.py            # Template re-exports (backward compatibility)
    ├── content_generator.py    # Dynamic content generation with CC/BCC
    ├── variations.py           # Content variation pools for realistic emails
    ├── attachments.py          # File attachment creation
    ├── threading.py            # Email thread management
    ├── graph_client.py         # Microsoft Graph API client (with EWS fallback)
    ├── ews_client.py           # Exchange Web Services client (for proper timestamps)
    ├── azure_ad_discovery.py   # Azure AD user/group discovery
    ├── user_pool.py            # User pool for CC/BCC recipient selection
    ├── utils.py                # Utility functions
    └── templates/              # Email template modules
        ├── __init__.py                       # Exports all templates
        ├── newsletter_templates.py           # Company & industry newsletters
        ├── sharepoint_templates.py           # Document sharing & site activity
        ├── attachment_templates.py           # Reports & documents for review
        ├── organisational_templates.py       # Announcements, HR, leadership
        ├── interdepartmental_templates.py    # Project updates, meetings, status
        ├── security_templates.py             # Account blocked, password reset
        ├── spam_templates.py                 # Promotional, phishing, scams
        └── external_business_templates.py    # External vendor/partner communications

config/
├── mailboxes.yaml              # Mailbox configuration
├── environments.json           # Tenant configuration
└── sites.json                  # SharePoint sites (for links)
```

### Email Creation Methods

The tool supports two methods for creating emails:

#### 1. Exchange Web Services (EWS) - Recommended

When `exchangelib` is installed, the tool uses EWS which provides:

| Feature | EWS Support |
|---------|-------------|
| Backdated timestamps | ✅ Full control over `datetime_received` and `datetime_sent` |
| No "[Draft]" prefix | ✅ Emails appear as received messages |
| Read/unread status | ✅ Full control |
| All folders | ✅ Inbox, Sent Items, Drafts, Deleted Items, Junk |

#### 2. Microsoft Graph API - Fallback

When `exchangelib` is not installed, the tool falls back to Graph API:

| Feature | Graph API Support |
|---------|-------------------|
| Backdated timestamps | ❌ Shows current time (read-only property) |
| No "[Draft]" prefix | ❌ May show as draft (read-only property) |
| Read/unread status | ✅ Full control |
| All folders | ✅ Inbox, Sent Items, Drafts, Deleted Items, Junk |

**Note**: The Graph API limitations are due to `receivedDateTime`, `sentDateTime`, and `isDraft` being read-only properties in the Microsoft Graph API.

### Content Variation System

The tool includes a comprehensive variation system (`variations.py`) that ensures each generated email is unique and realistic:

#### Greeting Variations
- **Formal**: "Dear {name}", "Good morning {name}"
- **Informal**: "Hi {name}", "Hey {name}"
- **Team**: "Hi Team", "Hello Everyone", "Dear Colleagues"

#### Closing Variations
- **Formal**: "Best regards", "Sincerely", "Respectfully"
- **Informal**: "Thanks", "Cheers", "Talk soon"
- **Professional**: "Best", "Kind regards", "Thank you"

#### Dynamic Content Pools
- **Project Names**: Codenames (Phoenix, Atlas, Titan) + Descriptors (Digital Transformation, Cloud Migration)
- **Document Names**: Department-specific documents with prefixes (Q1, FY2024, Draft)
- **Meeting Topics**: Various types (Sync, Review, Planning) with department context
- **Employee Names**: Diverse pool of first and last names for spotlights
- **Company News**: Product launches, team achievements, sustainability milestones
- **Events**: Town halls, training sessions, team building, workshops

#### Department-Specific Content
Each department has tailored:
- **Jargon**: Technical terms specific to the department
- **Priorities**: Current focus areas and initiatives
- **Key Points**: Metrics and KPIs relevant to the department
- **Contact Information**: Department-specific support channels

#### Time-Based Variations
- Current quarter and fiscal year references
- Random deadlines with multiple formats
- Date references (last Monday, this week, next month)

#### Tone Variations
- **Urgency Levels**: High (URGENT:), Medium (Important:), Low (FYI:)
- **Polite Requests**: "Could you please", "Would you mind", "I'd appreciate if"
- **Acknowledgments**: Various thank you phrases
- **Follow-up Phrases**: Different ways to invite questions

### Template Categories

| Category | File | Templates | Description |
|----------|------|-----------|-------------|
| Newsletters | `newsletter_templates.py` | 2 | Company and industry newsletters |
| SharePoint | `sharepoint_templates.py` | 2 | Document sharing, site activity |
| Attachments | `attachment_templates.py` | 2 | Reports, documents for review |
| Organisational | `organisational_templates.py` | 3 | Announcements, HR policies, leadership |
| Interdepartmental | `interdepartmental_templates.py` | 4 | Project updates, meetings, status reports |
| Security | `security_templates.py` | 5 | Account blocked, password reset, suspicious activity |
| External Business | `external_business_templates.py` | 12 | Vendor/partner communications, proposals, invoices |
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
