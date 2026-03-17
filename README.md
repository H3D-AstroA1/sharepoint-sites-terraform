# M365 Environment Population Tool

## 📋 Overview

This project provides a comprehensive toolkit for populating **Microsoft 365 environments** with realistic test data. It automates the creation and population of:

- **SharePoint Online Sites** - Create team sites, communication sites, and document libraries
- **Email Mailboxes** - Populate user mailboxes with realistic emails
- **Files & Documents** - Generate department-appropriate documents across sites

### Key Capabilities

| Feature | Description |
|---------|-------------|
| **SharePoint Site Deployment** | Create sites using Terraform with configurable templates |
| **File Population** | Generate realistic documents (Word, Excel, PDF, PowerPoint) |
| **Email Population** | Send realistic emails to user mailboxes |
| **Azure AD Discovery** | Discover users and groups from your tenant |
| **User/Domain Exclusions** | Exclude specific emails or domains from operations |
| **Deployment Tracking** | Track deployments with unique IDs for easy cleanup |

### Five SharePoint Deployment Modes

| Mode | Description | Use Case |
|------|-------------|----------|
| **Configuration File Only** | Define custom site names in a JSON file | When you know exactly which sites you need |
| **Config + Ad-hoc Sites** | Your custom sites + random ad-hoc sites | Realistic environments with your specific sites |
| **Department Sites** | Official department sites (HR, Finance, IT, etc.) | Organizational structure simulation |
| **Ad-hoc Sites** | User-created sites (projects, teams, events) | Organic SharePoint usage simulation |
| **Mixed Sites** | Combination of department + ad-hoc sites | Maximum realism |

> ⚠️ **Important Note**: SharePoint Online and Exchange Online are Microsoft 365 services, not Azure resources. This solution uses Azure resources (Resource Group, Key Vault) for supporting infrastructure, while SharePoint sites are created via Microsoft Graph API and emails are sent via Exchange.

---

## 📁 Project Structure

```
sharepoint-sites-terraform/
├── README.md                    # This file - main documentation
├── CONFIGURATION-GUIDE.md       # Detailed configuration instructions
├── PREREQUISITES.md             # What you need before starting
├── requirements.txt             # Python dependencies
├── .gitignore                   # Git ignore rules
│
├── config/                      # Configuration files
│   ├── sites.json               # ⭐ EDIT THIS FILE for custom sites
│   ├── mailboxes.yaml           # ⭐ EDIT THIS FILE for email mailboxes
│   └── environments.json        # ⭐ PRE-CONFIGURE YOUR ENVIRONMENTS HERE
│
├── terraform/                   # Terraform configuration files
│   ├── main.tf                  # Main Terraform configuration
│   ├── variables.tf             # Variable definitions
│   ├── outputs.tf               # Output definitions
│   ├── providers.tf             # Provider configurations
│   └── terraform.tfvars.example # Example variable values
│
├── scripts/                     # Automation scripts
│   ├── menu.py                  # ⭐ MAIN MENU (START HERE!)
│   ├── deploy.py                # Create SharePoint sites
│   ├── populate_files.py        # Populate sites with files
│   ├── populate_emails.py       # Populate mailboxes with emails
│   ├── cleanup.py               # Delete files/sites
│   ├── cleanup_emails.py        # Delete emails from mailboxes
│   └── email_generator/         # Email generation modules
│
└── docs/                        # Additional documentation
    ├── TROUBLESHOOTING.md       # Common issues and solutions
    └── EMAIL_POPULATION.md      # Email population tool guide
```

---

## 🚀 Quick Start Guide

### Easiest Way: Use the Main Menu

```bash
cd sharepoint-sites-terraform/scripts
python menu.py
```

This opens an interactive menu that guides you through all operations:

```
╔══════════════════════════════════════════════════════════════╗
║   M365 Environment Population Tool                           ║
╚══════════════════════════════════════════════════════════════╝

  [0] ✓ Check & Install Prerequisites    ← START HERE!
      Azure CLI, Terraform, Azure Login, PyYAML

  ────────────────────────────────────────────────────────────
  SHAREPOINT OPERATIONS:
  [1] 🏗️  Create SharePoint Sites
  [2] 📄 Populate Sites with Files
  [3] 🗑️  Delete Files or Sites
  [4] 📋 List SharePoint Sites
  [5] 📁 List Files in Sites

  ────────────────────────────────────────────────────────────
  EMAIL OPERATIONS:
  [6] 📧 Populate Mailboxes with Emails
  [7] 🗑️  Delete Emails from Mailboxes
  [8] 📬 List Mailboxes

  ────────────────────────────────────────────────────────────
  DISCOVERY & MANAGEMENT:
  [9] 🔍 Azure AD User Discovery
  [A] 🔐 Manage App Registration         ← For SharePoint/Mail permissions
  [R] 🧹 Remove Excluded Users from Sites

  [C] ⚙️  Edit Configuration
  [H] ❓ Help & Documentation
  [Q] 🚪 Quit
```

### Workflow Overview

| Category | Steps | Menu Options |
|----------|-------|--------------|
| **Setup** | Check prerequisites, configure app registration | `[0]`, `[A]` |
| **SharePoint** | Create sites → Populate files → Cleanup | `[1]`, `[2]`, `[3]` |
| **Email** | Configure mailboxes → Populate emails → Cleanup | `[6]`, `[7]`, `[8]` |
| **Discovery** | Discover Azure AD users and groups | `[9]` |

---

## 📧 Email Population

The tool can populate user mailboxes with realistic emails for testing and demonstration purposes.

### Configuration

Edit `config/mailboxes.yaml` to define your mailboxes and email settings:

```yaml
# Mailbox configuration
mailboxes:
  - email: user1@contoso.onmicrosoft.com
    display_name: John Smith
    department: IT
    email_count: 50
    
  - email: user2@contoso.onmicrosoft.com
    display_name: Jane Doe
    department: Finance
    email_count: 30

# Email generation settings
settings:
  default_email_count: 25
  date_range_days: 90
  include_attachments: true
  attachment_probability: 0.3

# Exclusions - emails/domains to skip
exclusions:
  enabled: true
  email_addresses:
    - admin@contoso.onmicrosoft.com
    - service@contoso.onmicrosoft.com
  domains:
    - external.com
  allowed_domains:  # Whitelist - only these domains are processed
    - contoso.onmicrosoft.com
```

### Running Email Population

```bash
# From the main menu
python menu.py
# Select [6] Populate Mailboxes with Emails

# Or run directly
python populate_emails.py
```

### Email Features

| Feature | Description |
|---------|-------------|
| **Realistic Content** | Department-appropriate email subjects and bodies |
| **Attachments** | Optional file attachments (configurable probability) |
| **Date Distribution** | Emails spread across configurable date range |
| **CC/BCC Support** | Realistic email threading with CC/BCC recipients |
| **Exclusions** | Skip specific addresses or domains |

---

## 🏗️ SharePoint Site Deployment

### Quick Start

```bash
cd sharepoint-sites-terraform/scripts

# Interactive mode (recommended)
python deploy.py

# Or use command line options
python deploy.py --random 10 --auto-approve
```

### Deployment Modes

When you run `python deploy.py`, you'll see this menu:

```
How would you like to define your SharePoint sites?

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CONFIGURATION FILE OPTIONS:

  [1] Use Configuration File Only (21 sites)
      - Uses config/sites.json for your custom site names
      - Full control over site names, descriptions, and settings

  [2] Configuration File + Ad-hoc Sites (21 config + you choose ad-hoc)
      - Uses your custom sites from config/sites.json
      - PLUS you choose how many ad-hoc sites (0-60 available)
      - Best for realistic environments with your specific sites

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RANDOM GENERATION OPTIONS:

  [3] Generate Department Sites (40 templates)
  [4] Generate Ad-hoc Sites (60 templates)
  [5] Generate Mixed Sites (Department + Ad-hoc)
```

### Site Configuration

Edit `config/sites.json` to define your sites:

```json
{
  "sites": [
    {
      "name": "executive-confidential",
      "display_name": "Executive Confidential",
      "description": "Confidential documents for executive team",
      "visibility": "Private"
    }
  ],
  "exclusions": {
    "enabled": true,
    "email_addresses": ["admin@contoso.com"],
    "domains": ["external.com"]
  },
  "deployment_tracking": {
    "enabled": true,
    "deployment_id": "auto"
  }
}
```

### Deployment Tracking

Sites are tagged with a unique deployment ID for easy tracking and cleanup:

- **Auto-generated ID**: Format `YYYYMMDD-HHMMSS-XXXX` (e.g., `20240315-143022-A7B3`)
- **Custom ID**: Specify your own ID in `sites.json`
- **Cleanup by ID**: Filter cleanup operations to only affect sites from a specific deployment

---

## 📄 File Population

After creating SharePoint sites, populate them with realistic documents.

### Quick Start

```bash
# Interactive mode
python populate_files.py

# Create 100 files distributed across all sites
python populate_files.py --files 100

# Create files only in HR-related sites
python populate_files.py --files 50 --site hr
```

### File Types Generated

| Department | Example Files |
|------------|---------------|
| **Executive** | Board Meeting Agenda, Strategic Plan, CEO Town Hall Presentation |
| **HR** | Employee Handbook, Job Descriptions, Onboarding Checklist |
| **Finance** | Monthly Financial Report, Annual Budget, Cash Flow Statement |
| **IT** | System Architecture Diagram, Security Policy, Disaster Recovery Plan |
| **Marketing** | Marketing Plan, Brand Guidelines, Campaign Performance |
| **Legal** | NDA Templates, Compliance Checklist, Privacy Policy |

### Realism Variation Levels

```bash
python populate_files.py --files 200 --variation-level high
```

- `low`: Lighter naming/folder variation (quick tests)
- `medium`: Balanced realism (default)
- `high`: Maximum variation for enterprise simulation

---

## 🔍 Azure AD Discovery

Discover users and groups from your Azure AD tenant for realistic site ownership and email population.

### Features

| Feature | Description |
|---------|-------------|
| **User Discovery** | Find all member users (filters out guests/service accounts) |
| **Group Discovery** | Get security groups and M365 groups |
| **Mailbox Validation** | Verify which users have Exchange mailboxes |
| **Department Grouping** | View users organized by department |
| **Caching** | Results cached for 60 minutes to reduce API calls |

### Mailbox Validation Options

When running Azure AD discovery, you can choose validation speed:

```
Mailbox Validation Options:
  [1] Skip mailbox validation (fastest)
  [2] Validate first 100 mailboxes (default)
  [3] Validate first 500 mailboxes
  [4] Validate all mailboxes (slow for large tenants)
```

---

## 🗑️ Cleanup Operations

### SharePoint Cleanup

```bash
# Interactive mode
python cleanup.py

# Delete all files from sites (keeps sites)
python cleanup.py --delete-files

# Delete SharePoint sites
python cleanup.py --delete-sites

# Filter by deployment ID
python cleanup.py --delete-sites --deployment-id 20240315-143022-A7B3
```

### Email Cleanup

```bash
# Interactive mode
python cleanup_emails.py

# Delete all emails from configured mailboxes
python cleanup_emails.py --all

# Delete emails from specific mailbox
python cleanup_emails.py --mailbox user@contoso.com
```

### Recycle Bin Purge

Deleted items go to recycle bins. To permanently remove:

```bash
# Purge M365 Groups from Azure AD recycle bin
python cleanup.py --purge-deleted

# Purge SharePoint sites from SharePoint recycle bin
python cleanup.py --purge-spo-recycle --tenant contoso

# Purge site document library recycle bins
python cleanup.py --purge-site-recycle --non-interactive --yes
```

---

## 🔐 App Registration

The scripts require Microsoft Graph API permissions. Use the built-in App Registration feature for easy setup.

### Setting Up

```bash
python menu.py
# Press [A] to open App Registration menu
```

```
╔══════════════════════════════════════════════════════════════╗
║   App Registration Management                                 ║
╚══════════════════════════════════════════════════════════════╝

  [1] 🆕 Create new app registration
  [2] 🔑 Grant admin consent (opens browser)
  [3] ✅ Check current app status
  [4] 🔄 Update app permissions
  [5] 🔑 Regenerate client secret
  [6] 🗑️  Delete app registration

  [B] ← Back to main menu
```

### Permissions Granted

| Permission | Purpose |
|------------|---------|
| `Sites.Read.All` | List SharePoint sites |
| `Sites.ReadWrite.All` | Create and modify sites |
| `Files.ReadWrite.All` | Upload and delete files |
| `Group.Read.All` | List Microsoft 365 Groups |
| `Group.ReadWrite.All` | Create and delete Groups |
| `User.Read.All` | Discover Azure AD users |
| `Mail.ReadWrite` | Read and write mailbox emails |
| `Mail.Send` | Send emails on behalf of users |

---

## 🚫 Exclusions

Exclude specific users or domains from all operations.

### Email Exclusions (mailboxes.yaml)

```yaml
exclusions:
  enabled: true
  email_addresses:
    - admin@contoso.onmicrosoft.com
    - service@contoso.onmicrosoft.com
  domains:
    - external.com
  allowed_domains:
    - contoso.onmicrosoft.com
```

### Site Exclusions (sites.json)

```json
{
  "exclusions": {
    "enabled": true,
    "email_addresses": ["admin@contoso.com"],
    "domains": ["external.com"],
    "patterns": ["*-test@*"]
  }
}
```

### Remove Excluded Users from Sites

If excluded users were previously added as site owners/members:

```bash
python menu.py
# Select [R] Remove Excluded Users from Sites
```

---

## 🌍 Pre-Configured Environments

Edit `config/environments.json` to save your Azure/M365 settings:

```json
{
  "environments": [
    {
      "name": "Default",
      "description": "Default environment",
      "azure": {
        "tenant_id": "YOUR-AZURE-TENANT-ID",
        "subscription_id": "YOUR-AZURE-SUBSCRIPTION-ID",
        "resource_group": "rg-sharepoint-sites",
        "location": "westus2"
      },
      "m365": {
        "tenant_name": "your-tenant",
        "admin_email": "admin@your-tenant.onmicrosoft.com"
      }
    }
  ],
  "default_environment": "Default"
}
```

---

## 🔧 Command Line Reference

### deploy.py

```bash
python deploy.py --help

Options:
  --config FILE        Use custom config file
  --random COUNT       Generate random sites (1-39)
  --environment NAME   Use pre-configured environment
  --skip-prerequisites Skip prerequisite checks
  --auto-approve       Skip Terraform confirmation
```

### populate_files.py

```bash
python populate_files.py --help

Options:
  -f, --files COUNT       Number of files to create (1-1000)
  -s, --site FILTER       Filter sites by name
  --variation-level LEVEL Realism: low, medium, high
  -l, --list-sites        List available sites
```

### cleanup.py

```bash
python cleanup.py --help

Options:
  --delete-files        Delete all files from sites
  --delete-sites        Delete SharePoint sites
  --delete-all          Delete both files and sites
  -s, --site FILTER     Filter sites by name
  --deployment-id ID    Filter by deployment ID
  --purge-deleted       Purge M365 Groups recycle bin
  --purge-spo-recycle   Purge SharePoint recycle bin
  --tenant NAME         SharePoint tenant name
  -y, --yes             Skip confirmation prompts
```

### populate_emails.py

```bash
python populate_emails.py --help

Options:
  --config FILE         Use custom mailboxes.yaml
  --mailbox EMAIL       Populate specific mailbox only
  --count COUNT         Override email count
  --dry-run             Preview without sending
```

### cleanup_emails.py

```bash
python cleanup_emails.py --help

Options:
  --all                 Delete from all configured mailboxes
  --mailbox EMAIL       Delete from specific mailbox
  --folder FOLDER       Target specific folder (inbox, sent, etc.)
  -y, --yes             Skip confirmation prompts
```

---

## 📖 Additional Documentation

| Document | Description |
|----------|-------------|
| [CONFIGURATION-GUIDE.md](./CONFIGURATION-GUIDE.md) | Step-by-step configuration instructions |
| [PREREQUISITES.md](./PREREQUISITES.md) | Detailed prerequisites and setup |
| [docs/TROUBLESHOOTING.md](./docs/TROUBLESHOOTING.md) | Common issues and solutions |
| [docs/EMAIL_POPULATION.md](./docs/EMAIL_POPULATION.md) | Email population tool guide |
| [plans/SHAREPOINT_ARCHITECTURE.md](./plans/SHAREPOINT_ARCHITECTURE.md) | SharePoint deployment architecture & implementation |
| [plans/EMAIL_POPULATION_ARCHITECTURE.md](./plans/EMAIL_POPULATION_ARCHITECTURE.md) | Email population architecture & implementation |

---

## 🔐 Security Considerations

1. **Never commit `terraform.tfvars`** - Contains sensitive information
2. **Never commit `config/` files with real data** - May contain emails and tenant info
3. **Use Azure Key Vault** - For storing secrets in production
4. **Review before applying** - Always check the Terraform plan
5. **Use exclusions** - Protect admin and service accounts

---

## 🐍 About the Python Scripts

The scripts are written in Python for cross-platform compatibility:

- **Works on**: Windows, macOS, Linux
- **Python version**: 3.8 or higher
- **Dependencies**: PyYAML (auto-installed), exchangelib (for email)
- **Auto-installs**: Azure CLI and Terraform if not present

### Automatic Prerequisite Installation

| Platform | Azure CLI | Terraform |
|----------|-----------|-----------|
| **Windows** | winget → chocolatey → manual | winget → chocolatey → manual |
| **macOS** | Homebrew | Homebrew |
| **Linux** | apt package manager | apt package manager |

---

## 📞 Support

If you encounter issues:

1. Check [TROUBLESHOOTING.md](./docs/TROUBLESHOOTING.md)
2. Ensure all prerequisites are met (`python --version`, `az --version`, `terraform --version`)
3. Verify your permissions in both Azure and Microsoft 365
4. Check the Terraform logs for detailed error messages

---

## 📝 License

This project is provided as-is for educational and automation purposes.
