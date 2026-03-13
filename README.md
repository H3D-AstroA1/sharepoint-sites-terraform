# SharePoint Sites Deployment with Terraform

## 📋 Overview

This project automates the creation of **SharePoint Online sites** using Terraform. It provides flexible options for defining which sites to create:

### Five Deployment Modes

| Mode | Description | Use Case |
|------|-------------|----------|
| **Configuration File Only** | Define custom site names in a JSON file | When you know exactly which sites you need |
| **Config + Ad-hoc Sites** | Your custom sites + random ad-hoc sites | Realistic environments with your specific sites |
| **Department Sites** | Official department sites (HR, Finance, IT, etc.) | Organizational structure simulation |
| **Ad-hoc Sites** | User-created sites (projects, teams, events) | Organic SharePoint usage simulation |
| **Mixed Sites** | Combination of department + ad-hoc sites | Maximum realism |

### Key Features

- **Azure AD User/Group Discovery** - Automatically discover real users and groups from your tenant
- **Random Owner Assignment** - Assign discovered users as site owners/members for realism
- **Key Vault Configuration** - Optional custom Key Vault name or auto-generated
- **60+ Ad-hoc Site Templates** - Projects, teams, events, clubs, regional offices
- **40 Department Site Templates** - HR, Finance, IT, Legal, Marketing, and more

> ⚠️ **Important Note**: SharePoint Online is a Microsoft 365 service, not an Azure resource. This solution uses Azure resources (Resource Group, Key Vault) for supporting infrastructure, while SharePoint sites are created via Microsoft Graph API.

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
│   ├── deploy.py                # Step 1: Create SharePoint sites
│   ├── populate_files.py        # Step 2: Populate sites with files
│   ├── populate_emails.py       # Populate mailboxes with emails
│   ├── cleanup.py               # Step 3: Delete files/sites
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
  [1] 🏗️  Create SharePoint Sites
  [2] 📄 Populate Sites with Files
  [3] 🗑️  Delete Files or Sites

  ────────────────────────────────────────────────────────────
  [4] 📋 List SharePoint Sites
  [5] 📁 List Files in Sites
  [6] 📧 Populate Mailboxes with Emails
  [7] 🗑️  Delete Emails from Mailboxes
  [8] 📬 List Mailboxes

  [A] 🔐 Manage App Registration         ← For SharePoint/Mail permissions
  [C] ⚙️  Edit Configuration
  [H] ❓ Help & Documentation
  [Q] 🚪 Quit
```

### Workflow Steps

| Step | Description | Menu Option |
|------|-------------|-------------|
| **Step 0** | Check & install prerequisites (Azure CLI, Terraform, login) | `[0]` |
| **Step 1** | Create SharePoint sites using Terraform | `[1]` |
| **Step 2** | Populate sites with realistic files | `[2]` |
| **Step 3** | Delete files or sites when done | `[3]` |

---

### Alternative: Run Scripts Directly

### Step 1: Prerequisites

The deployment script **automatically checks and installs** the required tools:

| Requirement | Auto-Install? | Notes |
|-------------|---------------|-------|
| **Python 3.8+** | ❌ Manual | Must be installed first to run the script |
| **Azure CLI** | ✅ Yes | Script will install if missing |
| **Terraform** | ✅ Yes | Script will install if missing |
| **Azure Subscription** | ❌ Manual | Check in Azure Portal |
| **Microsoft 365 Tenant** | ❌ Manual | With SharePoint Online license |
| **SharePoint Admin Role** | ❌ Manual | Check your role in Entra ID |

> 💡 **Only Python is required to start!** The script will detect and install Azure CLI and Terraform automatically on Windows (winget/chocolatey), macOS (Homebrew), and Linux (apt).

### Step 2: Choose Your Deployment Mode

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
      - Official department sites (HR, Finance, IT, Legal, etc.)
      - Realistic organizational structure
      - Mix of Private and Public visibility

  [4] Generate Ad-hoc Sites (60 templates)
      - User-created sites (projects, teams, events, clubs)
      - Simulates organic SharePoint usage by employees
      - Includes working groups, social clubs, regional offices

  [5] Generate Mixed Sites (Department + Ad-hoc)
      - Combines both types for maximum realism
      - Specify count for each type separately
```

#### Option A: Use Configuration File (Custom Site Names)

1. **Edit the configuration file** `config/sites.json`:

```json
{
  "sites": [
    {
      "name": "executive-confidential",
      "display_name": "Executive Confidential",
      "description": "Confidential documents for executive team"
    },
    {
      "name": "finance-internal",
      "display_name": "Finance Internal",
      "description": "Finance team documents"
    }
  ]
}
```

2. **Run the deployment and select option [1]**

#### Option B: Generate Department Sites

Sites will be created with realistic organizational department names like:
- `human-resources` - HR policies, employee handbook, benefits (Private)
- `finance-department` - Financial reports, budgets, accounting (Private)
- `it-helpdesk` - IT support documentation, troubleshooting guides (Public)
- `employee-intranet` - Central hub for all employees (Public, Communication site)
- `legal-compliance` - Legal documents, contracts, regulatory compliance (Private)

> 📊 **Department Sites**: 40 templates including 25 Private sites (64%) and 15 Public sites (36%)

#### Option C: Generate Ad-hoc Sites

User-created sites that simulate organic SharePoint usage:
- `q4-product-launch-2024` - Project team site (Private)
- `innovation-lab` - Working group for new ideas (Public)
- `coffee-club` - Social interest group (Public)
- `london-office` - Regional office site (Private)
- `hackathon-2024` - Event site (Public)

> 📊 **Ad-hoc Sites**: 60 templates including projects, teams, events, clubs, and regional offices

### Step 3: Follow the Interactive Prompts

The script will guide you through:

1. ✅ **Prerequisite check & auto-install** (Azure CLI, Terraform)
2. ✅ Site generation mode selection
3. ✅ Azure authentication
4. ✅ Environment selection (or manual tenant/subscription)
5. ✅ Resource group configuration (new or existing)
6. ✅ Key Vault configuration (new or existing)
7. ✅ Microsoft 365 settings
8. ✅ **Owner/Member Assignment** (NEW!)
9. ✅ Configuration review
10. ✅ Terraform deployment

> 💡 **Tip**: You can press `Q` to quit or `Ctrl+C` to cancel at any interactive prompt.

### Step 4: Configure Site Owners & Members (NEW!)

After selecting your sites, you'll be prompted to assign owners and members:

```
How would you like to assign site owners and members?

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  [1] Admin Only (Default)
      - Uses the SharePoint admin email as the sole owner
      - No additional members
      - Simplest option, good for testing

  [2] Discover Azure AD Users & Groups
      - Queries Microsoft Graph API for real users/groups
      - Randomly assigns users as owners (1-3 per site)
      - Randomly assigns users/groups as members (0-5 per site)
      - Most realistic option for production-like environments

  [3] Skip Owner Assignment
      - Leave owners/members empty
      - Sites will be created with default permissions
```

#### Azure AD Discovery Features

When you select option [2], the script will:

1. **Discover Users** - Query Microsoft Graph for member users (filters out guests/service accounts)
2. **Discover Groups** - Get security groups and M365 groups (filters out dynamic groups)
3. **Random Assignment** - Assign 1-3 random users as owners and 0-5 as members per site
4. **Optional Group Members** - Choose whether to include security groups as site members

Example output:
```
Discovered Azure AD Identities:

  Users (sample):
    • John Smith (john.smith@contoso.com) - IT Department
    • Jane Doe (jane.doe@contoso.com) - Finance
    • Bob Wilson (bob.wilson@contoso.com) - HR
    ... and 47 more

  Groups (sample):
    • IT-Admins (Security)
    • Finance-Team (M365)
    • All-Employees (Security)
    ... and 12 more

Include groups as site members? (y/N): y

Assigned 87 owners and 142 members across 21 sites
```

---

## 🔧 Command Line Options

```bash
# Show help
python deploy.py --help

# Interactive mode (recommended for first-time users)
python deploy.py

# Use custom config file
python deploy.py --config ./my-sites.json

# Generate random sites (specify count 1-39)
python deploy.py --random 15

# Use a pre-configured environment (from config/environments.json)
python deploy.py --environment Production

# Skip prerequisites check
python deploy.py --skip-prerequisites

# Auto-approve Terraform apply (no confirmation prompt)
python deploy.py --auto-approve

# Combine options
python deploy.py --random 5 --environment Development --auto-approve
```

### Common Examples

```bash
# First-time deployment (interactive)
python deploy.py

# Create 10 random test sites quickly
python deploy.py --random 10 --auto-approve

# Use your own sites configuration
python deploy.py --config /path/to/my-sites.json

# Deploy to Production environment with 5 random sites
python deploy.py --environment Production --random 5 --auto-approve

# Quick deployment skipping all prompts
python deploy.py --random 5 --skip-prerequisites --auto-approve
```

---

## 📄 File Population (Optional Step 2)

After creating SharePoint sites, you can populate them with realistic-looking files to simulate an actual organization's document structure.

### Quick Start

```bash
cd sharepoint-sites-terraform/scripts

# Interactive mode
python populate_files.py

# Create 100 files distributed across all sites
python populate_files.py --files 100

# Create 50 files only in HR-related sites
python populate_files.py --files 50 --site hr

# List available sites
python populate_files.py --list-sites
```

### File Types Generated

The script creates realistic files based on each site's department type:

| Department | Example Files |
|------------|---------------|
| **Executive** | Board Meeting Agenda, Strategic Plan, CEO Town Hall Presentation |
| **HR** | Employee Handbook, Job Descriptions, Onboarding Checklist, Benefits Summary |
| **Finance** | Monthly Financial Report, Annual Budget, Cash Flow Statement, Expense Reports |
| **Claims** | Claim Form Template, Claims Processing Guide, Claim Status Tracker, Settlement Authorization |
| **IT** | System Architecture Diagram, Security Policy, Disaster Recovery Plan |
| **Marketing** | Marketing Plan, Brand Guidelines, Campaign Performance, Social Media Calendar |
| **Sales** | Sales Proposals, Contracts, Pipeline Analysis, Territory Maps |
| **Legal** | NDA Templates, Compliance Checklist, Privacy Policy, Terms of Service |
| **Operations** | SOPs, Inventory Reports, Facility Maintenance, Quality Control |

### File Distribution

Files are randomly distributed across sites with:
- **Realistic folder structures** (e.g., "Board Materials", "Policies", "Reports")
- **Department-appropriate naming** (e.g., "Q3_Board_Meeting_Agenda_2024.docx")
- **Various file types** (Word, Excel, PowerPoint, PDF)
- **Random dates and version numbers** for authenticity

### Command Line Options

```bash
python populate_files.py --help

Options:
  -f, --files COUNT    Number of files to create (1-1000)
  -s, --site FILTER    Filter sites by name (e.g., "hr", "finance")
  -l, --list-sites     List available SharePoint sites and exit
```

---

## 🗑️ Cleanup Script (Optional)

When you need to reset your environment or remove test data, use the cleanup script.

> ⚠️ **WARNING**: This script performs DESTRUCTIVE operations. Deleted files and sites may not be recoverable!

### Quick Start

```bash
cd sharepoint-sites-terraform/scripts

# Interactive mode (safest - shows menu with all options)
python cleanup.py

# Delete all files from all sites (keeps sites)
python cleanup.py --delete-files

# Delete files from specific sites only
python cleanup.py --delete-files --site hr

# Delete SharePoint sites (requires admin permissions)
python cleanup.py --delete-sites

# Delete everything (files and sites)
python cleanup.py --delete-all

# List sites without deleting
python cleanup.py --list-sites

# List files in all sites
python cleanup.py --list-files

# List files in a SPECIFIC site (combine --list-files with --site)
python cleanup.py --list-files --site hr
python cleanup.py --list-files --site finance
python cleanup.py --list-files --site "executive"
```

### Interactive Selection (NEW!)

The cleanup script now supports **interactive selection** of specific sites and files:

```bash
# Interactively select specific sites from a numbered list
python cleanup.py --select-sites

# Interactively select specific files to delete
python cleanup.py --select-files

# Combine with site filter
python cleanup.py --select-files --site hr
```

#### Selection Syntax

When prompted to select items, you can use:

| Syntax | Example | Description |
|--------|---------|-------------|
| Single | `1` | Select item 1 |
| Multiple | `1,3,5` | Select items 1, 3, and 5 |
| Range | `1-5` | Select items 1 through 5 |
| Combined | `1,3,5-10` | Select 1, 3, and 5 through 10 |
| All | `*` | Select all items |

#### Interactive Mode Menu

When you run `python cleanup.py` without flags, you'll see:

```
What would you like to do?

  [1] Delete all FILES from sites (keeps sites)
  [2] Delete SITES (and all their content)
  [3] Delete BOTH files and sites
  [4] Select SPECIFIC sites to work with
  [5] Select SPECIFIC files to delete
  [6] List files in sites
  [7] Cancel
```

### Recycle Bin Purge (NEW!)

When you delete SharePoint sites, they go to **two separate recycle bins**:

1. **Microsoft 365 Groups Recycle Bin** (Azure AD) - Where the M365 Group is soft-deleted
2. **SharePoint Site Recycle Bin** (SharePoint Admin Center) - Where the site appears in "Deleted Sites"

#### Automatic Purge After Deletion

When you delete sites using option [5] or the `--delete-sites` flag, the script will **automatically offer to purge both recycle bins**:

```
Sites have been soft-deleted. They now exist in two recycle bins:
    1. Microsoft 365 Groups recycle bin (Azure AD)
    2. SharePoint site recycle bin (SharePoint Admin Center)

Purge recycle bins now? (Y/n):
```

- Press **Enter** or **Y**: Proceeds to purge both recycle bins
- Press **n**: Skips purge, sites remain in recycle bins for later cleanup

#### Manual Purge Commands

You can also manually purge the recycle bins at any time:

```bash
# Purge M365 Groups from Azure AD recycle bin
python cleanup.py --purge-deleted

# Purge SharePoint sites from SharePoint Admin Center recycle bin
# Replace 'contoso' with your tenant name (e.g., contoso.sharepoint.com)
python cleanup.py --purge-spo-recycle --tenant contoso
```

Or use the main menu:
- Option **[6]**: Purge M365 Groups recycle bin (Azure AD)
- Option **[7]**: Purge SharePoint site recycle bin

> ⚠️ **Note**: The SharePoint site recycle bin purge requires the **SharePoint Online PowerShell module** (`Microsoft.Online.SharePoint.PowerShell`). The script will automatically install it if not present (Windows only).

### Command Line Options

```bash
python cleanup.py --help

Options:
  --delete-files       Delete all files from SharePoint sites
  --delete-sites       Delete SharePoint sites (requires admin permissions)
  --delete-all         Delete both files and sites
  -s, --site FILTER    Filter sites by name (e.g., "hr", "finance")
  -l, --list-sites     List available SharePoint sites and exit
  --list-files         List files in SharePoint sites
  --select-sites       Interactively select specific sites from a numbered list
  --select-files       Interactively select specific files to delete
  --purge-deleted      Permanently delete M365 Groups from Azure AD recycle bin
  --purge-spo-recycle  Permanently delete sites from SharePoint recycle bin
  --tenant NAME        SharePoint tenant name (required with --purge-spo-recycle)
  -y, --yes            Skip confirmation prompts (use with caution!)
```

### Safety Features

| Feature | Description |
|---------|-------------|
| **Double confirmation** | Site deletion requires typing "DELETE" and "YES" |
| **Site filtering** | Only affect specific sites with `--site` flag |
| **Interactive selection** | Choose exactly which sites/files to delete |
| **List mode** | Preview sites/files before deleting |
| **Recycle bin** | Deleted sites go to SharePoint recycle bin for 93 days |
| **Auto-purge prompt** | After deletion, prompts to purge recycle bins (can skip with 'n') |
| **Manual purge** | Permanently remove items from both recycle bins on demand |

### Permissions Required

| Operation | Required Permission |
|-----------|---------------------|
| Delete files | Sites.ReadWrite.All, Files.ReadWrite.All |
| Delete sites | Sites.FullControl.All (SharePoint Admin) |

---

## 🔐 App Registration (Recommended)

The scripts require Microsoft Graph API permissions to access SharePoint sites. The easiest way to set this up is using the built-in App Registration feature.

### Why Use App Registration?

| Without App Registration | With App Registration |
|--------------------------|----------------------|
| Azure CLI token may lack permissions | Custom app with all required permissions |
| May get 403 Forbidden errors | Full access to SharePoint sites |
| Limited to current user's permissions | Works with client credentials flow |
| Requires manual admin consent | Guided admin consent process |

### Setting Up App Registration

```bash
cd sharepoint-sites-terraform/scripts
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
  [4] 🗑️  Delete app registration

  [B] ← Back to main menu
```

### Permissions Granted

The custom app is configured with these Microsoft Graph permissions:

| Permission | Purpose |
|------------|---------|
| `Sites.Read.All` | List SharePoint sites |
| `Sites.ReadWrite.All` | Create and modify sites |
| `Files.ReadWrite.All` | Upload and delete files |
| `Group.Read.All` | List Microsoft 365 Groups |
| `Group.ReadWrite.All` | Create and delete Groups |

### Admin Consent Required

After creating the app, a tenant administrator must grant consent:
1. The script opens a browser to the Azure consent page
2. Sign in with an admin account
3. Click "Accept" to grant permissions
4. Return to the script and press Enter

> 💡 **Tip**: If you're not an admin, ask your IT administrator to grant consent using the URL provided by the script.

---

## 🌍 Pre-Configured Environments

For easier deployment, you can pre-configure your Azure environment in `config/environments.json`. This saves you from entering tenant/subscription IDs manually each time.

### Setting Up Your Environment

Edit `config/environments.json` and replace the placeholder values:

```json
{
  "environments": [
    {
      "name": "Default",
      "description": "Default environment - edit with your Azure and M365 settings",
      "azure": {
        "tenant_id": "YOUR-AZURE-TENANT-ID",
        "tenant_name": "Your Organization",
        "subscription_id": "YOUR-AZURE-SUBSCRIPTION-ID",
        "subscription_name": "Your Subscription",
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

### Automatic Environment Detection

All scripts (`deploy.py`, `populate_files.py`, `cleanup.py`) automatically detect and use your configured environments:

| Scenario | Behavior |
|----------|----------|
| **1 environment configured** | Automatically uses it (no prompts) |
| **Multiple environments** | Prompts you to select which one |
| **No environments configured** | Falls back to current Azure CLI login |

### Using Environments

```bash
# Interactive mode - will offer your configured environment
python deploy.py

# Specify environment directly
python deploy.py --environment Default

# Combine with other options
python deploy.py --environment Default --random 10 --auto-approve
```

> 💡 **Tip**: To add more environments (e.g., Production, Development), copy the environment block in the JSON file and modify the values. The file includes an example showing how to do this.

### Edit Configuration from Menu

Press `[C]` in the main menu to access the configuration editor:

```
╔══════════════════════════════════════════════════════════════╗
║   Edit Configuration                                         ║
╚══════════════════════════════════════════════════════════════╝

  [1] 📝 Edit environments.json (in VS Code)
  [2] 📝 Edit sites.json (in VS Code)
  [3] 👁️  View environments.json
  [4] 👁️  View sites.json
  [5] ➕ Add new environment (wizard)

  [B] ← Back to main menu
```

| Option | Description |
|--------|-------------|
| **Edit in VS Code** | Opens the file in VS Code (or system default editor) |
| **View** | Displays the file contents in the terminal |
| **Add new environment** | Interactive wizard to add a new environment |

The wizard prompts for:
- Environment name (e.g., "Production", "Development")
- Azure Tenant ID
- Azure Subscription ID
- Resource Group name
- M365 Domain (e.g., "contoso.onmicrosoft.com")

---

## 📝 Configuration File Format

Edit `config/sites.json` to define your sites:

```json
{
  "sites": [
    {
      "name": "site-url-name",
      "display_name": "Site Display Name",
      "description": "Description of the site",
      "template": "STS#3",
      "visibility": "Private",
      "owners": ["owner@tenant.onmicrosoft.com"],
      "members": ["member@tenant.onmicrosoft.com"]
    }
  ]
}
```

| Field | Required | Description |
|-------|----------|-------------|
| `name` | Yes | URL-friendly name (no spaces, use hyphens) |
| `display_name` | Yes | Human-readable name |
| `description` | No | Site description |
| `template` | No | Site template (default: STS#3) |
| `visibility` | No | Private or Public (default: Private) |
| `owners` | No | Array of owner email addresses |
| `members` | No | Array of member email addresses |

### Site Templates

| Template | Description |
|----------|-------------|
| `STS#3` | Team site (no Microsoft 365 Group) - **Default** |
| `GROUP#0` | Team site (with Microsoft 365 Group) |
| `SITEPAGEPUBLISHING#0` | Communication site |

---

## ⚠️ Important: Random Site Behavior

### Maximum Random Sites

The random generation mode has a **maximum of 39 unique sites** (the number of department templates available). Each template includes realistic visibility settings:

| Category | Count | Visibility | Examples |
|----------|-------|------------|----------|
| Executive & Leadership | 3 | Private | Board of Directors, Senior Management |
| Human Resources | 3 | Private | HR Recruitment, Payroll & Benefits |
| Finance & Accounting | 3 | Private | Treasury, Accounts Payable |
| IT Department | 3 | Mixed | IT Security (Private), IT Helpdesk (Public) |
| Company-wide Resources | 10 | Public | Employee Intranet, Training, Announcements |
| Other Departments | 17 | Mixed | Legal, Marketing, Sales, Operations |

### Running the Script Multiple Times

| Scenario | What Happens |
|----------|--------------|
| Same random count twice | Different random selection each time |
| `--random 39` twice | Safe - all sites selected both times |
| `--random 10` after `--random 39` | Terraform "forgets" 29 sites (but they still exist in SharePoint) |

> ⚠️ **Important**: SharePoint sites are NOT automatically deleted when removed from Terraform. The `null_resource` approach only creates sites - it doesn't manage their lifecycle. To delete sites, use the SharePoint Admin Center.

### Best Practices

| Use Case | Recommendation |
|----------|----------------|
| **Testing/Development** | Use `--random` freely - sites accumulate |
| **Production** | Use `--config sites.json` with a fixed list |
| **Adding more sites** | Add to `sites.json` and re-run |
| **Deleting sites** | Manual deletion in SharePoint Admin Center |

---

## ❓ What Gets Created?

### SharePoint Sites (in Microsoft 365)

Sites are created based on your configuration:

| Mode | Example Sites |
|------|---------------|
| **Config File** | Whatever you define in `sites.json` |
| **Random** | `human-resources`, `finance-department`, `it-security`, etc. |

All sites are created at: `https://yourtenant.sharepoint.com/sites/site-name`

### Azure Resources (Supporting Infrastructure)

| Resource | Purpose |
|----------|---------|
| Resource Group | Container for all Azure resources |
| Key Vault | Secure storage for site URLs and metadata |

---

## 📖 Additional Documentation

| Document | Description |
|----------|-------------|
| [CONFIGURATION-GUIDE.md](./CONFIGURATION-GUIDE.md) | Step-by-step configuration instructions |
| [PREREQUISITES.md](./PREREQUISITES.md) | Detailed prerequisites and setup |
| [docs/TROUBLESHOOTING.md](./docs/TROUBLESHOOTING.md) | Common issues and solutions |

---

## 🔐 Security Considerations

1. **Never commit `terraform.tfvars`** - Contains sensitive information
2. **Never commit `config/sites.json` with sensitive data** - May contain owner emails
3. **Use Azure Key Vault** - For storing secrets in production
4. **Review before applying** - Always check the Terraform plan

---

## 🐍 About the Python Script

The deployment script ([`scripts/deploy.py`](sharepoint-sites-terraform/scripts/deploy.py)) is written in Python for simplicity:

- **Works on**: Windows, macOS, Linux
- **Python version**: 3.8 or higher
- **Dependencies**: None (uses only Python standard library)
- **Auto-installs**: Azure CLI and Terraform if not present

### Automatic Prerequisite Installation

The script automatically detects and installs missing tools:

| Platform | Azure CLI Installation | Terraform Installation |
|----------|------------------------|------------------------|
| **Windows** | winget → chocolatey → manual download | winget → chocolatey → manual download |
| **macOS** | Homebrew (`brew install azure-cli`) | Homebrew (`brew install terraform`) |
| **Linux** | apt package manager | apt package manager |

> 💡 Use `--skip-prerequisites` to bypass automatic installation if you prefer to install tools manually.

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
