# SharePoint Sites Deployment with Terraform

## 📋 Overview

This project automates the creation of **SharePoint Online sites** using Terraform. It provides flexible options for defining which sites to create:

### Two Deployment Modes

| Mode | Description | Use Case |
|------|-------------|----------|
| **Configuration File** | Define custom site names in a JSON file | When you know exactly which sites you need |
| **Random Generation** | Auto-generate sites with random names | For testing, demos, or bulk site creation |

> ⚠️ **Important Note**: SharePoint Online is a Microsoft 365 service, not an Azure resource. This solution uses Azure resources (Resource Group, Key Vault) for supporting infrastructure, while SharePoint sites are created via Microsoft Graph API.

---

## 📁 Project Structure

```
sharepoint-sites-terraform/
├── README.md                    # This file - main documentation
├── CONFIGURATION-GUIDE.md       # Detailed configuration instructions
├── PREREQUISITES.md             # What you need before starting
├── requirements.txt             # Python dependencies (none required!)
├── .gitignore                   # Git ignore rules
│
├── config/                      # Configuration files
│   ├── sites.json               # ⭐ EDIT THIS FILE for custom sites
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
│   └── cleanup.py               # Step 3: Delete files/sites
│
└── docs/                        # Additional documentation
    └── TROUBLESHOOTING.md       # Common issues and solutions
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
║   SharePoint Sites Management Tool                           ║
╚══════════════════════════════════════════════════════════════╝

  [0] ✓ Check & Install Prerequisites    ← START HERE!
      Azure CLI, Terraform, Azure Login

  ────────────────────────────────────────────────────────────
  [1] 🏗️  Create SharePoint Sites
  [2] 📄 Populate Sites with Files
  [3] 🗑️  Delete Files or Sites

  ────────────────────────────────────────────────────────────
  [4] 📋 List SharePoint Sites
  [5] 📁 List Files in Sites

  [C] ⚙️  Edit Configuration             ← NEW!
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
    },
    {
      "name": "my-custom-site",
      "display_name": "My Custom Site",
      "description": "Add as many sites as you need!"
    }
  ]
}
```

2. **Run the deployment:**
```bash
cd sharepoint-sites-terraform/scripts
python deploy.py
```

3. **Select option [1]** when prompted for Configuration File mode.

#### Option B: Generate Random Sites (Realistic Department Names)

1. **Run the deployment:**
```bash
cd sharepoint-sites-terraform/scripts
python deploy.py --random 10
```

2. Sites will be created with realistic organizational department names like:
   - `human-resources` - HR policies, employee handbook, benefits (Private)
   - `finance-department` - Financial reports, budgets, accounting (Private)
   - `it-helpdesk` - IT support documentation, troubleshooting guides (Public)
   - `employee-intranet` - Central hub for all employees (Public, Communication site)
   - `legal-compliance` - Legal documents, contracts, regulatory compliance (Private)

> 📊 **Random Site Distribution**: The 39 available templates include 25 Private sites (64%) and 14 Public sites (36%), with 6 Communication sites for company-wide announcements.

### Step 3: Follow the Interactive Prompts

The script will guide you through:

1. ✅ **Prerequisite check & auto-install** (Azure CLI, Terraform)
2. ✅ Site generation mode selection
3. ✅ Azure tenant selection
4. ✅ Subscription selection
5. ✅ Resource group configuration
6. ✅ Microsoft 365 settings
7. ✅ Configuration review
8. ✅ Terraform deployment

> 💡 **Tip**: You can press `Q` to quit or `Ctrl+C` to cancel at any interactive prompt.

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

### Permissions Required

| Operation | Required Permission |
|-----------|---------------------|
| Delete files | Sites.ReadWrite.All, Files.ReadWrite.All |
| Delete sites | Sites.FullControl.All (SharePoint Admin) |

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
