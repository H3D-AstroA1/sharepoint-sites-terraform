# Prerequisites Guide

## 📋 Overview

This document lists all prerequisites needed to deploy SharePoint sites using this Terraform solution.

> 💡 **Good News!** The deployment script (`scripts/deploy.py`) **automatically checks and installs** Azure CLI and Terraform if they are missing. You only need Python installed to get started!

---

## ✅ Prerequisites Checklist

### Required Software

| Software | Minimum Version | Auto-Install? | Notes |
|----------|-----------------|---------------|-------|
| **Python** | 3.8+ | ❌ Manual | Required to run the deployment script |
| **Azure CLI** | 2.50.0+ | ✅ Automatic | Script installs if missing |
| **Terraform** | 1.5.0+ | ✅ Automatic | Script installs if missing |

### Automatic Installation Support

The deployment script automatically installs missing tools on:

| Platform | Azure CLI | Terraform |
|----------|-----------|-----------|
| **Windows** | winget → chocolatey → manual | winget → chocolatey → manual |
| **macOS** | Homebrew | Homebrew |
| **Linux** | apt package manager | apt package manager |

### Required Permissions

| Platform | Required Role | Purpose |
|----------|---------------|---------|
| **Azure** | Contributor (on subscription) | Create Azure resources |
| **Microsoft 365** | SharePoint Administrator | Create SharePoint sites |
| **Microsoft Entra ID** | Application Administrator (optional) | Create app registrations |

### Required Accounts

| Account | Purpose |
|---------|---------|
| Azure account with active subscription | Deploy Azure resources |
| Microsoft 365 account with SharePoint license | Create SharePoint sites |

---

## 🐍 Installing Python (Required First)

Python is the only prerequisite you need to install manually. The deployment script handles the rest!

### Windows

**Option 1: Microsoft Store (Easiest)**
1. Open Microsoft Store
2. Search for "Python 3.11" (or latest 3.x version)
3. Click "Get" to install

**Option 2: Official Installer**
1. Download from [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Run the installer
3. ⚠️ **Important**: Check "Add Python to PATH" during installation

**Option 3: winget**
```powershell
winget install Python.Python.3.11
```

### macOS

```bash
# Using Homebrew (recommended)
brew install python@3.11

# Or download from python.org
```

### Linux (Ubuntu/Debian)

```bash
sudo apt update
sudo apt install python3 python3-pip
```

### Verify Installation

```bash
python --version
# Should show: Python 3.x.x
```

---

## 🚀 Quick Start (Recommended)

Once Python is installed, simply run the deployment script:

```bash
cd sharepoint-sites-terraform/scripts
python deploy.py
```

The script will:
1. ✅ Check for Azure CLI and Terraform
2. ✅ Offer to install them automatically if missing
3. ✅ Guide you through the entire deployment process

> 💡 **That's it!** You don't need to manually install Azure CLI or Terraform - the script handles everything.

---

## 🔧 Manual Installation (Optional)

If you prefer to install tools manually, or if automatic installation fails, follow the guides below.

### Installing Azure CLI

### Windows

**Option 1: MSI Installer (Recommended)**
1. Download the installer from [https://aka.ms/installazurecliwindows](https://aka.ms/installazurecliwindows)
2. Run the downloaded MSI file
3. Follow the installation wizard
4. Restart your terminal

**Option 2: PowerShell**
```powershell
# Using winget (Windows 10/11)
winget install Microsoft.AzureCLI

# Using Chocolatey
choco install azure-cli
```

### macOS

```bash
# Using Homebrew
brew update && brew install azure-cli
```

### Linux (Ubuntu/Debian)

```bash
# Install dependencies
sudo apt-get update
sudo apt-get install ca-certificates curl apt-transport-https lsb-release gnupg

# Add Microsoft signing key
curl -sL https://packages.microsoft.com/keys/microsoft.asc | gpg --dearmor | sudo tee /etc/apt/trusted.gpg.d/microsoft.gpg > /dev/null

# Add Azure CLI repository
AZ_REPO=$(lsb_release -cs)
echo "deb [arch=amd64] https://packages.microsoft.com/repos/azure-cli/ $AZ_REPO main" | sudo tee /etc/apt/sources.list.d/azure-cli.list

# Install Azure CLI
sudo apt-get update
sudo apt-get install azure-cli
```

### Verify Installation

```bash
az --version
# Should show: azure-cli 2.x.x
```

---

## 🔧 Installing Terraform

### Windows

**Option 1: Chocolatey**
```powershell
choco install terraform
```

**Option 2: winget**
```powershell
winget install Hashicorp.Terraform
```

**Option 3: Manual Installation**
1. Download from [https://www.terraform.io/downloads](https://www.terraform.io/downloads)
2. Extract the ZIP file
3. Move `terraform.exe` to a directory in your PATH (e.g., `C:\Windows\System32`)

### macOS

```bash
# Using Homebrew
brew tap hashicorp/tap
brew install hashicorp/tap/terraform
```

### Linux

```bash
# Ubuntu/Debian
wget -O- https://apt.releases.hashicorp.com/gpg | sudo gpg --dearmor -o /usr/share/keyrings/hashicorp-archive-keyring.gpg
echo "deb [signed-by=/usr/share/keyrings/hashicorp-archive-keyring.gpg] https://apt.releases.hashicorp.com $(lsb_release -cs) main" | sudo tee /etc/apt/sources.list.d/hashicorp.list
sudo apt update && sudo apt install terraform
```

### Verify Installation

```bash
terraform --version
# Should show: Terraform v1.x.x
```

---

## 🔐 Required Permissions

### Azure Permissions

You need **Contributor** role on the Azure subscription:

**Check Your Permissions:**
```bash
# List your role assignments
az role assignment list --assignee $(az ad signed-in-user show --query id -o tsv) --query "[].{Role:roleDefinitionName, Scope:scope}" -o table
```

**Request Access (if needed):**
1. Contact your Azure administrator
2. Request "Contributor" role on the target subscription
3. Or request a custom role with these permissions:
   - `Microsoft.Resources/subscriptions/resourceGroups/*`
   - `Microsoft.KeyVault/*`

### Microsoft 365 Permissions

You need **SharePoint Administrator** role:

**Check Your Permissions:**
1. Go to [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. Navigate to Users > Active users
3. Find your account and check assigned roles

**Request Access (if needed):**
1. Contact your Microsoft 365 administrator
2. Request "SharePoint Administrator" role
3. Or request "Global Administrator" (has all permissions)

---

## 🌐 Network Requirements

Ensure your network allows access to:

| Service | URLs | Ports |
|---------|------|-------|
| Azure Portal | `*.azure.com`, `*.microsoft.com` | 443 |
| Azure CLI | `management.azure.com` | 443 |
| Terraform Registry | `registry.terraform.io` | 443 |
| SharePoint Online | `*.sharepoint.com` | 443 |
| Microsoft Graph | `graph.microsoft.com` | 443 |
| Python Package Index | `pypi.org` | 443 |

### Proxy Configuration

If you're behind a corporate proxy:

**Windows (Command Prompt):**
```cmd
set HTTP_PROXY=http://proxy.company.com:8080
set HTTPS_PROXY=http://proxy.company.com:8080
```

**macOS/Linux (Bash):**
```bash
export HTTP_PROXY=http://proxy.company.com:8080
export HTTPS_PROXY=http://proxy.company.com:8080
```

---

## ✅ Verification

The deployment script automatically verifies all prerequisites. Simply run:

```bash
cd sharepoint-sites-terraform/scripts
python deploy.py
```

The script will:
1. Check if Python is running (you're already past this if the script runs!)
2. Check for Azure CLI and offer to install if missing
3. Check for Terraform and offer to install if missing
4. Verify Azure login status
5. Guide you through the deployment process

> 💡 **No manual verification needed!** The script handles everything automatically.

---

## 📚 Additional Resources

- [Python Downloads](https://www.python.org/downloads/)
- [Azure CLI Documentation](https://docs.microsoft.com/en-us/cli/azure/)
- [Terraform Documentation](https://www.terraform.io/docs)
- [SharePoint Online Administration](https://docs.microsoft.com/en-us/sharepoint/sharepoint-online)
- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview)

---

## 🆘 Need Help?

If you encounter issues with prerequisites:

1. Check the [TROUBLESHOOTING.md](./docs/TROUBLESHOOTING.md) guide
2. Ensure you have internet connectivity
3. Check if you're behind a corporate proxy
4. Verify you have sufficient permissions to install software
5. Try running the script with `--skip-prerequisites` and install tools manually
