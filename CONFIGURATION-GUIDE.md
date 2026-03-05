# Configuration Guide

## 📋 Overview

This guide explains how to configure the SharePoint sites deployment, including:
- Selecting your Azure tenant, subscription, and resource group
- Configuring Microsoft 365 settings
- Customizing SharePoint site properties

---

## 🎯 Quick Configuration Checklist

Before deploying, you need to gather the following information:

| Item | Where to Find It | Example |
|------|------------------|---------|
| Azure Tenant ID | Azure Portal > Entra ID > Overview | `12345678-1234-1234-1234-123456789abc` |
| Azure Subscription ID | Azure Portal > Subscriptions | `87654321-4321-4321-4321-cba987654321` |
| M365 Tenant Name | Your domain before .onmicrosoft.com | `contoso` |
| SharePoint Admin Email | Your admin account | `admin@contoso.onmicrosoft.com` |

---

## 🔧 Configuration Methods

### Method 1: Interactive Script (Recommended for Beginners)

The deployment script will prompt you for all required information:

```bash
# Works on Windows, macOS, and Linux
cd sharepoint-sites-terraform/scripts
python deploy.py
```

The script will:
1. ✅ Check and install prerequisites (Azure CLI, Terraform)
2. ✅ List available tenants for selection
3. ✅ List available subscriptions for selection
4. ✅ Allow you to create or select a resource group
5. ✅ Prompt for M365 configuration
6. ✅ Generate the Terraform configuration automatically

**Deployment Modes:**
- **Configuration File Mode**: Use `config/sites.json` for custom site names (unlimited sites)
- **Random Generation Mode**: Generate realistic organizational department sites (1-39 unique templates)

### Method 2: Pre-Configured Environments (Recommended for Teams)

Edit `config/environments.json` to pre-configure your Azure and M365 settings. This is the easiest method for novice users!

1. **Open the file** `config/environments.json`
2. **Replace all values** marked with `<<CHANGE THIS>>` with your actual values
3. **Run the script** - it will automatically detect your configured environment

```bash
python deploy.py
```

The script will offer your pre-configured environment, so users don't need to enter tenant IDs manually.

**To add more environments** (e.g., Production, Development):
- Follow the instructions in the `_HOW_TO_ADD_MORE_ENVIRONMENTS` section of the JSON file

### Method 3: Manual Configuration (For Advanced Users)

Edit the `terraform/terraform.tfvars` file directly:

```powershell
# Copy the example file
cd terraform
copy terraform.tfvars.example terraform.tfvars

# Edit with your values
notepad terraform.tfvars
```

---

## 📍 Finding Your Azure Tenant ID

### Step-by-Step Instructions

1. **Open Azure Portal**
   - Go to [https://portal.azure.com](https://portal.azure.com)
   - Sign in with your Azure account

2. **Navigate to Microsoft Entra ID**
   - In the search bar at the top, type "Microsoft Entra ID"
   - Click on "Microsoft Entra ID" in the results

3. **Find the Tenant ID**
   - On the Overview page, look for "Tenant ID"
   - It's displayed in the "Basic information" section
   - Click the copy icon to copy it

   ![Tenant ID Location](https://docs.microsoft.com/en-us/azure/active-directory/fundamentals/media/active-directory-how-to-find-tenant/portal-tenant-id.png)

### Using Azure CLI

```bash
# List all tenants you have access to
az account tenant list --query "[].{Name:displayName, TenantId:tenantId}" -o table

# Output example:
# Name              TenantId
# ----------------  ------------------------------------
# Contoso Ltd       12345678-1234-1234-1234-123456789abc
```

---

## 📍 Finding Your Azure Subscription ID

### Step-by-Step Instructions

1. **Open Azure Portal**
   - Go to [https://portal.azure.com](https://portal.azure.com)

2. **Navigate to Subscriptions**
   - In the search bar, type "Subscriptions"
   - Click on "Subscriptions" in the results

3. **Find Your Subscription**
   - You'll see a list of all subscriptions you have access to
   - Click on the subscription you want to use
   - The Subscription ID is displayed on the Overview page

### Using Azure CLI

```bash
# List all subscriptions
az account list --query "[].{Name:name, SubscriptionId:id, State:state}" -o table

# Output example:
# Name                    SubscriptionId                        State
# ----------------------  ------------------------------------  -------
# Production              87654321-4321-4321-4321-cba987654321  Enabled
# Development             11111111-2222-3333-4444-555555555555  Enabled
```

### Selecting a Subscription

```bash
# Set the active subscription
az account set --subscription "87654321-4321-4321-4321-cba987654321"

# Verify the selection
az account show --query "{Name:name, SubscriptionId:id}" -o table
```

---

## 📍 Finding Your M365 Tenant Name

### Step-by-Step Instructions

1. **Open Microsoft 365 Admin Center**
   - Go to [https://admin.microsoft.com](https://admin.microsoft.com)
   - Sign in with your admin account

2. **Find Your Tenant Name**
   - Look at the URL or your domain
   - Your tenant name is the part before `.onmicrosoft.com`
   
   **Example:**
   - If your domain is `contoso.onmicrosoft.com`
   - Your tenant name is `contoso`

3. **Alternative: Check SharePoint URL**
   - Go to SharePoint: [https://yourtenant.sharepoint.com](https://yourtenant.sharepoint.com)
   - The tenant name is in the URL

### Using PowerShell

```powershell
# Connect to Microsoft 365
Connect-MgGraph -Scopes "Organization.Read.All"

# Get organization details
Get-MgOrganization | Select-Object DisplayName, VerifiedDomains
```

---

## 🏢 Resource Group Configuration

### What is a Resource Group?

A Resource Group is a container that holds related Azure resources. All Azure resources created by this deployment will be placed in the specified resource group.

### Naming Conventions

We recommend using a descriptive naming convention:

```
rg-<project>-<environment>-<region>

Examples:
- rg-sharepoint-prod-uksouth
- rg-sharepoint-dev-westeurope
- rg-sp-automation-eastus
```

### Creating a New Resource Group

**Option 1: Let the Script Create It**
- The deployment script will create the resource group automatically
- Just provide the name and location when prompted

**Option 2: Create Manually via Azure Portal**
1. Go to Azure Portal
2. Search for "Resource groups"
3. Click "+ Create"
4. Fill in:
   - Subscription: Select your subscription
   - Resource group: Enter your name (e.g., `rg-sharepoint-automation`)
   - Region: Select your region (e.g., `UK South`)
5. Click "Review + create" then "Create"

**Option 3: Create via Azure CLI**
```bash
az group create \
  --name "rg-sharepoint-automation" \
  --location "uksouth" \
  --tags Environment=Production Project=SharePoint-Automation
```

---

## 🌐 SharePoint Sites Configuration

### Default Sites

The deployment creates 4 SharePoint sites by default:

| Site URL Name | Display Name | Purpose |
|---------------|--------------|---------|
| `executive-confidential` | Executive Confidential | Executive leadership documents |
| `finance-internal` | Finance Internal | Finance team collaboration |
| `claims-operations` | Claims Operations | Claims processing workspace |
| `it-infra` | IT Infrastructure | IT documentation and runbooks |

### Customizing Sites

Edit the `sharepoint_sites` variable in `terraform.tfvars`:

```hcl
sharepoint_sites = {
  "executive-confidential" = {
    display_name = "Executive Confidential"
    description  = "Your custom description here"
    template     = "STS#3"      # Site template
    visibility   = "Private"    # Private or Public
    owners       = [
      "ceo@yourtenant.onmicrosoft.com",
      "cfo@yourtenant.onmicrosoft.com"
    ]
    members      = [
      "executive.assistant@yourtenant.onmicrosoft.com"
    ]
  }
  # ... more sites
}
```

### Site Templates

| Template | Description | Use Case |
|----------|-------------|----------|
| `STS#3` | Team site (no M365 Group) | Simple document storage |
| `GROUP#0` | Team site (with M365 Group) | Full collaboration with Teams |
| `SITEPAGEPUBLISHING#0` | Communication site | News and announcements |

### Visibility Options

| Option | Description |
|--------|-------------|
| `Private` | Only site members can access |
| `Public` | Anyone in the organization can access |

---

## 🔐 Authentication Configuration

### Option 1: Azure CLI (Recommended for Development)

The simplest method - uses your logged-in Azure CLI credentials:

```bash
# Login to Azure
az login

# The deployment will use your credentials automatically
```

### Option 2: Service Principal (Recommended for CI/CD)

For automated deployments, create a Service Principal:

```bash
# Create Service Principal
az ad sp create-for-rbac \
  --name "sp-sharepoint-deployment" \
  --role "Contributor" \
  --scopes "/subscriptions/YOUR-SUBSCRIPTION-ID"

# Output:
# {
#   "appId": "CLIENT_ID",
#   "displayName": "sp-sharepoint-deployment",
#   "password": "CLIENT_SECRET",
#   "tenant": "TENANT_ID"
# }
```

Configure in `terraform.tfvars`:

```hcl
use_service_principal = true
azure_client_id       = "YOUR-CLIENT-ID"
azure_client_secret   = "YOUR-CLIENT-SECRET"
```

⚠️ **Security Warning**: Never commit client secrets to source control!

### Option 3: Environment Variables

Set credentials as environment variables:

```powershell
# PowerShell
$env:ARM_CLIENT_ID = "YOUR-CLIENT-ID"
$env:ARM_CLIENT_SECRET = "YOUR-CLIENT-SECRET"
$env:ARM_TENANT_ID = "YOUR-TENANT-ID"
$env:ARM_SUBSCRIPTION_ID = "YOUR-SUBSCRIPTION-ID"
```

```bash
# Bash
export ARM_CLIENT_ID="YOUR-CLIENT-ID"
export ARM_CLIENT_SECRET="YOUR-CLIENT-SECRET"
export ARM_TENANT_ID="YOUR-TENANT-ID"
export ARM_SUBSCRIPTION_ID="YOUR-SUBSCRIPTION-ID"
```

---

## 🏷️ Tags Configuration

Tags help organize and track Azure resources:

```hcl
tags = {
  Environment  = "Production"      # prod, dev, test, staging
  Project      = "SharePoint-Sites"
  Department   = "IT"
  Owner        = "IT-Team"
  CostCenter   = "IT-001"
  ManagedBy    = "Terraform"
}
```

### Recommended Tags

| Tag | Purpose | Example Values |
|-----|---------|----------------|
| `Environment` | Deployment environment | Production, Development, Test |
| `Project` | Project name | SharePoint-Automation |
| `Owner` | Responsible team/person | IT-Team, john.doe@company.com |
| `CostCenter` | Billing allocation | IT-001, FINANCE-002 |
| `ManagedBy` | How resource is managed | Terraform, Manual |

---

## ✅ Configuration Validation

Before deploying, validate your configuration:

```bash
cd terraform

# Initialize Terraform
terraform init

# Validate configuration
terraform validate

# Preview changes
terraform plan
```

### Common Validation Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Invalid GUID format" | Tenant/Subscription ID format wrong | Ensure format is `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| "Invalid email address" | Admin email format wrong | Use full email like `admin@tenant.onmicrosoft.com` |
| "Resource group name invalid" | Name contains invalid characters | Use only letters, numbers, hyphens, underscores |

---

## 📚 Complete Configuration Example

Here's a complete `terraform.tfvars` example:

```hcl
# ============================================
# AZURE CONFIGURATION
# ============================================
azure_tenant_id       = "12345678-1234-1234-1234-123456789abc"
azure_subscription_id = "87654321-4321-4321-4321-cba987654321"
resource_group_name   = "rg-sharepoint-automation"
location              = "uksouth"

# ============================================
# MICROSOFT 365 CONFIGURATION
# ============================================
m365_tenant_name       = "contoso"
sharepoint_admin_email = "admin@contoso.onmicrosoft.com"

# ============================================
# SHAREPOINT SITES
# ============================================
sharepoint_sites = {
  "executive-confidential" = {
    display_name = "Executive Confidential"
    description  = "Confidential documents for executive team"
    template     = "STS#3"
    visibility   = "Private"
    owners       = ["ceo@contoso.onmicrosoft.com"]
    members      = []
  }
  "finance-internal" = {
    display_name = "Finance Internal"
    description  = "Finance team documents and reports"
    template     = "STS#3"
    visibility   = "Private"
    owners       = ["cfo@contoso.onmicrosoft.com"]
    members      = ["finance-team@contoso.onmicrosoft.com"]
  }
  "claims-operations" = {
    display_name = "Claims Operations"
    description  = "Claims processing workspace"
    template     = "STS#3"
    visibility   = "Private"
    owners       = []
    members      = []
  }
  "it-infra" = {
    display_name = "IT Infrastructure"
    description  = "IT documentation and runbooks"
    template     = "STS#3"
    visibility   = "Private"
    owners       = ["it-manager@contoso.onmicrosoft.com"]
    members      = ["it-team@contoso.onmicrosoft.com"]
  }
}

# ============================================
# OPTIONAL SETTINGS
# ============================================
create_key_vault              = true
enable_soft_delete_protection = true
deployment_timeout_minutes    = 30

tags = {
  Environment = "Production"
  Project     = "SharePoint-Sites-Automation"
  Department  = "IT"
  Owner       = "IT-Team"
  CostCenter  = "IT-001"
  ManagedBy   = "Terraform"
}
```

---

## 🆘 Need Help?

- Check the [TROUBLESHOOTING.md](./docs/TROUBLESHOOTING.md) guide
- Review the [PREREQUISITES.md](./PREREQUISITES.md) document
- Ensure you have the required permissions in both Azure and Microsoft 365
