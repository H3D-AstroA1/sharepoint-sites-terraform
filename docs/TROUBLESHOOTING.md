# Troubleshooting Guide

## 📋 Overview

This guide helps you resolve common issues when deploying SharePoint sites with Terraform.

---

## 🔍 Quick Diagnostics

Run these commands to diagnose common issues:

```bash
# Check Python version
python --version

# Check Azure CLI login
az account show

# Check Terraform version
terraform --version

# Validate Terraform configuration
cd terraform
terraform validate

# Validate JSON configuration files
python -c "import json; json.load(open('../config/environments.json')); print('environments.json is valid')"
python -c "import json; json.load(open('../config/sites.json')); print('sites.json is valid')"
```

---

## 🚨 Common Issues and Solutions

### Issue 1: Azure CLI Not Logged In

**Error Message:**
```
ERROR: Please run 'az login' to setup account.
```

**Solution:**
```bash
# Login to Azure
az login

# If using a specific tenant
az login --tenant YOUR-TENANT-ID

# Verify login
az account show
```

---

### Issue 2: Azure CLI Not Found in PATH (Windows)

**Error Message:**
```
Error: unable to build authorizer for Resource Manager API: could not configure AzureCli Authorizer: could not parse Azure CLI version: launching Azure CLI: exec: "az": executable file not found in %PATH%
```

**Cause:**
Azure CLI was installed (e.g., via winget) but the installation path was not added to the system PATH environment variable.

**Solution:**

The Python scripts in this project automatically detect Azure CLI in its default installation locations:
- `C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd`
- `C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin\az.cmd`

**If you're running scripts directly:**
1. Use the main menu (`python scripts/menu.py`) which handles PATH automatically
2. Or add Azure CLI to your PATH manually:

```powershell
# PowerShell - Add to current session
$env:PATH = "C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin;" + $env:PATH

# PowerShell - Add permanently (requires admin)
[Environment]::SetEnvironmentVariable("PATH", "C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin;" + [Environment]::GetEnvironmentVariable("PATH", "Machine"), "Machine")
```

```cmd
# Command Prompt - Add to current session
set PATH=C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin;%PATH%
```

**Verify Azure CLI is accessible:**
```bash
az --version
```

---

### Issue 3: Terraform Init Fails

**Error Message:**
```
Error: Failed to query available provider packages
```

**Possible Causes:**
1. No internet connection
2. Corporate proxy blocking access
3. Terraform registry is down

**Solutions:**

**Check internet connectivity:**
```bash
# Test connectivity to Terraform registry
curl -I https://registry.terraform.io
```

**Configure proxy (if behind corporate proxy):**
```bash
# Set proxy environment variables
export HTTP_PROXY=http://proxy.company.com:8080
export HTTPS_PROXY=http://proxy.company.com:8080

# For PowerShell
$env:HTTP_PROXY = "http://proxy.company.com:8080"
$env:HTTPS_PROXY = "http://proxy.company.com:8080"
```

**Clear Terraform cache:**
```bash
# Remove .terraform directory and try again
rm -rf .terraform
rm .terraform.lock.hcl
terraform init
```

---

### Issue 4: Invalid Tenant ID or Subscription ID

**Error Message:**
```
Error: The azure_tenant_id must be a valid GUID format
```

**Solution:**

Ensure your IDs are in the correct format:
- Format: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
- Example: `12345678-1234-1234-1234-123456789abc`

**Find your Tenant ID:**
```bash
az account tenant list --query "[].tenantId" -o table
```

**Find your Subscription ID:**
```bash
az account list --query "[].{Name:name, Id:id}" -o table
```

---

### Issue 5: Environment Not Detected from environments.json

**Symptom:**
The script doesn't offer your pre-configured environment and goes straight to manual configuration.

**Possible Causes:**
1. The `environments.json` file still has placeholder values
2. The JSON file has syntax errors
3. The tenant_id still starts with `<<CHANGE THIS>>`

**Solutions:**

**Check if you've replaced all placeholders:**
```bash
# Look for any remaining placeholders
grep -i "CHANGE THIS" config/environments.json
```

**Validate the JSON syntax:**
```bash
python -c "import json; json.load(open('config/environments.json')); print('JSON is valid!')"
```

**Ensure tenant_id is a real GUID:**
- ❌ Wrong: `"tenant_id": "<<CHANGE THIS: Your Azure Tenant ID>>"`
- ✅ Correct: `"tenant_id": "12345678-1234-1234-1234-123456789abc"`

**Find your actual Tenant ID:**
```bash
az account tenant list --query "[].{Name:displayName, TenantId:tenantId}" -o table
```

---

### Issue 6: Invalid JSON in Configuration Files

**Error Message:**
```
json.decoder.JSONDecodeError: Expecting ',' delimiter
```

**Common Causes:**
1. Missing comma between array items
2. Trailing comma after last item
3. Using single quotes instead of double quotes

**Solution:**

Validate your JSON files:
```bash
# Check environments.json
python -c "import json; json.load(open('config/environments.json')); print('Valid!')"

# Check sites.json
python -c "import json; json.load(open('config/sites.json')); print('Valid!')"
```

**Common JSON mistakes:**
```json
// ❌ Wrong - missing comma
{
  "environments": [
    { "name": "Dev" }    // <-- Missing comma here!
    { "name": "Prod" }
  ]
}

// ❌ Wrong - trailing comma
{
  "environments": [
    { "name": "Dev" },
    { "name": "Prod" },  // <-- Extra comma here!
  ]
}

// ✅ Correct
{
  "environments": [
    { "name": "Dev" },
    { "name": "Prod" }
  ]
}
```

---

### Issue 7: Insufficient Permissions

**Error Message:**
```
Error: AuthorizationFailed: The client does not have authorization to perform action
```

**Solution:**

1. **Check your current permissions:**
```bash
az role assignment list --assignee $(az ad signed-in-user show --query id -o tsv) -o table
```

2. **Request required permissions:**
   - For Azure resources: Request "Contributor" role on the subscription
   - For SharePoint: Request "SharePoint Administrator" role in Microsoft 365

3. **Use a Service Principal with correct permissions:**
```bash
# Create Service Principal with Contributor role
az ad sp create-for-rbac \
  --name "sp-sharepoint-deployment" \
  --role "Contributor" \
  --scopes "/subscriptions/YOUR-SUBSCRIPTION-ID"
```

---

### Issue 8: Resource Group Already Exists

**Error Message:**
```
Error: A resource with the ID "/subscriptions/.../resourceGroups/rg-sharepoint-automation" already exists
```

**Solutions:**

**Option 1: Use the existing resource group (RECOMMENDED)**

When running `deploy.py`, select "Use an existing Resource Group" when prompted:
```
Would you like to:
  [1] Create a new Resource Group
  [2] Use an existing Resource Group

Enter your choice: 2
```

Or manually set in `terraform.tfvars`:
```hcl
resource_group_name         = "rg-sharepoint-automation"
use_existing_resource_group = true
```

**Option 2: Import existing resource group:**
```bash
terraform import azurerm_resource_group.main /subscriptions/YOUR-SUB-ID/resourceGroups/rg-sharepoint-automation
```

**Option 3: Use a different resource group name:**
Edit `terraform.tfvars`:
```hcl
resource_group_name = "rg-sharepoint-automation-new"
use_existing_resource_group = false
```

**Option 4: Delete existing resource group (if empty):**
```bash
az group delete --name rg-sharepoint-automation --yes
```

---

### Issue 9: Key Vault Name Already Taken

**Error Message:**
```
Error: Key Vault name 'kv-sharepoint-xxx' is already in use
```

**Cause:** Key Vault names must be globally unique across all of Azure.

**Solution:**

Specify a custom Key Vault name in `terraform.tfvars`:
```hcl
key_vault_name = "kv-sp-yourcompany-unique123"
```

Or disable Key Vault creation:
```hcl
create_key_vault = false
```

---

### Issue 10: SharePoint Site Creation Fails

**Error Message:**
```
Error: Failed to create SharePoint site
```

**Possible Causes:**
1. Not authenticated to SharePoint
2. Site already exists
3. Invalid site name
4. Insufficient permissions

**Solutions:**

**Check SharePoint authentication:**
```powershell
# Connect to SharePoint Admin
Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive

# Verify connection
Get-PnPTenantSite
```

**Check if site already exists:**
```powershell
# List all sites
Get-PnPTenantSite | Where-Object { $_.Url -like "*executive*" }
```

**Verify site name is valid:**
- No spaces (use hyphens instead)
- No special characters except hyphens
- Must be unique within your tenant

---

### Issue 11: PnP PowerShell Connection Issues

**Error Message:**
```
Connect-PnPOnline: AADSTS65001: The user or administrator has not consented to use the application
```

**Solution:**

Register PnP Management Shell in your tenant:
```powershell
# Run as Global Admin or Application Admin
Register-PnPManagementShellAccess

# This opens a browser for consent
# Grant consent for your organization
```

---

### Issue 12: Terraform State Lock

**Error Message:**
```
Error: Error acquiring the state lock
```

**Cause:** Another Terraform process is running or crashed without releasing the lock.

**Solution:**

**Force unlock (use with caution):**
```bash
terraform force-unlock LOCK-ID
```

**If using local state, delete the lock file:**
```bash
rm .terraform.tfstate.lock.info
```

---

### Issue 13: Module Not Found

**Error Message:**
```
The term 'Connect-PnPOnline' is not recognized as the name of a cmdlet
```

**Solution:**

Install the missing module:
```powershell
# Install PnP.PowerShell
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force

# Import the module
Import-Module PnP.PowerShell

# Verify installation
Get-Command Connect-PnPOnline
```

---

## 🔧 Advanced Troubleshooting

### Enable Terraform Debug Logging

```bash
# Set debug level
export TF_LOG=DEBUG
export TF_LOG_PATH=./terraform-debug.log

# Run Terraform
terraform plan

# Review the log
cat terraform-debug.log
```

### Enable PowerShell Verbose Output

```powershell
# Run script with verbose output
.\Deploy-SharePointSites.ps1 -Verbose

# Or set preference
$VerbosePreference = "Continue"
```

### Check Azure Activity Log

```bash
# View recent operations
az monitor activity-log list \
  --resource-group rg-sharepoint-automation \
  --start-time $(date -d '1 hour ago' -Iso8601) \
  --query "[].{Operation:operationName.localizedValue, Status:status.localizedValue, Time:eventTimestamp}" \
  -o table
```

### Check SharePoint Admin Logs

1. Go to [SharePoint Admin Center](https://yourtenant-admin.sharepoint.com)
2. Navigate to Reports > Usage
3. Check for any error messages

---

## 🔄 Recovery Procedures

### Recover from Failed Deployment

1. **Check current state:**
```bash
terraform state list
```

2. **Remove problematic resources from state:**
```bash
terraform state rm azurerm_resource_group.main
```

3. **Re-run deployment:**
```bash
terraform plan
terraform apply
```

### Rollback Changes

```bash
# Destroy all created resources
terraform destroy

# Confirm destruction
# Type 'yes' when prompted
```

### Clean Start

```bash
# Remove all Terraform files
rm -rf .terraform
rm .terraform.lock.hcl
rm terraform.tfstate*
rm tfplan

# Re-initialize
terraform init
```

---

## 📊 Diagnostic Commands Reference

| Command | Purpose |
|---------|---------|
| `az account show` | Check Azure login status |
| `az account list` | List available subscriptions |
| `terraform validate` | Validate Terraform configuration |
| `terraform state list` | List resources in state |
| `Get-PnPTenantSite` | List SharePoint sites |
| `Get-Module -ListAvailable` | List installed PowerShell modules |

---

## 📞 Getting Help

### Collect Diagnostic Information

Before seeking help, collect this information:

```powershell
# Save diagnostic info to file
$diagnostics = @{
    AzureCliVersion = (az --version 2>&1 | Select-Object -First 1)
    TerraformVersion = (terraform --version 2>&1 | Select-Object -First 1)
    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
    OS = [System.Environment]::OSVersion.VersionString
    Modules = (Get-Module -ListAvailable Az*, PnP* | Select-Object Name, Version)
}

$diagnostics | ConvertTo-Json | Out-File "diagnostics.json"
```

### Resources

- [Terraform Azure Provider Documentation](https://registry.terraform.io/providers/hashicorp/azurerm/latest/docs)
- [PnP PowerShell Documentation](https://pnp.github.io/powershell/)
- [SharePoint Online Administration](https://docs.microsoft.com/en-us/sharepoint/sharepoint-online)
- [Azure CLI Documentation](https://docs.microsoft.com/en-us/cli/azure/)

---

## 🔄 Random Site Generation Issues

### Issue 14: Terraform Wants to Destroy Sites After Re-running Random Mode

**Symptom:**
After running `--random 10` twice, Terraform shows it wants to destroy some sites and create others.

**Cause:**
Each run of `--random` shuffles the 39 available department templates randomly. If you get different sites in each run, Terraform sees the missing sites as "to be destroyed" and new sites as "to be created."

**Important Note:**
The `null_resource` approach does NOT actually delete SharePoint sites - they remain in your tenant. Only Terraform's tracking (state) changes.

**Solutions:**

**Option 1: Use Configuration File Mode for Consistent Deployments**
```bash
# Edit config/sites.json with your desired sites
python deploy.py --config config/sites.json
```

**Option 2: Use All 39 Sites for Consistent Random Deployments**
```bash
# Always generates the same 39 sites (just in different order)
python deploy.py --random 39
```

**Option 3: Accept the State Changes**
If you don't mind Terraform "forgetting" some sites (they still exist in SharePoint):
```bash
terraform apply
```

---

### Issue 15: Requested More Than 39 Random Sites

**Error Message:**
```
WARNING: Requested 50 sites, but only 39 unique templates available.
INFO: Generating 39 sites instead.
```

**Cause:**
The random generation mode has 39 unique department site templates. You cannot generate more than 39 unique sites in random mode.

**Solution:**
For more than 39 sites, use the configuration file mode:
```bash
# Edit config/sites.json to add as many sites as you need
python deploy.py --config config/sites.json
```

---

### Issue 16: SharePoint Sites Still Exist After Terraform Destroy

**Symptom:**
After running `terraform destroy`, the SharePoint sites still exist in your tenant.

**Cause:**
The Terraform configuration uses `null_resource` with only a creation provisioner. There is no destroy provisioner to delete SharePoint sites.

**This is by design** - it prevents accidental deletion of SharePoint sites and their content.

**Solution:**
To delete SharePoint sites, use the SharePoint Admin Center:
1. Go to [SharePoint Admin Center](https://yourtenant-admin.sharepoint.com)
2. Navigate to Sites > Active sites
3. Select the sites you want to delete
4. Click "Delete"

Or use PowerShell:
```powershell
# Connect to SharePoint Admin
Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive

# Delete a specific site
Remove-PnPTenantSite -Url "https://yourtenant.sharepoint.com/sites/site-name" -Force
```

---

### Issue 17: Resource Provider Registration Error

**Error Message:**
```
Error: Error ensuring Resource Providers are registered.
Terraform automatically attempts to register the Resource Providers it supports to ensure it's able to provision resources.
If you don't have permission to register Resource Providers you may wish to use the "skip_provider_registration" flag in the Provider block to disable this functionality.

Original Error: Cannot register providers: Microsoft.TimeSeriesInsights, Microsoft.Media, Microsoft.MixedReality.
```

**Cause:**
The Azure subscription has deprecated or unavailable resource providers that Terraform tries to register automatically. Some providers like `Microsoft.TimeSeriesInsights`, `Microsoft.Media`, and `Microsoft.MixedReality` may be deprecated or not available in your subscription.

**Solution:**

This project's `providers.tf` already includes the fix:
```hcl
provider "azurerm" {
  features { ... }
  
  # Skip automatic resource provider registration
  skip_provider_registration = true
}
```

If you're seeing this error, ensure your `providers.tf` has `skip_provider_registration = true` in the azurerm provider block.

**Alternative: Register Required Providers Manually**

If you prefer to register providers manually:
```bash
# Register only the providers you need
az provider register --namespace Microsoft.Storage
az provider register --namespace Microsoft.KeyVault
az provider register --namespace Microsoft.Resources

# Check registration status
az provider show --namespace Microsoft.Storage --query "registrationState"
```

---

### Issue 18: Deleted Sites Still Appear in SharePoint Admin Center

**Symptom:**
After deleting SharePoint sites (via cleanup.py or manually), the sites still appear in the SharePoint Admin Center under "Deleted Sites".

**Cause:**
When you delete a SharePoint site, it goes to **two separate recycle bins**:
1. **Microsoft 365 Groups Recycle Bin** (Azure AD) - The M365 Group is soft-deleted
2. **SharePoint Site Recycle Bin** (SharePoint Admin Center) - The site appears in "Deleted Sites"

**Solution:**
Use the cleanup script to purge both recycle bins:

```bash
cd sharepoint-sites-terraform/scripts

# Step 1: Purge M365 Groups from Azure AD recycle bin
python cleanup.py --purge-deleted

# Step 2: Purge SharePoint sites from SharePoint Admin Center recycle bin
# Replace 'contoso' with your tenant name (e.g., contoso.sharepoint.com)
python cleanup.py --purge-spo-recycle --tenant contoso
```

Or use the main menu:
```bash
python menu.py
# Select [3] Delete Files or Sites
# Select [6] Purge M365 Groups recycle bin
# Select [7] Purge SharePoint site recycle bin
# Select [8] Purge site files/folders recycle bin
```

---

### Issue 19: SharePoint Online PowerShell Module Installation Fails

**Error Message:**
```
ERROR: Failed to install SharePoint Online PowerShell module
```

**Cause:**
The SharePoint Online PowerShell module requires Windows PowerShell (not PowerShell 7) and may need the NuGet provider.

**Solution:**

1. **Ensure you're using Windows PowerShell** (not PowerShell 7):
   - The script automatically uses `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`

2. **Install NuGet provider manually** (if automatic installation fails):
   ```powershell
   # Run in Windows PowerShell as Administrator
   Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
   ```

3. **Install the SPO module manually**:
   ```powershell
   # Run in Windows PowerShell as Administrator
   Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber
   ```

4. **Verify installation**:
   ```powershell
   Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell
   ```

---

### Issue 20: Re-authentication Required for Each Site Deletion

**Symptom:**
When purging the SharePoint site recycle bin, you're prompted to authenticate for each site.

**Cause:**
This was a bug in earlier versions where each site deletion created a new PowerShell session.

**Solution:**
Update to the latest version of cleanup.py which uses batch deletion (single authentication for all sites):

```bash
# Pull latest changes
cd sharepoint-sites-terraform
git pull

# Run the purge command - now uses batch deletion
python scripts/cleanup.py --purge-spo-recycle --tenant contoso
```

The batch deletion connects once and deletes all selected sites in a single session.

---

### Issue 21: SPO Module Not Found After Installation

**Symptom:**
```
ERROR: SharePoint Online PowerShell module is not installed
```
Even though you just installed it.

**Cause:**
The module may be installed in a different PowerShell version's module path.

**Solution:**

1. **Check where the module is installed**:
   ```powershell
   # In Windows PowerShell
   Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell | Select-Object Path
   ```

2. **Ensure you're using Windows PowerShell** (not PowerShell 7):
   ```cmd
   C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -Command "Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell"
   ```

3. **Reinstall in the correct PowerShell**:
   ```powershell
   # Run in Windows PowerShell (not PowerShell 7)
   Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber -Scope CurrentUser
   ```

---

### Issue 22: Site Recycle Bin Shows Empty (PnP PowerShell Required)

**Symptom:**
When using menu option `[8] Purge site files/folders recycle bin`, one of these happens:
- The script reports all sites as empty unexpectedly
- One or more sites fail with app-only auth errors (for example `AADSTS700027`)
- You see inconsistent behavior across sites in the same run

**Cause:**
- SharePoint file/folder recycle bins are accessed through PnP PowerShell, not Graph API recycle endpoints.
- Non-interactive mode depends on app-only auth configuration in `scripts/.app_config.json`.
- Certificate assertion drift (missing/old app key) can break cert auth for specific sites unless fallback auth is available.

**Solution:**
Use the current non-interactive batch flow with preflight validation and certificate bootstrap:

```bash
cd sharepoint-sites-terraform/scripts

# Recommended first run (headless + auto certificate bootstrap)
python cleanup.py --purge-site-recycle --non-interactive --auto-setup-cert --yes

# Optional: tune chunk size for larger tenants
python cleanup.py --purge-site-recycle --non-interactive --yes --chunk-size 30
```

1. **If PnP.PowerShell is missing, install it**:
  ```powershell
  Install-Module -Name "PnP.PowerShell" -Scope CurrentUser -Force -AllowClobber
  ```

2. **If cert auth is not configured, run one-time setup**:
  ```bash
  python cleanup.py --setup-cert-auth
  ```

3. **Verify app permissions for headless mode**:
  - SharePoint application permission: `Sites.FullControl.All`
  - Admin consent granted in Entra app registration

4. **Expected runtime behavior**:
  - One non-interactive preflight against the first eligible site
  - Chunked batch purge in shared PowerShell sessions
  - Automatic fallback to client secret when certificate assertion fails (`AADSTS700027`) and secret is present

**Understanding the Three Recycle Bins:**

SharePoint has multiple recycle bins at different levels:

| Recycle Bin | What It Contains | How to Purge | Menu Option |
|-------------|------------------|--------------|-------------|
| **M365 Groups Recycle Bin** | Deleted M365 Groups (Azure AD) | `--purge-deleted` | [6] |
| **SharePoint Site Recycle Bin** | Deleted SharePoint sites | `--purge-spo-recycle` | [7] |
| **Site Document Library Recycle Bin** | Deleted files/folders within a site | `--purge-site-recycle` | [8] |

**Example Usage:**
```bash
# Via menu
python scripts/menu.py
# Select [3] Delete Files or Sites
# Select [8] Purge site files/folders recycle bin

# Via command line
python scripts/cleanup.py --purge-site-recycle --non-interactive --yes
```

---

### Issue 23: 403 Forbidden When Accessing SharePoint Sites

**Error Message:**
```
✗ Failed to get sites: 403 - Forbidden
✗ No SharePoint sites found
```

**Cause:**
The Azure CLI's access token doesn't have the required Microsoft Graph permissions to access SharePoint sites. The Azure CLI uses a first-party app registration (`04b07795-8ddb-461a-bbee-02f9e1bf7b46`) which by default doesn't have SharePoint permissions.

**Solution:**

**Option 1: Use the App Registration Menu (RECOMMENDED - Easiest)**

The main menu includes an automatic app registration feature that creates a custom app with all required permissions:

```bash
cd sharepoint-sites-terraform/scripts
python menu.py
# Press [A] to open App Registration menu
# Select [1] to create a new app registration
```

This will:
1. Create a custom app registration in your tenant
2. Configure all required Microsoft Graph permissions
3. Open a browser for admin consent
4. Save the credentials locally for future use

The app is granted these permissions automatically:
- `Sites.Read.All` - Read SharePoint sites
- `Sites.ReadWrite.All` - Create/modify SharePoint sites
- `Files.ReadWrite.All` - Upload/delete files
- `Group.Read.All` - Read Microsoft 365 Groups
- `Group.ReadWrite.All` - Create/delete Microsoft 365 Groups

**Option 2: Grant Admin Consent to Azure CLI**

A tenant administrator can grant the Azure CLI app the required permissions:

1. Go to **Azure Portal** > **Microsoft Entra ID** > **Enterprise Applications**
2. Search for **"Azure CLI"** (App ID: `04b07795-8ddb-461a-bbee-02f9e1bf7b46`)
3. Click on the Azure CLI application
4. Go to **Permissions** > **Grant admin consent for [your tenant]**
5. Grant these permissions:
   - `Sites.Read.All`
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
   - `Group.Read.All`
   - `Group.ReadWrite.All`

**Option 3: Re-login with Correct Scope**

Try logging in with the Microsoft Graph scope:
```bash
az login --scope https://graph.microsoft.com/.default
```

**Option 4: Use PowerShell with PnP**

If you can't get admin consent, use PnP PowerShell instead:
```powershell
# Install PnP PowerShell
Install-Module -Name PnP.PowerShell -Force

# Connect to SharePoint
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com" -Interactive

# List sites
Get-PnPTenantSite
```

**Verification:**

After setting up app registration or granting consent, verify the token has the correct permissions:
```bash
# Get a new token
az account get-access-token --resource https://graph.microsoft.com

# Test the API (should return sites)
curl -H "Authorization: Bearer $(az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv)" \
  "https://graph.microsoft.com/v1.0/sites?search=*"
```

---

### Issue 24: 403 Errors on Specific Sites During File Population

**Error Message:**
```
✗ Failed to upload file to "My workspace" - 403 Forbidden
✗ Failed to upload file to "Designer" - 403 Forbidden
✗ Failed to upload file to "Team Site" - 403 Forbidden
```

**Cause:**
Some SharePoint sites are system sites or personal sites that don't allow file uploads via the Microsoft Graph API, even with proper permissions. These include:
- "My workspace" - Personal OneDrive-like workspace
- "Designer" - Microsoft Designer integration site
- "Team Site" / "Communication Site" - Default template sites
- Sites containing "contenttypehub", "appcatalog", "search", "portal"
- Personal sites (URLs containing "/personal/")

**Solution:**

The `populate_files.py` and `menu.py` scripts now automatically filter out these problematic sites. The filtering happens automatically when:
- Listing sites for file population
- Uploading files to sites
- Listing files in sites

**Filtered Site Patterns:**
- `my workspace` - Personal OneDrive-like workspace
- `designer` - Microsoft Designer integration site
- `contenttypehub` - SharePoint content type hub
- `appcatalog` - SharePoint app catalog
- URLs containing `/personal/` - Personal OneDrive sites

**Manual Workaround:**

If you need to work with a specific site that's being filtered, you can use the `--site` flag to target it directly:
```bash
python populate_files.py --site "specific-site-name" --files 10
```

---

### Issue 25: M365 Groups Fallback When Sites API Returns 403

**Symptom:**
```
ℹ Using Microsoft 365 Groups API...
Found 15 sites via M365 Groups
```

**Cause:**
When the Microsoft Graph Sites API returns a 403 Forbidden error (due to missing permissions), the scripts automatically fall back to using the Microsoft 365 Groups API. This API can list groups that have associated SharePoint sites.

**This is expected behavior** and not an error. The fallback provides:
- Access to team sites created via Microsoft Teams
- Access to sites created via Microsoft 365 Groups
- Filtered list excluding system/personal sites

**To avoid the fallback:**

Use the App Registration feature to get proper permissions:
```bash
python menu.py
# Press [A] to open App Registration menu
# Select [1] to create a new app registration
```

Once the custom app is configured with admin consent, the scripts will use the Sites API directly.

---

### Issue 26: Deleted Sites Reappear in SharePoint Admin Center

**Symptom:**
After deleting SharePoint sites via the cleanup script, the sites still appear in the SharePoint Admin Center's "Active sites" list, or they reappear after being manually deleted.

**Cause:**
Microsoft 365 uses a **two-stage deletion process** for SharePoint sites:

1. **Stage 1 (Soft Delete)**: When you delete an M365 Group, the group goes to the "Deleted Groups" recycle bin AND the SharePoint site goes to the "Deleted Sites" recycle bin
2. **Stage 2 (Hard Delete)**: Sites must be purged from the SharePoint recycle bin to be permanently deleted

If only Stage 1 is completed, the sites remain in the recycle bin and may still appear in the Admin Center.

**Solution:**

The cleanup script now automatically handles both stages:

1. **Auto-detects tenant name** from site URLs (e.g., `contoso` from `https://contoso.sharepoint.com/sites/...`)
2. **Prompts to purge recycle bins** after deletion
3. **Purges both** the M365 Groups recycle bin AND the SharePoint site recycle bin

**Manual Purge:**

If sites are stuck in the recycle bin, use menu option `[7] Purge SharePoint site recycle bin`:

```bash
python menu.py
# Select [3] Delete Files or Sites
# Select [7] Purge SharePoint site recycle bin
# Select sites to permanently delete
# Type 'PURGE' to confirm
```

**PowerShell Alternative:**

```powershell
# Connect to SharePoint Online
Connect-SPOService -Url https://contoso-admin.sharepoint.com

# List deleted sites
Get-SPODeletedSite

# Permanently delete a specific site
Remove-SPODeletedSite -Identity https://contoso.sharepoint.com/sites/sitename

# Permanently delete all deleted sites
Get-SPODeletedSite | Remove-SPODeletedSite
```

---

### Issue 27: System Sites Protected from File Population

**Symptom:**
When running "Populate Sites with Files", some sites are skipped with a message like:
```
⚠ PROTECTED SYSTEM SITES (excluded from file population):
  These sites are protected and will not have files uploaded:
    • My workspace
    • Designer
    • Team Site
    • Communication site
```

**Cause:**
This is **expected behavior**. The following sites are protected system sites that should not have files uploaded:

| Site Type | Description |
|-----------|-------------|
| My workspace | Personal OneDrive-like workspace |
| Designer | Microsoft Designer integration site |
| Team Site | Default team site template |
| Communication Site | Default communication site template |
| Content Type Hub | SharePoint content type hub |
| App Catalog | SharePoint app catalog |
| Personal sites | URLs containing `/personal/` |
| Root site | The tenant root (e.g., `contoso.sharepoint.com`) |

**This is not an error** - the script is protecting these system sites from accidental modification.

**Solution:**

Only user-created sites (typically created via Terraform or M365 Groups) will have files uploaded. If you need to populate a specific site that's being filtered, you can:

1. Create a new site via the deployment script
2. Use the Microsoft Graph API directly
3. Upload files manually via the SharePoint web interface

---

## ✅ Prevention Tips

1. **Always run `terraform plan` before `terraform apply`**
2. **Keep Terraform and providers updated**
3. **Use version control for your configuration**
4. **Never commit sensitive values (use `.gitignore`)**
5. **Test in a development environment first**
6. **Document any manual changes made outside Terraform**
7. **Purge recycle bins after deleting sites** to fully remove them
8. **Grant Azure CLI admin consent** before using populate_files.py or cleanup.py
9. **Use the App Registration feature** for proper Microsoft Graph permissions
10. **System sites are protected** - only user-created sites can be modified
