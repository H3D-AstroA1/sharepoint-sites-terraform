# ============================================================================
# MAIN TERRAFORM CONFIGURATION
# ============================================================================
# This is the main entry point for the Terraform configuration.
# It orchestrates the creation of all resources.
#
# WHAT THIS FILE DOES:
# 1. Creates Azure supporting resources (Resource Group, Key Vault)
# 2. Sets up the infrastructure needed for SharePoint site creation
# 3. Triggers the SharePoint site creation via Microsoft Graph API
#
# EXECUTION ORDER:
# 1. Resource Group is created first
# 2. Key Vault is created (if enabled)
# 3. SharePoint sites are created via PowerShell/Graph API
# ============================================================================

# ============================================================================
# LOCAL VALUES
# ============================================================================
# Local values are like variables but computed within Terraform.
# They help simplify the configuration and avoid repetition.
# ============================================================================

locals {
  # Generate a unique suffix for globally unique resource names
  # This ensures Key Vault names don't conflict with other Azure users
  unique_suffix = random_string.suffix.result

  # Construct the Key Vault name if not provided
  key_vault_name = var.key_vault_name != "" ? var.key_vault_name : "kv-sp-${local.unique_suffix}"

  # SharePoint tenant URL
  sharepoint_url = "https://${var.m365_tenant_name}.sharepoint.com"

  # SharePoint admin URL
  sharepoint_admin_url = "https://${var.m365_tenant_name}-admin.sharepoint.com"

  # Common tags to apply to all resources
  common_tags = merge(var.tags, {
    DeployedAt = timestamp()
    TerraformManaged = "true"
  })

  # List of site names for iteration
  site_names = keys(var.sharepoint_sites)

  # URL segment used by the site creation script.
  # Team sites use alphanumeric-only mailNickname, and communication sites
  # may append "site" when created via fallback group path.
  site_url_segments = {
    for site_name, site_config in var.sharepoint_sites :
    site_name => (
      site_config.template == "SITEPAGEPUBLISHING#0"
      ? "${lower(replace(site_name, "/[^0-9A-Za-z]/", ""))}site"
      : lower(replace(site_name, "/[^0-9A-Za-z]/", ""))
    )
  }
}

# ============================================================================
# RANDOM STRING FOR UNIQUE NAMES
# ============================================================================
# Generates a random string to ensure globally unique resource names.
# ============================================================================

resource "random_string" "suffix" {
  length  = 8
  special = false
  upper   = false
  numeric = true
}

# ============================================================================
# DATA SOURCES
# ============================================================================
# Data sources fetch information about existing resources.
# ============================================================================

# Get current Azure client configuration
data "azurerm_client_config" "current" {}

# Get current Azure AD client configuration
data "azuread_client_config" "current" {}

# Get existing resource group (if use_existing_resource_group is true)
data "azurerm_resource_group" "existing" {
  count = var.use_existing_resource_group ? 1 : 0
  name  = var.resource_group_name
}

# ============================================================================
# RESOURCE GROUP
# ============================================================================
# The Resource Group is a container for all Azure resources.
# All resources in this deployment will be placed in this group.
#
# If use_existing_resource_group is true, we use the existing resource group.
# If false, we create a new one.
# ============================================================================

resource "azurerm_resource_group" "main" {
  count    = var.use_existing_resource_group ? 0 : 1
  name     = var.resource_group_name
  location = var.location
  tags     = local.common_tags
}

# Local values for resource group reference (works for both new and existing)
locals {
  resource_group_name     = var.use_existing_resource_group ? data.azurerm_resource_group.existing[0].name : azurerm_resource_group.main[0].name
  resource_group_location = var.use_existing_resource_group ? data.azurerm_resource_group.existing[0].location : azurerm_resource_group.main[0].location
  resource_group_id       = var.use_existing_resource_group ? data.azurerm_resource_group.existing[0].id : azurerm_resource_group.main[0].id
}

# ============================================================================
# AZURE KEY VAULT (OPTIONAL)
# ============================================================================
# Key Vault securely stores secrets, keys, and certificates.
# Used to store SharePoint site information and credentials.
# ============================================================================

resource "azurerm_key_vault" "main" {
  count = var.create_key_vault ? 1 : 0

  name                = local.key_vault_name
  location            = local.resource_group_location
  resource_group_name = local.resource_group_name
  tenant_id           = data.azurerm_client_config.current.tenant_id
  sku_name            = "standard"

  # Security settings
  enabled_for_deployment          = false
  enabled_for_disk_encryption     = false
  enabled_for_template_deployment = false
  enable_rbac_authorization       = true
  purge_protection_enabled        = var.enable_soft_delete_protection
  soft_delete_retention_days      = 90

  # Network rules - allow Azure services and current client
  network_acls {
    default_action = "Allow"
    bypass         = "AzureServices"
  }

  tags = local.common_tags

  # Depend on resource group (either new or existing)
  depends_on = [
    azurerm_resource_group.main,
    data.azurerm_resource_group.existing
  ]
}

# ============================================================================
# KEY VAULT ROLE ASSIGNMENT
# ============================================================================
# Grant the current user/service principal access to the Key Vault.
# ============================================================================

resource "azurerm_role_assignment" "key_vault_admin" {
  count = var.create_key_vault ? 1 : 0

  scope                = azurerm_key_vault.main[0].id
  role_definition_name = "Key Vault Administrator"
  principal_id         = data.azurerm_client_config.current.object_id
}

# ============================================================================
# SHAREPOINT SITE CREATION
# ============================================================================
# SharePoint sites are created using Microsoft Graph API via PowerShell.
# This uses a null_resource with local-exec provisioner.
#
# WHY POWERSHELL?
# - No official Terraform provider for SharePoint Online
# - Microsoft Graph API is the supported way to create SharePoint sites
# - PowerShell has excellent Microsoft 365 integration
# ============================================================================

resource "null_resource" "sharepoint_sites" {
  for_each = var.sharepoint_sites

  # Re-run if site configuration changes
  triggers = {
    site_name    = each.key
    display_name = each.value.display_name
    description  = each.value.description
    template     = each.value.template
    visibility   = each.value.visibility
    owners       = join(",", each.value.owners)
    members      = join(",", each.value.members)
  }

  # Create the SharePoint site using PowerShell
  provisioner "local-exec" {
    command = <<-EOT
      powershell -ExecutionPolicy Bypass -File "${path.module}/../scripts/Create-SharePointSite.ps1" `
        -SiteName "${each.key}" `
        -DisplayName "${each.value.display_name}" `
        -Description "${each.value.description}" `
        -Template "${each.value.template}" `
        -Visibility "${each.value.visibility}" `
        -TenantName "${var.m365_tenant_name}" `
        -AdminEmail "${var.sharepoint_admin_email}" `
        -Owners "${join(",", distinct(concat([lower(var.sharepoint_admin_email)], [for owner in each.value.owners : lower(owner)])))}"${length(each.value.members) > 0 ? " `\n        -Members \"${join(",", each.value.members)}\"" : ""}
    EOT

    interpreter = ["powershell", "-Command"]
  }

  # Depend on resource group (either new or existing)
  depends_on = [
    azurerm_resource_group.main,
    data.azurerm_resource_group.existing
  ]
}

# ============================================================================
# STORE SITE URLS IN KEY VAULT
# ============================================================================
# Store the SharePoint site URLs in Key Vault for reference.
# ============================================================================

resource "azurerm_key_vault_secret" "site_urls" {
  for_each = var.create_key_vault ? var.sharepoint_sites : {}

  name         = "sharepoint-site-${each.key}"
  value        = "${local.sharepoint_url}/sites/${local.site_url_segments[each.key]}"
  key_vault_id = azurerm_key_vault.main[0].id

  tags = {
    SiteName    = each.key
    DisplayName = each.value.display_name
    ManagedBy   = "Terraform"
  }

  depends_on = [
    azurerm_role_assignment.key_vault_admin,
    null_resource.sharepoint_sites
  ]
}

# ============================================================================
# TIME DELAY FOR SITE PROVISIONING
# ============================================================================
# SharePoint sites take time to fully provision. This adds a delay
# to ensure sites are ready before any dependent operations.
# ============================================================================

resource "time_sleep" "wait_for_sites" {
  depends_on = [null_resource.sharepoint_sites]

  create_duration = "60s"
}
