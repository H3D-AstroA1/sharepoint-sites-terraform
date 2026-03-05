# ============================================================================
# VARIABLES DEFINITION FILE
# ============================================================================
# This file defines all the variables used in the Terraform configuration.
# 
# HOW TO USE THIS FILE:
# - This file DEFINES variables (what they are, their type, description)
# - The actual VALUES are set in terraform.tfvars
# - DO NOT put sensitive values directly in this file
#
# VARIABLE TYPES:
# - string: Text values (e.g., "uksouth")
# - number: Numeric values (e.g., 30)
# - bool: True/false values
# - list: A list of values (e.g., ["item1", "item2"])
# - map: Key-value pairs (e.g., {key1 = "value1", key2 = "value2"})
# - object: Complex structured data
# ============================================================================

# ============================================================================
# AZURE CONFIGURATION VARIABLES
# ============================================================================
# These variables configure which Azure tenant, subscription, and resource
# group to deploy to. You MUST provide values for these.
# ============================================================================

variable "azure_tenant_id" {
  type        = string
  description = <<-EOT
    The Azure Tenant ID (also called Directory ID).
    
    WHERE TO FIND THIS:
    1. Go to Azure Portal (https://portal.azure.com)
    2. Search for "Microsoft Entra ID" (or "Azure Active Directory")
    3. Look for "Tenant ID" on the Overview page
    
    FORMAT: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx (GUID format)
  EOT

  validation {
    condition     = can(regex("^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", var.azure_tenant_id))
    error_message = "The azure_tenant_id must be a valid GUID format (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)."
  }
}

variable "azure_subscription_id" {
  type        = string
  description = <<-EOT
    The Azure Subscription ID where Azure resources will be created.
    
    WHERE TO FIND THIS:
    1. Go to Azure Portal (https://portal.azure.com)
    2. Search for "Subscriptions"
    3. Select your subscription
    4. Copy the "Subscription ID"
    
    FORMAT: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx (GUID format)
  EOT

  validation {
    condition     = can(regex("^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", var.azure_subscription_id))
    error_message = "The azure_subscription_id must be a valid GUID format (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)."
  }
}

variable "resource_group_name" {
  type        = string
  description = <<-EOT
    The name of the Azure Resource Group to create or use.
    
    NAMING CONVENTIONS:
    - Use lowercase letters, numbers, and hyphens
    - Start with a letter
    - Recommended prefix: "rg-" (e.g., "rg-sharepoint-automation")
    
    EXAMPLE: "rg-sharepoint-sites-prod"
  EOT
  default     = "rg-sharepoint-automation"

  validation {
    condition     = can(regex("^[a-zA-Z][a-zA-Z0-9-_]{1,88}[a-zA-Z0-9]$", var.resource_group_name))
    error_message = "Resource group name must be 3-90 characters, start with a letter, and contain only letters, numbers, hyphens, and underscores."
  }
}

variable "location" {
  type        = string
  description = <<-EOT
    The Azure region where resources will be deployed.
    
    COMMON UK/EUROPE REGIONS:
    - uksouth (UK South - London)
    - ukwest (UK West - Cardiff)
    - northeurope (North Europe - Ireland)
    - westeurope (West Europe - Netherlands)
    
    COMMON US REGIONS:
    - eastus (East US - Virginia)
    - westus2 (West US 2 - Washington)
    
    TIP: Choose a region close to your users for better performance.
  EOT
  default     = "westus2"

  validation {
    condition = contains([
      "uksouth", "ukwest", "northeurope", "westeurope",
      "eastus", "eastus2", "westus", "westus2", "centralus",
      "australiaeast", "southeastasia", "japaneast"
    ], var.location)
    error_message = "Please choose a valid Azure region."
  }
}

# ============================================================================
# MICROSOFT 365 / SHAREPOINT CONFIGURATION
# ============================================================================
# These variables configure the Microsoft 365 tenant and SharePoint settings.
# ============================================================================

variable "m365_tenant_name" {
  type        = string
  description = <<-EOT
    Your Microsoft 365 tenant name (without .onmicrosoft.com).
    
    WHERE TO FIND THIS:
    1. Go to Microsoft 365 Admin Center (https://admin.microsoft.com)
    2. Look at your domain - it's the part before .onmicrosoft.com
    
    EXAMPLE: If your domain is "contoso.onmicrosoft.com", enter "contoso"
    
    This is used to construct SharePoint URLs like:
    https://contoso.sharepoint.com/sites/your-site
  EOT

  validation {
    condition     = can(regex("^[a-zA-Z][a-zA-Z0-9-]{1,62}$", var.m365_tenant_name))
    error_message = "Tenant name must be 2-63 characters, start with a letter, and contain only letters, numbers, and hyphens."
  }
}

variable "sharepoint_admin_email" {
  type        = string
  description = <<-EOT
    Email address of the SharePoint administrator who will own the sites.
    
    REQUIREMENTS:
    - Must be a valid user in your Microsoft 365 tenant
    - Must have SharePoint Admin or Global Admin role
    
    EXAMPLE: "admin@contoso.onmicrosoft.com"
  EOT

  validation {
    condition     = can(regex("^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$", var.sharepoint_admin_email))
    error_message = "Please provide a valid email address."
  }
}

# ============================================================================
# SHAREPOINT SITES CONFIGURATION
# ============================================================================
# Configuration for the 4 SharePoint sites to be created.
# ============================================================================

variable "sharepoint_sites" {
  type = map(object({
    display_name = string
    description  = string
    template     = string
    visibility   = string
    owners       = list(string)
    members      = list(string)
  }))
  description = <<-EOT
    Configuration for SharePoint sites to create.
    
    SITE TEMPLATES:
    - "STS#3"     : Team site (no Microsoft 365 Group)
    - "GROUP#0"   : Team site (with Microsoft 365 Group)
    - "SITEPAGEPUBLISHING#0" : Communication site
    
    VISIBILITY OPTIONS:
    - "Private"   : Only members can access
    - "Public"    : Anyone in the organization can access
    
    EXAMPLE:
    {
      "site-name" = {
        display_name = "Site Display Name"
        description  = "Site description"
        template     = "STS#3"
        visibility   = "Private"
        owners       = ["owner@domain.com"]
        members      = ["member1@domain.com", "member2@domain.com"]
      }
    }
  EOT

  default = {
    "executive-confidential" = {
      display_name = "Executive Confidential"
      description  = "Confidential documents and communications for executive leadership team"
      template     = "STS#3"
      visibility   = "Private"
      owners       = []
      members      = []
    }
    "finance-internal" = {
      display_name = "Finance Internal"
      description  = "Internal finance team documents, reports, and collaboration space"
      template     = "STS#3"
      visibility   = "Private"
      owners       = []
      members      = []
    }
    "claims-operations" = {
      display_name = "Claims Operations"
      description  = "Claims operations team workspace for processing and documentation"
      template     = "STS#3"
      visibility   = "Private"
      owners       = []
      members      = []
    }
    "it-infra" = {
      display_name = "IT Infrastructure"
      description  = "IT infrastructure documentation, runbooks, and technical resources"
      template     = "STS#3"
      visibility   = "Private"
      owners       = []
      members      = []
    }
  }
}

# ============================================================================
# OPTIONAL: SERVICE PRINCIPAL AUTHENTICATION
# ============================================================================
# Use these variables if authenticating with a Service Principal instead of
# Azure CLI. Leave empty to use Azure CLI authentication.
# ============================================================================

variable "use_service_principal" {
  type        = bool
  description = <<-EOT
    Set to true to use Service Principal authentication instead of Azure CLI.
    
    WHEN TO USE SERVICE PRINCIPAL:
    - CI/CD pipelines (Azure DevOps, GitHub Actions)
    - Automated deployments without user interaction
    - Production environments
    
    WHEN TO USE AZURE CLI:
    - Local development
    - Testing and experimentation
    - When you want to use your own credentials
  EOT
  default     = false
}

variable "azure_client_id" {
  type        = string
  description = <<-EOT
    The Application (Client) ID of the Service Principal.
    Only required if use_service_principal = true.
    
    WHERE TO FIND THIS:
    1. Go to Azure Portal > Microsoft Entra ID > App registrations
    2. Select your app registration
    3. Copy the "Application (client) ID"
  EOT
  default     = ""
  sensitive   = true
}

variable "azure_client_secret" {
  type        = string
  description = <<-EOT
    The Client Secret of the Service Principal.
    Only required if use_service_principal = true.
    
    ⚠️ SECURITY WARNING: Never commit this value to source control!
    Use environment variables or Azure Key Vault instead.
  EOT
  default     = ""
  sensitive   = true
}

# ============================================================================
# OPTIONAL: AZURE KEY VAULT CONFIGURATION
# ============================================================================
# Configure Azure Key Vault for secure secret storage.
# ============================================================================

variable "create_key_vault" {
  type        = bool
  description = <<-EOT
    Whether to create an Azure Key Vault for storing secrets.
    
    RECOMMENDED: true for production environments
    
    The Key Vault will store:
    - Service Principal credentials (if used)
    - SharePoint site URLs and metadata
  EOT
  default     = true
}

variable "key_vault_name" {
  type        = string
  description = <<-EOT
    Name for the Azure Key Vault.
    
    NAMING RULES:
    - Must be globally unique across all of Azure
    - 3-24 characters
    - Letters, numbers, and hyphens only
    - Must start with a letter
    
    TIP: Include a unique identifier like your company name or project code.
  EOT
  default     = ""

  validation {
    condition     = var.key_vault_name == "" || can(regex("^[a-zA-Z][a-zA-Z0-9-]{1,22}[a-zA-Z0-9]$", var.key_vault_name))
    error_message = "Key Vault name must be 3-24 characters, start with a letter, and contain only letters, numbers, and hyphens."
  }
}

# ============================================================================
# TAGS
# ============================================================================
# Tags help organize and track Azure resources.
# ============================================================================

variable "tags" {
  type        = map(string)
  description = <<-EOT
    Tags to apply to all Azure resources.
    
    RECOMMENDED TAGS:
    - Environment: prod, dev, test, staging
    - Project: Project name or code
    - Owner: Team or person responsible
    - CostCenter: For billing purposes
    
    EXAMPLE:
    {
      Environment = "Production"
      Project     = "SharePoint-Automation"
      Owner       = "IT-Team"
      CostCenter  = "IT-001"
    }
  EOT
  default = {
    Environment = "Production"
    Project     = "SharePoint-Sites-Automation"
    ManagedBy   = "Terraform"
    Purpose     = "SharePoint-Site-Provisioning"
  }
}

# ============================================================================
# DEPLOYMENT OPTIONS
# ============================================================================
# Additional options to control the deployment behavior.
# ============================================================================

variable "enable_soft_delete_protection" {
  type        = bool
  description = <<-EOT
    Enable soft delete protection for Key Vault.
    
    WHAT THIS DOES:
    - When enabled, deleted Key Vaults are retained for 90 days
    - Allows recovery of accidentally deleted vaults
    
    RECOMMENDED: true for production, false for development/testing
  EOT
  default     = true
}

variable "deployment_timeout_minutes" {
  type        = number
  description = <<-EOT
    Maximum time (in minutes) to wait for SharePoint site creation.
    
    SharePoint site provisioning can take 5-15 minutes.
    Increase this value if you experience timeout errors.
  EOT
  default     = 30

  validation {
    condition     = var.deployment_timeout_minutes >= 5 && var.deployment_timeout_minutes <= 60
    error_message = "Deployment timeout must be between 5 and 60 minutes."
  }
}
