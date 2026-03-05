# ============================================================================
# PROVIDERS CONFIGURATION
# ============================================================================
# This file configures the Terraform providers needed for SharePoint deployment.
# 
# WHAT ARE PROVIDERS?
# Providers are plugins that Terraform uses to interact with cloud platforms
# and other services. We need two providers:
# 1. azurerm - For creating Azure resources (Key Vault, Resource Group, etc.)
# 2. azuread - For interacting with Microsoft Entra ID (Azure AD)
# 
# NOTE: SharePoint sites are created via Microsoft Graph API using a null_resource
# with local-exec provisioner, as there's no official SharePoint Terraform provider.
# ============================================================================

terraform {
  # Specify the minimum Terraform version required
  required_version = ">= 1.5.0"

  # Define required providers and their versions
  required_providers {
    # Azure Resource Manager provider - for Azure resources
    azurerm = {
      source  = "hashicorp/azurerm"
      version = "~> 3.85.0"
    }

    # Azure Active Directory provider - for Entra ID operations
    azuread = {
      source  = "hashicorp/azuread"
      version = "~> 2.47.0"
    }

    # Random provider - for generating unique names
    random = {
      source  = "hashicorp/random"
      version = "~> 3.6.0"
    }

    # Null provider - for running local scripts
    null = {
      source  = "hashicorp/null"
      version = "~> 3.2.0"
    }

    # Time provider - for adding delays if needed
    time = {
      source  = "hashicorp/time"
      version = "~> 0.10.0"
    }
  }
}

# ============================================================================
# AZURE RESOURCE MANAGER PROVIDER
# ============================================================================
# This provider manages Azure resources like Resource Groups, Key Vault, etc.
#
# AUTHENTICATION OPTIONS (in order of preference):
# 1. Azure CLI (az login) - Recommended for local development
# 2. Service Principal - Recommended for CI/CD pipelines
# 3. Managed Identity - For Azure-hosted automation
# ============================================================================

provider "azurerm" {
  features {
    # Key Vault configuration
    key_vault {
      # Don't purge Key Vault on destroy (keeps soft-deleted for recovery)
      purge_soft_delete_on_destroy    = false
      recover_soft_deleted_key_vaults = true
    }

    # Resource Group configuration
    resource_group {
      # Prevent accidental deletion of resource groups with resources
      prevent_deletion_if_contains_resources = true
    }
  }

  # Use the subscription ID from variables
  subscription_id = var.azure_subscription_id
  tenant_id       = var.azure_tenant_id

  # Optional: Use Service Principal authentication
  # Uncomment these lines if using Service Principal instead of Azure CLI
  # client_id       = var.azure_client_id
  # client_secret   = var.azure_client_secret
}

# ============================================================================
# AZURE ACTIVE DIRECTORY PROVIDER
# ============================================================================
# This provider interacts with Microsoft Entra ID (formerly Azure AD).
# Used for managing app registrations, service principals, and groups.
# ============================================================================

provider "azuread" {
  tenant_id = var.azure_tenant_id

  # Optional: Use Service Principal authentication
  # Uncomment these lines if using Service Principal instead of Azure CLI
  # client_id     = var.azure_client_id
  # client_secret = var.azure_client_secret
}

# ============================================================================
# RANDOM PROVIDER
# ============================================================================
# Used to generate unique suffixes for resource names to avoid conflicts.
# ============================================================================

provider "random" {
  # No configuration needed
}

# ============================================================================
# NULL PROVIDER
# ============================================================================
# Used to run local scripts (PowerShell/Bash) for SharePoint site creation.
# SharePoint sites are created via Microsoft Graph API calls.
# ============================================================================

provider "null" {
  # No configuration needed
}

# ============================================================================
# TIME PROVIDER
# ============================================================================
# Used to add delays between operations if needed (e.g., waiting for
# Azure AD replication or SharePoint site provisioning).
# ============================================================================

provider "time" {
  # No configuration needed
}
