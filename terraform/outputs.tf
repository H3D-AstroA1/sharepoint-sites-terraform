# ============================================================================
# OUTPUTS DEFINITION FILE
# ============================================================================
# Outputs display useful information after Terraform applies the configuration.
# These values are shown in the terminal and can be used by other Terraform
# configurations or scripts.
#
# HOW TO VIEW OUTPUTS:
# - After 'terraform apply': Outputs are displayed automatically
# - Anytime: Run 'terraform output'
# - Specific output: Run 'terraform output <output_name>'
# - JSON format: Run 'terraform output -json'
# ============================================================================

# ============================================================================
# AZURE RESOURCE OUTPUTS
# ============================================================================

output "resource_group_name" {
  description = "The name of the Azure Resource Group"
  value       = local.resource_group_name
}

output "resource_group_id" {
  description = "The Azure Resource ID of the Resource Group"
  value       = local.resource_group_id
}

output "resource_group_location" {
  description = "The Azure region where resources are deployed"
  value       = local.resource_group_location
}

# ============================================================================
# KEY VAULT OUTPUTS
# ============================================================================

output "key_vault_name" {
  description = "The name of the Azure Key Vault (if created)"
  value       = var.create_key_vault ? azurerm_key_vault.main[0].name : "Key Vault not created"
}

output "key_vault_uri" {
  description = "The URI of the Azure Key Vault (if created)"
  value       = var.create_key_vault ? azurerm_key_vault.main[0].vault_uri : "Key Vault not created"
}

output "key_vault_id" {
  description = "The Azure Resource ID of the Key Vault (if created)"
  value       = var.create_key_vault ? azurerm_key_vault.main[0].id : "Key Vault not created"
}

# ============================================================================
# SHAREPOINT SITE OUTPUTS
# ============================================================================

output "sharepoint_tenant_url" {
  description = "The base SharePoint Online URL for your tenant"
  value       = local.sharepoint_url
}

output "sharepoint_admin_url" {
  description = "The SharePoint Admin Center URL"
  value       = local.sharepoint_admin_url
}

output "sharepoint_sites_created" {
  description = "List of SharePoint sites that were created"
  value = {
    for site_name, site_config in var.sharepoint_sites : site_name => {
      display_name = site_config.display_name
      url          = "${local.sharepoint_url}/sites/${site_name}"
      description  = site_config.description
      template     = site_config.template
      visibility   = site_config.visibility
    }
  }
}

output "sharepoint_site_urls" {
  description = "Direct URLs to each SharePoint site"
  value = {
    for site_name, site_config in var.sharepoint_sites : site_name => "${local.sharepoint_url}/sites/${site_name}"
  }
}

# ============================================================================
# SUMMARY OUTPUT
# ============================================================================

output "deployment_summary" {
  description = "Summary of the deployment"
  value       = <<-EOT

    ╔══════════════════════════════════════════════════════════════════════════════╗
    ║                    SHAREPOINT SITES DEPLOYMENT SUMMARY                       ║
    ╠══════════════════════════════════════════════════════════════════════════════╣
    ║                                                                              ║
    ║  AZURE RESOURCES:                                                            ║
    ║  ├── Resource Group: ${local.resource_group_name}
    ║  ├── Location: ${local.resource_group_location}
    ║  └── Key Vault: ${var.create_key_vault ? azurerm_key_vault.main[0].name : "Not created"}
    ║                                                                              ║
    ║  SHAREPOINT SITES CREATED:                                                   ║
    ║  ├── executive-confidential                                                  ║
    ║  │   └── ${local.sharepoint_url}/sites/executive-confidential
    ║  ├── finance-internal                                                        ║
    ║  │   └── ${local.sharepoint_url}/sites/finance-internal
    ║  ├── claims-operations                                                       ║
    ║  │   └── ${local.sharepoint_url}/sites/claims-operations
    ║  └── it-infra                                                                ║
    ║      └── ${local.sharepoint_url}/sites/it-infra
    ║                                                                              ║
    ║  ADMIN CENTER:                                                               ║
    ║  └── ${local.sharepoint_admin_url}
    ║                                                                              ║
    ╚══════════════════════════════════════════════════════════════════════════════╝

  EOT
}

# ============================================================================
# NEXT STEPS OUTPUT
# ============================================================================

output "next_steps" {
  description = "Recommended next steps after deployment"
  value       = <<-EOT

    ╔══════════════════════════════════════════════════════════════════════════════╗
    ║                              NEXT STEPS                                      ║
    ╠══════════════════════════════════════════════════════════════════════════════╣
    ║                                                                              ║
    ║  1. VERIFY SITES:                                                            ║
    ║     Visit each SharePoint site URL to confirm they were created              ║
    ║                                                                              ║
    ║  2. CONFIGURE PERMISSIONS:                                                   ║
    ║     Add users and groups to each site as needed                              ║
    ║     Go to: Site Settings > Site Permissions                                  ║
    ║                                                                              ║
    ║  3. CUSTOMIZE SITES:                                                         ║
    ║     - Add document libraries                                                 ║
    ║     - Configure navigation                                                   ║
    ║     - Apply branding                                                         ║
    ║                                                                              ║
    ║  4. SET UP GOVERNANCE:                                                       ║
    ║     - Configure retention policies                                           ║
    ║     - Set up sensitivity labels                                              ║
    ║     - Enable auditing                                                        ║
    ║                                                                              ║
    ║  5. BACKUP TERRAFORM STATE:                                                  ║
    ║     Consider using Azure Storage for remote state management                 ║
    ║                                                                              ║
    ╚══════════════════════════════════════════════════════════════════════════════╝

  EOT
}
