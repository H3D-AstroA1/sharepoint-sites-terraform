<#
.SYNOPSIS
    Creates a SharePoint Online site using Microsoft Graph API.

.DESCRIPTION
    This script creates a SharePoint site using the Microsoft Graph API.
    It requires Azure CLI to be installed and logged in for authentication.

.PARAMETER SiteName
    The URL-friendly name for the site (e.g., "executive-confidential")

.PARAMETER DisplayName
    The display name for the site (e.g., "Executive Confidential")

.PARAMETER Description
    The description for the site

.PARAMETER Template
    The site template (STS#3 for Team Site, SITEPAGEPUBLISHING#0 for Communication Site)

.PARAMETER Visibility
    The visibility setting (Private or Public)

.PARAMETER TenantName
    The Microsoft 365 tenant name (without .onmicrosoft.com)

.PARAMETER AdminEmail
    The email of the site administrator/owner

.PARAMETER Owners
    Comma-separated list of owner email addresses

.PARAMETER Members
    Comma-separated list of member email addresses

.EXAMPLE
    .\Create-SharePointSite.ps1 -SiteName "test-site" -DisplayName "Test Site" -Description "A test site" -Template "STS#3" -Visibility "Private" -TenantName "contoso" -AdminEmail "admin@contoso.onmicrosoft.com"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteName,
    
    [Parameter(Mandatory=$true)]
    [string]$DisplayName,
    
    [Parameter(Mandatory=$false)]
    [string]$Description = "",
    
    [Parameter(Mandatory=$false)]
    [string]$Template = "STS#3",
    
    [Parameter(Mandatory=$false)]
    [string]$Visibility = "Private",
    
    [Parameter(Mandatory=$true)]
    [string]$TenantName,
    
    [Parameter(Mandatory=$true)]
    [string]$AdminEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$Owners = "",
    
    [Parameter(Mandatory=$false)]
    [string]$Members = ""
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to get access token using Azure CLI
function Get-GraphAccessToken {
    try {
        $token = az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv 2>$null
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrEmpty($token)) {
            throw "Failed to get access token"
        }
        return $token
    }
    catch {
        Write-ColorOutput "ERROR: Failed to get Microsoft Graph access token." "Red"
        Write-ColorOutput "Please ensure you are logged in with: az login" "Yellow"
        exit 1
    }
}

# Function to check if site already exists
function Test-SiteExists {
    param(
        [string]$SiteUrl,
        [string]$AccessToken
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }
    
    try {
        $encodedUrl = [System.Web.HttpUtility]::UrlEncode($SiteUrl)
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($TenantName).sharepoint.com:/sites/$SiteName" -Headers $headers -Method Get -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        return $false
    }
}

# Function to create a Microsoft 365 Group-connected Team Site
function New-TeamSite {
    param(
        [string]$AccessToken
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }
    
    # Create a Microsoft 365 Group (which creates a Team Site)
    $groupBody = @{
        displayName = $DisplayName
        description = $Description
        mailNickname = $SiteName -replace '[^a-zA-Z0-9]', ''
        mailEnabled = $true
        securityEnabled = $false
        groupTypes = @("Unified")
        visibility = $Visibility
    }
    
    # Add owners if specified
    if (-not [string]::IsNullOrEmpty($Owners)) {
        $ownerEmails = $Owners -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        if ($ownerEmails.Count -gt 0) {
            $ownerIds = @()
            foreach ($email in $ownerEmails) {
                try {
                    $user = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$email" -Headers $headers -Method Get -ErrorAction SilentlyContinue
                    if ($user.id) {
                        $ownerIds += "https://graph.microsoft.com/v1.0/users/$($user.id)"
                    }
                }
                catch {
                    Write-ColorOutput "WARNING: Could not find user: $email" "Yellow"
                }
            }
            if ($ownerIds.Count -gt 0) {
                $groupBody["owners@odata.bind"] = $ownerIds
            }
        }
    }
    
    $jsonBody = $groupBody | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups" -Headers $headers -Method Post -Body $jsonBody
        return $response
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $errorMessage = $reader.ReadToEnd()
        }
        throw "Failed to create group: $errorMessage"
    }
}

# Function to create a Communication Site using SharePoint REST API
function New-CommunicationSite {
    param(
        [string]$AccessToken
    )
    
    # Get SharePoint access token
    $spToken = az account get-access-token --resource "https://$TenantName.sharepoint.com" --query accessToken -o tsv 2>$null
    
    if ([string]::IsNullOrEmpty($spToken)) {
        throw "Failed to get SharePoint access token"
    }
    
    $headers = @{
        "Authorization" = "Bearer $spToken"
        "Content-Type" = "application/json"
        "Accept" = "application/json"
    }
    
    $siteUrl = "https://$TenantName.sharepoint.com/sites/$SiteName"
    
    $body = @{
        request = @{
            Title = $DisplayName
            Url = $siteUrl
            Description = $Description
            Classification = ""
            SiteDesignId = "00000000-0000-0000-0000-000000000000"
            WebTemplate = "SITEPAGEPUBLISHING#0"
            WebTemplateExtensionId = "00000000-0000-0000-0000-000000000000"
        }
    }
    
    $jsonBody = $body | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri "https://$TenantName.sharepoint.com/_api/SPSiteManager/create" -Headers $headers -Method Post -Body $jsonBody
        return $response
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $errorMessage = $reader.ReadToEnd()
        }
        throw "Failed to create communication site: $errorMessage"
    }
}

# Main execution
Write-ColorOutput "========================================" "Cyan"
Write-ColorOutput "SharePoint Site Creation Script" "Cyan"
Write-ColorOutput "========================================" "Cyan"
Write-ColorOutput ""
Write-ColorOutput "Site Name:    $SiteName" "White"
Write-ColorOutput "Display Name: $DisplayName" "White"
Write-ColorOutput "Template:     $Template" "White"
Write-ColorOutput "Visibility:   $Visibility" "White"
Write-ColorOutput "Tenant:       $TenantName" "White"
Write-ColorOutput ""

# Get access token
Write-ColorOutput "Getting Microsoft Graph access token..." "Yellow"
$accessToken = Get-GraphAccessToken
Write-ColorOutput "Access token obtained successfully." "Green"

# Check if site already exists
$siteUrl = "https://$TenantName.sharepoint.com/sites/$SiteName"
Write-ColorOutput "Checking if site already exists..." "Yellow"

if (Test-SiteExists -SiteUrl $siteUrl -AccessToken $accessToken) {
    Write-ColorOutput "Site already exists: $siteUrl" "Yellow"
    Write-ColorOutput "Skipping creation." "Yellow"
    exit 0
}

# Create the site based on template
Write-ColorOutput "Creating SharePoint site..." "Yellow"

try {
    if ($Template -eq "SITEPAGEPUBLISHING#0") {
        # Communication Site
        Write-ColorOutput "Creating Communication Site..." "Yellow"
        $result = New-CommunicationSite -AccessToken $accessToken
        Write-ColorOutput "Communication Site created successfully!" "Green"
    }
    else {
        # Team Site (STS#3 or GROUP#0)
        Write-ColorOutput "Creating Team Site (Microsoft 365 Group)..." "Yellow"
        $result = New-TeamSite -AccessToken $accessToken
        Write-ColorOutput "Team Site created successfully!" "Green"
        
        # Wait for site provisioning
        Write-ColorOutput "Waiting for site provisioning (30 seconds)..." "Yellow"
        Start-Sleep -Seconds 30
    }
    
    Write-ColorOutput "" "White"
    Write-ColorOutput "========================================" "Green"
    Write-ColorOutput "Site Created Successfully!" "Green"
    Write-ColorOutput "========================================" "Green"
    Write-ColorOutput "URL: $siteUrl" "Cyan"
    Write-ColorOutput "" "White"
    
    exit 0
}
catch {
    Write-ColorOutput "" "White"
    Write-ColorOutput "========================================" "Red"
    Write-ColorOutput "ERROR: Failed to create site" "Red"
    Write-ColorOutput "========================================" "Red"
    Write-ColorOutput $_.Exception.Message "Red"
    Write-ColorOutput "" "White"
    
    # Don't fail the Terraform apply - just warn
    # This allows the deployment to continue even if some sites fail
    Write-ColorOutput "WARNING: Site creation failed but continuing deployment." "Yellow"
    exit 0
}
