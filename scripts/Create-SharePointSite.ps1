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
        # Extract site path from URL
        $uri = [System.Uri]$SiteUrl
        $sitePath = $uri.AbsolutePath
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($TenantName).sharepoint.com:$sitePath" -Headers $headers -Method Get -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Function to get group's SharePoint site URL
function Get-GroupSiteUrl {
    param(
        [string]$GroupId,
        [string]$AccessToken,
        [int]$MaxRetries = 12,
        [int]$RetryDelaySeconds = 10
    )
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }
    
    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            Write-ColorOutput "  Checking for SharePoint site (attempt $i of $MaxRetries)..." "Yellow"
            $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/sites/root" -Headers $headers -Method Get -ErrorAction Stop
            if ($response.webUrl) {
                return $response.webUrl
            }
        }
        catch {
            if ($i -lt $MaxRetries) {
                Write-ColorOutput "  Site not ready yet, waiting $RetryDelaySeconds seconds..." "Yellow"
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
    }
    
    return $null
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
    
    # Use the SiteName directly as mailNickname (Graph API will sanitize it)
    # But we need to ensure it's valid - remove invalid characters
    $mailNickname = $SiteName -replace '[^a-zA-Z0-9]', ''
    
    # Create a Microsoft 365 Group (which creates a Team Site)
    $groupBody = @{
        displayName = $DisplayName
        description = $Description
        mailNickname = $mailNickname
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
        Write-ColorOutput "Creating Microsoft 365 Group..." "Yellow"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups" -Headers $headers -Method Post -Body $jsonBody
        
        if ($response.id) {
            Write-ColorOutput "Group created successfully! Group ID: $($response.id)" "Green"
            
            # Wait for SharePoint site to be provisioned
            Write-ColorOutput "Waiting for SharePoint site provisioning..." "Yellow"
            $siteUrl = Get-GroupSiteUrl -GroupId $response.id -AccessToken $AccessToken
            
            if ($siteUrl) {
                Write-ColorOutput "SharePoint site provisioned successfully!" "Green"
                return @{
                    GroupId = $response.id
                    SiteUrl = $siteUrl
                    Success = $true
                }
            }
            else {
                Write-ColorOutput "WARNING: Group created but SharePoint site not yet available." "Yellow"
                Write-ColorOutput "The site may still be provisioning. Check SharePoint Admin Center." "Yellow"
                return @{
                    GroupId = $response.id
                    SiteUrl = "https://$TenantName.sharepoint.com/sites/$mailNickname"
                    Success = $true
                    Pending = $true
                }
            }
        }
        else {
            throw "Group creation response did not contain an ID"
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $errorMessage = $reader.ReadToEnd()
            }
            catch {}
        }
        throw "Failed to create group: $errorMessage"
    }
}

# Function to create a Communication Site using SharePoint REST API
function New-CommunicationSite {
    param(
        [string]$AccessToken
    )
    
    $siteUrl = "https://$TenantName.sharepoint.com/sites/$SiteName"
    
    # Method 1: Try using SharePoint REST API with SharePoint-specific token
    Write-ColorOutput "Attempting to create Communication Site via SharePoint API..." "Yellow"
    
    try {
        # Get SharePoint access token
        $spToken = az account get-access-token --resource "https://$TenantName.sharepoint.com" --query accessToken -o tsv 2>$null
        
        if (-not [string]::IsNullOrEmpty($spToken)) {
            $headers = @{
                "Authorization" = "Bearer $spToken"
                "Content-Type" = "application/json;odata=verbose"
                "Accept" = "application/json;odata=verbose"
            }
            
            $body = @{
                request = @{
                    Title = $DisplayName
                    Url = $siteUrl
                    Description = $Description
                    Lcid = 1033
                    ShareByEmailEnabled = $false
                    Classification = ""
                    WebTemplate = "SITEPAGEPUBLISHING#0"
                    SiteDesignId = "00000000-0000-0000-0000-000000000000"
                    WebTemplateExtensionId = "00000000-0000-0000-0000-000000000000"
                    Owner = $AdminEmail
                }
            }
            
            $jsonBody = $body | ConvertTo-Json -Depth 10
            
            $response = Invoke-RestMethod -Uri "https://$TenantName.sharepoint.com/_api/SPSiteManager/create" -Headers $headers -Method Post -Body $jsonBody -ErrorAction Stop
            
            if ($response.SiteId -or ($response.d -and $response.d.SiteId)) {
                $siteId = if ($response.SiteId) { $response.SiteId } else { $response.d.SiteId }
                Write-ColorOutput "Communication Site created successfully! Site ID: $siteId" "Green"
                return @{
                    SiteId = $siteId
                    SiteUrl = $siteUrl
                    Success = $true
                }
            }
            elseif ($response.SiteStatus -eq 2 -or ($response.d -and $response.d.SiteStatus -eq 2)) {
                Write-ColorOutput "Communication Site created successfully!" "Green"
                return @{
                    SiteUrl = $siteUrl
                    Success = $true
                }
            }
        }
    }
    catch {
        $errorDetail = $_.Exception.Message
        Write-ColorOutput "SharePoint API method failed: $errorDetail" "Yellow"
        Write-ColorOutput "Trying alternative method..." "Yellow"
    }
    
    # Method 2: Create as a Team Site (fallback) since Communication Sites require SharePoint Admin permissions
    Write-ColorOutput "Creating as Team Site (Communication Site requires SharePoint Admin permissions)..." "Yellow"
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }
    
    # Create a Microsoft 365 Group but configure it to be more like a communication site
    $mailNickname = ($SiteName -replace '[^a-zA-Z0-9]', '') + "site"
    
    $groupBody = @{
        displayName = $DisplayName
        description = $Description
        mailNickname = $mailNickname
        mailEnabled = $true
        securityEnabled = $false
        groupTypes = @("Unified")
        visibility = $Visibility
        resourceBehaviorOptions = @("HideGroupInOutlook", "WelcomeEmailDisabled", "SubscribeMembersToCalendarEventsDisabled")
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
        Write-ColorOutput "Creating Microsoft 365 Group (fallback)..." "Yellow"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups" -Headers $headers -Method Post -Body $jsonBody
        
        if ($response.id) {
            Write-ColorOutput "Group created! Waiting for SharePoint site..." "Green"
            
            # Wait for SharePoint site to be provisioned
            $actualSiteUrl = Get-GroupSiteUrl -GroupId $response.id -AccessToken $AccessToken
            
            if ($actualSiteUrl) {
                Write-ColorOutput "Site provisioned successfully!" "Green"
                return @{
                    GroupId = $response.id
                    SiteUrl = $actualSiteUrl
                    Success = $true
                    IsFallback = $true
                }
            }
            else {
                return @{
                    GroupId = $response.id
                    SiteUrl = "https://$TenantName.sharepoint.com/sites/$mailNickname"
                    Success = $true
                    Pending = $true
                    IsFallback = $true
                }
            }
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $errorMessage = $reader.ReadToEnd()
            }
            catch {}
        }
        throw "Failed to create site: $errorMessage"
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

# Generate expected site URL based on template
$mailNickname = $SiteName -replace '[^a-zA-Z0-9]', ''
$expectedSiteUrl = "https://$TenantName.sharepoint.com/sites/$mailNickname"

# Check if site already exists
Write-ColorOutput "Checking if site already exists..." "Yellow"

if (Test-SiteExists -SiteUrl $expectedSiteUrl -AccessToken $accessToken) {
    Write-ColorOutput "Site already exists: $expectedSiteUrl" "Yellow"
    Write-ColorOutput "Skipping creation." "Yellow"
    Write-ColorOutput ""
    Write-ColorOutput "========================================" "Green"
    Write-ColorOutput "Site Already Exists" "Green"
    Write-ColorOutput "========================================" "Green"
    Write-ColorOutput "URL: $expectedSiteUrl" "Cyan"
    Write-ColorOutput ""
    exit 0
}

# Create the site based on template
Write-ColorOutput "Creating SharePoint site..." "Yellow"

try {
    $result = $null
    
    if ($Template -eq "SITEPAGEPUBLISHING#0") {
        # Communication Site
        Write-ColorOutput "Creating Communication Site..." "Yellow"
        $result = New-CommunicationSite -AccessToken $accessToken
    }
    else {
        # Team Site (STS#3 or GROUP#0)
        Write-ColorOutput "Creating Team Site (Microsoft 365 Group)..." "Yellow"
        $result = New-TeamSite -AccessToken $accessToken
    }
    
    if ($result -and $result.Success) {
        Write-ColorOutput "" "White"
        Write-ColorOutput "========================================" "Green"
        Write-ColorOutput "Site Created Successfully!" "Green"
        Write-ColorOutput "========================================" "Green"
        Write-ColorOutput "URL: $($result.SiteUrl)" "Cyan"
        
        if ($result.Pending) {
            Write-ColorOutput "" "White"
            Write-ColorOutput "NOTE: Site is still provisioning. It may take a few minutes to be accessible." "Yellow"
        }
        
        if ($result.IsFallback) {
            Write-ColorOutput "" "White"
            Write-ColorOutput "NOTE: Created as Team Site (Communication Site requires SharePoint Admin permissions)." "Yellow"
        }
        
        Write-ColorOutput "" "White"
        exit 0
    }
    else {
        throw "Site creation did not return a success result"
    }
}
catch {
    Write-ColorOutput "" "White"
    Write-ColorOutput "========================================" "Red"
    Write-ColorOutput "ERROR: Failed to create site" "Red"
    Write-ColorOutput "========================================" "Red"
    Write-ColorOutput $_.Exception.Message "Red"
    Write-ColorOutput "" "White"
    
    # Exit with error code so Terraform knows it failed
    exit 1
}
