#!/usr/bin/env python3
"""
SharePoint Sites Management - Main Menu

A unified interface for managing SharePoint sites:
  - Step 0: Check & Install Prerequisites
  - Step 1: Create SharePoint sites (deploy.py)
  - Step 2: Populate sites with files (populate_files.py)
  - Step 3: Delete files/sites (cleanup.py)

Usage:
    python menu.py

Requirements:
    - Python 3.8+
    - Azure CLI (will be installed automatically if missing)
    - Terraform (will be installed automatically if missing)
"""

import json
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Tuple, Optional, Dict, Any

# ============================================================================
# CONFIGURATION
# ============================================================================

SCRIPT_DIR = Path(__file__).parent.resolve()

# ============================================================================
# CONSOLE OUTPUT HELPERS
# ============================================================================

class Colors:
    """ANSI color codes for terminal output."""
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    NC = '\033[0m'
    BOLD = '\033[1m'
    DIM = '\033[2m'

    @classmethod
    def disable(cls) -> None:
        """Disable colors for non-terminal output."""
        cls.RED = cls.GREEN = cls.YELLOW = cls.BLUE = ''
        cls.MAGENTA = cls.CYAN = cls.WHITE = cls.NC = cls.BOLD = cls.DIM = ''

# Disable colors if not a terminal
if not sys.stdout.isatty():
    Colors.disable()

# Sites to exclude (system sites, personal sites, etc. that typically cause 403 errors)
# Note: We only exclude exact system site names, not patterns that might match user-created sites
EXCLUDED_SITE_PATTERNS = [
    "my workspace",      # Personal OneDrive-like workspace
    "designer",          # Microsoft Designer integration site
    "contenttypehub",    # SharePoint content type hub
    "appcatalog",        # SharePoint app catalog
    "team site",         # Default team site
    "communication site", # Root communication site
]


def is_system_site(site: dict) -> bool:
    """Check if a site is a protected system site that cannot be deleted."""
    site_name = site.get("displayName", site.get("name", "")).lower()
    web_url = site.get("webUrl", "").lower()
    
    # Check name patterns
    for pattern in EXCLUDED_SITE_PATTERNS:
        if pattern in site_name:
            return True
    
    # Check URL patterns
    if "/contentstorage/" in web_url:
        return True
    if web_url.endswith(".sharepoint.com") or web_url.endswith(".sharepoint.com/"):
        # Root site collection
        return True
    if "/sites/contenttypehub" in web_url:
        return True
    if "/sites/appcatalog" in web_url:
        return True
    if "/personal/" in web_url:
        return True
    
    return False


def categorize_sites(sites: list) -> tuple:
    """Categorize sites into deletable and system (protected) sites.
    
    Returns:
        Tuple of (deletable_sites, system_sites)
    """
    deletable = []
    system = []
    
    for site in sites:
        if is_system_site(site):
            system.append(site)
        else:
            deletable.append(site)
    
    return deletable, system


def filter_writable_sites(sites: list) -> list:
    """Filter out system/personal sites that typically cause 403 errors.
    
    Note: This function is kept for backward compatibility.
    Use categorize_sites() for more detailed categorization.
    """
    deletable, _ = categorize_sites(sites)
    return deletable


def clear_screen() -> None:
    """Clear the terminal screen."""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_logo() -> None:
    """Print the M365 management logo/banner."""
    print()
    print(f"  {Colors.CYAN}╔══════════════════════════════════════════════════════════════╗{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}                                                              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}███╗   ███╗{Colors.NC}{Colors.BLUE}██████╗ {Colors.NC}{Colors.YELLOW} ██████╗ {Colors.NC}{Colors.MAGENTA}███████╗{Colors.NC}                  {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}████╗ ████║{Colors.NC}{Colors.BLUE}╚════██╗{Colors.NC}{Colors.YELLOW}██╔════╝ {Colors.NC}{Colors.MAGENTA}██╔════╝{Colors.NC}                  {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}██╔████╔██║{Colors.NC}{Colors.BLUE} █████╔╝{Colors.NC}{Colors.YELLOW}███████╗ {Colors.NC}{Colors.MAGENTA}███████╗{Colors.NC}                  {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}██║╚██╔╝██║{Colors.NC}{Colors.BLUE} ╚═══██╗{Colors.NC}{Colors.YELLOW}██╔═══██╗{Colors.NC}{Colors.MAGENTA}╚════██║{Colors.NC}                  {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}██║ ╚═╝ ██║{Colors.NC}{Colors.BLUE}██████╔╝{Colors.NC}{Colors.YELLOW}╚██████╔╝{Colors.NC}{Colors.MAGENTA}███████║{Colors.NC}                  {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}╚═╝     ╚═╝{Colors.NC}{Colors.BLUE}╚═════╝ {Colors.NC}{Colors.YELLOW} ╚═════╝ {Colors.NC}{Colors.MAGENTA}╚══════╝{Colors.NC}                  {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}                                                              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.WHITE}{Colors.BOLD}M365 Environment Population Tool{Colors.NC}                         {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.DIM}SharePoint Sites + Email Mailboxes{Colors.NC}                        {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.DIM}Terraform + Microsoft Graph API{Colors.NC}                           {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}                                                              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}╚══════════════════════════════════════════════════════════════╝{Colors.NC}")
    print()

def print_success(message: str) -> None:
    print(f"  {Colors.GREEN}✓{Colors.NC} {message}")

def print_error(message: str) -> None:
    print(f"  {Colors.RED}✗{Colors.NC} {message}")

def print_warning(message: str) -> None:
    print(f"  {Colors.YELLOW}⚠{Colors.NC} {message}")

def print_info(message: str) -> None:
    print(f"  {Colors.BLUE}ℹ{Colors.NC} {message}")

# ============================================================================
# PREREQUISITES CHECK & INSTALL
# ============================================================================

def get_os_info() -> Tuple[str, str]:
    """Get operating system information."""
    system = platform.system().lower()
    if system == "darwin":
        return "macos", platform.machine()
    elif system == "windows":
        return "windows", platform.machine()
    else:
        return "linux", platform.machine()

# Default Azure CLI installation paths on Windows
AZURE_CLI_PATHS = [
    r"C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
    r"C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
]

def find_azure_cli_path() -> Optional[str]:
    """Find Azure CLI path, checking default installation locations on Windows."""
    # First check if 'az' is in PATH
    az_path = shutil.which("az")
    if az_path:
        return az_path
    
    # On Windows, check .cmd extension
    if platform.system().lower() == "windows":
        az_cmd_path = shutil.which("az.cmd")
        if az_cmd_path:
            return az_cmd_path
        
        # Check default installation paths
        for path in AZURE_CLI_PATHS:
            if os.path.exists(path):
                return path
    
    return None

def command_exists(command: str) -> bool:
    """Check if a command exists and actually works."""
    # Special handling for Azure CLI
    if command == "az":
        az_path = find_azure_cli_path()
        if az_path:
            try:
                subprocess.run(
                    [az_path, "--version"],
                    capture_output=True,
                    text=True,
                    timeout=10
                )
                return True
            except (FileNotFoundError, subprocess.TimeoutExpired, Exception):
                pass
        return False
    
    # First check if the command is in PATH
    if shutil.which(command) is None:
        # On Windows, also check for .cmd and .exe extensions
        if platform.system().lower() == "windows":
            if shutil.which(f"{command}.cmd") is None and shutil.which(f"{command}.exe") is None:
                return False
        else:
            return False
    
    # Actually try to run the command to verify it works
    try:
        subprocess.run(
            [command, "--version"],
            capture_output=True,
            text=True,
            timeout=10
        )
        return True
    except FileNotFoundError:
        return False
    except subprocess.TimeoutExpired:
        return True  # Command exists but took too long
    except Exception:
        return False

def get_command_version(command: str) -> Optional[str]:
    """Get the version of a command."""
    # Special handling for Azure CLI
    if command == "az":
        az_path = find_azure_cli_path()
        if az_path:
            try:
                result = subprocess.run(
                    [az_path, "--version"],
                    capture_output=True,
                    text=True,
                    check=True
                )
                return result.stdout.strip().split('\n')[0]
            except Exception:
                return None
        return None
    
    try:
        result = subprocess.run(
            [command, "--version"],
            capture_output=True,
            text=True,
            check=True
        )
        return result.stdout.strip().split('\n')[0]
    except Exception:
        return None

def check_azure_login() -> bool:
    """Check if user is logged into Azure CLI."""
    az_path = find_azure_cli_path()
    if not az_path:
        return False
    try:
        result = subprocess.run(
            [az_path, "account", "show"],
            capture_output=True,
            text=True,
            check=True
        )
        return True
    except Exception:
        return False


def get_azure_account_info() -> Optional[Dict]:
    """Get the current Azure account information."""
    az_path = find_azure_cli_path()
    if not az_path:
        return None
    try:
        result = subprocess.run(
            [az_path, "account", "show", "--output", "json"],
            capture_output=True,
            text=True,
            check=True
        )
        return json.loads(result.stdout)
    except Exception:
        return None

def install_azure_cli() -> bool:
    """Install Azure CLI based on the operating system."""
    os_type, _ = get_os_info()
    print_info("Installing Azure CLI...")
    
    try:
        if os_type == "windows":
            # Try winget first
            if command_exists("winget"):
                print_info("Using winget to install Azure CLI...")
                subprocess.run(
                    ["winget", "install", "-e", "--id", "Microsoft.AzureCLI"],
                    check=True
                )
                return True
            # Try chocolatey
            elif command_exists("choco"):
                print_info("Using Chocolatey to install Azure CLI...")
                subprocess.run(["choco", "install", "azure-cli", "-y"], check=True)
                return True
            else:
                print_error("Please install Azure CLI manually from:")
                print_info("https://docs.microsoft.com/en-us/cli/azure/install-azure-cli-windows")
                return False
                
        elif os_type == "macos":
            if command_exists("brew"):
                print_info("Using Homebrew to install Azure CLI...")
                subprocess.run(["brew", "install", "azure-cli"], check=True)
                return True
            else:
                print_error("Please install Homebrew first, then run: brew install azure-cli")
                return False
                
        else:  # Linux
            print_info("Installing Azure CLI on Linux...")
            subprocess.run(
                ["curl", "-sL", "https://aka.ms/InstallAzureCLIDeb", "-o", "/tmp/install_az.sh"],
                check=True
            )
            subprocess.run(["sudo", "bash", "/tmp/install_az.sh"], check=True)
            return True
            
    except Exception as e:
        print_error(f"Failed to install Azure CLI: {e}")
        return False

def install_terraform() -> bool:
    """Install Terraform based on the operating system."""
    os_type, arch = get_os_info()
    print_info("Installing Terraform...")
    
    try:
        if os_type == "windows":
            if command_exists("winget"):
                print_info("Using winget to install Terraform...")
                subprocess.run(
                    ["winget", "install", "-e", "--id", "Hashicorp.Terraform"],
                    check=True
                )
                return True
            elif command_exists("choco"):
                print_info("Using Chocolatey to install Terraform...")
                subprocess.run(["choco", "install", "terraform", "-y"], check=True)
                return True
            else:
                print_error("Please install Terraform manually from:")
                print_info("https://www.terraform.io/downloads")
                return False
                
        elif os_type == "macos":
            if command_exists("brew"):
                print_info("Using Homebrew to install Terraform...")
                subprocess.run(["brew", "tap", "hashicorp/tap"], check=True)
                subprocess.run(["brew", "install", "hashicorp/tap/terraform"], check=True)
                return True
            else:
                print_error("Please install Homebrew first, then run: brew install terraform")
                return False
                
        else:  # Linux
            print_info("Installing Terraform on Linux...")
            subprocess.run([
                "sudo", "apt-get", "update"
            ], check=True)
            subprocess.run([
                "sudo", "apt-get", "install", "-y", "gnupg", "software-properties-common"
            ], check=True)
            subprocess.run([
                "wget", "-O-", "https://apt.releases.hashicorp.com/gpg",
                "|", "sudo", "gpg", "--dearmor", "-o",
                "/usr/share/keyrings/hashicorp-archive-keyring.gpg"
            ], shell=True, check=True)
            subprocess.run([
                "sudo", "apt-get", "update"
            ], check=True)
            subprocess.run([
                "sudo", "apt-get", "install", "-y", "terraform"
            ], check=True)
            return True
            
    except Exception as e:
        print_error(f"Failed to install Terraform: {e}")
        return False

# ============================================================================
# MICROSOFT GRAPH API PERMISSIONS
# ============================================================================

# Azure CLI App ID (first-party Microsoft app)
AZURE_CLI_APP_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"

# Our custom app name for M365 environment management (SharePoint + Email)
CUSTOM_APP_NAME = "M365-Environment-Population-Tool"

# Config file to store custom app details
APP_CONFIG_FILE = SCRIPT_DIR / ".app_config.json"

# Microsoft Graph API ID (constant across all tenants)
MICROSOFT_GRAPH_API_ID = "00000003-0000-0000-c000-000000000000"

# Office 365 Exchange Online API ID (for EWS permissions)
EXCHANGE_ONLINE_API_ID = "00000002-0000-0ff1-ce00-000000000000"

# SharePoint Online API ID (for PnP PowerShell delegated permissions)
SHAREPOINT_ONLINE_API_ID = "00000003-0000-0ff1-ce00-000000000000"

# Microsoft Graph API permission IDs (Application permissions)
# These are the GUIDs for the specific permissions we need
GRAPH_PERMISSION_IDS = {
    # SharePoint permissions
    "Sites.Read.All": "332a536c-c7ef-4017-ab91-336970924f0d",
    "Sites.ReadWrite.All": "9492366f-7969-46a4-8d15-ed1a20078fff",
    "Sites.FullControl.All": "a82116e5-55eb-4c41-a434-62fe8a61c773",  # Full control for site deletion
    "Files.ReadWrite.All": "75359482-378d-4052-8f01-80520e7db3cd",
    "Group.Read.All": "5b567255-7703-4780-807c-7be8301ae99b",
    "Group.ReadWrite.All": "62a82d76-70ea-41e2-9197-370581804d09",
    # Mail permissions (for email population)
    "Mail.ReadWrite": "e2a3a72e-5f79-4c64-b1b1-878b674786c9",
    "User.Read.All": "df021288-bdef-4463-88db-98f22de89214"
}

# Exchange Online API permission IDs (for EWS access)
# These permissions allow full mailbox access via EWS
EWS_PERMISSION_IDS = {
    # full_access_as_app - Full access to all mailboxes via EWS
    "full_access_as_app": "dc890d15-9560-4a4c-9b7f-a736ec74ec40"
}

# SharePoint Online API permission IDs (Delegated permissions for PnP PowerShell)
# These are required for PnP PowerShell interactive authentication
SHAREPOINT_PERMISSION_IDS = {
    # AllSites.FullControl - Full control of all site collections (delegated)
    "AllSites.FullControl": "56680e0d-d2a3-4ae1-80d8-3c4f2c4b5f4c",
    # AllSites.Manage - Create, edit, and delete items and lists (delegated)
    "AllSites.Manage": "b3f70a70-8a4b-4f95-9573-d71c496a53f4",
    # AllSites.Write - Edit or delete items in all site collections (delegated)
    "AllSites.Write": "640ddd16-e5b7-4d71-9690-3f4022f5acd3",
    # AllSites.Read - Read items in all site collections (delegated)
    "AllSites.Read": "4e0d77b0-96ba-4398-af14-3baa780278f4",
}

# SharePoint Online API Application permission IDs (for app-only/client credentials access)
# These are required for REST API access without user context
SHAREPOINT_APP_PERMISSION_IDS = {
    # Sites.FullControl.All - Full control of all site collections (application)
    "Sites.FullControl.All": "678536fe-1083-478a-9c59-b99265e6b0d3",
    # Sites.ReadWrite.All - Read and write items in all site collections (application)
    "Sites.ReadWrite.All": "9bff6588-13f2-4c48-bbf2-ddab62256b36",
    # Sites.Read.All - Read items in all site collections (application)
    "Sites.Read.All": "d13f72ca-a275-4b96-b789-48ebcc4da984",
}

# Required permissions for SharePoint and Email operations
REQUIRED_GRAPH_PERMISSIONS = [
    "Sites.Read.All",
    "Sites.ReadWrite.All",
    "Sites.FullControl.All",  # Required for deleting SharePoint sites
    "Files.ReadWrite.All",
    "Group.Read.All",
    "Group.ReadWrite.All",
    "Mail.ReadWrite",
    "User.Read.All"
]

def get_graph_access_token_via_client_credentials(app_config: Dict[str, Any]) -> Optional[str]:
    """Get Microsoft Graph access token using client credentials flow."""
    import urllib.request
    import urllib.parse
    
    app_id = app_config.get("app_id")
    client_secret = app_config.get("client_secret")
    tenant_id = app_config.get("tenant_id")
    
    if not all([app_id, client_secret, tenant_id]):
        return None
    
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    data = urllib.parse.urlencode({
        "client_id": app_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }).encode('utf-8')
    
    try:
        req = urllib.request.Request(token_url, data=data)
        req.add_header("Content-Type", "application/x-www-form-urlencoded")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            if response.status == 200:
                token_data = json.loads(response.read().decode('utf-8'))
                return token_data.get("access_token")
    except Exception:
        pass
    
    return None

def get_graph_access_token() -> Optional[str]:
    """Get Microsoft Graph access token.
    
    First tries to use custom app credentials if available,
    then falls back to Azure CLI.
    """
    # Try custom app credentials first
    app_config = load_app_config()
    if app_config and app_config.get("client_secret"):
        token = get_graph_access_token_via_client_credentials(app_config)
        if token:
            return token
    
    # Fall back to Azure CLI
    az_path = find_azure_cli_path()
    if not az_path:
        return None
    try:
        result = subprocess.run(
            [az_path, "account", "get-access-token",
             "--resource", "https://graph.microsoft.com",
             "--query", "accessToken", "-o", "tsv"],
            capture_output=True,
            text=True,
            check=True
        )
        return result.stdout.strip()
    except Exception:
        return None

def check_graph_permissions() -> dict:
    """Check if the current token has required Microsoft Graph permissions.
    
    Returns a dict with:
        - has_permissions: bool - True if all required permissions are granted
        - can_access_sites: bool - True if we can access SharePoint sites
        - error: str or None - Error message if any
    """
    import urllib.request
    import urllib.error
    
    result = {
        "has_permissions": False,
        "can_access_sites": False,
        "error": None
    }
    
    token = get_graph_access_token()
    if not token:
        result["error"] = "Could not get Graph API token"
        return result
    
    # Try to access SharePoint sites to test permissions
    try:
        url = "https://graph.microsoft.com/v1.0/sites?search=*&$top=1"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=10) as response:
            if response.status == 200:
                result["has_permissions"] = True
                result["can_access_sites"] = True
    except urllib.error.HTTPError as e:
        if e.code == 403:
            result["error"] = "403 Forbidden - Admin consent required"
        elif e.code == 401:
            result["error"] = "401 Unauthorized - Token expired or invalid"
        else:
            result["error"] = f"{e.code} - {e.reason}"
    except Exception as e:
        result["error"] = str(e)
    
    return result

def get_tenant_id() -> Optional[str]:
    """Get the current tenant ID from Azure CLI."""
    account_info = get_azure_account_info()
    if account_info:
        return account_info.get("tenantId")
    return None

# ============================================================================
# CUSTOM APP REGISTRATION FOR SEAMLESS PERMISSIONS
# ============================================================================

def load_app_config() -> Optional[Dict[str, Any]]:
    """Load the custom app configuration from file."""
    if APP_CONFIG_FILE.exists():
        try:
            with open(APP_CONFIG_FILE, 'r') as f:
                return json.load(f)
        except Exception:
            pass
    return None

def save_app_config(config: Dict[str, Any]) -> bool:
    """Save the custom app configuration to file."""
    try:
        with open(APP_CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
        return True
    except Exception as e:
        print_error(f"Failed to save app config: {e}")
        return False

def find_existing_app() -> Optional[Dict[str, Any]]:
    """Check if our custom app already exists in the tenant."""
    az_path = find_azure_cli_path()
    if not az_path:
        return None
    
    try:
        result = subprocess.run(
            [az_path, "ad", "app", "list",
             "--display-name", CUSTOM_APP_NAME,
             "--query", "[0]",
             "-o", "json"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode == 0 and result.stdout.strip() and result.stdout.strip() != "null":
            return json.loads(result.stdout)
    except Exception:
        pass
    
    return None

def create_custom_app() -> Optional[Dict[str, Any]]:
    """Create a custom app registration with required permissions."""
    az_path = find_azure_cli_path()
    if not az_path:
        print_error("Azure CLI not found")
        return None
    
    print()
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}Creating Custom App Registration{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print()
    print(f"  {Colors.WHITE}App Name:{Colors.NC} {CUSTOM_APP_NAME}")
    print()
    
    try:
        # Step 1: Create the app registration with redirect URI
        print(f"  {Colors.DIM}Step 1/4: Creating app registration...{Colors.NC}")
        
        # Create app with redirect URIs:
        # - https://portal.azure.com - for admin consent (shows Azure Portal after consent)
        # - http://localhost - for PnP PowerShell interactive authentication (uses dynamic port)
        create_result = subprocess.run(
            [az_path, "ad", "app", "create",
             "--display-name", CUSTOM_APP_NAME,
             "--sign-in-audience", "AzureADMyOrg",
             "--web-redirect-uris", "https://portal.azure.com", "http://localhost",
             "-o", "json"],
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if create_result.returncode != 0:
            print_error(f"Failed to create app: {create_result.stderr}")
            return None
        
        app_data = json.loads(create_result.stdout)
        app_id = app_data.get("appId")
        object_id = app_data.get("id")
        
        print(f"  {Colors.GREEN}✓{Colors.NC} App created: {app_id}")
        
        # Step 2: Add required API permissions
        print(f"  {Colors.DIM}Step 2/4: Adding API permissions...{Colors.NC}")
        
        for perm_name, perm_id in GRAPH_PERMISSION_IDS.items():
            add_perm_result = subprocess.run(
                [az_path, "ad", "app", "permission", "add",
                 "--id", app_id,
                 "--api", MICROSOFT_GRAPH_API_ID,
                 "--api-permissions", f"{perm_id}=Role"],  # Role = Application permission
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if add_perm_result.returncode == 0:
                print(f"    {Colors.GREEN}✓{Colors.NC} Added: {perm_name}")
            else:
                print(f"    {Colors.YELLOW}⚠{Colors.NC} {perm_name}: {add_perm_result.stderr.strip()[:50]}")
        
        # Add EWS permissions (Exchange Online API)
        print(f"  {Colors.DIM}Adding EWS permissions for email operations...{Colors.NC}")
        for perm_name, perm_id in EWS_PERMISSION_IDS.items():
            add_perm_result = subprocess.run(
                [az_path, "ad", "app", "permission", "add",
                 "--id", app_id,
                 "--api", EXCHANGE_ONLINE_API_ID,
                 "--api-permissions", f"{perm_id}=Role"],  # Role = Application permission
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if add_perm_result.returncode == 0:
                print(f"    {Colors.GREEN}✓{Colors.NC} Added: EWS.{perm_name}")
            else:
                print(f"    {Colors.YELLOW}⚠{Colors.NC} EWS.{perm_name}: {add_perm_result.stderr.strip()[:50]}")
        
        # Add SharePoint Online API permissions (Delegated - for PnP PowerShell)
        print(f"  {Colors.DIM}Adding SharePoint Online delegated permissions for PnP PowerShell...{Colors.NC}")
        for perm_name, perm_id in SHAREPOINT_PERMISSION_IDS.items():
            add_perm_result = subprocess.run(
                [az_path, "ad", "app", "permission", "add",
                 "--id", app_id,
                 "--api", SHAREPOINT_ONLINE_API_ID,
                 "--api-permissions", f"{perm_id}=Scope"],  # Scope = Delegated permission
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if add_perm_result.returncode == 0:
                print(f"    {Colors.GREEN}✓{Colors.NC} Added: SharePoint.{perm_name} (delegated)")
            else:
                print(f"    {Colors.YELLOW}⚠{Colors.NC} SharePoint.{perm_name}: {add_perm_result.stderr.strip()[:50]}")
        
        # Add SharePoint Online application permissions (for REST API app-only access)
        print(f"  {Colors.DIM}Adding SharePoint Online application permissions for REST API...{Colors.NC}")
        for perm_name, perm_id in SHAREPOINT_APP_PERMISSION_IDS.items():
            add_perm_result = subprocess.run(
                [az_path, "ad", "app", "permission", "add",
                 "--id", app_id,
                 "--api", SHAREPOINT_ONLINE_API_ID,
                 "--api-permissions", f"{perm_id}=Role"],  # Role = Application permission
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if add_perm_result.returncode == 0:
                print(f"    {Colors.GREEN}✓{Colors.NC} Added: SharePoint.{perm_name} (application)")
            else:
                print(f"    {Colors.YELLOW}⚠{Colors.NC} SharePoint.{perm_name}: {add_perm_result.stderr.strip()[:50]}")
        
        # Add public client redirect URIs for PnP PowerShell interactive auth
        print(f"  {Colors.DIM}Adding public client redirect URIs for PnP PowerShell...{Colors.NC}")
        public_uris = [
            "http://localhost",
            "https://login.microsoftonline.com/common/oauth2/nativeclient",
            "urn:ietf:wg:oauth:2.0:oob"
        ]
        update_result = subprocess.run(
            [az_path, "ad", "app", "update", "--id", app_id, "--public-client-redirect-uris"] + public_uris,
            capture_output=True,
            text=True,
            timeout=30
        )
        if update_result.returncode == 0:
            print(f"    {Colors.GREEN}✓{Colors.NC} Added public client redirect URIs")
        else:
            print(f"    {Colors.YELLOW}⚠{Colors.NC} Could not add redirect URIs: {update_result.stderr.strip()[:50]}")
        
        # Enable public client flows (required for PnP PowerShell -Interactive without client_secret)
        print(f"  {Colors.DIM}Enabling public client flows for interactive auth...{Colors.NC}")
        public_client_result = subprocess.run(
            [az_path, "ad", "app", "update", "--id", app_id, "--is-fallback-public-client", "true"],
            capture_output=True,
            text=True,
            timeout=30
        )
        if public_client_result.returncode == 0:
            print(f"    {Colors.GREEN}✓{Colors.NC} Enabled public client flows")
        else:
            print(f"    {Colors.YELLOW}⚠{Colors.NC} Could not enable public client flows: {public_client_result.stderr.strip()[:50]}")
        
        # Step 3: Create a service principal for the app
        print(f"  {Colors.DIM}Step 3/4: Creating service principal...{Colors.NC}")
        
        sp_result = subprocess.run(
            [az_path, "ad", "sp", "create",
             "--id", app_id,
             "-o", "json"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if sp_result.returncode == 0:
            print(f"  {Colors.GREEN}✓{Colors.NC} Service principal created")
        else:
            # Service principal might already exist
            if "already exists" in sp_result.stderr.lower():
                print(f"  {Colors.GREEN}✓{Colors.NC} Service principal already exists")
            else:
                print(f"  {Colors.YELLOW}⚠{Colors.NC} SP creation: {sp_result.stderr.strip()[:50]}")
        
        # Step 4: Create a client secret for the app
        print(f"  {Colors.DIM}Step 4/4: Creating client secret...{Colors.NC}")
        
        secret_result = subprocess.run(
            [az_path, "ad", "app", "credential", "reset",
             "--id", app_id,
             "--append",
             "--display-name", "SharePoint-Tool-Secret",
             "--years", "1",
             "-o", "json"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        client_secret = None
        if secret_result.returncode == 0:
            secret_data = json.loads(secret_result.stdout)
            client_secret = secret_data.get("password")
            print(f"  {Colors.GREEN}✓{Colors.NC} Client secret created")
        else:
            print(f"  {Colors.YELLOW}⚠{Colors.NC} Could not create secret: {secret_result.stderr.strip()[:50]}")
        
        # Save the app config (including secret)
        config = {
            "app_id": app_id,
            "object_id": object_id,
            "display_name": CUSTOM_APP_NAME,
            "tenant_id": get_tenant_id(),
            "client_secret": client_secret
        }
        save_app_config(config)
        
        print()
        print(f"  {Colors.GREEN}✓{Colors.NC} App registration complete!")
        
        return config
        
    except subprocess.TimeoutExpired:
        print_error("Command timed out")
        return None
    except Exception as e:
        print_error(f"Failed to create app: {e}")
        return None

def grant_consent_for_custom_app(app_id: str) -> bool:
    """Grant admin consent for the custom app."""
    az_path = find_azure_cli_path()
    if not az_path:
        return False
    
    print()
    print(f"  {Colors.WHITE}Granting admin consent for app...{Colors.NC}")
    
    try:
        result = subprocess.run(
            [az_path, "ad", "app", "permission", "admin-consent",
             "--id", app_id],
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode == 0:
            print(f"  {Colors.GREEN}✓{Colors.NC} Admin consent granted!")
            return True
        else:
            # Check for specific errors
            if "Insufficient privileges" in result.stderr or "Authorization_RequestDenied" in result.stderr:
                print(f"  {Colors.YELLOW}⚠{Colors.NC} Insufficient privileges to grant consent.")
                print(f"    {Colors.DIM}A Global Administrator must grant consent.{Colors.NC}")
            else:
                print(f"  {Colors.YELLOW}⚠{Colors.NC} Could not grant consent: {result.stderr.strip()[:100]}")
            return False
            
    except subprocess.TimeoutExpired:
        print_error("Command timed out")
        return False
    except Exception as e:
        print_error(f"Failed to grant consent: {e}")
        return False

def regenerate_client_secret() -> bool:
    """Regenerate the client secret for the custom app.
    
    This is needed for SharePoint REST API access which requires
    client credentials flow with a client secret.
    """
    az_path = find_azure_cli_path()
    if not az_path:
        print_error("Azure CLI not found")
        return False
    
    # Load existing app config
    app_config = load_app_config()
    if not app_config:
        # Try to find existing app
        existing_app = find_existing_app()
        if not existing_app:
            print_error("No app registration found. Please create one first.")
            return False
        app_id = existing_app.get("appId")
        app_config = {
            "app_id": app_id,
            "object_id": existing_app.get("id"),
            "display_name": existing_app.get("displayName", CUSTOM_APP_NAME),
            "tenant_id": get_tenant_id()
        }
    else:
        app_id = app_config.get("app_id")
    
    if not app_id:
        print_error("App ID not found")
        return False
    
    print()
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}Regenerating Client Secret{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print()
    print(f"  {Colors.WHITE}App ID:{Colors.NC} {app_id}")
    print()
    
    try:
        print(f"  {Colors.DIM}Creating new client secret...{Colors.NC}")
        
        secret_result = subprocess.run(
            [az_path, "ad", "app", "credential", "reset",
             "--id", app_id,
             "--append",
             "--display-name", f"SharePoint-Tool-Secret-{int(__import__('time').time())}",
             "--years", "1",
             "-o", "json"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if secret_result.returncode == 0:
            secret_data = json.loads(secret_result.stdout)
            client_secret = secret_data.get("password")
            
            if client_secret:
                # Update the app config with the new secret
                app_config["client_secret"] = client_secret
                save_app_config(app_config)
                
                print(f"  {Colors.GREEN}✓{Colors.NC} Client secret created successfully!")
                print()
                print(f"  {Colors.DIM}The secret has been saved to the app configuration.{Colors.NC}")
                print(f"  {Colors.DIM}SharePoint REST API access should now work.{Colors.NC}")
                return True
            else:
                print_error("Secret was created but password not returned")
                return False
        else:
            print_error(f"Failed to create secret: {secret_result.stderr.strip()[:100]}")
            return False
            
    except subprocess.TimeoutExpired:
        print_error("Command timed out")
        return False
    except Exception as e:
        print_error(f"Failed to create secret: {e}")
        return False

def open_consent_url_for_custom_app(app_id: str) -> bool:
    """Open the admin consent URL for the custom app in the browser."""
    import webbrowser
    
    tenant_id = get_tenant_id()
    if not tenant_id:
        print_error("Could not determine tenant ID")
        return False
    
    consent_url = (
        f"https://login.microsoftonline.com/{tenant_id}/adminconsent"
        f"?client_id={app_id}"
    )
    
    print()
    print(f"  {Colors.WHITE}Opening browser for admin consent...{Colors.NC}")
    print()
    print(f"  {Colors.YELLOW}In the browser:{Colors.NC}")
    print(f"    1. Sign in as a Global Administrator")
    print(f"    2. Review the permissions")
    print(f"    3. Click 'Accept' to grant consent")
    print()
    
    try:
        webbrowser.open(consent_url)
        print(f"  {Colors.GREEN}✓{Colors.NC} Browser opened!")
        return True
    except Exception as e:
        print_error(f"Could not open browser: {e}")
        print(f"  Please open this URL manually:")
        print(f"  {Colors.CYAN}{consent_url}{Colors.NC}")
        return False

def setup_custom_app_with_consent() -> bool:
    """Set up custom app registration and grant consent - fully automated."""
    print()
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}Automatic App Registration & Consent{Colors.NC}")
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print()
    print(f"  {Colors.DIM}This will create a custom app registration in your tenant{Colors.NC}")
    print(f"  {Colors.DIM}and grant it the necessary permissions for SharePoint access.{Colors.NC}")
    print()
    
    # Check if app already exists
    print(f"  {Colors.WHITE}Checking for existing app...{Colors.NC}")
    existing_app = find_existing_app()
    
    if existing_app:
        app_id = existing_app.get("appId")
        print(f"  {Colors.GREEN}✓{Colors.NC} Found existing app: {app_id}")
        
        # Save config if not already saved
        config = load_app_config()
        if not config or config.get("app_id") != app_id:
            config = {
                "app_id": app_id,
                "object_id": existing_app.get("id"),
                "display_name": CUSTOM_APP_NAME,
                "tenant_id": get_tenant_id()
            }
            save_app_config(config)
    else:
        # Create new app
        config = create_custom_app()
        if not config:
            return False
        app_id = config.get("app_id")
    
    # Try to grant admin consent via CLI first
    if grant_consent_for_custom_app(app_id):
        # Wait a moment for permissions to propagate
        import time
        print()
        print(f"  {Colors.DIM}Waiting for permissions to propagate...{Colors.NC}")
        time.sleep(5)
        
        # Verify permissions work
        result = check_graph_permissions()
        if result.get("has_permissions"):
            print()
            print(f"  {Colors.GREEN}{'═' * 60}{Colors.NC}")
            print(f"  {Colors.GREEN}✓ Setup Complete! SharePoint access is now enabled.{Colors.NC}")
            print(f"  {Colors.GREEN}{'═' * 60}{Colors.NC}")
            return True
        else:
            print(f"  {Colors.YELLOW}⚠{Colors.NC} Permissions may still be propagating.")
            print(f"    Please wait a minute and try again.")
    else:
        # Fall back to browser-based consent
        print()
        print(f"  {Colors.WHITE}Falling back to browser-based consent...{Colors.NC}")
        
        if open_consent_url_for_custom_app(app_id):
            print()
            print(f"  {Colors.YELLOW}Waiting for you to complete consent in the browser...{Colors.NC}")
            print(f"  {Colors.DIM}(Checking every 2 seconds for up to 2 minutes){Colors.NC}")
            print()
            
            import time
            for i in range(60):  # Wait up to 2 minutes
                time.sleep(2)
                print(f"  Checking permissions... ({(i+1)*2}/120 seconds)", end='\r')
                result = check_graph_permissions()
                if result.get("has_permissions"):
                    print()
                    print()
                    print(f"  {Colors.GREEN}{'═' * 60}{Colors.NC}")
                    print(f"  {Colors.GREEN}✓ Setup Complete! SharePoint access is now enabled.{Colors.NC}")
                    print(f"  {Colors.GREEN}{'═' * 60}{Colors.NC}")
                    return True
            
            print()
            print()
            print(f"  {Colors.YELLOW}⚠{Colors.NC} Consent not detected yet.")
            print(f"    It may take a few minutes for permissions to propagate.")
            print(f"    Please try running the prerequisites check again.")
    
    return False

def delete_custom_app() -> bool:
    """Delete the custom app registration (cleanup)."""
    az_path = find_azure_cli_path()
    if not az_path:
        return False
    
    config = load_app_config()
    if not config:
        existing_app = find_existing_app()
        if existing_app:
            app_id = existing_app.get("appId")
        else:
            print_info("No custom app found to delete.")
            return True
    else:
        app_id = config.get("app_id")
    
    print(f"  {Colors.WHITE}Deleting app registration: {app_id}...{Colors.NC}")
    
    try:
        result = subprocess.run(
            [az_path, "ad", "app", "delete", "--id", app_id],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode == 0:
            print(f"  {Colors.GREEN}✓{Colors.NC} App deleted successfully")
            # Remove config file
            if APP_CONFIG_FILE.exists():
                APP_CONFIG_FILE.unlink()
            return True
        else:
            print_error(f"Failed to delete app: {result.stderr}")
            return False
    except Exception as e:
        print_error(f"Failed to delete app: {e}")
        return False

def update_app_redirect_uris(app_id: str) -> bool:
    """Update redirect URIs on an existing app registration.
    
    Adds public client redirect URI for PnP PowerShell interactive authentication.
    PnP PowerShell requires the app to be configured as a public client (mobile/desktop app).
    """
    az_path = find_azure_cli_path()
    if not az_path:
        return False
    
    try:
        # Check if public client redirect URIs are already configured
        app_result = subprocess.run(
            [az_path, "ad", "app", "show", "--id", app_id, "--query", "publicClient.redirectUris", "-o", "json"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        current_public_uris = []
        if app_result.returncode == 0 and app_result.stdout.strip():
            try:
                current_public_uris = json.loads(app_result.stdout) or []
            except json.JSONDecodeError:
                current_public_uris = []
        
        # Required public client redirect URIs for PnP PowerShell
        # These are the standard URIs used by PnP PowerShell for interactive auth
        required_public_uris = [
            "http://localhost",
            "https://login.microsoftonline.com/common/oauth2/nativeclient",
            "urn:ietf:wg:oauth:2.0:oob"
        ]
        
        # Check if we need to add any URIs
        uris_to_add = [uri for uri in required_public_uris if uri not in current_public_uris]
        
        if not uris_to_add:
            print(f"  {Colors.GREEN}✓{Colors.NC} Public client redirect URIs already configured")
        else:
            # Add missing URIs
            new_uris = current_public_uris + uris_to_add
            
            # Update the app with public client redirect URIs
            # This configures the app as a "Mobile and desktop applications" platform
            update_result = subprocess.run(
                [az_path, "ad", "app", "update", "--id", app_id, "--public-client-redirect-uris"] + new_uris,
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if update_result.returncode == 0:
                print(f"  {Colors.GREEN}✓{Colors.NC} Added public client redirect URIs for PnP PowerShell")
            else:
                print(f"  {Colors.YELLOW}⚠{Colors.NC} Could not update redirect URIs: {update_result.stderr.strip()[:80]}")
                return False
        
        # Always check and enable public client flows (required for PnP PowerShell -Interactive without client_secret)
        # This allows the app to be used without a client_secret for interactive authentication
        # First check if it's already enabled
        check_result = subprocess.run(
            [az_path, "ad", "app", "show", "--id", app_id, "--query", "isFallbackPublicClient", "-o", "tsv"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        is_public_client = check_result.returncode == 0 and check_result.stdout.strip().lower() == "true"
        
        if is_public_client:
            print(f"  {Colors.GREEN}✓{Colors.NC} Public client flows already enabled")
        else:
            public_client_result = subprocess.run(
                [az_path, "ad", "app", "update", "--id", app_id, "--is-fallback-public-client", "true"],
                capture_output=True,
                text=True,
                timeout=30
            )
            if public_client_result.returncode == 0:
                print(f"  {Colors.GREEN}✓{Colors.NC} Enabled public client flows for interactive auth")
            else:
                print(f"  {Colors.YELLOW}⚠{Colors.NC} Could not enable public client flows: {public_client_result.stderr.strip()[:50]}")
        
        return True
            
    except Exception as e:
        print(f"  {Colors.YELLOW}⚠{Colors.NC} Could not update redirect URIs: {str(e)[:50]}")
        return False

def update_app_permissions() -> bool:
    """Update all permissions on an existing app registration.
    
    This adds both Microsoft Graph API permissions and Exchange Online (EWS) permissions.
    First checks which permissions already exist to avoid duplicates.
    """
    az_path = find_azure_cli_path()
    if not az_path:
        print_error("Azure CLI not found")
        return False
    
    # Get app ID from config or find existing app
    config = load_app_config()
    if not config:
        existing_app = find_existing_app()
        if existing_app:
            app_id = existing_app.get("appId")
        else:
            print_error("No existing app registration found. Use option [1] to create one.")
            return False
    else:
        app_id = config.get("app_id")
    
    if not app_id:
        print_error("Could not determine app ID")
        return False
    
    print()
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}Updating App Permissions{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print()
    print(f"  {Colors.WHITE}App ID:{Colors.NC} {app_id}")
    print()
    
    # First, get existing permissions to avoid duplicates
    print(f"  {Colors.DIM}Checking existing permissions...{Colors.NC}")
    existing_perm_ids = set()
    
    try:
        app_result = subprocess.run(
            [az_path, "ad", "app", "show", "--id", app_id, "--query", "requiredResourceAccess", "-o", "json"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if app_result.returncode == 0 and app_result.stdout.strip():
            resource_access = json.loads(app_result.stdout)
            for resource in resource_access:
                for access in resource.get("resourceAccess", []):
                    existing_perm_ids.add(access.get("id", ""))
    except Exception:
        pass  # If we can't get existing permissions, we'll try to add all
    
    # Also update redirect URIs for PnP PowerShell support
    print(f"  {Colors.DIM}Checking redirect URIs...{Colors.NC}")
    update_app_redirect_uris(app_id)
    
    print()
    
    added_count = 0
    skipped_count = 0
    
    # Add Microsoft Graph API permissions
    print(f"  {Colors.WHITE}Microsoft Graph API permissions:{Colors.NC}")
    
    for perm_name, perm_id in GRAPH_PERMISSION_IDS.items():
        # Check if permission already exists
        if perm_id in existing_perm_ids:
            print(f"    {Colors.DIM}○ Already exists: {perm_name}{Colors.NC}")
            skipped_count += 1
            continue
        
        try:
            add_perm_result = subprocess.run(
                [az_path, "ad", "app", "permission", "add",
                 "--id", app_id,
                 "--api", MICROSOFT_GRAPH_API_ID,
                 "--api-permissions", f"{perm_id}=Role"],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if add_perm_result.returncode == 0:
                print(f"    {Colors.GREEN}✓{Colors.NC} Added: {perm_name}")
                added_count += 1
            else:
                # Check if permission already exists (fallback check)
                if "already exists" in add_perm_result.stderr.lower() or "already been added" in add_perm_result.stderr.lower():
                    print(f"    {Colors.DIM}○ Already exists: {perm_name}{Colors.NC}")
                    skipped_count += 1
                else:
                    print(f"    {Colors.YELLOW}⚠{Colors.NC} {perm_name}: {add_perm_result.stderr.strip()[:50]}")
        except Exception as e:
            print(f"    {Colors.RED}✗{Colors.NC} {perm_name}: {str(e)[:50]}")
    
    # Add Exchange Online (EWS) permissions
    print()
    print(f"  {Colors.WHITE}Exchange Online (EWS) permissions:{Colors.NC}")
    
    for perm_name, perm_id in EWS_PERMISSION_IDS.items():
        # Check if permission already exists
        if perm_id in existing_perm_ids:
            print(f"    {Colors.DIM}○ Already exists: EWS.{perm_name}{Colors.NC}")
            skipped_count += 1
            continue
        
        try:
            add_perm_result = subprocess.run(
                [az_path, "ad", "app", "permission", "add",
                 "--id", app_id,
                 "--api", EXCHANGE_ONLINE_API_ID,
                 "--api-permissions", f"{perm_id}=Role"],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if add_perm_result.returncode == 0:
                print(f"    {Colors.GREEN}✓{Colors.NC} Added: EWS.{perm_name}")
                added_count += 1
            else:
                # Check if permission already exists (fallback check)
                if "already exists" in add_perm_result.stderr.lower() or "already been added" in add_perm_result.stderr.lower():
                    print(f"    {Colors.DIM}○ Already exists: EWS.{perm_name}{Colors.NC}")
                    skipped_count += 1
                else:
                    print(f"    {Colors.YELLOW}⚠{Colors.NC} EWS.{perm_name}: {add_perm_result.stderr.strip()[:50]}")
        except Exception as e:
            print(f"    {Colors.RED}✗{Colors.NC} EWS.{perm_name}: {str(e)[:50]}")
    
    # Add SharePoint Online API permissions (Delegated - for PnP PowerShell)
    print()
    print(f"  {Colors.WHITE}SharePoint Online API delegated permissions (for PnP PowerShell):{Colors.NC}")
    
    for perm_name, perm_id in SHAREPOINT_PERMISSION_IDS.items():
        # Check if permission already exists
        if perm_id in existing_perm_ids:
            print(f"    {Colors.DIM}○ Already exists: SharePoint.{perm_name} (delegated){Colors.NC}")
            skipped_count += 1
            continue
        
        try:
            add_perm_result = subprocess.run(
                [az_path, "ad", "app", "permission", "add",
                 "--id", app_id,
                 "--api", SHAREPOINT_ONLINE_API_ID,
                 "--api-permissions", f"{perm_id}=Scope"],  # Scope = Delegated permission
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if add_perm_result.returncode == 0:
                print(f"    {Colors.GREEN}✓{Colors.NC} Added: SharePoint.{perm_name} (delegated)")
                added_count += 1
            else:
                # Check if permission already exists (fallback check)
                if "already exists" in add_perm_result.stderr.lower() or "already been added" in add_perm_result.stderr.lower():
                    print(f"    {Colors.DIM}○ Already exists: SharePoint.{perm_name} (delegated){Colors.NC}")
                    skipped_count += 1
                else:
                    print(f"    {Colors.YELLOW}⚠{Colors.NC} SharePoint.{perm_name}: {add_perm_result.stderr.strip()[:50]}")
        except Exception as e:
            print(f"    {Colors.RED}✗{Colors.NC} SharePoint.{perm_name}: {str(e)[:50]}")
    
    # Add SharePoint Online API application permissions (for REST API app-only access)
    print()
    print(f"  {Colors.WHITE}SharePoint Online API application permissions (for REST API):{Colors.NC}")
    
    for perm_name, perm_id in SHAREPOINT_APP_PERMISSION_IDS.items():
        # Check if permission already exists
        if perm_id in existing_perm_ids:
            print(f"    {Colors.DIM}○ Already exists: SharePoint.{perm_name} (application){Colors.NC}")
            skipped_count += 1
            continue
        
        try:
            add_perm_result = subprocess.run(
                [az_path, "ad", "app", "permission", "add",
                 "--id", app_id,
                 "--api", SHAREPOINT_ONLINE_API_ID,
                 "--api-permissions", f"{perm_id}=Role"],  # Role = Application permission
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if add_perm_result.returncode == 0:
                print(f"    {Colors.GREEN}✓{Colors.NC} Added: SharePoint.{perm_name} (application)")
                added_count += 1
            else:
                # Check if permission already exists (fallback check)
                if "already exists" in add_perm_result.stderr.lower() or "already been added" in add_perm_result.stderr.lower():
                    print(f"    {Colors.DIM}○ Already exists: SharePoint.{perm_name} (application){Colors.NC}")
                    skipped_count += 1
                else:
                    print(f"    {Colors.YELLOW}⚠{Colors.NC} SharePoint.{perm_name}: {add_perm_result.stderr.strip()[:50]}")
        except Exception as e:
            print(f"    {Colors.RED}✗{Colors.NC} SharePoint.{perm_name}: {str(e)[:50]}")
    
    print()
    print(f"  {Colors.WHITE}Summary:{Colors.NC}")
    print(f"    • Added: {added_count} permissions")
    print(f"    • Already existed: {skipped_count} permissions")
    
    # Grant admin consent for the updated permissions
    print()
    print(f"  {Colors.WHITE}Granting admin consent for new permissions...{Colors.NC}")
    
    if grant_consent_for_custom_app(app_id):
        print(f"  {Colors.GREEN}✓{Colors.NC} Admin consent granted!")
        return True
    else:
        print(f"  {Colors.YELLOW}⚠{Colors.NC} Could not auto-grant consent.")
        print(f"  {Colors.DIM}Opening browser for manual consent...{Colors.NC}")
        open_consent_url_for_custom_app(app_id)
        return True

def grant_admin_consent_via_cli() -> bool:
    """Try to grant admin consent using Azure CLI (requires admin privileges)."""
    az_path = find_azure_cli_path()
    if not az_path:
        return False
    
    print()
    print(f"  {Colors.WHITE}Attempting to grant admin consent via Azure CLI...{Colors.NC}")
    print(f"  {Colors.DIM}(This requires you to be logged in as a Global Administrator){Colors.NC}")
    print()
    
    try:
        # Try to grant admin consent for the Azure CLI app
        result = subprocess.run(
            [az_path, "ad", "app", "permission", "admin-consent",
             "--id", AZURE_CLI_APP_ID],
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode == 0:
            print(f"  {Colors.GREEN}✓{Colors.NC} Admin consent granted successfully!")
            return True
        else:
            # Check if it's a permission error
            if "Insufficient privileges" in result.stderr or "Authorization_RequestDenied" in result.stderr:
                print(f"  {Colors.YELLOW}⚠{Colors.NC} You don't have admin privileges to grant consent.")
                print(f"    {Colors.DIM}A Global Administrator must grant consent.{Colors.NC}")
            else:
                print(f"  {Colors.YELLOW}⚠{Colors.NC} Could not grant consent via CLI.")
                if result.stderr:
                    print(f"    {Colors.DIM}{result.stderr.strip()[:100]}{Colors.NC}")
            return False
    except subprocess.TimeoutExpired:
        print(f"  {Colors.YELLOW}⚠{Colors.NC} Command timed out.")
        return False
    except Exception as e:
        print(f"  {Colors.YELLOW}⚠{Colors.NC} CLI consent failed: {e}")
        return False

def open_admin_consent_url() -> bool:
    """Open the admin consent URL in the default browser."""
    import webbrowser
    
    tenant_id = get_tenant_id()
    if not tenant_id:
        print_error("Could not determine tenant ID. Please log in first.")
        return False
    
    # Admin consent URL format
    consent_url = (
        f"https://login.microsoftonline.com/{tenant_id}/adminconsent"
        f"?client_id={AZURE_CLI_APP_ID}"
    )
    
    print()
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}Opening Admin Consent Page{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print()
    print(f"  {Colors.WHITE}Required Permissions:{Colors.NC}")
    for perm in REQUIRED_GRAPH_PERMISSIONS:
        print(f"    • {perm}")
    print()
    print(f"  {Colors.WHITE}Opening browser...{Colors.NC}")
    print()
    
    try:
        webbrowser.open(consent_url)
        print(f"  {Colors.GREEN}✓{Colors.NC} Browser opened!")
        print()
        print(f"  {Colors.YELLOW}In the browser:{Colors.NC}")
        print(f"    1. Sign in as a Global Administrator")
        print(f"    2. Review the permissions")
        print(f"    3. Click 'Accept' to grant consent")
        print()
        return True
    except Exception as e:
        print_error(f"Could not open browser: {e}")
        print()
        print(f"  Please open this URL manually:")
        print(f"  {Colors.CYAN}{consent_url}{Colors.NC}")
        return False

def auto_grant_admin_consent() -> bool:
    """Automatically grant admin consent using custom app registration.
    
    This creates a custom app registration in the tenant (if it doesn't exist)
    and grants it the necessary permissions. This is more reliable than trying
    to grant consent for the Azure CLI app, which is a first-party Microsoft app.
    """
    # Use the custom app registration flow - this is the most reliable method
    return setup_custom_app_with_consent()

def get_app_permissions(app_id: str) -> Dict[str, Any]:
    """Get the current permissions assigned to an app registration.
    
    Returns a dict with:
    - configured_permissions: list of permissions configured on the app
    - consented_permissions: list of permissions that have admin consent
    - error: any error message
    """
    az_path = find_azure_cli_path()
    if not az_path:
        return {"configured_permissions": [], "consented_permissions": [], "error": "Azure CLI not found"}
    
    result = {
        "configured_permissions": [],
        "consented_permissions": [],
        "error": None
    }
    
    # Reverse lookup for permission IDs to names
    permission_id_to_name = {v: k for k, v in GRAPH_PERMISSION_IDS.items()}
    ews_permission_id_to_name = {v: k for k, v in EWS_PERMISSION_IDS.items()}
    sharepoint_permission_id_to_name = {v: k for k, v in SHAREPOINT_PERMISSION_IDS.items()}
    # Add SharePoint application permissions to the lookup
    sharepoint_permission_id_to_name.update({v: k for k, v in SHAREPOINT_APP_PERMISSION_IDS.items()})
    
    try:
        # Get app registration details including required resource access
        app_result = subprocess.run(
            [az_path, "ad", "app", "show", "--id", app_id, "--query", "requiredResourceAccess", "-o", "json"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if app_result.returncode == 0 and app_result.stdout.strip():
            resource_access = json.loads(app_result.stdout)
            
            for resource in resource_access:
                resource_app_id = resource.get("resourceAppId", "")
                
                for access in resource.get("resourceAccess", []):
                    perm_id = access.get("id", "")
                    perm_type = access.get("type", "")  # "Role" for application permissions
                    
                    # Check if it's a Graph permission
                    if resource_app_id == MICROSOFT_GRAPH_API_ID:
                        perm_name = permission_id_to_name.get(perm_id, f"Unknown ({perm_id[:8]}...)")
                        result["configured_permissions"].append({
                            "name": perm_name,
                            "id": perm_id,
                            "type": perm_type,
                            "api": "Microsoft Graph"
                        })
                    # Check if it's an Exchange Online permission
                    elif resource_app_id == EXCHANGE_ONLINE_API_ID:
                        perm_name = ews_permission_id_to_name.get(perm_id, f"Unknown ({perm_id[:8]}...)")
                        result["configured_permissions"].append({
                            "name": perm_name,
                            "id": perm_id,
                            "type": perm_type,
                            "api": "Exchange Online"
                        })
                    # Check if it's a SharePoint Online permission
                    elif resource_app_id == SHAREPOINT_ONLINE_API_ID:
                        perm_name = sharepoint_permission_id_to_name.get(perm_id, f"Unknown ({perm_id[:8]}...)")
                        result["configured_permissions"].append({
                            "name": perm_name,
                            "id": perm_id,
                            "type": perm_type,
                            "api": "SharePoint Online"
                        })
        
        # Get service principal to check consented permissions
        sp_result = subprocess.run(
            [az_path, "ad", "sp", "list", "--filter", f"appId eq '{app_id}'", "--query", "[0].appRoleAssignments", "-o", "json"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if sp_result.returncode == 0 and sp_result.stdout.strip() and sp_result.stdout.strip() != "null":
            try:
                role_assignments = json.loads(sp_result.stdout)
                if role_assignments:
                    for assignment in role_assignments:
                        perm_id = assignment.get("appRoleId", "")
                        perm_name = permission_id_to_name.get(perm_id, ews_permission_id_to_name.get(perm_id, f"Unknown ({perm_id[:8]}...)"))
                        result["consented_permissions"].append({
                            "name": perm_name,
                            "id": perm_id
                        })
            except json.JSONDecodeError:
                pass
                
    except subprocess.TimeoutExpired:
        result["error"] = "Timeout getting app permissions"
    except Exception as e:
        result["error"] = str(e)
    
    return result


def run_graph_permissions_check() -> None:
    """Run the Graph API permissions check and offer to grant consent."""
    print()
    print(f"  {Colors.WHITE}{Colors.BOLD}Checking Microsoft Graph API Permissions...{Colors.NC}")
    print()
    
    # First check if logged in
    if not check_azure_login():
        print(f"  {Colors.YELLOW}⚠{Colors.NC} Not logged into Azure. Please log in first.")
        return
    
    # Check if we have a custom app configured
    app_config = load_app_config()
    if app_config:
        app_id = app_config.get('app_id', 'Unknown')
        print(f"  {Colors.GREEN}✓{Colors.NC} Custom app registered: {app_id}")
        print(f"    {Colors.DIM}Name: {app_config.get('display_name', 'Unknown')}{Colors.NC}")
        print()
        
        # Get and display current permissions
        print(f"  {Colors.WHITE}{Colors.BOLD}Configured Permissions:{Colors.NC}")
        print()
        
        perm_info = get_app_permissions(app_id)
        
        if perm_info["error"]:
            print(f"    {Colors.YELLOW}⚠{Colors.NC} Could not retrieve permissions: {perm_info['error']}")
        elif perm_info["configured_permissions"]:
            # Group by API and deduplicate by permission name
            graph_perms = [p for p in perm_info["configured_permissions"] if p["api"] == "Microsoft Graph"]
            ews_perms = [p for p in perm_info["configured_permissions"] if p["api"] == "Exchange Online"]
            sp_perms = [p for p in perm_info["configured_permissions"] if p["api"] == "SharePoint Online"]
            
            # Deduplicate permissions by name (keep unique names only)
            seen_graph = set()
            unique_graph_perms = []
            for p in graph_perms:
                if p["name"] not in seen_graph:
                    seen_graph.add(p["name"])
                    unique_graph_perms.append(p)
            
            seen_ews = set()
            unique_ews_perms = []
            for p in ews_perms:
                if p["name"] not in seen_ews:
                    seen_ews.add(p["name"])
                    unique_ews_perms.append(p)
            
            seen_sp = set()
            unique_sp_perms = []
            for p in sp_perms:
                if p["name"] not in seen_sp:
                    seen_sp.add(p["name"])
                    unique_sp_perms.append(p)
            
            if unique_graph_perms:
                print(f"    {Colors.CYAN}Microsoft Graph API:{Colors.NC}")
                for perm in unique_graph_perms:
                    # Check if this permission is in our required list
                    is_required = perm["name"] in REQUIRED_GRAPH_PERMISSIONS
                    status_icon = Colors.GREEN + "✓" + Colors.NC if is_required else Colors.DIM + "○" + Colors.NC
                    print(f"      {status_icon} {perm['name']}")
                print()
            
            if unique_ews_perms:
                print(f"    {Colors.CYAN}Exchange Online (EWS):{Colors.NC}")
                for perm in unique_ews_perms:
                    print(f"      {Colors.GREEN}✓{Colors.NC} {perm['name']}")
                print()
            
            if unique_sp_perms:
                print(f"    {Colors.CYAN}SharePoint Online API:{Colors.NC}")
                # Separate delegated (Scope) and application (Role) permissions
                delegated_perms = [p for p in unique_sp_perms if p.get("type") == "Scope"]
                app_perms = [p for p in unique_sp_perms if p.get("type") == "Role"]
                
                if delegated_perms:
                    print(f"      {Colors.DIM}Delegated (for PnP PowerShell):{Colors.NC}")
                    for perm in delegated_perms:
                        print(f"        {Colors.GREEN}✓{Colors.NC} {perm['name']}")
                
                if app_perms:
                    print(f"      {Colors.DIM}Application (for REST API):{Colors.NC}")
                    for perm in app_perms:
                        print(f"        {Colors.GREEN}✓{Colors.NC} {perm['name']}")
                print()
            
            # Check for missing required permissions
            configured_names = {p["name"] for p in perm_info["configured_permissions"]}
            missing_perms = [p for p in REQUIRED_GRAPH_PERMISSIONS if p not in configured_names]
            
            if missing_perms:
                print(f"    {Colors.YELLOW}Missing Required Permissions:{Colors.NC}")
                for perm in missing_perms:
                    print(f"      {Colors.RED}✗{Colors.NC} {perm}")
                print()
                print(f"    {Colors.DIM}Use option [4] 'Update Permissions' to add missing permissions{Colors.NC}")
        else:
            print(f"    {Colors.YELLOW}⚠{Colors.NC} No permissions configured on this app")
            print(f"    {Colors.DIM}Use option [4] 'Update Permissions' to add required permissions{Colors.NC}")
        
        print()
    else:
        print(f"  {Colors.YELLOW}⚠{Colors.NC} No custom app configured")
        print(f"    {Colors.DIM}Use option [1] 'Setup App Registration' to create one{Colors.NC}")
        print()
    
    # Test actual API access
    print(f"  {Colors.WHITE}{Colors.BOLD}Testing API Access:{Colors.NC}")
    print()
    
    result = check_graph_permissions()
    
    if result["has_permissions"]:
        print(f"    {Colors.GREEN}✓{Colors.NC} Graph API access: Working")
        print(f"    {Colors.GREEN}✓{Colors.NC} Can list SharePoint sites: Yes")
    else:
        print(f"    {Colors.RED}✗{Colors.NC} Graph API access: Not working")
        if result["error"]:
            print(f"      {Colors.DIM}Error: {result['error']}{Colors.NC}")
        print()
        print(f"  {Colors.WHITE}Options to fix this:{Colors.NC}")
        print()
        print(f"    {Colors.CYAN}1{Colors.NC}. Automatic Setup (Recommended)")
        print(f"       {Colors.DIM}Creates a custom app registration and grants permissions{Colors.NC}")
        print()
        print(f"    {Colors.CYAN}2{Colors.NC}. Manual Browser Consent")
        print(f"       {Colors.DIM}Opens browser for manual admin consent{Colors.NC}")
        print()
        print(f"    {Colors.CYAN}3{Colors.NC}. Skip for now")
        print()
        
        choice = input(f"  Enter choice (1-3): ").strip()
        
        if choice == '1':
            # Automatic setup with custom app
            if setup_custom_app_with_consent():
                print()
                print(f"  {Colors.GREEN}✓{Colors.NC} Setup complete! You can now use SharePoint features.")
        elif choice == '2':
            # Manual browser consent
            open_admin_consent_url()
            print()
            input(f"  {Colors.YELLOW}Press Enter after granting consent to re-check...{Colors.NC}")
            print()
            # Re-check permissions
            result = check_graph_permissions()
            if result["has_permissions"]:
                print(f"  {Colors.GREEN}✓{Colors.NC} Permissions granted successfully!")
            else:
                print(f"  {Colors.YELLOW}⚠{Colors.NC} Permissions still not available.")
                print(f"    It may take a few minutes for consent to propagate.")
                print(f"    Try running the populate_files.py script again later.")
        else:
            print(f"  {Colors.DIM}Skipped. You can set up permissions later.{Colors.NC}")

def check_pyyaml_installed() -> tuple:
    """Check if PyYAML is installed and return (installed, version)."""
    try:
        import yaml
        version = getattr(yaml, '__version__', 'unknown')
        return True, version
    except ImportError:
        return False, None


def install_pyyaml() -> bool:
    """Install PyYAML using pip."""
    try:
        print(f"  {Colors.YELLOW}Installing PyYAML...{Colors.NC}")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "pyyaml>=6.0"],
            capture_output=True,
            text=True
        )
        if result.returncode == 0:
            print(f"  {Colors.GREEN}✓{Colors.NC} PyYAML installed successfully")
            return True
        else:
            print(f"  {Colors.RED}✗{Colors.NC} Failed to install PyYAML: {result.stderr}")
            return False
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Failed to install PyYAML: {e}")
        return False


def check_exchangelib_installed() -> tuple:
    """Check if exchangelib is installed and return (installed, version)."""
    try:
        import exchangelib
        version = getattr(exchangelib, '__version__', 'unknown')
        return True, version
    except ImportError:
        return False, None


def install_exchangelib() -> bool:
    """Install exchangelib using pip."""
    try:
        print(f"  {Colors.YELLOW}Installing exchangelib (for EWS support)...{Colors.NC}")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "exchangelib>=5.0"],
            capture_output=True,
            text=True,
            timeout=180  # May take a while due to dependencies
        )
        if result.returncode == 0:
            print(f"  {Colors.GREEN}✓{Colors.NC} exchangelib installed successfully")
            return True
        else:
            print(f"  {Colors.RED}✗{Colors.NC} Failed to install exchangelib: {result.stderr}")
            return False
    except subprocess.TimeoutExpired:
        print(f"  {Colors.RED}✗{Colors.NC} Installation timed out. Try manually: pip install exchangelib")
        return False
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Failed to install exchangelib: {e}")
        return False


def get_pwsh_executable() -> Optional[str]:
    """Get PowerShell 7 (pwsh) executable path if available."""
    pwsh_paths = [
        "pwsh",  # In PATH
        r"C:\Program Files\PowerShell\7\pwsh.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps\pwsh.exe"),
    ]
    
    for path in pwsh_paths:
        try:
            result = subprocess.run(
                [path, "-Version"],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0:
                return path
        except Exception:
            continue
    
    return None


def get_powershell_executable() -> str:
    """Get the PowerShell executable path."""
    # Try pwsh (PowerShell 7) first
    pwsh = get_pwsh_executable()
    if pwsh:
        return pwsh
    
    # Fall back to Windows PowerShell
    if os.name == 'nt':
        return "powershell.exe"
    return "pwsh"


def check_pnp_module_installed() -> tuple:
    """Check if PnP PowerShell module is installed and return (installed, version)."""
    ps_exe = get_powershell_executable()
    
    try:
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-NoProfile", "-Command",
             "if (Get-Module -ListAvailable -Name 'PnP.PowerShell') { "
             "$m = Get-Module -ListAvailable -Name 'PnP.PowerShell' | Select-Object -First 1; "
             "Write-Output \"INSTALLED:$($m.Version)\" } else { Write-Output 'NOT_INSTALLED' }"],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        output = result.stdout.strip()
        if output.startswith("INSTALLED:"):
            version = output.replace("INSTALLED:", "").strip()
            return True, version
        return False, None
    except Exception:
        return False, None


def install_pnp_module() -> bool:
    """Install PnP PowerShell module."""
    # Prefer PowerShell 7 (pwsh) for better module management
    pwsh = get_pwsh_executable()
    ps_exe = pwsh if pwsh else get_powershell_executable()
    
    if pwsh:
        print(f"  {Colors.YELLOW}Installing PnP.PowerShell module using PowerShell 7...{Colors.NC}")
    else:
        print(f"  {Colors.YELLOW}Installing PnP.PowerShell module using Windows PowerShell...{Colors.NC}")
    print(f"  {Colors.DIM}This may take a few minutes...{Colors.NC}")
    
    # PowerShell script for installation - handles PowerShellGet issues
    ps_script = '''
$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"

# Ensure TLS 1.2 is used
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Check if already installed
$existing = Get-Module -ListAvailable -Name "PnP.PowerShell" | Select-Object -First 1
if ($existing) {
    Write-Output "SUCCESS:Already installed"
    exit 0
}

# First, ensure NuGet provider is installed
try {
    $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [Version]"2.8.5.201" }
    if (-not $nuget) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
    }
} catch { }

# Set PSGallery as trusted
try {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
} catch { }

# Method 1: Standard Install-Module
try {
    Install-Module -Name "PnP.PowerShell" -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -AcceptLicense -ErrorAction Stop
    Write-Output "SUCCESS:Installed"
    exit 0
} catch { }

# Method 2: Try with SkipPublisherCheck
try {
    Install-Module -Name "PnP.PowerShell" -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
    Write-Output "SUCCESS:Installed"
    exit 0
} catch { }

# Method 3: Try updating PowerShellGet first
try {
    Import-Module PackageManagement -Force -ErrorAction SilentlyContinue
    Install-Module -Name PowerShellGet -Force -AllowClobber -Scope CurrentUser -ErrorAction SilentlyContinue
    Install-Module -Name "PnP.PowerShell" -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    Write-Output "SUCCESS:Installed"
    exit 0
} catch { }

# Method 4: Direct download from PSGallery
try {
    $modulePath = Join-Path $env:USERPROFILE "Documents\\WindowsPowerShell\\Modules\\PnP.PowerShell"
    if (-not (Test-Path $modulePath)) {
        New-Item -ItemType Directory -Path $modulePath -Force | Out-Null
    }
    Save-Module -Name "PnP.PowerShell" -Path (Split-Path $modulePath -Parent) -Force -ErrorAction Stop
    Write-Output "SUCCESS:Installed"
    exit 0
} catch { }

# Method 5: PSResourceGet for PowerShell 7
if ($PSVersionTable.PSVersion.Major -ge 7) {
    try {
        Install-PSResource -Name "PnP.PowerShell" -Scope CurrentUser -TrustRepository -ErrorAction Stop
        Write-Output "SUCCESS:Installed"
        exit 0
    } catch { }
}

Write-Output "FAILED"
exit 1
'''
    
    try:
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=600  # 10 minutes for installation
        )
        
        if "SUCCESS:" in result.stdout:
            print(f"  {Colors.GREEN}✓{Colors.NC} PnP.PowerShell module installed successfully")
            return True
        else:
            print(f"  {Colors.RED}✗{Colors.NC} Automatic installation failed")
            print(f"  {Colors.DIM}  Please install manually:{Colors.NC}")
            print(f"  {Colors.DIM}  1. Open PowerShell as Administrator{Colors.NC}")
            print(f"  {Colors.DIM}  2. Run: Install-Module -Name PnP.PowerShell -Force{Colors.NC}")
            return False
    except subprocess.TimeoutExpired:
        print(f"  {Colors.RED}✗{Colors.NC} Installation timed out")
        return False
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Failed to install PnP.PowerShell: {e}")
        return False


def check_prerequisites(auto_install: bool = False) -> dict:
    """Check all prerequisites and optionally install missing ones."""
    results = {
        "python": {"installed": True, "version": f"Python {sys.version.split()[0]}"},
        "azure_cli": {"installed": False, "version": None},
        "terraform": {"installed": False, "version": None},
        "pyyaml": {"installed": False, "version": None},
        "exchangelib": {"installed": False, "version": None},
        "pnp_powershell": {"installed": False, "version": None},
        "azure_login": {"logged_in": False, "account_info": None},
        "graph_permissions": {"has_permissions": False, "error": None}
    }
    
    # Check Azure CLI
    if command_exists("az"):
        results["azure_cli"]["installed"] = True
        results["azure_cli"]["version"] = get_command_version("az")
        # Check Azure login and get account info
        results["azure_login"]["logged_in"] = check_azure_login()
        if results["azure_login"]["logged_in"]:
            results["azure_login"]["account_info"] = get_azure_account_info()
            # Check Graph API permissions (only if logged in)
            graph_result = check_graph_permissions()
            results["graph_permissions"] = graph_result
    elif auto_install:
        if install_azure_cli():
            results["azure_cli"]["installed"] = True
            results["azure_cli"]["version"] = get_command_version("az")
    
    # Check Terraform
    if command_exists("terraform"):
        results["terraform"]["installed"] = True
        results["terraform"]["version"] = get_command_version("terraform")
    elif auto_install:
        if install_terraform():
            results["terraform"]["installed"] = True
            results["terraform"]["version"] = get_command_version("terraform")
    
    # Check PyYAML (required for email population)
    pyyaml_installed, pyyaml_version = check_pyyaml_installed()
    if pyyaml_installed:
        results["pyyaml"]["installed"] = True
        results["pyyaml"]["version"] = pyyaml_version
    elif auto_install:
        if install_pyyaml():
            pyyaml_installed, pyyaml_version = check_pyyaml_installed()
            results["pyyaml"]["installed"] = pyyaml_installed
            results["pyyaml"]["version"] = pyyaml_version
    
    # Check exchangelib (optional - for EWS support with proper timestamps)
    exchangelib_installed, exchangelib_version = check_exchangelib_installed()
    if exchangelib_installed:
        results["exchangelib"]["installed"] = True
        results["exchangelib"]["version"] = exchangelib_version
    elif auto_install:
        if install_exchangelib():
            exchangelib_installed, exchangelib_version = check_exchangelib_installed()
            results["exchangelib"]["installed"] = exchangelib_installed
            results["exchangelib"]["version"] = exchangelib_version
    
    # Check PnP.PowerShell (required for site recycle bin operations - Windows only)
    if os.name == 'nt':  # Windows only
        pnp_installed, pnp_version = check_pnp_module_installed()
        if pnp_installed:
            results["pnp_powershell"]["installed"] = True
            results["pnp_powershell"]["version"] = pnp_version
        elif auto_install:
            if install_pnp_module():
                pnp_installed, pnp_version = check_pnp_module_installed()
                results["pnp_powershell"]["installed"] = pnp_installed
                results["pnp_powershell"]["version"] = pnp_version
    else:
        # On non-Windows, mark as N/A
        results["pnp_powershell"]["installed"] = None  # N/A
        results["pnp_powershell"]["version"] = "N/A (Windows only)"
    
    return results

def display_prerequisites_status(results: dict) -> None:
    """Display the prerequisites check results."""
    print()
    print(f"  {Colors.WHITE}{Colors.BOLD}Prerequisites Status:{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
    print()
    
    # Python
    print(f"  {Colors.GREEN}✓{Colors.NC} Python: {results['python']['version']}")
    
    # Azure CLI
    if results["azure_cli"]["installed"]:
        print(f"  {Colors.GREEN}✓{Colors.NC} Azure CLI: {results['azure_cli']['version']}")
    else:
        print(f"  {Colors.RED}✗{Colors.NC} Azure CLI: Not installed")
    
    # Terraform
    if results["terraform"]["installed"]:
        print(f"  {Colors.GREEN}✓{Colors.NC} Terraform: {results['terraform']['version']}")
    else:
        print(f"  {Colors.RED}✗{Colors.NC} Terraform: Not installed")
    
    # PyYAML (for email population)
    if results.get("pyyaml", {}).get("installed"):
        print(f"  {Colors.GREEN}✓{Colors.NC} PyYAML: {results['pyyaml']['version']} (for email population)")
    else:
        print(f"  {Colors.YELLOW}⚠{Colors.NC} PyYAML: Not installed (required for email population)")
    
    # exchangelib (for EWS support - proper timestamps and no draft prefix)
    if results.get("exchangelib", {}).get("installed"):
        print(f"  {Colors.GREEN}✓{Colors.NC} exchangelib: {results['exchangelib']['version']} (for EWS email timestamps)")
    else:
        print(f"  {Colors.YELLOW}⚠{Colors.NC} exchangelib: Not installed (optional - enables proper email timestamps)")
    
    # PnP.PowerShell (for site recycle bin operations - Windows only)
    pnp_result = results.get("pnp_powershell", {})
    if pnp_result.get("installed") is None:
        # N/A on non-Windows
        print(f"  {Colors.DIM}─{Colors.NC} PnP.PowerShell: N/A (Windows only)")
    elif pnp_result.get("installed"):
        print(f"  {Colors.GREEN}✓{Colors.NC} PnP.PowerShell: {pnp_result['version']} (for site recycle bin)")
    else:
        print(f"  {Colors.YELLOW}⚠{Colors.NC} PnP.PowerShell: Not installed (required for site recycle bin purge)")
    
    # Azure Login
    if results["azure_cli"]["installed"]:
        if results["azure_login"]["logged_in"]:
            print(f"  {Colors.GREEN}✓{Colors.NC} Azure Login: Authenticated")
            # Show account details if available
            account_info = results["azure_login"].get("account_info")
            if account_info:
                user_name = account_info.get('user', {}).get('name', 'Unknown')
                subscription_name = account_info.get('name', 'Unknown')
                subscription_id = account_info.get('id', 'Unknown')
                tenant_id = account_info.get('tenantId', 'Unknown')
                print(f"    {Colors.CYAN}├─ User:{Colors.NC} {user_name}")
                print(f"    {Colors.CYAN}├─ Subscription:{Colors.NC} {subscription_name}")
                print(f"    {Colors.CYAN}├─ Subscription ID:{Colors.NC} {subscription_id}")
                print(f"    {Colors.CYAN}└─ Tenant ID:{Colors.NC} {tenant_id}")
        else:
            print(f"  {Colors.YELLOW}⚠{Colors.NC} Azure Login: Not logged in")
    
    # Graph API Permissions (only show if logged in)
    if results["azure_login"]["logged_in"]:
        graph_perms = results.get("graph_permissions", {})
        if graph_perms.get("has_permissions"):
            print(f"  {Colors.GREEN}✓{Colors.NC} Graph API Permissions: OK (SharePoint access)")
        else:
            print(f"  {Colors.YELLOW}⚠{Colors.NC} Graph API Permissions: Admin consent required")
            if graph_perms.get("error"):
                print(f"    {Colors.DIM}└─ {graph_perms['error']}{Colors.NC}")
    
    print()
    
    # Summary
    all_installed = (
        results["azure_cli"]["installed"] and
        results["terraform"]["installed"]
    )
    pyyaml_ok = results.get("pyyaml", {}).get("installed", False)
    
    # PnP.PowerShell check (only on Windows, None means N/A)
    pnp_result = results.get("pnp_powershell", {})
    pnp_ok = pnp_result.get("installed") is True or pnp_result.get("installed") is None  # True or N/A
    
    graph_ok = results.get("graph_permissions", {}).get("has_permissions", False)
    
    if all_installed and pyyaml_ok and pnp_ok and results["azure_login"]["logged_in"] and graph_ok:
        print(f"  {Colors.GREEN}{Colors.BOLD}✓ All prerequisites met! Ready to proceed.{Colors.NC}")
    elif all_installed and results["azure_login"]["logged_in"] and not graph_ok:
        print(f"  {Colors.YELLOW}{Colors.BOLD}⚠ Graph API permissions missing.{Colors.NC}")
        print(f"  {Colors.DIM}  Setup will be offered below, or use [A] from main menu.{Colors.NC}")
    elif all_installed and not pyyaml_ok:
        print(f"  {Colors.YELLOW}{Colors.BOLD}⚠ PyYAML not installed (required for email population).{Colors.NC}")
        print(f"  {Colors.DIM}  Run 'pip install pyyaml' or select auto-install below.{Colors.NC}")
    elif all_installed and not pnp_ok:
        print(f"  {Colors.YELLOW}{Colors.BOLD}⚠ PnP.PowerShell not installed (required for site recycle bin purge).{Colors.NC}")
        print(f"  {Colors.DIM}  Installation will be offered below.{Colors.NC}")
    elif all_installed:
        print(f"  {Colors.YELLOW}{Colors.BOLD}⚠ Tools installed but not logged into Azure.{Colors.NC}")
        print(f"  {Colors.DIM}  Run 'az login' to authenticate.{Colors.NC}")
    else:
        print(f"  {Colors.RED}{Colors.BOLD}✗ Some prerequisites are missing.{Colors.NC}")
        print(f"  {Colors.DIM}  Select option [0] to install missing tools.{Colors.NC}")
    
    print()

def run_prerequisites_check_menu() -> None:
    """Run the prerequisites check and install menu."""
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}Step 0: Check & Install Prerequisites{Colors.NC}")
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    
    print()
    print_info("Checking prerequisites...")
    print()
    
    results = check_prerequisites(auto_install=False)
    display_prerequisites_status(results)
    
    all_installed = (
        results["azure_cli"]["installed"] and 
        results["terraform"]["installed"]
    )
    
    if not all_installed:
        print(f"  {Colors.WHITE}Would you like to install missing tools?{Colors.NC}")
        print()
        print(f"    {Colors.GREEN}[Y]{Colors.NC} Yes, install missing tools automatically")
        print(f"    {Colors.RED}[N]{Colors.NC} No, I'll install them manually")
        print()
        
        choice = input(f"  {Colors.YELLOW}Choice:{Colors.NC} ").strip().lower()
        
        if choice == 'y':
            print()
            results = check_prerequisites(auto_install=True)
            display_prerequisites_status(results)
    
    if not results["azure_login"]["logged_in"] and results["azure_cli"]["installed"]:
        print(f"  {Colors.WHITE}Would you like to log in to Azure now?{Colors.NC}")
        print()
        print(f"    {Colors.GREEN}[Y]{Colors.NC} Yes, open browser for Azure login")
        print(f"    {Colors.RED}[N]{Colors.NC} No, I'll log in later")
        print()
        
        choice = input(f"  {Colors.YELLOW}Choice:{Colors.NC} ").strip().lower()
        
        if choice == 'y':
            print()
            print_info("Opening browser for Azure login...")
            az_path = find_azure_cli_path()
            if not az_path:
                print_error("Azure CLI is not installed or not in PATH.")
                print_info("Please install Azure CLI: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli")
                print_info("After installation, restart your terminal and try again.")
            else:
                try:
                    subprocess.run([az_path, "login"], check=True)
                    print()
                    print_success("Azure login successful!")
                    # Re-check prerequisites after login
                    results = check_prerequisites(auto_install=False)
                except subprocess.CalledProcessError as e:
                    print_error(f"Azure login failed: {e}")
                except Exception as e:
                    print_error(f"Azure login failed: {e}")
    
    # Check PyYAML and offer to install if missing
    pyyaml_status = results.get("pyyaml", {})
    if not pyyaml_status.get("installed", False):
        print()
        print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
        print(f"  {Colors.WHITE}{Colors.BOLD}PyYAML Installation{Colors.NC}")
        print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
        print()
        print(f"  {Colors.YELLOW}⚠{Colors.NC} PyYAML is required for email population features.")
        print()
        print(f"  {Colors.WHITE}Would you like to install PyYAML now?{Colors.NC}")
        print()
        print(f"    {Colors.GREEN}[Y]{Colors.NC} Yes, install PyYAML automatically")
        print(f"    {Colors.RED}[N]{Colors.NC} No, I'll install it later")
        print()
        
        choice = input(f"  {Colors.YELLOW}Choice:{Colors.NC} ").strip().lower()
        
        if choice == 'y':
            print()
            if install_pyyaml():
                print_success("PyYAML installed successfully!")
                results["pyyaml"] = {"installed": True, "version": "installed"}
            else:
                print_error("Failed to install PyYAML. Try manually: pip install pyyaml")
    
    # Check exchangelib and offer to install if missing (optional but recommended)
    exchangelib_status = results.get("exchangelib", {})
    if not exchangelib_status.get("installed", False):
        print()
        print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
        print(f"  {Colors.WHITE}{Colors.BOLD}exchangelib Installation (Recommended){Colors.NC}")
        print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
        print()
        print(f"  {Colors.YELLOW}⚠{Colors.NC} exchangelib enables proper email timestamps and removes '[Draft]' prefix.")
        print(f"  {Colors.DIM}  Without it, emails will show current time and may appear as drafts.{Colors.NC}")
        print()
        print(f"  {Colors.WHITE}Would you like to install exchangelib now?{Colors.NC}")
        print()
        print(f"    {Colors.GREEN}[Y]{Colors.NC} Yes, install exchangelib (recommended)")
        print(f"    {Colors.RED}[N]{Colors.NC} No, use Graph API fallback (limited features)")
        print()
        
        choice = input(f"  {Colors.YELLOW}Choice:{Colors.NC} ").strip().lower()
        
        if choice == 'y':
            print()
            if install_exchangelib():
                print_success("exchangelib installed successfully!")
                results["exchangelib"] = {"installed": True, "version": "installed"}
            else:
                print_error("Failed to install exchangelib. Try manually: pip install exchangelib")
                print_info("The tool will use Graph API fallback (with limitations).")
    
    # Check PnP.PowerShell and offer to install if missing (Windows only)
    if os.name == 'nt':  # Windows only
        pnp_status = results.get("pnp_powershell", {})
        if pnp_status.get("installed") is False:  # Explicitly False, not None (N/A)
            print()
            print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
            print(f"  {Colors.WHITE}{Colors.BOLD}PnP.PowerShell Installation{Colors.NC}")
            print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
            print()
            print(f"  {Colors.YELLOW}⚠{Colors.NC} PnP.PowerShell is required for site recycle bin purge operations.")
            print(f"  {Colors.DIM}  Without it, you cannot permanently delete items from site recycle bins.{Colors.NC}")
            print()
            print(f"  {Colors.WHITE}Would you like to install PnP.PowerShell now?{Colors.NC}")
            print()
            print(f"    {Colors.GREEN}[Y]{Colors.NC} Yes, install PnP.PowerShell automatically")
            print(f"    {Colors.RED}[N]{Colors.NC} No, I'll install it later")
            print()
            
            choice = input(f"  {Colors.YELLOW}Choice:{Colors.NC} ").strip().lower()
            
            if choice == 'y':
                print()
                if install_pnp_module():
                    print_success("PnP.PowerShell installed successfully!")
                    results["pnp_powershell"] = {"installed": True, "version": "installed"}
                else:
                    print_error("Automatic installation failed.")
                    print()
                    print_info("Please install PnP.PowerShell manually using one of these methods:")
                    print_info("  Option 1 - PowerShell 7 (recommended):")
                    print_info("    winget install Microsoft.PowerShell")
                    print_info("    pwsh -Command \"Install-Module -Name PnP.PowerShell -Force\"")
                    print()
                    print_info("  Option 2 - Windows PowerShell as Admin:")
                    print_info("    Install-Module -Name PnP.PowerShell -Force")
                    print()
                    print_info("  Option 3 - Using winget:")
                    print_info("    winget install --id=PnP.PowerShell -e")
    
    # Check Graph API permissions and automatically grant consent if needed
    if results["azure_login"]["logged_in"]:
        graph_perms = results.get("graph_permissions", {})
        if not graph_perms.get("has_permissions"):
            print()
            print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
            print(f"  {Colors.WHITE}{Colors.BOLD}Graph API Permissions{Colors.NC}")
            print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
            print()
            print(f"  {Colors.YELLOW}⚠{Colors.NC} SharePoint access requires Microsoft Graph API permissions.")
            print()
            print(f"  {Colors.WHITE}Would you like to set up permissions automatically?{Colors.NC}")
            print(f"  {Colors.DIM}This will create a custom app registration and grant the required permissions.{Colors.NC}")
            print(f"  {Colors.DIM}(Requires Global Administrator or Privileged Role Administrator){Colors.NC}")
            print()
            print(f"    {Colors.GREEN}[Y]{Colors.NC} Yes, set up automatically (recommended)")
            print(f"    {Colors.RED}[N]{Colors.NC} No, I'll do this later")
            print()
            
            choice = input(f"  {Colors.YELLOW}Choice:{Colors.NC} ").strip().lower()
            
            if choice == 'y':
                # Use the automated consent function (now uses custom app registration)
                if auto_grant_admin_consent():
                    # Re-check and update results
                    results["graph_permissions"] = check_graph_permissions()
    
    print()
    input(f"  {Colors.YELLOW}Press Enter to return to menu...{Colors.NC}")

# ============================================================================
# MENU DISPLAY
# ============================================================================

def print_menu(prereq_status: dict) -> None:
    """Print the main menu options."""
    # Determine status indicators
    prereq_ok = (
        prereq_status["azure_cli"]["installed"] and 
        prereq_status["terraform"]["installed"] and
        prereq_status["azure_login"]["logged_in"]
    )
    prereq_icon = f"{Colors.GREEN}✓{Colors.NC}" if prereq_ok else f"{Colors.YELLOW}⚠{Colors.NC}"
    
    print(f"  {Colors.WHITE}{Colors.BOLD}What would you like to do?{Colors.NC}")
    print()
    print(f"  {Colors.CYAN}╭──────────────────────────────────────────────────────────────╮{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.WHITE}[0]{Colors.NC} {prereq_icon} {Colors.WHITE}Check & Install Prerequisites{Colors.NC}                    {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Azure CLI, Terraform, Azure Login{Colors.NC}                     {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}├──────────────────────────────────────────────────────────────┤{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.GREEN}[1]{Colors.NC} {Colors.WHITE}🏗️  Create SharePoint Sites{Colors.NC}                           {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Deploy new sites using Terraform{Colors.NC}                      {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.BLUE}[2]{Colors.NC} {Colors.WHITE}📄 Populate Sites with Files{Colors.NC}                          {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Add realistic documents to existing sites{Colors.NC}             {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.RED}[3]{Colors.NC} {Colors.WHITE}🗑️  Delete Files or Sites{Colors.NC}                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Clean up files or remove SharePoint sites{Colors.NC}             {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}├──────────────────────────────────────────────────────────────┤{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.YELLOW}[4]{Colors.NC} {Colors.WHITE}📋 List SharePoint Sites{Colors.NC}                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}View all available sites{Colors.NC}                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.MAGENTA}[5]{Colors.NC} {Colors.WHITE}📁 List Files in Sites{Colors.NC}                                {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}View files in SharePoint sites{Colors.NC}                        {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.CYAN}[6]{Colors.NC} {Colors.WHITE}📧 Populate Mailboxes with Emails{Colors.NC}                     {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Add realistic emails to M365 mailboxes{Colors.NC}                 {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.RED}[7]{Colors.NC} {Colors.WHITE}🗑️  Delete Emails from Mailboxes{Colors.NC}                      {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Clean up emails from M365 mailboxes{Colors.NC}                    {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.YELLOW}[8]{Colors.NC} {Colors.WHITE}📬 List Mailboxes{Colors.NC}                                     {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}View configured mailboxes and status{Colors.NC}                   {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.BLUE}[9]{Colors.NC} {Colors.WHITE}🔍 Azure AD User Discovery{Colors.NC}                            {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Discover users & groups from Azure AD{Colors.NC}                  {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}├──────────────────────────────────────────────────────────────┤{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.CYAN}[C]{Colors.NC} {Colors.WHITE}⚙️  Edit Configuration{Colors.NC}                                 {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Environments, sites, and settings{Colors.NC}                     {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.MAGENTA}[A]{Colors.NC} {Colors.WHITE}🔑 Manage App Registration{Colors.NC}                            {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Setup or remove custom app for permissions{Colors.NC}            {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.WHITE}[H]{Colors.NC} {Colors.WHITE}❓ Help & Documentation{Colors.NC}                               {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.WHITE}[Q]{Colors.NC} {Colors.WHITE}🚪 Quit{Colors.NC}                                                {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}╰──────────────────────────────────────────────────────────────╯{Colors.NC}")
    print()

def print_help() -> None:
    """Print help information."""
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}SharePoint Sites Management - Help{Colors.NC}")
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print()
    print(f"  {Colors.WHITE}{Colors.BOLD}Step 0: Check & Install Prerequisites{Colors.NC}")
    print(f"  {Colors.DIM}─────────────────────────────────────{Colors.NC}")
    print(f"  Verify and install required tools:")
    print(f"    • {Colors.CYAN}Azure CLI{Colors.NC} - For Azure authentication")
    print(f"    • {Colors.CYAN}Terraform{Colors.NC} - For infrastructure deployment")
    print(f"    • {Colors.CYAN}Azure Login{Colors.NC} - Authenticate to your tenant")
    print()
    print(f"  {Colors.GREEN}{Colors.BOLD}Step 1: Create SharePoint Sites{Colors.NC}")
    print(f"  {Colors.DIM}─────────────────────────────────{Colors.NC}")
    print(f"  Use this option to deploy new SharePoint sites using Terraform.")
    print(f"  You can either:")
    print(f"    • Define custom sites in {Colors.YELLOW}config/sites.json{Colors.NC}")
    print(f"    • Generate random sites with realistic department names")
    print()
    print(f"  {Colors.BLUE}{Colors.BOLD}Step 2: Populate Sites with Files{Colors.NC}")
    print(f"  {Colors.DIM}───────────────────────────────────{Colors.NC}")
    print(f"  Add realistic-looking documents to your SharePoint sites.")
    print(f"  Files are department-appropriate (HR docs, Finance reports, etc.)")
    print(f"  Supports Word, Excel, PowerPoint, and PDF formats.")
    print()
    print(f"  {Colors.RED}{Colors.BOLD}Step 3: Delete Files or Sites{Colors.NC}")
    print(f"  {Colors.DIM}───────────────────────────────{Colors.NC}")
    print(f"  Clean up your environment by deleting:")
    print(f"    • All files from sites (keeps sites)")
    print(f"    • Specific files (interactive selection)")
    print(f"    • Entire SharePoint sites")
    print(f"    • Purge M365 Groups recycle bin (Azure AD)")
    print(f"    • Purge SharePoint site recycle bin")
    print()
    print(f"  {Colors.CYAN}{Colors.BOLD}Option 6: Populate Mailboxes with Emails{Colors.NC}")
    print(f"  {Colors.DIM}─────────────────────────────────────────{Colors.NC}")
    print(f"  Add realistic emails to M365 mailboxes:")
    print(f"    • Newsletters, organizational communications, project updates")
    print(f"    • Emails with attachments (Word, Excel, PowerPoint, PDF)")
    print(f"    • Email threads (replies, forwards, reply-all)")
    print(f"    • Backdated over 6-12 months with business hours bias")
    print(f"    • Microsoft 365 sensitivity labels")
    print(f"  Configure mailboxes in {Colors.YELLOW}config/mailboxes.yaml{Colors.NC}")
    print()
    print(f"  {Colors.RED}{Colors.BOLD}Option 7: Delete Emails from Mailboxes{Colors.NC}")
    print(f"  {Colors.DIM}───────────────────────────────────────{Colors.NC}")
    print(f"  Clean up emails from M365 mailboxes:")
    print(f"    • Delete from specific folders (inbox, sent, drafts)")
    print(f"    • Delete from specific mailboxes or all mailboxes")
    print(f"    • Soft delete (to Deleted Items) or permanent delete")
    print(f"    • Empty Deleted Items folder")
    print()
    print(f"  {Colors.YELLOW}{Colors.BOLD}Option 8: List Mailboxes{Colors.NC}")
    print(f"  {Colors.DIM}─────────────────────────{Colors.NC}")
    print(f"  View configured mailboxes and their status:")
    print(f"    • Shows all mailboxes from {Colors.YELLOW}config/mailboxes.yaml{Colors.NC}")
    print(f"    • Optional validation against Azure AD")
    print(f"    • Shows email count per mailbox")
    print(f"    • Department summary")
    print()
    print(f"  {Colors.BLUE}{Colors.BOLD}Option 9: Azure AD User Discovery{Colors.NC}")
    print(f"  {Colors.DIM}───────────────────────────────────{Colors.NC}")
    print(f"  Discover users and groups from Azure AD:")
    print(f"    • Discover all users (with/without mailboxes)")
    print(f"    • Discover M365 groups, security groups, distribution lists")
    print(f"    • Validate which users have mailboxes")
    print(f"    • View users by department")
    print(f"    • Cache results for faster email population")
    print(f"  Configure in {Colors.YELLOW}config/mailboxes.yaml{Colors.NC} under 'azure_ad' section")
    print()
    print(f"  {Colors.YELLOW}{Colors.BOLD}Quick Commands:{Colors.NC}")
    print(f"  {Colors.DIM}────────────────{Colors.NC}")
    print(f"    {Colors.CYAN}python deploy.py --random 10{Colors.NC}     Create 10 random sites")
    print(f"    {Colors.CYAN}python populate_files.py --files 50{Colors.NC}  Add 50 files")
    print(f"    {Colors.CYAN}python populate_emails.py --all --emails 100{Colors.NC}")
    print(f"                                      Add 100 emails per mailbox")
    print(f"    {Colors.CYAN}python cleanup.py --list-sites{Colors.NC}    List all sites")
    print(f"    {Colors.CYAN}python cleanup.py --select-files{Colors.NC}  Delete specific files")
    print(f"    {Colors.CYAN}python cleanup.py --delete-sites{Colors.NC}  Delete SharePoint sites")
    print(f"    {Colors.CYAN}python cleanup.py --purge-deleted{Colors.NC} Purge M365 Groups recycle bin")
    print(f"    {Colors.CYAN}python cleanup.py --purge-spo-recycle --tenant NAME{Colors.NC}")
    print(f"                                      Purge SharePoint site recycle bin")
    print()
    print(f"  {Colors.WHITE}For more information, see:{Colors.NC}")
    print(f"    • {Colors.BLUE}README.md{Colors.NC} - Main documentation")
    print(f"    • {Colors.BLUE}CONFIGURATION-GUIDE.md{Colors.NC} - Configuration details")
    print(f"    • {Colors.BLUE}docs/TROUBLESHOOTING.md{Colors.NC} - Common issues")
    print()
    input(f"  {Colors.YELLOW}Press Enter to return to menu...{Colors.NC}")

# ============================================================================
# APP REGISTRATION MANAGEMENT
# ============================================================================

def manage_app_registration_menu() -> None:
    """Show the app registration management menu."""
    while True:
        clear_screen()
        print()
        print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
        print(f"  {Colors.WHITE}{Colors.BOLD}🔑 Manage App Registration{Colors.NC}")
        print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
        print()
        
        # Check current app status
        app_config = load_app_config()
        existing_app = find_existing_app() if not app_config else None
        
        if app_config:
            print(f"  {Colors.GREEN}✓{Colors.NC} Custom app is configured:")
            print(f"    {Colors.DIM}App ID: {app_config.get('app_id', 'Unknown')}{Colors.NC}")
            print(f"    {Colors.DIM}Name: {app_config.get('display_name', CUSTOM_APP_NAME)}{Colors.NC}")
            print(f"    {Colors.DIM}Tenant: {app_config.get('tenant_id', 'Unknown')}{Colors.NC}")
        elif existing_app:
            print(f"  {Colors.YELLOW}⚠{Colors.NC} Found existing app (not in local config):")
            print(f"    {Colors.DIM}App ID: {existing_app.get('appId', 'Unknown')}{Colors.NC}")
        else:
            print(f"  {Colors.DIM}No custom app registration found.{Colors.NC}")
        
        # Check permissions
        print()
        if check_azure_login():
            result = check_graph_permissions()
            if result.get("has_permissions"):
                print(f"  {Colors.GREEN}✓{Colors.NC} Graph API permissions: Working")
            else:
                print(f"  {Colors.RED}✗{Colors.NC} Graph API permissions: Not working")
                if result.get("error"):
                    print(f"    {Colors.DIM}{result['error']}{Colors.NC}")
        else:
            print(f"  {Colors.YELLOW}⚠{Colors.NC} Not logged into Azure")
        
        print()
        print(f"  {Colors.WHITE}Options:{Colors.NC}")
        print()
        print(f"  {Colors.CYAN}╭──────────────────────────────────────────────────────────────╮{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.GREEN}[1]{Colors.NC} {Colors.WHITE}🔧 Setup App Registration{Colors.NC}                             {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Create app and grant permissions automatically{Colors.NC}        {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.YELLOW}[2]{Colors.NC} {Colors.WHITE}🔍 Check Permissions{Colors.NC}                                  {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Test if current permissions are working{Colors.NC}               {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.RED}[3]{Colors.NC} {Colors.WHITE}🗑️  Delete App Registration{Colors.NC}                            {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Remove the custom app from Azure AD{Colors.NC}                   {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.BLUE}[4]{Colors.NC} {Colors.WHITE}🔄 Update Permissions{Colors.NC}                                  {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Add all required permissions (Graph + EWS + SharePoint){Colors.NC}  {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.MAGENTA}[5]{Colors.NC} {Colors.WHITE}✅ Grant Admin Consent{Colors.NC}                                 {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Grant admin consent for all permissions{Colors.NC}                {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.YELLOW}[6]{Colors.NC} {Colors.WHITE}🔑 Regenerate Client Secret{Colors.NC}                            {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Create new secret for SharePoint REST API access{Colors.NC}        {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}├──────────────────────────────────────────────────────────────┤{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.WHITE}[B]{Colors.NC} {Colors.WHITE}← Back to Main Menu{Colors.NC}                                  {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}╰──────────────────────────────────────────────────────────────╯{Colors.NC}")
        print()
        
        choice = input(f"  {Colors.YELLOW}Enter choice:{Colors.NC} ").strip().lower()
        
        if choice == '1':
            # Setup app registration
            if not check_azure_login():
                print()
                print_error("Please log into Azure first (option 0 from main menu)")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue
            
            setup_custom_app_with_consent()
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == '2':
            # Check permissions - use the enhanced function that shows all permissions
            run_graph_permissions_check()
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == '3':
            # Delete app registration
            print()
            if not check_azure_login():
                print_error("Please log into Azure first")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue
            
            # Confirm deletion
            app_config = load_app_config()
            existing_app = find_existing_app()
            
            if not app_config and not existing_app:
                print_info("No custom app registration found to delete.")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue
            
            print(f"  {Colors.RED}{Colors.BOLD}⚠️  WARNING: This will delete the app registration!{Colors.NC}")
            print()
            print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
            print(f"  {Colors.WHITE}App Details:{Colors.NC}")
            if app_config:
                print(f"    • Name:      {Colors.YELLOW}{app_config.get('display_name', CUSTOM_APP_NAME)}{Colors.NC}")
                print(f"    • App ID:    {app_config.get('app_id', 'Unknown')}")
                print(f"    • Tenant:    {app_config.get('tenant_id', 'Unknown')}")
            elif existing_app:
                print(f"    • Name:      {Colors.YELLOW}{existing_app.get('displayName', CUSTOM_APP_NAME)}{Colors.NC}")
                print(f"    • App ID:    {existing_app.get('appId', 'Unknown')}")
            print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
            print()
            print(f"  {Colors.YELLOW}This action cannot be undone.{Colors.NC}")
            print(f"  {Colors.DIM}You will need to set up the app again to use SharePoint/Email features.{Colors.NC}")
            print()
            print(f"  Type '{Colors.RED}DELETE{Colors.NC}' to confirm, or '{Colors.GREEN}C{Colors.NC}' to cancel:")
            print()
            
            confirm = input(f"  {Colors.YELLOW}Your choice:{Colors.NC} ").strip()
            
            if confirm == 'DELETE':
                if delete_custom_app():
                    print()
                    print(f"  {Colors.GREEN}✓{Colors.NC} App registration deleted successfully!")
                else:
                    print()
                    print(f"  {Colors.RED}✗{Colors.NC} Failed to delete app registration.")
            elif confirm.upper() == 'C':
                print()
                print(f"  {Colors.GREEN}✓{Colors.NC} Operation cancelled. No changes made.")
            else:
                print()
                print(f"  {Colors.DIM}Deletion cancelled. (Type 'DELETE' to confirm or 'C' to cancel){Colors.NC}")
            
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == '4':
            # Update permissions on existing app
            print()
            if not check_azure_login():
                print_error("Please log into Azure first")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue
            
            # Check if app exists
            app_config = load_app_config()
            existing_app = find_existing_app()
            
            if not app_config and not existing_app:
                print_error("No existing app registration found.")
                print_info("Use option [1] to create a new app registration first.")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue
            
            # Get app ID for permission check
            app_id = app_config.get('app_id') if app_config else existing_app.get('appId')
            
            print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
            print(f"  {Colors.WHITE}App to Update:{Colors.NC}")
            if app_config:
                print(f"    • Name:      {Colors.YELLOW}{app_config.get('display_name', CUSTOM_APP_NAME)}{Colors.NC}")
                print(f"    • App ID:    {app_config.get('app_id', 'Unknown')}")
                print(f"    • Tenant:    {app_config.get('tenant_id', 'Unknown')}")
            elif existing_app:
                print(f"    • Name:      {Colors.YELLOW}{existing_app.get('displayName', CUSTOM_APP_NAME)}{Colors.NC}")
                print(f"    • App ID:    {existing_app.get('appId', 'Unknown')}")
            print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
            print()
            
            # Get existing permissions
            print(f"  {Colors.DIM}Checking existing permissions...{Colors.NC}")
            existing_perms = get_app_permissions(app_id)
            existing_perm_names = {p.get("name") for p in existing_perms.get("configured_permissions", [])}
            
            # Count what needs to be added
            graph_to_add = []
            graph_existing = []
            ews_to_add = []
            ews_existing = []
            sp_to_add = []
            sp_existing = []
            
            # Check Graph permissions
            for perm_name in GRAPH_PERMISSION_IDS.keys():
                if perm_name in existing_perm_names:
                    graph_existing.append(perm_name)
                else:
                    graph_to_add.append(perm_name)
            
            # Check EWS permissions
            for perm_name in EWS_PERMISSION_IDS.keys():
                if perm_name in existing_perm_names:
                    ews_existing.append(perm_name)
                else:
                    ews_to_add.append(perm_name)
            
            # Check SharePoint delegated permissions
            for perm_name in SHAREPOINT_PERMISSION_IDS.keys():
                if perm_name in existing_perm_names:
                    sp_existing.append(perm_name)
                else:
                    sp_to_add.append(perm_name)
            
            # Check SharePoint application permissions
            sp_app_to_add = []
            sp_app_existing = []
            for perm_name in SHAREPOINT_APP_PERMISSION_IDS.keys():
                if perm_name in existing_perm_names:
                    sp_app_existing.append(perm_name)
                else:
                    sp_app_to_add.append(perm_name)
            
            total_to_add = len(graph_to_add) + len(ews_to_add) + len(sp_to_add) + len(sp_app_to_add)
            total_existing = len(graph_existing) + len(ews_existing) + len(sp_existing) + len(sp_app_existing)
            
            print()
            print(f"  {Colors.WHITE}Permission Status:{Colors.NC}")
            print(f"    • {Colors.GREEN}Already configured:{Colors.NC} {total_existing} permissions")
            print(f"    • {Colors.YELLOW}To be added:{Colors.NC} {total_to_add} permissions")
            print()
            
            # Show Microsoft Graph permissions
            print(f"  {Colors.CYAN}Microsoft Graph API:{Colors.NC}")
            perm_descriptions = {
                "Mail.ReadWrite": "Read and write mail in all mailboxes",
                "User.Read.All": "Read all users' profiles",
                "Sites.Read.All": "Read SharePoint sites",
                "Sites.ReadWrite.All": "Read and write SharePoint sites",
                "Sites.FullControl.All": "Full control of SharePoint sites",
                "Files.ReadWrite.All": "Read and write files",
                "Group.Read.All": "Read groups",
                "Group.ReadWrite.All": "Read and write groups",
            }
            for perm_name in GRAPH_PERMISSION_IDS.keys():
                desc = perm_descriptions.get(perm_name, "")
                if perm_name in existing_perm_names:
                    print(f"    {Colors.GREEN}✓{Colors.NC} {Colors.DIM}{perm_name}{Colors.NC} - {desc}")
                else:
                    print(f"    {Colors.YELLOW}○{Colors.NC} {Colors.GREEN}{perm_name}{Colors.NC} - {desc}")
            print()
            
            # Show Exchange Online permissions
            print(f"  {Colors.CYAN}Exchange Online (EWS):{Colors.NC}")
            for perm_name in EWS_PERMISSION_IDS.keys():
                if perm_name in existing_perm_names:
                    print(f"    {Colors.GREEN}✓{Colors.NC} {Colors.DIM}{perm_name}{Colors.NC} - Full mailbox access via EWS")
                else:
                    print(f"    {Colors.YELLOW}○{Colors.NC} {Colors.GREEN}{perm_name}{Colors.NC} - Full mailbox access via EWS")
            print()
            
            # Show SharePoint Online permissions
            print(f"  {Colors.CYAN}SharePoint Online API (Delegated - for PnP PowerShell):{Colors.NC}")
            sp_descriptions = {
                "AllSites.FullControl": "Full control of all site collections",
                "AllSites.Manage": "Create, edit and delete items and lists",
                "AllSites.Write": "Edit or delete items in all site collections",
                "AllSites.Read": "Read items in all site collections",
            }
            for perm_name in SHAREPOINT_PERMISSION_IDS.keys():
                desc = sp_descriptions.get(perm_name, "")
                if perm_name in existing_perm_names:
                    print(f"    {Colors.GREEN}✓{Colors.NC} {Colors.DIM}{perm_name}{Colors.NC} - {desc}")
                else:
                    print(f"    {Colors.YELLOW}○{Colors.NC} {Colors.GREEN}{perm_name}{Colors.NC} - {desc}")
            print()
            
            # Show SharePoint Online application permissions
            print(f"  {Colors.CYAN}SharePoint Online API (Application - for REST API):{Colors.NC}")
            sp_app_descriptions = {
                "Sites.FullControl.All": "Full control of all site collections (app-only)",
                "Sites.ReadWrite.All": "Read and write items in all site collections (app-only)",
                "Sites.Read.All": "Read items in all site collections (app-only)",
            }
            for perm_name in SHAREPOINT_APP_PERMISSION_IDS.keys():
                desc = sp_app_descriptions.get(perm_name, "")
                if perm_name in existing_perm_names:
                    print(f"    {Colors.GREEN}✓{Colors.NC} {Colors.DIM}{perm_name}{Colors.NC} - {desc}")
                else:
                    print(f"    {Colors.YELLOW}○{Colors.NC} {Colors.GREEN}{perm_name}{Colors.NC} - {desc}")
            print()
            
            print(f"  {Colors.DIM}Legend: {Colors.GREEN}✓{Colors.NC}{Colors.DIM} = Already configured, {Colors.YELLOW}○{Colors.NC}{Colors.DIM} = Will be added{Colors.NC}")
            print()
            
            if total_to_add == 0:
                print(f"  {Colors.GREEN}All permissions are already configured!{Colors.NC}")
                print()
                # Still check and update redirect URIs even if permissions are complete
                print(f"  {Colors.DIM}Checking redirect URIs for PnP PowerShell...{Colors.NC}")
                update_app_redirect_uris(app_id)
                print()
                print(f"  {Colors.DIM}No permission changes needed.{Colors.NC}")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue
            
            
            confirm = input(f"  {Colors.YELLOW}Proceed? (Y/N/C to cancel):{Colors.NC} ").strip().lower()
            
            if confirm == 'y':
                if update_app_permissions():
                    print()
                    print(f"  {Colors.GREEN}✓{Colors.NC} Permissions updated successfully!")
                else:
                    print()
                    print(f"  {Colors.YELLOW}⚠{Colors.NC} Some permissions may need manual consent.")
            elif confirm in ['n', 'c']:
                print()
                print(f"  {Colors.GREEN}✓{Colors.NC} Operation cancelled. No changes made.")
            else:
                print()
                print(f"  {Colors.DIM}Update cancelled. (Enter Y to proceed, N or C to cancel){Colors.NC}")
            
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == '5':
            # Grant admin consent
            print()
            if not check_azure_login():
                print_error("Please log into Azure first")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue
            
            # Check if app exists
            app_config = load_app_config()
            existing_app = find_existing_app()
            
            if not app_config and not existing_app:
                print_error("No existing app registration found.")
                print_info("Use option [1] to create a new app registration first.")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue
            
            # Get app ID
            app_id = app_config.get('app_id') if app_config else existing_app.get('appId')
            app_name = app_config.get('display_name', CUSTOM_APP_NAME) if app_config else existing_app.get('displayName', CUSTOM_APP_NAME)
            
            print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
            print(f"  {Colors.WHITE}Granting Admin Consent for:{Colors.NC}")
            print(f"    • Name:      {Colors.YELLOW}{app_name}{Colors.NC}")
            print(f"    • App ID:    {app_id}")
            print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
            print()
            
            print(f"  {Colors.WHITE}This will grant admin consent for all configured permissions.{Colors.NC}")
            print(f"  {Colors.DIM}This is required for application permissions to work.{Colors.NC}")
            print()
            
            # Try CLI-based consent first
            print(f"  {Colors.DIM}Attempting to grant consent via Azure CLI...{Colors.NC}")
            if grant_consent_for_custom_app(app_id):
                print()
                print(f"  {Colors.GREEN}✓{Colors.NC} Admin consent granted successfully!")
                print()
                print(f"  {Colors.DIM}All permissions should now be active.{Colors.NC}")
                print(f"  {Colors.DIM}You can verify by using option [2] Check Permissions.{Colors.NC}")
            else:
                print()
                print(f"  {Colors.YELLOW}⚠{Colors.NC} CLI-based consent failed. Opening browser for manual consent...")
                print()
                if open_consent_url_for_custom_app(app_id):
                    print()
                    print(f"  {Colors.GREEN}✓{Colors.NC} Browser opened for admin consent.")
                    print(f"  {Colors.DIM}Please complete the consent in your browser.{Colors.NC}")
                    print(f"  {Colors.DIM}After consenting, use option [2] to verify permissions.{Colors.NC}")
                else:
                    print()
                    print(f"  {Colors.RED}✗{Colors.NC} Failed to open browser.")
                    print()
                    # Show manual URL
                    tenant_id = app_config.get('tenant_id') if app_config else get_tenant_id()
                    if tenant_id:
                        consent_url = f"https://login.microsoftonline.com/{tenant_id}/adminconsent?client_id={app_id}"
                        print(f"  {Colors.WHITE}Please open this URL manually:{Colors.NC}")
                        print(f"  {Colors.CYAN}{consent_url}{Colors.NC}")
            
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == '6':
            # Regenerate client secret
            print()
            if not check_azure_login():
                print_error("Please log into Azure first")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue
            
            # Check if app exists
            app_config = load_app_config()
            existing_app = find_existing_app()
            
            if not app_config and not existing_app:
                print_error("No existing app registration found.")
                print_info("Use option [1] to create a new app registration first.")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue
            
            # Check if secret already exists
            if app_config and app_config.get("client_secret"):
                print()
                print(f"  {Colors.YELLOW}⚠{Colors.NC} A client secret already exists in the configuration.")
                print(f"  {Colors.DIM}Creating a new secret will not invalidate the old one.{Colors.NC}")
                print()
                confirm = input(f"  {Colors.YELLOW}Create a new secret anyway? (y/N):{Colors.NC} ").strip().lower()
                if confirm != 'y':
                    print()
                    print(f"  {Colors.GREEN}✓{Colors.NC} Operation cancelled. Existing secret retained.")
                    input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                    continue
            
            if regenerate_client_secret():
                print()
                print(f"  {Colors.GREEN}✓{Colors.NC} Client secret regenerated successfully!")
                print(f"  {Colors.DIM}SharePoint REST API access should now work.{Colors.NC}")
            else:
                print()
                print(f"  {Colors.RED}✗{Colors.NC} Failed to regenerate client secret.")
            
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == 'b':
            break
        else:
            print()
            print_warning("Invalid choice. Please try again.")
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


def list_sharepoint_sites_menu() -> None:
    """List SharePoint sites with a clean, friendly interface."""
    import urllib.request
    import urllib.error
    import urllib.parse
    
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print(f"  {Colors.CYAN}{'📋 SHAREPOINT SITES':^60}{Colors.NC}")
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print()
    
    # Get access token
    print(f"  {Colors.WHITE}Connecting to Microsoft Graph...{Colors.NC}")
    token = get_graph_access_token()
    
    if not token:
        print(f"  {Colors.RED}✗{Colors.NC} Failed to get access token")
        print(f"  {Colors.YELLOW}ℹ{Colors.NC} Please ensure you're logged in to Azure CLI")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    # Get sites from both APIs to ensure we capture all sites
    # 1. SharePoint Sites API - returns system sites and some user sites
    # 2. M365 Groups API - returns group-connected sites (created via Terraform)
    
    sites = []
    site_urls_seen = set()  # Track URLs to avoid duplicates
    
    # Try SharePoint Sites API first (for system sites)
    try:
        url = "https://graph.microsoft.com/v1.0/sites?search=*&$top=100"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            for site in data.get("value", []):
                web_url = site.get("webUrl", "")
                if web_url and web_url not in site_urls_seen:
                    site_urls_seen.add(web_url)
                    sites.append(site)
    except urllib.error.HTTPError as e:
        if e.code != 403:
            print(f"  {Colors.RED}✗{Colors.NC} Sites API Error: {e.code} - {e.reason}")
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Sites API Error: {str(e)}")
    
    # ALWAYS try M365 Groups API to get group-connected sites (created via Terraform)
    # These are the sites users typically want to manage
    try:
        # Try without filter first to see all groups, then filter client-side
        # The $filter with groupTypes can be problematic with some permissions
        url = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName,groupTypes,createdDateTime,visibility&$top=100"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            groups = data.get("value", [])
            
            # Filter to only M365 Groups (Unified) - these have SharePoint sites
            unified_groups = [g for g in groups if 'Unified' in g.get('groupTypes', [])]
            
            # Convert unified groups to sites format
            for group in unified_groups:
                try:
                    site_url = f"https://graph.microsoft.com/v1.0/groups/{group['id']}/sites/root"
                    site_req = urllib.request.Request(site_url)
                    site_req.add_header("Authorization", f"Bearer {token}")
                    
                    with urllib.request.urlopen(site_req, timeout=10) as site_response:
                        site_data = json.loads(site_response.read().decode())
                        web_url = site_data.get("webUrl", "")
                        
                        # Only add if not already seen
                        if web_url and web_url not in site_urls_seen:
                            site_urls_seen.add(web_url)
                            sites.append({
                                "id": site_data.get("id", ""),
                                "displayName": group.get("displayName", "Unknown"),
                                "webUrl": web_url,
                                "createdDateTime": group.get("createdDateTime", ""),
                                "visibility": group.get("visibility", ""),
                                "isGroup": True,
                                "groupId": group.get("id", "")  # Store group ID for deletion
                            })
                except Exception:
                    # Group might not have a SharePoint site
                    pass
    except urllib.error.HTTPError as e:
        if e.code == 403:
            print(f"  {Colors.YELLOW}ℹ{Colors.NC} No access to Groups API (need Group.Read.All)")
        else:
            print(f"  {Colors.RED}✗{Colors.NC} Groups API Error: {e.code} - {e.reason}")
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Groups API Error: {str(e)}")
    
    # Also check for DELETED M365 Groups - these may still have SharePoint sites visible
    # This explains why sites "reappear" after deletion - the groups are soft-deleted
    deleted_groups_count = 0
    try:
        deleted_url = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group?$select=id,displayName,groupTypes,deletedDateTime&$top=100"
        del_req = urllib.request.Request(deleted_url)
        del_req.add_header("Authorization", f"Bearer {token}")
        del_req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(del_req, timeout=30) as del_response:
            del_data = json.loads(del_response.read().decode())
            deleted_groups = del_data.get("value", [])
            
            # Filter to M365 Groups only
            deleted_unified = [g for g in deleted_groups if 'Unified' in g.get('groupTypes', [])]
            deleted_groups_count = len(deleted_unified)
    except Exception:
        # Silently ignore errors checking deleted groups
        pass
    
    # Categorize sites into deletable and system sites
    deletable_sites, system_sites = categorize_sites(sites)
    
    if not deletable_sites and not system_sites:
        print()
        print(f"  {Colors.YELLOW}No SharePoint sites found.{Colors.NC}")
        print()
        print(f"  {Colors.WHITE}Possible reasons:{Colors.NC}")
        print(f"    • No sites have been created yet")
        print(f"    • Missing permissions (Sites.Read.All or Group.Read.All)")
        print(f"    • Use option [A] to set up app registration with proper permissions")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    # Display summary
    print()
    total_sites = len(deletable_sites) + len(system_sites)
    print(f"  {Colors.GREEN}✓{Colors.NC} Found {total_sites} SharePoint sites")
    if deleted_groups_count > 0:
        print(f"  {Colors.YELLOW}⚠{Colors.NC} {deleted_groups_count} deleted groups in recycle bin (use cleanup to purge)")
    print()
    
    # Display deletable sites first
    if deletable_sites:
        print(f"  {Colors.CYAN}{'─' * 70}{Colors.NC}")
        print(f"  {Colors.GREEN}✓ Deletable Sites ({len(deletable_sites)}):{Colors.NC}")
        print(f"  {Colors.DIM}  These sites can be deleted via cleanup{Colors.NC}")
        print(f"  {Colors.CYAN}{'─' * 70}{Colors.NC}")
        print()
        
        for i, site in enumerate(deletable_sites, 1):
            name = site.get("displayName", site.get("name", "Unknown"))
            web_url = site.get("webUrl", "")
            visibility = site.get("visibility", "")
            is_group = site.get("isGroup", False)
            
            # Icon based on visibility
            if visibility == "Private":
                icon = "🔒"
            else:
                icon = "🌐"
            
            # Show if it's a group-connected site
            group_tag = f" {Colors.CYAN}(M365 Group){Colors.NC}" if is_group else ""
            
            print(f"  [{i:2}] {icon} {name}{group_tag}")
            print(f"       {Colors.DIM}{web_url}{Colors.NC}")
            print()
    else:
        print()
        print(f"  {Colors.YELLOW}No deletable sites found.{Colors.NC}")
        print()
    
    # Display system sites separately
    if system_sites:
        print(f"  {Colors.YELLOW}{'─' * 70}{Colors.NC}")
        print(f"  {Colors.RED}🔒 Protected System Sites ({len(system_sites)}):{Colors.NC}")
        print(f"  {Colors.DIM}  These are built-in SharePoint sites that cannot be deleted{Colors.NC}")
        print(f"  {Colors.YELLOW}{'─' * 70}{Colors.NC}")
        print()
        
        for site in system_sites:
            name = site.get("displayName", site.get("name", "Unknown"))
            web_url = site.get("webUrl", "")
            print(f"      {Colors.DIM}• {name}{Colors.NC}")
            print(f"        {Colors.DIM}{web_url}{Colors.NC}")
            print()
    
    print(f"  {Colors.WHITE}{'─' * 70}{Colors.NC}")
    print()
    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


def get_folder_contents_recursive(site_id: str, folder_id: str, token: str, path: str = "") -> list:
    """Recursively get all files from a folder."""
    import urllib.request
    import urllib.error
    
    all_items = []
    try:
        if folder_id:
            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{folder_id}/children"
        else:
            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
        
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            items = data.get("value", [])
            
            for item in items:
                item_path = f"{path}/{item.get('name', '')}" if path else item.get('name', '')
                item['_path'] = item_path
                
                if "folder" in item:
                    # Recursively get contents of this folder
                    sub_items = get_folder_contents_recursive(site_id, item.get('id', ''), token, item_path)
                    all_items.extend(sub_items)
                else:
                    all_items.append(item)
    except Exception:
        pass
    
    return all_items


def list_files_in_sites_menu() -> None:
    """List files in SharePoint sites with a clean, friendly interface."""
    import urllib.request
    import urllib.error
    import urllib.parse
    
    # Get access token once at the start
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print(f"  {Colors.CYAN}{'📁 FILES IN SHAREPOINT SITES':^60}{Colors.NC}")
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print()
    
    print(f"  {Colors.WHITE}Connecting to Microsoft Graph...{Colors.NC}")
    token = get_graph_access_token()
    
    if not token:
        print(f"  {Colors.RED}✗{Colors.NC} Failed to get access token")
        print(f"  {Colors.YELLOW}ℹ{Colors.NC} Please ensure you're logged in to Azure CLI")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    # Get sites from both APIs to ensure we capture all sites
    sites = []
    site_urls_seen = set()
    
    # Try SharePoint Sites API first
    try:
        url = "https://graph.microsoft.com/v1.0/sites?search=*&$top=100"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            for site in data.get("value", []):
                web_url = site.get("webUrl", "")
                if web_url and web_url not in site_urls_seen:
                    site_urls_seen.add(web_url)
                    sites.append(site)
    except urllib.error.HTTPError:
        pass
    except Exception:
        pass
    
    # ALWAYS try M365 Groups API to get group-connected sites
    print(f"  {Colors.YELLOW}ℹ{Colors.NC} Checking Microsoft 365 Groups...")
    try:
        filter_param = urllib.parse.quote("groupTypes/any(c:c eq 'Unified')")
        url = f"https://graph.microsoft.com/v1.0/groups?$filter={filter_param}&$select=id,displayName&$top=100"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            groups = data.get("value", [])
            
            for group in groups:
                try:
                    site_url = f"https://graph.microsoft.com/v1.0/groups/{group['id']}/sites/root"
                    site_req = urllib.request.Request(site_url)
                    site_req.add_header("Authorization", f"Bearer {token}")
                    
                    with urllib.request.urlopen(site_req, timeout=10) as site_response:
                        site_data = json.loads(site_response.read().decode())
                        web_url = site_data.get("webUrl", "")
                        
                        if web_url and web_url not in site_urls_seen:
                            site_urls_seen.add(web_url)
                            sites.append({
                                "id": site_data.get("id", ""),
                                "displayName": group.get("displayName", "Unknown"),
                                "webUrl": web_url,
                                "isGroup": True
                            })
                except Exception:
                    pass
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Error: {str(e)}")
    
    # Categorize sites into deletable and system sites
    deletable_sites, system_sites = categorize_sites(sites)
    
    # For file listing, we use deletable sites (user-created sites)
    sites = deletable_sites
    
    if not sites and not system_sites:
        print()
        print(f"  {Colors.YELLOW}No SharePoint sites found.{Colors.NC}")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    if not sites:
        print()
        print(f"  {Colors.YELLOW}No user-created sites found.{Colors.NC}")
        print(f"  {Colors.DIM}  Only system sites exist (which typically don't have user files){Colors.NC}")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    # Helper function to display items based on view mode
    def display_items(items: list, view_mode: int, indent: str = "      ") -> int:
        """Display items based on view mode. Returns count of displayed items."""
        count = 0
        for item in items:
            name = item.get("name", "Unknown")
            is_folder = "folder" in item
            size = item.get("size", 0)
            path = item.get("_path", "")
            
            # Filter based on view mode
            if view_mode == 1 and not is_folder:  # Folders only
                continue
            if view_mode == 2 and is_folder:  # Files only
                continue
            # view_mode 3 and 4 show everything
            
            if is_folder:
                icon = "📁"
                size_str = ""
            else:
                icon = "📄"
                if size < 1024:
                    size_str = f"({size} B)"
                elif size < 1024 * 1024:
                    size_str = f"({size // 1024} KB)"
                else:
                    size_str = f"({size // (1024 * 1024)} MB)"
            
            # Show path for recursive mode
            if view_mode == 4 and path:
                print(f"{indent}{icon} {path} {Colors.DIM}{size_str}{Colors.NC}")
            else:
                print(f"{indent}{icon} {name} {Colors.DIM}{size_str}{Colors.NC}")
            count += 1
        return count
    
    # View mode labels
    mode_labels = {
        1: "FOLDERS",
        2: "FILES",
        3: "ALL ITEMS",
        4: "ALL FILES (RECURSIVE)"
    }
    
    # Main menu loop - stays in this menu until user chooses to go back
    while True:
        clear_screen()
        print()
        print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
        print(f"  {Colors.CYAN}{'📁 FILES IN SHAREPOINT SITES':^60}{Colors.NC}")
        print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
        print()
        print(f"  {Colors.GREEN}✓{Colors.NC} {len(sites)} SharePoint sites available")
        print()
        
        # Ask what to show
        print(f"  {Colors.WHITE}What would you like to view?{Colors.NC}")
        print()
        print(f"    [1] 📁 Folders only (top-level)")
        print(f"    [2] 📄 Files only (top-level)")
        print(f"    [3] 📋 All items (folders + files, top-level)")
        print(f"    [4] 📄 All files recursively (includes files in subfolders)")
        print()
        print(f"    [{Colors.RED}B{Colors.NC}]  Back to main menu")
        print()
        
        view_choice = input(f"  {Colors.YELLOW}Enter choice (1-4):{Colors.NC} ").strip().lower()
        
        if view_choice == 'b' or not view_choice:
            return
        
        try:
            view_mode = int(view_choice)
            if view_mode < 1 or view_mode > 4:
                view_mode = 3  # Default to all items
        except ValueError:
            view_mode = 3  # Default to all items
        
        # Let user select a site or all sites
        clear_screen()
        print()
        print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
        print(f"  {Colors.CYAN}{'📁 SELECT SITE':^60}{Colors.NC}")
        print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
        print()
        print(f"  {Colors.WHITE}Select a site to view:{Colors.NC}")
        print()
        print(f"    [{Colors.GREEN}*{Colors.NC}]  {Colors.GREEN}All sites{Colors.NC}")
        print()
        for i, site in enumerate(sites, 1):
            name = site.get("displayName", site.get("name", "Unknown"))
            print(f"    [{i:2}] {name}")
        print()
        print(f"    [{Colors.RED}B{Colors.NC}]  Back to view options")
        print()
        
        choice = input(f"  {Colors.YELLOW}Enter site number or * for all:{Colors.NC} ").strip().lower()
        
        if choice == 'b' or not choice:
            continue  # Go back to view options menu
        
        # Handle "all sites" option
        if choice == '*':
            clear_screen()
            print()
            print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
            print(f"  {Colors.CYAN}{f'📁 {mode_labels[view_mode]} IN ALL SITES':^60}{Colors.NC}")
            print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
            print()
            
            total_items = 0
            for site in sites:
                site_id = site.get("id", "")
                site_name = site.get("displayName", "Unknown")
                
                if not site_id:
                    continue
                
                try:
                    if view_mode == 4:
                        # Recursive mode - get all files from all folders
                        items = get_folder_contents_recursive(site_id, "", token)
                    else:
                        # Top-level only
                        drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
                        req = urllib.request.Request(drive_url)
                        req.add_header("Authorization", f"Bearer {token}")
                        req.add_header("Content-Type", "application/json")
                        
                        with urllib.request.urlopen(req, timeout=30) as response:
                            data = json.loads(response.read().decode())
                            items = data.get("value", [])
                    
                    # Filter items based on view mode for counting
                    if view_mode == 1:
                        filtered_items = [i for i in items if "folder" in i]
                    elif view_mode == 2 or view_mode == 4:
                        filtered_items = [i for i in items if "folder" not in i]
                    else:
                        filtered_items = items
                    
                    if filtered_items:
                        print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
                        print(f"  {Colors.WHITE}{Colors.BOLD}📁 {site_name}{Colors.NC} ({len(filtered_items)} items)")
                        print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
                        
                        count = display_items(items, view_mode)
                        total_items += count
                        print()
                except Exception:
                    pass
            
            print(f"  {Colors.WHITE}{'═' * 60}{Colors.NC}")
            print(f"  {Colors.GREEN}Total: {total_items} items across {len(sites)} sites{Colors.NC}")
            print()
            input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            continue  # Return to view options menu
        
        # Handle specific site selection
        try:
            site_idx = int(choice) - 1
            if site_idx < 0 or site_idx >= len(sites):
                print(f"  {Colors.RED}✗{Colors.NC} Invalid selection")
                input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                continue  # Return to view options menu
        except ValueError:
            print(f"  {Colors.RED}✗{Colors.NC} Invalid input")
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            continue  # Return to view options menu
        
        selected_site = sites[site_idx]
        site_id = selected_site.get("id", "")
        site_name = selected_site.get("displayName", "Unknown")
        
        if not site_id:
            print(f"  {Colors.RED}✗{Colors.NC} Could not get site ID")
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            continue  # Return to view options menu
        
        # Get files from the site's document library
        print()
        print(f"  {Colors.WHITE}Loading from '{site_name}'...{Colors.NC}")
        
        items = []
        try:
            if view_mode == 4:
                # Recursive mode - get all files from all folders
                items = get_folder_contents_recursive(site_id, "", token)
            else:
                # Top-level only
                drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
                req = urllib.request.Request(drive_url)
                req.add_header("Authorization", f"Bearer {token}")
                req.add_header("Content-Type", "application/json")
                
                with urllib.request.urlopen(req, timeout=30) as response:
                    data = json.loads(response.read().decode())
                    items = data.get("value", [])
        except urllib.error.HTTPError as e:
            print(f"  {Colors.RED}✗{Colors.NC} Error accessing files: {e.code} - {e.reason}")
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            continue  # Return to view options menu
        except Exception as e:
            print(f"  {Colors.RED}✗{Colors.NC} Error: {str(e)}")
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            continue  # Return to view options menu
        
        # Display items
        clear_screen()
        print()
        print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
        title = f"📁 {mode_labels[view_mode]} IN: {site_name[:35]}"
        print(f"  {Colors.CYAN}{title:^60}{Colors.NC}")
        print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
        print()
        
        # Filter items based on view mode for counting
        if view_mode == 1:
            filtered_items = [i for i in items if "folder" in i]
        elif view_mode == 2 or view_mode == 4:
            filtered_items = [i for i in items if "folder" not in i]
        else:
            filtered_items = items
        
        if not filtered_items:
            print(f"  {Colors.YELLOW}No items found matching the selected filter.{Colors.NC}")
        else:
            print(f"  {Colors.GREEN}✓{Colors.NC} Found {len(filtered_items)} items")
            print()
            print(f"  {Colors.WHITE}{'─' * 70}{Colors.NC}")
            print()
            
            display_items(items, view_mode, indent="  ")
            
            print()
            print(f"  {Colors.WHITE}{'─' * 70}{Colors.NC}")
        
        print()
        input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        # Loop continues back to view options menu


# ============================================================================
# MAILBOX LISTING
# ============================================================================

def list_mailboxes_menu() -> None:
    """List mailboxes from configuration with validation status."""
    import urllib.request
    import urllib.error
    
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print(f"  {Colors.CYAN}{'📬 CONFIGURED MAILBOXES':^60}{Colors.NC}")
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print()
    
    # Check if PyYAML is installed
    yaml_installed, _ = check_pyyaml_installed()
    if not yaml_installed:
        print(f"  {Colors.RED}✗{Colors.NC} PyYAML is not installed")
        print(f"  {Colors.YELLOW}ℹ{Colors.NC} Run option [0] to install prerequisites")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    import yaml
    
    # Load mailboxes configuration
    mailboxes_file = SCRIPT_DIR.parent / "config" / "mailboxes.yaml"
    if not mailboxes_file.exists():
        print(f"  {Colors.RED}✗{Colors.NC} Mailboxes configuration not found")
        print(f"  {Colors.DIM}Expected: {mailboxes_file}{Colors.NC}")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    try:
        with open(mailboxes_file, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Failed to load mailboxes.yaml: {e}")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    users = config.get("users", [])
    if not users:
        print(f"  {Colors.YELLOW}ℹ{Colors.NC} No mailboxes configured in mailboxes.yaml")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    print(f"  {Colors.WHITE}Found {len(users)} mailboxes in configuration{Colors.NC}")
    print()
    
    # Ask if user wants to validate against Azure AD
    print(f"  {Colors.WHITE}Would you like to validate mailboxes against Azure AD?{Colors.NC}")
    print(f"  {Colors.DIM}(This will check if mailboxes exist and are accessible){Colors.NC}")
    print()
    print(f"    {Colors.GREEN}[1]{Colors.NC} Yes, validate mailboxes")
    print(f"    {Colors.BLUE}[2]{Colors.NC} No, just show configuration")
    print()
    
    validate_choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip()
    validate = validate_choice == '1'
    
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'═' * 70}{Colors.NC}")
    print(f"  {Colors.CYAN}{'📬 MAILBOX LIST':^70}{Colors.NC}")
    print(f"  {Colors.CYAN}{'═' * 70}{Colors.NC}")
    print()
    
    # Get access token if validating
    token = None
    if validate:
        print(f"  {Colors.WHITE}Connecting to Microsoft Graph...{Colors.NC}")
        token = get_graph_access_token()
        
        if not token:
            print(f"  {Colors.YELLOW}⚠{Colors.NC} Could not get access token - showing config only")
            validate = False
        else:
            print(f"  {Colors.GREEN}✓{Colors.NC} Connected to Microsoft Graph")
        print()
    
    # Display mailboxes - different headers for validation vs config-only
    if validate:
        print(f"  {Colors.WHITE}{'#':<4} {'UPN':<35} {'Department':<15} {'Status':<15}{Colors.NC}")
    else:
        print(f"  {Colors.WHITE}{'#':<4} {'UPN':<35} {'Department':<15} {'Role':<20}{Colors.NC}")
    print(f"  {Colors.DIM}{'─' * 70}{Colors.NC}")
    
    valid_count = 0
    invalid_count = 0
    
    for idx, user in enumerate(users, 1):
        upn = user.get("upn", "Unknown")
        department = user.get("department", "N/A")
        job_title = user.get("job_title", "N/A")
        
        # Truncate long values
        if len(upn) > 33:
            upn_display = upn[:30] + "..."
        else:
            upn_display = upn
        
        if len(department) > 13:
            dept_display = department[:10] + "..."
        else:
            dept_display = department
        
        if validate and token:
            # Validate mailbox against Azure AD
            status = validate_mailbox_status(upn, token)
            if status["valid"]:
                valid_count += 1
                status_display = f"{Colors.GREEN}✓ Valid{Colors.NC}"
                if status.get("email_count", 0) > 0:
                    status_display += f" ({status['email_count']} emails)"
            else:
                invalid_count += 1
                status_display = f"{Colors.RED}✗ {status.get('error', 'Invalid')}{Colors.NC}"
            print(f"  {idx:<4} {upn_display:<35} {dept_display:<15} {status_display}")
        else:
            # Show job_title instead of status when not validating
            if len(job_title) > 18:
                job_title_display = job_title[:15] + "..."
            else:
                job_title_display = job_title
            print(f"  {idx:<4} {upn_display:<35} {dept_display:<15} {job_title_display}")
    
    print(f"  {Colors.DIM}{'─' * 70}{Colors.NC}")
    print()
    
    # Summary
    if validate:
        print(f"  {Colors.WHITE}Summary:{Colors.NC}")
        print(f"    {Colors.GREEN}✓{Colors.NC} Valid mailboxes: {valid_count}")
        print(f"    {Colors.RED}✗{Colors.NC} Invalid mailboxes: {invalid_count}")
        print(f"    {Colors.DIM}Total configured: {len(users)}{Colors.NC}")
    else:
        print(f"  {Colors.WHITE}Total configured mailboxes: {len(users)}{Colors.NC}")
    
    # Show departments summary
    departments: Dict[str, int] = {}
    for user in users:
        dept = user.get("department", "Unknown")
        departments[dept] = departments.get(dept, 0) + 1
    
    if departments:
        print()
        print(f"  {Colors.WHITE}Departments:{Colors.NC}")
        for dept, count in sorted(departments.items()):
            print(f"    • {dept}: {count} mailbox{'es' if count > 1 else ''}")
    
    print()
    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


def validate_mailbox_status(upn: str, token: str) -> Dict[str, Any]:
    """Validate a mailbox against Azure AD and get email count."""
    import urllib.request
    import urllib.error
    
    result: Dict[str, Any] = {"valid": False, "error": None, "email_count": 0}
    
    try:
        # Check if user exists
        user_url = f"https://graph.microsoft.com/v1.0/users/{upn}"
        req = urllib.request.Request(user_url)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=10) as response:
            # User exists, now check mailbox
            pass
        
        # Try to get email count
        try:
            mail_url = f"https://graph.microsoft.com/v1.0/users/{upn}/mailFolders/inbox/messages?$count=true&$top=1"
            mail_req = urllib.request.Request(mail_url)
            mail_req.add_header("Authorization", f"Bearer {token}")
            mail_req.add_header("Content-Type", "application/json")
            mail_req.add_header("ConsistencyLevel", "eventual")
            
            with urllib.request.urlopen(mail_req, timeout=10) as mail_response:
                mail_data = json.loads(mail_response.read().decode())
                result["email_count"] = mail_data.get("@odata.count", len(mail_data.get("value", [])))
                result["valid"] = True
        except urllib.error.HTTPError as e:
            if e.code == 404:
                result["error"] = "No mailbox"
            elif e.code == 403:
                result["valid"] = True  # User exists but no mail permission
                result["error"] = None
            else:
                result["valid"] = True  # User exists
        except Exception:
            result["valid"] = True  # User exists, mail check failed
            
    except urllib.error.HTTPError as e:
        if e.code == 404:
            result["error"] = "Not found"
        elif e.code == 403:
            result["error"] = "No access"
        else:
            result["error"] = f"Error {e.code}"
    except Exception as e:
        result["error"] = str(e)[:20]
    
    return result


# ============================================================================
# CONFIGURATION EDITING
# ============================================================================

CONFIG_DIR = SCRIPT_DIR.parent / "config"
ENVIRONMENTS_FILE = CONFIG_DIR / "environments.json"
SITES_FILE = CONFIG_DIR / "sites.json"
MAILBOXES_FILE = CONFIG_DIR / "mailboxes.yaml"

def open_file_in_editor(file_path: Path) -> bool:
    """Open a file in the system's default editor."""
    import platform
    
    if not file_path.exists():
        print_error(f"File not found: {file_path}")
        return False
    
    try:
        system = platform.system().lower()
        
        if system == "windows":
            # Try VS Code first, then notepad
            try:
                subprocess.run(["code", str(file_path)], check=True)
            except FileNotFoundError:
                os.startfile(str(file_path))
        elif system == "darwin":  # macOS
            subprocess.run(["open", str(file_path)], check=True)
        else:  # Linux
            # Try common editors
            for editor in ["code", "gedit", "nano", "vim"]:
                try:
                    subprocess.run([editor, str(file_path)], check=True)
                    break
                except FileNotFoundError:
                    continue
        
        return True
    except Exception as e:
        print_error(f"Could not open file: {e}")
        return False

def view_file_contents(file_path: Path) -> None:
    """Display the contents of a configuration file."""
    if not file_path.exists():
        print_error(f"File not found: {file_path}")
        return
    
    try:
        with open(file_path, 'r') as f:
            content = f.read()
        
        print()
        print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
        print(f"  {Colors.WHITE}Contents of {file_path.name}:{Colors.NC}")
        print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
        print()
        
        # Print with line numbers
        for i, line in enumerate(content.split('\n'), 1):
            print(f"  {Colors.DIM}{i:3}{Colors.NC} │ {line}")
        
        print()
        print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    except Exception as e:
        print_error(f"Could not read file: {e}")

def edit_configuration_menu() -> None:
    """Show the configuration editing menu."""
    while True:
        clear_screen()
        print()
        print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
        print(f"  {Colors.WHITE}{Colors.BOLD}⚙️  Edit Configuration{Colors.NC}")
        print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
        print()
        print(f"  {Colors.WHITE}Configuration Files:{Colors.NC}")
        print()
        print(f"  {Colors.CYAN}╭──────────────────────────────────────────────────────────────╮{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.GREEN}[1]{Colors.NC} {Colors.WHITE}📋 Edit environments.json{Colors.NC}                            {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Configure Azure tenants and subscriptions{Colors.NC}            {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.BLUE}[2]{Colors.NC} {Colors.WHITE}📋 Edit sites.json{Colors.NC}                                    {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Define custom SharePoint sites to create{Colors.NC}             {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.MAGENTA}[3]{Colors.NC} {Colors.WHITE}📧 Edit mailboxes.yaml{Colors.NC}                                {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Configure email mailboxes for population{Colors.NC}             {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}├──────────────────────────────────────────────────────────────┤{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.YELLOW}[4]{Colors.NC} {Colors.WHITE}👁️  View environments.json{Colors.NC}                            {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.YELLOW}[5]{Colors.NC} {Colors.WHITE}👁️  View sites.json{Colors.NC}                                   {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.YELLOW}[6]{Colors.NC} {Colors.WHITE}👁️  View mailboxes.yaml{Colors.NC}                               {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}├──────────────────────────────────────────────────────────────┤{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.CYAN}[7]{Colors.NC} {Colors.WHITE}➕ Add new environment{Colors.NC}                               {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}       {Colors.DIM}Interactive wizard to add a tenant{Colors.NC}                   {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.RED}[B]{Colors.NC} {Colors.WHITE}← Back to main menu{Colors.NC}                                 {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}╰──────────────────────────────────────────────────────────────╯{Colors.NC}")
        print()
        
        choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip().lower()
        
        if choice == '1':
            print()
            print_info(f"Opening {ENVIRONMENTS_FILE.name} in editor...")
            if open_file_in_editor(ENVIRONMENTS_FILE):
                print_success("File opened in editor")
            input(f"  {Colors.YELLOW}Press Enter when done editing...{Colors.NC}")
            
        elif choice == '2':
            print()
            print_info(f"Opening {SITES_FILE.name} in editor...")
            if open_file_in_editor(SITES_FILE):
                print_success("File opened in editor")
            input(f"  {Colors.YELLOW}Press Enter when done editing...{Colors.NC}")
            
        elif choice == '3':
            print()
            print_info(f"Opening {MAILBOXES_FILE.name} in editor...")
            if open_file_in_editor(MAILBOXES_FILE):
                print_success("File opened in editor")
            input(f"  {Colors.YELLOW}Press Enter when done editing...{Colors.NC}")
            
        elif choice == '4':
            view_file_contents(ENVIRONMENTS_FILE)
            input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == '5':
            view_file_contents(SITES_FILE)
            input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == '6':
            view_file_contents(MAILBOXES_FILE)
            input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == '7':
            add_environment_wizard()
            
        elif choice == 'b':
            break
        else:
            print_error("Invalid choice")
            input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")

def add_environment_wizard() -> None:
    """Interactive wizard to add a new environment."""
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}➕ Add New Environment{Colors.NC}")
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print()
    print(f"  {Colors.WHITE}This wizard will help you add a new environment configuration.{Colors.NC}")
    print()
    
    # Get environment name
    print(f"  {Colors.YELLOW}Step 1: Environment Name{Colors.NC}")
    name = input(f"  Enter a name for this environment (e.g., 'Production'): ").strip()
    if not name:
        print_error("Name is required")
        input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    print()
    print(f"  {Colors.YELLOW}Step 2: Azure Tenant ID{Colors.NC}")
    print(f"  {Colors.DIM}You can find this in Azure Portal > Azure Active Directory > Overview{Colors.NC}")
    tenant_id = input(f"  Enter Tenant ID (GUID): ").strip()
    
    print()
    print(f"  {Colors.YELLOW}Step 3: Azure Subscription ID{Colors.NC}")
    print(f"  {Colors.DIM}You can find this in Azure Portal > Subscriptions{Colors.NC}")
    subscription_id = input(f"  Enter Subscription ID (GUID): ").strip()
    
    print()
    print(f"  {Colors.YELLOW}Step 4: Resource Group (optional){Colors.NC}")
    resource_group = input(f"  Enter Resource Group name (or leave blank): ").strip()
    
    print()
    print(f"  {Colors.YELLOW}Step 5: M365 Domain (optional){Colors.NC}")
    print(f"  {Colors.DIM}e.g., contoso.onmicrosoft.com{Colors.NC}")
    m365_domain = input(f"  Enter M365 domain (or leave blank): ").strip()
    
    # Create the environment object
    new_env = {
        "name": name,
        "azure": {
            "tenant_id": tenant_id,
            "subscription_id": subscription_id
        }
    }
    
    if resource_group:
        new_env["azure"]["resource_group"] = resource_group
    
    if m365_domain:
        new_env["m365"] = {"domain": m365_domain}
    
    # Load existing environments
    try:
        if ENVIRONMENTS_FILE.exists():
            with open(ENVIRONMENTS_FILE, 'r') as f:
                data = json.load(f)
        else:
            data = {"environments": []}
        
        # Add new environment
        if "environments" not in data:
            data["environments"] = []
        
        data["environments"].append(new_env)
        
        # Save
        with open(ENVIRONMENTS_FILE, 'w') as f:
            json.dump(data, f, indent=2)
        
        print()
        print_success(f"Environment '{name}' added successfully!")
        print_info(f"Saved to: {ENVIRONMENTS_FILE}")
        
    except Exception as e:
        print_error(f"Failed to save environment: {e}")
    
    print()
    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")

# ============================================================================
# SCRIPT EXECUTION
# ============================================================================

def get_script_env() -> dict:
    """Get environment variables for running scripts with Azure CLI path included."""
    env = os.environ.copy()
    
    # Find Azure CLI path and add its directory to PATH
    az_path = find_azure_cli_path()
    if az_path and az_path != "az":
        az_dir = os.path.dirname(az_path)
        current_path = env.get('PATH', '')
        if az_dir not in current_path:
            env['PATH'] = az_dir + os.pathsep + current_path
    
    return env

def run_script(script_name: str, args: list = None) -> None:  # type: ignore
    """Run a Python script with optional arguments."""
    script_path = SCRIPT_DIR / script_name
    
    if not script_path.exists():
        print(f"  {Colors.RED}✗{Colors.NC} Script not found: {script_name}")
        input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    cmd = [sys.executable, str(script_path)]
    if args:
        cmd.extend(args)
    
    print()
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}Running: {Colors.YELLOW}{script_name}{Colors.NC}")
    if args:
        print(f"  {Colors.DIM}Arguments: {' '.join(args)}{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print()
    
    try:
        subprocess.run(cmd, cwd=SCRIPT_DIR, env=get_script_env())
    except KeyboardInterrupt:
        print()
        print(f"  {Colors.YELLOW}⚠{Colors.NC} Operation cancelled")
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Error: {e}")
    
    print()
    input(f"  {Colors.YELLOW}Press Enter to return to menu...{Colors.NC}")

def get_site_filter() -> str:
    """Prompt user for optional site filter."""
    print()
    print(f"  {Colors.WHITE}Filter by site name (optional):{Colors.NC}")
    print(f"  {Colors.DIM}Leave blank to include all sites, or enter a filter (e.g., 'hr', 'finance'){Colors.NC}")
    print()
    filter_input = input(f"  {Colors.YELLOW}Site filter:{Colors.NC} ").strip()
    return filter_input


def get_variation_level() -> str:
    """Prompt user for variation intensity for file/folder generation."""
    print()
    print(f"  {Colors.WHITE}Variation intensity:{Colors.NC}")
    print(f"    {Colors.GREEN}[1]{Colors.NC} Low    (lighter variation)")
    print(f"    {Colors.BLUE}[2]{Colors.NC} Medium (balanced, default)")
    print(f"    {Colors.MAGENTA}[3]{Colors.NC} High   (maximum variation)")
    print()

    while True:
        choice = input(f"  {Colors.YELLOW}Select variation level [1-3, Enter=2]:{Colors.NC} ").strip().lower()
        if choice in ["", "2", "medium", "m"]:
            return "medium"
        if choice in ["1", "low", "l"]:
            return "low"
        if choice in ["3", "high", "h"]:
            return "high"
        print(f"  {Colors.RED}✗{Colors.NC} Invalid selection. Choose 1, 2, or 3.")

# ============================================================================
# AZURE AD DISCOVERY MENU
# ============================================================================

def azure_ad_discovery_menu() -> None:
    """Show the Azure AD user discovery menu."""
    while True:
        clear_screen()
        print()
        print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
        print(f"  {Colors.WHITE}{Colors.BOLD}🔍 Azure AD User Discovery{Colors.NC}")
        print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
        print()
        
        # Check if we have a valid token
        token = get_graph_access_token()
        if not token:
            print(f"  {Colors.RED}✗{Colors.NC} Not authenticated. Please check prerequisites first.")
            print()
            input(f"  {Colors.YELLOW}Press Enter to return to menu...{Colors.NC}")
            return
        
        # Try to load existing cache
        cache_info = _get_azure_ad_cache_info()
        
        if cache_info:
            print(f"  {Colors.GREEN}✓{Colors.NC} Cached data available:")
            print(f"    {Colors.DIM}Users: {cache_info.get('total_users', 0)} ({cache_info.get('mailbox_users', 0)} with mailboxes){Colors.NC}")
            print(f"    {Colors.DIM}Groups: {cache_info.get('groups', 0)}{Colors.NC}")
            print(f"    {Colors.DIM}Last updated: {cache_info.get('timestamp', 'Unknown')}{Colors.NC}")
            print(f"    {Colors.DIM}Cache valid: {'Yes' if cache_info.get('valid', False) else 'No (expired)'}{Colors.NC}")
        else:
            print(f"  {Colors.YELLOW}ℹ{Colors.NC} No cached data. Run discovery to populate.")
        
        print()
        print(f"  {Colors.WHITE}What would you like to do?{Colors.NC}")
        print()
        print(f"    {Colors.GREEN}[1]{Colors.NC} Run full discovery (users + groups)")
        print(f"    {Colors.BLUE}[2]{Colors.NC} Discover users only")
        print(f"    {Colors.YELLOW}[3]{Colors.NC} Discover groups only")
        print(f"    {Colors.MAGENTA}[4]{Colors.NC} View discovery statistics")
        print(f"    {Colors.CYAN}[5]{Colors.NC} View users by department")
        print(f"    {Colors.RED}[6]{Colors.NC} Clear cache")
        print(f"    {Colors.WHITE}[B]{Colors.NC} Back to main menu")
        print()
        
        choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip().lower()
        
        if choice == '1':
            _run_azure_ad_discovery(token, discover_users=True, discover_groups=True)
        elif choice == '2':
            _run_azure_ad_discovery(token, discover_users=True, discover_groups=False)
        elif choice == '3':
            _run_azure_ad_discovery(token, discover_users=False, discover_groups=True)
        elif choice == '4':
            _show_discovery_statistics()
        elif choice == '5':
            _show_users_by_department()
        elif choice == '6':
            _clear_azure_ad_cache()
        elif choice == 'b':
            return
        else:
            print(f"  {Colors.RED}✗{Colors.NC} Invalid choice.")
            input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


def _get_azure_ad_cache_info() -> Optional[Dict]:
    """Get information about the Azure AD cache."""
    try:
        cache_path = SCRIPT_DIR.parent / "config" / ".azure_ad_cache.json"
        if not cache_path.exists():
            return None
        
        with open(cache_path, 'r') as f:
            data = json.load(f)
        
        from datetime import datetime, timedelta
        
        timestamp_str = data.get("timestamp")
        if timestamp_str:
            timestamp = datetime.fromisoformat(timestamp_str)
            ttl_minutes = 60  # Default TTL
            is_valid = datetime.now() - timestamp < timedelta(minutes=ttl_minutes)
            timestamp_display = timestamp.strftime("%Y-%m-%d %H:%M:%S")
        else:
            is_valid = False
            timestamp_display = "Unknown"
        
        users = data.get("users", [])
        mailbox_users = len([u for u in users if u.get("has_mailbox", False)])
        
        return {
            "total_users": len(users),
            "mailbox_users": mailbox_users,
            "groups": len(data.get("groups", [])),
            "timestamp": timestamp_display,
            "valid": is_valid,
        }
    except Exception:
        return None


def _run_azure_ad_discovery(token: str, discover_users: bool = True, discover_groups: bool = True) -> None:
    """Run Azure AD discovery."""
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}Running Azure AD Discovery{Colors.NC}")
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print()
    
    try:
        # Import the discovery module
        from email_generator.azure_ad_discovery import AzureADDiscovery
        from email_generator.config import load_mailbox_config
        
        # Load configuration
        try:
            config = load_mailbox_config()
        except Exception:
            config = {}
        
        # Create discovery instance
        discovery = AzureADDiscovery(token, config)
        
        def progress_callback(current: int, total: int, message: str) -> None:
            percentage = (current / total) * 100 if total > 0 else 0
            print(f"\r  {Colors.CYAN}[{percentage:5.1f}%]{Colors.NC} {message:<50}", end='', flush=True)
        
        if discover_users:
            print(f"  {Colors.BLUE}ℹ{Colors.NC} Discovering users from Azure AD...")
            print()
            users = discovery.discover_users(
                validate_mailboxes=True,
                progress_callback=progress_callback
            )
            print()
            print(f"  {Colors.GREEN}✓{Colors.NC} Found {len(users)} users")
            mailbox_count = len([u for u in users if u.has_mailbox])
            print(f"    {Colors.DIM}{mailbox_count} users have mailboxes{Colors.NC}")
            print()
        
        if discover_groups:
            print(f"  {Colors.BLUE}ℹ{Colors.NC} Discovering groups from Azure AD...")
            print()
            groups = discovery.discover_groups(
                progress_callback=progress_callback
            )
            print()
            print(f"  {Colors.GREEN}✓{Colors.NC} Found {len(groups)} groups")
            print()
        
        # Save cache
        discovery._save_cache()
        print(f"  {Colors.GREEN}✓{Colors.NC} Cache saved successfully")
        
    except ImportError as e:
        print(f"  {Colors.RED}✗{Colors.NC} Failed to import discovery module: {e}")
        print(f"    {Colors.DIM}Make sure PyYAML is installed: pip install pyyaml{Colors.NC}")
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Discovery failed: {e}")
    
    print()
    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


def _show_discovery_statistics() -> None:
    """Show discovery statistics."""
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}Discovery Statistics{Colors.NC}")
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print()
    
    cache_info = _get_azure_ad_cache_info()
    
    if not cache_info:
        print(f"  {Colors.YELLOW}ℹ{Colors.NC} No cached data available. Run discovery first.")
    else:
        print(f"  {Colors.WHITE}{Colors.BOLD}Users:{Colors.NC}")
        print(f"    Total users:        {cache_info.get('total_users', 0)}")
        print(f"    With mailboxes:     {cache_info.get('mailbox_users', 0)}")
        print(f"    Without mailboxes:  {cache_info.get('total_users', 0) - cache_info.get('mailbox_users', 0)}")
        print()
        print(f"  {Colors.WHITE}{Colors.BOLD}Groups:{Colors.NC}")
        print(f"    Total groups:       {cache_info.get('groups', 0)}")
        print()
        print(f"  {Colors.WHITE}{Colors.BOLD}Cache:{Colors.NC}")
        print(f"    Last updated:       {cache_info.get('timestamp', 'Unknown')}")
        print(f"    Status:             {'Valid' if cache_info.get('valid', False) else 'Expired'}")
    
    print()
    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


def _show_users_by_department() -> None:
    """Show users grouped by department."""
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print(f"  {Colors.WHITE}{Colors.BOLD}Users by Department{Colors.NC}")
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print()
    
    try:
        cache_path = SCRIPT_DIR.parent / "config" / ".azure_ad_cache.json"
        if not cache_path.exists():
            print(f"  {Colors.YELLOW}ℹ{Colors.NC} No cached data available. Run discovery first.")
            print()
            input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            return
        
        with open(cache_path, 'r') as f:
            data = json.load(f)
        
        users = data.get("users", [])
        
        # Group by department
        departments: Dict[str, Dict[str, int]] = {}
        for user in users:
            dept = user.get("department") or "Unknown"
            if dept not in departments:
                departments[dept] = {"total": 0, "mailbox": 0}
            departments[dept]["total"] += 1
            if user.get("has_mailbox", False):
                departments[dept]["mailbox"] += 1
        
        # Sort by total users
        sorted_depts = sorted(departments.items(), key=lambda x: x[1]["total"], reverse=True)
        
        print(f"  {'Department':<35} {'Total':>8} {'Mailbox':>10}")
        print(f"  {'-' * 35} {'-' * 8} {'-' * 10}")
        
        for dept, counts in sorted_depts:
            dept_display = dept[:33] + ".." if len(dept) > 35 else dept
            print(f"  {dept_display:<35} {counts['total']:>8} {counts['mailbox']:>10}")
        
        print()
        print(f"  {Colors.DIM}Total: {len(users)} users across {len(departments)} departments{Colors.NC}")
        
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Failed to load data: {e}")
    
    print()
    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


def _clear_azure_ad_cache() -> None:
    """Clear the Azure AD cache."""
    print()
    print(f"  {Colors.YELLOW}⚠{Colors.NC} This will delete all cached Azure AD data.")
    confirm = input(f"  {Colors.YELLOW}Type 'CLEAR' to confirm:{Colors.NC} ").strip()
    
    if confirm == 'CLEAR':
        try:
            cache_path = SCRIPT_DIR.parent / "config" / ".azure_ad_cache.json"
            if cache_path.exists():
                cache_path.unlink()
                print(f"  {Colors.GREEN}✓{Colors.NC} Cache cleared successfully")
            else:
                print(f"  {Colors.YELLOW}ℹ{Colors.NC} No cache file found")
        except Exception as e:
            print(f"  {Colors.RED}✗{Colors.NC} Failed to clear cache: {e}")
    else:
        print(f"  {Colors.YELLOW}Operation cancelled{Colors.NC}")
    
    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def main() -> None:
    """Main entry point."""
    # Initial prerequisites check (silent)
    prereq_status = check_prerequisites(auto_install=False)
    
    while True:
        clear_screen()
        print_logo()
        print_menu(prereq_status)
        
        choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip().lower()
        
        if choice == '0':
            # Check & Install Prerequisites
            run_prerequisites_check_menu()
            # Refresh status after check
            prereq_status = check_prerequisites(auto_install=False)
            
        elif choice == '1':
            # Create SharePoint Sites
            clear_screen()
            print()
            print(f"  {Colors.GREEN}{Colors.BOLD}🏗️  Step 1: Create SharePoint Sites{Colors.NC}")
            print(f"  {Colors.CYAN}{'─' * 40}{Colors.NC}")
            print()
            print(f"  {Colors.WHITE}How would you like to create sites?{Colors.NC}")
            print()
            print(f"    {Colors.GREEN}[1]{Colors.NC} Interactive mode (guided setup)")
            print(f"    {Colors.BLUE}[2]{Colors.NC} Quick: Create 5 random sites")
            print(f"    {Colors.YELLOW}[3]{Colors.NC} Quick: Create 10 random sites")
            print(f"    {Colors.MAGENTA}[4]{Colors.NC} Custom: Specify number of random sites")
            print(f"    {Colors.WHITE}[5]{Colors.NC} Use configuration file (config/sites.json)")
            print(f"    {Colors.RED}[B]{Colors.NC} Back to main menu")
            print()
            
            sub_choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip().lower()
            
            if sub_choice == '1':
                run_script("deploy.py", ["--skip-prerequisites"])
            elif sub_choice == '2':
                run_script("deploy.py", ["--skip-prerequisites", "--random", "5", "--auto-approve"])
            elif sub_choice == '3':
                run_script("deploy.py", ["--skip-prerequisites", "--random", "10", "--auto-approve"])
            elif sub_choice == '4':
                print()
                count = input(f"  {Colors.YELLOW}Number of sites (1-39):{Colors.NC} ").strip()
                if count.isdigit() and 1 <= int(count) <= 39:
                    run_script("deploy.py", ["--skip-prerequisites", "--random", count])
                else:
                    print(f"  {Colors.RED}✗{Colors.NC} Invalid number. Must be between 1 and 39.")
                    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            elif sub_choice == '5':
                run_script("deploy.py", ["--skip-prerequisites", "--config", str(SCRIPT_DIR.parent / "config" / "sites.json")])
            # else: back to menu
            
        elif choice == '2':
            # Populate Sites with Files
            clear_screen()
            print()
            print(f"  {Colors.BLUE}{Colors.BOLD}📄 Step 2: Populate Sites with Files{Colors.NC}")
            print(f"  {Colors.CYAN}{'─' * 40}{Colors.NC}")
            print()
            print(f"  {Colors.WHITE}How many files would you like to create?{Colors.NC}")
            print()
            print(f"    {Colors.GREEN}[1]{Colors.NC} Interactive mode")
            print(f"    {Colors.BLUE}[2]{Colors.NC} Quick: Create 25 files")
            print(f"    {Colors.YELLOW}[3]{Colors.NC} Quick: Create 50 files")
            print(f"    {Colors.MAGENTA}[4]{Colors.NC} Quick: Create 100 files")
            print(f"    {Colors.CYAN}[5]{Colors.NC} Custom: Specify number of files")
            print(f"    {Colors.RED}[B]{Colors.NC} Back to main menu")
            print()
            
            sub_choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip().lower()
            
            if sub_choice == '1':
                variation_level = get_variation_level()
                run_script("populate_files.py", ["--variation-level", variation_level])
            elif sub_choice == '2':
                variation_level = get_variation_level()
                site_filter = get_site_filter()
                args = ["--files", "25", "--variation-level", variation_level]
                if site_filter:
                    args.extend(["--site", site_filter])
                run_script("populate_files.py", args)
            elif sub_choice == '3':
                variation_level = get_variation_level()
                site_filter = get_site_filter()
                args = ["--files", "50", "--variation-level", variation_level]
                if site_filter:
                    args.extend(["--site", site_filter])
                run_script("populate_files.py", args)
            elif sub_choice == '4':
                variation_level = get_variation_level()
                site_filter = get_site_filter()
                args = ["--files", "100", "--variation-level", variation_level]
                if site_filter:
                    args.extend(["--site", site_filter])
                run_script("populate_files.py", args)
            elif sub_choice == '5':
                print()
                count = input(f"  {Colors.YELLOW}Number of files (1-1000):{Colors.NC} ").strip()
                if count.isdigit() and 1 <= int(count) <= 1000:
                    variation_level = get_variation_level()
                    site_filter = get_site_filter()
                    args = ["--files", count, "--variation-level", variation_level]
                    if site_filter:
                        args.extend(["--site", site_filter])
                    run_script("populate_files.py", args)
                else:
                    print(f"  {Colors.RED}✗{Colors.NC} Invalid number. Must be between 1 and 1000.")
                    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            # else: back to menu
            
        elif choice == '3':
            # Delete Files or Sites
            clear_screen()
            print()
            print(f"  {Colors.RED}{Colors.BOLD}🗑️  Step 3: Delete Files or Sites{Colors.NC}")
            print(f"  {Colors.CYAN}{'─' * 40}{Colors.NC}")
            print()
            print(f"  {Colors.YELLOW}⚠ WARNING: These operations are DESTRUCTIVE!{Colors.NC}")
            print()
            print(f"    {Colors.GREEN}[1]{Colors.NC} Interactive mode (safest)")
            print(f"    {Colors.BLUE}[2]{Colors.NC} Select specific SITES to work with")
            print(f"    {Colors.YELLOW}[3]{Colors.NC} Select specific FILES to delete")
            print(f"    {Colors.MAGENTA}[4]{Colors.NC} Delete ALL files from sites (keeps sites)")
            print(f"    {Colors.RED}[5]{Colors.NC} Delete SharePoint SITES (requires admin)")
            print()
            print(f"  {Colors.CYAN}{'─' * 40}{Colors.NC}")
            print(f"  {Colors.WHITE}{Colors.BOLD}Recycle Bin Operations:{Colors.NC}")
            print()
            print(f"    {Colors.CYAN}[6]{Colors.NC} Purge M365 Groups recycle bin (Azure AD)")
            print(f"    {Colors.CYAN}[7]{Colors.NC} Purge SharePoint site recycle bin")
            print(f"    {Colors.CYAN}[8]{Colors.NC} Purge site files/folders recycle bin")
            print(f"    {Colors.WHITE}[B]{Colors.NC} Back to main menu")
            print()
            
            sub_choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip().lower()
            
            if sub_choice == '1':
                run_script("cleanup.py")
            elif sub_choice == '2':
                run_script("cleanup.py", ["--select-sites"])
            elif sub_choice == '3':
                site_filter = get_site_filter()
                args = ["--select-files"]
                if site_filter:
                    args.extend(["--site", site_filter])
                run_script("cleanup.py", args)
            elif sub_choice == '4':
                site_filter = get_site_filter()
                args = ["--delete-files"]
                if site_filter:
                    args.extend(["--site", site_filter])
                run_script("cleanup.py", args)
            elif sub_choice == '5':
                site_filter = get_site_filter()
                args = ["--delete-sites"]
                if site_filter:
                    args.extend(["--site", site_filter])
                run_script("cleanup.py", args)
            elif sub_choice == '6':
                # Purge M365 Groups recycle bin
                run_script("cleanup.py", ["--purge-deleted"])
            elif sub_choice == '7':
                # Purge SharePoint site recycle bin
                print()
                print(f"  {Colors.WHITE}Enter your SharePoint tenant name{Colors.NC}")
                print(f"  {Colors.DIM}(e.g., 'contoso' for contoso.sharepoint.com){Colors.NC}")
                print()
                tenant = input(f"  {Colors.YELLOW}Tenant name:{Colors.NC} ").strip()
                if tenant:
                    run_script("cleanup.py", ["--purge-spo-recycle", "--tenant", tenant])
                else:
                    print(f"  {Colors.RED}✗{Colors.NC} Tenant name is required")
                    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            elif sub_choice == '8':
                # Purge site files/folders recycle bin
                site_filter = get_site_filter()
                args = ["--purge-site-recycle", "--non-interactive", "--auto-setup-cert", "--yes"]
                if site_filter:
                    args.extend(["--site", site_filter])
                run_script("cleanup.py", args)
            # else: back to menu
            
        elif choice == '4':
            # List SharePoint Sites
            list_sharepoint_sites_menu()
            
        elif choice == '5':
            # List Files in Sites
            list_files_in_sites_menu()
            
        elif choice == '6':
            # Populate Mailboxes with Emails
            clear_screen()
            print()
            print(f"  {Colors.CYAN}{Colors.BOLD}📧 Populate Mailboxes with Emails{Colors.NC}")
            print(f"  {Colors.CYAN}{'─' * 40}{Colors.NC}")
            print()
            print(f"  {Colors.WHITE}How would you like to populate emails?{Colors.NC}")
            print()
            print(f"    {Colors.GREEN}[1]{Colors.NC} Interactive mode (guided setup)")
            print(f"    {Colors.BLUE}[2]{Colors.NC} Quick: 50 emails per mailbox (all mailboxes)")
            print(f"    {Colors.YELLOW}[3]{Colors.NC} Quick: 100 emails per mailbox (all mailboxes)")
            print(f"    {Colors.MAGENTA}[4]{Colors.NC} Custom: Specify number of emails")
            print(f"    {Colors.CYAN}[5]{Colors.NC} Custom: Select specific mailboxes")
            print(f"    {Colors.RED}[B]{Colors.NC} Back to main menu")
            print()
            
            sub_choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip().lower()
            
            if sub_choice == '1':
                run_script("populate_emails.py")
            elif sub_choice == '2':
                run_script("populate_emails.py", ["--all", "--emails", "50"])
            elif sub_choice == '3':
                run_script("populate_emails.py", ["--all", "--emails", "100"])
            elif sub_choice == '4':
                print()
                count = input(f"  {Colors.YELLOW}Number of emails per mailbox (1-500):{Colors.NC} ").strip()
                if count.isdigit() and 1 <= int(count) <= 500:
                    run_script("populate_emails.py", ["--all", "--emails", count])
                else:
                    print(f"  {Colors.RED}✗{Colors.NC} Invalid number. Must be between 1 and 500.")
                    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            elif sub_choice == '5':
                print()
                print(f"  {Colors.WHITE}Enter mailbox UPNs (comma-separated):{Colors.NC}")
                print(f"  {Colors.DIM}Example: user1@contoso.com, user2@contoso.com{Colors.NC}")
                print()
                mailboxes = input(f"  {Colors.YELLOW}Mailboxes:{Colors.NC} ").strip()
                if mailboxes:
                    print()
                    count = input(f"  {Colors.YELLOW}Number of emails per mailbox (1-500):{Colors.NC} ").strip()
                    if count.isdigit() and 1 <= int(count) <= 500:
                        run_script("populate_emails.py", ["--mailboxes", mailboxes, "--emails", count])
                    else:
                        print(f"  {Colors.RED}✗{Colors.NC} Invalid number. Must be between 1 and 500.")
                        input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
                else:
                    print(f"  {Colors.RED}✗{Colors.NC} No mailboxes specified.")
                    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            # else: back to menu
            
        elif choice == '7':
            # Delete Emails from Mailboxes
            clear_screen()
            print()
            print(f"  {Colors.RED}{Colors.BOLD}🗑️  Delete Emails from Mailboxes{Colors.NC}")
            print(f"  {Colors.CYAN}{'─' * 40}{Colors.NC}")
            print()
            print(f"  {Colors.YELLOW}⚠ WARNING: This will DELETE emails!{Colors.NC}")
            print()
            print(f"  {Colors.WHITE}How would you like to delete emails?{Colors.NC}")
            print()
            print(f"    {Colors.GREEN}[1]{Colors.NC} Interactive mode (guided setup)")
            print(f"    {Colors.BLUE}[2]{Colors.NC} Delete from all mailboxes (inbox)")
            print(f"    {Colors.YELLOW}[3]{Colors.NC} Delete from specific mailboxes")
            print(f"    {Colors.MAGENTA}[4]{Colors.NC} Empty Deleted Items (all mailboxes)")
            print(f"    {Colors.RED}[B]{Colors.NC} Back to main menu")
            print()
            
            sub_choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip().lower()
            
            if sub_choice == '1':
                run_script("cleanup_emails.py")
            elif sub_choice == '2':
                print()
                print(f"  {Colors.RED}⚠ This will delete ALL emails from inbox in ALL mailboxes!{Colors.NC}")
                confirm = input(f"  {Colors.YELLOW}Type 'DELETE' to confirm:{Colors.NC} ").strip()
                if confirm == 'DELETE':
                    run_script("cleanup_emails.py", ["--all", "--folder", "inbox"])
                else:
                    print(f"  {Colors.YELLOW}Operation cancelled{Colors.NC}")
                    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            elif sub_choice == '3':
                print()
                print(f"  {Colors.WHITE}Enter mailbox UPNs (comma-separated):{Colors.NC}")
                print(f"  {Colors.DIM}Example: user1@contoso.com, user2@contoso.com{Colors.NC}")
                print()
                mailboxes = input(f"  {Colors.YELLOW}Mailboxes:{Colors.NC} ").strip()
                if mailboxes:
                    run_script("cleanup_emails.py", ["--mailboxes", mailboxes])
                else:
                    print(f"  {Colors.RED}✗{Colors.NC} No mailboxes specified.")
                    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            elif sub_choice == '4':
                print()
                print(f"  {Colors.RED}⚠ This will PERMANENTLY delete all items in Deleted Items!{Colors.NC}")
                confirm = input(f"  {Colors.YELLOW}Type 'EMPTY' to confirm:{Colors.NC} ").strip()
                if confirm == 'EMPTY':
                    run_script("cleanup_emails.py", ["--all", "--folder", "deleteditems", "--permanent"])
                else:
                    print(f"  {Colors.YELLOW}Operation cancelled{Colors.NC}")
                    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            # else: back to menu
            
        elif choice == '8':
            # List Mailboxes
            list_mailboxes_menu()
            
        elif choice == '9':
            # Azure AD User Discovery
            azure_ad_discovery_menu()
            
        elif choice == 'c':
            # Edit Configuration
            edit_configuration_menu()
            
        elif choice == 'a':
            # Manage App Registration
            manage_app_registration_menu()
            # Refresh prereq status after app management
            prereq_status = check_prerequisites(auto_install=False)
            
        elif choice == 'h':
            print_help()
            
        elif choice == 'q':
            clear_screen()
            print()
            print(f"  {Colors.GREEN}Thank you for using SharePoint Sites Management Tool!{Colors.NC}")
            print()
            sys.exit(0)
            
        else:
            print(f"  {Colors.RED}✗{Colors.NC} Invalid choice. Please try again.")
            input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print()
        print(f"  {Colors.YELLOW}⚠{Colors.NC} Goodbye!")
        sys.exit(0)
