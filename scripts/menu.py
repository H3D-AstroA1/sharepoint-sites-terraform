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

def clear_screen() -> None:
    """Clear the terminal screen."""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_logo() -> None:
    """Print the SharePoint logo/banner."""
    print()
    print(f"  {Colors.CYAN}╔══════════════════════════════════════════════════════════════╗{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}                                                              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}███████╗{Colors.NC}{Colors.BLUE}██╗  ██╗{Colors.NC}{Colors.YELLOW}  █████╗ {Colors.NC}{Colors.MAGENTA}██████╗ {Colors.NC}{Colors.CYAN}███████╗{Colors.NC}              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}██╔════╝{Colors.NC}{Colors.BLUE}██║  ██║{Colors.NC}{Colors.YELLOW} ██╔══██╗{Colors.NC}{Colors.MAGENTA}██╔══██╗{Colors.NC}{Colors.CYAN}██╔════╝{Colors.NC}              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}███████╗{Colors.NC}{Colors.BLUE}███████║{Colors.NC}{Colors.YELLOW} ███████║{Colors.NC}{Colors.MAGENTA}██████╔╝{Colors.NC}{Colors.CYAN}█████╗  {Colors.NC}              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}╚════██║{Colors.NC}{Colors.BLUE}██╔══██║{Colors.NC}{Colors.YELLOW} ██╔══██║{Colors.NC}{Colors.MAGENTA}██╔══██╗{Colors.NC}{Colors.CYAN}██╔══╝  {Colors.NC}              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}███████║{Colors.NC}{Colors.BLUE}██║  ██║{Colors.NC}{Colors.YELLOW} ██║  ██║{Colors.NC}{Colors.MAGENTA}██║  ██║{Colors.NC}{Colors.CYAN}███████╗{Colors.NC}              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.GREEN}╚══════╝{Colors.NC}{Colors.BLUE}╚═╝  ╚═╝{Colors.NC}{Colors.YELLOW} ╚═╝  ╚═╝{Colors.NC}{Colors.MAGENTA}╚═╝  ╚═╝{Colors.NC}{Colors.CYAN}╚══════╝{Colors.NC}              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}                                                              {Colors.CYAN}║{Colors.NC}")
    print(f"  {Colors.CYAN}║{Colors.NC}   {Colors.WHITE}{Colors.BOLD}SharePoint Sites Management Tool{Colors.NC}                         {Colors.CYAN}║{Colors.NC}")
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

# Our custom app name for SharePoint management
CUSTOM_APP_NAME = "SharePoint-Sites-Terraform-Tool"

# Config file to store custom app details
APP_CONFIG_FILE = SCRIPT_DIR / ".app_config.json"

# Microsoft Graph API ID (constant across all tenants)
MICROSOFT_GRAPH_API_ID = "00000003-0000-0000-c000-000000000000"

# Microsoft Graph API permission IDs (Application permissions)
# These are the GUIDs for the specific permissions we need
GRAPH_PERMISSION_IDS = {
    "Sites.Read.All": "332a536c-c7ef-4017-ab91-336970924f0d",
    "Sites.ReadWrite.All": "9492366f-7969-46a4-8d15-ed1a20078fff",
    "Files.ReadWrite.All": "75359482-378d-4052-8f01-80520e7db3cd",
    "Group.Read.All": "5b567255-7703-4780-807c-7be8301ae99b",
    "Group.ReadWrite.All": "62a82d76-70ea-41e2-9197-370581804d09"
}

# Required permissions for SharePoint operations
REQUIRED_GRAPH_PERMISSIONS = [
    "Sites.Read.All",
    "Sites.ReadWrite.All",
    "Files.ReadWrite.All",
    "Group.Read.All",
    "Group.ReadWrite.All"
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
        
        # Create app with a web redirect URI (required for admin consent)
        # Using portal.azure.com as redirect - shows Azure Portal after consent (cleaner UX)
        create_result = subprocess.run(
            [az_path, "ad", "app", "create",
             "--display-name", CUSTOM_APP_NAME,
             "--sign-in-audience", "AzureADMyOrg",
             "--web-redirect-uris", "https://portal.azure.com",
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
        print(f"  {Colors.GREEN}✓{Colors.NC} Custom app registered: {app_config.get('app_id', 'Unknown')}")
    
    result = check_graph_permissions()
    
    if result["has_permissions"]:
        print(f"  {Colors.GREEN}✓{Colors.NC} Graph API permissions: OK")
        print(f"  {Colors.GREEN}✓{Colors.NC} Can access SharePoint sites: Yes")
    else:
        print(f"  {Colors.RED}✗{Colors.NC} Graph API permissions: Missing")
        if result["error"]:
            print(f"    {Colors.DIM}Error: {result['error']}{Colors.NC}")
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

def check_prerequisites(auto_install: bool = False) -> dict:
    """Check all prerequisites and optionally install missing ones."""
    results = {
        "python": {"installed": True, "version": f"Python {sys.version.split()[0]}"},
        "azure_cli": {"installed": False, "version": None},
        "terraform": {"installed": False, "version": None},
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
    
    graph_ok = results.get("graph_permissions", {}).get("has_permissions", False)
    
    if all_installed and results["azure_login"]["logged_in"] and graph_ok:
        print(f"  {Colors.GREEN}{Colors.BOLD}✓ All prerequisites met! Ready to proceed.{Colors.NC}")
    elif all_installed and results["azure_login"]["logged_in"] and not graph_ok:
        print(f"  {Colors.YELLOW}{Colors.BOLD}⚠ Graph API permissions missing.{Colors.NC}")
        print(f"  {Colors.DIM}  Setup will be offered below, or use [A] from main menu.{Colors.NC}")
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
    print(f"  {Colors.YELLOW}{Colors.BOLD}Quick Commands:{Colors.NC}")
    print(f"  {Colors.DIM}────────────────{Colors.NC}")
    print(f"    {Colors.CYAN}python deploy.py --random 10{Colors.NC}     Create 10 random sites")
    print(f"    {Colors.CYAN}python populate_files.py --files 50{Colors.NC}  Add 50 files")
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
            # Check permissions
            print()
            if not check_azure_login():
                print_error("Please log into Azure first")
            else:
                result = check_graph_permissions()
                if result.get("has_permissions"):
                    print(f"  {Colors.GREEN}✓{Colors.NC} Permissions are working correctly!")
                    print(f"  {Colors.GREEN}✓{Colors.NC} You can access SharePoint sites.")
                else:
                    print(f"  {Colors.RED}✗{Colors.NC} Permissions are not working.")
                    if result.get("error"):
                        print(f"    {Colors.DIM}Error: {result['error']}{Colors.NC}")
                    print()
                    print(f"  {Colors.WHITE}Use option [1] to set up app registration.{Colors.NC}")
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
            if app_config:
                print(f"  App to delete: {app_config.get('app_id', 'Unknown')}")
            elif existing_app:
                print(f"  App to delete: {existing_app.get('appId', 'Unknown')}")
            print()
            print(f"  {Colors.YELLOW}This action cannot be undone.{Colors.NC}")
            print(f"  {Colors.DIM}You will need to set up the app again to use SharePoint features.{Colors.NC}")
            print()
            
            confirm = input(f"  Type 'DELETE' to confirm: ").strip()
            
            if confirm == 'DELETE':
                if delete_custom_app():
                    print()
                    print(f"  {Colors.GREEN}✓{Colors.NC} App registration deleted successfully!")
                else:
                    print()
                    print(f"  {Colors.RED}✗{Colors.NC} Failed to delete app registration.")
            else:
                print()
                print(f"  {Colors.DIM}Deletion cancelled.{Colors.NC}")
            
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
    
    # Try SharePoint Sites API first
    sites = []
    try:
        url = "https://graph.microsoft.com/v1.0/sites?search=*&$top=100"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            sites = data.get("value", [])
    except urllib.error.HTTPError as e:
        if e.code == 403:
            # Fall back to M365 Groups API
            pass
        else:
            print(f"  {Colors.RED}✗{Colors.NC} Error: {e.code} - {e.reason}")
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Error: {str(e)}")
    
    # If no sites from Sites API, try M365 Groups
    if not sites:
        print(f"  {Colors.YELLOW}ℹ{Colors.NC} Using Microsoft 365 Groups API...")
        try:
            url = "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c eq 'Unified')&$select=id,displayName,createdDateTime,visibility&$top=100"
            req = urllib.request.Request(url)
            req.add_header("Authorization", f"Bearer {token}")
            req.add_header("Content-Type", "application/json")
            
            with urllib.request.urlopen(req, timeout=30) as response:
                data = json.loads(response.read().decode())
                groups = data.get("value", [])
                
                # Convert groups to sites format
                for group in groups:
                    try:
                        site_url = f"https://graph.microsoft.com/v1.0/groups/{group['id']}/sites/root"
                        site_req = urllib.request.Request(site_url)
                        site_req.add_header("Authorization", f"Bearer {token}")
                        
                        with urllib.request.urlopen(site_req, timeout=10) as site_response:
                            site_data = json.loads(site_response.read().decode())
                            sites.append({
                                "displayName": group.get("displayName", "Unknown"),
                                "webUrl": site_data.get("webUrl", ""),
                                "createdDateTime": group.get("createdDateTime", ""),
                                "visibility": group.get("visibility", ""),
                                "isGroup": True
                            })
                    except Exception:
                        # Group might not have a SharePoint site
                        pass
        except Exception as e:
            print(f"  {Colors.RED}✗{Colors.NC} Error: {str(e)}")
    
    if not sites:
        print()
        print(f"  {Colors.YELLOW}No SharePoint sites found.{Colors.NC}")
        print()
        print(f"  {Colors.WHITE}Possible reasons:{Colors.NC}")
        print(f"    • No sites have been created yet")
        print(f"    • Missing permissions (Sites.Read.All or Group.Read.All)")
        print(f"    • Use option [A] to set up app registration with proper permissions")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    # Display sites
    print()
    print(f"  {Colors.GREEN}✓{Colors.NC} Found {len(sites)} SharePoint sites")
    print()
    print(f"  {Colors.WHITE}{'─' * 70}{Colors.NC}")
    print()
    
    for i, site in enumerate(sites, 1):
        name = site.get("displayName", site.get("name", "Unknown"))
        web_url = site.get("webUrl", "")
        visibility = site.get("visibility", "")
        is_group = site.get("isGroup", False)
        
        # Icon based on visibility
        if visibility == "Private":
            icon = "🔒"
        else:
            icon = "🌐"
        
        print(f"  [{i:2}] {icon} {name}")
        print(f"       {Colors.DIM}{web_url}{Colors.NC}")
        print()
    
    print(f"  {Colors.WHITE}{'─' * 70}{Colors.NC}")
    print()
    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


def list_files_in_sites_menu() -> None:
    """List files in SharePoint sites with a clean, friendly interface."""
    import urllib.request
    import urllib.error
    
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print(f"  {Colors.CYAN}{'📁 FILES IN SHAREPOINT SITES':^60}{Colors.NC}")
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
    
    # Get sites first (using M365 Groups API as fallback)
    sites = []
    try:
        url = "https://graph.microsoft.com/v1.0/sites?search=*&$top=100"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            sites = data.get("value", [])
    except urllib.error.HTTPError as e:
        if e.code == 403:
            pass  # Fall back to M365 Groups
    except Exception:
        pass
    
    # If no sites from Sites API, try M365 Groups
    if not sites:
        print(f"  {Colors.YELLOW}ℹ{Colors.NC} Using Microsoft 365 Groups API...")
        try:
            url = "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c eq 'Unified')&$select=id,displayName&$top=100"
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
                            sites.append({
                                "id": site_data.get("id", ""),
                                "displayName": group.get("displayName", "Unknown"),
                                "webUrl": site_data.get("webUrl", "")
                            })
                    except Exception:
                        pass
        except Exception as e:
            print(f"  {Colors.RED}✗{Colors.NC} Error: {str(e)}")
    
    if not sites:
        print()
        print(f"  {Colors.YELLOW}No SharePoint sites found.{Colors.NC}")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    print(f"  {Colors.GREEN}✓{Colors.NC} Found {len(sites)} SharePoint sites")
    print()
    
    # Let user select a site
    print(f"  {Colors.WHITE}Select a site to view files:{Colors.NC}")
    print()
    for i, site in enumerate(sites, 1):
        name = site.get("displayName", site.get("name", "Unknown"))
        print(f"    [{i:2}] {name}")
    print()
    print(f"    [{Colors.RED}B{Colors.NC}]  Back to main menu")
    print()
    
    choice = input(f"  {Colors.YELLOW}Enter site number:{Colors.NC} ").strip().lower()
    
    if choice == 'b' or not choice:
        return
    
    try:
        site_idx = int(choice) - 1
        if site_idx < 0 or site_idx >= len(sites):
            print(f"  {Colors.RED}✗{Colors.NC} Invalid selection")
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            return
    except ValueError:
        print(f"  {Colors.RED}✗{Colors.NC} Invalid input")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    selected_site = sites[site_idx]
    site_id = selected_site.get("id", "")
    site_name = selected_site.get("displayName", "Unknown")
    
    if not site_id:
        print(f"  {Colors.RED}✗{Colors.NC} Could not get site ID")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    # Get files from the site's document library
    print()
    print(f"  {Colors.WHITE}Loading files from '{site_name}'...{Colors.NC}")
    
    files = []
    try:
        # Get the default document library (drive)
        drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
        req = urllib.request.Request(drive_url)
        req.add_header("Authorization", f"Bearer {token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            files = data.get("value", [])
    except urllib.error.HTTPError as e:
        print(f"  {Colors.RED}✗{Colors.NC} Error accessing files: {e.code} - {e.reason}")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    except Exception as e:
        print(f"  {Colors.RED}✗{Colors.NC} Error: {str(e)}")
        input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
        return
    
    # Display files
    clear_screen()
    print()
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print(f"  {Colors.CYAN}{'📁 FILES IN: ' + site_name[:45]:^60}{Colors.NC}")
    print(f"  {Colors.CYAN}{'═' * 60}{Colors.NC}")
    print()
    
    if not files:
        print(f"  {Colors.YELLOW}No files found in this site's document library.{Colors.NC}")
    else:
        print(f"  {Colors.GREEN}✓{Colors.NC} Found {len(files)} items")
        print()
        print(f"  {Colors.WHITE}{'─' * 70}{Colors.NC}")
        print()
        
        for i, item in enumerate(files, 1):
            name = item.get("name", "Unknown")
            is_folder = "folder" in item
            size = item.get("size", 0)
            
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
            
            print(f"  [{i:3}] {icon} {name} {Colors.DIM}{size_str}{Colors.NC}")
        
        print()
        print(f"  {Colors.WHITE}{'─' * 70}{Colors.NC}")
    
    print()
    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")


# ============================================================================
# CONFIGURATION EDITING
# ============================================================================

CONFIG_DIR = SCRIPT_DIR.parent / "config"
ENVIRONMENTS_FILE = CONFIG_DIR / "environments.json"
SITES_FILE = CONFIG_DIR / "sites.json"

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
        print(f"  {Colors.CYAN}├──────────────────────────────────────────────────────────────┤{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.YELLOW}[3]{Colors.NC} {Colors.WHITE}👁️  View environments.json{Colors.NC}                            {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.YELLOW}[4]{Colors.NC} {Colors.WHITE}👁️  View sites.json{Colors.NC}                                   {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}├──────────────────────────────────────────────────────────────┤{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}                                                              {Colors.CYAN}│{Colors.NC}")
        print(f"  {Colors.CYAN}│{Colors.NC}   {Colors.MAGENTA}[5]{Colors.NC} {Colors.WHITE}➕ Add new environment{Colors.NC}                               {Colors.CYAN}│{Colors.NC}")
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
            view_file_contents(ENVIRONMENTS_FILE)
            input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == '4':
            view_file_contents(SITES_FILE)
            input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            
        elif choice == '5':
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
                run_script("populate_files.py")
            elif sub_choice == '2':
                site_filter = get_site_filter()
                args = ["--files", "25"]
                if site_filter:
                    args.extend(["--site", site_filter])
                run_script("populate_files.py", args)
            elif sub_choice == '3':
                site_filter = get_site_filter()
                args = ["--files", "50"]
                if site_filter:
                    args.extend(["--site", site_filter])
                run_script("populate_files.py", args)
            elif sub_choice == '4':
                site_filter = get_site_filter()
                args = ["--files", "100"]
                if site_filter:
                    args.extend(["--site", site_filter])
                run_script("populate_files.py", args)
            elif sub_choice == '5':
                print()
                count = input(f"  {Colors.YELLOW}Number of files (1-1000):{Colors.NC} ").strip()
                if count.isdigit() and 1 <= int(count) <= 1000:
                    site_filter = get_site_filter()
                    args = ["--files", count]
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
            # else: back to menu
            
        elif choice == '4':
            # List SharePoint Sites
            list_sharepoint_sites_menu()
            
        elif choice == '5':
            # List Files in Sites
            list_files_in_sites_menu()
            
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
