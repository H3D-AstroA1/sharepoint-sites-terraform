#!/usr/bin/env python3
"""
SharePoint Cleanup Script

This script provides options to delete files and/or SharePoint sites.
Use with caution - deletions are permanent!

Usage:
    python cleanup.py                           # Interactive mode
    python cleanup.py --delete-files            # Delete all files from sites
    python cleanup.py --delete-sites            # Delete SharePoint sites
    python cleanup.py --delete-all              # Delete both files and sites
    python cleanup.py --site hr --delete-files  # Delete files from HR sites only
    python cleanup.py --select-sites            # Interactively select specific sites
    python cleanup.py --select-files            # Interactively select specific files
    python cleanup.py --list-sites              # List available sites
    python cleanup.py --list-files              # List files in all sites
    python cleanup.py --list-files --site hr    # List files in a SPECIFIC site
    python cleanup.py --list-groups             # List Microsoft 365 Groups
    python cleanup.py --delete-groups           # Delete Microsoft 365 Groups
    python cleanup.py --list-deleted            # List deleted groups in recycle bin
    python cleanup.py --purge-deleted           # Permanently delete groups from recycle bin
    python cleanup.py --purge-spo-recycle       # Purge SharePoint site recycle bin
    python cleanup.py --purge-spo-recycle --tenant contoso  # With tenant name
    python cleanup.py --help                    # Show help

Requirements:
    - Python 3.8+
    - Azure CLI (logged in)
    - Microsoft Graph API permissions (Sites.ReadWrite.All, Files.ReadWrite.All)
    - SharePoint Online PowerShell module (for --purge-spo-recycle)

WARNING: This script performs DESTRUCTIVE operations. Always backup important data first!
"""

import argparse
import json
import os
import platform
import shutil
import subprocess
import sys
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple, Set

# ============================================================================
# CONFIGURATION
# ============================================================================

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_DIR = SCRIPT_DIR.parent
CONFIG_DIR = PROJECT_DIR / "config"
ENVIRONMENTS_FILE = CONFIG_DIR / "environments.json"
# App config file path (same as menu.py - in scripts folder as hidden file)
APP_CONFIG_FILE = SCRIPT_DIR / ".app_config.json"

# Default Azure CLI installation paths on Windows
AZURE_CLI_PATHS = [
    r"C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
    r"C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
]

def find_azure_cli_path() -> str:
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
    
    return "az"  # Return 'az' as fallback

# ============================================================================
# SYSTEM SITES DETECTION
# ============================================================================

# System sites that cannot be deleted (protected by SharePoint)
SYSTEM_SITE_PATTERNS = [
    "my workspace",
    "designer",
    "team site",
    "communication site",
    "contentstorage",  # Content storage sites (My workspace, Designer, etc.)
    "contenttypehub",  # Content Type Hub
    "appcatalog",      # App Catalog
    "search",          # Search Center
]

def is_system_site(site: Dict[str, Any]) -> bool:
    """Check if a site is a protected system site that cannot be deleted.
    
    System sites include:
    - My workspace (personal content storage)
    - Designer (design content storage)
    - Team Site (default team site)
    - Communication site (root site)
    - Content Type Hub
    - App Catalog
    - Search Center
    """
    name = site.get("displayName", site.get("name", "")).lower()
    web_url = site.get("webUrl", "").lower()
    
    # Check name patterns
    for pattern in SYSTEM_SITE_PATTERNS:
        if pattern in name:
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
    
    return False

def categorize_sites(sites: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
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

def print_banner(text: str) -> None:
    """Print a banner with the given text."""
    print()
    print(f"  {Colors.RED}{'=' * 60}{Colors.NC}")
    print(f"  {Colors.BOLD}{Colors.WHITE}{text.center(60)}{Colors.NC}")
    print(f"  {Colors.RED}{'=' * 60}{Colors.NC}")
    print()

def print_step(number: int, title: str) -> None:
    """Print a step header."""
    print()
    print(f"  {Colors.CYAN}[Step {number}]{Colors.NC} {Colors.BOLD}{title}{Colors.NC}")
    print(f"  {Colors.CYAN}{'-' * 50}{Colors.NC}")

def print_success(message: str) -> None:
    print(f"  {Colors.GREEN}✓{Colors.NC} {message}")

def print_error(message: str) -> None:
    print(f"  {Colors.RED}✗{Colors.NC} {message}")

def print_warning(message: str) -> None:
    print(f"  {Colors.YELLOW}⚠{Colors.NC} {message}")

def print_info(message: str) -> None:
    print(f"  {Colors.BLUE}ℹ{Colors.NC} {message}")

def print_danger(message: str) -> None:
    print(f"  {Colors.RED}{Colors.BOLD}⚠ DANGER:{Colors.NC} {Colors.RED}{message}{Colors.NC}")

def print_progress(current: int, total: int, message: str) -> None:
    """Print progress indicator."""
    percentage = (current / total) * 100 if total > 0 else 0
    bar_length = 30
    filled = int(bar_length * current / total) if total > 0 else 0
    bar = '█' * filled + '░' * (bar_length - filled)
    print(f"\r  {Colors.CYAN}[{bar}]{Colors.NC} {percentage:5.1f}% - {message:<40}", end='', flush=True)

# ============================================================================
# ENVIRONMENT CONFIGURATION
# ============================================================================

def load_environments() -> Dict:
    """Load pre-configured environments from environments.json."""
    if not ENVIRONMENTS_FILE.exists():
        return {"environments": []}
    
    try:
        with open(ENVIRONMENTS_FILE, 'r') as f:
            data = json.load(f)
            return data
    except Exception as e:
        print_warning(f"Could not load environments.json: {e}")
        return {"environments": []}

def get_environment_tenant(env: Dict) -> Optional[str]:
    """Get the tenant ID from an environment configuration."""
    azure_config = env.get("azure", {})
    return azure_config.get("tenant_id")

def select_environment() -> Optional[Dict]:
    """Auto-select or prompt for environment selection."""
    env_data = load_environments()
    environments = env_data.get("environments", [])
    
    if not environments:
        # No environments configured, use current Azure CLI login
        return None
    
    if len(environments) == 1:
        # Only one environment, use it automatically
        env = environments[0]
        print_info(f"Using environment: {env.get('name', 'Default')}")
        return env
    
    # Multiple environments, prompt user to select
    print()
    print(f"  {Colors.WHITE}Available Environments:{Colors.NC}")
    print()
    for i, env in enumerate(environments, 1):
        name = env.get("name", f"Environment {i}")
        tenant = get_environment_tenant(env) or "Not configured"
        print(f"    [{i}] {name}")
        print(f"        Tenant: {tenant[:20]}..." if len(tenant) > 20 else f"        Tenant: {tenant}")
    print()
    
    while True:
        try:
            choice = input(f"  {Colors.YELLOW}Select environment (1-{len(environments)}):{Colors.NC} ").strip()
            idx = int(choice) - 1
            if 0 <= idx < len(environments):
                env = environments[idx]
                print_success(f"Selected: {env.get('name', 'Environment')}")
                return env
        except ValueError:
            pass
        print_error("Invalid selection")

def switch_to_tenant(tenant_id: str) -> bool:
    """Switch Azure CLI to the specified tenant."""
    if not tenant_id:
        return True
    
    az_path = find_azure_cli_path()
    try:
        # Check if already logged into this tenant
        result = subprocess.run(
            [az_path, "account", "show", "--query", "tenantId", "-o", "tsv"],
            capture_output=True,
            text=True,
            check=True
        )
        current_tenant = result.stdout.strip()
        
        if current_tenant == tenant_id:
            print_info(f"Already logged into tenant: {tenant_id[:20]}...")
            return True
        
        # Need to switch tenant
        print_info(f"Switching to tenant: {tenant_id[:20]}...")
        subprocess.run(
            [az_path, "login", "--tenant", tenant_id],
            check=True
        )
        return True
        
    except subprocess.CalledProcessError:
        print_error(f"Failed to switch to tenant: {tenant_id}")
        return False

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def run_command(command: List[str], capture_output: bool = True) -> Optional[subprocess.CompletedProcess]:
    """Run a shell command and return the result."""
    # Resolve Azure CLI path if command starts with 'az'
    if command and command[0] == 'az':
        command = [find_azure_cli_path()] + command[1:]
    
    try:
        result = subprocess.run(
            command,
            capture_output=capture_output,
            text=True,
            check=True
        )
        return result
    except subprocess.CalledProcessError:
        return None

def check_azure_login() -> bool:
    """Check if user is logged into Azure CLI."""
    result = run_command(["az", "account", "show"])
    return result is not None

def azure_login() -> bool:
    """Perform Azure CLI login."""
    print_info("Opening browser for Azure login...")
    az_path = find_azure_cli_path()
    try:
        subprocess.run([az_path, "login"], check=True)
        return True
    except FileNotFoundError:
        print_error("Azure CLI is not installed or not in PATH.")
        print_info("Please install Azure CLI: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli")
        print_info("Or run the main menu (menu.py) and use option [0] to install prerequisites.")
        return False
    except subprocess.CalledProcessError:
        return False

def load_app_config() -> Optional[Dict[str, Any]]:
    """Load the custom app configuration from file."""
    if APP_CONFIG_FILE.exists():
        try:
            with open(APP_CONFIG_FILE, 'r') as f:
                return json.load(f)
        except Exception:
            pass
    return None


def get_graph_access_token_via_client_credentials(app_config: Dict[str, Any]) -> Optional[str]:
    """Get Microsoft Graph access token using client credentials flow."""
    tenant_id = app_config.get("tenant_id")
    client_id = app_config.get("app_id")
    client_secret = app_config.get("client_secret")
    
    if not all([tenant_id, client_id, client_secret]):
        return None
    
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    data = urllib.parse.urlencode({
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }).encode()
    
    try:
        req = urllib.request.Request(token_url, data=data)
        req.add_header("Content-Type", "application/x-www-form-urlencoded")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            result = json.loads(response.read().decode())
            return result.get("access_token")
    except Exception as e:
        print_error(f"Failed to get token via client credentials: {e}")
    return None


def get_access_token() -> Optional[str]:
    """Get Microsoft Graph access token.
    
    First tries to use custom app credentials if available,
    then falls back to Azure CLI.
    """
    # Try custom app credentials first
    app_config = load_app_config()
    if app_config and app_config.get("client_secret"):
        print_info("Using custom app credentials for authentication...")
        token = get_graph_access_token_via_client_credentials(app_config)
        if token:
            print_success("Authenticated via custom app")
            return token
        else:
            print_warning("Custom app authentication failed, falling back to Azure CLI...")
    
    # Fall back to Azure CLI
    try:
        result = run_command([
            "az", "account", "get-access-token",
            "--resource", "https://graph.microsoft.com",
            "--query", "accessToken",
            "-o", "tsv"
        ])
        if result:
            return result.stdout.strip()
    except Exception as e:
        print_error(f"Failed to get access token: {e}")
    return None

def parse_selection(selection: str, max_value: int) -> Set[int]:
    """Parse a selection string like '1,3,5-7' or '*' for all into a set of integers."""
    selected = set()
    
    # Handle wildcard for all items
    if selection.strip() == '*':
        return set(range(1, max_value + 1))
    
    parts = selection.replace(' ', '').split(',')
    
    for part in parts:
        if not part:
            continue
        if '-' in part:
            try:
                start, end = part.split('-', 1)
                start_num = int(start)
                end_num = int(end)
                if 1 <= start_num <= max_value and 1 <= end_num <= max_value:
                    for i in range(min(start_num, end_num), max(start_num, end_num) + 1):
                        selected.add(i)
            except ValueError:
                continue
        else:
            try:
                num = int(part)
                if 1 <= num <= max_value:
                    selected.add(num)
            except ValueError:
                continue
    
    return selected

# ============================================================================
# SHAREPOINT OPERATIONS
# ============================================================================

def get_sharepoint_sites(access_token: str) -> List[Dict[str, Any]]:
    """Get list of SharePoint sites using Microsoft Graph API."""
    sites = []
    url = "https://graph.microsoft.com/v1.0/sites?search=*"
    
    try:
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            sites = data.get("value", [])
    except urllib.error.HTTPError as e:
        print_error(f"Failed to get sites: {e.code} - {e.reason}")
        if e.code == 403:
            print()
            print(f"  {Colors.YELLOW}{'─' * 60}{Colors.NC}")
            print(f"  {Colors.YELLOW}PERMISSION ERROR - Admin Consent Required{Colors.NC}")
            print(f"  {Colors.YELLOW}{'─' * 60}{Colors.NC}")
            print()
            print(f"  The Azure CLI app needs admin consent for SharePoint permissions.")
            print()
            print(f"  {Colors.WHITE}Option 1: Grant admin consent (recommended){Colors.NC}")
            print(f"  A tenant admin must grant the Azure CLI app these permissions:")
            print(f"    • Sites.Read.All")
            print(f"    • Sites.ReadWrite.All")
            print(f"    • Files.ReadWrite.All")
            print()
            print(f"  {Colors.CYAN}Steps:{Colors.NC}")
            print(f"  1. Go to Azure Portal > Microsoft Entra ID > Enterprise Applications")
            print(f"  2. Search for 'Azure CLI' (App ID: 04b07795-8ddb-461a-bbee-02f9e1bf7b46)")
            print(f"  3. Go to Permissions > Grant admin consent")
            print()
            print(f"  {Colors.WHITE}Option 2: Use PowerShell with PnP{Colors.NC}")
            print(f"  Install PnP.PowerShell and use Connect-PnPOnline with interactive login")
            print()
            print(f"  {Colors.WHITE}Option 3: Re-login with correct scope{Colors.NC}")
            print(f"  Try: az login --scope https://graph.microsoft.com/.default")
            print()
    except Exception as e:
        print_error(f"Error getting sites: {e}")
    
    return sites


def get_m365_groups(access_token: str) -> List[Dict[str, Any]]:
    """Get list of Microsoft 365 Groups (Unified Groups) that have SharePoint sites."""
    groups = []
    # Filter for Unified groups (Microsoft 365 Groups) which have SharePoint sites
    # URL encode the filter parameter to avoid issues with special characters
    filter_param = urllib.parse.quote("groupTypes/any(c:c eq 'Unified')")
    select_param = "id,displayName,description,mail,createdDateTime,visibility"
    url = f"https://graph.microsoft.com/v1.0/groups?$filter={filter_param}&$select={select_param}"
    
    try:
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            groups = data.get("value", [])
            
            # Get the SharePoint site URL for each group
            for group in groups:
                try:
                    site_url = f"https://graph.microsoft.com/v1.0/groups/{group['id']}/sites/root"
                    site_req = urllib.request.Request(site_url)
                    site_req.add_header("Authorization", f"Bearer {access_token}")
                    site_req.add_header("Content-Type", "application/json")
                    
                    with urllib.request.urlopen(site_req, timeout=30) as site_response:
                        site_data = json.loads(site_response.read().decode())
                        group["siteUrl"] = site_data.get("webUrl", "")
                        group["siteId"] = site_data.get("id", "")
                except Exception:
                    group["siteUrl"] = ""
                    group["siteId"] = ""
                    
    except urllib.error.HTTPError as e:
        print_error(f"Failed to get groups: {e.code} - {e.reason}")
        if e.code == 403:
            print_info("You may need Group.Read.All permission to list groups")
    except Exception as e:
        print_error(f"Error getting groups: {e}")
    
    return groups


def delete_m365_group(group_id: str, access_token: str) -> bool:
    """Delete a Microsoft 365 Group (which also deletes its SharePoint site)."""
    url = f"https://graph.microsoft.com/v1.0/groups/{group_id}"
    
    try:
        req = urllib.request.Request(url, method="DELETE")
        req.add_header("Authorization", f"Bearer {access_token}")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            return True
    except urllib.error.HTTPError as e:
        if e.code == 204:
            return True  # 204 No Content is success for DELETE
        print_error(f"Failed to delete group: {e.code} - {e.reason}")
        if e.code == 403:
            print_info("You may need Group.ReadWrite.All permission to delete groups")
        return False
    except Exception as e:
        print_error(f"Error deleting group: {e}")
        return False


def get_deleted_m365_groups(access_token: str) -> List[Dict[str, Any]]:
    """Get list of deleted Microsoft 365 Groups from the recycle bin."""
    groups = []
    
    # Use the directory/deletedItems endpoint to get deleted groups
    filter_param = urllib.parse.quote("groupTypes/any(c:c eq 'Unified')")
    select_param = "id,displayName,description,mail,deletedDateTime"
    url = f"https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group?$filter={filter_param}&$select={select_param}"
    
    try:
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            groups = data.get("value", [])
            
            # Handle pagination
            while "@odata.nextLink" in data:
                next_url = data["@odata.nextLink"]
                req = urllib.request.Request(next_url)
                req.add_header("Authorization", f"Bearer {access_token}")
                req.add_header("Content-Type", "application/json")
                with urllib.request.urlopen(req, timeout=30) as response:
                    data = json.loads(response.read().decode())
                    groups.extend(data.get("value", []))
                    
    except urllib.error.HTTPError as e:
        print_error(f"Failed to get deleted groups: {e.code} - {e.reason}")
        if e.code == 403:
            print_info("You may need Directory.Read.All permission to list deleted groups")
    except Exception as e:
        print_error(f"Error getting deleted groups: {e}")
    
    return groups


def permanently_delete_m365_group(group_id: str, access_token: str) -> bool:
    """Permanently delete a Microsoft 365 Group from the recycle bin."""
    url = f"https://graph.microsoft.com/v1.0/directory/deletedItems/{group_id}"
    
    try:
        req = urllib.request.Request(url, method="DELETE")
        req.add_header("Authorization", f"Bearer {access_token}")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            return True
    except urllib.error.HTTPError as e:
        if e.code == 204:
            return True  # 204 No Content is success for DELETE
        print_error(f"Failed to permanently delete group: {e.code} - {e.reason}")
        if e.code == 403:
            print_info("You may need Directory.ReadWrite.All permission to permanently delete groups")
        return False
    except Exception as e:
        print_error(f"Error permanently deleting group: {e}")
        return False


def display_deleted_groups_for_selection(groups: List[Dict[str, Any]]) -> None:
    """Display deleted groups in a numbered list for selection."""
    print()
    print(f"  {Colors.WHITE}{Colors.BOLD}Deleted Microsoft 365 Groups (Recycle Bin):{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 70}{Colors.NC}")
    print()
    
    for i, group in enumerate(groups, 1):
        name = group.get("displayName", "Unknown")
        deleted_date = group.get("deletedDateTime", "Unknown")
        if deleted_date and deleted_date != "Unknown":
            # Parse and format the date
            try:
                dt = datetime.fromisoformat(deleted_date.replace('Z', '+00:00'))
                deleted_date = dt.strftime("%Y-%m-%d %H:%M")
            except:
                pass
        
        print(f"  [{Colors.YELLOW}{i:2d}{Colors.NC}] 🗑️  {Colors.WHITE}{name}{Colors.NC}")
        print(f"       Deleted: {deleted_date}")
        print()


def interactive_select_deleted_groups(groups: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Allow user to select deleted groups from a numbered list."""
    display_deleted_groups_for_selection(groups)
    
    print(f"  Enter group numbers to select (e.g., 1,3,5 or 1-5 or * for all):")
    selection = input(f"  > ").strip()
    
    if not selection:
        return []
    
    selected_indices = parse_selection(selection, len(groups))
    
    if not selected_indices:
        print_warning("No valid selection made")
        return []
    
    selected_groups = [groups[i-1] for i in sorted(selected_indices) if 1 <= i <= len(groups)]
    print_info(f"Selected {len(selected_groups)} group(s)")
    
    return selected_groups


def purge_deleted_groups_mode(groups: List[Dict[str, Any]], access_token: str, auto_confirm: bool = False) -> None:
    """Permanently delete selected groups from the recycle bin."""
    if not groups:
        print_warning("No deleted groups to purge")
        return
    
    selected_groups = interactive_select_deleted_groups(groups)
    
    if not selected_groups:
        print_info("No groups selected for permanent deletion")
        return
    
    # Show confirmation
    print()
    print_danger("WARNING: Permanently deleting groups cannot be undone!")
    print_danger("This will free up the site URLs for reuse.")
    print()
    print(f"  {Colors.WHITE}Groups to be permanently deleted:{Colors.NC}")
    for group in selected_groups:
        print(f"    {Colors.RED}✗{Colors.NC} {group.get('displayName', 'Unknown')}")
    print()
    
    if not auto_confirm:
        confirm = input(f"  Type 'PURGE' to confirm permanent deletion: ").strip()
        if confirm != "PURGE":
            print_info("Permanent deletion cancelled")
            return
    
    # Delete the groups
    print()
    total = len(selected_groups)
    success_count = 0
    
    for i, group in enumerate(selected_groups, 1):
        name = group.get("displayName", "Unknown")
        group_id = group.get("id", "")
        
        print_progress(i, total, f"Purging {name}...")
        
        if permanently_delete_m365_group(group_id, access_token):
            print_success(f"Permanently deleted: {name}")
            success_count += 1
        else:
            print_error(f"Failed to permanently delete: {name}")
    
    print()
    print_info(f"Permanently deleted {success_count} of {total} groups")
    if success_count > 0:
        print_info("The site URLs are now available for reuse")


# ============================================================================
# SHAREPOINT ONLINE POWERSHELL FUNCTIONS (for SharePoint site recycle bin)
# ============================================================================

def get_powershell_executable() -> str:
    """Get the appropriate PowerShell executable (prefer Windows PowerShell for SPO module)."""
    # Windows PowerShell (powershell.exe) is more compatible with SPO module
    # PowerShell 7 (pwsh) has issues with PowerShellGet
    if platform.system() == "Windows":
        # Check for Windows PowerShell first
        win_ps = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        if os.path.exists(win_ps):
            return win_ps
    return "powershell"


def check_spo_module_installed() -> bool:
    """Check if SharePoint Online PowerShell module is installed."""
    ps_exe = get_powershell_executable()
    try:
        # Check in all possible module paths
        check_script = '''
$ErrorActionPreference = "SilentlyContinue"
$found = $false

# Check standard module paths
$modulePaths = @(
    "$env:USERPROFILE\\Documents\\WindowsPowerShell\\Modules",
    "$env:ProgramFiles\\WindowsPowerShell\\Modules",
    "C:\\Program Files\\WindowsPowerShell\\Modules",
    "$env:USERPROFILE\\Documents\\PowerShell\\Modules"
)

foreach ($path in $modulePaths) {
    if (Test-Path "$path\\Microsoft.Online.SharePoint.PowerShell") {
        $found = $true
        break
    }
}

# Also check using Get-Module
if (-not $found) {
    $module = Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell
    if ($module) {
        $found = $true
    }
}

if ($found) {
    Write-Output "FOUND"
} else {
    Write-Output "NOT_FOUND"
}
'''
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-Command", check_script],
            capture_output=True,
            text=True,
            timeout=30
        )
        return "FOUND" in result.stdout
    except Exception:
        return False


def install_spo_module() -> bool:
    """Install SharePoint Online PowerShell module."""
    ps_exe = get_powershell_executable()
    print_info("Installing SharePoint Online PowerShell module...")
    print_info(f"Using: {ps_exe}")
    
    # First, try to repair/update PowerShellGet and PackageManagement
    print_info("Updating PowerShellGet module...")
    try:
        repair_script = '''
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "SilentlyContinue"

# Try to update NuGet provider
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null

# Try to update PowerShellGet
Install-Module -Name PowerShellGet -Force -AllowClobber -Scope CurrentUser | Out-Null

# Now install SPO module
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber -Scope CurrentUser

if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell) {
    Write-Output "SUCCESS"
} else {
    Write-Error "Module not found after installation"
    exit 1
}
'''
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-Command", repair_script],
            capture_output=True,
            text=True,
            timeout=600  # 10 minutes for full installation
        )
        
        if "SUCCESS" in result.stdout:
            print_success("SharePoint Online PowerShell module installed")
            return True
        else:
            # Try alternative method using direct download
            print_warning("Standard installation failed, trying alternative method...")
            return install_spo_module_alternative(ps_exe)
            
    except subprocess.TimeoutExpired:
        print_error("Installation timed out")
        return False
    except Exception as e:
        print_error(f"Error installing module: {e}")
        return False


def install_spo_module_alternative(ps_exe: str) -> bool:
    """Alternative method to install SPO module using direct download."""
    print_info("Attempting alternative installation method...")
    
    alt_script = '''
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"

try {
    # Register PSGallery if not registered
    if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
        Register-PSRepository -Default -ErrorAction SilentlyContinue
    }
    
    # Set PSGallery as trusted
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    
    # Install using Save-Module and manual copy
    $modulePath = "$env:USERPROFILE\\Documents\\WindowsPowerShell\\Modules"
    if (-not (Test-Path $modulePath)) {
        New-Item -ItemType Directory -Path $modulePath -Force | Out-Null
    }
    
    # Try direct installation one more time with verbose output
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber -Scope CurrentUser -Verbose
    
    if (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell) {
        Write-Output "SUCCESS"
    } else {
        throw "Module still not available"
    }
} catch {
    Write-Error $_.Exception.Message
    exit 1
}
'''
    
    try:
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-Command", alt_script],
            capture_output=True,
            text=True,
            timeout=600
        )
        
        if "SUCCESS" in result.stdout:
            print_success("SharePoint Online PowerShell module installed (alternative method)")
            return True
        else:
            print_error(f"Alternative installation also failed: {result.stderr}")
            return False
    except Exception as e:
        print_error(f"Error with alternative installation: {e}")
        return False


def get_spo_deleted_sites(admin_url: str) -> List[Dict[str, Any]]:
    """Get deleted SharePoint sites using PowerShell."""
    ps_exe = get_powershell_executable()
    sites = []
    
    # PowerShell script to get deleted sites
    ps_script = f'''
$ErrorActionPreference = "Stop"
try {{
    # Connect to SharePoint Online (will prompt for credentials if needed)
    Connect-SPOService -Url "{admin_url}" -ErrorAction Stop
    
    # Get deleted sites
    $deletedSites = Get-SPODeletedSite -ErrorAction Stop
    
    # Output as JSON
    $deletedSites | Select-Object Url, DeletionTime, SiteId, Status | ConvertTo-Json -Compress
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}
'''
    
    try:
        print_info("Connecting to SharePoint Online...")
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode == 0 and result.stdout.strip():
            output = result.stdout.strip()
            # Handle single item vs array
            if output.startswith('['):
                sites = json.loads(output)
            elif output.startswith('{'):
                sites = [json.loads(output)]
        elif result.stderr:
            print_error(f"PowerShell error: {result.stderr}")
            
    except subprocess.TimeoutExpired:
        print_error("Connection timed out. Please try again.")
    except json.JSONDecodeError:
        print_warning("Could not parse deleted sites response")
    except Exception as e:
        print_error(f"Error getting deleted sites: {e}")
    
    return sites


def purge_spo_deleted_site(admin_url: str, site_url: str) -> bool:
    """Permanently delete a SharePoint site from the recycle bin (single site)."""
    ps_exe = get_powershell_executable()
    ps_script = f'''
$ErrorActionPreference = "Stop"
try {{
    Connect-SPOService -Url "{admin_url}" -ErrorAction SilentlyContinue
    Remove-SPODeletedSite -Identity "{site_url}" -Confirm:$false -ErrorAction Stop
    Write-Output "SUCCESS"
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}
'''
    
    try:
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=120
        )
        
        return result.returncode == 0 and "SUCCESS" in result.stdout
    except Exception as e:
        print_error(f"Error purging site: {e}")
        return False


def purge_spo_deleted_sites_batch(admin_url: str, site_urls: List[str]) -> Dict[str, bool]:
    """Permanently delete multiple SharePoint sites in a single session (no re-auth)."""
    ps_exe = get_powershell_executable()
    
    # Build the list of sites to delete
    sites_array = ", ".join([f'"{url}"' for url in site_urls])
    
    ps_script = f'''
$ErrorActionPreference = "Stop"
$results = @{{}}
$sites = @({sites_array})

try {{
    # Connect once at the start
    Write-Host "Connecting to SharePoint Online..."
    Connect-SPOService -Url "{admin_url}" -ErrorAction Stop
    Write-Host "Connected successfully!"
    
    foreach ($site in $sites) {{
        try {{
            Write-Host "Deleting: $site"
            Remove-SPODeletedSite -Identity $site -Confirm:$false -ErrorAction Stop
            $results[$site] = "SUCCESS"
            Write-Host "SUCCESS: $site"
        }} catch {{
            $results[$site] = "FAILED: $($_.Exception.Message)"
            Write-Host "FAILED: $site - $($_.Exception.Message)"
        }}
    }}
    
    # Output results as JSON
    $results | ConvertTo-Json -Compress
}} catch {{
    Write-Error "Connection failed: $($_.Exception.Message)"
    exit 1
}}
'''
    
    try:
        print_info("Connecting to SharePoint Online (authenticate once)...")
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=600  # 10 minutes for batch operation
        )
        
        if result.returncode == 0:
            # Parse results from output
            results = {}
            for url in site_urls:
                if f"SUCCESS: {url}" in result.stdout or f'"{url}":"SUCCESS"' in result.stdout:
                    results[url] = True
                else:
                    results[url] = False
            return results
        else:
            print_error(f"Batch deletion failed: {result.stderr}")
            return {url: False for url in site_urls}
            
    except subprocess.TimeoutExpired:
        print_error("Batch operation timed out")
        return {url: False for url in site_urls}
    except Exception as e:
        print_error(f"Error in batch purge: {e}")
        return {url: False for url in site_urls}


def display_spo_deleted_sites_for_selection(sites: List[Dict[str, Any]]) -> None:
    """Display deleted SharePoint sites in a numbered list for selection."""
    print()
    print(f"  {Colors.WHITE}{Colors.BOLD}Deleted SharePoint Sites (SharePoint Recycle Bin):{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 70}{Colors.NC}")
    print()
    
    for i, site in enumerate(sites, 1):
        url = site.get("Url", "Unknown")
        deleted_time = site.get("DeletionTime", "Unknown")
        
        # Extract site name from URL
        site_name = url.split("/sites/")[-1] if "/sites/" in url else url
        
        print(f"  [{Colors.YELLOW}{i:2d}{Colors.NC}] 🗑️  {Colors.WHITE}{site_name}{Colors.NC}")
        print(f"       {Colors.CYAN}{url}{Colors.NC}")
        print(f"       Deleted: {deleted_time}")
        print()


def interactive_select_spo_deleted_sites(sites: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Allow user to select deleted SharePoint sites from a numbered list."""
    display_spo_deleted_sites_for_selection(sites)
    
    print(f"  Enter site numbers to select (e.g., 1,3,5 or 1-5 or * for all):")
    selection = input(f"  > ").strip()
    
    if not selection:
        return []
    
    selected_indices = parse_selection(selection, len(sites))
    
    if not selected_indices:
        print_warning("No valid selection made")
        return []
    
    selected_sites = [sites[i-1] for i in sorted(selected_indices) if 1 <= i <= len(sites)]
    print_info(f"Selected {len(selected_sites)} site(s)")
    
    return selected_sites


def purge_spo_deleted_sites_mode(admin_url: str, auto_confirm: bool = False) -> None:
    """List and permanently delete SharePoint sites from the recycle bin."""
    
    # Check if SPO module is installed
    if not check_spo_module_installed():
        print_warning("SharePoint Online PowerShell module is not installed")
        install_choice = input("  Would you like to install it now? (y/n): ").strip().lower()
        if install_choice == 'y':
            if not install_spo_module():
                print_error("Could not install SharePoint Online PowerShell module")
                print_info("You can install it manually with:")
                print_info("  Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force")
                return
        else:
            print_info("Skipping SharePoint site recycle bin cleanup")
            return
    
    # Get deleted sites
    print_info("Fetching deleted SharePoint sites...")
    sites = get_spo_deleted_sites(admin_url)
    
    if not sites:
        print_warning("No deleted SharePoint sites found in recycle bin")
        print_info("The SharePoint recycle bin is empty")
        return
    
    print_success(f"Found {len(sites)} deleted SharePoint site(s)")
    
    # Let user select sites
    selected_sites = interactive_select_spo_deleted_sites(sites)
    
    if not selected_sites:
        print_info("No sites selected for permanent deletion")
        return
    
    # Show confirmation
    print()
    print_danger("WARNING: Permanently deleting sites cannot be undone!")
    print_danger("This will free up the site URLs for reuse.")
    print()
    print(f"  {Colors.WHITE}Sites to be permanently deleted:{Colors.NC}")
    for site in selected_sites:
        url = site.get("Url", "Unknown")
        print(f"    {Colors.RED}✗{Colors.NC} {url}")
    print()
    
    if not auto_confirm:
        confirm = input(f"  Type 'PURGE' to confirm permanent deletion: ").strip()
        if confirm != "PURGE":
            print_info("Permanent deletion cancelled")
            return
    
    # Delete the sites using batch operation (single authentication)
    print()
    site_urls = [site.get("Url", "") for site in selected_sites if site.get("Url")]
    
    print_info(f"Deleting {len(site_urls)} sites in a single session...")
    print_info("You will only need to authenticate once.")
    print()
    
    results = purge_spo_deleted_sites_batch(admin_url, site_urls)
    
    # Show results
    success_count = 0
    for url, success in results.items():
        site_name = url.split("/sites/")[-1] if "/sites/" in url else url
        if success:
            print_success(f"Permanently deleted: {site_name}")
            success_count += 1
        else:
            print_error(f"Failed to permanently delete: {site_name}")
    
    print()
    print_info(f"Permanently deleted {success_count} of {len(site_urls)} sites")
    if success_count > 0:
        print_info("The site URLs are now available for reuse")


def display_groups_for_selection(groups: List[Dict[str, Any]]) -> None:
    """Display groups in a numbered list for selection."""
    print()
    print(f"  {Colors.WHITE}{Colors.BOLD}Microsoft 365 Groups with SharePoint Sites:{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 70}{Colors.NC}")
    print()
    
    for i, group in enumerate(groups, 1):
        name = group.get("displayName", "Unknown")
        site_url = group.get("siteUrl", "No site")
        visibility = group.get("visibility", "Unknown")
        created = group.get("createdDateTime", "")[:10] if group.get("createdDateTime") else ""
        
        visibility_icon = "🔒" if visibility == "Private" else "🌐"
        
        print(f"  [{i:2d}] {visibility_icon} {Colors.WHITE}{name}{Colors.NC}")
        if site_url:
            print(f"       {Colors.CYAN}{site_url}{Colors.NC}")
        if created:
            print(f"       {Colors.YELLOW}Created: {created}{Colors.NC}")
        print()


def interactive_select_groups(groups: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Allow user to interactively select groups from a numbered list."""
    if not groups:
        return []
    
    display_groups_for_selection(groups)
    
    print(f"  {Colors.WHITE}Enter group numbers to select (e.g., 1,3,5 or 1-5 or * for all):{Colors.NC}")
    selection = input("  > ").strip()
    
    if not selection:
        return []
    
    selected_indices = parse_selection(selection, len(groups))
    
    if not selected_indices:
        print_warning("No valid selection made")
        return []
    
    selected_groups = [groups[i-1] for i in sorted(selected_indices) if 0 < i <= len(groups)]
    
    print()
    print_info(f"Selected {len(selected_groups)} group(s)")
    
    return selected_groups


def delete_groups_mode(groups: List[Dict[str, Any]], access_token: str, auto_confirm: bool = False,
                       tenant: Optional[str] = None, skip_recycle_purge: bool = False) -> None:
    """Delete selected Microsoft 365 Groups (and their SharePoint sites).
    
    Args:
        groups: List of M365 groups to delete
        access_token: Microsoft Graph access token
        auto_confirm: Skip confirmation prompts
        tenant: SharePoint tenant name for SPO recycle bin purge (e.g., 'contoso')
        skip_recycle_purge: Skip automatic recycle bin purge after deletion
    """
    if not groups:
        print_warning("No groups to delete")
        return
    
    print()
    print_danger("WARNING: Deleting groups will also delete their SharePoint sites!")
    print_danger("This action cannot be undone!")
    print()
    
    print(f"  {Colors.WHITE}Groups to be deleted:{Colors.NC}")
    for group in groups:
        print(f"    {Colors.RED}✗{Colors.NC} {group.get('displayName', 'Unknown')}")
        if group.get("siteUrl"):
            print(f"      {Colors.CYAN}{group.get('siteUrl')}{Colors.NC}")
    print()
    
    if not auto_confirm:
        confirm = input(f"  {Colors.RED}Type 'DELETE' to confirm deletion: {Colors.NC}").strip()
        if confirm != "DELETE":
            print_warning("Deletion cancelled")
            return
    
    print()
    deleted = 0
    failed = 0
    
    for group in groups:
        name = group.get("displayName", "Unknown")
        group_id = group.get("id")
        
        if not group_id:
            print_error(f"No ID for group: {name}")
            failed += 1
            continue
        
        print_progress(deleted + failed + 1, len(groups), f"Deleting {name}...")
        
        if delete_m365_group(group_id, access_token):
            print_success(f"Deleted: {name}")
            deleted += 1
        else:
            print_error(f"Failed to delete: {name}")
            failed += 1
    
    print()
    print_info(f"Deleted: {deleted}, Failed: {failed}")
    
    # Automatically purge recycle bins after deletion
    if deleted > 0 and not skip_recycle_purge:
        print()
        print_banner("RECYCLE BIN CLEANUP")
        print()
        print_info("Sites have been soft-deleted. They now exist in two recycle bins:")
        print(f"    {Colors.CYAN}1.{Colors.NC} Microsoft 365 Groups recycle bin (Azure AD)")
        print(f"    {Colors.CYAN}2.{Colors.NC} SharePoint site recycle bin (SharePoint Admin Center)")
        print()
        
        # Try to auto-detect tenant name from site URLs if not provided
        if not tenant:
            for group in groups:
                site_url = group.get("siteUrl", "")
                if site_url and ".sharepoint.com" in site_url:
                    # Extract tenant from URL like https://contoso.sharepoint.com/sites/...
                    import re
                    match = re.search(r'https://([^.]+)\.sharepoint\.com', site_url)
                    if match:
                        tenant = match.group(1)
                        print_info(f"Auto-detected tenant name: {tenant}")
                        break
        
        # Ask if user wants to skip recycle bin purge
        if not auto_confirm:
            skip_choice = input(f"  {Colors.YELLOW}Purge recycle bins now? (Y/n): {Colors.NC}").strip().lower()
            if skip_choice == 'n':
                print_warning("Skipping recycle bin purge. Sites remain in recycle bins.")
                print_info("You can purge them later using menu options [6] and [7]")
                return
        
        # Step 1: Purge M365 Groups recycle bin
        print()
        print_step(1, "Purging M365 Groups recycle bin (Azure AD)")
        
        # Wait a moment for Azure AD to process the deletions
        print_info("Waiting for Azure AD to process deletions...")
        import time
        time.sleep(3)
        
        deleted_groups = get_deleted_m365_groups(access_token)
        if deleted_groups:
            print_info(f"Found {len(deleted_groups)} deleted groups in recycle bin")
            purge_deleted_groups_mode(deleted_groups, access_token, auto_confirm=True)
        else:
            print_info("No deleted groups found in Azure AD recycle bin")
        
        # Step 2: Purge SharePoint site recycle bin
        print()
        print_step(2, "Purging SharePoint site recycle bin")
        
        # Get tenant name if not provided
        if not tenant:
            print()
            print(f"  {Colors.WHITE}Enter your SharePoint tenant name{Colors.NC}")
            print(f"  {Colors.DIM}(e.g., 'contoso' for contoso.sharepoint.com){Colors.NC}")
            print()
            tenant = input(f"  {Colors.YELLOW}Tenant name (or press Enter to skip): {Colors.NC}").strip()
        
        if tenant:
            admin_url = f"https://{tenant}-admin.sharepoint.com"
            purge_spo_deleted_sites_mode(admin_url, auto_confirm=True)
        else:
            print_warning("Skipping SharePoint recycle bin purge (no tenant provided)")
            print_info("You can purge it later using menu option [7]")

def get_site_files(site_id: str, access_token: str) -> List[Dict[str, Any]]:
    """Get all files from a SharePoint site's document library."""
    files = []
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
    
    try:
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            items = data.get("value", [])
            
            for item in items:
                if "file" in item:
                    files.append(item)
                elif "folder" in item:
                    # Recursively get files from folders
                    folder_files = get_folder_files(site_id, item["id"], access_token)
                    files.extend(folder_files)
    except urllib.error.HTTPError as e:
        if e.code != 404:
            print_error(f"Failed to get files: {e.code}")
    except Exception as e:
        print_error(f"Error getting files: {e}")
    
    return files

def get_folder_files(site_id: str, folder_id: str, access_token: str) -> List[Dict[str, Any]]:
    """Get all files from a folder recursively."""
    files = []
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{folder_id}/children"
    
    try:
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            items = data.get("value", [])
            
            for item in items:
                if "file" in item:
                    files.append(item)
                elif "folder" in item:
                    # Recursively get files from subfolders
                    subfolder_files = get_folder_files(site_id, item["id"], access_token)
                    files.extend(subfolder_files)
    except Exception:
        pass
    
    return files

def get_all_site_items(site_id: str, access_token: str, include_folders: bool = True) -> List[Dict[str, Any]]:
    """Get all items (files and optionally folders) from a site with full path info."""
    items = []
    
    def get_items_recursive(folder_id: Optional[str] = None, path: str = "") -> None:
        if folder_id:
            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{folder_id}/children"
        else:
            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
        
        try:
            req = urllib.request.Request(url)
            req.add_header("Authorization", f"Bearer {access_token}")
            req.add_header("Content-Type", "application/json")
            
            with urllib.request.urlopen(req, timeout=30) as response:
                data = json.loads(response.read().decode())
                for item in data.get("value", []):
                    item_name = item.get("name", "")
                    item_path = f"{path}/{item_name}" if path else item_name
                    item["_full_path"] = item_path
                    
                    if "folder" in item:
                        if include_folders:
                            items.append(item)
                        get_items_recursive(item["id"], item_path)
                    else:
                        items.append(item)
        except Exception:
            pass
    
    get_items_recursive()
    return items

def delete_file(site_id: str, item_id: str, access_token: str) -> bool:
    """Delete a file from SharePoint."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
    
    try:
        req = urllib.request.Request(url, method="DELETE")
        req.add_header("Authorization", f"Bearer {access_token}")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            return response.status in [200, 204]
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return True  # Already deleted
        return False
    except Exception:
        return False

def delete_folder(site_id: str, folder_id: str, access_token: str) -> bool:
    """Delete a folder from SharePoint."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{folder_id}"
    
    try:
        req = urllib.request.Request(url, method="DELETE")
        req.add_header("Authorization", f"Bearer {access_token}")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            return response.status in [200, 204]
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return True  # Already deleted
        return False
    except Exception:
        return False

def get_site_root_items(site_id: str, access_token: str) -> List[Dict[str, Any]]:
    """Get all items (files and folders) from site root."""
    items = []
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
    
    try:
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            items = data.get("value", [])
    except Exception:
        pass
    
    return items


def check_pnp_module_installed() -> bool:
    """Check if PnP PowerShell module is installed.
    
    Prefers PowerShell 7 (pwsh) since PnP.PowerShell is typically installed there.
    Falls back to Windows PowerShell if pwsh is not available.
    """
    # Prefer PowerShell 7 (pwsh) since that's where PnP is typically installed
    pwsh = get_pwsh_executable()
    ps_exe = pwsh if pwsh else get_powershell_executable()
    
    ps_script = '''
$module = Get-Module -ListAvailable -Name "PnP.PowerShell" | Select-Object -First 1
if ($module) {
    Write-Output "INSTALLED"
} else {
    Write-Output "NOT_INSTALLED"
}
'''
    
    try:
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        # Check for exact match to avoid false positives
        output = result.stdout.strip()
        is_installed = output == "INSTALLED"
        
        # Debug output
        if not is_installed:
            ps_type = "pwsh" if pwsh else "powershell"
            print_info(f"  PnP module check ({ps_type}): {output}")
        
        return is_installed
    except Exception as e:
        print_warning(f"  Could not check PnP module: {e}")
        return False


def get_pwsh_executable() -> Optional[str]:
    """Get PowerShell 7 (pwsh) executable path if available."""
    # Check common locations for pwsh
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


def install_pnp_module() -> bool:
    """Install PnP PowerShell module."""
    # Prefer PowerShell 7 (pwsh) for better module management
    pwsh = get_pwsh_executable()
    ps_exe = pwsh if pwsh else get_powershell_executable()
    
    if pwsh:
        print_info("Installing PnP.PowerShell module using PowerShell 7...")
    else:
        print_info("Installing PnP.PowerShell module using Windows PowerShell...")
    print_info("This may take a few minutes...")
    
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

Write-Host "Installing PnP.PowerShell module..."

# First, ensure NuGet provider is installed (required for Install-Module)
Write-Host "Ensuring NuGet provider is available..."
try {
    $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [Version]"2.8.5.201" }
    if (-not $nuget) {
        Write-Host "Installing NuGet provider..."
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
    }
} catch {
    Write-Host "NuGet provider setup: $($_.Exception.Message)"
}

# Set PSGallery as trusted to avoid prompts
try {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
} catch {
    # Ignore errors
}

# Method 1: Standard Install-Module
Write-Host "Trying standard Install-Module..."
try {
    Install-Module -Name "PnP.PowerShell" -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -AcceptLicense -ErrorAction Stop
    Write-Output "SUCCESS:Installed"
    exit 0
} catch {
    Write-Host "Method 1 failed: $($_.Exception.Message)"
}

# Method 2: Try with SkipPublisherCheck
Write-Host "Trying Install-Module with SkipPublisherCheck..."
try {
    Install-Module -Name "PnP.PowerShell" -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
    Write-Output "SUCCESS:Installed"
    exit 0
} catch {
    Write-Host "Method 2 failed: $($_.Exception.Message)"
}

# Method 3: Try updating PowerShellGet first
Write-Host "Trying to update PowerShellGet..."
try {
    # Import PackageManagement first
    Import-Module PackageManagement -Force -ErrorAction SilentlyContinue
    
    # Try to install/update PowerShellGet
    Install-Module -Name PowerShellGet -Force -AllowClobber -Scope CurrentUser -ErrorAction SilentlyContinue
    
    # Now try PnP again
    Install-Module -Name "PnP.PowerShell" -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    Write-Output "SUCCESS:Installed after PowerShellGet update"
    exit 0
} catch {
    Write-Host "Method 3 failed: $($_.Exception.Message)"
}

# Method 4: Direct download from PSGallery using Save-Module
Write-Host "Trying Save-Module approach..."
try {
    $modulePath = Join-Path $env:USERPROFILE "Documents\\WindowsPowerShell\\Modules\\PnP.PowerShell"
    if (-not (Test-Path $modulePath)) {
        New-Item -ItemType Directory -Path $modulePath -Force | Out-Null
    }
    Save-Module -Name "PnP.PowerShell" -Path (Split-Path $modulePath -Parent) -Force -ErrorAction Stop
    Write-Output "SUCCESS:Installed via Save-Module"
    exit 0
} catch {
    Write-Host "Method 4 failed: $($_.Exception.Message)"
}

# Method 5: Use PowerShell 7's PSResourceGet if available
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Write-Host "Trying PSResourceGet (PowerShell 7)..."
    try {
        # PSResourceGet is built into PS7
        Install-PSResource -Name "PnP.PowerShell" -Scope CurrentUser -TrustRepository -ErrorAction Stop
        Write-Output "SUCCESS:Installed via PSResourceGet"
        exit 0
    } catch {
        Write-Host "PSResourceGet method failed: $($_.Exception.Message)"
    }
}

# Method 6: Direct web download as last resort
Write-Host "Trying direct download from PSGallery..."
try {
    $apiUrl = "https://www.powershellgallery.com/api/v2/package/PnP.PowerShell"
    $tempZip = Join-Path $env:TEMP "PnP.PowerShell.zip"
    $modulePath = Join-Path $env:USERPROFILE "Documents\\WindowsPowerShell\\Modules\\PnP.PowerShell"
    
    # Download the nupkg (it's a zip file)
    Invoke-WebRequest -Uri $apiUrl -OutFile $tempZip -UseBasicParsing -ErrorAction Stop
    
    # Extract to module path
    if (Test-Path $modulePath) {
        Remove-Item $modulePath -Recurse -Force
    }
    Expand-Archive -Path $tempZip -DestinationPath $modulePath -Force
    
    # Clean up temp file
    Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
    
    # Verify installation
    $installed = Get-Module -ListAvailable -Name "PnP.PowerShell" | Select-Object -First 1
    if ($installed) {
        Write-Output "SUCCESS:Installed via direct download"
        exit 0
    }
} catch {
    Write-Host "Direct download failed: $($_.Exception.Message)"
}

Write-Host "FAILED:All installation methods failed"
exit 1
'''
    
    try:
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=600  # 10 minutes for installation
        )
        
        # Check if SUCCESS is in output
        if "SUCCESS:" in result.stdout:
            print_success("PnP.PowerShell module installed successfully")
            return True
        else:
            # Show the actual error
            all_output = result.stdout + "\n" + result.stderr
            print_warning("Automatic installation failed")
            print()
            
            # Show what was tried
            if result.stdout:
                for line in result.stdout.split('\n'):
                    if line.strip() and not line.startswith('SUCCESS'):
                        print(f"    {line.strip()}")
            
            # Check if it's a permissions issue
            if "administrator" in all_output.lower() or "access denied" in all_output.lower():
                print()
                print_info("This may require administrator privileges.")
            
            # Provide manual installation instructions
            print()
            print_info("Please install PnP.PowerShell manually:")
            print_info("  Option 1 - Using PowerShell 7 (recommended):")
            print_info("    1. Install PowerShell 7: winget install Microsoft.PowerShell")
            print_info("    2. Open PowerShell 7 (pwsh)")
            print_info("    3. Run: Install-Module -Name PnP.PowerShell -Force")
            print()
            print_info("  Option 2 - Using Windows PowerShell as Admin:")
            print_info("    1. Open PowerShell as Administrator")
            print_info("    2. Run: Install-Module -Name PnP.PowerShell -Force")
            print()
            print_info("  Option 3 - Using winget:")
            print_info("    winget install --id=PnP.PowerShell -e")
            return False
    except subprocess.TimeoutExpired:
        print_error("Installation timed out")
        return False
    except Exception as e:
        print_error(f"Error installing PnP.PowerShell: {e}")
        return False


def get_site_url_from_id(site_id: str, access_token: str) -> Optional[str]:
    """Get the SharePoint site URL from site ID using Graph API.
    
    Note: Some special sites (like 'My workspace', OneDrive, etc.) may not have
    accessible URLs via Graph API.
    """
    try:
        site_url_req = f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        req = urllib.request.Request(site_url_req)
        req.add_header("Authorization", f"Bearer {access_token}")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            site_data = json.loads(response.read().decode())
            web_url = site_data.get("webUrl", "")
            
            # Validate the URL looks like a SharePoint site URL
            if web_url and "sharepoint.com" in web_url:
                return web_url
            return None
    except urllib.error.HTTPError as e:
        # 403 = Access denied, 404 = Not found
        # These are expected for special sites like OneDrive
        return None
    except Exception:
        return None


def get_site_recycle_bin_items_pnp(site_url: str) -> List[Dict[str, Any]]:
    """Get all items from a site's recycle bin using PnP PowerShell.
    
    This properly authenticates to SharePoint and can access the recycle bin.
    Prefers PowerShell 7 (pwsh) since PnP.PowerShell is typically installed there.
    """
    # Prefer PowerShell 7 (pwsh) since that's where PnP is typically installed
    pwsh = get_pwsh_executable()
    ps_exe = pwsh if pwsh else get_powershell_executable()
    items = []
    
    ps_script = f'''
$ErrorActionPreference = "Stop"
try {{
    # Connect to SharePoint site (will prompt for credentials if needed)
    Connect-PnPOnline -Url "{site_url}" -Interactive -ErrorAction Stop
    
    # Get recycle bin items
    $recycleBinItems = Get-PnPRecycleBinItem -ErrorAction Stop
    
    if ($recycleBinItems) {{
        # Output as JSON
        $recycleBinItems | Select-Object Id, Title, ItemType, DirName, DeletedByEmail, DeletedDate, ItemState, Size | ConvertTo-Json -Compress
    }} else {{
        Write-Output "[]"
    }}
    
    # Disconnect
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}
'''
    
    try:
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode == 0 and result.stdout.strip():
            output = result.stdout.strip()
            # Handle single item vs array
            if output.startswith('['):
                items = json.loads(output)
            elif output.startswith('{'):
                items = [json.loads(output)]
        elif result.stderr:
            # Don't print error here, let caller handle it
            pass
            
    except subprocess.TimeoutExpired:
        pass
    except json.JSONDecodeError:
        pass
    except Exception:
        pass
    
    return items


def get_site_recycle_bin_items(site_id: str, access_token: str) -> List[Dict[str, Any]]:
    """Get all items from a site's document library recycle bin.
    
    Note: Graph API doesn't have a direct recycle bin endpoint for sites.
    This function returns an empty list - use get_site_recycle_bin_items_pnp()
    with the site URL for actual recycle bin access.
    """
    # Graph API doesn't support site recycle bin access
    # Return empty list - the PnP PowerShell method should be used instead
    return []


def get_site_recycle_bin_count(site_id: str, access_token: str) -> int:
    """Get count of items in site recycle bin.
    
    Note: This always returns 0 as Graph API doesn't support recycle bin access.
    Use PnP PowerShell for actual recycle bin operations.
    """
    return 0


def add_current_user_as_site_admin_pnp(site_url: str, client_id: Optional[str] = None) -> bool:
    """Add the current user as Site Collection Administrator using PnP PowerShell.
    
    This is required to access the recycle bin via PnP PowerShell.
    
    Args:
        site_url: The SharePoint site URL
        client_id: Optional Azure AD app client ID for authentication.
    
    Returns True if successful, False otherwise.
    """
    pwsh = get_pwsh_executable()
    ps_exe = pwsh if pwsh else get_powershell_executable()
    
    if not client_id:
        app_config = load_app_config()
        if app_config:
            client_id = app_config.get("app_id")
    
    # Build connection command
    if client_id:
        connect_cmd = f'Connect-PnPOnline -Url "{site_url}" -Interactive -ClientId "{client_id}" -ErrorAction Stop'
    else:
        connect_cmd = f'Connect-PnPOnline -Url "{site_url}" -Interactive -ErrorAction Stop'
    
    ps_script = f'''
$ErrorActionPreference = "Stop"
try {{
    {connect_cmd}
    
    # Get current user
    $currentUser = Get-PnPProperty -ClientObject (Get-PnPWeb) -Property CurrentUser
    $userEmail = $currentUser.Email
    
    if ($userEmail) {{
        Write-Host "Adding $userEmail as Site Collection Administrator..."
        Set-PnPSiteCollectionAdmin -Owners $userEmail -ErrorAction Stop
        Write-Host "SUCCESS"
    }} else {{
        Write-Host "Could not determine current user email"
        Write-Host "FAILED"
    }}
    
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}} catch {{
    Write-Host "Error: $_"
    Write-Host "FAILED"
}}
'''
    
    try:
        print_info("  Adding current user as Site Collection Administrator...")
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=120
        )
        
        output = result.stdout + result.stderr
        if "SUCCESS" in output:
            print(f"    {Colors.GREEN}✓{Colors.NC} Added as Site Collection Administrator")
            return True
        else:
            print_error(f"  Failed to add as Site Collection Admin: {output[:100]}")
            return False
            
    except subprocess.TimeoutExpired:
        print_error("  Operation timed out")
        return False
    except Exception as e:
        print_error(f"  Error: {e}")
        return False


def purge_site_recycle_bin_pnp(site_url: str, first_stage_only: bool = False, client_id: Optional[str] = None) -> Tuple[int, int]:
    """Permanently delete all items from a site's recycle bin using PnP PowerShell.
    
    Args:
        site_url: The SharePoint site URL
        first_stage_only: If True, only clear first-stage recycle bin.
                         If False, clear both first and second stage.
        client_id: Optional Azure AD app client ID for authentication.
                  If provided, uses -Interactive with -ClientId.
                  If not provided, falls back to -DeviceLogin.
    
    Returns (success_count, fail_count).
    
    Prefers PowerShell 7 (pwsh) since PnP.PowerShell is typically installed there.
    """
    # Prefer PowerShell 7 (pwsh) since that's where PnP is typically installed
    pwsh = get_pwsh_executable()
    ps_exe = pwsh if pwsh else get_powershell_executable()
    
    # Try to get client_id from app config if not provided
    if not client_id:
        app_config = load_app_config()
        if app_config:
            client_id = app_config.get("app_id")
    
    # Build the PowerShell script
    # Note: second_stage is now handled in the main script flow (checked first)
    # This variable is kept for backwards compatibility but is now empty
    second_stage_cmd = ""
    
    # Build connection command
    # Try to get client_id from app config if not provided
    if not client_id:
        app_config = load_app_config()
        if app_config:
            client_id = app_config.get("app_id")
    
    if client_id:
        # Use -Interactive with ClientId - this uses the app registration
        connect_cmd = f'Connect-PnPOnline -Url "{site_url}" -Interactive -ClientId "{client_id}" -ErrorAction Stop'
        connect_admin_cmd = f'Connect-PnPOnline -Url $adminUrl -Interactive -ClientId "{client_id}" -ErrorAction Stop'
        auth_method = "Interactive with registered app"
    else:
        # Fallback to DeviceLogin if no app registration
        connect_cmd = f'Connect-PnPOnline -Url "{site_url}" -DeviceLogin -ErrorAction Stop'
        connect_admin_cmd = 'Connect-PnPOnline -Url $adminUrl -DeviceLogin -ErrorAction Stop'
        auth_method = "Device Login"
    
    ps_script = f'''
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
try {{
    # Connect to SharePoint site using {auth_method}
    Write-Host "Connecting to SharePoint site..."
    Write-Host "Authentication method: {auth_method}"
    
    # Connect using Interactive authentication
    {connect_cmd}
    Write-Host "Connected successfully!"
    
    # Note: To access the Site Collection Recycle Bin (second-stage), you need to be a Site Collection Admin.
    # If you're not seeing items, you may need to be added as a Site Collection Admin via SharePoint Admin Center.
    Write-Host "Checking current user permissions..."
    try {{
        $web = Get-PnPWeb
        $currentUser = Get-PnPProperty -ClientObject $web -Property CurrentUser
        $userEmail = $currentUser.Email
        Write-Host "Current user: $userEmail"
        
        # Check if user is Site Collection Admin
        $site = Get-PnPSite -Includes Owner
        $siteOwner = $site.Owner.Email
        Write-Host "Site Owner: $siteOwner"
        
        if ($userEmail -eq $siteOwner) {{
            Write-Host "You are the Site Owner - full access to recycle bin"
        }} else {{
            Write-Host "Note: You may need Site Collection Admin access to see all recycle bin items"
            Write-Host "If recycle bin appears empty, ask a SharePoint Admin to add you as Site Collection Admin"
        }}
    }} catch {{
        Write-Host "Could not check permissions: $_"
    }}
    
    $totalCount = 0
    
    # Check SECOND-STAGE recycle bin FIRST (Site Collection Recycle Bin)
    # Items deleted from first-stage go here, and AdminRecycleBin.aspx?view=5 shows these
    Write-Host "Checking second-stage recycle bin (Site Collection Recycle Bin)..."
    $secondStageItems = Get-PnPRecycleBinItem -SecondStage -RowLimit 5000 -ErrorAction SilentlyContinue
    
    if ($secondStageItems -and @($secondStageItems).Count -gt 0) {{
        $secondStageCount = @($secondStageItems).Count
        Write-Host "Found $secondStageCount items in second-stage recycle bin"
        
        # List first few items for debugging
        Write-Host "Items found:"
        $secondStageItems | Select-Object -First 5 | ForEach-Object {{ Write-Host "  - $($_.Title) ($($_.ItemType))" }}
        
        # Clear the second-stage recycle bin
        Write-Host "Clearing second-stage recycle bin..."
        Clear-PnPRecycleBinItem -SecondStage -Force -ErrorAction Stop
        Write-Host "Second-stage recycle bin cleared!"
        $totalCount = $totalCount + $secondStageCount
    }} else {{
        Write-Host "Second-stage recycle bin is empty"
    }}
    
    # Now check first-stage recycle bin
    Write-Host "Checking first-stage recycle bin..."
    $recycleBinItems = Get-PnPRecycleBinItem -FirstStage -RowLimit 5000 -ErrorAction SilentlyContinue
    
    if ($recycleBinItems -and @($recycleBinItems).Count -gt 0) {{
        $firstStageCount = @($recycleBinItems).Count
        Write-Host "Found $firstStageCount items in first-stage recycle bin"
        
        # List first few items for debugging
        Write-Host "Items found:"
        $recycleBinItems | Select-Object -First 5 | ForEach-Object {{ Write-Host "  - $($_.Title) ($($_.ItemType))" }}
        
        # Clear the first-stage recycle bin
        Write-Host "Clearing first-stage recycle bin..."
        Clear-PnPRecycleBinItem -All -Force -ErrorAction Stop
        Write-Host "First-stage recycle bin cleared!"
        $totalCount = $totalCount + $firstStageCount
    }} else {{
        Write-Host "First-stage recycle bin is empty"
    }}
    {second_stage_cmd}
    # Output result
    Write-Output "SUCCESS:$totalCount"
    
    # Disconnect
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}
'''
    
    try:
        # Create a temp file to capture the result
        import tempfile
        with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as result_file:
            result_path = result_file.name
        
        # Modify script to write result to file
        ps_script_with_output = ps_script.replace(
            'Write-Output "SUCCESS:$totalCount"',
            f'Write-Output "SUCCESS:$totalCount" | Out-File -FilePath "{result_path}" -Encoding UTF8'
        )
        
        # Run WITHOUT capture_output so browser can open for interactive auth
        # This allows the user to see the authentication prompts
        print_info("    Opening browser for authentication...")
        result = subprocess.run(
            [ps_exe, "-ExecutionPolicy", "Bypass", "-NoProfile", "-Command", ps_script_with_output],
            timeout=300,  # 5 minutes for large recycle bins
            # Don't capture output - let it go to console so user can see auth prompts
        )
        
        # Read result from temp file
        try:
            with open(result_path, 'r', encoding='utf-8') as f:
                output = f.read().strip()
            os.unlink(result_path)  # Clean up temp file
        except Exception:
            output = ""
            try:
                os.unlink(result_path)
            except Exception:
                pass
        
        if result.returncode == 0:
            # Parse the success count from file
            if output.startswith("SUCCESS:"):
                count = int(output.split(":")[1].strip())
                return count, 0
            return 0, 0
        else:
            print_error(f"  PnP PowerShell failed (exit code: {result.returncode})")
            return 0, 1
            
    except subprocess.TimeoutExpired:
        print_error("Operation timed out")
        return 0, 1
    except Exception as e:
        print_error(f"Error purging recycle bin: {e}")
        return 0, 1


def get_sharepoint_access_token(site_url: str) -> Optional[str]:
    """Get a SharePoint-specific access token for REST API calls.
    
    Uses the app registration's client credentials to get a token
    scoped to SharePoint Online.
    """
    app_config = load_app_config()
    if not app_config:
        print_error("No app configuration found. Please set up app registration first.")
        return None
    
    client_id = app_config.get("app_id")
    client_secret = app_config.get("client_secret")
    tenant_id = app_config.get("tenant_id")
    
    if not client_id:
        print_error("App ID not found in configuration.")
        return None
    
    if not client_secret:
        print_error("Client secret not found in app configuration.")
        print_info("Please use 'Manage App Registration' > 'Regenerate Client Secret' to create one.")
        return None
    
    if not tenant_id:
        print_error("Tenant ID not found in configuration.")
        return None
    
    # Extract the SharePoint host from the site URL
    # e.g., https://contoso.sharepoint.com/sites/mysite -> contoso.sharepoint.com
    import re
    match = re.match(r'https://([^/]+)', site_url)
    if not match:
        print_error(f"Could not extract SharePoint host from URL: {site_url}")
        return None
    
    sharepoint_host = match.group(1)
    resource = f"https://{sharepoint_host}"
    
    # Try v2.0 endpoint first with .default scope
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    data = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": f"{resource}/.default"
    }).encode()
    
    try:
        req = urllib.request.Request(token_url, data=data)
        req.add_header("Content-Type", "application/x-www-form-urlencoded")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            result = json.loads(response.read().decode())
            token = result.get("access_token")
            if token:
                return token
    except urllib.error.HTTPError as e:
        # If v2.0 fails, try v1.0 endpoint (legacy but sometimes required for SharePoint)
        print_info(f"    v2.0 token request failed ({e.code}), trying v1.0 endpoint...")
    except Exception as e:
        print_info(f"    v2.0 token request failed: {e}, trying v1.0 endpoint...")
    
    # Try v1.0 endpoint (legacy Azure AD endpoint - sometimes required for SharePoint app-only)
    token_url_v1 = f"https://login.microsoftonline.com/{tenant_id}/oauth2/token"
    
    data_v1 = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "resource": resource
    }).encode()
    
    try:
        req = urllib.request.Request(token_url_v1, data=data_v1)
        req.add_header("Content-Type", "application/x-www-form-urlencoded")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            result = json.loads(response.read().decode())
            return result.get("access_token")
    except Exception as e:
        print_error(f"Failed to get SharePoint token: {e}")
        return None


def get_recycle_bin_items_rest(site_url: str, access_token: str) -> Optional[List[Dict[str, Any]]]:
    """Get recycle bin items using SharePoint REST API.
    
    Args:
        site_url: The SharePoint site URL
        access_token: SharePoint access token
    
    Returns list of recycle bin items, or None if authentication/access failed.
    Empty list means recycle bin is empty. None means we couldn't access it.
    """
    # SharePoint REST API endpoint for recycle bin
    api_url = f"{site_url.rstrip('/')}/_api/web/RecycleBin?$top=5000"
    
    try:
        req = urllib.request.Request(api_url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Accept", "application/json;odata=verbose")
        
        with urllib.request.urlopen(req, timeout=60) as response:
            data = json.loads(response.read().decode())
            items = data.get("d", {}).get("results", [])
            return items
    except urllib.error.HTTPError as e:
        if e.code == 401:
            print_error(f"  Authentication failed (401). SharePoint REST API requires certificate-based auth.")
            return None  # Return None to indicate auth failure, not empty recycle bin
        elif e.code == 403:
            print_error(f"  Access denied to recycle bin (403). Check app permissions.")
            return None  # Return None to indicate access failure
        elif e.code == 404:
            print_error(f"  Site not found (404)")
            return None
        else:
            print_error(f"  HTTP error {e.code}: {e.reason}")
            return None  # Return None for any HTTP error
    except Exception as e:
        print_error(f"  Error getting recycle bin: {e}")
        return None


def delete_recycle_bin_item_rest(site_url: str, item_id: str, access_token: str) -> bool:
    """Delete a single item from the recycle bin using SharePoint REST API.
    
    Args:
        site_url: The SharePoint site URL
        item_id: The recycle bin item ID (GUID)
        access_token: SharePoint access token
    
    Returns True if successful.
    """
    # SharePoint REST API endpoint to delete recycle bin item
    api_url = f"{site_url.rstrip('/')}/_api/web/RecycleBin('{item_id}')"
    
    try:
        req = urllib.request.Request(api_url, method="DELETE")
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Accept", "application/json;odata=verbose")
        req.add_header("IF-MATCH", "*")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            return True
    except urllib.error.HTTPError as e:
        if e.code == 204:  # No content = success
            return True
        return False
    except Exception:
        return False


def clear_recycle_bin_rest(site_url: str, access_token: str) -> bool:
    """Clear all items from the recycle bin using SharePoint REST API.
    
    Args:
        site_url: The SharePoint site URL
        access_token: SharePoint access token
    
    Returns True if successful.
    """
    # SharePoint REST API endpoint to delete all recycle bin items
    api_url = f"{site_url.rstrip('/')}/_api/web/RecycleBin/DeleteAll()"
    
    try:
        req = urllib.request.Request(api_url, method="POST")
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Accept", "application/json;odata=verbose")
        req.add_header("Content-Length", "0")
        
        with urllib.request.urlopen(req, timeout=60) as response:
            return True
    except urllib.error.HTTPError as e:
        if e.code == 204 or e.code == 200:  # Success
            return True
        print_error(f"  HTTP error {e.code}: {e.reason}")
        return False
    except Exception as e:
        print_error(f"  Error clearing recycle bin: {e}")
        return False


def purge_site_recycle_bin_rest(site_url: str) -> Tuple[int, int]:
    """Purge site recycle bin using SharePoint REST API.
    
    This uses the app registration's client credentials (application permissions)
    which should have Sites.FullControl.All permission.
    
    Note: SharePoint REST API with client credentials typically requires
    certificate-based authentication, not client secret. This method may fail
    with 401 errors even with correct permissions.
    
    Args:
        site_url: The SharePoint site URL
    
    Returns (success_count, fail_count).
        - (0, 1) means access failed (should fall back to PnP PowerShell)
        - (0, 0) means recycle bin is empty (no fallback needed)
        - (n, 0) means n items were purged successfully
    """
    # Get SharePoint access token
    print_info("  Getting SharePoint access token...")
    sp_token = get_sharepoint_access_token(site_url)
    
    if not sp_token:
        print_error("  Could not get SharePoint access token")
        return 0, 1  # Return failure to trigger PnP fallback
    
    # Get recycle bin items
    print_info("  Checking recycle bin via REST API...")
    items = get_recycle_bin_items_rest(site_url, sp_token)
    
    # None means access failed (auth error, permission error, etc.)
    # Empty list [] means recycle bin is actually empty
    if items is None:
        print_warning("  REST API access failed - will try PnP PowerShell")
        return 0, 1  # Return failure to trigger PnP fallback
    
    if len(items) == 0:
        print_info("  Recycle bin is empty (confirmed via REST API)")
        return 0, 0  # Actually empty, no fallback needed
    
    item_count = len(items)
    print_info(f"  Found {item_count} items in recycle bin")
    
    # Show first few items
    for item in items[:5]:
        title = item.get("Title", "Unknown")
        item_type = item.get("ItemType", 0)
        type_name = "Folder" if item_type == 1 else "File"
        print(f"      - {title} ({type_name})")
    
    if item_count > 5:
        print(f"      ... and {item_count - 5} more items")
    
    # Clear all items
    print_info("  Clearing recycle bin...")
    if clear_recycle_bin_rest(site_url, sp_token):
        print(f"    {Colors.GREEN}✓{Colors.NC} Cleared {item_count} items from recycle bin")
        return item_count, 0
    else:
        print_error("  Failed to clear recycle bin")
        return 0, 1


def purge_site_recycle_bin(site_id: str, access_token: str) -> Tuple[int, int]:
    """Permanently delete all items from a site's recycle bin.
    
    Note: This function uses Graph API which doesn't support recycle bin operations.
    Use purge_site_recycle_bin_pnp() with the site URL for actual purging.
    
    Returns (success_count, fail_count).
    """
    # Graph API doesn't support recycle bin operations
    # This is a stub - use purge_site_recycle_bin_pnp() instead
    return 0, 0


def delete_site(site_id: str, access_token: str) -> bool:
    """Delete a SharePoint site.
    
    Note: This requires admin permissions and the site goes to recycle bin first.
    """
    # SharePoint sites are deleted via the SharePoint Admin API, not Graph API directly
    # We need to use the site URL to delete it
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}"
    
    try:
        # First, get the site details to get the web URL
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            site_data = json.loads(response.read().decode())
            web_url = site_data.get("webUrl", "")
            
        # Delete the site using the admin API
        # Note: This requires Sites.FullControl.All permission
        delete_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        
        req = urllib.request.Request(delete_url, method="DELETE")
        req.add_header("Authorization", f"Bearer {access_token}")
        
        with urllib.request.urlopen(req, timeout=60) as response:
            return response.status in [200, 204]
            
    except urllib.error.HTTPError as e:
        if e.code == 403:
            print_error("Insufficient permissions to delete site. Need Sites.FullControl.All")
        elif e.code == 404:
            return True  # Already deleted
        return False
    except Exception as e:
        print_error(f"Error deleting site: {e}")
        return False

def delete_all_files_from_site(site: Dict[str, Any], access_token: str) -> Tuple[int, int]:
    """Delete all files and folders from a SharePoint site."""
    site_id = site.get("id", "")
    site_name = site.get("displayName", site.get("name", "Unknown"))
    
    success_count = 0
    fail_count = 0
    
    # Get all root items
    items = get_site_root_items(site_id, access_token)
    total_items = len(items)
    
    if total_items == 0:
        print_info(f"No files found in {site_name}")
        return 0, 0
    
    print_info(f"Found {total_items} items in {site_name}")
    
    for i, item in enumerate(items):
        item_name = item.get("name", "Unknown")
        item_id = item.get("id", "")
        
        # Delete the item (works for both files and folders)
        if "folder" in item:
            success = delete_folder(site_id, item_id, access_token)
        else:
            success = delete_file(site_id, item_id, access_token)
        
        if success:
            success_count += 1
        else:
            fail_count += 1
        
        print_progress(i + 1, total_items, f"Deleting: {item_name[:30]}...")
    
    print()  # New line after progress
    return success_count, fail_count

# ============================================================================
# INTERACTIVE SELECTION FUNCTIONS
# ============================================================================

def display_sites_for_selection(sites: List[Dict[str, Any]]) -> None:
    """Display sites in a numbered list for selection."""
    print()
    print(f"  {Colors.WHITE}Available SharePoint Sites:{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print()
    
    for i, site in enumerate(sites, 1):
        name = site.get("displayName", site.get("name", "Unknown"))
        web_url = site.get("webUrl", "")
        print(f"    {Colors.YELLOW}[{i:3}]{Colors.NC} {name}")
        print(f"          {Colors.BLUE}{web_url}{Colors.NC}")
    
    print()
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")

def display_files_for_selection(files: List[Dict[str, Any]], site_name: str) -> None:
    """Display files in a numbered list for selection."""
    print()
    print(f"  {Colors.WHITE}Files in {site_name}:{Colors.NC}")
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
    print()
    
    for i, file in enumerate(files, 1):
        name = file.get("name", "Unknown")
        path = file.get("_full_path", name)
        size = file.get("size", 0)
        is_folder = "folder" in file
        
        # Format size
        if size < 1024:
            size_str = f"{size} B"
        elif size < 1024 * 1024:
            size_str = f"{size / 1024:.1f} KB"
        else:
            size_str = f"{size / (1024 * 1024):.1f} MB"
        
        if is_folder:
            print(f"    {Colors.YELLOW}[{i:3}]{Colors.NC} {Colors.CYAN}📁 {path}/{Colors.NC}")
        else:
            print(f"    {Colors.YELLOW}[{i:3}]{Colors.NC} 📄 {path} ({size_str})")
    
    print()
    print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")

def interactive_select_sites(sites: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Interactively select specific sites from a list."""
    display_sites_for_selection(sites)
    
    print(f"  {Colors.WHITE}Enter site numbers to select:{Colors.NC}")
    print(f"  {Colors.BLUE}  - Single: 1{Colors.NC}")
    print(f"  {Colors.BLUE}  - Multiple: 1,3,5{Colors.NC}")
    print(f"  {Colors.BLUE}  - Range: 1-5{Colors.NC}")
    print(f"  {Colors.BLUE}  - Combined: 1,3,5-10{Colors.NC}")
    print(f"  {Colors.BLUE}  - All: *{Colors.NC}")
    print()
    
    selection = input(f"  {Colors.YELLOW}Selection:{Colors.NC} ").strip()
    
    if selection == '*':
        return sites
    
    selected_indices = parse_selection(selection, len(sites))
    
    if not selected_indices:
        print_warning("No valid selection made")
        return []
    
    selected_sites = [sites[i - 1] for i in sorted(selected_indices)]
    
    print()
    print_success(f"Selected {len(selected_sites)} site(s):")
    for site in selected_sites:
        name = site.get("displayName", site.get("name", "Unknown"))
        print(f"    - {name}")
    
    return selected_sites

def interactive_select_files(site: Dict[str, Any], access_token: str) -> List[Dict[str, Any]]:
    """Interactively select specific files from a site."""
    site_id = site.get("id", "")
    site_name = site.get("displayName", site.get("name", "Unknown"))
    
    print_info(f"Loading files from {site_name}...")
    files = get_all_site_items(site_id, access_token, include_folders=True)
    
    if not files:
        print_warning(f"No files found in {site_name}")
        return []
    
    display_files_for_selection(files, site_name)
    
    print(f"  {Colors.WHITE}Enter file numbers to select:{Colors.NC}")
    print(f"  {Colors.BLUE}  - Single: 1{Colors.NC}")
    print(f"  {Colors.BLUE}  - Multiple: 1,3,5{Colors.NC}")
    print(f"  {Colors.BLUE}  - Range: 1-5{Colors.NC}")
    print(f"  {Colors.BLUE}  - Combined: 1,3,5-10{Colors.NC}")
    print(f"  {Colors.BLUE}  - All: *{Colors.NC}")
    print()
    
    selection = input(f"  {Colors.YELLOW}Selection:{Colors.NC} ").strip()
    
    if selection == '*':
        return files
    
    selected_indices = parse_selection(selection, len(files))
    
    if not selected_indices:
        print_warning("No valid selection made")
        return []
    
    selected_files = [files[i - 1] for i in sorted(selected_indices)]
    
    print()
    print_success(f"Selected {len(selected_files)} file(s):")
    for f in selected_files[:10]:
        name = f.get("_full_path", f.get("name", "Unknown"))
        print(f"    - {name}")
    if len(selected_files) > 10:
        print(f"    ... and {len(selected_files) - 10} more")
    
    return selected_files

# ============================================================================
# MAIN FUNCTIONS
# ============================================================================

def list_files_mode(sites: List[Dict[str, Any]], access_token: str) -> None:
    """List all files in the selected sites."""
    print_step(5, "List Files in Sites")
    
    for site in sites:
        site_id = site.get("id", "")
        site_name = site.get("displayName", site.get("name", "Unknown"))
        
        print()
        print(f"  {Colors.WHITE}📁 {site_name}{Colors.NC}")
        print(f"  {Colors.CYAN}{'─' * 50}{Colors.NC}")
        
        files = get_all_site_items(site_id, access_token, include_folders=True)
        
        if not files:
            print(f"    {Colors.YELLOW}(No files){Colors.NC}")
            continue
        
        for f in files:
            path = f.get("_full_path", f.get("name", "Unknown"))
            size = f.get("size", 0)
            is_folder = "folder" in f
            
            if size < 1024:
                size_str = f"{size} B"
            elif size < 1024 * 1024:
                size_str = f"{size / 1024:.1f} KB"
            else:
                size_str = f"{size / (1024 * 1024):.1f} MB"
            
            if is_folder:
                print(f"    {Colors.CYAN}📁 {path}/{Colors.NC}")
            else:
                print(f"    📄 {path} ({size_str})")
        
        print(f"    {Colors.GREEN}Total: {len(files)} items{Colors.NC}")

def delete_files_mode(sites: List[Dict[str, Any]], access_token: str, purge_recycle_bin: bool = False) -> None:
    """Delete all files from selected sites.
    
    Args:
        sites: List of SharePoint sites to delete files from
        access_token: Microsoft Graph access token
        purge_recycle_bin: If True, also permanently delete files from recycle bin
    """
    print_step(5, "Delete Files from Sites")
    
    total_success = 0
    total_fail = 0
    
    for site in sites:
        site_name = site.get("displayName", site.get("name", "Unknown"))
        print()
        print_info(f"Processing: {site_name}")
        
        success, fail = delete_all_files_from_site(site, access_token)
        total_success += success
        total_fail += fail
    
    print()
    print_success(f"Deleted {total_success} items successfully")
    if total_fail > 0:
        print_warning(f"Failed to delete {total_fail} items")
    
    # Ask about purging recycle bin if not already specified
    if total_success > 0:
        print()
        print_info("Deleted files are now in the site recycle bin (recoverable for 93 days)")
        
        if not purge_recycle_bin:
            print()
            purge_choice = input(f"  {Colors.YELLOW}Permanently delete from recycle bin? (y/N): {Colors.NC}").strip().lower()
            purge_recycle_bin = purge_choice == 'y'
        
        if purge_recycle_bin:
            print()
            print_step(6, "Purging Site Recycle Bins (using PnP PowerShell)")
            
            # Check if PnP module is installed
            if not check_pnp_module_installed():
                print_warning("PnP.PowerShell module is not installed")
                install_choice = input(f"  {Colors.YELLOW}Install PnP.PowerShell module? (Y/n): {Colors.NC}").strip().lower()
                if install_choice != 'n':
                    if not install_pnp_module():
                        print_error("Could not install PnP.PowerShell. Skipping recycle bin purge.")
                        return
                else:
                    print_info("Skipping recycle bin purge")
                    return
            
            total_purged = 0
            total_purge_fail = 0
            
            for site in sites:
                site_id = site.get("id", "")
                site_name = site.get("displayName", site.get("name", "Unknown"))
                
                # Get site URL from Graph API
                site_url = get_site_url_from_id(site_id, access_token)
                
                if not site_url:
                    print_warning(f"  Could not get URL for site: {site_name}")
                    total_purge_fail += 1
                    continue
                
                print()
                print_info(f"Purging recycle bin: {site_name}")
                print_info(f"  Site URL: {site_url}")
                
                # Try REST API first (uses app permissions), fall back to PnP PowerShell
                purged, purge_fail = purge_site_recycle_bin_rest(site_url)
                if purge_fail > 0:
                    # REST API failed, try PnP PowerShell as fallback
                    print_info("  REST API failed, trying PnP PowerShell...")
                    purged, purge_fail = purge_site_recycle_bin_pnp(site_url)
                total_purged += purged
                total_purge_fail += purge_fail
                
                if purged > 0:
                    print_success(f"  Purged {purged} items from recycle bin")
                elif purge_fail > 0:
                    print_warning(f"  Failed to purge recycle bin")
                else:
                    print_info(f"  Recycle bin is empty")
            
            print()
            if total_purged > 0:
                print_success(f"Permanently deleted {total_purged} items from recycle bins")
            if total_purge_fail > 0:
                print_warning(f"Failed to purge {total_purge_fail} sites (may require additional permissions)")

def delete_selected_files_mode(sites: List[Dict[str, Any]], access_token: str) -> None:
    """Interactively select and delete specific files from sites."""
    print_step(5, "Select and Delete Specific Files")
    
    total_success = 0
    total_fail = 0
    
    for site in sites:
        site_id = site.get("id", "")
        site_name = site.get("displayName", site.get("name", "Unknown"))
        
        print()
        print(f"  {Colors.WHITE}Processing: {site_name}{Colors.NC}")
        
        # Let user select files from this site
        selected_files = interactive_select_files(site, access_token)
        
        if not selected_files:
            continue
        
        # Confirm deletion
        print()
        print_warning(f"About to delete {len(selected_files)} file(s) from {site_name}")
        confirm = input(f"  {Colors.YELLOW}Continue? (y/n):{Colors.NC} ").strip().lower()
        
        if confirm != 'y':
            print_warning("Skipped")
            continue
        
        # Delete selected files
        for i, f in enumerate(selected_files):
            item_id = f.get("id", "")
            item_name = f.get("name", "Unknown")
            is_folder = "folder" in f
            
            if is_folder:
                success = delete_folder(site_id, item_id, access_token)
            else:
                success = delete_file(site_id, item_id, access_token)
            
            if success:
                total_success += 1
            else:
                total_fail += 1
            
            print_progress(i + 1, len(selected_files), f"Deleting: {item_name[:30]}...")
        
        print()  # New line after progress
    
    print()
    print_success(f"Deleted {total_success} items successfully")
    if total_fail > 0:
        print_warning(f"Failed to delete {total_fail} items")

def delete_sites_mode(sites: List[Dict[str, Any]], access_token: str, tenant: Optional[str] = None) -> None:
    """Delete selected SharePoint sites.
    
    For group-connected sites (created via Terraform/M365 Groups), this deletes the M365 Group
    which automatically deletes the associated SharePoint site.
    
    System sites (My workspace, Designer, Team Site, Communication site) are automatically
    filtered out as they cannot be deleted.
    
    Args:
        sites: List of SharePoint sites to delete (may include group info)
        access_token: Microsoft Graph access token
        tenant: SharePoint tenant name for SPO recycle bin purge (e.g., 'contoso')
    """
    print_step(5, "Delete SharePoint Sites")
    
    # First, filter out system sites
    deletable_sites, system_sites = categorize_sites(sites)
    
    # Warn about system sites that will be skipped
    if system_sites:
        print()
        print(f"  {Colors.YELLOW}{'─' * 60}{Colors.NC}")
        print(f"  {Colors.YELLOW}⚠ The following system sites will be SKIPPED:{Colors.NC}")
        print(f"  {Colors.DIM}  (These are protected SharePoint system sites){Colors.NC}")
        print()
        for site in system_sites:
            site_name = site.get("displayName", site.get("name", "Unknown"))
            print(f"    {Colors.DIM}🔒 {site_name}{Colors.NC}")
        print()
        print(f"  {Colors.YELLOW}{'─' * 60}{Colors.NC}")
    
    if not deletable_sites:
        print()
        print_warning("No deletable sites found. All selected sites are protected system sites.")
        return
    
    print()
    print_danger(f"This will permanently delete the following {len(deletable_sites)} site(s):")
    print()
    
    # Categorize deletable sites by deletion method
    group_sites = []  # Sites with M365 Group (delete via group)
    standalone_sites = []  # Sites without group (delete directly)
    
    for site in deletable_sites:
        site_name = site.get("displayName", site.get("name", "Unknown"))
        web_url = site.get("webUrl", site.get("siteUrl", ""))
        
        # Check if this is a group-connected site
        # Sites from get_m365_groups have 'id' as group ID and 'siteId' as SharePoint site ID
        # Sites from get_sharepoint_sites have 'id' as SharePoint site ID
        has_group = "siteId" in site  # If siteId exists, 'id' is the group ID
        
        if has_group:
            group_sites.append(site)
            print(f"    - {site_name} {Colors.CYAN}(M365 Group){Colors.NC}")
        else:
            standalone_sites.append(site)
            print(f"    - {site_name} {Colors.DIM}(Standalone){Colors.NC}")
        print(f"      {web_url}")
    print()
    
    if group_sites:
        print_info(f"{len(group_sites)} site(s) will be deleted via M365 Group deletion")
    if standalone_sites:
        print_info(f"{len(standalone_sites)} site(s) will be deleted directly (requires Sites.FullControl.All)")
    print()
    
    # Double confirmation for site deletion
    confirm1 = input(f"  {Colors.RED}Type 'DELETE' to confirm site deletion:{Colors.NC} ").strip()
    if confirm1 != "DELETE":
        print_warning("Site deletion cancelled")
        return
    
    confirm2 = input(f"  {Colors.RED}Are you ABSOLUTELY sure? Type 'YES' to proceed:{Colors.NC} ").strip()
    if confirm2 != "YES":
        print_warning("Site deletion cancelled")
        return
    
    print()
    success_count = 0
    fail_count = 0
    
    # Delete group-connected sites via M365 Group deletion
    if group_sites:
        print_info("Deleting group-connected sites via M365 Groups...")
        for i, site in enumerate(group_sites):
            site_name = site.get("displayName", site.get("name", "Unknown"))
            group_id = site.get("id", "")  # For group sites, 'id' is the group ID
            
            print_progress(i + 1, len(group_sites), f"Deleting group: {site_name[:30]}...")
            
            if delete_m365_group(group_id, access_token):
                success_count += 1
            else:
                fail_count += 1
        print()
    
    # Delete standalone sites directly
    if standalone_sites:
        print_info("Deleting standalone sites directly...")
        for i, site in enumerate(standalone_sites):
            site_name = site.get("displayName", site.get("name", "Unknown"))
            site_id = site.get("id", "")
            
            print_progress(i + 1, len(standalone_sites), f"Deleting site: {site_name[:30]}...")
            
            if delete_site(site_id, access_token):
                success_count += 1
            else:
                fail_count += 1
        print()
    
    print()
    print()
    
    if success_count > 0:
        print_success(f"Deleted {success_count} sites successfully")
        print_info("Note: Deleted sites go to the SharePoint recycle bin for 93 days")
        
        # Automatically purge recycle bins after deletion
        print()
        print_banner("RECYCLE BIN CLEANUP")
        print()
        print_info("Sites have been soft-deleted. They now exist in two recycle bins:")
        print(f"    {Colors.CYAN}1.{Colors.NC} Microsoft 365 Groups recycle bin (Azure AD)")
        print(f"    {Colors.CYAN}2.{Colors.NC} SharePoint site recycle bin (SharePoint Admin Center)")
        print()
        
        # Try to auto-detect tenant name from site URLs if not provided
        if not tenant:
            for site in sites:
                site_url = site.get("webUrl", site.get("siteUrl", ""))
                if site_url and ".sharepoint.com" in site_url:
                    # Extract tenant from URL like https://contoso.sharepoint.com/sites/...
                    import re
                    match = re.search(r'https://([^.]+)\.sharepoint\.com', site_url)
                    if match:
                        tenant = match.group(1)
                        print_info(f"Auto-detected tenant name: {tenant}")
                        break
        
        # Ask if user wants to skip recycle bin purge
        skip_choice = input(f"  {Colors.YELLOW}Purge recycle bins now? (Y/n): {Colors.NC}").strip().lower()
        if skip_choice == 'n':
            print_warning("Skipping recycle bin purge. Sites remain in recycle bins.")
            print_info("You can purge them later using menu options [6] and [7]")
        else:
            # Step 1: Purge M365 Groups recycle bin
            print()
            print_step(1, "Purging M365 Groups recycle bin (Azure AD)")
            
            # Wait a moment for Azure AD to process the deletions
            print_info("Waiting for Azure AD to process deletions...")
            import time
            time.sleep(3)
            
            deleted_groups = get_deleted_m365_groups(access_token)
            if deleted_groups:
                print_info(f"Found {len(deleted_groups)} deleted groups in recycle bin")
                purge_deleted_groups_mode(deleted_groups, access_token, auto_confirm=True)
            else:
                print_info("No deleted groups found in Azure AD recycle bin")
            
            # Step 2: Purge SharePoint site recycle bin
            print()
            print_step(2, "Purging SharePoint site recycle bin")
            
            # Get tenant name if not provided (and not auto-detected above)
            if not tenant:
                print()
                print(f"  {Colors.WHITE}Enter your SharePoint tenant name{Colors.NC}")
                print(f"  {Colors.DIM}(e.g., 'contoso' for contoso.sharepoint.com){Colors.NC}")
                print()
                tenant = input(f"  {Colors.YELLOW}Tenant name (or press Enter to skip): {Colors.NC}").strip()
            
            if tenant:
                admin_url = f"https://{tenant}-admin.sharepoint.com"
                purge_spo_deleted_sites_mode(admin_url, auto_confirm=True)
            else:
                print_warning("Skipping SharePoint recycle bin purge (no tenant provided)")
                print_info("You can purge it later using menu option [7]")
        
    if fail_count > 0:
        print_warning(f"Failed to delete {fail_count} sites")
        print_info("You may need SharePoint Admin permissions to delete sites")

def main() -> None:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Delete files and/or SharePoint sites",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
WARNING: This script performs DESTRUCTIVE operations!

Examples:
    python cleanup.py                           # Interactive mode
    python cleanup.py --delete-files            # Delete all files from all sites
    python cleanup.py --delete-sites            # Delete SharePoint sites
    python cleanup.py --delete-all              # Delete both files and sites
    python cleanup.py --site hr --delete-files  # Delete files from HR sites only
    python cleanup.py --select-sites            # Interactively select specific sites
    python cleanup.py --select-files            # Interactively select specific files to delete
    python cleanup.py --list-sites              # List available sites
    python cleanup.py --list-files              # List files in all sites
    python cleanup.py --list-files --site hr    # List files in a SPECIFIC site

Selection Syntax (for --select-sites and --select-files):
    - Single item:    1
    - Multiple items: 1,3,5
    - Range:          1-5
    - Combined:       1,3,5-10
    - All items:      *
        """
    )
    
    parser.add_argument(
        '--delete-files',
        action='store_true',
        help='Delete all files from SharePoint sites'
    )
    parser.add_argument(
        '--delete-sites',
        action='store_true',
        help='Delete SharePoint sites (requires admin permissions)'
    )
    parser.add_argument(
        '--delete-all',
        action='store_true',
        help='Delete both files and sites'
    )
    parser.add_argument(
        '-s', '--site',
        type=str,
        metavar='FILTER',
        help='Filter sites by name (e.g., "hr", "finance")'
    )
    parser.add_argument(
        '-l', '--list-sites',
        action='store_true',
        help='List available SharePoint sites and exit'
    )
    parser.add_argument(
        '--list-files',
        action='store_true',
        help='List files in SharePoint sites'
    )
    parser.add_argument(
        '--select-sites',
        action='store_true',
        help='Interactively select specific sites from a numbered list'
    )
    parser.add_argument(
        '--select-files',
        action='store_true',
        help='Interactively select specific files to delete'
    )
    parser.add_argument(
        '-y', '--yes',
        action='store_true',
        help='Skip confirmation prompts (use with caution!)'
    )
    parser.add_argument(
        '--list-groups',
        action='store_true',
        help='List Microsoft 365 Groups (which have SharePoint sites)'
    )
    parser.add_argument(
        '--delete-groups',
        action='store_true',
        help='Delete Microsoft 365 Groups (and their SharePoint sites)'
    )
    parser.add_argument(
        '--list-deleted',
        action='store_true',
        help='List deleted Microsoft 365 Groups in the recycle bin'
    )
    parser.add_argument(
        '--purge-deleted',
        action='store_true',
        help='Permanently delete groups from the recycle bin (frees up URLs)'
    )
    parser.add_argument(
        '--purge-spo-recycle',
        action='store_true',
        help='Permanently delete sites from SharePoint recycle bin (requires SPO PowerShell)'
    )
    parser.add_argument(
        '--purge-site-recycle',
        action='store_true',
        help='Purge files/folders from site document library recycle bins'
    )
    parser.add_argument(
        '--tenant',
        type=str,
        metavar='NAME',
        help='SharePoint tenant name (e.g., "contoso" for contoso.sharepoint.com)'
    )
    
    args = parser.parse_args()
    
    # Determine if this is a read-only operation
    is_read_only = (args.list_sites or args.list_files or args.list_groups or args.list_deleted) and not (
        args.delete_files or args.delete_sites or args.delete_all or
        args.delete_groups or args.purge_deleted or args.purge_spo_recycle or args.purge_site_recycle or args.select_files
    )
    
    # Clear screen and show appropriate banner
    os.system('cls' if os.name == 'nt' else 'clear')
    
    if is_read_only:
        print_banner("📋 SHAREPOINT SITES")
        print()
    else:
        print_banner("⚠️  SHAREPOINT CLEANUP  ⚠️")
        print_danger("This script performs DESTRUCTIVE operations!")
        print_danger("Deleted files and sites may not be recoverable!")
        print()
    
    # Step 1: Select Environment (auto-detect from environments.json)
    print_step(1, "Select Environment")
    
    selected_env = select_environment()
    
    if selected_env:
        tenant_id = get_environment_tenant(selected_env)
        if tenant_id:
            if not switch_to_tenant(tenant_id):
                print_error("Failed to switch to configured tenant")
                sys.exit(1)
    else:
        print_info("No environment configured, using current Azure CLI login")
    
    # Step 2: Check Azure login
    print_step(2, "Check Azure Authentication")
    
    if not check_azure_login():
        print_warning("Not logged into Azure CLI")
        if not azure_login():
            print_error("Azure login failed")
            sys.exit(1)
    
    print_success("Azure CLI authenticated")
    
    # Step 3: Get access token
    print_step(3, "Get Microsoft Graph Access Token")
    
    access_token = get_access_token()
    if not access_token:
        print_error("Failed to get access token")
        print_info("Make sure you have the required permissions:")
        print_info("  - Sites.ReadWrite.All (for file deletion)")
        print_info("  - Sites.FullControl.All (for site deletion)")
        sys.exit(1)
    
    print_success("Access token obtained")
    
    # Handle Microsoft 365 Groups mode (separate from SharePoint sites)
    if args.list_groups or args.delete_groups:
        print_step(4, "Discover Microsoft 365 Groups")
        
        groups = get_m365_groups(access_token)
        
        if not groups:
            print_warning("No Microsoft 365 Groups found")
            print_info("Note: Only groups with SharePoint sites are shown")
            sys.exit(0)
        
        print_success(f"Found {len(groups)} Microsoft 365 Groups with SharePoint sites")
        
        # Filter groups if specified
        if args.site:
            filter_term = args.site.lower()
            groups = [g for g in groups if filter_term in g.get("displayName", "").lower()
                     or filter_term in g.get("mail", "").lower()]
            if not groups:
                print_error(f"No groups found matching '{args.site}'")
                sys.exit(1)
            print_info(f"Filtered to {len(groups)} groups matching '{args.site}'")
        
        if args.list_groups:
            display_groups_for_selection(groups)
            sys.exit(0)
        
        if args.delete_groups:
            selected_groups = interactive_select_groups(groups)
            if selected_groups:
                delete_groups_mode(selected_groups, access_token, args.yes, tenant=args.tenant)
            sys.exit(0)
    
    # Handle deleted groups (recycle bin) mode
    if args.list_deleted or args.purge_deleted:
        print_step(4, "Discover Deleted Microsoft 365 Groups (Recycle Bin)")
        
        deleted_groups = get_deleted_m365_groups(access_token)
        
        if not deleted_groups:
            print_warning("No deleted Microsoft 365 Groups found in recycle bin")
            print_info("The recycle bin is empty - all URLs are available for reuse")
            sys.exit(0)
        
        print_success(f"Found {len(deleted_groups)} deleted groups in recycle bin")
        
        # Filter groups if specified
        if args.site:
            filter_term = args.site.lower()
            deleted_groups = [g for g in deleted_groups if filter_term in g.get("displayName", "").lower()]
            if not deleted_groups:
                print_error(f"No deleted groups found matching '{args.site}'")
                sys.exit(1)
            print_info(f"Filtered to {len(deleted_groups)} deleted groups matching '{args.site}'")
        
        if args.list_deleted:
            display_deleted_groups_for_selection(deleted_groups)
            print()
            print_info("To permanently delete these groups and free up URLs, run:")
            print_info("  python cleanup.py --purge-deleted")
            sys.exit(0)
        
        if args.purge_deleted:
            purge_deleted_groups_mode(deleted_groups, access_token, args.yes)
            sys.exit(0)
    
    # Handle SharePoint site recycle bin (requires SPO PowerShell)
    if args.purge_spo_recycle:
        print_step(4, "Purge SharePoint Site Recycle Bin")
        
        # Get tenant name from args or environment
        tenant_name = args.tenant
        if not tenant_name:
            # Try to get from environment config
            if selected_env:
                m365_config = selected_env.get("m365", {})
                tenant_name = m365_config.get("tenant_name", "")
        
        if not tenant_name:
            print_warning("Tenant name not specified")
            tenant_name = input("  Enter your SharePoint tenant name (e.g., 'contoso' for contoso.sharepoint.com): ").strip()
            if not tenant_name:
                print_error("Tenant name is required for SharePoint recycle bin operations")
                sys.exit(1)
        
        admin_url = f"https://{tenant_name}-admin.sharepoint.com"
        print_info(f"SharePoint Admin URL: {admin_url}")
        
        purge_spo_deleted_sites_mode(admin_url, args.yes)
        sys.exit(0)
    
    # Handle site document library recycle bin purge
    if args.purge_site_recycle:
        print_step(4, "Purge Site Files/Folders Recycle Bins (using PnP PowerShell)")
        
        # Check if PnP module is installed
        if not check_pnp_module_installed():
            print_warning("PnP.PowerShell module is not installed")
            print_info("This module is required to access site recycle bins")
            print()
            install_choice = input(f"  {Colors.YELLOW}Install PnP.PowerShell module? (Y/n): {Colors.NC}").strip().lower()
            if install_choice != 'n':
                if not install_pnp_module():
                    print_error("Could not install PnP.PowerShell. Cannot proceed.")
                    sys.exit(1)
            else:
                print_warning("Cannot purge recycle bins without PnP.PowerShell")
                sys.exit(0)
        
        # Get sites first
        sites = get_sharepoint_sites(access_token)
        
        if not sites:
            print_warning("No SharePoint sites found")
            sys.exit(0)
        
        # Filter sites if specified
        if args.site:
            filter_term = args.site.lower()
            sites = [s for s in sites if filter_term in s.get("displayName", "").lower() or
                     filter_term in s.get("name", "").lower()]
            if not sites:
                print_error(f"No sites found matching '{args.site}'")
                sys.exit(1)
            print_info(f"Filtered to {len(sites)} sites matching '{args.site}'")
        
        print_success(f"Found {len(sites)} sites")
        print()
        
        # Confirm before purging
        if not args.yes:
            print_warning("This will permanently delete all items from site recycle bins!")
            print_info("You will be prompted to authenticate to each site via browser")
            print()
            for site in sites[:10]:
                name = site.get("displayName", site.get("name", "Unknown"))
                print(f"    - {name}")
            if len(sites) > 10:
                print(f"    ... and {len(sites) - 10} more")
            print()
            confirm = input(f"  {Colors.YELLOW}Proceed with purging recycle bins? (y/N): {Colors.NC}").strip().lower()
            if confirm != 'y':
                print_warning("Operation cancelled")
                sys.exit(0)
        
        # Purge recycle bins using PnP PowerShell
        total_purged = 0
        total_failed = 0
        
        for site in sites:
            site_id = site.get("id", "")
            site_name = site.get("displayName", site.get("name", "Unknown"))
            
            # Get site URL from Graph API
            site_url = get_site_url_from_id(site_id, access_token)
            
            if not site_url:
                print_warning(f"  Could not get URL for site: {site_name}")
                total_failed += 1
                continue
            
            print()
            print_info(f"Purging recycle bin: {site_name}")
            print_info(f"  Site URL: {site_url}")
            
            # Try REST API first (uses app permissions), fall back to PnP PowerShell
            purged, failed = purge_site_recycle_bin_rest(site_url)
            if failed > 0:
                # REST API failed, try PnP PowerShell as fallback
                print_info("  REST API failed, trying PnP PowerShell...")
                purged, failed = purge_site_recycle_bin_pnp(site_url)
            total_purged += purged
            total_failed += failed
            
            if purged > 0:
                print_success(f"  Purged {purged} items from recycle bin")
            elif failed > 0:
                print_warning(f"  Failed to purge recycle bin")
            else:
                print_info(f"  Recycle bin is empty")
        
        print()
        if total_purged > 0:
            print_success(f"Purged recycle bins from {len(sites)} sites ({total_purged} total items)")
        if total_failed > 0:
            print_warning(f"Failed to purge {total_failed} sites (may require additional permissions)")
        
        sys.exit(0)
    
    # Step 4: Get SharePoint sites (or fall back to Groups if Sites API fails)
    print_step(4, "Discover SharePoint Sites")
    
    sites = get_sharepoint_sites(access_token)
    
    if not sites:
        print_warning("Could not list SharePoint sites (may require Sites.Read.All permission)")
        print_info("Falling back to Microsoft 365 Groups API...")
        print()
        
        # Fall back to Groups API
        groups = get_m365_groups(access_token)
        
        if groups:
            print_success(f"Found {len(groups)} Microsoft 365 Groups with SharePoint sites")
            print()
            print(f"  {Colors.YELLOW}Note: These are group-connected SharePoint sites.{Colors.NC}")
            print(f"  {Colors.YELLOW}To delete them, you must delete the associated Microsoft 365 Group.{Colors.NC}")
            print()
            
            # Convert groups to a sites-like format for display
            display_groups_for_selection(groups)
            
            # If --list-sites was passed, just list and exit (don't prompt for deletion)
            if args.list_sites or args.list_files:
                print()
                sys.exit(0)
            
            print()
            print(f"  {Colors.WHITE}Would you like to delete any of these groups (and their SharePoint sites)?{Colors.NC}")
            print()
            confirm = input("  Enter 'yes' to select groups to delete, or press Enter to exit: ").strip().lower()
            
            if confirm == 'yes':
                selected_groups = interactive_select_groups(groups)
                if selected_groups:
                    delete_groups_mode(selected_groups, access_token, args.yes, tenant=args.tenant)
            
            sys.exit(0)
        else:
            print_error("No SharePoint sites or Microsoft 365 Groups found")
            print_info("This may be a permissions issue. Required permissions:")
            print_info("  - Sites.Read.All (for SharePoint sites)")
            print_info("  - Group.Read.All (for Microsoft 365 Groups)")
            sys.exit(1)
    
    print_success(f"Found {len(sites)} SharePoint sites")
    
    # Filter sites if specified
    if args.site:
        filter_term = args.site.lower()
        sites = [s for s in sites if filter_term in s.get("name", "").lower()
                 or filter_term in s.get("displayName", "").lower()]
        if not sites:
            print_error(f"No sites found matching '{args.site}'")
            sys.exit(1)
        print_info(f"Filtered to {len(sites)} sites matching '{args.site}'")
    
    # Interactive site selection mode
    if args.select_sites:
        sites = interactive_select_sites(sites)
        if not sites:
            print_warning("No sites selected")
            sys.exit(0)
    
    # List sites mode
    if args.list_sites:
        print()
        print(f"  {Colors.WHITE}Available SharePoint Sites:{Colors.NC}")
        print()
        for i, site in enumerate(sites, 1):
            name = site.get("displayName", site.get("name", "Unknown"))
            web_url = site.get("webUrl", "")
            print(f"    {i:3}. {name}")
            print(f"         {web_url}")
        print()
        sys.exit(0)
    
    # List files mode
    if args.list_files:
        list_files_mode(sites, access_token)
        sys.exit(0)
    
    # Interactive file selection mode
    if args.select_files:
        delete_selected_files_mode(sites, access_token)
        sys.exit(0)
    
    # Determine operation mode
    delete_files = args.delete_files or args.delete_all
    delete_sites = args.delete_sites or args.delete_all
    
    # Interactive mode if no flags specified
    while not delete_files and not delete_sites:
        print()
        print(f"  {Colors.WHITE}What would you like to do?{Colors.NC}")
        print()
        print(f"    {Colors.CYAN}[1]{Colors.NC} 📋 List all sites")
        print(f"    {Colors.CYAN}[2]{Colors.NC} 📁 List files in sites")
        print()
        print(f"    {Colors.RED}[3]{Colors.NC} 🗑️  Delete all SITES (and all their content)")
        print(f"    {Colors.YELLOW}[4]{Colors.NC} 🗑️  Delete all FILES from sites (keeps sites)")
        print()
        print(f"    {Colors.RED}[5]{Colors.NC} 🎯 Select specific SITES to delete")
        print(f"    {Colors.YELLOW}[6]{Colors.NC} 🎯 Select specific FILES to delete")
        print()
        print(f"    {Colors.DIM}[7] Cancel / Exit{Colors.NC}")
        print()
        
        choice = input("  Enter your choice (1-7): ").strip()
        
        if choice == "1":
            # List all sites - categorize into deletable and system sites
            deletable_sites, system_sites = categorize_sites(sites)
            
            print()
            print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
            print(f"  {Colors.WHITE}{Colors.BOLD}SharePoint Sites ({len(sites)} total){Colors.NC}")
            print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
            
            # Show deletable sites first
            if deletable_sites:
                print()
                print(f"  {Colors.GREEN}✓ Deletable Sites ({len(deletable_sites)}):{Colors.NC}")
                print(f"  {Colors.DIM}  These sites can be deleted via M365 Groups or Sites API{Colors.NC}")
                print()
                for i, site in enumerate(deletable_sites, 1):
                    name = site.get("displayName", site.get("name", "Unknown"))
                    url = site.get("webUrl", "")
                    # Check if it's a group-connected site
                    is_group = "siteId" in site
                    group_tag = f" {Colors.CYAN}(M365 Group){Colors.NC}" if is_group else ""
                    print(f"    [{i:2}] {name}{group_tag}")
                    if url:
                        print(f"         {Colors.DIM}{url}{Colors.NC}")
            else:
                print()
                print(f"  {Colors.YELLOW}No deletable sites found.{Colors.NC}")
            
            # Show system sites separately
            if system_sites:
                print()
                print(f"  {Colors.RED}🔒 Protected System Sites ({len(system_sites)}):{Colors.NC}")
                print(f"  {Colors.DIM}  These are built-in SharePoint sites that cannot be deleted{Colors.NC}")
                print()
                for site in system_sites:
                    name = site.get("displayName", site.get("name", "Unknown"))
                    url = site.get("webUrl", "")
                    print(f"    {Colors.DIM}•{Colors.NC} {Colors.DIM}{name}{Colors.NC}")
                    if url:
                        print(f"      {Colors.DIM}{url}{Colors.NC}")
            
            print()
            print(f"  {Colors.CYAN}{'─' * 60}{Colors.NC}")
            print()
            input(f"  {Colors.YELLOW}Press Enter to return to menu...{Colors.NC}")
            # Loop continues - shows menu again
        elif choice == "2":
            list_files_mode(sites, access_token)
            input(f"\n  {Colors.YELLOW}Press Enter to return to menu...{Colors.NC}")
            # Loop continues - shows menu again
        elif choice == "3":
            delete_sites = True
            # Exit loop to proceed with deletion
        elif choice == "4":
            delete_files = True
            # Exit loop to proceed with deletion
        elif choice == "5":
            selected_sites = interactive_select_sites(sites)
            if not selected_sites:
                print_warning("No sites selected")
                input(f"\n  {Colors.YELLOW}Press Enter to return to menu...{Colors.NC}")
                continue
            sites = selected_sites
            delete_sites = True
            # Exit loop to proceed with deletion
        elif choice == "6":
            delete_selected_files_mode(sites, access_token)
            input(f"\n  {Colors.YELLOW}Press Enter to return to menu...{Colors.NC}")
            # Loop continues - shows menu again
        elif choice == "7" or choice.lower() == "q":
            print()
            print(f"  {Colors.GREEN}✓{Colors.NC} Exiting cleanup. No changes made.")
            sys.exit(0)
        else:
            print_warning("Invalid choice. Please enter 1-7.")
            input(f"\n  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            # Loop continues - shows menu again
    
    # Show what will be affected
    print()
    print(f"  {Colors.WHITE}Sites that will be affected:{Colors.NC}")
    for site in sites[:10]:
        name = site.get("displayName", site.get("name", "Unknown"))
        print(f"    - {name}")
    if len(sites) > 10:
        print(f"    ... and {len(sites) - 10} more")
    print()
    
    # Confirmation
    if not args.yes:
        if delete_sites:
            print_danger("You are about to DELETE SharePoint sites!")
        elif delete_files:
            print_warning("You are about to delete all files from these sites!")
        
        confirm = input(f"\n  {Colors.YELLOW}Continue? (y/n):{Colors.NC} ").strip().lower()
        if confirm != 'y':
            print_warning("Operation cancelled")
            sys.exit(0)
    
    # Execute operations
    start_time = datetime.now()
    
    if delete_files and not delete_sites:
        delete_files_mode(sites, access_token)
    elif delete_sites and not delete_files:
        delete_sites_mode(sites, access_token, tenant=args.tenant)
    elif delete_files and delete_sites:
        # Delete files first, then sites
        print_info("Deleting files first, then sites...")
        delete_files_mode(sites, access_token)
        delete_sites_mode(sites, access_token, tenant=args.tenant)
    
    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds()
    
    print()
    print(f"  {Colors.BLUE}⏱ Total duration:{Colors.NC} {duration:.1f} seconds")
    print()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print()
        print_warning("Operation cancelled by user")
        sys.exit(0)
    except Exception as e:
        print_error(f"Unexpected error: {e}")
        sys.exit(1)
