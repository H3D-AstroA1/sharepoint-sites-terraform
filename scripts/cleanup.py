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
    python cleanup.py --help                    # Show help

Requirements:
    - Python 3.8+
    - Azure CLI (logged in)
    - Microsoft Graph API permissions (Sites.ReadWrite.All, Files.ReadWrite.All)

WARNING: This script performs DESTRUCTIVE operations. Always backup important data first!
"""

import argparse
import json
import os
import subprocess
import sys
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple

# ============================================================================
# CONFIGURATION
# ============================================================================

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_DIR = SCRIPT_DIR.parent
CONFIG_DIR = PROJECT_DIR / "config"

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

    @classmethod
    def disable(cls) -> None:
        """Disable colors for non-terminal output."""
        cls.RED = cls.GREEN = cls.YELLOW = cls.BLUE = ''
        cls.MAGENTA = cls.CYAN = cls.WHITE = cls.NC = cls.BOLD = ''

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
# HELPER FUNCTIONS
# ============================================================================

def run_command(command: List[str], capture_output: bool = True) -> Optional[subprocess.CompletedProcess]:
    """Run a shell command and return the result."""
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
    try:
        subprocess.run(["az", "login"], check=True)
        return True
    except subprocess.CalledProcessError:
        return False

def get_access_token() -> Optional[str]:
    """Get Microsoft Graph access token using Azure CLI."""
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
    except Exception as e:
        print_error(f"Error getting sites: {e}")
    
    return sites

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
# MAIN FUNCTIONS
# ============================================================================

def delete_files_mode(sites: List[Dict[str, Any]], access_token: str) -> None:
    """Delete all files from selected sites."""
    print_step(4, "Delete Files from Sites")
    
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

def delete_sites_mode(sites: List[Dict[str, Any]], access_token: str) -> None:
    """Delete selected SharePoint sites."""
    print_step(4, "Delete SharePoint Sites")
    
    print()
    print_danger("This will permanently delete the following sites:")
    print()
    for site in sites:
        site_name = site.get("displayName", site.get("name", "Unknown"))
        web_url = site.get("webUrl", "")
        print(f"    - {site_name}")
        print(f"      {web_url}")
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
    
    for i, site in enumerate(sites):
        site_name = site.get("displayName", site.get("name", "Unknown"))
        site_id = site.get("id", "")
        
        print_progress(i + 1, len(sites), f"Deleting: {site_name[:30]}...")
        
        if delete_site(site_id, access_token):
            success_count += 1
        else:
            fail_count += 1
    
    print()
    print()
    
    if success_count > 0:
        print_success(f"Deleted {success_count} sites successfully")
        print_info("Note: Deleted sites go to the SharePoint recycle bin for 93 days")
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
    python cleanup.py --list-sites              # List available sites
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
        '-y', '--yes',
        action='store_true',
        help='Skip confirmation prompts (use with caution!)'
    )
    
    args = parser.parse_args()
    
    # Clear screen and show banner
    os.system('cls' if os.name == 'nt' else 'clear')
    print_banner("⚠️  SHAREPOINT CLEANUP  ⚠️")
    
    print_danger("This script performs DESTRUCTIVE operations!")
    print_danger("Deleted files and sites may not be recoverable!")
    print()
    
    # Step 1: Check Azure login
    print_step(1, "Check Azure Authentication")
    
    if not check_azure_login():
        print_warning("Not logged into Azure CLI")
        if not azure_login():
            print_error("Azure login failed")
            sys.exit(1)
    
    print_success("Azure CLI authenticated")
    
    # Step 2: Get access token
    print_step(2, "Get Microsoft Graph Access Token")
    
    access_token = get_access_token()
    if not access_token:
        print_error("Failed to get access token")
        print_info("Make sure you have the required permissions:")
        print_info("  - Sites.ReadWrite.All (for file deletion)")
        print_info("  - Sites.FullControl.All (for site deletion)")
        sys.exit(1)
    
    print_success("Access token obtained")
    
    # Step 3: Get SharePoint sites
    print_step(3, "Discover SharePoint Sites")
    
    sites = get_sharepoint_sites(access_token)
    
    if not sites:
        print_error("No SharePoint sites found")
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
    
    # Determine operation mode
    delete_files = args.delete_files or args.delete_all
    delete_sites = args.delete_sites or args.delete_all
    
    # Interactive mode if no flags specified
    if not delete_files and not delete_sites:
        print()
        print(f"  {Colors.WHITE}What would you like to do?{Colors.NC}")
        print()
        print("    [1] Delete all FILES from sites (keeps sites)")
        print("    [2] Delete SITES (and all their content)")
        print("    [3] Delete BOTH files and sites")
        print("    [4] Cancel")
        print()
        
        choice = input("  Enter your choice (1-4): ").strip()
        
        if choice == "1":
            delete_files = True
        elif choice == "2":
            delete_sites = True
        elif choice == "3":
            delete_files = True
            delete_sites = True
        else:
            print_warning("Operation cancelled")
            sys.exit(0)
    
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
        delete_sites_mode(sites, access_token)
    elif delete_files and delete_sites:
        # Delete files first, then sites
        print_info("Deleting files first, then sites...")
        delete_files_mode(sites, access_token)
        delete_sites_mode(sites, access_token)
    
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
