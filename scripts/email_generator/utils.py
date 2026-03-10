"""
Utility functions for M365 email population.

Common helper functions used across the email generator package.
"""

import os
import platform
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any, List


# =============================================================================
# CONSOLE OUTPUT HELPERS
# =============================================================================

class Colors:
    """ANSI color codes for terminal output."""
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    NC = '\033[0m'  # No Color
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


def print_banner(text: str) -> None:
    """Print a banner with the given text."""
    print()
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print(f"  {Colors.BOLD}{Colors.WHITE}{text.center(60)}{Colors.NC}")
    print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
    print()


def print_step(number: int, title: str) -> None:
    """Print a step header."""
    print()
    print(f"  {Colors.CYAN}[Step {number}]{Colors.NC} {Colors.BOLD}{title}{Colors.NC}")
    print(f"  {Colors.CYAN}{'-' * 50}{Colors.NC}")


def print_success(message: str) -> None:
    """Print a success message."""
    print(f"  {Colors.GREEN}✓{Colors.NC} {message}")


def print_error(message: str) -> None:
    """Print an error message."""
    print(f"  {Colors.RED}✗{Colors.NC} {message}")


def print_warning(message: str) -> None:
    """Print a warning message."""
    print(f"  {Colors.YELLOW}⚠{Colors.NC} {message}")


def print_info(message: str) -> None:
    """Print an info message."""
    print(f"  {Colors.BLUE}ℹ{Colors.NC} {message}")


def print_progress(current: int, total: int, message: str) -> None:
    """Print progress indicator."""
    percentage = (current / total) * 100 if total > 0 else 0
    bar_length = 30
    filled = int(bar_length * current / total) if total > 0 else 0
    bar = '█' * filled + '░' * (bar_length - filled)
    print(f"\r  {Colors.CYAN}[{bar}]{Colors.NC} {percentage:5.1f}% - {message:<40}", end='', flush=True)


def print_summary_box(title: str, items: List[tuple]) -> None:
    """Print a summary box with key-value pairs."""
    print()
    print(f"  {Colors.CYAN}╭{'─' * 58}╮{Colors.NC}")
    print(f"  {Colors.CYAN}│{Colors.NC} {Colors.BOLD}{title:<56}{Colors.NC} {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}├{'─' * 58}┤{Colors.NC}")
    for label, value in items:
        print(f"  {Colors.CYAN}│{Colors.NC}   {label:<25} {str(value):<29} {Colors.CYAN}│{Colors.NC}")
    print(f"  {Colors.CYAN}╰{'─' * 58}╯{Colors.NC}")
    print()


# =============================================================================
# AZURE CLI HELPERS
# =============================================================================

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
            timeout=10,
        )
        return result.returncode == 0
    except Exception:
        return False


def get_azure_account_info() -> Optional[Dict[str, Any]]:
    """Get the current Azure account information."""
    import json
    
    az_path = find_azure_cli_path()
    if not az_path:
        return None
    
    try:
        result = subprocess.run(
            [az_path, "account", "show", "--output", "json"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        if result.returncode == 0:
            return json.loads(result.stdout)
    except Exception:
        pass
    
    return None


def azure_login() -> bool:
    """Perform Azure CLI login."""
    az_path = find_azure_cli_path()
    if not az_path:
        print_error("Azure CLI is not installed or not in PATH.")
        print_info("Please install Azure CLI: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli")
        return False
    
    print_info("Opening browser for Azure login...")
    try:
        subprocess.run([az_path, "login"], check=True)
        return True
    except subprocess.CalledProcessError:
        return False
    except Exception as e:
        print_error(f"Azure login failed: {e}")
        return False


def switch_to_tenant(tenant_id: str) -> bool:
    """Switch Azure CLI to the specified tenant."""
    if not tenant_id:
        return True
    
    az_path = find_azure_cli_path()
    if not az_path:
        return False
    
    try:
        # Check if already logged into this tenant
        result = subprocess.run(
            [az_path, "account", "show", "--query", "tenantId", "-o", "tsv"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        current_tenant = result.stdout.strip()
        
        if current_tenant == tenant_id:
            print_info(f"Already logged into tenant: {tenant_id[:20]}...")
            return True
        
        # Need to switch tenant
        print_info(f"Switching to tenant: {tenant_id[:20]}...")
        subprocess.run(
            [az_path, "login", "--tenant", tenant_id],
            check=True,
        )
        return True
        
    except subprocess.CalledProcessError:
        print_error(f"Failed to switch to tenant: {tenant_id}")
        return False
    except Exception as e:
        print_error(f"Error switching tenant: {e}")
        return False


# =============================================================================
# DATE/TIME HELPERS
# =============================================================================

def format_duration(seconds: float) -> str:
    """Format duration in seconds to human-readable string."""
    if seconds < 60:
        return f"{seconds:.1f} seconds"
    elif seconds < 3600:
        minutes = seconds / 60
        return f"{minutes:.1f} minutes"
    else:
        hours = seconds / 3600
        return f"{hours:.1f} hours"


def format_datetime(dt: datetime) -> str:
    """Format datetime for display."""
    return dt.strftime("%B %d, %Y at %I:%M %p")


def format_date(dt: datetime) -> str:
    """Format date for display."""
    return dt.strftime("%B %d, %Y")


# =============================================================================
# FILE SIZE HELPERS
# =============================================================================

def format_size(size_bytes: int) -> str:
    """Format file size for display."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"


# =============================================================================
# VALIDATION HELPERS
# =============================================================================

def validate_email(email: str) -> bool:
    """Basic email validation."""
    if not email or "@" not in email:
        return False
    
    parts = email.split("@")
    if len(parts) != 2:
        return False
    
    local, domain = parts
    if not local or not domain:
        return False
    
    if "." not in domain:
        return False
    
    return True


def validate_upn_domain(upn: str, expected_domain: str) -> bool:
    """Validate that UPN matches expected domain."""
    if not upn or "@" not in upn:
        return False
    
    domain = upn.split("@")[-1].lower()
    return domain == expected_domain.lower()


# =============================================================================
# USER INTERACTION HELPERS
# =============================================================================

def confirm_action(message: str, default: bool = False) -> bool:
    """Ask user to confirm an action."""
    default_str = "Y/n" if default else "y/N"
    response = input(f"  {message} ({default_str}): ").strip().lower()
    
    if not response:
        return default
    
    return response in ['y', 'yes']


def get_user_input(prompt: str, default: Optional[str] = None) -> str:
    """Get user input with optional default."""
    if default:
        response = input(f"  {prompt} [{default}]: ").strip()
        return response if response else default
    else:
        return input(f"  {prompt}: ").strip()


def get_numeric_input(
    prompt: str,
    min_val: int = 1,
    max_val: int = 1000,
    default: Optional[int] = None
) -> int:
    """Get numeric input from user with validation."""
    while True:
        if default:
            response = input(f"  {prompt} [{default}]: ").strip()
            if not response:
                return default
        else:
            response = input(f"  {prompt}: ").strip()
        
        try:
            value = int(response)
            if min_val <= value <= max_val:
                return value
            print_warning(f"Please enter a number between {min_val} and {max_val}")
        except ValueError:
            print_warning("Please enter a valid number")


def select_from_list(
    items: List[Any],
    prompt: str,
    display_func: Optional[callable] = None
) -> Optional[Any]:
    """Let user select an item from a list."""
    if not items:
        return None
    
    print()
    for i, item in enumerate(items, 1):
        if display_func:
            display = display_func(item)
        else:
            display = str(item)
        print(f"    [{i}] {display}")
    print()
    
    while True:
        response = input(f"  {prompt} (1-{len(items)}): ").strip()
        
        try:
            idx = int(response) - 1
            if 0 <= idx < len(items):
                return items[idx]
            print_warning(f"Please enter a number between 1 and {len(items)}")
        except ValueError:
            print_warning("Please enter a valid number")


# =============================================================================
# STATISTICS HELPERS
# =============================================================================

def calculate_statistics(values: List[float]) -> Dict[str, float]:
    """Calculate basic statistics for a list of values."""
    if not values:
        return {"min": 0, "max": 0, "avg": 0, "total": 0}
    
    return {
        "min": min(values),
        "max": max(values),
        "avg": sum(values) / len(values),
        "total": sum(values),
    }


def format_rate(count: int, seconds: float) -> str:
    """Format rate (items per second)."""
    if seconds <= 0:
        return "N/A"
    
    rate = count / seconds
    if rate >= 1:
        return f"{rate:.1f}/sec"
    else:
        return f"{rate * 60:.1f}/min"
