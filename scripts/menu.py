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

def command_exists(command: str) -> bool:
    """Check if a command exists and actually works."""
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
    try:
        result = subprocess.run(
            ["az", "account", "show"],
            capture_output=True,
            text=True,
            check=True
        )
        return True
    except Exception:
        return False

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

def check_prerequisites(auto_install: bool = False) -> dict:
    """Check all prerequisites and optionally install missing ones."""
    results = {
        "python": {"installed": True, "version": f"Python {sys.version.split()[0]}"},
        "azure_cli": {"installed": False, "version": None},
        "terraform": {"installed": False, "version": None},
        "azure_login": {"logged_in": False}
    }
    
    # Check Azure CLI
    if command_exists("az"):
        results["azure_cli"]["installed"] = True
        results["azure_cli"]["version"] = get_command_version("az")
        # Check Azure login
        results["azure_login"]["logged_in"] = check_azure_login()
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
        else:
            print(f"  {Colors.YELLOW}⚠{Colors.NC} Azure Login: Not logged in")
    
    print()
    
    # Summary
    all_installed = (
        results["azure_cli"]["installed"] and 
        results["terraform"]["installed"]
    )
    
    if all_installed and results["azure_login"]["logged_in"]:
        print(f"  {Colors.GREEN}{Colors.BOLD}✓ All prerequisites met! Ready to proceed.{Colors.NC}")
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
            try:
                subprocess.run(["az", "login"], check=True)
                print()
                print_success("Azure login successful!")
            except FileNotFoundError:
                print_error("Azure CLI is not installed or not in PATH.")
                print_info("Please install Azure CLI: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli")
                print_info("After installation, restart your terminal and try again.")
            except subprocess.CalledProcessError as e:
                print_error(f"Azure login failed: {e}")
            except Exception as e:
                print_error(f"Azure login failed: {e}")
    
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
    print()
    print(f"  {Colors.YELLOW}{Colors.BOLD}Quick Commands:{Colors.NC}")
    print(f"  {Colors.DIM}────────────────{Colors.NC}")
    print(f"    {Colors.CYAN}python deploy.py --random 10{Colors.NC}     Create 10 random sites")
    print(f"    {Colors.CYAN}python populate_files.py --files 50{Colors.NC}  Add 50 files")
    print(f"    {Colors.CYAN}python cleanup.py --list-sites{Colors.NC}    List all sites")
    print(f"    {Colors.CYAN}python cleanup.py --select-files{Colors.NC}  Delete specific files")
    print()
    print(f"  {Colors.WHITE}For more information, see:{Colors.NC}")
    print(f"    • {Colors.BLUE}README.md{Colors.NC} - Main documentation")
    print(f"    • {Colors.BLUE}CONFIGURATION-GUIDE.md{Colors.NC} - Configuration details")
    print(f"    • {Colors.BLUE}docs/TROUBLESHOOTING.md{Colors.NC} - Common issues")
    print()
    input(f"  {Colors.YELLOW}Press Enter to return to menu...{Colors.NC}")

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
        subprocess.run(cmd, cwd=SCRIPT_DIR)
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
            # else: back to menu
            
        elif choice == '4':
            # List SharePoint Sites
            run_script("cleanup.py", ["--list-sites"])
            
        elif choice == '5':
            # List Files in Sites
            clear_screen()
            print()
            print(f"  {Colors.MAGENTA}{Colors.BOLD}📁 List Files in Sites{Colors.NC}")
            print(f"  {Colors.CYAN}{'─' * 40}{Colors.NC}")
            
            site_filter = get_site_filter()
            args = ["--list-files"]
            if site_filter:
                args.extend(["--site", site_filter])
            run_script("cleanup.py", args)
            
        elif choice == 'c':
            # Edit Configuration
            edit_configuration_menu()
            
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
