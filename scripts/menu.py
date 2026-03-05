#!/usr/bin/env python3
"""
SharePoint Sites Management - Main Menu

A unified interface for managing SharePoint sites:
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

import os
import subprocess
import sys
from pathlib import Path

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

def print_menu() -> None:
    """Print the main menu options."""
    print(f"  {Colors.WHITE}{Colors.BOLD}What would you like to do?{Colors.NC}")
    print()
    print(f"  {Colors.CYAN}╭──────────────────────────────────────────────────────────────╮{Colors.NC}")
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

def main() -> None:
    """Main entry point."""
    while True:
        clear_screen()
        print_logo()
        print_menu()
        
        choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip().lower()
        
        if choice == '1':
            # Create SharePoint Sites
            clear_screen()
            print()
            print(f"  {Colors.GREEN}{Colors.BOLD}🏗️  Create SharePoint Sites{Colors.NC}")
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
                run_script("deploy.py")
            elif sub_choice == '2':
                run_script("deploy.py", ["--random", "5", "--auto-approve"])
            elif sub_choice == '3':
                run_script("deploy.py", ["--random", "10", "--auto-approve"])
            elif sub_choice == '4':
                print()
                count = input(f"  {Colors.YELLOW}Number of sites (1-39):{Colors.NC} ").strip()
                if count.isdigit() and 1 <= int(count) <= 39:
                    run_script("deploy.py", ["--random", count])
                else:
                    print(f"  {Colors.RED}✗{Colors.NC} Invalid number. Must be between 1 and 39.")
                    input(f"  {Colors.YELLOW}Press Enter to continue...{Colors.NC}")
            elif sub_choice == '5':
                run_script("deploy.py", ["--config", str(SCRIPT_DIR.parent / "config" / "sites.json")])
            # else: back to menu
            
        elif choice == '2':
            # Populate Sites with Files
            clear_screen()
            print()
            print(f"  {Colors.BLUE}{Colors.BOLD}📄 Populate Sites with Files{Colors.NC}")
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
            print(f"  {Colors.RED}{Colors.BOLD}🗑️  Delete Files or Sites{Colors.NC}")
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
