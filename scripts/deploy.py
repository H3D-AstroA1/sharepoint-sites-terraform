#!/usr/bin/env python3
"""
SharePoint Sites Deployment Script

This cross-platform Python script automates the deployment of SharePoint sites
using Terraform. It supports two modes:
1. Configuration File Mode - Use a JSON file with custom site names
2. Random Generation Mode - Generate sites with random names

Usage:
    python deploy.py                          # Interactive mode
    python deploy.py --config sites.json      # Use config file
    python deploy.py --random 10              # Generate 10 random sites (max 39)
    python deploy.py --help                   # Show help

Requirements:
    - Python 3.8+
    - Azure CLI (az) - Will be installed if missing
    - Terraform - Will be installed if missing
"""

import argparse
import json
import os
import platform
import random
import re
import subprocess
import sys
import urllib.request
import urllib.parse
import zipfile
import shutil
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Set

# ============================================================================
# CONFIGURATION
# ============================================================================

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_DIR = SCRIPT_DIR.parent
TERRAFORM_DIR = PROJECT_DIR / "terraform"
CONFIG_DIR = PROJECT_DIR / "config"
DEFAULT_CONFIG_FILE = CONFIG_DIR / "sites.json"
ENVIRONMENTS_FILE = CONFIG_DIR / "environments.json"
APP_CONFIG_FILE = SCRIPT_DIR / ".app_config.json"

# Realistic organizational department site templates with visibility settings
# Sites marked as "Public" are accessible to all employees
# Sites marked as "Private" are restricted to specific teams
DEPARTMENT_SITES = [
    # Executive & Leadership (Private - sensitive)
    {"name": "executive-leadership", "display_name": "Executive Leadership", "description": "Executive team communications, board materials, and strategic planning documents", "visibility": "Private", "template": "STS#3"},
    {"name": "board-of-directors", "display_name": "Board of Directors", "description": "Board meeting materials, governance documents, and director communications", "visibility": "Private", "template": "STS#3"},
    {"name": "senior-management", "display_name": "Senior Management", "description": "Senior leadership team collaboration and decision-making resources", "visibility": "Private", "template": "STS#3"},
    
    # Human Resources (Private - confidential employee data)
    {"name": "human-resources", "display_name": "Human Resources", "description": "HR policies, employee handbook, benefits information, and recruitment materials", "visibility": "Private", "template": "STS#3"},
    {"name": "hr-recruitment", "display_name": "HR Recruitment", "description": "Job postings, candidate tracking, and recruitment process documentation", "visibility": "Private", "template": "STS#3"},
    {"name": "hr-payroll-benefits", "display_name": "HR Payroll & Benefits", "description": "Payroll information, benefits enrollment, and compensation resources", "visibility": "Private", "template": "STS#3"},
    
    # Finance & Accounting (Private - financial data)
    {"name": "finance-department", "display_name": "Finance Department", "description": "Financial reports, budgets, forecasts, and accounting documentation", "visibility": "Private", "template": "STS#3"},
    {"name": "finance-accounts-payable", "display_name": "Finance Accounts Payable", "description": "Vendor payments, invoice processing, and AP documentation", "visibility": "Private", "template": "STS#3"},
    {"name": "finance-treasury", "display_name": "Finance Treasury", "description": "Cash management, banking relationships, and treasury operations", "visibility": "Private", "template": "STS#3"},
    
    # Claims (Private - sensitive claim data)
    {"name": "claims-department", "display_name": "Claims Department", "description": "Claims processing, case management, and claims documentation", "visibility": "Private", "template": "STS#3"},
    
    # Information Technology (Mixed)
    {"name": "it-department", "display_name": "IT Department", "description": "IT policies, system documentation, and technology resources", "visibility": "Private", "template": "STS#3"},
    {"name": "it-helpdesk", "display_name": "IT Helpdesk", "description": "IT support documentation, troubleshooting guides, and user assistance", "visibility": "Public", "template": "STS#3"},
    {"name": "it-security", "display_name": "IT Security", "description": "Cybersecurity policies, incident response, and security documentation", "visibility": "Private", "template": "STS#3"},
    
    # Legal & Compliance (Private - legal matters)
    {"name": "legal-department", "display_name": "Legal Department", "description": "Legal documents, contracts, and compliance materials", "visibility": "Private", "template": "STS#3"},
    {"name": "legal-compliance", "display_name": "Legal Compliance", "description": "Regulatory compliance, policy documentation, and audit materials", "visibility": "Private", "template": "STS#3"},
    
    # Marketing & Communications (Mixed)
    {"name": "marketing-department", "display_name": "Marketing Department", "description": "Marketing campaigns, brand assets, and promotional materials", "visibility": "Private", "template": "STS#3"},
    {"name": "marketing-brand", "display_name": "Marketing Brand", "description": "Brand guidelines, logos, templates, and brand assets", "visibility": "Public", "template": "STS#3"},
    {"name": "corporate-communications", "display_name": "Corporate Communications", "description": "Internal communications, press releases, and corporate messaging", "visibility": "Public", "template": "SITEPAGEPUBLISHING#0"},
    
    # Sales & Business Development (Private - client data)
    {"name": "sales-department", "display_name": "Sales Department", "description": "Sales resources, client information, and sales documentation", "visibility": "Private", "template": "STS#3"},
    {"name": "sales-enablement", "display_name": "Sales Enablement", "description": "Sales training, product information, and selling tools", "visibility": "Private", "template": "STS#3"},
    {"name": "customer-success", "display_name": "Customer Success", "description": "Customer onboarding, retention strategies, and success metrics", "visibility": "Private", "template": "STS#3"},
    
    # Operations (Mixed)
    {"name": "operations-department", "display_name": "Operations Department", "description": "Operational procedures, process documentation, and efficiency initiatives", "visibility": "Private", "template": "STS#3"},
    {"name": "operations-facilities", "display_name": "Operations Facilities", "description": "Facility management, maintenance, and workplace services", "visibility": "Public", "template": "STS#3"},
    
    # Product & Engineering (Private - IP)
    {"name": "product-management", "display_name": "Product Management", "description": "Product roadmaps, requirements, and product documentation", "visibility": "Private", "template": "STS#3"},
    {"name": "research-development", "display_name": "Research & Development", "description": "R&D projects, innovation initiatives, and research documentation", "visibility": "Private", "template": "STS#3"},
    {"name": "quality-assurance", "display_name": "Quality Assurance", "description": "Testing documentation, QA processes, and quality metrics", "visibility": "Private", "template": "STS#3"},
    
    # Customer Service (Private - customer data)
    {"name": "customer-service", "display_name": "Customer Service", "description": "Customer support resources, service documentation, and support tools", "visibility": "Private", "template": "STS#3"},
    
    # Project Management (Private)
    {"name": "project-management-office", "display_name": "Project Management Office", "description": "PMO resources, project templates, and methodology documentation", "visibility": "Private", "template": "STS#3"},
    
    # Risk & Audit (Private - sensitive)
    {"name": "risk-management", "display_name": "Risk Management", "description": "Risk assessments, mitigation strategies, and risk documentation", "visibility": "Private", "template": "STS#3"},
    {"name": "internal-audit", "display_name": "Internal Audit", "description": "Audit reports, compliance reviews, and audit documentation", "visibility": "Private", "template": "STS#3"},
    
    # Company-wide Resources (Public - all employees)
    {"name": "company-announcements", "display_name": "Company Announcements", "description": "Company-wide news, announcements, and organizational updates", "visibility": "Public", "template": "SITEPAGEPUBLISHING#0"},
    {"name": "employee-intranet", "display_name": "Employee Intranet", "description": "Central hub for all employees - news, resources, and company information", "visibility": "Public", "template": "SITEPAGEPUBLISHING#0"},
    {"name": "employee-recognition", "display_name": "Employee Recognition", "description": "Employee awards, achievements, and recognition programs", "visibility": "Public", "template": "SITEPAGEPUBLISHING#0"},
    {"name": "training-development", "display_name": "Training & Development", "description": "Employee training materials, learning resources, and professional development", "visibility": "Public", "template": "STS#3"},
    {"name": "onboarding-portal", "display_name": "Onboarding Portal", "description": "New employee onboarding, orientation materials, and welcome resources", "visibility": "Public", "template": "SITEPAGEPUBLISHING#0"},
    {"name": "health-safety", "display_name": "Health & Safety", "description": "Workplace safety, health policies, and HSE documentation", "visibility": "Public", "template": "STS#3"},
    {"name": "policies-procedures", "display_name": "Policies & Procedures", "description": "Company policies, standard procedures, and compliance documentation", "visibility": "Public", "template": "STS#3"},
    {"name": "templates-forms", "display_name": "Templates & Forms", "description": "Document templates, forms library, and standardized documents", "visibility": "Public", "template": "STS#3"},
    {"name": "social-committee", "display_name": "Social Committee", "description": "Company events, social activities, and team building initiatives", "visibility": "Public", "template": "STS#3"},
    {"name": "diversity-inclusion", "display_name": "Diversity & Inclusion", "description": "D&I initiatives, employee resource groups, and inclusion programs", "visibility": "Public", "template": "SITEPAGEPUBLISHING#0"},
]

# ============================================================================
# AD-HOC / USER-CREATED SITES
# ============================================================================
# These sites simulate organic, user-created sites that employees might create
# for projects, teams, events, working groups, and regional offices.
# They represent the "messy" reality of SharePoint usage in organizations.
# ============================================================================

ADHOC_SITES = [
    # Project Sites (Private - project teams)
    {"name": "q4-product-launch-2024", "display_name": "Q4 Product Launch 2024", "description": "Planning and coordination for Q4 2024 product launch", "visibility": "Private", "template": "STS#3"},
    {"name": "website-redesign-project", "display_name": "Website Redesign Project", "description": "Corporate website redesign initiative", "visibility": "Private", "template": "STS#3"},
    {"name": "crm-migration-team", "display_name": "CRM Migration Team", "description": "Salesforce to Dynamics 365 migration project", "visibility": "Private", "template": "STS#3"},
    {"name": "office-relocation-planning", "display_name": "Office Relocation Planning", "description": "HQ office move coordination and planning", "visibility": "Private", "template": "STS#3"},
    {"name": "erp-implementation", "display_name": "ERP Implementation", "description": "SAP implementation project team", "visibility": "Private", "template": "STS#3"},
    {"name": "mobile-app-development", "display_name": "Mobile App Development", "description": "Customer mobile app development project", "visibility": "Private", "template": "STS#3"},
    {"name": "data-warehouse-project", "display_name": "Data Warehouse Project", "description": "Enterprise data warehouse modernization", "visibility": "Private", "template": "STS#3"},
    {"name": "cloud-migration-2024", "display_name": "Cloud Migration 2024", "description": "Azure cloud migration initiative", "visibility": "Private", "template": "STS#3"},
    {"name": "process-automation-team", "display_name": "Process Automation Team", "description": "RPA and workflow automation project", "visibility": "Private", "template": "STS#3"},
    {"name": "customer-portal-redesign", "display_name": "Customer Portal Redesign", "description": "Customer self-service portal project", "visibility": "Private", "template": "STS#3"},
    
    # Working Groups & Committees (Mixed visibility)
    {"name": "innovation-lab", "display_name": "Innovation Lab", "description": "Cross-functional innovation and ideation team", "visibility": "Private", "template": "STS#3"},
    {"name": "sustainability-committee", "display_name": "Sustainability Committee", "description": "Environmental sustainability initiatives", "visibility": "Public", "template": "STS#3"},
    {"name": "digital-transformation", "display_name": "Digital Transformation", "description": "Company-wide digital transformation working group", "visibility": "Private", "template": "STS#3"},
    {"name": "employee-engagement-team", "display_name": "Employee Engagement Team", "description": "Employee satisfaction and engagement initiatives", "visibility": "Private", "template": "STS#3"},
    {"name": "cost-reduction-taskforce", "display_name": "Cost Reduction Taskforce", "description": "Operational efficiency and cost optimization", "visibility": "Private", "template": "STS#3"},
    {"name": "merger-integration-team", "display_name": "Merger Integration Team", "description": "Post-merger integration planning", "visibility": "Private", "template": "STS#3"},
    {"name": "compliance-working-group", "display_name": "Compliance Working Group", "description": "Cross-departmental compliance coordination", "visibility": "Private", "template": "STS#3"},
    {"name": "vendor-management-team", "display_name": "Vendor Management Team", "description": "Strategic vendor relationships and contracts", "visibility": "Private", "template": "STS#3"},
    
    # Social & Interest Groups (Public - employee communities)
    {"name": "coffee-club", "display_name": "Coffee Club", "description": "Coffee enthusiasts and office coffee machine updates", "visibility": "Public", "template": "STS#3"},
    {"name": "book-club", "display_name": "Book Club", "description": "Monthly book discussions and recommendations", "visibility": "Public", "template": "STS#3"},
    {"name": "running-club", "display_name": "Running Club", "description": "Lunchtime running group and race events", "visibility": "Public", "template": "STS#3"},
    {"name": "photography-club", "display_name": "Photography Club", "description": "Photography enthusiasts and photo sharing", "visibility": "Public", "template": "STS#3"},
    {"name": "gaming-community", "display_name": "Gaming Community", "description": "Video game enthusiasts and gaming events", "visibility": "Public", "template": "STS#3"},
    {"name": "parents-network", "display_name": "Parents Network", "description": "Working parents support and resources", "visibility": "Public", "template": "STS#3"},
    {"name": "pet-lovers", "display_name": "Pet Lovers", "description": "Pet photos and pet-friendly workplace initiatives", "visibility": "Public", "template": "STS#3"},
    {"name": "volunteer-corps", "display_name": "Volunteer Corps", "description": "Community volunteering and charity events", "visibility": "Public", "template": "STS#3"},
    
    # Employee Resource Groups (Public - D&I)
    {"name": "women-in-tech", "display_name": "Women in Tech", "description": "Women in technology employee resource group", "visibility": "Public", "template": "STS#3"},
    {"name": "pride-network", "display_name": "Pride Network", "description": "LGBTQ+ employee resource group", "visibility": "Public", "template": "STS#3"},
    {"name": "veterans-network", "display_name": "Veterans Network", "description": "Military veterans employee resource group", "visibility": "Public", "template": "STS#3"},
    {"name": "young-professionals", "display_name": "Young Professionals", "description": "Early career professionals network", "visibility": "Public", "template": "STS#3"},
    {"name": "accessibility-advocates", "display_name": "Accessibility Advocates", "description": "Disability inclusion and accessibility", "visibility": "Public", "template": "STS#3"},
    {"name": "multicultural-network", "display_name": "Multicultural Network", "description": "Cultural diversity and inclusion", "visibility": "Public", "template": "STS#3"},
    
    # Event Sites (Public - company events)
    {"name": "annual-company-retreat-2024", "display_name": "Annual Company Retreat 2024", "description": "Planning for the 2024 company retreat", "visibility": "Public", "template": "STS#3"},
    {"name": "hackathon-2024", "display_name": "Hackathon 2024", "description": "Annual innovation hackathon event", "visibility": "Public", "template": "STS#3"},
    {"name": "customer-summit-planning", "display_name": "Customer Summit Planning", "description": "Annual customer conference planning", "visibility": "Private", "template": "STS#3"},
    {"name": "holiday-party-2024", "display_name": "Holiday Party 2024", "description": "End of year celebration planning", "visibility": "Public", "template": "STS#3"},
    {"name": "sales-kickoff-2025", "display_name": "Sales Kickoff 2025", "description": "Annual sales kickoff event planning", "visibility": "Private", "template": "STS#3"},
    {"name": "town-hall-archives", "display_name": "Town Hall Archives", "description": "Recordings and materials from company town halls", "visibility": "Public", "template": "STS#3"},
    {"name": "charity-fundraiser", "display_name": "Charity Fundraiser", "description": "Annual charity fundraising campaign", "visibility": "Public", "template": "STS#3"},
    {"name": "wellness-week-2024", "display_name": "Wellness Week 2024", "description": "Employee wellness week activities", "visibility": "Public", "template": "STS#3"},
    
    # Regional & Office Sites (Mixed visibility)
    {"name": "london-office", "display_name": "London Office", "description": "London office team and local announcements", "visibility": "Public", "template": "STS#3"},
    {"name": "new-york-team", "display_name": "New York Team", "description": "New York office collaboration space", "visibility": "Public", "template": "STS#3"},
    {"name": "apac-region", "display_name": "APAC Region", "description": "Asia-Pacific regional team", "visibility": "Private", "template": "STS#3"},
    {"name": "emea-operations", "display_name": "EMEA Operations", "description": "Europe, Middle East & Africa operations", "visibility": "Private", "template": "STS#3"},
    {"name": "remote-workers-hub", "display_name": "Remote Workers Hub", "description": "Resources and community for remote employees", "visibility": "Public", "template": "STS#3"},
    {"name": "singapore-office", "display_name": "Singapore Office", "description": "Singapore office team collaboration", "visibility": "Public", "template": "STS#3"},
    {"name": "sydney-team", "display_name": "Sydney Team", "description": "Sydney office local team site", "visibility": "Public", "template": "STS#3"},
    {"name": "berlin-hub", "display_name": "Berlin Hub", "description": "Berlin office and DACH region", "visibility": "Public", "template": "STS#3"},
    
    # Miscellaneous User-Created Sites
    {"name": "lunch-orders", "display_name": "Lunch Orders", "description": "Weekly lunch order coordination", "visibility": "Public", "template": "STS#3"},
    {"name": "parking-coordination", "display_name": "Parking Coordination", "description": "Office parking space management", "visibility": "Public", "template": "STS#3"},
    {"name": "office-plants", "display_name": "Office Plants", "description": "Office plant care and watering schedule", "visibility": "Public", "template": "STS#3"},
    {"name": "lost-and-found", "display_name": "Lost and Found", "description": "Office lost and found items", "visibility": "Public", "template": "STS#3"},
    {"name": "carpool-network", "display_name": "Carpool Network", "description": "Employee carpooling coordination", "visibility": "Public", "template": "STS#3"},
    {"name": "office-supplies-requests", "display_name": "Office Supplies Requests", "description": "Office supply requests and inventory", "visibility": "Public", "template": "STS#3"},
    {"name": "meeting-room-tips", "display_name": "Meeting Room Tips", "description": "Meeting room booking tips and AV guides", "visibility": "Public", "template": "STS#3"},
    {"name": "new-hire-buddies", "display_name": "New Hire Buddies", "description": "Buddy program for new employees", "visibility": "Public", "template": "STS#3"},
]

# ============================================================================
# CONSOLE OUTPUT HELPERS
# ============================================================================

class Colors:
    """ANSI color codes for terminal output."""
    RED = '\033[0;31m'
    GREEN = '\033[0;32m'
    YELLOW = '\033[1;33m'
    BLUE = '\033[0;34m'
    CYAN = '\033[0;36m'
    WHITE = '\033[1;37m'
    BOLD = '\033[1m'
    DIM = '\033[2m'  # Dim/faint text
    NC = '\033[0m'  # No Color

    @classmethod
    def disable(cls):
        """Disable colors (for non-TTY output)."""
        cls.RED = cls.GREEN = cls.YELLOW = cls.BLUE = cls.CYAN = cls.WHITE = cls.BOLD = cls.DIM = cls.NC = ''


# Disable colors if not a TTY or on Windows without color support
if not sys.stdout.isatty():
    Colors.disable()
elif sys.platform == "win32":
    # Enable ANSI colors on Windows 10+
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
    except Exception:
        Colors.disable()


def print_banner(text: str) -> None:
    """Print a banner with the given text."""
    width = 80
    padding = (width - len(text) - 2) // 2
    line = "=" * width
    
    print()
    print(f"{Colors.CYAN}{line}{Colors.NC}")
    print(f"{Colors.CYAN}={'=' * padding} {text} {'=' * (width - padding - len(text) - 3)}={Colors.NC}")
    print(f"{Colors.CYAN}{line}{Colors.NC}")
    print()


def print_step(number: int, title: str) -> None:
    """Print a step header."""
    print()
    print(f"{Colors.YELLOW}+{'-' * 77}+{Colors.NC}")
    print(f"{Colors.YELLOW}| STEP {number}: {title}{' ' * (69 - len(title))}|{Colors.NC}")
    print(f"{Colors.YELLOW}+{'-' * 77}+{Colors.NC}")
    print()


def print_info(message: str) -> None:
    """Print an info message."""
    print(f"  {Colors.CYAN}[INFO]{Colors.NC} {message}")


def print_success(message: str) -> None:
    """Print a success message."""
    print(f"  {Colors.GREEN}[SUCCESS]{Colors.NC} {message}")


def print_warning(message: str) -> None:
    """Print a warning message."""
    print(f"  {Colors.YELLOW}[WARNING]{Colors.NC} {message}")


def print_error(message: str) -> None:
    """Print an error message."""
    print(f"  {Colors.RED}[ERROR]{Colors.NC} {message}")


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

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

def run_command(command: List[str], capture_output: bool = True, check: bool = True) -> subprocess.CompletedProcess:
    """Run a shell command and return the result."""
    # Resolve Azure CLI path if command starts with 'az'
    if command and command[0] == 'az':
        command = [find_azure_cli_path()] + command[1:]
    
    try:
        result = subprocess.run(
            command,
            capture_output=capture_output,
            text=True,
            check=check
        )
        return result
    except subprocess.CalledProcessError as e:
        if capture_output:
            print_error(f"Command failed: {' '.join(command)}")
            if e.stderr:
                print_error(f"Error: {e.stderr}")
        raise


def command_exists(command: str) -> bool:
    """Check if a command exists in the system PATH."""
    # Special handling for Azure CLI
    if command == "az":
        az_path = find_azure_cli_path()
        if az_path and az_path != "az":
            # Verify it actually works
            try:
                subprocess.run([az_path, "--version"], capture_output=True, text=True, timeout=10)
                return True
            except Exception:
                pass
        # Fall through to standard check
    
    try:
        if sys.platform == "win32":
            result = subprocess.run(["where", command], capture_output=True, text=True)
        else:
            result = subprocess.run(["which", command], capture_output=True, text=True)
        return result.returncode == 0
    except Exception:
        return False


def get_command_version(command: str) -> str:
    """Get the version of a command."""
    try:
        result = subprocess.run([command, '--version'], capture_output=True, text=True, check=False)
        if result.returncode == 0:
            return result.stdout.split('\n')[0]
        return "unknown"
    except Exception:
        return "unknown"


def prompt_input(prompt: str, default: str = "", required: bool = False, 
                 validation_pattern: str = "") -> str:
    """Prompt the user for input with optional default and validation."""
    while True:
        if default:
            user_input = input(f"  {prompt} [{default}]: ").strip()
            if not user_input:
                user_input = default
        else:
            user_input = input(f"  {prompt}: ").strip()
        
        if required and not user_input:
            print_warning("This field is required. Please enter a value.")
            continue
        
        if validation_pattern and user_input:
            if not re.match(validation_pattern, user_input):
                print_warning("Invalid format. Please try again.")
                continue
        
        return user_input


def prompt_selection(prompt: str, max_value: int) -> int:
    """Prompt the user to select a number from 1 to max_value."""
    while True:
        try:
            selection = int(input(f"  {prompt} (1-{max_value}): "))
            if 1 <= selection <= max_value:
                return selection
            print_warning(f"Please enter a number between 1 and {max_value}.")
        except ValueError:
            print_warning("Please enter a valid number.")


def confirm(prompt: str) -> bool:
    """Ask the user for confirmation (y/n)."""
    response = input(f"  {prompt} (y/n): ").strip().lower()
    return response in ('y', 'yes')


# ============================================================================
# PREREQUISITE INSTALLATION FUNCTIONS
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


def install_azure_cli() -> bool:
    """Install Azure CLI based on the operating system."""
    os_type, arch = get_os_info()
    
    print_info("Installing Azure CLI...")
    
    try:
        if os_type == "windows":
            # Download and run the MSI installer
            print_info("Downloading Azure CLI installer for Windows...")
            msi_url = "https://aka.ms/installazurecliwindows"
            msi_path = Path(os.environ.get('TEMP', '/tmp')) / "AzureCLI.msi"
            
            urllib.request.urlretrieve(msi_url, msi_path)
            print_info("Running installer (this may take a few minutes)...")
            
            # Run MSI installer silently
            subprocess.run(
                ["msiexec", "/i", str(msi_path), "/quiet", "/norestart"],
                check=True
            )
            
            print_success("Azure CLI installed successfully!")
            print_warning("You may need to restart your terminal for changes to take effect.")
            return True
            
        elif os_type == "macos":
            # Check if Homebrew is installed
            if command_exists("brew"):
                print_info("Installing Azure CLI via Homebrew...")
                subprocess.run(["brew", "install", "azure-cli"], check=True)
                print_success("Azure CLI installed successfully!")
                return True
            else:
                print_error("Homebrew is not installed.")
                print_info("Please install Homebrew first: https://brew.sh")
                print_info("Then run: brew install azure-cli")
                return False
                
        else:  # Linux
            print_info("Installing Azure CLI for Linux...")
            # Use the official installation script
            subprocess.run(
                ["bash", "-c", "curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash"],
                check=True
            )
            print_success("Azure CLI installed successfully!")
            return True
            
    except subprocess.CalledProcessError as e:
        print_error(f"Failed to install Azure CLI: {e}")
        return False
    except Exception as e:
        print_error(f"Error during Azure CLI installation: {e}")
        return False


def install_terraform() -> bool:
    """Install Terraform based on the operating system."""
    os_type, arch = get_os_info()
    
    print_info("Installing Terraform...")
    
    try:
        if os_type == "windows":
            # Check for package managers
            if command_exists("winget"):
                print_info("Installing Terraform via winget...")
                subprocess.run(
                    ["winget", "install", "--id", "Hashicorp.Terraform", "-e", "--source", "winget"],
                    check=True
                )
                print_success("Terraform installed successfully!")
                return True
            elif command_exists("choco"):
                print_info("Installing Terraform via Chocolatey...")
                subprocess.run(["choco", "install", "terraform", "-y"], check=True)
                print_success("Terraform installed successfully!")
                return True
            else:
                # Manual installation
                print_info("Downloading Terraform manually...")
                
                # Determine architecture
                tf_arch = "amd64" if "64" in arch or arch == "AMD64" else "386"
                
                # Get latest version from API
                try:
                    with urllib.request.urlopen("https://api.github.com/repos/hashicorp/terraform/releases/latest") as response:
                        data = json.loads(response.read().decode())
                        version = data['tag_name'].lstrip('v')
                except Exception:
                    version = "1.6.0"  # Fallback version
                
                download_url = f"https://releases.hashicorp.com/terraform/{version}/terraform_{version}_windows_{tf_arch}.zip"
                zip_path = Path(os.environ.get('TEMP', '/tmp')) / "terraform.zip"
                install_dir = Path(os.environ.get('LOCALAPPDATA', '')) / "Programs" / "Terraform"
                
                print_info(f"Downloading Terraform {version}...")
                urllib.request.urlretrieve(download_url, zip_path)
                
                # Extract
                install_dir.mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(install_dir)
                
                # Add to PATH
                current_path = os.environ.get('PATH', '')
                if str(install_dir) not in current_path:
                    print_info(f"Adding {install_dir} to PATH...")
                    subprocess.run(
                        ['setx', 'PATH', f'{current_path};{install_dir}'],
                        check=True,
                        capture_output=True
                    )
                    os.environ['PATH'] = f"{current_path};{install_dir}"
                
                print_success("Terraform installed successfully!")
                print_warning("You may need to restart your terminal for PATH changes to take effect.")
                return True
                
        elif os_type == "macos":
            if command_exists("brew"):
                print_info("Installing Terraform via Homebrew...")
                subprocess.run(["brew", "tap", "hashicorp/tap"], check=False)
                subprocess.run(["brew", "install", "hashicorp/tap/terraform"], check=True)
                print_success("Terraform installed successfully!")
                return True
            else:
                print_error("Homebrew is not installed.")
                print_info("Please install Homebrew first: https://brew.sh")
                print_info("Then run: brew install hashicorp/tap/terraform")
                return False
                
        else:  # Linux
            print_info("Installing Terraform for Linux...")
            
            # Add HashiCorp GPG key and repository
            commands = [
                "wget -O- https://apt.releases.hashicorp.com/gpg | sudo gpg --dearmor -o /usr/share/keyrings/hashicorp-archive-keyring.gpg",
                'echo "deb [signed-by=/usr/share/keyrings/hashicorp-archive-keyring.gpg] https://apt.releases.hashicorp.com $(lsb_release -cs) main" | sudo tee /etc/apt/sources.list.d/hashicorp.list',
                "sudo apt update && sudo apt install terraform -y"
            ]
            
            for cmd in commands:
                subprocess.run(["bash", "-c", cmd], check=True)
            
            print_success("Terraform installed successfully!")
            return True
            
    except subprocess.CalledProcessError as e:
        print_error(f"Failed to install Terraform: {e}")
        return False
    except Exception as e:
        print_error(f"Error during Terraform installation: {e}")
        return False


def check_and_install_prerequisites() -> bool:
    """Check for prerequisites and offer to install if missing."""
    print_step(1, "Checking Prerequisites")
    
    all_met = True
    missing = []
    azure_logged_in = False
    
    # Check Python version
    python_version = sys.version_info
    if python_version >= (3, 8):
        print_success(f"Python {python_version.major}.{python_version.minor}.{python_version.micro} - OK")
    else:
        print_error(f"Python 3.8+ required (current: {python_version.major}.{python_version.minor})")
        print_info("Please upgrade Python from: https://www.python.org/downloads/")
        all_met = False
    
    # Check Azure CLI
    if command_exists('az'):
        version = get_command_version('az')
        print_success(f"Azure CLI - OK ({version})")
        
        # Check Azure login status and show account info
        if check_azure_login():
            azure_logged_in = True
            account_info = get_azure_account_info()
            if account_info:
                user_name = account_info.get('user', {}).get('name', 'Unknown')
                subscription_name = account_info.get('name', 'Unknown')
                subscription_id = account_info.get('id', 'Unknown')
                tenant_id = account_info.get('tenantId', 'Unknown')
                print_success(f"Azure Login - Authenticated")
                print(f"  {Colors.CYAN}├─ User:{Colors.NC} {user_name}")
                print(f"  {Colors.CYAN}├─ Subscription:{Colors.NC} {subscription_name}")
                print(f"  {Colors.CYAN}├─ Subscription ID:{Colors.NC} {subscription_id}")
                print(f"  {Colors.CYAN}└─ Tenant ID:{Colors.NC} {tenant_id}")
            else:
                print_success(f"Azure Login - Authenticated")
        else:
            print_warning("Azure Login - NOT LOGGED IN")
            print_info("  You will be prompted to login later in the deployment process.")
    else:
        print_warning("Azure CLI - NOT INSTALLED")
        missing.append('az')
        all_met = False
    
    # Check Terraform
    if command_exists('terraform'):
        version = get_command_version('terraform')
        print_success(f"Terraform - OK ({version})")
    else:
        print_warning("Terraform - NOT INSTALLED")
        missing.append('terraform')
        all_met = False
    
    # If prerequisites are missing, offer to install
    if missing:
        print()
        print_info("Some prerequisites are missing. Would you like to install them?")
        print()
        
        for tool in missing:
            tool_name = "Azure CLI" if tool == "az" else "Terraform"
            if confirm(f"Install {tool_name}?"):
                print()
                if tool == "az":
                    if install_azure_cli():
                        # Verify installation
                        if command_exists('az'):
                            print_success("Azure CLI verified!")
                            all_met = True if 'terraform' not in missing or command_exists('terraform') else all_met
                        else:
                            print_warning("Azure CLI installed but not in PATH. Please restart your terminal.")
                    else:
                        print_error("Failed to install Azure CLI.")
                        print_info("Please install manually from: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli")
                else:
                    if install_terraform():
                        # Verify installation
                        if command_exists('terraform'):
                            print_success("Terraform verified!")
                            all_met = True if 'az' not in missing or command_exists('az') else all_met
                        else:
                            print_warning("Terraform installed but not in PATH. Please restart your terminal.")
                    else:
                        print_error("Failed to install Terraform.")
                        print_info("Please install manually from: https://www.terraform.io/downloads")
            else:
                print()
                print_info(f"Skipping {tool_name} installation.")
                if tool == "az":
                    print_info("Install manually from: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli")
                else:
                    print_info("Install manually from: https://www.terraform.io/downloads")
    
    return all_met


# ============================================================================
# SITE GENERATION FUNCTIONS
# ============================================================================

def read_sites_from_config(config_path: Path) -> List[Dict]:
    """Read site definitions from a JSON configuration file."""
    if not config_path.exists():
        raise FileNotFoundError(f"Configuration file not found: {config_path}")
    
    with open(config_path, 'r') as f:
        config = json.load(f)
    
    sites = config.get('sites', [])
    if not sites:
        raise ValueError("No sites defined in configuration file")
    
    return sites


def get_department_site_templates(config_path: Optional[Path] = None, warn_on_error: bool = False) -> List[Dict]:
    """Get department site templates using config baseline plus hardcoded extras.

    Baseline source is config/sites.json (or provided config path). Any additional
    hardcoded department templates not present in the config baseline are appended.
    Matching is done by site name (case-insensitive), and config entries win.
    """
    template_path = config_path or DEFAULT_CONFIG_FILE
    baseline_sites: List[Dict] = []

    if template_path.exists():
        try:
            baseline_sites = read_sites_from_config(template_path)
        except Exception as e:
            if warn_on_error:
                print_warning(f"Could not read department baseline from {template_path}: {e}")
                print_info("Falling back to built-in department templates only")

    merged_sites: List[Dict] = []
    seen_names: Set[str] = set()

    def add_site_template(template: Dict) -> None:
        name = str(template.get("name", "")).strip()
        display_name = str(template.get("display_name", "")).strip()
        description = str(template.get("description", "")).strip()

        if not name or not display_name or not description:
            return

        key = name.lower()
        if key in seen_names:
            return

        merged_sites.append({
            "name": name,
            "display_name": display_name,
            "description": description,
            "template": template.get("template", "STS#3"),
            "visibility": template.get("visibility", "Private"),
            "owners": template.get("owners", []),
            "members": template.get("members", []),
        })
        seen_names.add(key)

    for site in baseline_sites:
        add_site_template(site)

    for site in DEPARTMENT_SITES:
        add_site_template(site)

    return merged_sites


def generate_random_sites(count: int, department_templates: Optional[List[Dict]] = None) -> List[Dict]:
    """Generate a list of realistic organizational department sites.
    
    The maximum number of unique sites is limited to the number of templates
    in DEPARTMENT_SITES (currently 39). Each site includes realistic visibility
    settings (Private or Public) and appropriate templates.
    """
    templates = department_templates if department_templates is not None else get_department_site_templates()
    max_sites = len(templates)
    
    if count > max_sites:
        print_warning(f"Requested {count} sites, but only {max_sites} unique templates available.")
        print_info(f"Generating {max_sites} sites instead.")
        count = max_sites
    
    # Shuffle the department sites to get random selection
    available_sites = templates.copy()
    random.shuffle(available_sites)
    
    # Generate sites from the shuffled list
    sites = []
    for i in range(count):
        site_template = available_sites[i]
        
        # Create a copy with all required fields, preserving visibility and template
        site = {
            'name': site_template['name'],
            'display_name': site_template['display_name'],
            'description': site_template['description'],
            'template': site_template.get('template', 'STS#3'),
            'visibility': site_template.get('visibility', 'Private'),
            'owners': [],
            'members': []
        }
        
        sites.append(site)
    
    return sites


def generate_adhoc_sites(count: int) -> List[Dict]:
    """Generate a list of realistic ad-hoc/user-created sites.
    
    These sites simulate organic, user-created sites that employees might create
    for projects, teams, events, working groups, and regional offices.
    
    The maximum number of unique sites is limited to the number of templates
    in ADHOC_SITES (currently 60+). Each site includes realistic visibility
    settings (Private or Public) and appropriate templates.
    """
    max_sites = len(ADHOC_SITES)
    
    if count > max_sites:
        print_warning(f"Requested {count} sites, but only {max_sites} unique ad-hoc templates available.")
        print_info(f"Generating {max_sites} sites instead.")
        count = max_sites
    
    # Shuffle the ad-hoc sites to get random selection
    available_sites = ADHOC_SITES.copy()
    random.shuffle(available_sites)
    
    # Generate sites from the shuffled list
    sites = []
    for i in range(count):
        site_template = available_sites[i]
        
        # Create a copy with all required fields, preserving visibility and template
        site = {
            'name': site_template['name'],
            'display_name': site_template['display_name'],
            'description': site_template['description'],
            'template': site_template.get('template', 'STS#3'),
            'visibility': site_template.get('visibility', 'Private'),
            'owners': [],
            'members': []
        }
        
        sites.append(site)
    
    return sites


def generate_mixed_sites(
    dept_count: int,
    adhoc_count: int,
    department_templates: Optional[List[Dict]] = None,
) -> List[Dict]:
    """Generate a mix of department sites and ad-hoc sites.
    
    This creates a more realistic environment with both official department
    sites and organic user-created sites.
    
    Args:
        dept_count: Number of department sites to generate
        adhoc_count: Number of ad-hoc sites to generate
    
    Returns:
        Combined list of sites
    """
    dept_sites = generate_random_sites(dept_count, department_templates)
    adhoc_sites = generate_adhoc_sites(adhoc_count)
    
    # Combine and shuffle for a more natural mix
    all_sites = dept_sites + adhoc_sites
    random.shuffle(all_sites)
    
    return all_sites


# ============================================================================
# AZURE AD DISCOVERY FUNCTIONS
# ============================================================================

def load_app_config() -> Optional[Dict]:
    """Load the custom app configuration from file."""
    if APP_CONFIG_FILE.exists():
        try:
            with open(APP_CONFIG_FILE, 'r') as f:
                return json.load(f)
        except Exception:
            pass
    return None


def get_graph_access_token_via_client_credentials(app_config: Dict) -> Optional[str]:
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
    except Exception:
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
    try:
        result = run_command([
            "az", "account", "get-access-token",
            "--resource", "https://graph.microsoft.com",
            "--query", "accessToken",
            "-o", "tsv"
        ])
        if result:
            return result.stdout.strip()
    except Exception:
        pass
    return None


def discover_azure_ad_users(access_token: str, max_users: int = 100) -> List[Dict]:
    """Discover Azure AD users from the tenant.
    
    Returns a list of users with their UPN (userPrincipalName) and display name.
    Filters out guest users and service accounts.
    """
    users = []
    
    try:
        # Get users, filtering for member users (not guests)
        url = f"https://graph.microsoft.com/v1.0/users?$filter=userType eq 'Member'&$select=id,userPrincipalName,displayName,department,jobTitle&$top={max_users}"
        
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            result = json.loads(response.read().decode())
            
            for user in result.get('value', []):
                upn = user.get('userPrincipalName', '')
                # Skip service accounts and system accounts
                if upn and not upn.startswith('sync_') and not upn.startswith('admin_') and '#EXT#' not in upn:
                    users.append({
                        'id': user.get('id'),
                        'upn': upn,
                        'displayName': user.get('displayName', ''),
                        'department': user.get('department', ''),
                        'jobTitle': user.get('jobTitle', '')
                    })
    except Exception as e:
        print_warning(f"Failed to discover Azure AD users: {e}")
    
    return users


def discover_azure_ad_groups(access_token: str, max_groups: int = 50) -> List[Dict]:
    """Discover Azure AD groups from the tenant.
    
    Returns a list of security groups and M365 groups.
    Filters out dynamic groups and system groups.
    """
    groups = []
    
    try:
        # Get groups - both security groups and M365 groups
        url = f"https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,groupTypes,securityEnabled,mailEnabled&$top={max_groups}"
        
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            result = json.loads(response.read().decode())
            
            for group in result.get('value', []):
                display_name = group.get('displayName', '')
                group_types = group.get('groupTypes', [])
                
                # Skip dynamic groups and system groups
                if 'DynamicMembership' in group_types:
                    continue
                if display_name.startswith('All ') or display_name.startswith('System'):
                    continue
                
                # Determine group type
                is_m365 = 'Unified' in group_types
                is_security = group.get('securityEnabled', False)
                
                groups.append({
                    'id': group.get('id'),
                    'displayName': display_name,
                    'description': group.get('description', ''),
                    'type': 'M365' if is_m365 else ('Security' if is_security else 'Distribution'),
                    'isM365': is_m365,
                    'isSecurity': is_security
                })
    except Exception as e:
        print_warning(f"Failed to discover Azure AD groups: {e}")
    
    return groups


def assign_random_owners_members(
    sites: List[Dict],
    users: List[Dict],
    groups: List[Dict],
    min_owners: int = 1,
    max_owners: int = 3,
    min_members: int = 0,
    max_members: int = 5,
    include_groups: bool = True
) -> List[Dict]:
    """Assign random Azure AD users and groups as owners/members to sites.
    
    Args:
        sites: List of site dictionaries to modify
        users: List of discovered Azure AD users
        groups: List of discovered Azure AD groups
        min_owners: Minimum number of owners per site
        max_owners: Maximum number of owners per site
        min_members: Minimum number of members per site
        max_members: Maximum number of members per site
        include_groups: Whether to include groups as members
    
    Returns:
        Modified list of sites with owners and members assigned
    """
    if not users:
        print_warning("No users available for assignment")
        return sites
    
    # Get user UPNs for assignment
    user_upns = [u['upn'] for u in users]
    
    # Get group IDs for assignment (only security groups work well as members)
    group_ids = [g['id'] for g in groups if g.get('isSecurity')] if include_groups else []
    
    for site in sites:
        # Assign random owners (always users, not groups)
        num_owners = random.randint(min_owners, min(max_owners, len(user_upns)))
        site['owners'] = random.sample(user_upns, num_owners)
        
        # Assign random members (can be users or groups)
        num_members = random.randint(min_members, max_members)
        if num_members > 0:
            # Mix of users and groups
            available_members = user_upns.copy()
            if include_groups and group_ids:
                available_members.extend(group_ids)
            
            # Remove owners from potential members to avoid duplicates
            available_members = [m for m in available_members if m not in site['owners']]
            
            if available_members:
                num_members = min(num_members, len(available_members))
                site['members'] = random.sample(available_members, num_members)
            else:
                site['members'] = []
        else:
            site['members'] = []
    
    return sites


def select_owner_assignment_mode(sites: List[Dict], admin_email: str) -> List[Dict]:
    """Interactive menu to select how owners/members should be assigned to sites.
    
    Args:
        sites: List of site dictionaries
        admin_email: The SharePoint admin email (used as default owner)
    
    Returns:
        Modified list of sites with owners/members assigned
    """
    print()
    print(f"  {Colors.WHITE}How would you like to assign site owners and members?{Colors.NC}")
    print()
    print(f"  {Colors.CYAN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.NC}")
    print()
    print(f"    [1] Admin Only (Default)")
    print("        - Uses the SharePoint admin email as the sole owner")
    print("        - No additional members")
    print("        - Simplest option, good for testing")
    print()
    print(f"    [2] Discover Azure AD Users & Groups")
    print("        - Queries Microsoft Graph API for real users/groups")
    print("        - Randomly assigns users as owners (1-3 per site)")
    print("        - Randomly assigns users/groups as members (0-5 per site)")
    print("        - Most realistic option for production-like environments")
    print()
    print(f"    [3] Skip Owner Assignment")
    print("        - Leave owners/members empty")
    print("        - Sites will be created with default permissions")
    print()
    print(f"  {Colors.CYAN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.NC}")
    print()
    
    while True:
        choice_input = input(f"  Enter your choice (1-3): ").strip()
        try:
            choice = int(choice_input)
            if 1 <= choice <= 3:
                break
            print_warning("Please enter 1, 2, or 3.")
        except ValueError:
            print_warning("Please enter 1, 2, or 3.")
    
    if choice == 1:
        # Admin only - use the admin email as owner
        print()
        print_info("Using admin email as sole owner for all sites")
        for site in sites:
            site['owners'] = [admin_email]
            site['members'] = []
        print_success(f"Assigned {admin_email} as owner to all {len(sites)} sites")
        
    elif choice == 2:
        # Discover Azure AD users and groups
        print()
        print_info("Discovering Azure AD users and groups...")
        
        # Get access token
        access_token = get_graph_access_token()
        if not access_token:
            print_warning("Could not get access token. Falling back to admin-only mode.")
            print_info("Make sure you have set up the app registration via the main menu.")
            for site in sites:
                site['owners'] = [admin_email]
                site['members'] = []
            return sites
        
        # Discover users
        print_info("Querying Azure AD for users...")
        users = discover_azure_ad_users(access_token, max_users=100)
        print_success(f"Found {len(users)} users")
        
        # Discover groups
        print_info("Querying Azure AD for groups...")
        groups = discover_azure_ad_groups(access_token, max_groups=50)
        print_success(f"Found {len(groups)} groups")
        
        if not users:
            print_warning("No users found. Falling back to admin-only mode.")
            for site in sites:
                site['owners'] = [admin_email]
                site['members'] = []
            return sites
        
        # Show discovered users/groups summary
        print()
        print(f"  {Colors.WHITE}Discovered Azure AD Identities:{Colors.NC}")
        print()
        
        # Show sample users
        print(f"    {Colors.CYAN}Users (sample):{Colors.NC}")
        for user in users[:5]:
            dept = f" - {user['department']}" if user.get('department') else ""
            print(f"      • {user['displayName']} ({user['upn']}){dept}")
        if len(users) > 5:
            print(f"      ... and {len(users) - 5} more")
        
        # Show sample groups
        if groups:
            print()
            print(f"    {Colors.CYAN}Groups (sample):{Colors.NC}")
            for group in groups[:5]:
                print(f"      • {group['displayName']} ({group['type']})")
            if len(groups) > 5:
                print(f"      ... and {len(groups) - 5} more")
        
        print()
        
        # Ask about including groups as members
        include_groups = False
        if groups:
            include_groups_input = input("  Include groups as site members? (y/N): ").strip().lower()
            include_groups = include_groups_input == 'y'
        
        # Assign random owners/members
        print()
        print_info("Assigning random owners and members to sites...")
        sites = assign_random_owners_members(
            sites, users, groups,
            min_owners=1, max_owners=3,
            min_members=0, max_members=5,
            include_groups=include_groups
        )
        
        # Show assignment summary
        total_owners = sum(len(s.get('owners', [])) for s in sites)
        total_members = sum(len(s.get('members', [])) for s in sites)
        print_success(f"Assigned {total_owners} owners and {total_members} members across {len(sites)} sites")
        
    else:
        # Skip owner assignment
        print()
        print_info("Skipping owner assignment - sites will use default permissions")
        for site in sites:
            site['owners'] = []
            site['members'] = []
    
    return sites


def format_terraform_sites_block(sites: List[Dict]) -> str:
    """Format sites as a Terraform HCL block."""
    lines = ["sharepoint_sites = {"]
    
    for site in sites:
        lines.append(f'  "{site["name"]}" = {{')
        lines.append(f'    display_name = "{site.get("display_name", site["name"])}"')
        lines.append(f'    description  = "{site.get("description", "")}"')
        lines.append(f'    template     = "{site.get("template", "STS#3")}"')
        lines.append(f'    visibility   = "{site.get("visibility", "Private")}"')
        
        owners = site.get('owners', [])
        owners_str = ', '.join(f'"{o}"' for o in owners)
        lines.append(f'    owners       = [{owners_str}]')
        
        members = site.get('members', [])
        members_str = ', '.join(f'"{m}"' for m in members)
        lines.append(f'    members      = [{members_str}]')
        
        lines.append('  }')
    
    lines.append("}")
    return '\n'.join(lines)


def get_site_url_name(site_name: str, template: str) -> str:
    """Convert a site name to the URL segment used by the creator script."""
    url_name = ''.join(c for c in site_name if c.isalnum())
    if template == "SITEPAGEPUBLISHING#0":
        url_name = url_name + "site"
    return url_name


def get_sharepoint_site_url_candidates(site: Dict, m365_tenant: str) -> List[str]:
    """Return likely SharePoint URLs for a site across current/legacy creation paths."""
    site_name = site['name']
    template = site.get('template', 'STS#3')
    sanitized_name = ''.join(c for c in site_name if c.isalnum())

    candidates = [
        f"https://{m365_tenant}.sharepoint.com/sites/{site_name}",
        f"https://{m365_tenant}.sharepoint.com/sites/{sanitized_name}",
    ]

    if template == "SITEPAGEPUBLISHING#0":
        candidates.append(f"https://{m365_tenant}.sharepoint.com/sites/{sanitized_name}site")

    unique_candidates: List[str] = []
    seen: Set[str] = set()
    for candidate in candidates:
        normalized_candidate = candidate.rstrip('/')
        if normalized_candidate not in seen:
            seen.add(normalized_candidate)
            unique_candidates.append(normalized_candidate)
    return unique_candidates


def test_sharepoint_site_exists(site_url: str, m365_tenant: str, access_token: str) -> Optional[str]:
    """Return the resolved SharePoint site URL if Graph can find it, else None."""
    parsed_url = urllib.parse.urlparse(site_url)
    site_path = parsed_url.path.rstrip('/')
    if not site_path:
        return None

    request = urllib.request.Request(
        f"https://graph.microsoft.com/v1.0/sites/{m365_tenant}.sharepoint.com:{site_path}",
        headers={
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
    )

    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            payload = json.loads(response.read().decode('utf-8'))
            web_url = payload.get('webUrl')
            return web_url.rstrip('/') if web_url else site_url.rstrip('/')
    except Exception:
        return None


def resolve_sharepoint_site_url(site: Dict, m365_tenant: str, access_token: str) -> Optional[str]:
    """Resolve the actual SharePoint URL for a site by probing likely URL patterns."""
    for candidate_url in get_sharepoint_site_url_candidates(site, m365_tenant):
        resolved_url = test_sharepoint_site_exists(candidate_url, m365_tenant, access_token)
        if resolved_url:
            return resolved_url
    return None


def get_verified_sharepoint_site_urls(sites: List[Dict], m365_tenant: str) -> Dict[str, str]:
    """Return a map of site name to verified SharePoint URL for sites that currently exist."""
    access_token = get_graph_access_token()
    if not access_token:
        return {}

    verified_urls: Dict[str, str] = {}
    for site in sites:
        resolved_url = resolve_sharepoint_site_url(site, m365_tenant, access_token)
        if resolved_url:
            verified_urls[site['name']] = resolved_url
    return verified_urls


def get_terraform_state_resources() -> Set[str]:
    """Return all Terraform resource addresses currently tracked in state."""
    result = subprocess.run(
        ['terraform', 'state', 'list'],
        cwd=TERRAFORM_DIR,
        capture_output=True,
        text=True,
        check=False,
        env=get_terraform_env()
    )
    if result.returncode != 0:
        return set()
    return {line.strip() for line in result.stdout.splitlines() if line.strip()}


def remove_terraform_state_resource(resource_address: str) -> bool:
    """Remove a single resource address from Terraform state."""
    try:
        subprocess.run(
            ['terraform', 'state', 'rm', resource_address],
            cwd=TERRAFORM_DIR,
            check=True,
            env=get_terraform_env()
        )
        return True
    except subprocess.CalledProcessError:
        print_warning(f"  Failed to remove stale state: {resource_address}")
        return False


def reconcile_stale_sharepoint_site_state(sites: List[Dict], m365_tenant: str) -> List[str]:
    """Remove stale null_resource state entries for SharePoint sites that no longer exist."""
    access_token = get_graph_access_token()
    if not access_token:
        print_warning("Could not get a Microsoft Graph token to validate existing SharePoint sites.")
        return []

    state_resources = get_terraform_state_resources()
    if not state_resources:
        return []

    stale_resources: List[str] = []
    for site in sites:
        resource_address = f'null_resource.sharepoint_sites["{site["name"]}"]'
        if resource_address not in state_resources:
            continue

        resolved_url = resolve_sharepoint_site_url(site, m365_tenant, access_token)
        if resolved_url is None:
            stale_resources.append(resource_address)

    if not stale_resources:
        return []

    print_warning(
        f"Detected {len(stale_resources)} SharePoint site resource(s) in Terraform state that are missing in Microsoft 365."
    )
    print_info("Removing stale Terraform state so the missing sites will be recreated...")

    removed_resources: List[str] = []
    for resource_address in stale_resources:
        print_info(f"  Removing stale state: {resource_address}")
        if remove_terraform_state_resource(resource_address):
            removed_resources.append(resource_address)

    return removed_resources


# ============================================================================
# AZURE CLI FUNCTIONS
# ============================================================================

def check_azure_login() -> bool:
    """Check if the user is logged in to Azure CLI."""
    try:
        result = run_command(['az', 'account', 'show'], check=False)
        return result.returncode == 0
    except Exception:
        return False


def get_azure_account_info() -> Optional[Dict]:
    """Get the current Azure account information."""
    try:
        result = run_command(['az', 'account', 'show', '--output', 'json'], check=False)
        if result.returncode == 0:
            return json.loads(result.stdout)
        return None
    except Exception:
        return None


def azure_login() -> bool:
    """Perform Azure CLI login."""
    print_info("Opening browser for Azure login...")
    az_path = find_azure_cli_path()
    try:
        subprocess.run([az_path, 'login'], check=True)
        return True
    except FileNotFoundError:
        print_error("Azure CLI is not installed or not in PATH.")
        print_info("Please install Azure CLI: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli")
        print_info("Or run the main menu (menu.py) and use option [0] to install prerequisites.")
        return False
    except subprocess.CalledProcessError:
        return False


def get_azure_tenants() -> List[Dict]:
    """Get list of Azure tenants."""
    result = run_command([
        'az', 'account', 'tenant', 'list',
        '--query', '[].{tenantId:tenantId, displayName:displayName}',
        '-o', 'json'
    ])
    return json.loads(result.stdout)


def get_azure_subscriptions(tenant_id: str) -> List[Dict]:
    """Get list of Azure subscriptions for a tenant."""
    result = run_command([
        'az', 'account', 'list',
        '--query', f"[?tenantId=='{tenant_id}' && state=='Enabled'].{{id:id, name:name}}",
        '-o', 'json'
    ])
    return json.loads(result.stdout)


def set_azure_subscription(subscription_id: str) -> None:
    """Set the active Azure subscription."""
    run_command(['az', 'account', 'set', '--subscription', subscription_id])


def get_resource_groups() -> List[Dict]:
    """Get list of resource groups."""
    result = run_command([
        'az', 'group', 'list',
        '--query', '[].{name:name, location:location}',
        '-o', 'json'
    ])
    return json.loads(result.stdout)


# ============================================================================
# TERRAFORM FUNCTIONS
# ============================================================================

def get_terraform_env() -> dict:
    """Get environment variables for Terraform with Azure CLI path included."""
    env = os.environ.copy()
    
    # Find Azure CLI path and add its directory to PATH
    az_path = find_azure_cli_path()
    if az_path and az_path != "az":
        az_dir = os.path.dirname(az_path)
        current_path = env.get('PATH', '')
        if az_dir not in current_path:
            env['PATH'] = az_dir + os.pathsep + current_path
    
    return env

def terraform_init() -> bool:
    """Initialize Terraform."""
    print_info("Running terraform init...")
    try:
        subprocess.run(['terraform', 'init'], cwd=TERRAFORM_DIR, check=True, env=get_terraform_env())
        return True
    except subprocess.CalledProcessError:
        return False


def terraform_plan() -> bool:
    """Run Terraform plan."""
    print_info("Running terraform plan...")
    try:
        subprocess.run(['terraform', 'plan', '-out=tfplan'], cwd=TERRAFORM_DIR, check=True, env=get_terraform_env())
        return True
    except subprocess.CalledProcessError:
        return False


def _terraform_import_resource(resource_address: str, resource_id: str) -> bool:
    """Import a single existing resource into Terraform state."""
    print_info(f"  Importing: {resource_address}")
    try:
        subprocess.run(
            ['terraform', 'import', resource_address, resource_id],
            cwd=TERRAFORM_DIR, check=True, env=get_terraform_env()
        )
        return True
    except subprocess.CalledProcessError:
        print_warning(f"  Failed to import: {resource_address}")
        return False


def terraform_apply(auto_approve: bool = False) -> bool:
    """Run Terraform apply.

    On failure, automatically detects state drift ('already exists') errors,
    imports the conflicting resources, re-plans, and retries apply so the user
    never needs to run terraform import manually.
    """
    print_info("Running terraform apply...")
    env = get_terraform_env()

    cmd = ['terraform', 'apply']
    if auto_approve:
        cmd.append('-auto-approve')
    cmd.append('tfplan')

    # Stream output to the console while buffering for error analysis.
    # Use utf-8 explicitly so Terraform's box-drawing chars (╷ ╵ │) are
    # decoded correctly on Windows (which defaults to cp1252).
    output_lines: List[str] = []
    try:
        proc = subprocess.Popen(
            cmd, cwd=TERRAFORM_DIR, env=env,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            encoding='utf-8', errors='replace', bufsize=1
        )
        assert proc.stdout is not None
        for line in proc.stdout:
            print(line, end='', flush=True)
            output_lines.append(line)
        proc.wait()
    except Exception:
        return False

    if proc.returncode == 0:
        return True

    # ------------------------------------------------------------------ #
    # State drift detection: scan for "already exists" error blocks and   #
    # extract ALL (resource_id, resource_address) pairs in one pass.      #
    # Avoids relying on ╷/╵ block delimiters (encoding-fragile).          #
    # Each Terraform error emits the ID before the "with ..." address.    #
    # ------------------------------------------------------------------ #
    full_output = ''.join(output_lines)
    import_pairs: List[Tuple[str, str]] = []

    # Non-greedy DOTALL match pairs each ID with the nearest following address.
    raw_pairs = re.findall(
        r'A resource with the ID "([^"]+)" already exists.*?'
        r'with (azurerm_\S+),',
        full_output, re.DOTALL | re.IGNORECASE
    )
    import_pairs = [(addr, resource_id) for resource_id, addr in raw_pairs]

    if not import_pairs:
        # Not a state drift error — nothing we can auto-fix.
        return False

    print_warning(f"\nDetected {len(import_pairs)} pre-existing resource(s) not tracked in Terraform state.")
    print_info("Auto-importing into state — no manual action needed...")

    all_imported = all(
        _terraform_import_resource(addr, rid)
        for addr, rid in import_pairs
    )

    if not all_imported:
        print_error("One or more imports failed. Please resolve the remaining resources manually.")
        return False

    print_success(f"Imported {len(import_pairs)} resource(s) successfully.")

    # Re-plan so the plan reflects the updated state, then retry apply.
    print_info("Re-planning after import...")
    try:
        subprocess.run(
            ['terraform', 'plan', '-out=tfplan'],
            cwd=TERRAFORM_DIR, check=True, env=env
        )
    except subprocess.CalledProcessError:
        print_error("Re-plan after import failed.")
        return False

    print_info("Retrying terraform apply...")
    try:
        subprocess.run(
            ['terraform', 'apply', '-auto-approve', 'tfplan'],
            cwd=TERRAFORM_DIR, check=True, env=env
        )
        return True
    except subprocess.CalledProcessError:
        return False


def terraform_output() -> None:
    """Display Terraform outputs."""
    try:
        subprocess.run(['terraform', 'output', 'deployment_summary'], cwd=TERRAFORM_DIR, check=False, env=get_terraform_env())
    except Exception:
        pass


# ============================================================================
# ENVIRONMENT CONFIGURATION FUNCTIONS
# ============================================================================

def load_environments() -> Optional[Dict]:
    """Load pre-configured environments from environments.json."""
    if not ENVIRONMENTS_FILE.exists():
        return None
    
    try:
        with open(ENVIRONMENTS_FILE, 'r') as f:
            data = json.load(f)
        
        # Check if environments are configured (not just template values)
        environments = data.get('environments', [])
        configured_envs = []
        
        for env in environments:
            azure = env.get('azure', {})
            # Check if tenant_id is configured (not placeholder)
            tenant_id = azure.get('tenant_id', '')
            if tenant_id and not tenant_id.startswith('YOUR-') and not tenant_id.startswith('<<CHANGE'):
                configured_envs.append(env)
        
        if configured_envs:
            data['environments'] = configured_envs
            return data
        return None
    except Exception as e:
        print_warning(f"Could not load environments file: {e}")
        return None


def select_environment() -> Optional[Dict]:
    """Let user select from pre-configured environments or manual configuration."""
    env_data = load_environments()
    
    if not env_data:
        return None
    
    environments = env_data.get('environments', [])
    default_env = env_data.get('default_environment', '')
    
    print()
    print(f"  {Colors.WHITE}Pre-configured environments found!{Colors.NC}")
    print()
    print(f"  {Colors.CYAN}[0] Manual Configuration{Colors.NC}")
    print("      Enter tenant, subscription, and resource group manually")
    print()
    
    default_idx = 0
    for i, env in enumerate(environments, 1):
        name = env.get('name', f'Environment {i}')
        desc = env.get('description', '')
        azure = env.get('azure', {})
        tenant_name = azure.get('tenant_name', 'Unknown')
        sub_name = azure.get('subscription_name', 'Unknown')
        
        marker = ""
        if name == default_env:
            marker = f" {Colors.GREEN}(default){Colors.NC}"
            default_idx = i
        
        print(f"  {Colors.CYAN}[{i}] {name}{marker}{Colors.NC}")
        print(f"      {desc}")
        print(f"      Tenant: {tenant_name} | Subscription: {sub_name}")
        print()
    
    while True:
        try:
            default_str = f" (default: {default_idx})" if default_idx > 0 else ""
            choice = input(f"  Select environment [0-{len(environments)}]{default_str}: ").strip()
            
            if choice == "" and default_idx > 0:
                choice = str(default_idx)
            
            choice_int = int(choice)
            
            if choice_int == 0:
                return None  # Manual configuration
            elif 1 <= choice_int <= len(environments):
                selected = environments[choice_int - 1]
                print()
                print_success(f"Selected environment: {selected.get('name')}")
                return selected
            else:
                print_error(f"Please enter a number between 0 and {len(environments)}")
        except ValueError:
            print_error("Please enter a valid number")


def use_environment_config(env: Dict) -> Tuple[str, str, str, str, str, str, str]:
    """Extract configuration from selected environment.
    
    Returns:
        Tuple of (tenant_id, subscription_id, resource_group, location, m365_tenant, admin_email, key_vault_name)
    """
    azure = env.get('azure', {})
    m365 = env.get('m365', {})
    
    tenant_id = azure.get('tenant_id', '')
    subscription_id = azure.get('subscription_id', '')
    resource_group = azure.get('resource_group', '')
    location = azure.get('location', 'westus2')
    m365_tenant = m365.get('tenant_name', '')
    admin_email = m365.get('admin_email', '')
    
    # Key Vault name - empty string means auto-generate
    key_vault_name = azure.get('key_vault_name', '')
    # Clean up placeholder text
    if key_vault_name.startswith('<<'):
        key_vault_name = ''
    
    return tenant_id, subscription_id, resource_group, location, m365_tenant, admin_email, key_vault_name


def get_environment_by_name(name: str) -> Optional[Dict]:
    """Get a specific environment by name from environments.json."""
    env_data = load_environments()
    
    if not env_data:
        return None
    
    environments = env_data.get('environments', [])
    
    for env in environments:
        if env.get('name', '').lower() == name.lower():
            return env
    
    return None


# ============================================================================
# MAIN DEPLOYMENT LOGIC
# ============================================================================

def select_site_mode(args, step_num: int = 1) -> Tuple[str, List[Dict]]:
    """Select site generation mode and return sites."""
    print_step(step_num, "Select Site Generation Mode")
    
    sites = []
    mode = ""
    config_path = Path(args.config) if args.config else DEFAULT_CONFIG_FILE
    department_templates = get_department_site_templates(config_path, warn_on_error=True)
    department_template_count = len(department_templates)
    
    # Check command line arguments
    if args.random and args.random > 0:
        mode = "random"
        print_info("Using random generation mode (from command line)")
    elif args.config:
        mode = "config"
        print_info("Using configuration file mode (from command line)")
    else:
        # Interactive selection
        # Try to get config file site count for display
        config_site_count = 0
        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    config_data = json.load(f)
                    config_site_count = len(config_data.get('sites', []))
            except Exception:
                pass
        
        config_count_str = f" ({config_site_count} sites)" if config_site_count > 0 else " (edit to add sites)"
        
        print()
        print(f"  {Colors.WHITE}How would you like to define your SharePoint sites?{Colors.NC}")
        print()
        print(f"  {Colors.CYAN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.NC}")
        print(f"  {Colors.WHITE}CONFIGURATION FILE OPTIONS:{Colors.NC}")
        print()
        print(f"    [1] Use Configuration File Only{config_count_str}")
        print(f"        {Colors.BLUE}Source:{Colors.NC} config/sites.json")
        print("        - Full control over site names, descriptions, and settings")
        print()
        print(f"    [2] Configuration File + Ad-hoc Sites ({config_site_count} config + you choose ad-hoc)")
        print(f"        {Colors.BLUE}Source:{Colors.NC} config/sites.json + Ad-hoc templates")
        print(f"        - PLUS you choose how many ad-hoc sites (0-{len(ADHOC_SITES)} available)")
        print()
        print(f"  {Colors.CYAN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.NC}")
        print(f"  {Colors.WHITE}RANDOM GENERATION OPTIONS:{Colors.NC}")
        print()
        print(f"    [3] Generate Department Sites ({department_template_count} templates)")
        print(f"        {Colors.BLUE}Source:{Colors.NC} config baseline + department templates")
        print("        - Official department sites (HR, Finance, IT, Legal, etc.)")
        print()
        print(f"    [4] Generate Ad-hoc Sites ({len(ADHOC_SITES)} templates)")
        print(f"        {Colors.BLUE}Source:{Colors.NC} Ad-hoc templates")
        print("        - User-created sites (projects, teams, events, clubs)")
        print()
        print(f"    [5] Generate Mixed Sites (Department + Ad-hoc) {Colors.GREEN}(Recommended){Colors.NC}")
        print(f"        {Colors.BLUE}Source:{Colors.NC} Department + Ad-hoc templates")
        print("        - Combines both types for maximum realism")
        print()
        print(f"  {Colors.CYAN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.NC}")
        print()
        print(f"    {Colors.RED}[Q] Quit{Colors.NC}")
        print()
        
        while True:
            choice_input = input(f"  Enter your choice (1-5, Q to quit): ").strip().lower()
            if choice_input == 'q':
                print_warning("Deployment cancelled.")
                sys.exit(0)
            try:
                choice = int(choice_input)
                if 1 <= choice <= 5:
                    break
                print_warning("Please enter 1, 2, 3, 4, 5, or Q to quit.")
            except ValueError:
                print_warning("Please enter 1, 2, 3, 4, 5, or Q to quit.")
        
        if choice == 1:
            mode = "config"
        elif choice == 2:
            mode = "config_adhoc"
        elif choice == 3:
            mode = "random"
        elif choice == 4:
            mode = "adhoc"
        else:
            mode = "mixed"
    
    # Process based on mode
    if mode == "config":
        print()
        print_info("Configuration File Mode selected")
        
        config_path = Path(args.config) if args.config else DEFAULT_CONFIG_FILE
        
        if not config_path.exists():
            print_warning(f"Config file not found at: {config_path}")
            config_path = Path(prompt_input("Enter path to configuration file", required=True))
        
        print_info(f"Reading sites from: {config_path}")
        
        try:
            sites = read_sites_from_config(config_path)
            print_success(f"Loaded {len(sites)} sites from configuration file")
        except Exception as e:
            print_error(f"Failed to read configuration file: {e}")
            sys.exit(1)
    
    elif mode == "config_adhoc":
        print()
        print_info("Configuration File + Ad-hoc Sites Mode selected")
        print(f"  {Colors.CYAN}ℹ{Colors.NC} Your custom sites + random ad-hoc sites")
        print()
        
        # Load config file sites
        config_path = Path(args.config) if args.config else DEFAULT_CONFIG_FILE
        
        if not config_path.exists():
            print_warning(f"Config file not found at: {config_path}")
            config_path = Path(prompt_input("Enter path to configuration file", required=True))
        
        print_info(f"Reading sites from: {config_path}")
        
        try:
            config_sites = read_sites_from_config(config_path)
            print_success(f"Loaded {len(config_sites)} sites from configuration file")
        except Exception as e:
            print_error(f"Failed to read configuration file: {e}")
            sys.exit(1)
        
        # Ask for ad-hoc sites count
        print()
        max_adhoc = len(ADHOC_SITES)
        while True:
            try:
                adhoc_count = int(prompt_input(f"How many ad-hoc sites to add? (0-{max_adhoc})", required=True))
                if 0 <= adhoc_count <= max_adhoc:
                    break
                print_warning(f"Please enter a number between 0 and {max_adhoc}")
            except ValueError:
                print_warning("Please enter a valid number")
        
        if adhoc_count > 0:
            adhoc_sites = generate_adhoc_sites(adhoc_count)
            sites = config_sites + adhoc_sites
            random.shuffle(sites)  # Mix them up for realism
            print_success(f"Combined {len(config_sites)} config sites + {adhoc_count} ad-hoc sites = {len(sites)} total")
        else:
            sites = config_sites
            print_info("No ad-hoc sites added, using config file only")
    
    elif mode == "random":
        print()
        print_info("Department Sites Mode selected")
        print(f"  {Colors.CYAN}ℹ{Colors.NC} These are official department sites (HR, Finance, IT, etc.)")
        print()
        
        max_sites = department_template_count
        count = args.random if args.random and args.random > 0 else 0
        if count == 0:
            while True:
                try:
                    count = int(prompt_input(f"How many department sites would you like to create? (1-{max_sites})", required=True))
                    if 1 <= count <= max_sites:
                        break
                    print_warning(f"Please enter a number between 1 and {max_sites}")
                except ValueError:
                    print_warning("Please enter a valid number")
        
        print_info(f"Generating {count} department sites...")
        sites = generate_random_sites(count, department_templates)
        print_success(f"Generated {len(sites)} department sites")
    
    elif mode == "adhoc":
        print()
        print_info("Ad-hoc Sites Mode selected")
        print(f"  {Colors.CYAN}ℹ{Colors.NC} These are user-created sites (projects, teams, events, clubs)")
        print()
        
        max_sites = len(ADHOC_SITES)
        while True:
            try:
                count = int(prompt_input(f"How many ad-hoc sites would you like to create? (1-{max_sites})", required=True))
                if 1 <= count <= max_sites:
                    break
                print_warning(f"Please enter a number between 1 and {max_sites}")
            except ValueError:
                print_warning("Please enter a valid number")
        
        print_info(f"Generating {count} ad-hoc sites...")
        sites = generate_adhoc_sites(count)
        print_success(f"Generated {len(sites)} ad-hoc sites")
    
    else:  # mixed mode
        print()
        print_info("Mixed Sites Mode selected")
        print(f"  {Colors.CYAN}ℹ{Colors.NC} This creates a realistic mix of department and ad-hoc sites")
        print()
        
        max_dept = department_template_count
        max_adhoc = len(ADHOC_SITES)
        
        print(f"  {Colors.WHITE}Step 1: Department Sites{Colors.NC}")
        while True:
            try:
                dept_count = int(prompt_input(f"How many department sites? (0-{max_dept})", required=True))
                if 0 <= dept_count <= max_dept:
                    break
                print_warning(f"Please enter a number between 0 and {max_dept}")
            except ValueError:
                print_warning("Please enter a valid number")
        
        print()
        print(f"  {Colors.WHITE}Step 2: Ad-hoc Sites{Colors.NC}")
        while True:
            try:
                adhoc_count = int(prompt_input(f"How many ad-hoc sites? (0-{max_adhoc})", required=True))
                if 0 <= adhoc_count <= max_adhoc:
                    break
                print_warning(f"Please enter a number between 0 and {max_adhoc}")
            except ValueError:
                print_warning("Please enter a valid number")
        
        if dept_count == 0 and adhoc_count == 0:
            print_error("You must create at least one site")
            sys.exit(1)
        
        print()
        print_info(f"Generating {dept_count} department sites + {adhoc_count} ad-hoc sites...")
        sites = generate_mixed_sites(dept_count, adhoc_count, department_templates)
        print_success(f"Generated {len(sites)} total sites ({dept_count} department + {adhoc_count} ad-hoc)")
    
    # Display sites
    print()
    print(f"  {Colors.WHITE}Sites to be created:{Colors.NC}")
    for site in sites:
        # Show the actual URL name (without hyphens/special chars) that SharePoint will use
        url_name = ''.join(c for c in site['name'] if c.isalnum())
        if site.get('template') == 'SITEPAGEPUBLISHING#0':
            url_name = url_name + 'site'
        print(f"    - {site['display_name']}")
        print(f"      {Colors.DIM}URL: /sites/{url_name}{Colors.NC}")
    
    print()
    if not confirm("Continue with these sites?"):
        print_warning("Deployment cancelled.")
        sys.exit(0)
    
    return mode, sites


def select_azure_tenant() -> str:
    """Select Azure tenant and return tenant ID."""
    print_step(4, "Select Azure Tenant")
    
    print_info("Fetching available tenants...")
    tenants = get_azure_tenants()
    
    if not tenants:
        print_error("No tenants found. Please check your Azure account.")
        sys.exit(1)
    
    if len(tenants) == 1:
        tenant = tenants[0]
        print_info(f"Only one tenant available: {tenant['displayName']}")
    else:
        print()
        print(f"  {Colors.WHITE}Available Tenants:{Colors.NC}")
        print()
        for i, tenant in enumerate(tenants, 1):
            print(f"    [{i}] {tenant['displayName']} ({tenant['tenantId']})")
        print()
        
        selection = prompt_selection("Select tenant", len(tenants))
        tenant = tenants[selection - 1]
    
    print_success(f"Selected Tenant: {tenant['displayName']}")
    return tenant['tenantId']


def select_azure_subscription(tenant_id: str) -> str:
    """Select Azure subscription and return subscription ID."""
    print_step(5, "Select Azure Subscription")
    
    print_info("Fetching subscriptions for tenant...")
    subscriptions = get_azure_subscriptions(tenant_id)
    
    if not subscriptions:
        print_error("No enabled subscriptions found in this tenant.")
        sys.exit(1)
    
    if len(subscriptions) == 1:
        subscription = subscriptions[0]
        print_info(f"Only one subscription available: {subscription['name']}")
    else:
        print()
        print(f"  {Colors.WHITE}Available Subscriptions:{Colors.NC}")
        print()
        for i, sub in enumerate(subscriptions, 1):
            print(f"    [{i}] {sub['name']} ({sub['id']})")
        print()
        
        selection = prompt_selection("Select subscription", len(subscriptions))
        subscription = subscriptions[selection - 1]
    
    print_success(f"Selected Subscription: {subscription['name']}")
    set_azure_subscription(subscription['id'])
    return subscription['id']


def configure_resource_group() -> Tuple[str, str, bool]:
    """Configure resource group and return (name, location, use_existing)."""
    print_step(6, "Configure Resource Group")
    
    print()
    print(f"  {Colors.WHITE}Would you like to:{Colors.NC}")
    print("    [1] Create a new Resource Group")
    print("    [2] Use an existing Resource Group")
    print()
    
    choice = prompt_selection("Enter your choice", 2)
    
    if choice == 2:
        print_info("Fetching existing resource groups...")
        resource_groups = get_resource_groups()
        
        if not resource_groups:
            print_warning("No existing resource groups found. Creating a new one.")
            choice = 1
        else:
            print()
            print(f"  {Colors.WHITE}Available Resource Groups:{Colors.NC}")
            print()
            for i, rg in enumerate(resource_groups, 1):
                print(f"    [{i}] {rg['name']} ({rg['location']})")
            print()
            
            selection = prompt_selection("Select resource group", len(resource_groups))
            rg = resource_groups[selection - 1]
            print_success(f"Using existing Resource Group: {rg['name']}")
            return rg['name'], rg['location'], True  # True = use existing
    
    # Create new resource group
    rg_name = prompt_input("Enter Resource Group name", default="rg-sharepoint-automation", required=True)
    
    print()
    print(f"  {Colors.WHITE}Available Locations:{Colors.NC}")
    locations = [
        ("uksouth", "UK South - London"),
        ("ukwest", "UK West - Cardiff"),
        ("northeurope", "North Europe - Ireland"),
        ("westeurope", "West Europe - Netherlands"),
        ("eastus", "East US - Virginia"),
        ("westus2", "West US 2 - Washington")
    ]
    for i, (code, name) in enumerate(locations, 1):
        print(f"    [{i}] {code} ({name})")
    print()
    
    selection = prompt_selection("Select location", len(locations))
    location = locations[selection - 1][0]
    
    print_success(f"Resource Group: {rg_name}")
    print_success(f"Location: {location}")
    
    return rg_name, location, False  # False = create new


def configure_m365_settings() -> Tuple[str, str]:
    """Configure Microsoft 365 settings and return (tenant_name, admin_email)."""
    print_step(7, "Configure Microsoft 365 Settings")
    
    print()
    print(f"  {Colors.WHITE}Now we need your Microsoft 365 tenant information.{Colors.NC}")
    print()
    print("  Your M365 tenant name is the part before .onmicrosoft.com")
    print("  For example, if your domain is 'contoso.onmicrosoft.com',")
    print("  your tenant name is 'contoso'")
    print()
    
    tenant_name = prompt_input("Enter your M365 tenant name", required=True)
    admin_email = prompt_input(
        "Enter SharePoint admin email",
        required=True,
        validation_pattern=r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    )
    
    print_success(f"M365 Tenant: {tenant_name}")
    print_success(f"Admin Email: {admin_email}")
    
    return tenant_name, admin_email


def generate_terraform_config(
    tenant_id: str,
    subscription_id: str,
    rg_name: str,
    location: str,
    m365_tenant: str,
    admin_email: str,
    sites: List[Dict],
    mode: str,
    use_existing_rg: bool = False,
    key_vault_name: str = "",
    use_existing_kv: bool = False,
    step_num: int = 5
) -> None:
    """Generate the terraform.tfvars file."""
    print_step(step_num, "Generating Terraform Configuration")
    
    sites_block = format_terraform_sites_block(sites)
    
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mode_desc = "Configuration File" if mode == "config" else f"Random Generation ({len(sites)} sites)"
    use_existing_str = "true" if use_existing_rg else "false"
    
    # Key Vault configuration
    kv_name_line = f'key_vault_name                = "{key_vault_name}"' if key_vault_name else '# key_vault_name              = ""  # Leave empty for auto-generated name'
    create_kv_str = "false" if use_existing_kv else "true"
    
    config = f'''# Auto-generated by deploy.py
# Generated at: {timestamp}
# Mode: {mode_desc}

# Azure Configuration
azure_tenant_id       = "{tenant_id}"
azure_subscription_id = "{subscription_id}"
resource_group_name   = "{rg_name}"
location              = "{location}"

# Resource Group Mode
# Set to true to use an existing resource group, false to create a new one
use_existing_resource_group = {use_existing_str}

# Microsoft 365 Configuration
m365_tenant_name       = "{m365_tenant}"
sharepoint_admin_email = "{admin_email}"

# SharePoint Sites Configuration
{sites_block}

# Key Vault Settings
{kv_name_line}
create_key_vault              = {create_kv_str}
enable_soft_delete_protection = true

# Optional Settings
deployment_timeout_minutes    = 30

tags = {{
  Environment = "Production"
  Project     = "SharePoint-Sites-Automation"
  ManagedBy   = "Terraform"
  DeployedBy  = "{admin_email}"
  SiteCount   = "{len(sites)}"
}}
'''
    
    tfvars_path = TERRAFORM_DIR / "terraform.tfvars"
    with open(tfvars_path, 'w') as f:
        f.write(config)
    
    print_success(f"Terraform configuration generated: {tfvars_path}")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Deploy SharePoint sites using Terraform",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python deploy.py                          # Interactive mode
  python deploy.py --config sites.json      # Use custom config file
  python deploy.py --random 10              # Generate 10 random sites
  python deploy.py --random 5 --auto-approve  # Generate 5 sites, auto-approve
        """
    )
    parser.add_argument(
        '-c', '--config',
        help='Path to sites configuration JSON file'
    )
    parser.add_argument(
        '-r', '--random',
        type=int,
        metavar='COUNT',
        help='Generate COUNT random department sites'
    )
    parser.add_argument(
        '-s', '--skip-prerequisites',
        action='store_true',
        help='Skip prerequisite validation'
    )
    parser.add_argument(
        '-a', '--auto-approve',
        action='store_true',
        help='Auto-approve Terraform apply'
    )
    parser.add_argument(
        '-e', '--environment',
        metavar='NAME',
        help='Use pre-configured environment by name (from config/environments.json)'
    )
    
    args = parser.parse_args()
    
    # Validate random count
    validation_config_path = Path(args.config) if args.config else DEFAULT_CONFIG_FILE
    max_random_sites = len(get_department_site_templates(validation_config_path))
    if args.random is not None and (args.random < 1 or args.random > max_random_sites):
        print_error(f"Random count must be between 1 and {max_random_sites}")
        sys.exit(1)
    
    # Clear screen and show banner
    os.system('cls' if os.name == 'nt' else 'clear')
    print_banner("SHAREPOINT SITES DEPLOYMENT")
    
    print(f"  {Colors.WHITE}Welcome to the SharePoint Sites Deployment Script!{Colors.NC}")
    print()
    
    # Step 1: Check prerequisites (if not skipped)
    if not args.skip_prerequisites:
        if not check_and_install_prerequisites():
            print()
            print_error("Prerequisites not met. Please install missing tools and try again.")
            print_info("You can skip this check with: python deploy.py --skip-prerequisites")
            sys.exit(1)
        print()
        print_success("All prerequisites validated!")
    
    # Track step numbers dynamically
    current_step = 1 if args.skip_prerequisites else 2
    
    # Select site generation mode
    mode, sites = select_site_mode(args, current_step)
    current_step += 1
    
    # Azure authentication
    print_step(current_step, "Azure Authentication")
    print_info("Checking Azure CLI authentication status...")
    
    if not check_azure_login():
        print_warning("You are not logged in to Azure CLI.")
        if not azure_login():
            print_error("Azure login failed. Please try again.")
            sys.exit(1)
    
    print_success("Azure CLI authentication successful!")
    current_step += 1
    
    # Select environment or manual configuration
    print_step(current_step, "Select Environment")
    
    # Check if environment was specified via command line
    selected_env = None
    if args.environment:
        selected_env = get_environment_by_name(args.environment)
        if not selected_env:
            print_error(f"Environment '{args.environment}' not found in config/environments.json")
            print_info("Available environments:")
            env_data = load_environments()
            if env_data:
                for env in env_data.get('environments', []):
                    print(f"  - {env.get('name')}")
            sys.exit(1)
        print_info(f"Using environment from command line: {args.environment}")
    else:
        selected_env = select_environment()
    
    if selected_env:
        # Use pre-configured environment
        tenant_id, subscription_id, rg_name, location, m365_tenant, admin_email, key_vault_name = use_environment_config(selected_env)
        # For pre-configured environments, ask if the resource group already exists
        print()
        print_info(f"Using pre-configured environment: {selected_env.get('name')}")
        print(f"    Tenant ID:       {tenant_id}")
        print(f"    Subscription ID: {subscription_id}")
        print(f"    Resource Group:  {rg_name}")
        print(f"    Location:        {location}")
        print(f"    Key Vault:       {key_vault_name if key_vault_name else '(auto-generated)'}")
        print(f"    M365 Tenant:     {m365_tenant}")
        print(f"    Admin Email:     {admin_email}")
        print()
        print(f"  {Colors.WHITE}Does the resource group '{rg_name}' already exist in Azure?{Colors.NC}")
        use_existing_rg = confirm("Use existing resource group?")
        if use_existing_rg:
            print_info(f"Will use existing resource group: {rg_name}")
        else:
            print_info(f"Will create new resource group: {rg_name}")
        
        # Ask about Key Vault if a name is configured
        use_existing_kv = False
        if key_vault_name:
            print()
            print(f"  {Colors.WHITE}Does the Key Vault '{key_vault_name}' already exist in Azure?{Colors.NC}")
            use_existing_kv = confirm("Use existing Key Vault?")
            if use_existing_kv:
                print_info(f"Will use existing Key Vault: {key_vault_name}")
            else:
                print_info(f"Will create new Key Vault: {key_vault_name}")
        else:
            print_info("Key Vault name will be auto-generated")
    else:
        # Manual configuration
        print_info("Using manual configuration mode...")
        current_step += 1
        
        # Select tenant
        print_step(current_step, "Select Azure Tenant (Manual)")
        tenant_id = select_azure_tenant()
        current_step += 1
        
        # Select subscription
        print_step(current_step, "Select Azure Subscription")
        subscription_id = select_azure_subscription(tenant_id)
        current_step += 1
        
        # Configure resource group
        print_step(current_step, "Configure Resource Group")
        rg_name, location, use_existing_rg = configure_resource_group()
        current_step += 1
        
        # Configure M365 settings
        print_step(current_step, "Configure Microsoft 365 Settings")
        m365_tenant, admin_email = configure_m365_settings()
        
        # For manual mode, no Key Vault name specified (auto-generate)
        key_vault_name = ""
        use_existing_kv = False
    
    current_step += 1
    
    # Owner/Member Assignment
    print_step(current_step, "Configure Site Owners & Members")
    sites = select_owner_assignment_mode(sites, admin_email)
    current_step += 1
    
    # Review configuration
    print_step(current_step, "Review Configuration")
    
    rg_mode = "Existing" if use_existing_rg else "New"
    kv_mode = "Existing" if use_existing_kv else "New"
    kv_display = key_vault_name if key_vault_name else "(auto-generated)"
    
    print()
    print(f"  {Colors.WHITE}Please review your configuration:{Colors.NC}")
    print()
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    print(f"  {Colors.CYAN}| AZURE CONFIGURATION{' ' * 55}|{Colors.NC}")
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    print(f"  | Tenant ID:        {tenant_id}")
    print(f"  | Subscription ID:  {subscription_id}")
    print(f"  | Resource Group:   {rg_name} ({rg_mode})")
    print(f"  | Location:         {location}")
    print(f"  | Key Vault:        {kv_display} ({kv_mode})")
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    print(f"  {Colors.CYAN}| MICROSOFT 365 CONFIGURATION{' ' * 47}|{Colors.NC}")
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    print(f"  | M365 Tenant:      {m365_tenant}")
    print(f"  | Admin Email:      {admin_email}")
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    # Calculate owner/member totals
    total_owners = sum(len(s.get('owners', [])) for s in sites)
    total_members = sum(len(s.get('members', [])) for s in sites)
    
    print(f"  {Colors.CYAN}| SHAREPOINT SITES ({len(sites)} sites, {total_owners} owners, {total_members} members){' ' * max(0, 30 - len(str(len(sites))) - len(str(total_owners)) - len(str(total_members)))}|{Colors.NC}")
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    
    for i, site in enumerate(sites, 1):
        owners_count = len(site.get('owners', []))
        members_count = len(site.get('members', []))
        owner_info = f" [{owners_count}O/{members_count}M]" if owners_count > 0 or members_count > 0 else ""
        print(f"  | {i}. {site['name']}{owner_info}")
    
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    print()
    
    if not confirm("Is this configuration correct?"):
        print_warning("Deployment cancelled. Please run the script again.")
        sys.exit(0)
    
    current_step += 1
    
    # Generate Terraform configuration
    generate_terraform_config(
        tenant_id, subscription_id, rg_name, location,
        m365_tenant, admin_email, sites, mode, use_existing_rg,
        key_vault_name, use_existing_kv, current_step
    )
    
    current_step += 1
    
    # Initialize Terraform
    print_step(current_step, "Initializing Terraform")
    
    if not terraform_init():
        print_error("Terraform init failed!")
        sys.exit(1)
    
    print_success("Terraform initialized successfully!")

    removed_stale_state = reconcile_stale_sharepoint_site_state(sites, m365_tenant)
    if removed_stale_state:
        print_success(f"Removed stale Terraform state for {len(removed_stale_state)} missing SharePoint site(s).")
    
    current_step += 1
    
    # Terraform plan
    print_step(current_step, "Planning Deployment")
    
    if not terraform_plan():
        print_error("Terraform plan failed!")
        sys.exit(1)
    
    print_success("Terraform plan completed!")
    
    current_step += 1
    
    # Terraform apply
    print_step(current_step, "Deploying Resources")
    
    if not args.auto_approve:
        print()
        print(f"  {Colors.CYAN}╔══════════════════════════════════════════════════════════════════════════════╗{Colors.NC}")
        print(f"  {Colors.CYAN}║                         READY TO CREATE RESOURCES                           ║{Colors.NC}")
        print(f"  {Colors.CYAN}╠══════════════════════════════════════════════════════════════════════════════╣{Colors.NC}")
        print(f"  {Colors.CYAN}║{Colors.NC}                                                                              {Colors.CYAN}║{Colors.NC}")
        print(f"  {Colors.CYAN}║{Colors.NC}  The plan above shows what will be created in Azure and SharePoint.         {Colors.CYAN}║{Colors.NC}")
        print(f"  {Colors.CYAN}║{Colors.NC}                                                                              {Colors.CYAN}║{Colors.NC}")
        print(f"  {Colors.CYAN}║{Colors.NC}  {Colors.GREEN}Type 'y' and press Enter{Colors.NC} - to CREATE the SharePoint sites             {Colors.CYAN}║{Colors.NC}")
        print(f"  {Colors.CYAN}║{Colors.NC}  {Colors.RED}Type 'n' and press Enter{Colors.NC} - to CANCEL and exit without changes         {Colors.CYAN}║{Colors.NC}")
        print(f"  {Colors.CYAN}║{Colors.NC}                                                                              {Colors.CYAN}║{Colors.NC}")
        print(f"  {Colors.CYAN}╚══════════════════════════════════════════════════════════════════════════════╝{Colors.NC}")
        print()
        if not confirm("Proceed with deployment?"):
            print_warning("Deployment cancelled by user.")
            sys.exit(0)
    
    if not terraform_apply(args.auto_approve):
        print_error("Terraform apply failed!")
        sys.exit(1)
    
    print_success("Terraform apply completed successfully!")
    verified_site_urls = get_verified_sharepoint_site_urls(sites, m365_tenant)
    
    current_step += 1
    
    # Display results
    print_step(current_step, "Deployment Complete")
    
    print()
    terraform_output()
    
    print()
    print_banner("DEPLOYMENT SUCCESSFUL")
    print()
    print(f"  {Colors.GREEN}Your SharePoint sites have been created!{Colors.NC}")
    print()

    # Separate sites by visibility
    private_sites = [s for s in sites if s.get('visibility', 'Private').lower() == 'private']
    public_sites = [s for s in sites if s.get('visibility', 'Private').lower() == 'public']
    unresolved_sites: List[str] = []
    
    if private_sites:
        print(f"  {Colors.WHITE}{Colors.BOLD}Private Sites:{Colors.NC}")
        for site in private_sites:
            verified_url = verified_site_urls.get(site['name'])
            fallback_url = f"https://{m365_tenant}.sharepoint.com/sites/{get_site_url_name(site['name'], site.get('template', 'STS#3'))}"
            display_url = verified_url or fallback_url
            print(f"    {Colors.YELLOW}🔒{Colors.NC} {site.get('display_name', site['name'])}")
            print(f"       {Colors.CYAN}{display_url}{Colors.NC}")
            if not verified_url:
                unresolved_sites.append(site.get('display_name', site['name']))
        print()
    
    if public_sites:
        print(f"  {Colors.WHITE}{Colors.BOLD}Public Sites:{Colors.NC}")
        for site in public_sites:
            verified_url = verified_site_urls.get(site['name'])
            fallback_url = f"https://{m365_tenant}.sharepoint.com/sites/{get_site_url_name(site['name'], site.get('template', 'STS#3'))}"
            display_url = verified_url or fallback_url
            print(f"    {Colors.GREEN}🌐{Colors.NC} {site.get('display_name', site['name'])}")
            print(f"       {Colors.CYAN}{display_url}{Colors.NC}")
            if not verified_url:
                unresolved_sites.append(site.get('display_name', site['name']))
        print()
    
    print(f"  {Colors.WHITE}{Colors.BOLD}SharePoint Admin Center:{Colors.NC}")
    print(f"    {Colors.CYAN}https://{m365_tenant}-admin.sharepoint.com{Colors.NC}")
    print()

    if unresolved_sites:
        print_warning(
            f"Could not verify {len(unresolved_sites)} site(s) in SharePoint yet. They may still be provisioning or need another deploy run if they were previously deleted outside Terraform."
        )
        for site_name in unresolved_sites:
            print(f"    {Colors.YELLOW}- {site_name}{Colors.NC}")
        print()
    
    print(f"  {Colors.YELLOW}Note: Sites may take a few minutes to be fully accessible after creation.{Colors.NC}")
    print()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print()
        print_warning("Deployment cancelled by user.")
        sys.exit(0)
    except Exception as e:
        print_error(f"An unexpected error occurred: {e}")
        sys.exit(1)