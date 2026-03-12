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
import zipfile
import shutil
from pathlib import Path
from typing import Dict, List, Tuple, Optional

# ============================================================================
# CONFIGURATION
# ============================================================================

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_DIR = SCRIPT_DIR.parent
TERRAFORM_DIR = PROJECT_DIR / "terraform"
CONFIG_DIR = PROJECT_DIR / "config"
DEFAULT_CONFIG_FILE = CONFIG_DIR / "sites.json"
ENVIRONMENTS_FILE = CONFIG_DIR / "environments.json"

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
    CYAN = '\033[0;36m'
    WHITE = '\033[1;37m'
    BOLD = '\033[1m'
    NC = '\033[0m'  # No Color

    @classmethod
    def disable(cls):
        """Disable colors (for non-TTY output)."""
        cls.RED = cls.GREEN = cls.YELLOW = cls.CYAN = cls.WHITE = cls.BOLD = cls.NC = ''


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


def generate_random_sites(count: int) -> List[Dict]:
    """Generate a list of realistic organizational department sites.
    
    The maximum number of unique sites is limited to the number of templates
    in DEPARTMENT_SITES (currently 39). Each site includes realistic visibility
    settings (Private or Public) and appropriate templates.
    """
    max_sites = len(DEPARTMENT_SITES)
    
    if count > max_sites:
        print_warning(f"Requested {count} sites, but only {max_sites} unique templates available.")
        print_info(f"Generating {max_sites} sites instead.")
        count = max_sites
    
    # Shuffle the department sites to get random selection
    available_sites = DEPARTMENT_SITES.copy()
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


def generate_mixed_sites(dept_count: int, adhoc_count: int) -> List[Dict]:
    """Generate a mix of department sites and ad-hoc sites.
    
    This creates a more realistic environment with both official department
    sites and organic user-created sites.
    
    Args:
        dept_count: Number of department sites to generate
        adhoc_count: Number of ad-hoc sites to generate
    
    Returns:
        Combined list of sites
    """
    dept_sites = generate_random_sites(dept_count)
    adhoc_sites = generate_adhoc_sites(adhoc_count)
    
    # Combine and shuffle for a more natural mix
    all_sites = dept_sites + adhoc_sites
    random.shuffle(all_sites)
    
    return all_sites


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


def terraform_apply(auto_approve: bool = False) -> bool:
    """Run Terraform apply."""
    print_info("Running terraform apply...")
    try:
        cmd = ['terraform', 'apply']
        if auto_approve:
            cmd.append('-auto-approve')
        cmd.append('tfplan')
        subprocess.run(cmd, cwd=TERRAFORM_DIR, check=True, env=get_terraform_env())
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


def use_environment_config(env: Dict) -> Tuple[str, str, str, str, str, str]:
    """Extract configuration from selected environment."""
    azure = env.get('azure', {})
    m365 = env.get('m365', {})
    
    tenant_id = azure.get('tenant_id', '')
    subscription_id = azure.get('subscription_id', '')
    resource_group = azure.get('resource_group', '')
    location = azure.get('location', 'westus2')
    m365_tenant = m365.get('tenant_name', '')
    admin_email = m365.get('admin_email', '')
    
    return tenant_id, subscription_id, resource_group, location, m365_tenant, admin_email


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
    
    # Check command line arguments
    if args.random and args.random > 0:
        mode = "random"
        print_info("Using random generation mode (from command line)")
    elif args.config:
        mode = "config"
        print_info("Using configuration file mode (from command line)")
    else:
        # Interactive selection
        print()
        print(f"  {Colors.WHITE}How would you like to define your SharePoint sites?{Colors.NC}")
        print()
        print(f"  {Colors.CYAN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.NC}")
        print(f"  {Colors.WHITE}CONFIGURATION FILE OPTIONS:{Colors.NC}")
        print()
        print("    [1] Use Configuration File Only (config/sites.json)")
        print("        - Edit the JSON file to add your custom site names")
        print("        - Full control over site names, descriptions, and settings")
        print()
        print(f"    [2] Configuration File + Ad-hoc Sites")
        print("        - Uses your custom sites from config/sites.json")
        print("        - PLUS random ad-hoc sites (projects, teams, events)")
        print("        - Best for realistic environments with your specific sites")
        print()
        print(f"  {Colors.CYAN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Colors.NC}")
        print(f"  {Colors.WHITE}RANDOM GENERATION OPTIONS:{Colors.NC}")
        print()
        print(f"    [3] Generate Department Sites ({len(DEPARTMENT_SITES)} templates)")
        print("        - Official department sites (HR, Finance, IT, Legal, etc.)")
        print("        - Realistic organizational structure")
        print("        - Mix of Private and Public visibility")
        print()
        print(f"    [4] Generate Ad-hoc Sites ({len(ADHOC_SITES)} templates)")
        print("        - User-created sites (projects, teams, events, clubs)")
        print("        - Simulates organic SharePoint usage by employees")
        print("        - Includes working groups, social clubs, regional offices")
        print()
        print(f"    [5] Generate Mixed Sites (Department + Ad-hoc)")
        print("        - Combines both types for maximum realism")
        print("        - Specify count for each type separately")
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
        
        max_sites = len(DEPARTMENT_SITES)
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
        sites = generate_random_sites(count)
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
        
        max_dept = len(DEPARTMENT_SITES)
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
        sites = generate_mixed_sites(dept_count, adhoc_count)
        print_success(f"Generated {len(sites)} total sites ({dept_count} department + {adhoc_count} ad-hoc)")
    
    # Display sites
    print()
    print(f"  {Colors.WHITE}Sites to be created:{Colors.NC}")
    for site in sites:
        print(f"    - {site['name']} ({site['display_name']})")
    
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
    step_num: int = 5
) -> None:
    """Generate the terraform.tfvars file."""
    print_step(step_num, "Generating Terraform Configuration")
    
    sites_block = format_terraform_sites_block(sites)
    
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mode_desc = "Configuration File" if mode == "config" else f"Random Generation ({len(sites)} sites)"
    use_existing_str = "true" if use_existing_rg else "false"
    
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

# Optional Settings
create_key_vault              = true
enable_soft_delete_protection = true
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
        help=f'Generate COUNT random sites (1-{len(DEPARTMENT_SITES)})'
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
    max_random_sites = len(DEPARTMENT_SITES)
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
        tenant_id, subscription_id, rg_name, location, m365_tenant, admin_email = use_environment_config(selected_env)
        # For pre-configured environments, ask if the resource group already exists
        print()
        print_info(f"Using pre-configured environment: {selected_env.get('name')}")
        print(f"    Tenant ID:       {tenant_id}")
        print(f"    Subscription ID: {subscription_id}")
        print(f"    Resource Group:  {rg_name}")
        print(f"    Location:        {location}")
        print(f"    M365 Tenant:     {m365_tenant}")
        print(f"    Admin Email:     {admin_email}")
        print()
        print(f"  {Colors.WHITE}Does the resource group '{rg_name}' already exist in Azure?{Colors.NC}")
        use_existing_rg = confirm("Use existing resource group?")
        if use_existing_rg:
            print_info(f"Will use existing resource group: {rg_name}")
        else:
            print_info(f"Will create new resource group: {rg_name}")
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
    
    current_step += 1
    
    # Review configuration
    print_step(current_step, "Review Configuration")
    
    rg_mode = "Existing" if use_existing_rg else "New"
    
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
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    print(f"  {Colors.CYAN}| MICROSOFT 365 CONFIGURATION{' ' * 47}|{Colors.NC}")
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    print(f"  | M365 Tenant:      {m365_tenant}")
    print(f"  | Admin Email:      {admin_email}")
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    print(f"  {Colors.CYAN}| SHAREPOINT SITES ({len(sites)} sites){' ' * (53 - len(str(len(sites))))}|{Colors.NC}")
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    
    for i, site in enumerate(sites, 1):
        print(f"  | {i}. {site['name']}")
    
    print(f"  {Colors.CYAN}+{'-' * 75}+{Colors.NC}")
    print()
    
    if not confirm("Is this configuration correct?"):
        print_warning("Deployment cancelled. Please run the script again.")
        sys.exit(0)
    
    current_step += 1
    
    # Generate Terraform configuration
    generate_terraform_config(
        tenant_id, subscription_id, rg_name, location,
        m365_tenant, admin_email, sites, mode, use_existing_rg, current_step
    )
    
    current_step += 1
    
    # Initialize Terraform
    print_step(current_step, "Initializing Terraform")
    
    if not terraform_init():
        print_error("Terraform init failed!")
        sys.exit(1)
    
    print_success("Terraform initialized successfully!")
    
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
    
    # Helper function to convert site name to URL-safe mailNickname
    def get_site_url_name(site_name: str, template: str) -> str:
        """Convert site name to the actual URL name used by SharePoint."""
        # Remove all non-alphanumeric characters (same as PowerShell script)
        url_name = ''.join(c for c in site_name if c.isalnum())
        # Communication sites created as fallback get 'site' suffix
        if template == "SITEPAGEPUBLISHING#0":
            url_name = url_name + "site"
        return url_name
    
    # Separate sites by visibility
    private_sites = [s for s in sites if s.get('visibility', 'Private').lower() == 'private']
    public_sites = [s for s in sites if s.get('visibility', 'Private').lower() == 'public']
    
    if private_sites:
        print(f"  {Colors.WHITE}{Colors.BOLD}Private Sites:{Colors.NC}")
        for site in private_sites:
            url_name = get_site_url_name(site['name'], site.get('template', 'STS#3'))
            print(f"    {Colors.YELLOW}🔒{Colors.NC} {site.get('display_name', site['name'])}")
            print(f"       {Colors.CYAN}https://{m365_tenant}.sharepoint.com/sites/{url_name}{Colors.NC}")
        print()
    
    if public_sites:
        print(f"  {Colors.WHITE}{Colors.BOLD}Public Sites:{Colors.NC}")
        for site in public_sites:
            url_name = get_site_url_name(site['name'], site.get('template', 'STS#3'))
            print(f"    {Colors.GREEN}🌐{Colors.NC} {site.get('display_name', site['name'])}")
            print(f"       {Colors.CYAN}https://{m365_tenant}.sharepoint.com/sites/{url_name}{Colors.NC}")
        print()
    
    print(f"  {Colors.WHITE}{Colors.BOLD}SharePoint Admin Center:{Colors.NC}")
    print(f"    {Colors.CYAN}https://{m365_tenant}-admin.sharepoint.com{Colors.NC}")
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