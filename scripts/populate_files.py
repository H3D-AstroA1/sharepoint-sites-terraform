#!/usr/bin/env python3
"""
SharePoint File Population Script

This script populates SharePoint sites with realistic-looking files to simulate
an actual organization's document structure. It creates various file types
(Word, Excel, PowerPoint, PDF) with department-appropriate content.

Usage:
    python populate_files.py                      # Interactive mode
    python populate_files.py --files 100          # Create 100 files across all sites
    python populate_files.py --files 50 --site hr # Create 50 files in HR site only
    python populate_files.py --help               # Show help

Requirements:
    - Python 3.8+
    - Azure CLI (logged in)
    - Microsoft Graph API permissions (Sites.ReadWrite.All, Files.ReadWrite.All)
"""

import argparse
import json
import os
import platform
import random
import shutil
import subprocess
import sys
import tempfile
import base64
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any

# ============================================================================
# CONFIGURATION
# ============================================================================

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_DIR = SCRIPT_DIR.parent
CONFIG_DIR = PROJECT_DIR / "config"
ENVIRONMENTS_FILE = CONFIG_DIR / "environments.json"

# Maximum files that can be created in one run
MAX_FILES = 1000

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
# REALISTIC FILE TEMPLATES BY DEPARTMENT
# ============================================================================

FILE_TEMPLATES = {
    "executive-leadership": {
        "folders": ["Board Materials", "Strategic Planning", "Executive Reports", "Confidential"],
        "files": [
            {"name": "Q{quarter}_Board_Meeting_Agenda_{year}.docx", "type": "word"},
            {"name": "Strategic_Plan_{year}-{next_year}.pptx", "type": "powerpoint"},
            {"name": "Executive_Summary_Report_{month}_{year}.docx", "type": "word"},
            {"name": "Annual_Budget_Overview_{year}.xlsx", "type": "excel"},
            {"name": "Leadership_Team_OKRs_Q{quarter}_{year}.xlsx", "type": "excel"},
            {"name": "Board_Resolution_{number}_{year}.pdf", "type": "pdf"},
            {"name": "CEO_Town_Hall_Presentation_{month}_{year}.pptx", "type": "powerpoint"},
            {"name": "Investor_Relations_Update_Q{quarter}.pptx", "type": "powerpoint"},
        ]
    },
    "human-resources": {
        "folders": ["Policies", "Recruitment", "Training", "Employee Records", "Benefits"],
        "files": [
            {"name": "Employee_Handbook_v{version}.pdf", "type": "pdf"},
            {"name": "Job_Description_{role}.docx", "type": "word"},
            {"name": "Interview_Questions_Template.docx", "type": "word"},
            {"name": "Onboarding_Checklist_{year}.xlsx", "type": "excel"},
            {"name": "Benefits_Summary_{year}.pdf", "type": "pdf"},
            {"name": "Performance_Review_Template.docx", "type": "word"},
            {"name": "Training_Schedule_Q{quarter}_{year}.xlsx", "type": "excel"},
            {"name": "HR_Metrics_Dashboard_{month}_{year}.xlsx", "type": "excel"},
            {"name": "Leave_Policy_v{version}.pdf", "type": "pdf"},
            {"name": "Compensation_Guidelines_{year}.xlsx", "type": "excel"},
        ]
    },
    "finance-department": {
        "folders": ["Reports", "Budgets", "Invoices", "Audits", "Tax"],
        "files": [
            {"name": "Monthly_Financial_Report_{month}_{year}.xlsx", "type": "excel"},
            {"name": "Annual_Budget_{year}.xlsx", "type": "excel"},
            {"name": "Cash_Flow_Statement_Q{quarter}_{year}.xlsx", "type": "excel"},
            {"name": "Expense_Report_Template.xlsx", "type": "excel"},
            {"name": "Invoice_Processing_Guide.pdf", "type": "pdf"},
            {"name": "Audit_Findings_{year}.docx", "type": "word"},
            {"name": "Tax_Filing_Checklist_{year}.xlsx", "type": "excel"},
            {"name": "Vendor_Payment_Schedule_{month}.xlsx", "type": "excel"},
            {"name": "Financial_Forecast_{year}-{next_year}.xlsx", "type": "excel"},
            {"name": "Cost_Center_Report_{month}_{year}.xlsx", "type": "excel"},
        ]
    },
    "claims-department": {
        "folders": ["Active Claims", "Closed Claims", "Templates", "Reports", "Policies"],
        "files": [
            {"name": "Claim_Form_Template_v{version}.docx", "type": "word"},
            {"name": "Claims_Processing_Guide.pdf", "type": "pdf"},
            {"name": "Monthly_Claims_Report_{month}_{year}.xlsx", "type": "excel"},
            {"name": "Claim_Status_Tracker_{year}.xlsx", "type": "excel"},
            {"name": "Claims_Policy_v{version}.pdf", "type": "pdf"},
            {"name": "Investigation_Checklist.docx", "type": "word"},
            {"name": "Settlement_Authorization_Template.docx", "type": "word"},
            {"name": "Claims_Metrics_Dashboard_Q{quarter}_{year}.xlsx", "type": "excel"},
            {"name": "Fraud_Detection_Guidelines.pdf", "type": "pdf"},
            {"name": "Appeals_Process_Documentation.docx", "type": "word"},
        ]
    },
    "it-department": {
        "folders": ["Documentation", "Procedures", "Architecture", "Security", "Projects"],
        "files": [
            {"name": "System_Architecture_Diagram_v{version}.pptx", "type": "powerpoint"},
            {"name": "IT_Security_Policy_v{version}.pdf", "type": "pdf"},
            {"name": "Network_Topology_{year}.pptx", "type": "powerpoint"},
            {"name": "Disaster_Recovery_Plan_v{version}.docx", "type": "word"},
            {"name": "Software_Inventory_{year}.xlsx", "type": "excel"},
            {"name": "Change_Management_Process.docx", "type": "word"},
            {"name": "Incident_Response_Playbook.pdf", "type": "pdf"},
            {"name": "IT_Budget_Request_{year}.xlsx", "type": "excel"},
            {"name": "Vendor_Assessment_{vendor}.docx", "type": "word"},
            {"name": "Project_Status_Report_{month}_{year}.pptx", "type": "powerpoint"},
        ]
    },
    "marketing-department": {
        "folders": ["Campaigns", "Brand Assets", "Analytics", "Content", "Events"],
        "files": [
            {"name": "Marketing_Plan_{year}.pptx", "type": "powerpoint"},
            {"name": "Brand_Guidelines_v{version}.pdf", "type": "pdf"},
            {"name": "Campaign_Performance_Q{quarter}_{year}.xlsx", "type": "excel"},
            {"name": "Social_Media_Calendar_{month}_{year}.xlsx", "type": "excel"},
            {"name": "Press_Release_Template.docx", "type": "word"},
            {"name": "Event_Planning_Checklist.xlsx", "type": "excel"},
            {"name": "Customer_Survey_Results_{year}.pptx", "type": "powerpoint"},
            {"name": "Competitor_Analysis_{year}.docx", "type": "word"},
            {"name": "Marketing_Budget_{year}.xlsx", "type": "excel"},
            {"name": "Content_Strategy_{year}.pptx", "type": "powerpoint"},
        ]
    },
    "sales-department": {
        "folders": ["Proposals", "Contracts", "Reports", "Training", "Territories"],
        "files": [
            {"name": "Sales_Proposal_Template.pptx", "type": "powerpoint"},
            {"name": "Contract_Template_v{version}.docx", "type": "word"},
            {"name": "Sales_Report_Q{quarter}_{year}.xlsx", "type": "excel"},
            {"name": "Pipeline_Analysis_{month}_{year}.xlsx", "type": "excel"},
            {"name": "Product_Pricing_Sheet_{year}.xlsx", "type": "excel"},
            {"name": "Sales_Training_Materials.pptx", "type": "powerpoint"},
            {"name": "Territory_Map_{region}.pptx", "type": "powerpoint"},
            {"name": "Customer_Case_Study_{company}.docx", "type": "word"},
            {"name": "Commission_Structure_{year}.xlsx", "type": "excel"},
            {"name": "Quarterly_Forecast_Q{quarter}_{year}.xlsx", "type": "excel"},
        ]
    },
    "legal-department": {
        "folders": ["Contracts", "Compliance", "Policies", "Litigation", "Templates"],
        "files": [
            {"name": "NDA_Template_v{version}.docx", "type": "word"},
            {"name": "Employment_Agreement_Template.docx", "type": "word"},
            {"name": "Compliance_Checklist_{year}.xlsx", "type": "excel"},
            {"name": "Legal_Hold_Notice_Template.docx", "type": "word"},
            {"name": "Privacy_Policy_v{version}.pdf", "type": "pdf"},
            {"name": "Terms_of_Service_v{version}.pdf", "type": "pdf"},
            {"name": "Vendor_Contract_Template.docx", "type": "word"},
            {"name": "Regulatory_Update_{month}_{year}.docx", "type": "word"},
            {"name": "IP_Portfolio_Summary_{year}.xlsx", "type": "excel"},
            {"name": "Legal_Matter_Tracker_{year}.xlsx", "type": "excel"},
        ]
    },
    "operations-department": {
        "folders": ["Procedures", "Reports", "Inventory", "Facilities", "Quality"],
        "files": [
            {"name": "Standard_Operating_Procedure_{process}.docx", "type": "word"},
            {"name": "Operations_Dashboard_{month}_{year}.xlsx", "type": "excel"},
            {"name": "Inventory_Report_{month}_{year}.xlsx", "type": "excel"},
            {"name": "Facility_Maintenance_Schedule_{year}.xlsx", "type": "excel"},
            {"name": "Quality_Control_Checklist.xlsx", "type": "excel"},
            {"name": "Vendor_Performance_Review_{vendor}.docx", "type": "word"},
            {"name": "Process_Improvement_Plan_{year}.pptx", "type": "powerpoint"},
            {"name": "Safety_Inspection_Report_{month}_{year}.docx", "type": "word"},
            {"name": "Capacity_Planning_{year}.xlsx", "type": "excel"},
            {"name": "Logistics_Report_Q{quarter}_{year}.xlsx", "type": "excel"},
        ]
    },
    "product-management": {
        "folders": ["Roadmaps", "Requirements", "Research", "Releases", "Feedback"],
        "files": [
            {"name": "Product_Roadmap_{year}.pptx", "type": "powerpoint"},
            {"name": "PRD_{feature}_v{version}.docx", "type": "word"},
            {"name": "User_Research_Findings_{month}_{year}.pptx", "type": "powerpoint"},
            {"name": "Release_Notes_v{version}.docx", "type": "word"},
            {"name": "Feature_Prioritization_Matrix.xlsx", "type": "excel"},
            {"name": "Customer_Feedback_Analysis_Q{quarter}.xlsx", "type": "excel"},
            {"name": "Competitive_Analysis_{year}.pptx", "type": "powerpoint"},
            {"name": "Product_Metrics_Dashboard_{month}.xlsx", "type": "excel"},
            {"name": "Sprint_Planning_Template.xlsx", "type": "excel"},
            {"name": "Go_To_Market_Plan_{product}.pptx", "type": "powerpoint"},
        ]
    },
    "customer-service": {
        "folders": ["Scripts", "Training", "Reports", "Knowledge Base", "Escalations"],
        "files": [
            {"name": "Customer_Service_Script_{scenario}.docx", "type": "word"},
            {"name": "Training_Manual_v{version}.pdf", "type": "pdf"},
            {"name": "Support_Metrics_Report_{month}_{year}.xlsx", "type": "excel"},
            {"name": "FAQ_Document_v{version}.docx", "type": "word"},
            {"name": "Escalation_Procedures.docx", "type": "word"},
            {"name": "Customer_Satisfaction_Survey_Q{quarter}.xlsx", "type": "excel"},
            {"name": "Call_Quality_Scorecard.xlsx", "type": "excel"},
            {"name": "Knowledge_Base_Article_{topic}.docx", "type": "word"},
            {"name": "SLA_Performance_Report_{month}.xlsx", "type": "excel"},
            {"name": "Agent_Performance_Dashboard_{month}.xlsx", "type": "excel"},
        ]
    },
    "default": {
        "folders": ["Documents", "Reports", "Templates", "Archive"],
        "files": [
            {"name": "Meeting_Notes_{date}.docx", "type": "word"},
            {"name": "Project_Status_Report_{month}_{year}.pptx", "type": "powerpoint"},
            {"name": "Budget_Tracker_{year}.xlsx", "type": "excel"},
            {"name": "Team_Directory_{year}.xlsx", "type": "excel"},
            {"name": "Process_Documentation_v{version}.docx", "type": "word"},
            {"name": "Quarterly_Review_Q{quarter}_{year}.pptx", "type": "powerpoint"},
            {"name": "Action_Items_Tracker.xlsx", "type": "excel"},
            {"name": "Policy_Document_v{version}.pdf", "type": "pdf"},
        ]
    }
}

# Variable replacements for file names
ROLES = ["Software_Engineer", "Product_Manager", "Data_Analyst", "Marketing_Specialist", 
         "Sales_Representative", "HR_Coordinator", "Financial_Analyst", "Project_Manager"]
VENDORS = ["Acme_Corp", "TechSolutions", "GlobalServices", "DataPro", "CloudFirst"]
REGIONS = ["North_America", "EMEA", "APAC", "LATAM"]
COMPANIES = ["Contoso", "Fabrikam", "Northwind", "Adventure_Works", "Wide_World"]
PROCESSES = ["Onboarding", "Procurement", "Expense_Approval", "Change_Request", "Incident_Management"]
FEATURES = ["User_Authentication", "Dashboard", "Reporting", "Integration", "Mobile_App"]
PRODUCTS = ["Enterprise_Suite", "Cloud_Platform", "Mobile_App", "Analytics_Tool"]
SCENARIOS = ["Billing_Inquiry", "Technical_Support", "Account_Setup", "Complaint_Resolution"]
TOPICS = ["Password_Reset", "VPN_Setup", "Email_Configuration", "Software_Installation"]

# ============================================================================
# USER-NAMED FILE TEMPLATES (for Azure AD mode)
# ============================================================================
# These templates create files that appear to be created by specific users
# Format: {user_name} is replaced with actual Azure AD user display names

USER_FILE_TEMPLATES = [
    # Personal documents
    {"name": "{user_name} - Meeting Notes {date}.docx", "type": "word"},
    {"name": "{user_name} - Project Update {date}.docx", "type": "word"},
    {"name": "{user_name} - Weekly Report {date}.docx", "type": "word"},
    {"name": "{user_name} - Action Items {date}.docx", "type": "word"},
    {"name": "{user_name} - Draft Proposal.docx", "type": "word"},
    {"name": "{user_name} - Research Notes.docx", "type": "word"},
    
    # Spreadsheets
    {"name": "{user_name} - Expense Report {month}_{year}.xlsx", "type": "excel"},
    {"name": "{user_name} - Time Tracking {month}_{year}.xlsx", "type": "excel"},
    {"name": "{user_name} - Budget Analysis.xlsx", "type": "excel"},
    {"name": "{user_name} - Data Analysis {date}.xlsx", "type": "excel"},
    {"name": "{user_name} - Project Timeline.xlsx", "type": "excel"},
    
    # Presentations
    {"name": "{user_name} - Team Presentation {date}.pptx", "type": "powerpoint"},
    {"name": "{user_name} - Status Update Q{quarter}_{year}.pptx", "type": "powerpoint"},
    {"name": "{user_name} - Training Materials.pptx", "type": "powerpoint"},
    {"name": "{user_name} - Project Overview.pptx", "type": "powerpoint"},
    
    # PDFs
    {"name": "{user_name} - Signed Agreement {date}.pdf", "type": "pdf"},
    {"name": "{user_name} - Certificate.pdf", "type": "pdf"},
    {"name": "{user_name} - Reference Document.pdf", "type": "pdf"},
]

# Shared/collaborative file templates
SHARED_FILE_TEMPLATES = [
    {"name": "Shared by {user_name} - {topic} Document.docx", "type": "word"},
    {"name": "Review Request from {user_name} - {topic}.docx", "type": "word"},
    {"name": "Feedback from {user_name} - {topic}.docx", "type": "word"},
    {"name": "{user_name} and Team - Collaboration Notes.docx", "type": "word"},
    {"name": "Approved by {user_name} - {topic}.pdf", "type": "pdf"},
]

# Topics for shared files
SHARED_TOPICS = [
    "Budget", "Contract", "Proposal", "Report", "Policy", "Procedure",
    "Guidelines", "Strategy", "Plan", "Analysis", "Review", "Assessment"
]

# ============================================================================
# AZURE AD DISCOVERY FUNCTIONS
# ============================================================================

def discover_azure_ad_users(access_token: str, max_users: int = 50) -> List[Dict[str, Any]]:
    """Discover Azure AD users from the tenant.
    
    Returns a list of users with their display names and UPNs.
    Filters out guest users and service accounts.
    """
    users = []
    try:
        # Get member users only (exclude guests)
        # URL-encode the filter parameter to handle spaces
        filter_param = urllib.parse.quote("userType eq 'Member'")
        url = f"https://graph.microsoft.com/v1.0/users?$filter={filter_param}&$select=id,userPrincipalName,displayName,department,jobTitle&$top={max_users}"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            for user in data.get("value", []):
                upn = user.get("userPrincipalName", "")
                display_name = user.get("displayName", "")
                
                # Skip service accounts and system users
                if not display_name or "@" not in upn:
                    continue
                if any(skip in upn.lower() for skip in ["admin", "service", "system", "sync", "mailbox"]):
                    continue
                if any(skip in display_name.lower() for skip in ["admin", "service", "system"]):
                    continue
                
                users.append({
                    "id": user.get("id", ""),
                    "upn": upn,
                    "displayName": display_name,
                    "department": user.get("department", ""),
                    "jobTitle": user.get("jobTitle", "")
                })
    except Exception as e:
        print_warning(f"Could not discover Azure AD users: {e}")
    
    return users


def discover_azure_ad_groups(access_token: str, max_groups: int = 30) -> List[Dict[str, Any]]:
    """Discover Azure AD groups from the tenant.
    
    Returns a list of groups suitable for folder naming.
    Filters out system groups and dynamic groups.
    """
    groups = []
    try:
        url = f"https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,groupTypes,securityEnabled,mailEnabled&$top={max_groups}"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            for group in data.get("value", []):
                display_name = group.get("displayName", "")
                group_types = group.get("groupTypes", [])
                
                # Skip dynamic groups and system groups
                if "DynamicMembership" in group_types:
                    continue
                if not display_name:
                    continue
                if any(skip in display_name.lower() for skip in ["all users", "all employees", "everyone", "system"]):
                    continue
                
                groups.append({
                    "id": group.get("id", ""),
                    "displayName": display_name,
                    "description": group.get("description", ""),
                    "isM365Group": "Unified" in group_types
                })
    except Exception as e:
        print_warning(f"Could not discover Azure AD groups: {e}")
    
    return groups


def generate_user_file_name(template: Dict[str, str], user: Dict[str, Any]) -> str:
    """Generate a file name using a user-based template."""
    name = template["name"]
    now = datetime.now()
    
    # Replace user placeholder
    user_name = user.get("displayName", "User")
    # Clean user name for file system (remove special chars)
    user_name = "".join(c for c in user_name if c.isalnum() or c in " -_").strip()
    name = name.replace("{user_name}", user_name)
    
    # Replace date/time placeholders
    name = name.replace("{date}", now.strftime("%Y-%m-%d"))
    name = name.replace("{month}", now.strftime("%B"))
    name = name.replace("{year}", str(now.year))
    name = name.replace("{quarter}", str((now.month - 1) // 3 + 1))
    
    # Replace topic placeholder
    name = name.replace("{topic}", random.choice(SHARED_TOPICS))
    
    return name


def create_user_folder(
    site_id: str,
    user: Dict[str, Any],
    access_token: str
) -> bool:
    """Create a personal folder for a user in the site."""
    user_name = user.get("displayName", "User")
    # Clean folder name
    folder_name = "".join(c for c in user_name if c.isalnum() or c in " -_").strip()
    
    return create_folder_in_sharepoint(site_id, folder_name, access_token)


def create_group_folder(
    site_id: str,
    group: Dict[str, Any],
    access_token: str
) -> bool:
    """Create a folder for a group in the site."""
    group_name = group.get("displayName", "Team")
    # Clean folder name
    folder_name = "".join(c for c in group_name if c.isalnum() or c in " -_").strip()
    
    return create_folder_in_sharepoint(site_id, folder_name, access_token)


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
    print(f"  {Colors.GREEN}✓{Colors.NC} {message}")

def print_error(message: str) -> None:
    print(f"  {Colors.RED}✗{Colors.NC} {message}")

def print_warning(message: str) -> None:
    print(f"  {Colors.YELLOW}⚠{Colors.NC} {message}")

def print_info(message: str) -> None:
    print(f"  {Colors.BLUE}ℹ{Colors.NC} {message}")

def print_progress(current: int, total: int, message: str) -> None:
    """Print progress indicator."""
    percentage = (current / total) * 100
    bar_length = 30
    filled = int(bar_length * current / total)
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
    except subprocess.CalledProcessError as e:
        print_error(f"Command failed: {e}")
        return None

def get_random_date(days_back: int = 365) -> datetime:
    """Get a random date within the specified number of days back."""
    return datetime.now() - timedelta(days=random.randint(0, days_back))

def generate_file_name(template: Dict[str, str], site_type: str) -> str:
    """Generate a realistic file name from a template."""
    name = template["name"]
    now = datetime.now()
    random_date = get_random_date()
    
    # Replace placeholders
    replacements = {
        "{year}": str(random.choice([now.year, now.year - 1])),
        "{next_year}": str(now.year + 1),
        "{month}": random_date.strftime("%B"),
        "{quarter}": str(random.randint(1, 4)),
        "{version}": f"{random.randint(1, 5)}.{random.randint(0, 9)}",
        "{number}": str(random.randint(100, 999)),
        "{date}": random_date.strftime("%Y-%m-%d"),
        "{role}": random.choice(ROLES),
        "{vendor}": random.choice(VENDORS),
        "{region}": random.choice(REGIONS),
        "{company}": random.choice(COMPANIES),
        "{process}": random.choice(PROCESSES),
        "{feature}": random.choice(FEATURES),
        "{product}": random.choice(PRODUCTS),
        "{scenario}": random.choice(SCENARIOS),
        "{topic}": random.choice(TOPICS),
    }
    
    for placeholder, value in replacements.items():
        name = name.replace(placeholder, value)
    
    return name

def get_site_type(site_name: str) -> str:
    """Determine the site type based on site name."""
    site_lower = site_name.lower()
    
    type_mappings = {
        "executive": "executive-leadership",
        "leadership": "executive-leadership",
        "board": "executive-leadership",
        "hr": "human-resources",
        "human-resources": "human-resources",
        "recruitment": "human-resources",
        "payroll": "human-resources",
        "finance": "finance-department",
        "accounting": "finance-department",
        "treasury": "finance-department",
        "claims": "claims-department",
        "it": "it-department",
        "technology": "it-department",
        "security": "it-department",
        "helpdesk": "it-department",
        "marketing": "marketing-department",
        "brand": "marketing-department",
        "communications": "marketing-department",
        "sales": "sales-department",
        "customer-success": "sales-department",
        "legal": "legal-department",
        "compliance": "legal-department",
        "operations": "operations-department",
        "facilities": "operations-department",
        "product": "product-management",
        "research": "product-management",
        "quality": "product-management",
        "customer-service": "customer-service",
        "support": "customer-service",
    }
    
    for keyword, site_type in type_mappings.items():
        if keyword in site_lower:
            return site_type
    
    return "default"

def create_minimal_docx() -> bytes:
    """Create a minimal valid DOCX file with some content."""
    # Minimal DOCX structure
    content = b'PK\x03\x04\x14\x00\x00\x00\x08\x00'
    content += b'\x00' * 100  # Padding for minimal valid structure
    return content

def create_minimal_xlsx() -> bytes:
    """Create a minimal valid XLSX file."""
    content = b'PK\x03\x04\x14\x00\x00\x00\x08\x00'
    content += b'\x00' * 100
    return content

def create_minimal_pptx() -> bytes:
    """Create a minimal valid PPTX file."""
    content = b'PK\x03\x04\x14\x00\x00\x00\x08\x00'
    content += b'\x00' * 100
    return content

def create_minimal_pdf() -> bytes:
    """Create a minimal valid PDF file."""
    pdf_content = b"""%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R >>
endobj
4 0 obj
<< /Length 44 >>
stream
BT
/F1 12 Tf
100 700 Td
(Sample Document) Tj
ET
endstream
endobj
xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000206 00000 n 
trailer
<< /Size 5 /Root 1 0 R >>
startxref
300
%%EOF"""
    return pdf_content

def create_text_file(file_name: str, site_type: str) -> bytes:
    """Create a text file with realistic content."""
    content = f"""Document: {file_name}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Department: {site_type.replace('-', ' ').title()}

This is a sample document created for testing purposes.
It simulates realistic organizational content.

---
This document is confidential and intended for internal use only.
"""
    return content.encode('utf-8')

def create_file_content(file_type: str, file_name: str = "", site_type: str = "") -> bytes:
    """Create file content based on type."""
    if file_type == "word":
        return create_minimal_docx()
    elif file_type == "excel":
        return create_minimal_xlsx()
    elif file_type == "powerpoint":
        return create_minimal_pptx()
    elif file_type == "pdf":
        return create_minimal_pdf()
    else:
        return create_text_file(file_name, site_type)

# ============================================================================
# SHAREPOINT OPERATIONS
# ============================================================================

# App config file path (same as menu.py)
APP_CONFIG_FILE = SCRIPT_DIR / ".app_config.json"


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
    except Exception:
        return None


def get_access_token() -> Optional[str]:
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
    except Exception as e:
        print_error(f"Failed to get access token: {e}")
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

def get_m365_groups_with_sites(access_token: str) -> List[Dict[str, Any]]:
    """Get Microsoft 365 Groups that have SharePoint sites (fallback method)."""
    sites = []
    url = "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c eq 'Unified')&$select=id,displayName,createdDateTime,visibility&$top=100"
    
    try:
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            groups = data.get("value", [])
            
            for group in groups:
                try:
                    # Get the SharePoint site for this group
                    site_url = f"https://graph.microsoft.com/v1.0/groups/{group['id']}/sites/root"
                    site_req = urllib.request.Request(site_url)
                    site_req.add_header("Authorization", f"Bearer {access_token}")
                    
                    with urllib.request.urlopen(site_req, timeout=10) as site_response:
                        site_data = json.loads(site_response.read().decode())
                        sites.append({
                            "id": site_data.get("id", ""),
                            "name": group.get("displayName", "Unknown"),
                            "displayName": group.get("displayName", "Unknown"),
                            "webUrl": site_data.get("webUrl", ""),
                            "groupId": group.get("id", ""),
                            "isGroupSite": True
                        })
                except Exception:
                    # Group might not have a SharePoint site
                    pass
    except Exception as e:
        print_error(f"Error getting M365 groups: {e}")
    
    return sites


# Sites to exclude (system sites, personal sites, etc. that typically cause 403 errors)
# These are protected system sites that should NOT have files uploaded to them
# Note: We only exclude exact system site names, not patterns that might match user-created sites
EXCLUDED_SITE_PATTERNS = [
    "my workspace",       # Personal OneDrive-like workspace
    "designer",           # Microsoft Designer integration site
    "contenttypehub",     # SharePoint content type hub
    "appcatalog",         # SharePoint app catalog
    "team site",          # Default team site template
    "communication site", # Root communication site template
]


def is_system_site(site: Dict[str, Any]) -> bool:
    """Check if a site is a protected system site that should not be modified.
    
    System sites include:
    - Personal OneDrive sites (/personal/)
    - Content storage sites (/contentstorage/)
    - Root site collections (tenant.sharepoint.com)
    - Default template sites (Team Site, Communication Site)
    - System sites (Designer, My Workspace, Content Type Hub, App Catalog)
    
    Returns:
        True if the site is a system site that should be protected
    """
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


def filter_writable_sites(sites: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Filter out system/personal sites that should not have files uploaded.
    
    IMPORTANT: This function protects system sites from accidental modification.
    System sites include default templates, personal sites, and SharePoint system sites.
    Only user-created sites (typically via Terraform/M365 Groups) will be returned.
    """
    filtered = []
    excluded_sites = []
    
    for site in sites:
        if is_system_site(site):
            excluded_sites.append(site)
        else:
            filtered.append(site)
    
    if excluded_sites:
        print()
        print(f"  {Colors.YELLOW}⚠ PROTECTED SYSTEM SITES (excluded from file population):{Colors.NC}")
        print(f"  {Colors.CYAN}  These sites are protected and will not have files uploaded:{Colors.NC}")
        for site in excluded_sites[:5]:  # Show first 5
            name = site.get("displayName", site.get("name", "Unknown"))
            print(f"      {Colors.CYAN}• {name}{Colors.NC}")
        if len(excluded_sites) > 5:
            print(f"      {Colors.CYAN}... and {len(excluded_sites) - 5} more{Colors.NC}")
        print()
        print_info(f"Filtered out {len(excluded_sites)} system/protected sites")
    
    return filtered


def get_sharepoint_sites(access_token: str) -> List[Dict[str, Any]]:
    """Get list of SharePoint sites using Microsoft Graph API.
    
    Falls back to M365 Groups API if Sites API returns 403.
    Filters out system/personal sites that typically cause 403 errors.
    """
    sites = []
    url = "https://graph.microsoft.com/v1.0/sites?search=*"
    use_groups_fallback = False
    
    try:
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            sites = data.get("value", [])
    except urllib.error.HTTPError as e:
        if e.code == 403:
            print_warning("Sites API returned 403, trying M365 Groups API...")
            use_groups_fallback = True
        else:
            print_error(f"Failed to get sites: {e.code} - {e.reason}")
    except Exception as e:
        print_error(f"Error getting sites: {e}")
    
    # Fall back to M365 Groups if Sites API failed with 403
    if use_groups_fallback or not sites:
        groups_sites = get_m365_groups_with_sites(access_token)
        if groups_sites:
            print_success(f"Found {len(groups_sites)} sites via M365 Groups API")
            # Filter out system/personal sites
            return filter_writable_sites(groups_sites)
    
    # Filter out system/personal sites
    return filter_writable_sites(sites)

def create_folder_in_sharepoint(
    site_id: str,
    folder_name: str,
    access_token: str
) -> bool:
    """Create a folder in SharePoint document library."""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
    
    folder_data = json.dumps({
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename"
    }).encode('utf-8')
    
    try:
        req = urllib.request.Request(url, data=folder_data, method="POST")
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            return response.status in [200, 201]
    except urllib.error.HTTPError as e:
        if e.code == 409:  # Conflict - folder already exists
            return True
        print_error(f"Failed to create folder {folder_name}: {e.code}")
        return False
    except Exception as e:
        print_error(f"Error creating folder {folder_name}: {e}")
        return False

def generate_random_past_date(months_back: int = 12) -> datetime:
    """Generate a random date within the past N months."""
    now = datetime.now()
    days_back = random.randint(1, months_back * 30)
    return now - timedelta(days=days_back)


def generate_random_modified_date(created_date: datetime) -> datetime:
    """Generate a random modified date between created date and now."""
    now = datetime.now()
    if created_date >= now:
        return now
    days_diff = (now - created_date).days
    if days_diff <= 0:
        return now
    days_after_created = random.randint(0, days_diff)
    return created_date + timedelta(days=days_after_created)


def get_sharepoint_site_url(site_id: str, access_token: str) -> Optional[str]:
    """Get the SharePoint site URL from site ID."""
    try:
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            web_url = data.get("webUrl")
            return web_url if web_url else None
    except Exception:
        return None


def get_list_item_id(site_id: str, file_path: str, access_token: str) -> Optional[str]:
    """Get the list item ID for a file in SharePoint.
    
    Uses the Graph API with $select to explicitly request sharepointIds.
    """
    try:
        encoded_path = urllib.parse.quote(file_path)
        # Request sharepointIds explicitly using $select
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{encoded_path}?$select=id,name,sharepointIds"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            # Get the SharePoint list item info
            sp_item = data.get("sharepointIds", {})
            list_item_id = sp_item.get("listItemId")
            if list_item_id:
                return list_item_id
            
            # Fallback: try to get the item ID from the listItem endpoint
            item_id = data.get("id")
            if item_id:
                # Try to get the list item directly
                list_item_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/listItem?$select=id"
                list_req = urllib.request.Request(list_item_url)
                list_req.add_header("Authorization", f"Bearer {access_token}")
                with urllib.request.urlopen(list_req, timeout=30) as list_response:
                    list_data = json.loads(list_response.read().decode())
                    return list_data.get("id")
            return None
    except Exception:
        return None


def update_file_metadata_via_graph(
    site_id: str,
    list_item_id: str,
    access_token: str,
    created_date: Optional[datetime] = None,
    modified_date: Optional[datetime] = None,
    author_id: Optional[str] = None,
    editor_id: Optional[str] = None
) -> bool:
    """Update file metadata (timestamps) via Microsoft Graph API.
    
    Uses the Graph API to update list item fields for Created/Modified dates.
    Note: SharePoint's Created/Modified fields are system-managed and may not be
    directly updatable via Graph API. This function attempts to update custom
    date fields if available.
    
    Falls back gracefully if the operation fails.
    """
    try:
        # Build the update payload for list item fields
        # Note: Standard Created/Modified fields are read-only in SharePoint
        # We'll try to update them anyway - some tenants may allow it
        update_data = {"fields": {}}
        
        if created_date:
            # Try multiple field names that might work
            update_data["fields"]["Created"] = created_date.strftime("%Y-%m-%dT%H:%M:%SZ")
        if modified_date:
            update_data["fields"]["Modified"] = modified_date.strftime("%Y-%m-%dT%H:%M:%SZ")
        
        if not update_data["fields"]:
            return True  # Nothing to update
        
        # Graph API endpoint for updating list item
        # First, we need to get the list ID for the Documents library
        list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists?$filter=displayName eq 'Documents'&$select=id"
        list_req = urllib.request.Request(list_url)
        list_req.add_header("Authorization", f"Bearer {access_token}")
        
        list_id = None
        try:
            with urllib.request.urlopen(list_req, timeout=30) as list_response:
                list_data = json.loads(list_response.read().decode())
                lists = list_data.get("value", [])
                if lists:
                    list_id = lists[0].get("id")
        except Exception:
            pass
        
        if not list_id:
            return False
        
        # Update the list item
        update_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{list_item_id}/fields"
        
        payload = json.dumps(update_data["fields"]).encode('utf-8')
        
        req = urllib.request.Request(update_url, data=payload, method="PATCH")
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")
        
        with urllib.request.urlopen(req, timeout=30) as response:
            return response.status in [200, 204]
    except Exception:
        # Silently fail - file was still uploaded, just without custom metadata
        return False


# Keep the old function name as an alias for backwards compatibility
def update_file_metadata_via_rest(
    site_url: str,
    list_item_id: str,
    access_token: str,
    created_date: Optional[datetime] = None,
    modified_date: Optional[datetime] = None,
    author_id: Optional[str] = None,
    editor_id: Optional[str] = None
) -> bool:
    """Legacy function - redirects to Graph API version.
    
    Note: This function signature is kept for backwards compatibility but
    the site_url parameter is no longer used. The site_id should be passed
    to update_file_metadata_via_graph instead.
    """
    # This function is no longer used directly - the upload function
    # now calls update_file_metadata_via_graph with site_id
    return False


# Global counters for REST API metadata updates
_metadata_update_success = 0
_metadata_update_failed = 0
_metadata_update_skipped = 0
_metadata_skip_no_site_url = 0
_metadata_skip_no_list_item = 0


def reset_metadata_counters():
    """Reset the metadata update counters."""
    global _metadata_update_success, _metadata_update_failed, _metadata_update_skipped
    global _metadata_skip_no_site_url, _metadata_skip_no_list_item
    _metadata_update_success = 0
    _metadata_update_failed = 0
    _metadata_update_skipped = 0
    _metadata_skip_no_site_url = 0
    _metadata_skip_no_list_item = 0


def get_metadata_stats() -> Dict[str, int]:
    """Get the metadata update statistics."""
    return {
        "success": _metadata_update_success,
        "failed": _metadata_update_failed,
        "skipped": _metadata_update_skipped,
        "skip_no_site_url": _metadata_skip_no_site_url,
        "skip_no_list_item": _metadata_skip_no_list_item
    }


def upload_file_to_sharepoint(
    site_id: str,
    folder_path: str,
    file_name: str,
    file_content: bytes,
    access_token: str,
    user: Optional[Dict[str, Any]] = None,
    set_custom_dates: bool = True
) -> bool:
    """Upload a file to SharePoint using Microsoft Graph API.
    
    Optionally sets custom timestamps via SharePoint REST API for realism.
    Tracks metadata update success/failure for reporting.
    """
    global _metadata_update_success, _metadata_update_failed, _metadata_update_skipped
    global _metadata_skip_no_site_url, _metadata_skip_no_list_item
    # Encode the file path
    if folder_path:
        file_path = f"{folder_path}/{file_name}"
        encoded_path = urllib.parse.quote(file_path)
    else:
        file_path = file_name
        encoded_path = urllib.parse.quote(file_name)
    
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{encoded_path}:/content"
    
    try:
        req = urllib.request.Request(url, data=file_content, method="PUT")
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/octet-stream")
        
        with urllib.request.urlopen(req, timeout=60) as response:
            upload_success = response.status in [200, 201]
            
            # Try to set custom timestamps if upload succeeded
            if upload_success and set_custom_dates:
                try:
                    # Generate random past dates for realism
                    created_date = generate_random_past_date(months_back=12)
                    modified_date = generate_random_modified_date(created_date)
                    
                    # Get list item ID for the uploaded file
                    list_item_id = get_list_item_id(site_id, file_path, access_token)
                    if list_item_id:
                        # Update metadata via Graph API (more reliable than REST API)
                        metadata_result = update_file_metadata_via_graph(
                            site_id=site_id,
                            list_item_id=list_item_id,
                            access_token=access_token,
                            created_date=created_date,
                            modified_date=modified_date,
                            author_id=user.get("id") if user else None
                        )
                        if metadata_result:
                            _metadata_update_success += 1
                        else:
                            _metadata_update_failed += 1
                    else:
                        _metadata_update_skipped += 1
                        _metadata_skip_no_list_item += 1
                except Exception:
                    _metadata_update_failed += 1
            elif upload_success:
                _metadata_update_skipped += 1  # Custom dates disabled
            
            return upload_success
    except urllib.error.HTTPError as e:
        if e.code == 409:  # Conflict - file already exists
            return True
        # Don't print error for every file to avoid spam
        return False
    except Exception:
        return False

# ============================================================================
# FILE GENERATION
# ============================================================================

def generate_files_for_site(
    site: Dict[str, Any],
    num_files: int,
    access_token: str
) -> Tuple[int, int]:
    """Generate and upload files to a SharePoint site."""
    site_id = site.get("id", "")
    site_name = site.get("name", "Unknown")
    site_type = get_site_type(site_name)
    
    templates = FILE_TEMPLATES.get(site_type, FILE_TEMPLATES["default"])
    folders = templates["folders"]
    file_templates = templates["files"]
    
    success_count = 0
    fail_count = 0
    
    # Create folders first
    for folder in folders:
        create_folder_in_sharepoint(site_id, folder, access_token)
    
    # Generate and upload files
    for i in range(num_files):
        # Pick a random template and folder
        template = random.choice(file_templates)
        folder = random.choice(folders)
        
        # Generate file name
        file_name = generate_file_name(template, site_type)
        file_type = template["type"]
        
        # Create file content
        file_content = create_file_content(file_type, file_name, site_type)
        
        # Upload file
        if upload_file_to_sharepoint(site_id, folder, file_name, file_content, access_token):
            success_count += 1
        else:
            fail_count += 1
        
        # Update progress
        print_progress(i + 1, num_files, f"{site_name}: {file_name[:30]}...")
    
    print()  # New line after progress bar
    return success_count, fail_count


def generate_files_for_site_with_users(
    site: Dict[str, Any],
    num_files: int,
    access_token: str,
    users: List[Dict[str, Any]],
    groups: List[Dict[str, Any]]
) -> Tuple[int, int]:
    """Generate and upload files with Azure AD user attribution.
    
    Creates:
    - User-named folders (e.g., "John Smith")
    - Group-named folders (e.g., "Marketing Team")
    - Files with user names in the filename
    """
    site_id = site.get("id", "")
    site_name = site.get("name", "Unknown")
    site_type = get_site_type(site_name)
    
    success_count = 0
    fail_count = 0
    
    # Get department-specific templates as fallback
    templates = FILE_TEMPLATES.get(site_type, FILE_TEMPLATES["default"])
    dept_folders = templates["folders"]
    
    # Create department folders
    for folder in dept_folders:
        create_folder_in_sharepoint(site_id, folder, access_token)
    
    # Create user folders (select random subset of users)
    user_folders = []
    if users:
        selected_users = random.sample(users, min(5, len(users)))
        for user in selected_users:
            user_name = user.get("displayName", "User")
            folder_name = "".join(c for c in user_name if c.isalnum() or c in " -_").strip()
            if create_folder_in_sharepoint(site_id, folder_name, access_token):
                user_folders.append({"folder": folder_name, "user": user})
    
    # Create group folders (select random subset of groups)
    group_folders = []
    if groups:
        selected_groups = random.sample(groups, min(3, len(groups)))
        for group in selected_groups:
            group_name = group.get("displayName", "Team")
            folder_name = "".join(c for c in group_name if c.isalnum() or c in " -_").strip()
            if create_folder_in_sharepoint(site_id, folder_name, access_token):
                group_folders.append({"folder": folder_name, "group": group})
    
    # Combine all available folders
    all_folders = dept_folders.copy()
    all_folders.extend([uf["folder"] for uf in user_folders])
    all_folders.extend([gf["folder"] for gf in group_folders])
    
    # Generate and upload files
    for i in range(num_files):
        # Decide file type: 60% user-named, 20% shared, 20% department
        file_type_choice = random.random()
        
        if file_type_choice < 0.6 and users:
            # User-named file
            user = random.choice(users)
            template = random.choice(USER_FILE_TEMPLATES)
            file_name = generate_user_file_name(template, user)
            file_type = template["type"]
            
            # Put in user's folder if available, otherwise random folder
            if user_folders:
                matching_folders = [uf["folder"] for uf in user_folders
                                   if user.get("displayName", "") in uf["folder"]]
                folder = matching_folders[0] if matching_folders else random.choice(all_folders)
            else:
                folder = random.choice(all_folders)
                
        elif file_type_choice < 0.8 and users:
            # Shared file (attributed to a user)
            user = random.choice(users)
            template = random.choice(SHARED_FILE_TEMPLATES)
            file_name = generate_user_file_name(template, user)
            file_type = template["type"]
            
            # Put in group folder if available, otherwise department folder
            if group_folders:
                folder = random.choice([gf["folder"] for gf in group_folders])
            else:
                folder = random.choice(dept_folders)
        else:
            # Department file (original behavior)
            dept_templates = templates["files"]
            template = random.choice(dept_templates)
            file_name = generate_file_name(template, site_type)
            file_type = template["type"]
            folder = random.choice(dept_folders)
        
        # Create file content
        file_content = create_file_content(file_type, file_name, site_type)
        
        # Upload file
        if upload_file_to_sharepoint(site_id, folder, file_name, file_content, access_token):
            success_count += 1
        else:
            fail_count += 1
        
        # Update progress
        print_progress(i + 1, num_files, f"{site_name}: {file_name[:30]}...")
    
    print()  # New line after progress bar
    return success_count, fail_count


def distribute_files_with_users(
    sites: List[Dict[str, Any]],
    total_files: int,
    access_token: str,
    users: List[Dict[str, Any]],
    groups: List[Dict[str, Any]]
) -> Tuple[int, int]:
    """Distribute files with Azure AD user attribution across multiple sites."""
    if not sites:
        print_error("No sites available")
        return 0, 0
    
    total_success = 0
    total_fail = 0
    
    # Calculate files per site (with some randomness)
    files_per_site = []
    remaining = total_files
    
    for i, site in enumerate(sites):
        if i == len(sites) - 1:
            files_per_site.append(remaining)
        else:
            avg = remaining // (len(sites) - i)
            variance = max(1, avg // 3)
            count = max(1, random.randint(avg - variance, avg + variance))
            count = min(count, remaining - (len(sites) - i - 1))
            files_per_site.append(count)
            remaining -= count
    
    # Upload files to each site
    for site, num_files in zip(sites, files_per_site):
        site_name = site.get("displayName", site.get("name", "Unknown"))
        print_info(f"Uploading {num_files} files to: {site_name}")
        
        success, fail = generate_files_for_site_with_users(
            site, num_files, access_token, users, groups
        )
        total_success += success
        total_fail += fail
    
    return total_success, total_fail


def select_population_mode(users: List[Dict[str, Any]], groups: List[Dict[str, Any]]) -> str:
    """Let user select the file population mode."""
    print()
    print(f"  {Colors.WHITE}Select file population mode:{Colors.NC}")
    print()
    
    if users or groups:
        # Option 1: Azure AD Realistic Mix (Recommended)
        print(f"  {Colors.CYAN}[1]{Colors.NC} Azure AD Realistic Mix {Colors.GREEN}(Recommended){Colors.NC}")
        print(f"      {Colors.BLUE}Source:{Colors.NC} 100% Azure AD users/groups")
        if users:
            print(f"      {Colors.GREEN}✓{Colors.NC} {len(users)} Azure AD users discovered")
        if groups:
            print(f"      {Colors.GREEN}✓{Colors.NC} {len(groups)} Azure AD groups discovered")
        print(f"      {Colors.WHITE}60% personal + 20% shared + 20% collaborative files{Colors.NC}")
        print(f"      {Colors.WHITE}Consistent with Azure AD site ownership{Colors.NC}")
        print()
        
        # Option 2: Department Files Only
        print(f"  {Colors.CYAN}[2]{Colors.NC} Department Files Only")
        print(f"      {Colors.BLUE}Source:{Colors.NC} Config templates (FILE_TEMPLATES)")
        print(f"      {Colors.WHITE}Official documents (Budget_Report.xlsx, Policy_v2.pdf){Colors.NC}")
        print()
        
        # Option 3: Combined Sources Mix
        print(f"  {Colors.CYAN}[3]{Colors.NC} Combined Sources Mix")
        print(f"      {Colors.BLUE}Source:{Colors.NC} 70% Azure AD + 30% Config templates")
        print(f"      {Colors.WHITE}Maximum variety - combines both sources{Colors.NC}")
        print()
        
        # Option 4: User-Named Files Only
        print(f"  {Colors.CYAN}[4]{Colors.NC} User-Named Files Only")
        print(f"      {Colors.BLUE}Source:{Colors.NC} 100% Azure AD users/groups")
        print(f"      {Colors.WHITE}Example: \"John Smith - Meeting Notes.docx\"{Colors.NC}")
        print()
    else:
        # No Azure AD users/groups found - show limited options
        print(f"  {Colors.CYAN}[1]{Colors.NC} Department Files Only")
        print(f"      {Colors.BLUE}Source:{Colors.NC} Config templates (FILE_TEMPLATES)")
        print(f"      {Colors.WHITE}Official documents (Budget_Report.xlsx, Policy_v2.pdf){Colors.NC}")
        print()
        print(f"  {Colors.YELLOW}[2-4]{Colors.NC} Azure AD Modes {Colors.YELLOW}(No users/groups found){Colors.NC}")
        print(f"      {Colors.WHITE}Requires User.Read.All and Group.Read.All permissions{Colors.NC}")
        print()
    
    while True:
        choice = input(f"  {Colors.YELLOW}Enter choice (1-4, Q to quit):{Colors.NC} ").strip().lower()
        
        if choice == 'q':
            return "quit"
        elif choice == '1' and (users or groups):
            return "azure_ad_mixed"
        elif choice == '1':
            return "generic"
        elif choice == '2' and (users or groups):
            return "generic"
        elif choice == '3' and (users or groups):
            return "combined"
        elif choice == '4' and (users or groups):
            return "azure_ad"
        elif choice in ['2', '3', '4']:
            print_warning("Azure AD mode requires discovered users/groups. Using Department Files mode.")
            return "generic"
        else:
            print_warning("Invalid choice. Please enter 1, 2, 3, 4, or Q.")

def distribute_files_across_sites(
    sites: List[Dict[str, Any]],
    total_files: int,
    access_token: str
) -> Tuple[int, int]:
    """Distribute files randomly across multiple sites."""
    if not sites:
        print_error("No sites available")
        return 0, 0
    
    total_success = 0
    total_fail = 0
    
    # Calculate files per site (with some randomness)
    files_per_site = []
    remaining = total_files
    
    for i, site in enumerate(sites):
        if i == len(sites) - 1:
            # Last site gets remaining files
            files_per_site.append(remaining)
        else:
            # Random distribution with minimum of 1
            avg = remaining // (len(sites) - i)
            variance = max(1, avg // 3)
            count = max(1, random.randint(avg - variance, avg + variance))
            count = min(count, remaining - (len(sites) - i - 1))  # Ensure others get at least 1
            files_per_site.append(count)
            remaining -= count
    
    # Upload files to each site
    for site, num_files in zip(sites, files_per_site):
        site_name = site.get("displayName", site.get("name", "Unknown"))
        print_info(f"Uploading {num_files} files to: {site_name}")
        
        success, fail = generate_files_for_site(site, num_files, access_token)
        total_success += success
        total_fail += fail
    
    return total_success, total_fail

# ============================================================================
# MAIN FUNCTION
# ============================================================================

def main() -> None:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Populate SharePoint sites with realistic files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python populate_files.py                      # Interactive mode
    python populate_files.py --files 100          # Create 100 files across all sites
    python populate_files.py --files 50 --site hr # Create 50 files in sites containing 'hr'
    python populate_files.py --list-sites         # List available SharePoint sites
        """
    )
    
    parser.add_argument(
        '-f', '--files',
        type=int,
        default=0,
        metavar='COUNT',
        help=f'Number of files to create (1-{MAX_FILES})'
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
    
    args = parser.parse_args()
    
    # Clear screen and show banner
    os.system('cls' if os.name == 'nt' else 'clear')
    print_banner("SHAREPOINT FILE POPULATION")
    
    print(f"  {Colors.WHITE}This script populates SharePoint sites with realistic files{Colors.NC}")
    print(f"  {Colors.WHITE}to simulate an actual organization's document structure.{Colors.NC}")
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
        print_info("  - Sites.ReadWrite.All")
        print_info("  - Files.ReadWrite.All")
        sys.exit(1)
    
    print_success("Access token obtained")
    
    # Step 4: Get SharePoint sites
    print_step(4, "Discover SharePoint Sites")
    
    sites = get_sharepoint_sites(access_token)
    
    if not sites:
        print_error("No SharePoint sites found")
        print_info("Make sure you have access to SharePoint sites in your tenant")
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
            site_type = get_site_type(name)
            print(f"    {i:3}. {name} ({site_type})")
        print()
        sys.exit(0)
    
    # Step 5: Discover Azure AD Users & Groups
    print_step(5, "Discover Azure AD Users & Groups")
    
    print_info("Discovering Azure AD users and groups for realistic file naming...")
    users = discover_azure_ad_users(access_token)
    groups = discover_azure_ad_groups(access_token)
    
    if users:
        print_success(f"Found {len(users)} Azure AD users")
    else:
        print_warning("No Azure AD users found (User.Read.All permission may be missing)")
    
    if groups:
        print_success(f"Found {len(groups)} Azure AD groups")
    else:
        print_warning("No Azure AD groups found (Group.Read.All permission may be missing)")
    
    # Step 6: Select Population Mode
    print_step(6, "Select File Population Mode")
    
    population_mode = select_population_mode(users, groups)
    if population_mode == "quit":
        print_warning("Operation cancelled.")
        sys.exit(0)
    
    print_success(f"Selected mode: {population_mode.replace('_', ' ').title()}")
    
    # Step 7: Get file count
    print_step(7, "Configure File Generation")
    
    num_files = args.files
    if num_files <= 0:
        print()
        print(f"  {Colors.WHITE}How many files would you like to create?{Colors.NC}")
        print(f"  (Files will be distributed across {len(sites)} sites)")
        print(f"  {Colors.CYAN}(Enter Q to quit){Colors.NC}")
        print()
        while True:
            user_input = input(f"  Enter number of files (1-{MAX_FILES}, Q to quit): ").strip()
            if user_input.lower() == 'q':
                print_warning("Operation cancelled.")
                sys.exit(0)
            try:
                num_files = int(user_input)
                if 1 <= num_files <= MAX_FILES:
                    break
                print_warning(f"Please enter a number between 1 and {MAX_FILES}")
            except ValueError:
                print_warning("Please enter a valid number or Q to quit")
    
    # Validate file count
    if num_files < 1 or num_files > MAX_FILES:
        print_error(f"File count must be between 1 and {MAX_FILES}")
        sys.exit(1)
    
    print_success(f"Will create {num_files} files across {len(sites)} sites")
    
    # Step 8: Confirm
    print_step(8, "Confirm File Generation")
    
    # Map mode to friendly display name
    mode_names = {
        "generic": "Department Files Only",
        "azure_ad": "User-Named Files Only (100% Azure AD)",
        "azure_ad_mixed": "Azure AD Realistic Mix (100% Azure AD)",
        "combined": "Combined Sources Mix (Recommended)"
    }
    mode_display = mode_names.get(population_mode, population_mode)
    
    print()
    print(f"  {Colors.WHITE}Summary:{Colors.NC}")
    print(f"    - Population mode: {mode_display}")
    print(f"    - Total files to create: {num_files}")
    print(f"    - Target sites: {len(sites)}")
    print(f"    - Average files per site: ~{num_files // len(sites)}")
    if population_mode in ["azure_ad", "azure_ad_mixed", "combined"] and users:
        print(f"    - Azure AD users: {len(users)}")
    if population_mode in ["azure_ad", "azure_ad_mixed", "combined"] and groups:
        print(f"    - Azure AD groups: {len(groups)}")
    print()
    print(f"  {Colors.WHITE}Sites to populate:{Colors.NC}")
    for site in sites[:10]:  # Show first 10
        name = site.get("displayName", site.get("name", "Unknown"))
        print(f"    - {name}")
    if len(sites) > 10:
        print(f"    ... and {len(sites) - 10} more")
    print()
    
    confirm = input("  Proceed with file generation? (y/n): ").strip().lower()
    if confirm != 'y':
        print_warning("Operation cancelled")
        sys.exit(0)
    
    # Step 9: Generate files
    print_step(9, "Generate and Upload Files")
    
    print()
    # Reset metadata counters before starting
    reset_metadata_counters()
    start_time = datetime.now()
    
    # Use appropriate distribution function based on mode
    if population_mode == "generic":
        # Option 1: Department files only (from config templates)
        success, fail = distribute_files_across_sites(sites, num_files, access_token)
    elif population_mode == "azure_ad":
        # Option 2: User-named files only (100% Azure AD - personal documents)
        success, fail = distribute_files_with_users(sites, num_files, access_token, users, groups)
    elif population_mode == "azure_ad_mixed":
        # Option 3: Azure AD Realistic Mix (100% Azure AD - varied file types)
        # Uses the same function but the generate_files_for_site_with_users already
        # creates a mix of personal (60%), shared (20%), and collaborative (20%) files
        print_info("Azure AD Realistic Mix: 60% personal + 20% shared + 20% collaborative files")
        success, fail = distribute_files_with_users(sites, num_files, access_token, users, groups)
    else:  # combined mode - Option 4 (Recommended)
        # Split files: 30% department/generic, 70% Azure AD user-named
        # This reflects real organizations where most files are created by individuals
        # but also have official department documents
        generic_count = int(num_files * 0.30)
        azure_count = num_files - generic_count
        
        print_info(f"Combined Sources Mix: {azure_count} Azure AD files + {generic_count} department files")
        
        success1, fail1 = distribute_files_across_sites(sites, generic_count, access_token)
        success2, fail2 = distribute_files_with_users(sites, azure_count, access_token, users, groups)
        success = success1 + success2
        fail = fail1 + fail2
    
    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds()
    
    # Step 10: Summary
    print_step(10, "Summary")
    
    print()
    print(f"  {Colors.WHITE}File Generation Complete!{Colors.NC}")
    print()
    print(f"    {Colors.GREEN}✓ Successfully uploaded:{Colors.NC} {success} files")
    if fail > 0:
        print(f"    {Colors.RED}✗ Failed:{Colors.NC} {fail} files")
    print(f"    {Colors.BLUE}⏱ Duration:{Colors.NC} {duration:.1f} seconds")
    if duration > 0:
        print(f"    {Colors.BLUE}📊 Rate:{Colors.NC} {success / duration:.1f} files/second")
    print()
    
    # Display metadata update statistics (REST API for custom timestamps)
    metadata_stats = get_metadata_stats()
    total_metadata_attempts = metadata_stats["success"] + metadata_stats["failed"] + metadata_stats["skipped"]
    if total_metadata_attempts > 0:
        print(f"  {Colors.WHITE}Custom Timestamps (REST API):{Colors.NC}")
        if metadata_stats["success"] > 0:
            print(f"    {Colors.GREEN}✓ Timestamps set:{Colors.NC} {metadata_stats['success']} files")
        if metadata_stats["failed"] > 0:
            print(f"    {Colors.RED}✗ REST API failed:{Colors.NC} {metadata_stats['failed']} files")
            print(f"      {Colors.YELLOW}(May require Sites.FullControl.All permission){Colors.NC}")
        if metadata_stats["skipped"] > 0:
            print(f"    {Colors.YELLOW}⊘ Skipped:{Colors.NC} {metadata_stats['skipped']} files")
            # Show detailed breakdown of why skipped
            if metadata_stats.get("skip_no_site_url", 0) > 0:
                print(f"      - No site URL: {metadata_stats['skip_no_site_url']} (Graph API lookup failed)")
            if metadata_stats.get("skip_no_list_item", 0) > 0:
                print(f"      - No list item ID: {metadata_stats['skip_no_list_item']} (SharePoint IDs not available)")
        print()
    
    if success > 0:
        print_success("Files have been uploaded to your SharePoint sites!")
        print_info("You can view them in the SharePoint document libraries")
    else:
        print_error("No files were uploaded successfully")
        print_info("Check your permissions and try again")


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