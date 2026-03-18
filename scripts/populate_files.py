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
import concurrent.futures
import json
import os
import platform
import random
import shutil
import subprocess
import sys
import tempfile
import base64
import threading
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any, Set

# ============================================================================
# CONFIGURATION
# ============================================================================

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_DIR = SCRIPT_DIR.parent
CONFIG_DIR = PROJECT_DIR / "config"
DEFAULT_CONFIG_FILE = CONFIG_DIR / "sites.json"
ENVIRONMENTS_FILE = CONFIG_DIR / "environments.json"

# Maximum files that can be created in one run
MAX_FILES = 1000
METADATA_UPDATE_WORKERS = 8
VARIATION_LEVEL = "medium"

VARIATION_PROFILES = {
    "low": {
        "dynamic_file_patterns": 2,
        "extra_folder_min": 2,
        "extra_folder_max": 4,
        "rev_suffix_probability": 0.10,
        "version_suffix_probability": 0.06,
    },
    "medium": {
        "dynamic_file_patterns": 5,
        "extra_folder_min": 5,
        "extra_folder_max": 10,
        "rev_suffix_probability": 0.18,
        "version_suffix_probability": 0.12,
    },
    "high": {
        "dynamic_file_patterns": 8,
        "extra_folder_min": 8,
        "extra_folder_max": 14,
        "rev_suffix_probability": 0.28,
        "version_suffix_probability": 0.16,
    },
}

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
# CONFIGURATION & DEPARTMENT TEMPLATE MANAGEMENT
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


def load_sites_config() -> Dict[str, Any]:
    """Load the full sites.json configuration."""
    try:
        with open(DEFAULT_CONFIG_FILE, 'r') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return {}


def load_deployment_tracking() -> Dict[str, Any]:
    """Load deployment tracking settings from sites.json."""
    config = load_sites_config()
    return config.get("deployment_tracking", {})


def get_configured_site_names() -> Set[str]:
    """Get the set of site names defined in sites.json (lowercase for matching)."""
    config = load_sites_config()
    sites = config.get("sites", [])
    names = set()
    for site in sites:
        display_name = site.get("display_name", site.get("name", ""))
        if display_name:
            names.add(display_name.lower())
    return names


def filter_sites_by_config(sites: List[Dict[str, Any]], configured_names: Set[str]) -> List[Dict[str, Any]]:
    """Filter sites to only those matching names in sites.json."""
    matching = []
    for site in sites:
        site_name = site.get("displayName", site.get("name", "")).lower()
        if site_name in configured_names:
            matching.append(site)
    return matching


def filter_sites_by_deployment_id(sites: List[Dict[str, Any]], deployment_id: str) -> List[Dict[str, Any]]:
    """Filter sites to only those with matching deployment ID in description."""
    import re
    matching = []
    for site in sites:
        description = site.get("description", "") or ""
        # Check for deployment ID pattern in description
        if f"Ref: {deployment_id}" in description:
            matching.append(site)
    return matching


def select_site_scope(sites: List[Dict[str, Any]], access_token: str) -> Tuple[List[Dict[str, Any]], str]:
    """Present scope selection menu and return filtered sites.
    
    Returns:
        Tuple of (filtered_sites, scope_description)
    """
    # Load configuration
    configured_names = get_configured_site_names()
    tracking = load_deployment_tracking()
    deployment_id = tracking.get("deployment_id", "")
    
    print()
    print(f"  {Colors.CYAN}Select which sites to populate with files:{Colors.NC}")
    print()
    print(f"    {Colors.WHITE}[1]{Colors.NC} Filter by Deployment ID", end="")
    if deployment_id:
        print(f" {Colors.GREEN}(current: {deployment_id}){Colors.NC}")
    else:
        print(f" {Colors.DIM}(no ID configured){Colors.NC}")
    print(f"        {Colors.DIM}Process sites with matching deployment ID in description{Colors.NC}")
    print()
    print(f"    {Colors.WHITE}[2]{Colors.NC} Filter by sites.json names", end="")
    if configured_names:
        print(f" {Colors.GREEN}({len(configured_names)} sites){Colors.NC}")
    else:
        print(f" {Colors.DIM}(no sites configured){Colors.NC}")
    print(f"        {Colors.DIM}Process only sites defined in sites.json{Colors.NC}")
    print()
    print(f"    {Colors.WHITE}[3]{Colors.NC} Process ALL discovered sites {Colors.YELLOW}(use with caution!){Colors.NC}")
    print(f"        {Colors.DIM}Process all {len(sites)} sites found in tenant{Colors.NC}")
    print()
    print(f"    {Colors.WHITE}[Q]{Colors.NC} Cancel")
    print()
    
    scope_choice = input(f"  {Colors.CYAN}Select option (1-3, Q):{Colors.NC} ").strip().lower()
    
    if scope_choice == 'q':
        return [], "cancelled"
    
    if scope_choice == '1':
        # Filter by deployment ID
        if not deployment_id:
            print()
            print(f"  {Colors.YELLOW}No deployment ID configured.{Colors.NC}")
            print(f"  Enter a deployment ID (format: PRJ-XXXXXX) or press Enter to cancel:")
            user_id = input(f"  {Colors.CYAN}Deployment ID:{Colors.NC} ").strip().upper()
            
            if not user_id:
                return [], "cancelled"
            
            import re
            if not re.match(r'^PRJ-[A-Z0-9]{3,20}$', user_id):
                print(f"  {Colors.RED}✗{Colors.NC} Invalid format. Expected PRJ-XXXXX (3-20 alphanumeric characters)")
                return [], "invalid"
            
            deployment_id = user_id
        
        # Need to get M365 Groups to check descriptions (sites don't have descriptions)
        print()
        print_info(f"Filtering by deployment ID: {deployment_id}")
        
        # Get M365 Groups to match by description (with pagination)
        try:
            filter_param = urllib.parse.quote("groupTypes/any(c:c eq 'Unified')")
            url: Optional[str] = f"https://graph.microsoft.com/v1.0/groups?$filter={filter_param}&$select=id,displayName,description&$top=200"
            
            groups = []
            while url:
                req = urllib.request.Request(url)
                req.add_header("Authorization", f"Bearer {access_token}")
                req.add_header("Content-Type", "application/json")
                
                with urllib.request.urlopen(req, timeout=30) as response:
                    result = json.loads(response.read().decode())
                    groups.extend(result.get('value', []))
                    url = result.get('@odata.nextLink')
                    if url and len(groups) >= 2000:
                        break  # Safety limit
            
            # Find groups with matching deployment ID
            matching_group_names = set()
            for group in groups:
                description = group.get('description', '') or ''
                if f"Ref: {deployment_id}" in description:
                    matching_group_names.add(group.get('displayName', '').lower())
            
            if not matching_group_names:
                print(f"  {Colors.YELLOW}⚠{Colors.NC} No sites found with deployment ID: {deployment_id}")
                return [], "no_matches"
            
            # Filter sites by matching group names
            filtered = [s for s in sites if s.get("displayName", s.get("name", "")).lower() in matching_group_names]
            
            if filtered:
                print_success(f"Found {len(filtered)} sites matching deployment ID")
            return filtered, f"deployment ID: {deployment_id}"
            
        except Exception as e:
            print(f"  {Colors.RED}✗{Colors.NC} Failed to get groups: {e}")
            return [], "error"
    
    elif scope_choice == '2':
        # Filter by sites.json names
        if not configured_names:
            print(f"  {Colors.YELLOW}⚠{Colors.NC} No sites found in sites.json")
            return [], "no_config"
        
        print()
        print_info(f"Filtering by sites.json ({len(configured_names)} sites)...")
        
        filtered = filter_sites_by_config(sites, configured_names)
        
        if not filtered:
            print(f"  {Colors.YELLOW}⚠{Colors.NC} No discovered sites match sites.json configuration")
            print(f"  {Colors.DIM}This could mean the sites haven't been created yet.{Colors.NC}")
            return [], "no_matches"
        
        print_success(f"Found {len(filtered)} sites matching sites.json")
        return filtered, f"sites.json ({len(filtered)} sites)"
    
    elif scope_choice == '3':
        # Process all sites - confirm first
        print()
        print(f"  {Colors.YELLOW}{'─' * 50}{Colors.NC}")
        print(f"  {Colors.YELLOW}⚠ WARNING: This will populate ALL {len(sites)} sites!{Colors.NC}")
        print(f"  {Colors.YELLOW}  This includes sites you may not have created.{Colors.NC}")
        print(f"  {Colors.YELLOW}{'─' * 50}{Colors.NC}")
        print()
        confirm = input(f"  {Colors.YELLOW}Type 'YES' to confirm:{Colors.NC} ").strip()
        
        if confirm != 'YES':
            return [], "cancelled"
        
        return sites, f"all sites ({len(sites)})"
    
    else:
        print(f"  {Colors.RED}✗{Colors.NC} Invalid option")
        return [], "invalid"


def get_department_file_templates(config_path: Optional[Path] = None, warn_on_error: bool = False) -> Dict[str, Dict]:
    """Get file templates for departments using config baseline plus hardcoded extras.

    Baseline source is config/sites.json (or provided config path). For each department
    found in config, check if FILE_TEMPLATES has a custom template. If not, use a
    sensible default template. Hardcoded entries in FILE_TEMPLATES are preserved if
    they're not already covered by config departments.
    
    Returns a merged dictionary of department -> file template mappings.
    """
    template_path = config_path or DEFAULT_CONFIG_FILE
    baseline_departments: List[str] = []
    
    if template_path.exists():
        try:
            sites = read_sites_from_config(template_path)
            # Extract department names from config
            baseline_departments = [s.get("name", "") for s in sites if s.get("name", "")]
        except Exception as e:
            if warn_on_error:
                print(f"Warning: Could not read department baseline from {template_path}: {e}")
                print("Using built-in department templates only")
    
    # Create merged templates dictionary
    merged_templates: Dict[str, Dict] = {}
    seen_names: set = set()
    
    # Default template for unmapped departments
    default_template = {
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
    
    # First, add all baseline departments (config wins)
    for dept_name in baseline_departments:
        dept_lower = dept_name.lower()
        if dept_lower not in seen_names:
            # Use existing FILE_TEMPLATES entry if available, else use default
            merged_templates[dept_name] = FILE_TEMPLATES.get(dept_name, default_template)
            seen_names.add(dept_lower)
    
    # Then, add any hardcoded FILE_TEMPLATES entries not already in config
    for dept_name in FILE_TEMPLATES.keys():
        if dept_name != "default" and dept_name.lower() not in seen_names:
            merged_templates[dept_name] = FILE_TEMPLATES[dept_name]
            seen_names.add(dept_name.lower())
    
    # Always include default if it exists in FILE_TEMPLATES
    merged_templates["default"] = FILE_TEMPLATES.get("default", default_template)
    
    return merged_templates


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

# Enterprise-scale naming and taxonomy variables for higher realism
BUSINESS_UNITS = ["Corporate", "Shared_Services", "Regional_Ops", "Digital", "Risk", "Transformation"]
CONFIDENTIALITY_LEVELS = ["Public", "Internal", "Confidential", "Restricted"]
DOC_STATUSES = ["Draft", "In_Review", "Approved", "Final"]
INITIATIVES = ["Phoenix", "Atlas", "Horizon", "Aurora", "North_Star", "Velocity"]
ENTITY_TYPES = ["Policy", "Procedure", "Runbook", "Playbook", "Assessment", "Brief"]
REGION_CODES = ["NA", "EMEA", "APAC", "LATAM"]
CYCLES = ["Weekly", "Monthly", "Quarterly", "Annual"]

ENTERPRISE_COMMON_FILE_TEMPLATES = [
    {"name": "{confidentiality}_{status}_{entity}_{year}_{doc_id}.docx", "type": "word"},
    {"name": "{initiative}_Program_Update_Q{quarter}_{year}_{status}.pptx", "type": "powerpoint"},
    {"name": "KPI_Scorecard_{business_unit}_{month}_{year}.xlsx", "type": "excel"},
    {"name": "Governance_Minutes_{date}_{doc_id}.docx", "type": "word"},
    {"name": "Risk_Register_{region_code}_FY{fiscal_year}.xlsx", "type": "excel"},
    {"name": "Control_Testing_Evidence_{month}_{year}_{number}.pdf", "type": "pdf"},
]

SITE_ENTERPRISE_TEMPLATE_AUGMENTS = {
    "executive-leadership": [
        {"name": "SteerCo_Pack_Q{quarter}_{year}_{status}.pptx", "type": "powerpoint"},
        {"name": "Board_Action_Log_FY{fiscal_year}.xlsx", "type": "excel"},
    ],
    "finance-department": [
        {"name": "GL_Reconciliation_{month}_{year}_{doc_id}.xlsx", "type": "excel"},
        {"name": "Capex_Approval_{business_unit}_FY{fiscal_year}.pdf", "type": "pdf"},
    ],
    "it-department": [
        {"name": "Change_Record_{number}_{month}_{year}.docx", "type": "word"},
        {"name": "Incident_Postmortem_{initiative}_{date}.docx", "type": "word"},
    ],
    "legal-department": [
        {"name": "Contract_Deviation_Log_{year}.xlsx", "type": "excel"},
        {"name": "Regulatory_Advisory_{region_code}_{month}_{year}.pdf", "type": "pdf"},
    ],
    "operations-department": [
        {"name": "Site_Readiness_Checklist_{region_code}_{date}.xlsx", "type": "excel"},
        {"name": "Vendor_QBR_{vendor}_Q{quarter}_{year}.pptx", "type": "powerpoint"},
    ],
}

SITE_SUBJECTS = {
    "executive-leadership": ["Portfolio_Review", "Investment_Thesis", "Operating_Model", "Strategic_Priorities"],
    "human-resources": ["Workforce_Planning", "Talent_Review", "Learning_Pathway", "Compensation_Framework"],
    "finance-department": ["Close_Process", "Cost_Optimization", "Working_Capital", "Revenue_Assurance"],
    "claims-department": ["Case_Assessment", "Fraud_Review", "Claims_Quality", "Settlement_Strategy"],
    "it-department": ["Platform_Upgrade", "Service_Reliability", "Identity_Access", "Incident_Trend"],
    "marketing-department": ["Campaign_Attribution", "Audience_Strategy", "Brand_Performance", "Lead_Quality"],
    "sales-department": ["Pipeline_Health", "Deal_Desk", "Territory_Coverage", "Win_Loss_Analysis"],
    "legal-department": ["Contract_Risk", "Policy_Exception", "Regulatory_Tracking", "Matter_Prioritization"],
    "operations-department": ["Throughput_Optimization", "Supplier_Risk", "Process_Control", "Capacity_Model"],
    "product-management": ["Roadmap_Validation", "Feature_Adoption", "User_Journey", "Release_Readiness"],
    "customer-service": ["Escalation_Trends", "Knowledge_Quality", "Agent_Coaching", "Resolution_Time"],
    "default": ["Operational_Review", "Governance_Update", "Delivery_Status", "Program_Health"],
}

ENTERPRISE_DYNAMIC_FILE_PATTERNS = [
    {"name": "{cycle}_{subject}_{month}_{year}.docx", "type": "word"},
    {"name": "{subject}_Decision_Log_Q{quarter}_{year}.docx", "type": "word"},
    {"name": "{subject}_Dashboard_{region_code}_{month}_{year}.xlsx", "type": "excel"},
    {"name": "{subject}_Scorecard_FY{fiscal_year}.xlsx", "type": "excel"},
    {"name": "{initiative}_{subject}_SteerCo_Q{quarter}_{year}.pptx", "type": "powerpoint"},
    {"name": "{business_unit}_{subject}_Operating_Review_{month}_{year}.pptx", "type": "powerpoint"},
    {"name": "Control_Evidence_{subject}_{date}_{doc_id}.pdf", "type": "pdf"},
    {"name": "Audit_Pack_{subject}_FY{fiscal_year}_{status}.pdf", "type": "pdf"},
]

SITE_FOLDER_AUGMENTS = {
    "executive-leadership": ["SteerCo", "Portfolio Governance", "Investor Materials"],
    "human-resources": ["Talent Reviews", "Org Design", "People Analytics"],
    "finance-department": ["Close Calendar", "Controllership", "Treasury"],
    "claims-department": ["Severity Triage", "Escalated Claims", "Reserving"],
    "it-department": ["Service Operations", "Cloud Platform", "Architecture Review Board"],
    "marketing-department": ["Demand Generation", "Performance Marketing", "Content Operations"],
    "sales-department": ["Deal Desk", "Account Plans", "Quota Planning"],
    "legal-department": ["Legal Intake", "Regulatory Affairs", "Outside Counsel"],
    "operations-department": ["Service Delivery", "Continuous Improvement", "Site Governance"],
    "product-management": ["Portfolio Council", "Discovery", "Product Operations"],
    "customer-service": ["Voice of Customer", "Quality Assurance", "Service Operations"],
    "default": ["Governance", "Operations", "Workstreams"],
}

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
    fiscal_year = now.year + 1 if now.month >= 7 else now.year
    
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
        "{business_unit}": random.choice(BUSINESS_UNITS),
        "{confidentiality}": random.choice(CONFIDENTIALITY_LEVELS),
        "{status}": random.choice(DOC_STATUSES),
        "{initiative}": random.choice(INITIATIVES),
        "{entity}": random.choice(ENTITY_TYPES),
        "{region_code}": random.choice(REGION_CODES),
        "{cycle}": random.choice(CYCLES),
        "{subject}": random.choice(SITE_SUBJECTS.get(site_type, SITE_SUBJECTS["default"])),
        "{fiscal_year}": str(fiscal_year),
        "{doc_id}": f"DOC-{now.year}-{random.randint(10000, 99999)}",
    }
    
    for placeholder, value in replacements.items():
        name = name.replace(placeholder, value)

    profile = VARIATION_PROFILES.get(VARIATION_LEVEL, VARIATION_PROFILES["medium"])

    # Add light enterprise variation to avoid repetitive names at scale.
    if random.random() < profile["rev_suffix_probability"]:
        name = name.replace(".", f"_REV{random.randint(1, 7)}.", 1)
    elif random.random() < profile["version_suffix_probability"]:
        name = name.replace(".", f"_v{random.randint(1, 5)}.{random.randint(0, 9)}.", 1)
    
    return name


def build_template_pool(site_type: str) -> Dict[str, Any]:
    """Build a richer template pool using the enterprise variation pipeline.
    
    Uses merged department templates from config baseline + hardcoded extras
    to ensure sites.json changes automatically propagate.
    """
    # Get merged templates (config baseline + hardcoded, config wins on conflicts)
    merged = get_department_file_templates(warn_on_error=False)
    base = merged.get(site_type, merged.get("default", FILE_TEMPLATES["default"]))
    return {
        "folders": list(base.get("folders", [])),
        "files": build_file_pattern_pool(site_type, list(base.get("files", []))),
    }


def build_file_pattern_pool(site_type: str, base_files: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """Compose file patterns in one place for easier maintenance and tuning.

    Order is intentional:
    1. department baseline templates
    2. enterprise common templates
    3. enterprise dynamic templates
    4. department-specific enterprise augmentations
    """
    files = list(base_files)
    profile = VARIATION_PROFILES.get(VARIATION_LEVEL, VARIATION_PROFILES["medium"])

    files.extend(ENTERPRISE_COMMON_FILE_TEMPLATES)

    dynamic_patterns = list(ENTERPRISE_DYNAMIC_FILE_PATTERNS)
    random.shuffle(dynamic_patterns)
    files.extend(dynamic_patterns[: profile["dynamic_file_patterns"]])

    files.extend(SITE_ENTERPRISE_TEMPLATE_AUGMENTS.get(site_type, []))
    return files


def build_folder_candidate_pool(site_type: str) -> List[str]:
    """Build enterprise folder candidates independent of department defaults."""
    now = datetime.now()
    fiscal_year = now.year + 1 if now.month >= 7 else now.year
    current_quarter = (now.month - 1) // 3 + 1

    candidates = [
        f"FY{fiscal_year} Planning",
        f"Q{current_quarter}_{now.year} Priorities",
        f"Archive_{now.year - 1}",
        f"Archive_{now.year - 2}",
        "Working Drafts",
        "Approvals",
        "Governance Packs",
        "Governance",
        "Vendor Management",
        "Risk and Controls",
        f"Regional_{random.choice(REGION_CODES)}",
        f"Program_{random.choice(INITIATIVES)}",
        f"{random.choice(CONFIDENTIALITY_LEVELS)} Documents",
        f"{random.choice(CYCLES)} Reporting",
        f"{random.choice(BUSINESS_UNITS)} Operations",
    ]
    candidates.extend(SITE_FOLDER_AUGMENTS.get(site_type, SITE_FOLDER_AUGMENTS["default"]))
    return candidates


def build_folder_pool(site_type: str, base_folders: List[str]) -> List[str]:
    """Expand base department folders with enterprise-like organizational structure."""
    profile = VARIATION_PROFILES.get(VARIATION_LEVEL, VARIATION_PROFILES["medium"])
    extra_candidates = build_folder_candidate_pool(site_type)

    folder_pool = list(base_folders)
    random.shuffle(extra_candidates)
    folder_pool.extend(
        extra_candidates[: random.randint(profile["extra_folder_min"], profile["extra_folder_max"])]
    )

    # Preserve order while de-duplicating.
    deduped: List[str] = []
    seen = set()
    for folder in folder_pool:
        if folder not in seen:
            seen.add(folder)
            deduped.append(folder)
    return deduped

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


_site_drive_endpoint_cache: Dict[str, str] = {}
_site_drive_endpoint_lock = threading.Lock()


def resolve_site_drive_endpoint(
    site_id: str,
    access_token: str,
    force_refresh: bool = False
) -> str:
    """Resolve a stable drive endpoint for a site.

    Some sites intermittently return 404 for /sites/{id}/drive even when a
    document library exists. We resolve and cache a concrete /drives/{id}
    endpoint when possible, with a fallback to the original site-drive path.
    """
    default_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"

    if not force_refresh:
        with _site_drive_endpoint_lock:
            cached_endpoint = _site_drive_endpoint_cache.get(site_id)
        if cached_endpoint:
            return cached_endpoint

    resolved_endpoint = default_endpoint

    try:
        drive_req = urllib.request.Request(f"{default_endpoint}?$select=id")
        drive_req.add_header("Authorization", f"Bearer {access_token}")

        with urllib.request.urlopen(drive_req, timeout=30) as response:
            drive_data = json.loads(response.read().decode())
            drive_id = drive_data.get("id")
            if drive_id:
                resolved_endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
    except Exception:
        pass

    if resolved_endpoint == default_endpoint:
        try:
            drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives?$select=id,driveType,name&$top=50"
            drives_req = urllib.request.Request(drives_url)
            drives_req.add_header("Authorization", f"Bearer {access_token}")

            with urllib.request.urlopen(drives_req, timeout=30) as response:
                drives_data = json.loads(response.read().decode())
                drives = drives_data.get("value", [])

                preferred_drive = next(
                    (drive for drive in drives if drive.get("driveType") == "documentLibrary" and drive.get("id")),
                    None
                )
                selected_drive = preferred_drive or next((drive for drive in drives if drive.get("id")), None)

                if selected_drive:
                    resolved_endpoint = f"https://graph.microsoft.com/v1.0/drives/{selected_drive['id']}"
        except Exception:
            pass

    with _site_drive_endpoint_lock:
        _site_drive_endpoint_cache[site_id] = resolved_endpoint

    return resolved_endpoint

def create_folder_in_sharepoint(
    site_id: str,
    folder_name: str,
    access_token: str
) -> bool:
    """Create a folder in SharePoint document library."""
    def _create_folder_at_endpoint(drive_endpoint: str) -> bool:
        url = f"{drive_endpoint}/root/children"

        req = urllib.request.Request(url, data=folder_data, method="POST")
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/json")

        with urllib.request.urlopen(req, timeout=30) as response:
            return response.status in [200, 201]

    def _queue_folder_metadata_update() -> None:
        created_date = generate_random_past_date(months_back=12)
        modified_date = generate_random_modified_date(created_date)
        queue_metadata_update(
            site_id=site_id,
            file_path=folder_name,
            graph_access_token=access_token,
            created_date=created_date,
            modified_date=modified_date
        )
    
    folder_data = json.dumps({
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename"
    }).encode('utf-8')
    
    primary_endpoint = resolve_site_drive_endpoint(site_id, access_token)

    try:
        created = _create_folder_at_endpoint(primary_endpoint)
        if created:
            _queue_folder_metadata_update()
            return True
        return False
    except urllib.error.HTTPError as e:
        if e.code == 409:  # Conflict - folder already exists, skip timestamp
            return True
        if e.code == 404:
            retry_endpoint = resolve_site_drive_endpoint(site_id, access_token, force_refresh=True)
            if retry_endpoint != primary_endpoint:
                try:
                    created = _create_folder_at_endpoint(retry_endpoint)
                    if created:
                        _queue_folder_metadata_update()
                        return True
                    return False
                except urllib.error.HTTPError as retry_error:
                    if retry_error.code == 409:
                        return True
                    print_error(f"Failed to create folder {folder_name}: {retry_error.code}")
                    return False
                except Exception as retry_exception:
                    print_error(f"Error creating folder {folder_name}: {retry_exception}")
                    return False
        print_error(f"Failed to create folder {folder_name}: {e.code}")
        return False
    except Exception as e:
        print_error(f"Error creating folder {folder_name}: {e}")
        return False

def generate_random_past_date(months_back: int = 12) -> datetime:
    """Generate a realistic past timestamp within the past N months."""
    now = datetime.now(timezone.utc).replace(microsecond=0)
    max_days_back = max(1, months_back * 30)

    for _ in range(12):
        candidate = now - timedelta(
            days=random.randint(0, max_days_back),
            hours=random.randint(0, 23),
            minutes=random.randint(0, 59),
            seconds=random.randint(0, 59)
        )
        candidate = candidate.replace(
            hour=random.randint(7, 18),
            minute=random.randint(0, 59),
            second=random.randint(0, 59),
            microsecond=0
        )
        if candidate.weekday() < 5 and candidate <= now:
            return candidate

    fallback = now - timedelta(days=random.randint(1, max_days_back))
    return fallback.replace(hour=9, minute=0, second=0, microsecond=0)


def generate_random_modified_date(created_date: datetime) -> datetime:
    """Generate a realistic modified timestamp between created date and now."""
    now = datetime.now(timezone.utc).replace(microsecond=0)
    if created_date >= now:
        return now

    total_seconds = int((now - created_date).total_seconds())
    if total_seconds <= 0:
        return now

    modified_date = created_date + timedelta(seconds=random.randint(0, total_seconds))
    if modified_date > now:
        return now
    return modified_date.replace(microsecond=0)


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


def get_drive_root_relative_path(site_id: str, access_token: str) -> Optional[str]:
    """Get the drive root server-relative path (includes library segment)."""
    try:
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive?$select=webUrl"
        req = urllib.request.Request(url)
        req.add_header("Authorization", f"Bearer {access_token}")

        with urllib.request.urlopen(req, timeout=30) as response:
            data = json.loads(response.read().decode())
            drive_web_url = data.get("webUrl", "")
            drive_path = urllib.parse.urlparse(drive_web_url).path.rstrip("/")
            if drive_path:
                return urllib.parse.unquote(drive_path)
            return None
    except Exception:
        return None


def decode_jwt_claim(token: str, claim_name: str) -> str:
    """Decode a JWT claim without signature verification for diagnostics."""
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return ""
        payload = parts[1]
        padding = "=" * ((4 - len(payload) % 4) % 4)
        decoded = base64.urlsafe_b64decode(payload + padding)
        claims = json.loads(decoded.decode("utf-8"))
        value = claims.get(claim_name, "")
        return str(value) if value is not None else ""
    except Exception:
        return ""


def get_sharepoint_access_token(site_url: str) -> Optional[str]:
    """Get a SharePoint-specific access token for REST API calls."""
    match = urllib.parse.urlparse(site_url)
    sharepoint_host = match.netloc
    if not sharepoint_host:
        return None

    resource = f"https://{sharepoint_host}"

    # Prefer Azure CLI delegated token first. In this workflow, it is more
    # likely to have effective site permissions than app-only credentials.
    result = run_command([
        "az", "account", "get-access-token",
        "--resource", resource,
        "--query", "accessToken",
        "-o", "tsv"
    ])
    if result:
        token = result.stdout.strip()
        if token:
            return token

    # Fallback: some Azure CLI environments work better with v2 scope syntax.
    result = run_command([
        "az", "account", "get-access-token",
        "--scope", f"{resource}/.default",
        "--query", "accessToken",
        "-o", "tsv"
    ])
    if result:
        token = result.stdout.strip()
        if token:
            return token

    app_config = load_app_config()
    if app_config and app_config.get("client_secret"):
        client_id = app_config.get("app_id")
        client_secret = app_config.get("client_secret")
        tenant_id = app_config.get("tenant_id")

        if client_id and client_secret and tenant_id:
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
            except Exception:
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
                        token = result.get("access_token")
                        if token:
                            return token
                except Exception:
                    pass

    return None


def get_sharepoint_form_digest(site_url: str, sharepoint_access_token: str) -> Optional[str]:
    """Get a SharePoint form digest value for REST POST operations."""
    contextinfo_url = f"{site_url.rstrip('/')}/_api/contextinfo"
    try:
        req = urllib.request.Request(contextinfo_url, data=b"", method="POST")
        req.add_header("Authorization", f"Bearer {sharepoint_access_token}")
        req.add_header("Accept", "application/json;odata=verbose")
        req.add_header("Content-Type", "application/json;odata=verbose")

        with urllib.request.urlopen(req, timeout=30) as response:
            body = json.loads(response.read().decode())
            return (
                body.get("d", {})
                .get("GetContextWebInformation", {})
                .get("FormDigestValue")
            )
    except Exception:
        return None


def build_server_relative_file_path(
    site_url: str,
    file_path: str,
    drive_root_relative_path: Optional[str] = None
) -> str:
    """Build the SharePoint server-relative path for a file."""
    site_path = urllib.parse.urlparse(site_url).path.rstrip("/")
    base_path = drive_root_relative_path.rstrip("/") if drive_root_relative_path else site_path
    normalized_file_path = file_path.strip("/")
    if not base_path:
        return f"/{normalized_file_path}"
    return f"{base_path}/{normalized_file_path}"


def format_sharepoint_datetime(value: datetime) -> str:
    """Format a UTC datetime for SharePoint field updates."""
    return value.strftime("%Y-%m-%dT%H:%M:%SZ")


def update_file_metadata_via_sharepoint_rest(
    site_url: str,
    file_path: str,
    sharepoint_access_token: str,
    drive_root_relative_path: Optional[str] = None,
    created_date: Optional[datetime] = None,
    modified_date: Optional[datetime] = None
) -> bool:
    """Update Created/Modified using SharePoint REST ValidateUpdateListItem."""
    global _metadata_last_error
    form_values = []

    if created_date:
        form_values.append({
            "FieldName": "Created",
            "FieldValue": format_sharepoint_datetime(created_date)
        })
    if modified_date:
        form_values.append({
            "FieldName": "Modified",
            "FieldValue": format_sharepoint_datetime(modified_date)
        })

    if not form_values:
        return True

    server_relative_path = build_server_relative_file_path(site_url, file_path, drive_root_relative_path)
    encoded_relative_path = urllib.parse.quote(server_relative_path, safe="/")
    update_url = (
        f"{site_url.rstrip('/')}/_api/web/"
        "GetFileByServerRelativePath(decodedurl=@a1)/ListItemAllFields/ValidateUpdateListItem()"
        f"?@a1='{encoded_relative_path}'"
    )
    payload = json.dumps({
        "formValues": form_values,
        "bNewDocumentUpdate": True
    }).encode("utf-8")

    form_digest = get_sharepoint_form_digest(site_url, sharepoint_access_token)

    try:
        req = urllib.request.Request(update_url, data=payload, method="POST")
        req.add_header("Authorization", f"Bearer {sharepoint_access_token}")
        req.add_header("Accept", "application/json;odata=nometadata")
        req.add_header("Content-Type", "application/json;odata=nometadata")
        if form_digest:
            req.add_header("X-RequestDigest", form_digest)

        with urllib.request.urlopen(req, timeout=30) as response:
            if response.status not in [200, 204]:
                return False

            response_body = response.read().decode().strip()
            if not response_body:
                return True

            result = json.loads(response_body)
            validation_results = result.get("value", result)
            if isinstance(validation_results, list):
                return not any(entry.get("HasException") for entry in validation_results)
            return True
    except urllib.error.HTTPError as e:
        response_snippet = ""
        try:
            response_snippet = e.read().decode("utf-8", errors="ignore").strip().replace("\n", " ")[:200]
        except Exception:
            response_snippet = ""
        with _metadata_lock:
            if not _metadata_last_error:
                _metadata_last_error = f"HTTP {e.code}" + (f": {response_snippet}" if response_snippet else "")
        return False
    except Exception:
        with _metadata_lock:
            if not _metadata_last_error:
                _metadata_last_error = "Request failed"
        return False


def update_file_metadata_via_graph(
    site_id: str,
    list_item_id: str,
    access_token: str,
    created_date: Optional[datetime] = None,
    modified_date: Optional[datetime] = None,
    author_id: Optional[str] = None,
    editor_id: Optional[str] = None
) -> bool:
    """Legacy compatibility wrapper.

    SharePoint system fields are now updated through SharePoint REST rather than
    Microsoft Graph, so this legacy Graph path is intentionally disabled.
    """
    return False


# Keep the old function name as an alias for backwards compatibility
def update_file_metadata_via_rest(
    site_url: str,
    file_path: str,
    access_token: str,
    created_date: Optional[datetime] = None,
    modified_date: Optional[datetime] = None,
    author_id: Optional[str] = None,
    editor_id: Optional[str] = None
) -> bool:
    """Compatibility wrapper for the SharePoint REST metadata updater."""
    return update_file_metadata_via_sharepoint_rest(
        site_url=site_url,
        file_path=file_path,
        sharepoint_access_token=access_token,
        created_date=created_date,
        modified_date=modified_date
    )


# Global counters for metadata updates
_metadata_update_success = 0
_metadata_update_failed = 0
_metadata_update_skipped = 0
_metadata_skip_no_site_url = 0
_metadata_skip_no_sharepoint_token = 0
_metadata_last_error = ""
_metadata_last_token_source = ""
_metadata_last_token_aud = ""
_metadata_lock = threading.Lock()
_site_context_cache: Dict[str, Dict[str, Optional[str]]] = {}
_site_context_cache_lock = threading.Lock()
_metadata_executor: Optional[concurrent.futures.ThreadPoolExecutor] = None
_metadata_futures: List[concurrent.futures.Future] = []


def reset_metadata_counters():
    """Reset the metadata update counters."""
    global _metadata_update_success, _metadata_update_failed, _metadata_update_skipped
    global _metadata_skip_no_site_url, _metadata_skip_no_sharepoint_token, _metadata_last_error
    global _metadata_last_token_source, _metadata_last_token_aud
    global _metadata_futures, _site_context_cache
    with _metadata_lock:
        _metadata_update_success = 0
        _metadata_update_failed = 0
        _metadata_update_skipped = 0
        _metadata_skip_no_site_url = 0
        _metadata_skip_no_sharepoint_token = 0
        _metadata_last_error = ""
        _metadata_last_token_source = ""
        _metadata_last_token_aud = ""
        _metadata_futures = []
    with _site_context_cache_lock:
        _site_context_cache = {}


def get_metadata_stats() -> Dict[str, Any]:
    """Get the metadata update statistics."""
    with _metadata_lock:
        return {
            "success": _metadata_update_success,
            "failed": _metadata_update_failed,
            "skipped": _metadata_update_skipped,
            "skip_no_site_url": _metadata_skip_no_site_url,
            "skip_no_sharepoint_token": _metadata_skip_no_sharepoint_token,
            "last_error": _metadata_last_error,
            "last_token_source": _metadata_last_token_source,
            "last_token_aud": _metadata_last_token_aud
        }


def get_site_metadata_context(site_id: str, graph_access_token: str) -> Dict[str, Optional[str]]:
    """Get cached SharePoint metadata context for a site."""
    with _site_context_cache_lock:
        cached_context = _site_context_cache.get(site_id)
    if cached_context is not None:
        return cached_context

    site_url = get_sharepoint_site_url(site_id, graph_access_token)
    sharepoint_token = get_sharepoint_access_token(site_url) if site_url else None
    token_aud = decode_jwt_claim(sharepoint_token, "aud") if sharepoint_token else ""
    token_idtyp = decode_jwt_claim(sharepoint_token, "idtyp") if sharepoint_token else ""
    token_source = "unknown"
    if sharepoint_token:
        token_source = "app" if token_idtyp == "app" else "delegated"

    context = {
        "site_url": site_url,
        "drive_root_relative_path": get_drive_root_relative_path(site_id, graph_access_token),
        "sharepoint_token": sharepoint_token,
        "sharepoint_token_aud": token_aud,
        "sharepoint_token_source": token_source
    }

    with _site_context_cache_lock:
        _site_context_cache[site_id] = context

    return context


def ensure_metadata_executor() -> concurrent.futures.ThreadPoolExecutor:
    """Create the metadata executor on first use."""
    global _metadata_executor
    with _metadata_lock:
        if _metadata_executor is None:
            _metadata_executor = concurrent.futures.ThreadPoolExecutor(
                max_workers=METADATA_UPDATE_WORKERS,
                thread_name_prefix="sp-metadata"
            )
        return _metadata_executor


def apply_realistic_timestamps(
    site_id: str,
    file_path: str,
    graph_access_token: str,
    created_date: datetime,
    modified_date: datetime
) -> None:
    """Apply realistic timestamps to SharePoint Created/Modified fields.

    Preferred path is SharePoint REST ValidateUpdateListItem (most reliable for
    document library system fields). Graph fileSystemInfo PATCH is retained as
    a fallback for environments where REST token acquisition is unavailable.
    """
    global _metadata_update_success, _metadata_update_failed, _metadata_update_skipped
    global _metadata_skip_no_site_url, _metadata_skip_no_sharepoint_token
    global _metadata_last_error, _metadata_last_token_source, _metadata_last_token_aud

    context = get_site_metadata_context(site_id, graph_access_token)
    site_url = context.get("site_url")
    drive_root_relative_path = context.get("drive_root_relative_path")
    sharepoint_token = context.get("sharepoint_token")
    token_source = context.get("sharepoint_token_source", "")
    token_aud = context.get("sharepoint_token_aud", "")

    if token_source or token_aud:
        with _metadata_lock:
            _metadata_last_token_source = token_source
            _metadata_last_token_aud = token_aud

    if site_url and sharepoint_token:
        rest_success = update_file_metadata_via_sharepoint_rest(
            site_url=site_url,
            file_path=file_path,
            sharepoint_access_token=sharepoint_token,
            drive_root_relative_path=drive_root_relative_path,
            created_date=created_date,
            modified_date=modified_date
        )
        if rest_success:
            with _metadata_lock:
                _metadata_update_success += 1
            return
    else:
        with _metadata_lock:
            _metadata_update_skipped += 1
            if not site_url:
                _metadata_skip_no_site_url += 1
            if not sharepoint_token:
                _metadata_skip_no_sharepoint_token += 1

    # Fallback to Graph fileSystemInfo patch.
    encoded_path = urllib.parse.quote(file_path)
    drive_endpoint = resolve_site_drive_endpoint(site_id, graph_access_token)
    url = f"{drive_endpoint}/root:/{encoded_path}:"

    payload = json.dumps({
        "fileSystemInfo": {
            "createdDateTime": created_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "lastModifiedDateTime": modified_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
    }).encode("utf-8")

    try:
        req = urllib.request.Request(url, data=payload, method="PATCH")
        req.add_header("Authorization", f"Bearer {graph_access_token}")
        req.add_header("Content-Type", "application/json")

        with urllib.request.urlopen(req, timeout=30) as response:
            if response.status in [200, 201]:
                with _metadata_lock:
                    _metadata_update_success += 1
            else:
                with _metadata_lock:
                    _metadata_update_failed += 1
    except urllib.error.HTTPError as e:
        if e.code == 404:
            retry_endpoint = resolve_site_drive_endpoint(site_id, graph_access_token, force_refresh=True)
            if retry_endpoint != drive_endpoint:
                retry_url = f"{retry_endpoint}/root:/{encoded_path}:"
                try:
                    retry_req = urllib.request.Request(retry_url, data=payload, method="PATCH")
                    retry_req.add_header("Authorization", f"Bearer {graph_access_token}")
                    retry_req.add_header("Content-Type", "application/json")
                    with urllib.request.urlopen(retry_req, timeout=30) as response:
                        if response.status in [200, 201]:
                            with _metadata_lock:
                                _metadata_update_success += 1
                            return
                except Exception:
                    pass

        response_snippet = ""
        try:
            response_snippet = e.read().decode("utf-8", errors="ignore").strip().replace("\n", " ")[:200]
        except Exception:
            response_snippet = ""
        with _metadata_lock:
            _metadata_update_failed += 1
            if not _metadata_last_error:
                _metadata_last_error = f"HTTP {e.code}" + (f": {response_snippet}" if response_snippet else "")
    except Exception as ex:
        with _metadata_lock:
            _metadata_update_failed += 1
            if not _metadata_last_error:
                _metadata_last_error = str(ex)[:200]


def queue_metadata_update(
    site_id: str,
    file_path: str,
    graph_access_token: str,
    created_date: datetime,
    modified_date: datetime
) -> None:
    """Queue a metadata update so uploads do not block on REST calls."""
    executor = ensure_metadata_executor()
    future = executor.submit(
        apply_realistic_timestamps,
        site_id,
        file_path,
        graph_access_token,
        created_date,
        modified_date
    )
    with _metadata_lock:
        _metadata_futures.append(future)


def wait_for_metadata_updates() -> None:
    """Wait for queued metadata updates to finish before summarizing."""
    global _metadata_executor, _metadata_futures

    with _metadata_lock:
        pending_futures = list(_metadata_futures)
        _metadata_futures = []

    for future in pending_futures:
        try:
            future.result()
        except Exception:
            with _metadata_lock:
                _metadata_update_failed += 1

    executor = _metadata_executor
    _metadata_executor = None
    if executor is not None:
        executor.shutdown(wait=True)


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

    Realistic timestamp updates are queued separately so uploads stay fast.
    """
    global _metadata_update_success, _metadata_update_failed, _metadata_update_skipped
    global _metadata_skip_no_site_url, _metadata_skip_no_sharepoint_token
    # Encode the file path
    if folder_path:
        file_path = f"{folder_path}/{file_name}"
        encoded_path = urllib.parse.quote(file_path)
    else:
        file_path = file_name
        encoded_path = urllib.parse.quote(file_name)
    
    primary_drive_endpoint = resolve_site_drive_endpoint(site_id, access_token)
    url = f"{primary_drive_endpoint}/root:/{encoded_path}:/content"

    def _upload(upload_url: str) -> bool:
        req = urllib.request.Request(upload_url, data=file_content, method="PUT")
        req.add_header("Authorization", f"Bearer {access_token}")
        req.add_header("Content-Type", "application/octet-stream")

        with urllib.request.urlopen(req, timeout=60) as response:
            upload_success = response.status in [200, 201]

            if upload_success and set_custom_dates:
                created_date = generate_random_past_date(months_back=12)
                modified_date = generate_random_modified_date(created_date)
                queue_metadata_update(
                    site_id=site_id,
                    file_path=file_path,
                    graph_access_token=access_token,
                    created_date=created_date,
                    modified_date=modified_date
                )
            elif upload_success:
                with _metadata_lock:
                    _metadata_update_skipped += 1

            return upload_success
    
    try:
        return _upload(url)
    except urllib.error.HTTPError as e:
        if e.code == 409:  # Conflict - file already exists
            return True
        if e.code == 404:
            retry_endpoint = resolve_site_drive_endpoint(site_id, access_token, force_refresh=True)
            if retry_endpoint != primary_drive_endpoint:
                retry_url = f"{retry_endpoint}/root:/{encoded_path}:/content"
                try:
                    return _upload(retry_url)
                except urllib.error.HTTPError as retry_error:
                    if retry_error.code == 409:
                        return True
                    return False
                except Exception:
                    return False
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
    
    templates = build_template_pool(site_type)
    folders = build_folder_pool(site_type, templates["folders"])
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
    templates = build_template_pool(site_type)
    dept_folders = build_folder_pool(site_type, templates["folders"])
    
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
    global VARIATION_LEVEL

    parser = argparse.ArgumentParser(
        description="Populate SharePoint sites with realistic files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python populate_files.py                      # Interactive mode
    python populate_files.py --files 100          # Create 100 files across all sites
    python populate_files.py --files 50 --site hr # Create 50 files in sites containing 'hr'
    python populate_files.py --files 300 --variation-level high
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
    parser.add_argument(
        '--variation-level',
        type=str,
        choices=['low', 'medium', 'high'],
        default='medium',
        help='Controls template/folder variation density (default: medium)'
    )
    
    args = parser.parse_args()
    VARIATION_LEVEL = args.variation_level
    
    # Clear screen and show banner
    os.system('cls' if os.name == 'nt' else 'clear')
    print_banner("SHAREPOINT FILE POPULATION")
    
    print(f"  {Colors.WHITE}This script populates SharePoint sites with realistic files{Colors.NC}")
    print(f"  {Colors.WHITE}to simulate an actual organization's document structure.{Colors.NC}")
    print(f"  {Colors.BLUE}Variation level:{Colors.NC} {VARIATION_LEVEL}")
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
    
    # Filter sites if specified via command line
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
    
    # Step 5: Select Site Scope (only in interactive mode without --site filter)
    if not args.site:
        print_step(5, "Select Site Scope")
        
        print()
        print(f"  {Colors.WHITE}To ensure files are only populated in sites you created,{Colors.NC}")
        print(f"  {Colors.WHITE}select how to filter the discovered sites:{Colors.NC}")
        
        sites, scope_desc = select_site_scope(sites, access_token)
        
        if not sites:
            if scope_desc == "cancelled":
                print_warning("Operation cancelled.")
            elif scope_desc == "no_matches":
                print_warning("No matching sites found.")
            elif scope_desc == "invalid":
                print_error("Invalid selection.")
            sys.exit(0)
        
        print()
        print_success(f"Selected {len(sites)} sites ({scope_desc})")
    
    # Step 6: Discover Azure AD Users & Groups
    print_step(6, "Discover Azure AD Users & Groups")
    
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
    
    # Step 7: Select Population Mode
    print_step(7, "Select File Population Mode")
    
    population_mode = select_population_mode(users, groups)
    if population_mode == "quit":
        print_warning("Operation cancelled.")
        sys.exit(0)
    
    print_success(f"Selected mode: {population_mode.replace('_', ' ').title()}")
    
    # Step 8: Get file count
    print_step(8, "Configure File Generation")
    
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
    
    # Step 9: Confirm
    print_step(9, "Confirm File Generation")
    
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
    
    # Step 10: Generate files
    print_step(10, "Generate and Upload Files")
    
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
    
    wait_for_metadata_updates()

    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds()
    
    # Step 11: Summary
    print_step(11, "Summary")
    
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
    
    # Display metadata update statistics
    metadata_stats = get_metadata_stats()
    total_metadata_attempts = metadata_stats["success"] + metadata_stats["failed"] + metadata_stats["skipped"]
    if total_metadata_attempts > 0:
        print(f"  {Colors.WHITE}Custom Timestamps (SharePoint REST + Graph fallback):{Colors.NC}")
        if metadata_stats["success"] > 0:
            print(f"    {Colors.GREEN}✓ Timestamps set:{Colors.NC} {metadata_stats['success']} files")
        if metadata_stats["failed"] > 0:
            print(f"    {Colors.RED}✗ Timestamp update failed:{Colors.NC} {metadata_stats['failed']} files")
            if metadata_stats.get("last_error"):
                print(f"      - Sample error: {metadata_stats['last_error']}")
        if metadata_stats["skipped"] > 0:
            print(f"    {Colors.YELLOW}⊘ Skipped:{Colors.NC} {metadata_stats['skipped']} files")
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