#!/usr/bin/env python3
"""
M365 Email Population Script

Populates Microsoft 365 mailboxes with realistic organisational emails
to simulate an actual enterprise environment.

Usage:
    python populate_emails.py                           # Interactive mode
    python populate_emails.py --all                     # All mailboxes
    python populate_emails.py --count 10                # 10 random mailboxes
    python populate_emails.py --mailboxes "user@domain" # Specific mailboxes
    python populate_emails.py --emails-per-mailbox 100  # 100 emails each
    python populate_emails.py --dry-run                 # Preview only
    python populate_emails.py --list-mailboxes          # List configured mailboxes

Requirements:
    - Python 3.8+
    - Azure CLI (logged in)
    - PyYAML (pip install pyyaml)
    - Microsoft Graph API permissions (Mail.ReadWrite)
"""

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any

# Add parent directory to path for imports
SCRIPT_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(SCRIPT_DIR))

# Local imports
from email_generator.config import (
    load_mailbox_config,
    load_sites_config,
    load_environment_config,
    get_environment,
    get_tenant_id,
    get_tenant_domain,
    get_all_users,
    check_yaml_installed,
    is_azure_ad_enabled,
    get_cc_bcc_config,
    is_exclusions_enabled,
    is_email_excluded,
    should_log_exclusions,
    filter_excluded_users,
)
from email_generator.content_generator import EmailContentGenerator
from email_generator.graph_client import GraphClient
from email_generator.threading import ThreadManager
from email_generator.attachments import AttachmentGenerator
from email_generator.user_pool import UserPool
from email_generator.utils import (
    Colors,
    clear_screen,
    print_banner,
    print_step,
    print_success,
    print_error,
    print_warning,
    print_info,
    print_progress,
    print_summary_box,
    check_azure_login,
    azure_login,
    switch_to_tenant,
    format_duration,
    confirm_action,
    get_numeric_input,
)

# Configuration paths
PROJECT_DIR = SCRIPT_DIR.parent
CONFIG_DIR = PROJECT_DIR / "config"
MAILBOXES_FILE = CONFIG_DIR / "mailboxes.yaml"
SITES_FILE = CONFIG_DIR / "sites.json"
ENVIRONMENTS_FILE = CONFIG_DIR / "environments.json"

# Limits
MAX_EMAILS_PER_MAILBOX = 500
MAX_TOTAL_EMAILS = 10000


class EmailPopulator:
    """Main class for populating M365 mailboxes with emails."""
    
    def __init__(self):
        """Initialize the email populator."""
        self.config: Dict[str, Any] = {}
        self.sites: Dict[str, Any] = {}
        self.environment: Optional[Dict[str, Any]] = None
        self.graph_client: Optional[GraphClient] = None
        self.content_generator: Optional[EmailContentGenerator] = None
        self.thread_manager: Optional[ThreadManager] = None
        self.attachment_generator: Optional[AttachmentGenerator] = None
        self.user_pool: Optional[UserPool] = None
        
        # Statistics
        self.stats = {
            "total_emails": 0,
            "successful": 0,
            "failed": 0,
            "mailboxes_processed": 0,
            "start_time": None,
            "end_time": None,
            "by_category": {},
        }
    
    def load_configuration(self) -> bool:
        """Load all configuration files."""
        try:
            # Check PyYAML is installed
            if not check_yaml_installed():
                print_error("PyYAML is required. Install with: pip install pyyaml")
                return False
            
            # Load mailbox configuration
            if not MAILBOXES_FILE.exists():
                print_error(f"Mailbox configuration not found: {MAILBOXES_FILE}")
                print_info("Please create config/mailboxes.yaml with your mailbox definitions.")
                return False
            
            self.config = load_mailbox_config(MAILBOXES_FILE)
            print_success(f"Loaded {len(self.config.get('users', []))} mailboxes from configuration")
            
            # Load sites configuration
            if SITES_FILE.exists():
                self.sites = load_sites_config(SITES_FILE)
                print_success(f"Loaded {len(self.sites.get('sites', []))} SharePoint sites")
            else:
                print_warning("No SharePoint sites configuration found")
                self.sites = {"sites": []}
            
            # Load environment configuration
            if ENVIRONMENTS_FILE.exists():
                env_config = load_environment_config(ENVIRONMENTS_FILE)
                self.environment = get_environment(env_config)
                if self.environment:
                    print_success(f"Using environment: {self.environment.get('name', 'Default')}")
            
            return True
            
        except Exception as e:
            print_error(f"Failed to load configuration: {e}")
            return False
    
    def initialize_components(self) -> bool:
        """Initialize all components."""
        try:
            # Get rate limiting configuration
            settings = self.config.get("settings", {})
            rate_limit_config = settings.get("rate_limiting", {})
            
            # Initialize Graph client with rate limiting
            self.graph_client = GraphClient(rate_limit_config=rate_limit_config)
            
            # Log rate limiting settings if configured
            if rate_limit_config:
                print_info(f"Rate limiting: {rate_limit_config.get('request_delay_ms', 100)}ms between requests, "
                          f"{rate_limit_config.get('batch_delay_ms', 500)}ms between batches, "
                          f"{rate_limit_config.get('max_retries', 5)} max retries")
            
            # Initialize user pool for CC/BCC support
            # This uses the cc_bcc configuration from mailboxes.yaml
            cc_bcc_config = get_cc_bcc_config(self.config)
            if cc_bcc_config.get("cc", {}).get("enabled", True) or cc_bcc_config.get("bcc", {}).get("enabled", True):
                self.user_pool = UserPool(config=self.config)
                print_success("User pool initialized for CC/BCC support")
            
            # Initialize content generator with user pool
            self.content_generator = EmailContentGenerator(self.config, self.sites, self.user_pool)
            
            # Initialize thread manager
            self.thread_manager = ThreadManager(self.config)
            
            # Initialize attachment generator
            self.attachment_generator = AttachmentGenerator()
            
            return True
            
        except Exception as e:
            print_error(f"Failed to initialize components: {e}")
            return False
    
    def authenticate(self) -> bool:
        """Authenticate with Azure and Graph API."""
        # Check Azure CLI login
        if not check_azure_login():
            print_warning("Not logged into Azure CLI")
            if not azure_login():
                print_error("Azure login failed")
                return False
        
        print_success("Azure CLI authenticated")
        
        # Switch tenant if configured
        if self.environment:
            tenant_id = get_tenant_id(self.environment)
            if tenant_id and not tenant_id.startswith("<<"):
                if not switch_to_tenant(tenant_id):
                    print_warning("Could not switch to configured tenant")
        
        # Authenticate Graph client
        if not self.graph_client.authenticate():
            print_error("Failed to authenticate with Microsoft Graph API")
            print_info("Make sure you have the required permissions:")
            print_info("  - Mail.ReadWrite")
            print_info("  - User.Read.All")
            return False
        
        print_success("Microsoft Graph API authenticated")
        
        # Verify permissions
        perm_check = self.graph_client.verify_permissions()
        if not perm_check.get("has_permissions"):
            print_warning(f"Permission check: {perm_check.get('error', 'Unknown error')}")
        
        return True
    
    def get_mailboxes(
        self,
        all_mailboxes: bool = False,
        count: Optional[int] = None,
        specific: Optional[List[str]] = None,
        department: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Get list of mailboxes to populate based on selection criteria.
        
        Automatically filters out excluded email addresses and domains
        based on the exclusions configuration in mailboxes.yaml.
        """
        import random
        
        users = get_all_users(self.config)
        
        if not users:
            return []
        
        # Filter out excluded users based on exclusions configuration
        if is_exclusions_enabled(self.config):
            included_users, excluded_users = filter_excluded_users(
                users,
                self.config,
                log_func=print_warning if should_log_exclusions(self.config) else None
            )
            
            if excluded_users:
                print_info(f"Filtered out {len(excluded_users)} excluded mailbox(es)")
            
            users = included_users
        
        # Filter by specific mailboxes
        if specific:
            specific_lower = [s.lower() for s in specific]
            users = [u for u in users if u.get("upn", "").lower() in specific_lower]
        
        # Filter by department
        if department:
            dept_lower = department.lower()
            users = [u for u in users if dept_lower in u.get("department", "").lower()]
        
        # Select count
        if count and count < len(users):
            users = random.sample(users, count)
        
        return users
    
    def _select_folder_weighted(self, folders: List[str], category: str) -> str:
        """
        Select a folder using weighted distribution for realism.
        
        ONLY spam emails go to junkemail folder.
        Legitimate emails use realistic distribution:
        - inbox: 55% (most emails land here)
        - sentitems: 25% (sent emails)
        - drafts: 10% (unfinished emails)
        - deleteditems: 10% (deleted emails)
        
        Args:
            folders: List of available folders.
            category: Email category (spam, newsletters, etc.).
            
        Returns:
            Selected folder name.
        """
        import random
        
        # SPAM emails: ALWAYS go to junkemail folder
        # They should NEVER go to inbox, sentitems, drafts, or deleteditems
        if category == "spam":
            if "junkemail" in folders:
                return "junkemail"
            # If junkemail not available, still try to use it
            return "junkemail"
        
        # For legitimate emails, exclude junkemail from available folders
        # Legitimate emails should NEVER go to junk folder
        legitimate_folders = [f for f in folders if f != "junkemail"]
        
        # If no legitimate folders available, fall back to inbox
        if not legitimate_folders:
            return "inbox" if "inbox" in folders else folders[0]
        
        # Define realistic folder weights (no junkemail for legitimate emails)
        folder_weights = {
            "inbox": 55,
            "sentitems": 25,
            "drafts": 10,
            "deleteditems": 10,
        }
        
        # Filter to only available legitimate folders and get their weights
        available_weights = []
        available_folders = []
        for folder in legitimate_folders:
            if folder in folder_weights:
                available_folders.append(folder)
                available_weights.append(folder_weights[folder])
            else:
                # Unknown folder gets default weight (but not junkemail)
                if folder != "junkemail":
                    available_folders.append(folder)
                    available_weights.append(10)
        
        if not available_folders:
            return "inbox" if "inbox" in folders else (folders[0] if folders else "inbox")
        
        return random.choices(available_folders, weights=available_weights)[0]
    
    def populate_mailbox(
        self,
        user: Dict[str, Any],
        email_count: int,
        dry_run: bool = False,
        folders: Optional[List[str]] = None
    ) -> Dict[str, int]:
        """
        Populate a single mailbox with emails.
        
        Args:
            user: User configuration dictionary.
            email_count: Number of emails to create.
            dry_run: If True, don't actually create emails.
            folders: List of folders to distribute emails to. Default is ["inbox"].
            
        Returns:
            Dictionary with success and failure counts.
        """
        import random
        
        results = {"success": 0, "failed": 0}
        upn = user.get("upn", "")
        
        # Default to inbox only
        if not folders:
            folders = ["inbox"]
        
        # Track folder distribution
        folder_counts: Dict[str, int] = {f: 0 for f in folders}
        
        for i in range(email_count):
            try:
                # Generate email
                email = self.content_generator.generate_email(user)
                
                # Handle threading
                if self.thread_manager.should_create_thread():
                    email = self.thread_manager.create_thread(email, user)
                
                # Handle attachments
                if email.get("has_attachment"):
                    attachment = self.attachment_generator.generate(
                        email.get("attachment_type"),
                        email.get("attachment_department", user.get("department", "General"))
                    )
                    email["attachments"] = [attachment]
                
                # Track category
                category = email.get("category", "other")
                self.stats["by_category"][category] = self.stats["by_category"].get(category, 0) + 1
                
                # Select folder for this email (weighted distribution)
                folder = self._select_folder_weighted(folders, category)
                folder_counts[folder] += 1
                
                # Create email (or simulate in dry run)
                if dry_run:
                    results["success"] += 1
                else:
                    if self.graph_client.create_email(upn, email, folder=folder):
                        results["success"] += 1
                    else:
                        results["failed"] += 1
                
                # Update progress with folder info
                folder_short = folder[:5] if len(folder) > 5 else folder
                print_progress(
                    i + 1,
                    email_count,
                    f"[{folder_short}] {email.get('subject', 'Email')[:30]}..."
                )
                
            except Exception as e:
                results["failed"] += 1
        
        print()  # New line after progress bar
        
        # Store folder distribution in results
        results["folders"] = folder_counts  # type: ignore
        
        return results
    
    def run(
        self,
        all_mailboxes: bool = False,
        count: Optional[int] = None,
        specific: Optional[List[str]] = None,
        department: Optional[str] = None,
        emails_per_mailbox: Optional[int] = None,
        emails_min: Optional[int] = None,
        emails_max: Optional[int] = None,
        dry_run: bool = False,
        folders: Optional[List[str]] = None
    ) -> bool:
        """
        Main execution method.
        
        Args:
            all_mailboxes: Populate all configured mailboxes.
            count: Number of mailboxes to populate.
            specific: List of specific mailbox UPNs.
            department: Filter by department.
            emails_per_mailbox: Fixed number of emails per mailbox.
            emails_min: Minimum emails per mailbox (for random).
            emails_max: Maximum emails per mailbox (for random).
            dry_run: Preview without creating emails.
            folders: List of folders to distribute emails to.
            
        Returns:
            True if successful, False otherwise.
        """
        import random
        
        self.stats["start_time"] = datetime.now()
        
        # Get mailboxes to populate
        mailboxes = self.get_mailboxes(all_mailboxes, count, specific, department)
        
        if not mailboxes:
            print_error("No mailboxes selected for population")
            return False
        
        print_success(f"Selected {len(mailboxes)} mailboxes for population")
        
        # Validate mailboxes exist in Azure AD (skip in dry run)
        if not dry_run:
            print_info("Validating mailboxes...")
            valid_mailboxes = []
            invalid_mailboxes = []
            
            for user in mailboxes:
                upn = user.get("upn", "")
                verification = self.graph_client.verify_mailbox(upn)
                
                if verification.get("exists") and verification.get("has_mailbox"):
                    valid_mailboxes.append(user)
                    print(f"    {Colors.GREEN}✓{Colors.NC} {upn}")
                else:
                    invalid_mailboxes.append({
                        "upn": upn,
                        "error": verification.get("error", "Unknown error")
                    })
                    print(f"    {Colors.RED}✗{Colors.NC} {upn} - {verification.get('error', 'Not found')}")
            
            print()
            
            if invalid_mailboxes:
                print_warning(f"Found {len(invalid_mailboxes)} invalid mailbox(es):")
                for mb in invalid_mailboxes:
                    print(f"    • {mb['upn']}: {mb['error']}")
                print()
                
                if not valid_mailboxes:
                    print_error("No valid mailboxes found. Please check your mailboxes.yaml configuration.")
                    return False
                
                # Ask user if they want to continue with valid mailboxes only
                if not confirm_action(f"Continue with {len(valid_mailboxes)} valid mailbox(es)?"):
                    print_warning("Operation cancelled")
                    return False
                
                mailboxes = valid_mailboxes
                print_success(f"Proceeding with {len(mailboxes)} valid mailboxes")
            else:
                print_success(f"All {len(mailboxes)} mailboxes validated successfully")
        
        # Determine email counts
        settings = self.config.get("settings", {})
        default_count = settings.get("default_email_count", 50)
        
        # Calculate emails per mailbox
        mailbox_email_counts = []
        for user in mailboxes:
            if emails_per_mailbox:
                count = emails_per_mailbox
            elif emails_min and emails_max:
                count = random.randint(emails_min, emails_max)
            else:
                # Use user's configured volume or default
                count = user.get("email_count", default_count)
            
            # Apply limits
            count = min(count, MAX_EMAILS_PER_MAILBOX)
            mailbox_email_counts.append(count)
        
        total_emails = sum(mailbox_email_counts)
        
        # Check total limit
        if total_emails > MAX_TOTAL_EMAILS:
            print_warning(f"Total emails ({total_emails}) exceeds limit ({MAX_TOTAL_EMAILS})")
            print_info("Reducing email counts proportionally...")
            factor = MAX_TOTAL_EMAILS / total_emails
            mailbox_email_counts = [int(c * factor) for c in mailbox_email_counts]
            total_emails = sum(mailbox_email_counts)
        
        # Default folders if not specified
        if not folders:
            folders = ["inbox"]
        
        # Show summary
        folder_display = ", ".join(folders) if len(folders) <= 3 else f"{len(folders)} folders"
        print_summary_box("Email Population Summary", [
            ("Mailboxes:", len(mailboxes)),
            ("Total Emails:", total_emails),
            ("Avg per Mailbox:", total_emails // len(mailboxes) if mailboxes else 0),
            ("Folders:", folder_display),
            ("Mode:", "DRY RUN" if dry_run else "LIVE"),
        ])
        
        # List mailboxes
        print(f"  {Colors.WHITE}Mailboxes to populate:{Colors.NC}")
        for i, (user, email_count) in enumerate(zip(mailboxes[:10], mailbox_email_counts[:10])):
            print(f"    • {user.get('upn')} ({email_count} emails)")
        if len(mailboxes) > 10:
            print(f"    ... and {len(mailboxes) - 10} more")
        print()
        
        # Confirm
        if not dry_run:
            if not confirm_action("Proceed with email population?"):
                print_warning("Operation cancelled")
                return False
        
        # Process mailboxes
        print()
        for i, (user, email_count) in enumerate(zip(mailboxes, mailbox_email_counts)):
            upn = user.get("upn", "Unknown")
            department = user.get("department", "General")
            
            print(f"  {Colors.CYAN}[{i+1}/{len(mailboxes)}]{Colors.NC} {upn}")
            print(f"       Department: {department} | Emails: {email_count}")
            
            results = self.populate_mailbox(user, email_count, dry_run, folders)
            
            self.stats["successful"] += results["success"]
            self.stats["failed"] += results["failed"]
            self.stats["total_emails"] += results["success"] + results["failed"]
            self.stats["mailboxes_processed"] += 1
            
            # Track folder distribution
            if "folders" in results:
                if "by_folder" not in self.stats:
                    self.stats["by_folder"] = {}
                for folder, folder_count in results["folders"].items():
                    self.stats["by_folder"][folder] = self.stats["by_folder"].get(folder, 0) + folder_count
            
            if results["failed"] > 0:
                print_warning(f"  {results['success']} succeeded, {results['failed']} failed")
            else:
                print_success(f"  {results['success']} emails created")
            print()
        
        self.stats["end_time"] = datetime.now()
        
        return True
    
    def print_final_summary(self) -> None:
        """Print final summary of the operation."""
        # Handle case where start_time or end_time is None (e.g., early exit)
        if self.stats.get("start_time") and self.stats.get("end_time"):
            duration = (self.stats["end_time"] - self.stats["start_time"]).total_seconds()
        else:
            duration = 0
        
        print()
        print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
        print(f"  {Colors.BOLD}{Colors.WHITE}{'EMAIL POPULATION COMPLETE':^60}{Colors.NC}")
        print(f"  {Colors.CYAN}{'=' * 60}{Colors.NC}")
        print()
        
        print(f"    {Colors.GREEN}✓ Successfully created:{Colors.NC} {self.stats['successful']} emails")
        if self.stats["failed"] > 0:
            print(f"    {Colors.RED}✗ Failed:{Colors.NC} {self.stats['failed']} emails")
        print(f"    {Colors.BLUE}📬 Mailboxes processed:{Colors.NC} {self.stats['mailboxes_processed']}")
        print(f"    {Colors.BLUE}⏱ Duration:{Colors.NC} {format_duration(duration)}")
        
        if duration > 0:
            rate = self.stats["successful"] / duration
            print(f"    {Colors.BLUE}📊 Rate:{Colors.NC} {rate:.1f} emails/second")
        
        # Folder breakdown
        if self.stats.get("by_folder"):
            print()
            print(f"  {Colors.WHITE}Folder Distribution:{Colors.NC}")
            total = sum(self.stats["by_folder"].values())
            folder_icons = {
                "inbox": "📥",
                "sentitems": "📤",
                "deleteditems": "🗑️",
                "drafts": "📝",
            }
            for folder, folder_count in sorted(self.stats["by_folder"].items()):
                pct = (folder_count / total * 100) if total > 0 else 0
                icon = folder_icons.get(folder, "📁")
                print(f"    {icon} {folder}: {folder_count} ({pct:.0f}%)")
        
        # Category breakdown
        if self.stats["by_category"]:
            print()
            print(f"  {Colors.WHITE}Email Type Distribution:{Colors.NC}")
            total = sum(self.stats["by_category"].values())
            for category, count in sorted(self.stats["by_category"].items()):
                pct = (count / total * 100) if total > 0 else 0
                icon = {
                    "newsletters": "📰",
                    "links": "🔗",
                    "attachments": "📎",
                    "organisational": "📢",
                    "interdepartmental": "💬",
                }.get(category, "📧")
                print(f"    {icon} {category.capitalize()}: {count} ({pct:.0f}%)")
        
        print()
        
        if self.stats["successful"] > 0:
            print_success("Emails have been created in your M365 mailboxes!")
        else:
            print_error("No emails were created successfully")


def list_mailboxes(config: Dict[str, Any]) -> None:
    """List all configured mailboxes."""
    users = get_all_users(config)
    
    if not users:
        print_warning("No mailboxes configured")
        return
    
    print()
    print(f"  {Colors.WHITE}Configured Mailboxes ({len(users)}):{Colors.NC}")
    print()
    
    # Group by department
    by_dept: Dict[str, List] = {}
    for user in users:
        dept = user.get("department", "General")
        if dept not in by_dept:
            by_dept[dept] = []
        by_dept[dept].append(user)
    
    for dept, dept_users in sorted(by_dept.items()):
        print(f"  {Colors.CYAN}{dept}{Colors.NC}")
        for user in dept_users:
            upn = user.get("upn", "Unknown")
            job_title = user.get("job_title", "")
            email_count = user.get("email_count", 50)
            print(f"    • {upn}")
            print(f"      {Colors.DIM}Job Title: {job_title} | Emails: {email_count}{Colors.NC}")
        print()


def main() -> None:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Populate M365 mailboxes with realistic emails",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python populate_emails.py                           # Interactive mode
    python populate_emails.py --all                     # All mailboxes
    python populate_emails.py --count 10                # 10 random mailboxes
    python populate_emails.py --mailboxes "user@domain" # Specific mailboxes
    python populate_emails.py --emails-per-mailbox 100  # 100 emails each
    python populate_emails.py --folders all             # All folders (inbox, sent, deleted, drafts)
    python populate_emails.py --folders inbox,sentitems # Specific folders
    python populate_emails.py --dry-run                 # Preview only
    python populate_emails.py --list-mailboxes          # List configured mailboxes
    python populate_emails.py --auto-discover           # Use Azure AD discovery
    python populate_emails.py --auto-discover --refresh-cache  # Refresh Azure AD cache
    python populate_emails.py --enable-cc --enable-bcc  # Enable CC/BCC recipients

Available folders:
    inbox        - Inbox folder
    sentitems    - Sent Items folder
    deleteditems - Deleted Items folder
    drafts       - Drafts folder
        """
    )
    
    # Selection options
    selection = parser.add_argument_group("Mailbox Selection")
    selection.add_argument(
        '-a', '--all',
        action='store_true',
        help='Populate all configured mailboxes'
    )
    selection.add_argument(
        '-c', '--count',
        type=int,
        metavar='N',
        help='Number of mailboxes to populate (randomly selected)'
    )
    selection.add_argument(
        '-m', '--mailboxes',
        type=str,
        metavar='UPNS',
        help='Comma-separated list of specific mailbox UPNs'
    )
    selection.add_argument(
        '-d', '--department',
        type=str,
        metavar='DEPT',
        help='Filter mailboxes by department'
    )
    
    # Email count options
    emails = parser.add_argument_group("Email Count")
    emails.add_argument(
        '-e', '--emails-per-mailbox',
        type=int,
        metavar='N',
        help=f'Fixed number of emails per mailbox (max {MAX_EMAILS_PER_MAILBOX})'
    )
    emails.add_argument(
        '--emails-min',
        type=int,
        metavar='N',
        help='Minimum emails per mailbox (for random distribution)'
    )
    emails.add_argument(
        '--emails-max',
        type=int,
        metavar='N',
        help='Maximum emails per mailbox (for random distribution)'
    )
    
    # Folder options
    folder_opts = parser.add_argument_group("Folder Options")
    folder_opts.add_argument(
        '-f', '--folders',
        type=str,
        metavar='FOLDERS',
        help='Comma-separated list of folders (inbox,sentitems,deleteditems,drafts) or "all"'
    )
    
    # Azure AD Discovery options
    discovery = parser.add_argument_group("Azure AD Discovery")
    discovery.add_argument(
        '--auto-discover',
        action='store_true',
        help='Use Azure AD auto-discovery for users and groups'
    )
    discovery.add_argument(
        '--refresh-cache',
        action='store_true',
        help='Force refresh of Azure AD cache before population'
    )
    discovery.add_argument(
        '--use-cache',
        action='store_true',
        help='Use existing Azure AD cache without refresh'
    )
    
    # CC/BCC options
    cc_bcc = parser.add_argument_group("CC/BCC Options")
    cc_bcc.add_argument(
        '--enable-cc',
        action='store_true',
        default=None,
        help='Enable CC recipients (uses config default if not specified)'
    )
    cc_bcc.add_argument(
        '--disable-cc',
        action='store_true',
        help='Disable CC recipients'
    )
    cc_bcc.add_argument(
        '--enable-bcc',
        action='store_true',
        default=None,
        help='Enable BCC recipients (uses config default if not specified)'
    )
    cc_bcc.add_argument(
        '--disable-bcc',
        action='store_true',
        help='Disable BCC recipients'
    )
    
    # Other options
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview without creating emails'
    )
    parser.add_argument(
        '-l', '--list-mailboxes',
        action='store_true',
        help='List configured mailboxes and exit'
    )
    
    args = parser.parse_args()
    
    # Clear screen and show banner
    clear_screen()
    print_banner("M365 EMAIL POPULATION")
    
    print(f"  {Colors.WHITE}This script populates M365 mailboxes with realistic emails{Colors.NC}")
    print(f"  {Colors.WHITE}to simulate an actual organisation's communication patterns.{Colors.NC}")
    print()
    
    # Create populator
    populator = EmailPopulator()
    
    # Step 1: Load configuration
    print_step(1, "Load Configuration")
    if not populator.load_configuration():
        sys.exit(1)
    
    # Handle list-mailboxes option
    if args.list_mailboxes:
        list_mailboxes(populator.config)
        sys.exit(0)
    
    # Step 2: Initialize components
    print_step(2, "Initialize Components")
    if not populator.initialize_components():
        sys.exit(1)
    print_success("All components initialized")
    
    # Step 3: Authenticate
    print_step(3, "Authenticate")
    if not populator.authenticate():
        sys.exit(1)
    
    # Step 4: Configure population
    print_step(4, "Configure Email Population")
    
    # Determine selection mode
    all_mailboxes = args.all
    count = args.count
    specific = args.mailboxes.split(",") if args.mailboxes else None
    department = args.department
    
    # If no selection specified, go interactive
    if not any([all_mailboxes, count, specific, department]):
        print()
        print(f"  {Colors.WHITE}How would you like to select mailboxes?{Colors.NC}")
        print()
        print(f"    [1] All configured mailboxes")
        print(f"    [2] Specific number of mailboxes")
        print(f"    [3] Specific mailboxes by email")
        print(f"    [4] By department")
        print()
        
        choice = input(f"  {Colors.YELLOW}Enter choice (1-4):{Colors.NC} ").strip()
        
        if choice == '1':
            all_mailboxes = True
        elif choice == '2':
            count = get_numeric_input("How many mailboxes?", 1, 100, 10)
        elif choice == '3':
            upns = input("  Enter mailbox UPNs (comma-separated): ").strip()
            specific = [u.strip() for u in upns.split(",") if u.strip()]
        elif choice == '4':
            department = input("  Enter department name: ").strip()
        else:
            all_mailboxes = True
    
    # Determine email count
    emails_per_mailbox = args.emails_per_mailbox
    emails_min = args.emails_min
    emails_max = args.emails_max
    
    if not any([emails_per_mailbox, emails_min, emails_max]):
        print()
        print(f"  {Colors.WHITE}How many emails per mailbox?{Colors.NC}")
        print()
        print(f"    [1] Use configured defaults")
        print(f"    [2] Fixed number for all")
        print(f"    [3] Random range")
        print()
        
        choice = input(f"  {Colors.YELLOW}Enter choice (1-3):{Colors.NC} ").strip()
        
        if choice == '2':
            emails_per_mailbox = get_numeric_input(
                "Emails per mailbox?", 1, MAX_EMAILS_PER_MAILBOX, 50
            )
        elif choice == '3':
            emails_min = get_numeric_input("Minimum emails?", 1, MAX_EMAILS_PER_MAILBOX, 20)
            emails_max = get_numeric_input(
                "Maximum emails?", emails_min, MAX_EMAILS_PER_MAILBOX, 100
            )
    
    # Determine folders
    ALL_FOLDERS = ["inbox", "sentitems", "deleteditems", "drafts"]
    folders: Optional[List[str]] = None
    
    if args.folders:
        if args.folders.lower() == "all":
            folders = ALL_FOLDERS
        else:
            folders = [f.strip().lower() for f in args.folders.split(",") if f.strip()]
            # Validate folder names
            valid_folders = []
            for f in folders:
                if f in ALL_FOLDERS:
                    valid_folders.append(f)
                else:
                    print_warning(f"Unknown folder '{f}' - skipping")
            folders = valid_folders if valid_folders else None
    else:
        # Interactive folder selection
        print()
        print(f"  {Colors.WHITE}Which folders should receive emails?{Colors.NC}")
        print()
        print(f"    [1] Inbox only (default)")
        print(f"    [2] All folders (inbox, sent, deleted, drafts, junk)")
        print(f"    [3] Select specific folders")
        print()
        
        choice = input(f"  {Colors.YELLOW}Enter choice (1-3):{Colors.NC} ").strip()
        
        if choice == '2':
            folders = ALL_FOLDERS + ["junkemail"]  # Add junk folder
        elif choice == '3':
            print()
            print(f"  {Colors.WHITE}Available folders:{Colors.NC}")
            print(f"    [1] inbox")
            print(f"    [2] sentitems (Sent Items)")
            print(f"    [3] deleteditems (Deleted Items)")
            print(f"    [4] drafts")
            print(f"    [5] junkemail (Junk/Spam)")
            print()
            folder_input = input(f"  {Colors.YELLOW}Enter folder numbers (comma-separated, e.g., 1,2,3):{Colors.NC} ").strip()
            
            folder_map = {"1": "inbox", "2": "sentitems", "3": "deleteditems", "4": "drafts", "5": "junkemail"}
            selected = []
            for num in folder_input.split(","):
                num = num.strip()
                if num in folder_map:
                    selected.append(folder_map[num])
            folders = selected if selected else ["inbox"]
        else:
            folders = ["inbox"]
    
    # Step 5: Run population
    print_step(5, "Populate Mailboxes")
    
    success = populator.run(
        all_mailboxes=all_mailboxes,
        count=count,
        specific=specific,
        department=department,
        emails_per_mailbox=emails_per_mailbox,
        emails_min=emails_min,
        emails_max=emails_max,
        dry_run=args.dry_run,
        folders=folders,
    )
    
    # Step 6: Summary
    print_step(6, "Summary")
    populator.print_final_summary()
    
    sys.exit(0 if success else 1)


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
