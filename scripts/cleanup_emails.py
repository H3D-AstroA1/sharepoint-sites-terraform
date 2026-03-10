#!/usr/bin/env python3
"""
M365 Email Cleanup Script

Deletes emails from Microsoft 365 mailboxes. Supports:
- Delete all emails from specific mailboxes
- Delete emails from specific folders
- Permanently delete or move to Deleted Items
- Empty Deleted Items folder

Usage:
    python cleanup_emails.py --all                    # Delete from all configured mailboxes
    python cleanup_emails.py --mailboxes user@domain.com
    python cleanup_emails.py --folder inbox --permanent
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any

# Add parent directory to path for imports
SCRIPT_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(SCRIPT_DIR))

from email_generator.config import load_mailbox_config, get_all_users
from email_generator.graph_client import GraphClient
from email_generator.utils import (
    Colors,
    print_banner,
    print_step,
    print_success,
    print_error,
    print_warning,
    print_info,
    print_summary_box,
    confirm_action,
    print_progress,
)


class EmailCleaner:
    """Handles email cleanup operations."""
    
    def __init__(self, config_path: Optional[str] = None):
        """Initialize the email cleaner."""
        self.config = load_mailbox_config(config_path)
        self.graph_client: Optional[GraphClient] = None
        self.stats = {
            "mailboxes_processed": 0,
            "emails_deleted": 0,
            "emails_failed": 0,
            "start_time": None,
            "end_time": None,
        }
    
    def initialize(self) -> bool:
        """Initialize the Graph client and authenticate."""
        print_step(1, "Authenticating with Microsoft Graph")
        
        self.graph_client = GraphClient()
        
        if not self.graph_client.authenticate():
            print_error("Failed to authenticate with Microsoft Graph")
            print_info("Make sure you're logged in with 'az login' or have app credentials configured")
            return False
        
        print_success("Authentication successful")
        
        # Verify permissions
        print_info("Verifying permissions...")
        perm_check = self.graph_client.verify_permissions()
        
        if not perm_check.get("has_permissions"):
            print_warning("Some permissions may be missing")
            if perm_check.get("error"):
                print_info(f"  {perm_check['error']}")
        else:
            print_success("Permissions verified")
        
        return True
    
    def get_mailboxes(
        self,
        all_mailboxes: bool = False,
        specific: Optional[List[str]] = None
    ) -> List[Dict[str, Any]]:
        """Get list of mailboxes to clean."""
        users = get_all_users(self.config)
        
        if not users:
            return []
        
        # Filter by specific mailboxes
        if specific:
            specific_lower = [s.lower() for s in specific]
            users = [u for u in users if u.get("upn", "").lower() in specific_lower]
        
        return users
    
    def get_mailbox_email_count(self, mailbox: str, folder: str = "inbox") -> int:
        """Get the number of emails in a mailbox folder."""
        if not self.graph_client:
            return 0
        return self.graph_client.get_email_count(mailbox, folder)
    
    def delete_emails_from_mailbox(
        self,
        mailbox: str,
        folder: str = "inbox",
        permanent: bool = False,
        max_emails: int = 10000
    ) -> Dict[str, int]:
        """Delete emails from a single mailbox."""
        if not self.graph_client:
            return {"success": 0, "failed": 0}
        
        return self.graph_client.delete_all_emails(mailbox, folder, permanent, max_emails)
    
    def empty_deleted_items(self, mailbox: str) -> Dict[str, int]:
        """Empty the Deleted Items folder for a mailbox."""
        if not self.graph_client:
            return {"success": 0, "failed": 0}
        
        return self.graph_client.empty_deleted_items(mailbox)
    
    def run(
        self,
        all_mailboxes: bool = False,
        specific: Optional[List[str]] = None,
        folder: str = "inbox",
        permanent: bool = False,
        empty_trash: bool = False,
        dry_run: bool = False
    ) -> bool:
        """
        Main execution method.
        
        Args:
            all_mailboxes: Clean all configured mailboxes.
            specific: List of specific mailbox UPNs.
            folder: Folder to clean (inbox, sentitems, etc.).
            permanent: Permanently delete (skip Deleted Items).
            empty_trash: Also empty Deleted Items folder.
            dry_run: Preview without deleting.
            
        Returns:
            True if successful, False otherwise.
        """
        self.stats["start_time"] = datetime.now()
        
        # Get mailboxes
        mailboxes = self.get_mailboxes(all_mailboxes, specific)
        
        if not mailboxes:
            print_error("No mailboxes selected for cleanup")
            return False
        
        print_success(f"Selected {len(mailboxes)} mailboxes for cleanup")
        
        # Validate mailboxes
        if not dry_run:
            print_info("Validating mailboxes...")
            valid_mailboxes = []
            
            for user in mailboxes:
                upn = user.get("upn", "")
                verification = self.graph_client.verify_mailbox(upn)
                
                if verification.get("exists") and verification.get("has_mailbox"):
                    valid_mailboxes.append(user)
                    print(f"    {Colors.GREEN}✓{Colors.NC} {upn}")
                else:
                    print(f"    {Colors.RED}✗{Colors.NC} {upn} - {verification.get('error', 'Not found')}")
            
            print()
            
            if not valid_mailboxes:
                print_error("No valid mailboxes found")
                return False
            
            mailboxes = valid_mailboxes
        
        # Get email counts
        print_info("Getting email counts...")
        mailbox_counts = []
        total_emails = 0
        
        for user in mailboxes:
            upn = user.get("upn", "")
            if dry_run:
                count = "N/A (dry run)"
            else:
                count = self.get_mailbox_email_count(upn, folder)
                total_emails += count if isinstance(count, int) else 0
            mailbox_counts.append((user, count))
            print(f"    {upn}: {count} emails in {folder}")
        
        print()
        
        # Show summary
        print_summary_box("Email Cleanup Summary", [
            ("Mailboxes:", len(mailboxes)),
            ("Folder:", folder),
            ("Total Emails:", total_emails if not dry_run else "N/A"),
            ("Delete Mode:", "PERMANENT" if permanent else "Move to Deleted Items"),
            ("Empty Trash:", "Yes" if empty_trash else "No"),
            ("Mode:", "DRY RUN" if dry_run else "LIVE"),
        ])
        
        # Confirm
        if not dry_run:
            warning_msg = "⚠️  WARNING: This will DELETE emails!"
            if permanent:
                warning_msg += " PERMANENTLY!"
            print(f"  {Colors.RED}{warning_msg}{Colors.NC}")
            print()
            
            if not confirm_action("Proceed with email deletion?"):
                print_warning("Operation cancelled")
                return False
        
        # Process mailboxes
        print()
        for i, (user, count) in enumerate(mailbox_counts):
            upn = user.get("upn", "Unknown")
            
            print(f"  {Colors.CYAN}[{i+1}/{len(mailboxes)}]{Colors.NC} {upn}")
            
            if dry_run:
                print(f"       {Colors.YELLOW}[DRY RUN]{Colors.NC} Would delete emails from {folder}")
                self.stats["mailboxes_processed"] += 1
            else:
                # Delete emails
                print(f"       Deleting emails from {folder}...")
                results = self.delete_emails_from_mailbox(upn, folder, permanent)
                
                self.stats["emails_deleted"] += results["success"]
                self.stats["emails_failed"] += results["failed"]
                self.stats["mailboxes_processed"] += 1
                
                print(f"       {Colors.GREEN}✓{Colors.NC} Deleted: {results['success']}, Failed: {results['failed']}")
                
                # Empty trash if requested
                if empty_trash:
                    print(f"       Emptying Deleted Items...")
                    trash_results = self.empty_deleted_items(upn)
                    print(f"       {Colors.GREEN}✓{Colors.NC} Emptied: {trash_results['success']} items")
            
            print()
        
        self.stats["end_time"] = datetime.now()
        
        # Final summary
        duration = (self.stats["end_time"] - self.stats["start_time"]).total_seconds()
        
        print_summary_box("Cleanup Complete", [
            ("Mailboxes Processed:", self.stats["mailboxes_processed"]),
            ("Emails Deleted:", self.stats["emails_deleted"]),
            ("Emails Failed:", self.stats["emails_failed"]),
            ("Duration:", f"{duration:.1f} seconds"),
        ])
        
        return True


def interactive_mode():
    """Run in interactive mode."""
    print_banner("M365 Email Cleanup")
    print()
    
    cleaner = EmailCleaner()
    
    if not cleaner.initialize():
        return
    
    print()
    print_step(2, "Select Mailboxes")
    print()
    print(f"  {Colors.WHITE}Which mailboxes do you want to clean?{Colors.NC}")
    print()
    print(f"    {Colors.GREEN}[1]{Colors.NC} All mailboxes from config")
    print(f"    {Colors.BLUE}[2]{Colors.NC} Specific mailboxes")
    print(f"    {Colors.RED}[Q]{Colors.NC} Quit")
    print()
    
    choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip().lower()
    
    if choice == 'q':
        return
    
    all_mailboxes = choice == '1'
    specific = None
    
    if choice == '2':
        print()
        print(f"  {Colors.WHITE}Enter mailbox UPNs (comma-separated):{Colors.NC}")
        mailbox_input = input(f"  {Colors.YELLOW}Mailboxes:{Colors.NC} ").strip()
        if mailbox_input:
            specific = [m.strip() for m in mailbox_input.split(",")]
        else:
            print_error("No mailboxes specified")
            return
    
    # Select folder
    print()
    print_step(3, "Select Folder")
    print()
    print(f"  {Colors.WHITE}Which folder do you want to clean?{Colors.NC}")
    print()
    print(f"    {Colors.GREEN}[1]{Colors.NC} Inbox")
    print(f"    {Colors.BLUE}[2]{Colors.NC} Sent Items")
    print(f"    {Colors.YELLOW}[3]{Colors.NC} Deleted Items")
    print(f"    {Colors.MAGENTA}[4]{Colors.NC} All folders (inbox + sent)")
    print()
    
    folder_choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip()
    
    folder_map = {
        '1': 'inbox',
        '2': 'sentitems',
        '3': 'deleteditems',
        '4': 'inbox',  # Will handle both
    }
    folder = folder_map.get(folder_choice, 'inbox')
    
    # Delete mode
    print()
    print_step(4, "Delete Mode")
    print()
    print(f"  {Colors.WHITE}How do you want to delete emails?{Colors.NC}")
    print()
    print(f"    {Colors.GREEN}[1]{Colors.NC} Move to Deleted Items (recoverable)")
    print(f"    {Colors.RED}[2]{Colors.NC} Permanently delete (not recoverable)")
    print()
    
    mode_choice = input(f"  {Colors.YELLOW}Enter your choice:{Colors.NC} ").strip()
    permanent = mode_choice == '2'
    
    # Empty trash option
    empty_trash = False
    if not permanent:
        print()
        print(f"  {Colors.WHITE}Also empty Deleted Items folder?{Colors.NC}")
        empty_choice = input(f"  {Colors.YELLOW}[y/N]:{Colors.NC} ").strip().lower()
        empty_trash = empty_choice == 'y'
    
    # Run cleanup
    print()
    print_step(5, "Running Cleanup")
    print()
    
    cleaner.run(
        all_mailboxes=all_mailboxes,
        specific=specific,
        folder=folder,
        permanent=permanent,
        empty_trash=empty_trash,
        dry_run=False
    )


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Delete emails from M365 mailboxes",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python cleanup_emails.py                           # Interactive mode
  python cleanup_emails.py --all                     # Delete from all mailboxes
  python cleanup_emails.py --mailboxes user@domain.com
  python cleanup_emails.py --all --folder sentitems  # Delete sent items
  python cleanup_emails.py --all --permanent         # Permanently delete
  python cleanup_emails.py --all --empty-trash       # Also empty Deleted Items
  python cleanup_emails.py --all --dry-run           # Preview only
        """
    )
    
    # Mailbox selection
    selection = parser.add_mutually_exclusive_group()
    selection.add_argument(
        "--all",
        action="store_true",
        help="Clean all mailboxes from config"
    )
    selection.add_argument(
        "--mailboxes",
        type=str,
        help="Comma-separated list of mailbox UPNs"
    )
    
    # Options
    parser.add_argument(
        "--folder",
        type=str,
        default="inbox",
        choices=["inbox", "sentitems", "deleteditems", "drafts"],
        help="Folder to clean (default: inbox)"
    )
    parser.add_argument(
        "--permanent",
        action="store_true",
        help="Permanently delete (skip Deleted Items)"
    )
    parser.add_argument(
        "--empty-trash",
        action="store_true",
        help="Also empty Deleted Items folder"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview without deleting"
    )
    parser.add_argument(
        "--config",
        type=str,
        help="Path to mailboxes.yaml config file"
    )
    
    args = parser.parse_args()
    
    # If no arguments, run interactive mode
    if not args.all and not args.mailboxes:
        interactive_mode()
        return
    
    # Command line mode
    print_banner("M365 Email Cleanup")
    print()
    
    cleaner = EmailCleaner(args.config)
    
    if not cleaner.initialize():
        sys.exit(1)
    
    print()
    
    # Parse mailboxes
    specific = None
    if args.mailboxes:
        specific = [m.strip() for m in args.mailboxes.split(",")]
    
    # Run cleanup
    success = cleaner.run(
        all_mailboxes=args.all,
        specific=specific,
        folder=args.folder,
        permanent=args.permanent,
        empty_trash=args.empty_trash,
        dry_run=args.dry_run
    )
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print()
        print(f"  {Colors.YELLOW}⚠{Colors.NC} Operation cancelled by user")
        sys.exit(0)
