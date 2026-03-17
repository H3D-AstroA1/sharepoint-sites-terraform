"""
User Pool Module for M365 Email Population.

This module manages pools of users and groups for email generation,
supporting both YAML-configured users and Azure AD discovered users.
Includes support for exclusion filtering.
"""

import random
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional, Set, Tuple

from .azure_ad_discovery import (
    AzureADUser,
    AzureADGroup,
    DiscoveryCache,
    UserCategory,
    RecipientType,
)
from .config import (
    is_exclusions_enabled,
    is_email_excluded,
    should_log_exclusions,
)
from .utils import print_info, print_warning


class UserSource(Enum):
    """Source of user data."""
    YAML = "yaml"           # From mailboxes.yaml configuration
    AZURE_AD = "azure_ad"   # From Azure AD discovery
    MIXED = "mixed"         # Combination of both


@dataclass
class EmailRecipient:
    """Represents an email recipient (To, CC, or BCC)."""
    email: str
    display_name: str
    recipient_type: RecipientType = RecipientType.USER
    has_mailbox: bool = False
    department: Optional[str] = None
    
    def to_graph_format(self) -> Dict[str, Any]:
        """Convert to Microsoft Graph API format."""
        return {
            "emailAddress": {
                "address": self.email,
                "name": self.display_name,
            }
        }
    
    @classmethod
    def from_azure_ad_user(cls, user: AzureADUser) -> "EmailRecipient":
        """Create from AzureADUser."""
        return cls(
            email=user.email or user.upn,
            display_name=user.display_name,
            recipient_type=RecipientType.USER,
            has_mailbox=user.has_mailbox,
            department=user.department,
        )
    
    @classmethod
    def from_azure_ad_group(cls, group: AzureADGroup) -> "EmailRecipient":
        """Create from AzureADGroup."""
        return cls(
            email=group.email or "",
            display_name=group.display_name,
            recipient_type=group.group_type,
            has_mailbox=False,  # Groups don't have mailboxes
            department=None,
        )
    
    @classmethod
    def from_yaml_config(cls, config: Dict[str, Any]) -> "EmailRecipient":
        """Create from YAML configuration."""
        return cls(
            email=config.get("upn", config.get("email", "")),
            display_name=config.get("display_name", config.get("name", "")),
            recipient_type=RecipientType.USER,
            has_mailbox=True,  # YAML users are assumed to have mailboxes
            department=config.get("department"),
        )


@dataclass
class RecipientSelection:
    """Represents a selection of recipients for an email."""
    to: List[EmailRecipient] = field(default_factory=list)
    cc: List[EmailRecipient] = field(default_factory=list)
    bcc: List[EmailRecipient] = field(default_factory=list)
    
    def to_graph_format(self) -> Dict[str, Any]:
        """Convert to Microsoft Graph API format."""
        result: Dict[str, Any] = {}
        
        if self.to:
            result["toRecipients"] = [r.to_graph_format() for r in self.to]
        
        if self.cc:
            result["ccRecipients"] = [r.to_graph_format() for r in self.cc]
        
        if self.bcc:
            result["bccRecipients"] = [r.to_graph_format() for r in self.bcc]
        
        return result
    
    @property
    def all_recipients(self) -> List[EmailRecipient]:
        """Get all recipients."""
        return self.to + self.cc + self.bcc
    
    @property
    def total_count(self) -> int:
        """Get total recipient count."""
        return len(self.to) + len(self.cc) + len(self.bcc)


class UserPool:
    """
    Manages pools of users and groups for email generation.
    
    Supports:
    - YAML-configured users (existing functionality)
    - Azure AD discovered users
    - Mixed mode (both sources)
    - CC/BCC recipient selection
    - Department-based filtering
    - Weighted selection based on mailbox status
    """
    
    def __init__(
        self,
        config: Optional[Dict[str, Any]] = None,
        discovery_cache: Optional[DiscoveryCache] = None,
    ):
        """
        Initialize the user pool.
        
        Args:
            config: Configuration dictionary (from mailboxes.yaml)
            discovery_cache: Azure AD discovery cache
        """
        self.config = config or {}
        self.discovery_cache = discovery_cache
        
        # User pools
        self._mailbox_users: List[EmailRecipient] = []
        self._non_mailbox_users: List[EmailRecipient] = []
        self._all_users: List[EmailRecipient] = []
        self._groups: List[EmailRecipient] = []
        
        # Configuration
        self._cc_bcc_config = self.config.get("cc_bcc", {})
        self._source = UserSource.YAML
        
        # Initialize pools
        self._initialize_pools()
    
    def _initialize_pools(self) -> None:
        """Initialize user pools from available sources."""
        # Load from YAML config
        yaml_users = self._load_yaml_users()
        
        # Load from Azure AD cache
        azure_users, azure_groups = self._load_azure_ad_data()
        
        # Determine source
        if yaml_users and azure_users:
            self._source = UserSource.MIXED
        elif azure_users:
            self._source = UserSource.AZURE_AD
        else:
            self._source = UserSource.YAML
        
        # Combine users (Azure AD takes precedence for duplicates)
        user_map: Dict[str, EmailRecipient] = {}
        
        # Add YAML users first
        for user in yaml_users:
            user_map[user.email.lower()] = user
        
        # Add/override with Azure AD users
        for user in azure_users:
            user_map[user.email.lower()] = user
        
        # Track exclusion statistics
        excluded_count = 0
        should_log = should_log_exclusions(self.config)
        
        # Categorize users (with exclusion filtering)
        for user in user_map.values():
            # Check if user should be excluded
            is_excluded, reason = is_email_excluded(user.email, self.config)
            
            if is_excluded:
                excluded_count += 1
                if should_log:
                    print_warning(f"Excluding {user.email}: {reason}")
                continue
            
            self._all_users.append(user)
            if user.has_mailbox:
                self._mailbox_users.append(user)
            else:
                self._non_mailbox_users.append(user)
        
        # Add groups (also filter excluded domains)
        for group in azure_groups:
            is_excluded, reason = is_email_excluded(group.email, self.config)
            if not is_excluded:
                self._groups.append(group)
            elif should_log:
                print_warning(f"Excluding group {group.email}: {reason}")
        
        # Log summary
        if excluded_count > 0:
            print_info(f"User pool initialized: {len(self._mailbox_users)} mailbox users, "
                       f"{len(self._non_mailbox_users)} non-mailbox users, "
                       f"{len(self._groups)} groups "
                       f"({excluded_count} excluded)")
        else:
            print_info(f"User pool initialized: {len(self._mailbox_users)} mailbox users, "
                       f"{len(self._non_mailbox_users)} non-mailbox users, "
                       f"{len(self._groups)} groups")
    
    def _load_yaml_users(self) -> List[EmailRecipient]:
        """Load users from YAML configuration."""
        users: List[EmailRecipient] = []
        
        mailboxes = self.config.get("mailboxes", [])
        for mailbox in mailboxes:
            if mailbox.get("enabled", True):
                users.append(EmailRecipient.from_yaml_config(mailbox))
        
        return users
    
    def _load_azure_ad_data(self) -> Tuple[List[EmailRecipient], List[EmailRecipient]]:
        """Load users and groups from Azure AD discovery cache."""
        users: List[EmailRecipient] = []
        groups: List[EmailRecipient] = []
        
        if not self.discovery_cache:
            return users, groups
        
        # Load users
        for azure_user in self.discovery_cache.users:
            if azure_user.email or azure_user.upn:
                users.append(EmailRecipient.from_azure_ad_user(azure_user))
        
        # Load groups
        for azure_group in self.discovery_cache.groups:
            if azure_group.email:
                groups.append(EmailRecipient.from_azure_ad_group(azure_group))
        
        return users, groups
    
    @property
    def source(self) -> UserSource:
        """Get the current user source."""
        return self._source
    
    @property
    def mailbox_users(self) -> List[EmailRecipient]:
        """Get users with mailboxes."""
        return self._mailbox_users
    
    @property
    def non_mailbox_users(self) -> List[EmailRecipient]:
        """Get users without mailboxes."""
        return self._non_mailbox_users
    
    @property
    def all_users(self) -> List[EmailRecipient]:
        """Get all users."""
        return self._all_users
    
    @property
    def groups(self) -> List[EmailRecipient]:
        """Get all groups."""
        return self._groups
    
    def get_random_sender(
        self,
        exclude_upn: Optional[str] = None,
        department: Optional[str] = None,
        require_mailbox: bool = True,
    ) -> Optional[EmailRecipient]:
        """
        Get a random sender from the pool.
        
        Args:
            exclude_upn: UPN to exclude (e.g., the recipient)
            department: Prefer senders from this department
            require_mailbox: Only select users with mailboxes
            
        Returns:
            Random sender or None if no suitable sender found
        """
        pool = self._mailbox_users if require_mailbox else self._all_users
        
        if not pool:
            return None
        
        # Filter by department if specified
        if department:
            dept_users = [u for u in pool if u.department == department]
            if dept_users:
                pool = dept_users
        
        # Exclude specific UPN
        if exclude_upn:
            pool = [u for u in pool if u.email.lower() != exclude_upn.lower()]
        
        if not pool:
            return None
        
        return random.choice(pool)
    
    def get_random_recipients(
        self,
        count: int = 1,
        exclude_upns: Optional[List[str]] = None,
        department: Optional[str] = None,
        include_non_mailbox: bool = True,
        include_groups: bool = False,
    ) -> List[EmailRecipient]:
        """
        Get random recipients from the pool.
        
        Args:
            count: Number of recipients to select
            exclude_upns: UPNs to exclude
            department: Prefer recipients from this department
            include_non_mailbox: Include users without mailboxes
            include_groups: Include groups in selection
            
        Returns:
            List of random recipients
        """
        pool: List[EmailRecipient] = []
        
        # Build pool
        if include_non_mailbox:
            pool.extend(self._all_users)
        else:
            pool.extend(self._mailbox_users)
        
        if include_groups:
            pool.extend(self._groups)
        
        if not pool:
            return []
        
        # Filter by department if specified
        if department:
            dept_pool = [r for r in pool if r.department == department or r.recipient_type != RecipientType.USER]
            if dept_pool:
                pool = dept_pool
        
        # Exclude specific UPNs
        if exclude_upns:
            exclude_set = {upn.lower() for upn in exclude_upns}
            pool = [r for r in pool if r.email.lower() not in exclude_set]
        
        if not pool:
            return []
        
        # Select random recipients
        count = min(count, len(pool))
        return random.sample(pool, count)
    
    def generate_recipient_selection(
        self,
        mailbox_upn: str,
        sender_email: str,
        category: str = "general",
        folder: str = "inbox",
    ) -> RecipientSelection:
        """
        Generate a complete recipient selection for an email.
        
        Args:
            mailbox_upn: The mailbox UPN (for inbox, this is the recipient)
            sender_email: The sender's email
            category: Email category (affects CC/BCC probability)
            folder: Target folder (inbox, sentitems, drafts)
            
        Returns:
            RecipientSelection with To, CC, and BCC recipients
        """
        selection = RecipientSelection()
        
        # Get CC/BCC configuration
        cc_config = self._cc_bcc_config.get("cc", {})
        bcc_config = self._cc_bcc_config.get("bcc", {})
        
        cc_enabled = cc_config.get("enabled", True)
        bcc_enabled = bcc_config.get("enabled", True)
        
        cc_probability = cc_config.get("probability", 0.3)
        bcc_probability = bcc_config.get("probability", 0.1)
        
        cc_max = cc_config.get("max_recipients", 5)
        bcc_max = bcc_config.get("max_recipients", 3)
        
        # Adjust probabilities based on category
        category_adjustments = {
            "newsletter": {"cc": 0.1, "bcc": 0.05},
            "security": {"cc": 0.2, "bcc": 0.3},
            "spam": {"cc": 0.0, "bcc": 0.0},
            "organisational": {"cc": 0.5, "bcc": 0.2},
            "interdepartmental": {"cc": 0.4, "bcc": 0.1},
        }
        
        if category in category_adjustments:
            cc_probability = category_adjustments[category]["cc"]
            bcc_probability = category_adjustments[category]["bcc"]
        
        # Exclude sender and mailbox owner from recipients
        exclude_upns = [mailbox_upn, sender_email]
        
        # For inbox: mailbox owner is the primary recipient
        # For sentitems/drafts: need to select recipients
        if folder == "inbox":
            # The mailbox owner is the recipient
            mailbox_user = self._find_user_by_email(mailbox_upn)
            if mailbox_user:
                selection.to.append(mailbox_user)
            else:
                # Create a basic recipient
                selection.to.append(EmailRecipient(
                    email=mailbox_upn,
                    display_name=mailbox_upn.split("@")[0].replace(".", " ").title(),
                    has_mailbox=True,
                ))
        else:
            # For sent items and drafts, select random recipients
            to_count = random.randint(1, 3)
            selection.to = self.get_random_recipients(
                count=to_count,
                exclude_upns=exclude_upns,
                include_non_mailbox=True,
                include_groups=random.random() < 0.2,  # 20% chance of group
            )
            
            # Update exclude list
            exclude_upns.extend([r.email for r in selection.to])
        
        # Add CC recipients
        if cc_enabled and random.random() < cc_probability:
            cc_count = random.randint(1, cc_max)
            selection.cc = self.get_random_recipients(
                count=cc_count,
                exclude_upns=exclude_upns,
                include_non_mailbox=True,
                include_groups=random.random() < 0.1,
            )
            exclude_upns.extend([r.email for r in selection.cc])
        
        # Add BCC recipients
        if bcc_enabled and random.random() < bcc_probability:
            bcc_count = random.randint(1, bcc_max)
            selection.bcc = self.get_random_recipients(
                count=bcc_count,
                exclude_upns=exclude_upns,
                include_non_mailbox=True,
                include_groups=False,  # BCC typically doesn't include groups
            )
        
        return selection
    
    def _find_user_by_email(self, email: str) -> Optional[EmailRecipient]:
        """Find a user by email address."""
        email_lower = email.lower()
        for user in self._all_users:
            if user.email.lower() == email_lower:
                return user
        return None
    
    def get_users_by_department(self, department: str) -> List[EmailRecipient]:
        """Get all users in a specific department."""
        return [u for u in self._all_users if u.department == department]
    
    def get_departments(self) -> List[str]:
        """Get list of all departments."""
        departments: Set[str] = set()
        for user in self._all_users:
            if user.department:
                departments.add(user.department)
        return sorted(departments)
    
    def get_statistics(self) -> Dict[str, Any]:
        """Get pool statistics."""
        return {
            "source": self._source.value,
            "total_users": len(self._all_users),
            "mailbox_users": len(self._mailbox_users),
            "non_mailbox_users": len(self._non_mailbox_users),
            "groups": len(self._groups),
            "departments": len(self.get_departments()),
            "cc_bcc_enabled": {
                "cc": self._cc_bcc_config.get("cc", {}).get("enabled", True),
                "bcc": self._cc_bcc_config.get("bcc", {}).get("enabled", True),
            },
        }
    
    def refresh_from_cache(self, discovery_cache: DiscoveryCache) -> None:
        """Refresh the pool from a new discovery cache."""
        self.discovery_cache = discovery_cache
        self._mailbox_users.clear()
        self._non_mailbox_users.clear()
        self._all_users.clear()
        self._groups.clear()
        self._initialize_pools()


class SenderPool:
    """
    Specialized pool for selecting email senders.
    
    Provides weighted selection based on:
    - Department matching
    - Sender type (internal, external, system)
    - Historical patterns
    """
    
    def __init__(self, user_pool: UserPool, config: Optional[Dict[str, Any]] = None):
        """
        Initialize the sender pool.
        
        Args:
            user_pool: The main user pool
            config: Sender-specific configuration
        """
        self.user_pool = user_pool
        self.config = config or {}
        
        # Sender type weights
        self._weights = self.config.get("sender_weights", {
            "internal": 0.7,
            "external": 0.2,
            "system": 0.1,
        })
    
    def select_sender(
        self,
        recipient: EmailRecipient,
        category: str = "general",
        sender_type: Optional[str] = None,
    ) -> Dict[str, str]:
        """
        Select an appropriate sender for an email.
        
        Args:
            recipient: The primary recipient
            category: Email category
            sender_type: Force a specific sender type
            
        Returns:
            Sender dictionary with name and email
        """
        # Determine sender type if not specified
        if not sender_type:
            sender_type = self._select_sender_type(category)
        
        if sender_type == "internal":
            return self._get_internal_sender(recipient)
        elif sender_type == "external":
            return self._get_external_sender(category)
        else:
            return self._get_system_sender(recipient)
    
    def _select_sender_type(self, category: str) -> str:
        """Select sender type based on category and weights."""
        # Override for specific categories
        if category == "spam":
            return "external"
        elif category == "security":
            return random.choice(["system", "internal"])
        elif category == "newsletter":
            return random.choice(["system", "internal"])
        
        # Use weighted random selection
        types = list(self._weights.keys())
        weights = list(self._weights.values())
        return random.choices(types, weights=weights)[0]
    
    def _get_internal_sender(self, recipient: EmailRecipient) -> Dict[str, str]:
        """Get an internal sender."""
        sender = self.user_pool.get_random_sender(
            exclude_upn=recipient.email,
            department=recipient.department,  # Prefer same department
            require_mailbox=True,
        )
        
        if sender:
            return {
                "name": sender.display_name,
                "email": sender.email,
            }
        
        # Fallback to generic internal sender
        return {
            "name": "Internal User",
            "email": "user@contoso.com",
        }
    
    def _get_external_sender(self, category: str) -> Dict[str, str]:
        """Get an external sender."""
        external_domains = [
            "partner.com", "vendor.com", "client.com",
            "consultant.com", "supplier.com", "agency.com",
        ]
        
        first_names = ["John", "Jane", "Michael", "Sarah", "David", "Emily"]
        last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Davis"]
        
        first = random.choice(first_names)
        last = random.choice(last_names)
        domain = random.choice(external_domains)
        
        return {
            "name": f"{first} {last}",
            "email": f"{first.lower()}.{last.lower()}@{domain}",
        }
    
    def _get_system_sender(self, recipient: EmailRecipient) -> Dict[str, str]:
        """Get a system sender."""
        # Extract domain from recipient
        domain = recipient.email.split("@")[-1] if "@" in recipient.email else "contoso.com"
        
        system_senders = [
            {"name": "IT Support", "prefix": "it-support"},
            {"name": "HR Department", "prefix": "hr"},
            {"name": "Security Team", "prefix": "security"},
            {"name": "Admin", "prefix": "admin"},
            {"name": "Notifications", "prefix": "noreply"},
        ]
        
        sender = random.choice(system_senders)
        return {
            "name": sender["name"],
            "email": f"{sender['prefix']}@{domain}",
        }
