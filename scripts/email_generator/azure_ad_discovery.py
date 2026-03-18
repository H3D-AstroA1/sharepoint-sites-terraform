"""
Azure AD Auto-Discovery Module for M365 Email Population.

This module provides functionality to discover users, groups, and distribution
lists from Azure AD for use in email population scenarios.
"""

import json
import urllib.request
import urllib.error
import urllib.parse
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Set, Tuple
from .utils import print_info, print_error, print_warning, print_success


class UserCategory(Enum):
    """Categories of users in Azure AD."""
    MAILBOX_USER = "mailbox"           # Has Exchange license/mailbox
    NON_MAILBOX_USER = "non_mailbox"   # Azure AD user, no mailbox
    EXTERNAL_USER = "external"          # Guest users
    SERVICE_ACCOUNT = "service"         # Service/app accounts


class RecipientType(Enum):
    """Types of recipients for email addressing."""
    USER = "user"                       # Individual user
    M365_GROUP = "m365_group"           # Microsoft 365 Group
    SECURITY_GROUP = "security_group"   # Security Group (mail-enabled)
    DISTRIBUTION_LIST = "distribution"  # Distribution List
    MAIL_CONTACT = "mail_contact"       # External mail contact


@dataclass
class AzureADUser:
    """Represents a user from Azure AD."""
    id: str
    upn: str
    display_name: str
    email: Optional[str] = None
    department: Optional[str] = None
    job_title: Optional[str] = None
    manager_id: Optional[str] = None
    office_location: Optional[str] = None
    has_mailbox: bool = False
    category: UserCategory = UserCategory.NON_MAILBOX_USER
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization."""
        return {
            "id": self.id,
            "upn": self.upn,
            "display_name": self.display_name,
            "email": self.email,
            "department": self.department,
            "job_title": self.job_title,
            "manager_id": self.manager_id,
            "office_location": self.office_location,
            "has_mailbox": self.has_mailbox,
            "category": self.category.value,
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AzureADUser":
        """Create from dictionary."""
        return cls(
            id=data["id"],
            upn=data["upn"],
            display_name=data["display_name"],
            email=data.get("email"),
            department=data.get("department"),
            job_title=data.get("job_title"),
            manager_id=data.get("manager_id"),
            office_location=data.get("office_location"),
            has_mailbox=data.get("has_mailbox", False),
            category=UserCategory(data.get("category", "non_mailbox")),
        )
    
    @classmethod
    def from_graph_response(cls, data: Dict[str, Any]) -> "AzureADUser":
        """Create from Microsoft Graph API response."""
        # Determine category
        category = UserCategory.NON_MAILBOX_USER
        upn = data.get("userPrincipalName", "")
        
        if "#EXT#" in upn:
            category = UserCategory.EXTERNAL_USER
        elif data.get("userType") == "Guest":
            category = UserCategory.EXTERNAL_USER
        
        return cls(
            id=data.get("id", ""),
            upn=upn,
            display_name=data.get("displayName", ""),
            email=data.get("mail"),
            department=data.get("department"),
            job_title=data.get("jobTitle"),
            manager_id=data.get("manager", {}).get("id") if isinstance(data.get("manager"), dict) else None,
            office_location=data.get("officeLocation"),
            has_mailbox=False,  # Will be set later after mailbox check
            category=category,
        )


@dataclass
class AzureADGroup:
    """Represents a group from Azure AD."""
    id: str
    display_name: str
    email: Optional[str] = None
    description: Optional[str] = None
    group_type: RecipientType = RecipientType.M365_GROUP
    member_count: int = 0
    members: List[str] = field(default_factory=list)  # List of user IDs
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization."""
        return {
            "id": self.id,
            "display_name": self.display_name,
            "email": self.email,
            "description": self.description,
            "group_type": self.group_type.value,
            "member_count": self.member_count,
            "members": self.members,
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AzureADGroup":
        """Create from dictionary."""
        return cls(
            id=data["id"],
            display_name=data["display_name"],
            email=data.get("email"),
            description=data.get("description"),
            group_type=RecipientType(data.get("group_type", "m365_group")),
            member_count=data.get("member_count", 0),
            members=data.get("members", []),
        )
    
    @classmethod
    def from_graph_response(cls, data: Dict[str, Any]) -> "AzureADGroup":
        """Create from Microsoft Graph API response."""
        # Determine group type
        group_types = data.get("groupTypes", [])
        mail_enabled = data.get("mailEnabled", False)
        security_enabled = data.get("securityEnabled", False)
        
        if "Unified" in group_types:
            group_type = RecipientType.M365_GROUP
        elif mail_enabled and not security_enabled:
            group_type = RecipientType.DISTRIBUTION_LIST
        elif security_enabled and mail_enabled:
            group_type = RecipientType.SECURITY_GROUP
        else:
            group_type = RecipientType.SECURITY_GROUP
        
        return cls(
            id=data.get("id", ""),
            display_name=data.get("displayName", ""),
            email=data.get("mail"),
            description=data.get("description"),
            group_type=group_type,
            member_count=0,  # Will be populated later
            members=[],
        )


@dataclass
class DiscoveryCache:
    """Cache for discovered Azure AD data."""
    users: List[AzureADUser] = field(default_factory=list)
    groups: List[AzureADGroup] = field(default_factory=list)
    timestamp: Optional[datetime] = None
    tenant_id: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization."""
        return {
            "users": [u.to_dict() for u in self.users],
            "groups": [g.to_dict() for g in self.groups],
            "timestamp": self.timestamp.isoformat() if self.timestamp else None,
            "tenant_id": self.tenant_id,
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DiscoveryCache":
        """Create from dictionary."""
        return cls(
            users=[AzureADUser.from_dict(u) for u in data.get("users", [])],
            groups=[AzureADGroup.from_dict(g) for g in data.get("groups", [])],
            timestamp=datetime.fromisoformat(data["timestamp"]) if data.get("timestamp") else None,
            tenant_id=data.get("tenant_id"),
        )
    
    def is_valid(self, ttl_minutes: int = 60) -> bool:
        """Check if cache is still valid."""
        if not self.timestamp:
            return False
        return datetime.now() - self.timestamp < timedelta(minutes=ttl_minutes)
    
    @property
    def mailbox_users(self) -> List[AzureADUser]:
        """Get users with mailboxes."""
        return [u for u in self.users if u.has_mailbox]
    
    @property
    def non_mailbox_users(self) -> List[AzureADUser]:
        """Get users without mailboxes."""
        return [u for u in self.users if not u.has_mailbox]
    
    def get_users_by_department(self, department: str) -> List[AzureADUser]:
        """Get users in a specific department."""
        return [u for u in self.users if u.department == department]
    
    def get_departments(self) -> Dict[str, Dict[str, int]]:
        """Get department statistics."""
        departments: Dict[str, Dict[str, int]] = {}
        for user in self.users:
            dept = user.department or "Unknown"
            if dept not in departments:
                departments[dept] = {"total": 0, "mailbox": 0}
            departments[dept]["total"] += 1
            if user.has_mailbox:
                departments[dept]["mailbox"] += 1
        return departments


class AzureADDiscovery:
    """
    Azure AD Discovery Service.
    
    Discovers users, groups, and distribution lists from Azure AD
    for use in email population scenarios.
    """
    
    CACHE_FILE = ".azure_ad_cache.json"
    
    def __init__(self, token: str, config: Optional[Dict[str, Any]] = None):
        """
        Initialize the discovery service.
        
        Args:
            token: Microsoft Graph API access token
            config: Optional configuration dictionary
        """
        self.token = token
        self.config = config or {}
        self.cache = DiscoveryCache()
        self._load_cache()
    
    def _make_request(
        self,
        url: str,
        method: str = "GET",
        timeout: int = 30
    ) -> Optional[Dict[str, Any]]:
        """Make a request to Microsoft Graph API."""
        try:
            req = urllib.request.Request(url, method=method)
            req.add_header("Authorization", f"Bearer {self.token}")
            req.add_header("Content-Type", "application/json")
            
            with urllib.request.urlopen(req, timeout=timeout) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            print_error(f"Graph API error: {e.code} - {e.reason}")
            return None
        except Exception as e:
            print_error(f"Request failed: {e}")
            return None
    
    def _paginate_request(
        self,
        url: str,
        max_items: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """Make paginated requests to Microsoft Graph API."""
        items: List[Dict[str, Any]] = []
        next_url: Optional[str] = url
        
        while next_url:
            response = self._make_request(next_url)
            if not response:
                break
            
            items.extend(response.get("value", []))
            
            # Check if we've reached the limit
            if max_items and len(items) >= max_items:
                items = items[:max_items]
                break
            
            # Get next page URL
            next_url = response.get("@odata.nextLink")
        
        return items
    
    def discover_users(
        self,
        max_users: Optional[int] = None,
        include_departments: Optional[List[str]] = None,
        exclude_departments: Optional[List[str]] = None,
        exclude_external: bool = False,
        exclude_service_accounts: bool = True,
        validate_mailboxes: bool = True,
        mailbox_validation_limit: int = 100,
        progress_callback: Optional[Callable] = None
    ) -> List[AzureADUser]:
        """
        Discover users from Azure AD.
        
        Args:
            max_users: Maximum number of users to discover
            include_departments: Only include these departments
            exclude_departments: Exclude these departments
            exclude_external: Exclude guest/external users
            exclude_service_accounts: Exclude service accounts
            validate_mailboxes: Check if users have mailboxes
            mailbox_validation_limit: Max users to validate for mailboxes (0 = unlimited, default 100)
            progress_callback: Callback for progress updates
            
        Returns:
            List of discovered users
        """
        print_info("Discovering users from Azure AD...")
        
        # Build query
        select_fields = "id,userPrincipalName,displayName,mail,department,jobTitle,officeLocation,accountEnabled,userType"
        url = f"https://graph.microsoft.com/v1.0/users?$select={select_fields}&$top=999"
        
        # Add filter for enabled accounts (URL-encoded)
        filters = ["accountEnabled eq true"]
        
        if filters:
            filter_str = urllib.parse.quote(' and '.join(filters))
            url += f"&$filter={filter_str}"
        
        # Fetch users
        raw_users = self._paginate_request(url, max_users)
        print_info(f"Found {len(raw_users)} users in Azure AD")
        
        # Convert to AzureADUser objects
        users: List[AzureADUser] = []
        for i, raw_user in enumerate(raw_users):
            user = AzureADUser.from_graph_response(raw_user)
            
            # Apply filters
            if exclude_external and user.category == UserCategory.EXTERNAL_USER:
                continue
            
            if exclude_service_accounts:
                upn_lower = user.upn.lower()
                if any(pattern in upn_lower for pattern in ["service", "admin", "system", "noreply", "mailbox"]):
                    continue
            
            if include_departments and user.department not in include_departments:
                continue
            
            if exclude_departments and user.department in exclude_departments:
                continue
            
            users.append(user)
            
            if progress_callback and i % 100 == 0:
                progress_callback(i, len(raw_users), "Filtering users...")
        
        print_info(f"Filtered to {len(users)} users")
        
        # Validate mailboxes (with optional limit for performance)
        if validate_mailboxes:
            users = self._validate_mailboxes(users, mailbox_validation_limit, progress_callback)
        
        self.cache.users = users
        self.cache.timestamp = datetime.now()
        
        return users
    
    def _validate_mailboxes(
        self,
        users: List[AzureADUser],
        limit: int = 100,
        progress_callback: Optional[Callable] = None
    ) -> List[AzureADUser]:
        """Validate which users have mailboxes.
        
        Args:
            users: List of users to validate
            limit: Maximum number of users to validate (0 = unlimited, default 100)
            progress_callback: Callback for progress updates
            
        Returns:
            List of users (with has_mailbox flag set for validated users)
        """
        total_users = len(users)
        
        # Apply limit if specified
        if limit > 0 and total_users > limit:
            print_info(f"Validating mailboxes for first {limit} of {total_users} users (limit applied for performance)")
            users_to_validate = users[:limit]
            remaining_users = users[limit:]
        else:
            print_info(f"Validating mailboxes for {total_users} users...")
            users_to_validate = users
            remaining_users = []
        
        mailbox_count = 0
        for i, user in enumerate(users_to_validate):
            # Check if user has a mailbox by trying to access their inbox
            url = f"https://graph.microsoft.com/v1.0/users/{user.id}/mailFolders/inbox"
            response = self._make_request(url)
            
            if response:
                user.has_mailbox = True
                user.category = UserCategory.MAILBOX_USER
                mailbox_count += 1
            
            if progress_callback and i % 50 == 0:
                progress_callback(i, len(users_to_validate), f"Validating mailboxes... ({mailbox_count} found)")
        
        print_info(f"Found {mailbox_count} users with mailboxes (validated {len(users_to_validate)} of {total_users})")
        
        # Return all users (validated + remaining unvalidated)
        return users_to_validate + remaining_users
    
    def discover_groups(
        self,
        max_groups: Optional[int] = None,
        include_m365_groups: bool = True,
        include_security_groups: bool = True,
        include_distribution_lists: bool = True,
        exclude_patterns: Optional[List[str]] = None,
        min_members: int = 2,
        progress_callback: Optional[Callable] = None
    ) -> List[AzureADGroup]:
        """
        Discover groups from Azure AD.
        
        Args:
            max_groups: Maximum number of groups to discover
            include_m365_groups: Include Microsoft 365 groups
            include_security_groups: Include mail-enabled security groups
            include_distribution_lists: Include distribution lists
            exclude_patterns: Patterns to exclude (e.g., ["Test*", "Dev*"])
            min_members: Minimum number of members
            progress_callback: Callback for progress updates
            
        Returns:
            List of discovered groups
        """
        print_info("Discovering groups from Azure AD...")
        
        groups: List[AzureADGroup] = []
        
        # Discover M365 groups
        if include_m365_groups:
            filter_str = urllib.parse.quote("groupTypes/any(c:c eq 'Unified')")
            url = f"https://graph.microsoft.com/v1.0/groups?$filter={filter_str}&$select=id,displayName,mail,description,groupTypes,mailEnabled,securityEnabled&$top=999"
            raw_groups = self._paginate_request(url, max_groups)
            
            for raw_group in raw_groups:
                group = AzureADGroup.from_graph_response(raw_group)
                if group.email:  # Only include mail-enabled groups
                    groups.append(group)
        
        # Discover distribution lists
        if include_distribution_lists:
            filter_str = urllib.parse.quote("mailEnabled eq true and securityEnabled eq false")
            url = f"https://graph.microsoft.com/v1.0/groups?$filter={filter_str}&$select=id,displayName,mail,description,groupTypes,mailEnabled,securityEnabled&$top=999"
            raw_groups = self._paginate_request(url, max_groups)
            
            for raw_group in raw_groups:
                group = AzureADGroup.from_graph_response(raw_group)
                if group.email and group.id not in [g.id for g in groups]:
                    groups.append(group)
        
        # Discover mail-enabled security groups
        if include_security_groups:
            filter_str = urllib.parse.quote("mailEnabled eq true and securityEnabled eq true")
            url = f"https://graph.microsoft.com/v1.0/groups?$filter={filter_str}&$select=id,displayName,mail,description,groupTypes,mailEnabled,securityEnabled&$top=999"
            raw_groups = self._paginate_request(url, max_groups)
            
            for raw_group in raw_groups:
                group = AzureADGroup.from_graph_response(raw_group)
                if group.email and group.id not in [g.id for g in groups]:
                    groups.append(group)
        
        print_info(f"Found {len(groups)} groups")
        
        # Apply filters
        filtered_groups: List[AzureADGroup] = []
        for group in groups:
            # Check exclude patterns
            if exclude_patterns:
                excluded = False
                for pattern in exclude_patterns:
                    if pattern.endswith("*"):
                        if group.display_name.lower().startswith(pattern[:-1].lower()):
                            excluded = True
                            break
                    elif pattern.lower() in group.display_name.lower():
                        excluded = True
                        break
                if excluded:
                    continue
            
            filtered_groups.append(group)
        
        # Get member counts
        for i, group in enumerate(filtered_groups):
            url = f"https://graph.microsoft.com/v1.0/groups/{group.id}/members/$count"
            # Use a simpler approach - get members and count
            members_url = f"https://graph.microsoft.com/v1.0/groups/{group.id}/members?$select=id&$top=999"
            members_response = self._make_request(members_url)
            if members_response:
                group.member_count = len(members_response.get("value", []))
                group.members = [m.get("id") for m in members_response.get("value", [])]
            
            if progress_callback and i % 10 == 0:
                progress_callback(i, len(filtered_groups), "Getting group members...")
        
        # Filter by minimum members
        filtered_groups = [g for g in filtered_groups if g.member_count >= min_members]
        
        print_info(f"Filtered to {len(filtered_groups)} groups")
        
        self.cache.groups = filtered_groups
        self.cache.timestamp = datetime.now()
        
        return filtered_groups
    
    def discover_all(
        self,
        config: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable] = None
    ) -> DiscoveryCache:
        """
        Discover all users and groups.
        
        Args:
            config: Configuration dictionary
            progress_callback: Callback for progress updates
            
        Returns:
            DiscoveryCache with all discovered data
        """
        cfg = config or self.config.get("azure_ad", {})
        
        # Discover users
        user_config = cfg.get("users", {})
        self.discover_users(
            max_users=user_config.get("max_users"),
            include_departments=user_config.get("include_departments"),
            exclude_departments=user_config.get("exclude_departments"),
            exclude_external=user_config.get("exclude_external_users", False),
            exclude_service_accounts=user_config.get("exclude_service_accounts", True),
            validate_mailboxes=user_config.get("validate_mailbox_exists", True),
            progress_callback=progress_callback,
        )
        
        # Discover groups
        group_config = cfg.get("groups", {})
        if group_config.get("enabled", True):
            self.discover_groups(
                max_groups=group_config.get("max_groups"),
                include_m365_groups=group_config.get("include_m365_groups", True),
                include_security_groups=group_config.get("include_security_groups", True),
                include_distribution_lists=group_config.get("include_distribution_lists", True),
                exclude_patterns=group_config.get("exclude_patterns"),
                min_members=group_config.get("min_members", 2),
                progress_callback=progress_callback,
            )
        
        self._save_cache()
        return self.cache
    
    def _get_cache_path(self) -> Path:
        """Get the cache file path."""
        cache_config = self.config.get("azure_ad", {}).get("cache", {})
        cache_file = cache_config.get("cache_file", self.CACHE_FILE)
        return Path(cache_file)
    
    def _load_cache(self) -> None:
        """Load cache from file."""
        cache_path = self._get_cache_path()
        if cache_path.exists():
            try:
                with open(cache_path, "r") as f:
                    data = json.load(f)
                    self.cache = DiscoveryCache.from_dict(data)
                    print_info(f"Loaded cache with {len(self.cache.users)} users and {len(self.cache.groups)} groups")
            except Exception as e:
                print_warning(f"Failed to load cache: {e}")
                self.cache = DiscoveryCache()
    
    def _save_cache(self) -> None:
        """Save cache to file."""
        cache_path = self._get_cache_path()
        try:
            with open(cache_path, "w") as f:
                json.dump(self.cache.to_dict(), f, indent=2)
            print_info(f"Saved cache to {cache_path}")
        except Exception as e:
            print_warning(f"Failed to save cache: {e}")
    
    def clear_cache(self) -> None:
        """Clear the cache."""
        self.cache = DiscoveryCache()
        cache_path = self._get_cache_path()
        if cache_path.exists():
            cache_path.unlink()
            print_info("Cache cleared")
    
    def get_cache(self) -> DiscoveryCache:
        """Get the current cache."""
        return self.cache
    
    def is_cache_valid(self) -> bool:
        """Check if cache is valid."""
        ttl = self.config.get("azure_ad", {}).get("cache", {}).get("ttl_minutes", 60)
        return self.cache.is_valid(ttl)
    
    def get_statistics(self) -> Dict[str, Any]:
        """Get discovery statistics."""
        return {
            "total_users": len(self.cache.users),
            "mailbox_users": len(self.cache.mailbox_users),
            "non_mailbox_users": len(self.cache.non_mailbox_users),
            "groups": len(self.cache.groups),
            "m365_groups": len([g for g in self.cache.groups if g.group_type == RecipientType.M365_GROUP]),
            "distribution_lists": len([g for g in self.cache.groups if g.group_type == RecipientType.DISTRIBUTION_LIST]),
            "security_groups": len([g for g in self.cache.groups if g.group_type == RecipientType.SECURITY_GROUP]),
            "departments": self.cache.get_departments(),
            "cache_timestamp": self.cache.timestamp.isoformat() if self.cache.timestamp else None,
            "cache_valid": self.is_cache_valid(),
        }
