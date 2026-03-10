"""
M365 Email Generator Package

This package provides functionality to generate and populate realistic
organisational emails in Microsoft 365 mailboxes.

Modules:
    - config: Configuration loading and validation
    - templates: Email template definitions
    - content_generator: Dynamic content generation
    - attachments: Attachment file generation
    - threading: Email thread management
    - graph_client: Microsoft Graph API client
    - azure_ad_discovery: Azure AD user/group discovery
    - user_pool: User and group pool management
    - utils: Utility functions
"""

from .config import (
    load_mailbox_config,
    load_sites_config,
    load_environment_config,
    get_user_email_count,
    validate_config,
    is_azure_ad_enabled,
    get_azure_ad_config,
    get_cc_bcc_config,
)

from .content_generator import EmailContentGenerator
from .graph_client import GraphClient
from .threading import ThreadManager
from .attachments import AttachmentGenerator

# Azure AD discovery imports
from .azure_ad_discovery import (
    AzureADDiscovery,
    AzureADUser,
    AzureADGroup,
    DiscoveryCache,
    UserCategory,
    RecipientType,
)

# User pool imports
from .user_pool import (
    UserPool,
    SenderPool,
    EmailRecipient,
    RecipientSelection,
    UserSource,
)

__version__ = "1.1.0"
__all__ = [
    # Config
    "load_mailbox_config",
    "load_sites_config",
    "load_environment_config",
    "get_user_email_count",
    "validate_config",
    "is_azure_ad_enabled",
    "get_azure_ad_config",
    "get_cc_bcc_config",
    # Core generators
    "EmailContentGenerator",
    "GraphClient",
    "ThreadManager",
    "AttachmentGenerator",
    # Azure AD discovery
    "AzureADDiscovery",
    "AzureADUser",
    "AzureADGroup",
    "DiscoveryCache",
    "UserCategory",
    "RecipientType",
    # User pool
    "UserPool",
    "SenderPool",
    "EmailRecipient",
    "RecipientSelection",
    "UserSource",
]
