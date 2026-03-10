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
    - utils: Utility functions
"""

from .config import (
    load_mailbox_config,
    load_sites_config,
    load_environment_config,
    get_user_email_count,
    validate_config,
)

from .content_generator import EmailContentGenerator
from .graph_client import GraphClient
from .threading import ThreadManager
from .attachments import AttachmentGenerator

__version__ = "1.0.0"
__all__ = [
    "load_mailbox_config",
    "load_sites_config", 
    "load_environment_config",
    "get_user_email_count",
    "validate_config",
    "EmailContentGenerator",
    "GraphClient",
    "ThreadManager",
    "AttachmentGenerator",
]
