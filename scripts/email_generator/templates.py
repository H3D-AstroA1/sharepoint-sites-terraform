"""
Email templates for M365 email population.

This module re-exports all templates from the templates/ package for backward compatibility.
All templates are now organized in separate files under the templates/ folder.

Template Categories:
- newsletters: Company and industry newsletters
- links: SharePoint document sharing and site activity
- attachments: Reports and documents for review
- organisational: Company announcements, HR policies, leadership messages
- interdepartmental: Project updates, meeting requests, status reports
- security: Account blocked, password reset, suspicious activity alerts
- spam: Promotional spam, phishing simulations, lottery scams
"""

# Import directly from the templates subpackage modules to avoid circular imports
from .templates.newsletter_templates import (
    COMPANY_NEWSLETTER,
    INDUSTRY_NEWSLETTER,
    NEWSLETTER_TEMPLATES,
)

from .templates.sharepoint_templates import (
    SHAREPOINT_DOCUMENT_SHARED,
    SHAREPOINT_SITE_ACTIVITY,
    SHAREPOINT_TEMPLATES,
)

from .templates.attachment_templates import (
    REPORT_WITH_ATTACHMENT,
    DOCUMENT_FOR_REVIEW,
    ATTACHMENT_TEMPLATES,
)

from .templates.organisational_templates import (
    COMPANY_ANNOUNCEMENT,
    HR_POLICY_UPDATE,
    LEADERSHIP_MESSAGE,
    ORGANISATIONAL_TEMPLATES,
)

from .templates.interdepartmental_templates import (
    PROJECT_UPDATE,
    MEETING_REQUEST,
    STATUS_REPORT,
    COLLABORATION_REQUEST,
    INTERDEPARTMENTAL_TEMPLATES,
    THREADABLE_TEMPLATES,
)

from .templates.security_templates import (
    ACCOUNT_BLOCKED,
    PASSWORD_RESET_WITH_TEMP,
    PASSWORD_RESET_LINK,
    ACCOUNT_UNLOCKED,
    SUSPICIOUS_ACTIVITY,
    SECURITY_TEMPLATES,
    PASSWORD_TEMPLATES,
)

from .templates.spam_templates import (
    PROMOTIONAL_SPAM,
    PHISHING_BANK,
    LOTTERY_SCAM,
    FAKE_INVOICE,
    NEWSLETTER_SPAM,
    SPAM_TEMPLATES,
    SPAM_SENDER_DOMAINS,
    SPAM_SENDER_NAMES,
)

from .templates.external_business_templates import (
    EXTERNAL_BUSINESS_TEMPLATES,
    EXTERNAL_BUSINESS_DOMAINS,
    EXTERNAL_SENDER_PROFILES,
)

from typing import Dict, List, Any

# =============================================================================
# COMBINED TEMPLATE COLLECTIONS
# =============================================================================

# All templates organized by category (main export)
EMAIL_TEMPLATES: Dict[str, List[Dict[str, Any]]] = {
    "newsletters": NEWSLETTER_TEMPLATES,
    "links": SHAREPOINT_TEMPLATES,
    "attachments": ATTACHMENT_TEMPLATES,
    "organisational": ORGANISATIONAL_TEMPLATES,
    "interdepartmental": INTERDEPARTMENTAL_TEMPLATES,
    "security": SECURITY_TEMPLATES,
    "spam": SPAM_TEMPLATES,
    "external_business": EXTERNAL_BUSINESS_TEMPLATES,
}

# All templates that support email threading
ALL_THREADABLE_TEMPLATES = THREADABLE_TEMPLATES

# All templates that require attachments
ALL_ATTACHMENT_TEMPLATES = ATTACHMENT_TEMPLATES

# All templates with temporary passwords (for special handling)
ALL_PASSWORD_TEMPLATES = PASSWORD_TEMPLATES

__all__ = [
    # Main template dictionary
    "EMAIL_TEMPLATES",
    
    # Newsletter templates
    "COMPANY_NEWSLETTER",
    "INDUSTRY_NEWSLETTER",
    "NEWSLETTER_TEMPLATES",
    
    # SharePoint templates
    "SHAREPOINT_DOCUMENT_SHARED",
    "SHAREPOINT_SITE_ACTIVITY",
    "SHAREPOINT_TEMPLATES",
    
    # Attachment templates
    "REPORT_WITH_ATTACHMENT",
    "DOCUMENT_FOR_REVIEW",
    "ATTACHMENT_TEMPLATES",
    
    # Organisational templates
    "COMPANY_ANNOUNCEMENT",
    "HR_POLICY_UPDATE",
    "LEADERSHIP_MESSAGE",
    "ORGANISATIONAL_TEMPLATES",
    
    # Interdepartmental templates
    "PROJECT_UPDATE",
    "MEETING_REQUEST",
    "STATUS_REPORT",
    "COLLABORATION_REQUEST",
    "INTERDEPARTMENTAL_TEMPLATES",
    "THREADABLE_TEMPLATES",
    
    # Security templates
    "ACCOUNT_BLOCKED",
    "PASSWORD_RESET_WITH_TEMP",
    "PASSWORD_RESET_LINK",
    "ACCOUNT_UNLOCKED",
    "SUSPICIOUS_ACTIVITY",
    "SECURITY_TEMPLATES",
    "PASSWORD_TEMPLATES",
    
    # Spam templates
    "PROMOTIONAL_SPAM",
    "PHISHING_BANK",
    "LOTTERY_SCAM",
    "FAKE_INVOICE",
    "NEWSLETTER_SPAM",
    "SPAM_TEMPLATES",
    "SPAM_SENDER_DOMAINS",
    "SPAM_SENDER_NAMES",
    
    # External business templates
    "EXTERNAL_BUSINESS_TEMPLATES",
    "EXTERNAL_BUSINESS_DOMAINS",
    "EXTERNAL_SENDER_PROFILES",
    
    # Combined collections
    "ALL_THREADABLE_TEMPLATES",
    "ALL_ATTACHMENT_TEMPLATES",
    "ALL_PASSWORD_TEMPLATES",
]
