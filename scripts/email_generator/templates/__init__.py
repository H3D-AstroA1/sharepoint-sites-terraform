"""
Email templates package for M365 email population.

This package contains all email templates organized by category:
- Newsletter templates (company and industry newsletters)
- SharePoint templates (document sharing, site activity)
- Attachment templates (reports, documents for review)
- Organisational templates (announcements, HR policies, leadership messages)
- Interdepartmental templates (project updates, meeting requests, status reports)
- Security templates (account blocked, password reset, suspicious activity)
- Spam templates (promotional, phishing simulations, lottery scams)
"""

from typing import Dict, List, Any

# Import newsletter templates
from .newsletter_templates import (
    COMPANY_NEWSLETTER,
    INDUSTRY_NEWSLETTER,
    NEWSLETTER_TEMPLATES,
)

# Import SharePoint templates
from .sharepoint_templates import (
    SHAREPOINT_DOCUMENT_SHARED,
    SHAREPOINT_SITE_ACTIVITY,
    SHAREPOINT_TEMPLATES,
)

# Import attachment templates
from .attachment_templates import (
    REPORT_WITH_ATTACHMENT,
    DOCUMENT_FOR_REVIEW,
    ATTACHMENT_TEMPLATES,
)

# Import organisational templates
from .organisational_templates import (
    COMPANY_ANNOUNCEMENT,
    HR_POLICY_UPDATE,
    LEADERSHIP_MESSAGE,
    ORGANISATIONAL_TEMPLATES,
)

# Import interdepartmental templates
from .interdepartmental_templates import (
    PROJECT_UPDATE,
    MEETING_REQUEST,
    STATUS_REPORT,
    COLLABORATION_REQUEST,
    INTERDEPARTMENTAL_TEMPLATES,
    THREADABLE_TEMPLATES,
)

# Import security templates
from .security_templates import (
    ACCOUNT_BLOCKED,
    PASSWORD_RESET_WITH_TEMP,
    PASSWORD_RESET_LINK,
    ACCOUNT_UNLOCKED,
    SUSPICIOUS_ACTIVITY,
    SECURITY_TEMPLATES,
    PASSWORD_TEMPLATES,
)

# Import spam templates
from .spam_templates import (
    PROMOTIONAL_SPAM,
    PHISHING_BANK,
    LOTTERY_SCAM,
    FAKE_INVOICE,
    NEWSLETTER_SPAM,
    SPAM_TEMPLATES,
    SPAM_SENDER_DOMAINS,
    SPAM_SENDER_NAMES,
)

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
}

# All templates that support email threading
ALL_THREADABLE_TEMPLATES = THREADABLE_TEMPLATES

# All templates that require attachments
ALL_ATTACHMENT_TEMPLATES = ATTACHMENT_TEMPLATES

# All templates with temporary passwords (for special handling)
ALL_PASSWORD_TEMPLATES = PASSWORD_TEMPLATES

# =============================================================================
# PUBLIC API
# =============================================================================

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
    
    # Combined collections
    "ALL_THREADABLE_TEMPLATES",
    "ALL_ATTACHMENT_TEMPLATES",
    "ALL_PASSWORD_TEMPLATES",
]
