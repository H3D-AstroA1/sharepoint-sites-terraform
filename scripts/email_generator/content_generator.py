"""
Email content generation with realistic organisational content.

Generates dynamic email content based on templates, recipient context,
and SharePoint site integration. Uses the variations module for
extensive content randomization and variation.
"""

import random
import string
import secrets
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any, Tuple

from .templates import EMAIL_TEMPLATES, THREADABLE_TEMPLATES, ATTACHMENT_TEMPLATES, PASSWORD_TEMPLATES
from . import variations


class EmailContentGenerator:
    """Generates realistic email content based on templates and context."""
    
    def __init__(self, config: Dict[str, Any], sites: Dict[str, Any], user_pool: Any = None):
        """
        Initialize the content generator.
        
        Args:
            config: Mailbox configuration dictionary.
            sites: SharePoint sites configuration dictionary.
            user_pool: Optional UserPool instance for CC/BCC generation.
        """
        self.config = config
        self.sites = sites
        self.settings = config.get("settings", {})
        self.departments = config.get("departments", {})
        self.external_senders = config.get("external_senders", {})
        self.user_pool = user_pool
        
        # Build site lookup
        self.site_lookup = self._build_site_lookup()
        
    def _build_site_lookup(self) -> Dict[str, Dict]:
        """Build a lookup dictionary for SharePoint sites."""
        lookup = {}
        for site in self.sites.get("sites", []):
            name = site.get("name", "")
            lookup[name] = site
        return lookup
    
    def generate_email(self, recipient: Dict[str, Any], folder: str = "inbox") -> Dict[str, Any]:
        """
        Generate a complete email for a recipient.
        
        Args:
            recipient: Recipient user configuration.
            folder: Target folder (inbox, sentitems, drafts).
            
        Returns:
            Dictionary containing email data.
        """
        # Select email category based on distribution
        category = self._select_category()
        
        # Select template from category
        template = self._select_template(category)
        
        # Generate sender
        sender = self._select_sender(template, recipient)
        
        # Generate date
        email_date = self._generate_date()
        
        # Generate subject
        subject = self._generate_subject(template, recipient, sender)
        
        # Generate body
        body = self._generate_body(template, recipient, sender)
        
        # Select sensitivity label
        sensitivity = self._select_sensitivity(category, recipient)
        
        # Generate importance level
        importance = self._generate_importance(category, template)
        
        # Generate flag status (follow-up flags)
        flag_status = self._generate_flag_status(email_date, category)
        
        # Generate categories (color tags)
        categories = self._generate_categories(category, recipient)
        
        # Generate threading info (reply/forward)
        threading_info = self._generate_threading_info(category, template, sender, recipient)
        
        # Apply Re:/Fwd: prefix to subject based on threading info
        if threading_info.get("is_reply") or threading_info.get("is_reply_all"):
            if not subject.lower().startswith("re:"):
                subject = f"Re: {subject}"
        elif threading_info.get("is_forward"):
            if not subject.lower().startswith("fwd:") and not subject.lower().startswith("fw:"):
                subject = f"Fwd: {subject}"
        
        # Build email object
        email = {
            "category": category,
            "template_name": template.get("subject_templates", [""])[0][:30],
            "subject": subject,
            "body": body,
            "sender": sender,
            "recipient": {
                "email": recipient.get("upn"),
                "name": recipient.get("display_name", ""),
            },
            "date": email_date,
            "sensitivity_label": sensitivity,
            "has_attachment": template.get("has_attachment", False),
            "supports_threading": template.get("supports_threading", False),
            # New realistic properties
            "importance": importance,
            "flag_status": flag_status,
            "categories": categories,
            "threading": threading_info,
        }
        
        # Add CC/BCC recipients if user_pool is available
        if self.user_pool:
            cc_bcc = self._generate_cc_bcc_recipients(
                mailbox_upn=recipient.get("upn", ""),
                sender_email=sender.get("email", ""),
                category=category,
                folder=folder
            )
            if cc_bcc.get("cc_recipients"):
                email["cc_recipients"] = cc_bcc["cc_recipients"]
            if cc_bcc.get("bcc_recipients"):
                email["bcc_recipients"] = cc_bcc["bcc_recipients"]
        
        # Add attachment type if needed
        if email["has_attachment"]:
            email["attachment_type"] = self._select_attachment_type(recipient)
            email["attachment_department"] = recipient.get("department", "General")
        
        return email
    
    def _generate_cc_bcc_recipients(
        self,
        mailbox_upn: str,
        sender_email: str,
        category: str,
        folder: str
    ) -> Dict[str, Any]:
        """
        Generate CC and BCC recipients using the user pool.
        
        Args:
            mailbox_upn: The mailbox UPN.
            sender_email: The sender's email.
            category: Email category.
            folder: Target folder.
            
        Returns:
            Dictionary with cc_recipients and bcc_recipients lists.
        """
        result: Dict[str, Any] = {}
        
        if not self.user_pool:
            return result
        
        try:
            # Use the user pool's generate_recipient_selection method
            selection = self.user_pool.generate_recipient_selection(
                mailbox_upn=mailbox_upn,
                sender_email=sender_email,
                category=category,
                folder=folder
            )
            
            # Convert CC recipients to dict format
            if selection.cc:
                result["cc_recipients"] = [
                    {"name": r.display_name, "email": r.email}
                    for r in selection.cc
                ]
            
            # Convert BCC recipients to dict format
            if selection.bcc:
                result["bcc_recipients"] = [
                    {"name": r.display_name, "email": r.email}
                    for r in selection.bcc
                ]
        except Exception:
            # If anything fails, just return empty - CC/BCC is optional
            pass
        
        return result
    
    def _select_category(self) -> str:
        """Select email category based on configured distribution."""
        distribution = self.settings.get("email_distribution", {
            "newsletters": 12,
            "links": 18,
            "attachments": 15,
            "organisational": 15,
            "interdepartmental": 18,
            "security": 12,
            "spam": 10,  # Realistic spam percentage
        })
        
        categories = list(distribution.keys())
        weights = list(distribution.values())
        
        return random.choices(categories, weights=weights)[0]
    
    def _select_template(self, category: str) -> Dict[str, Any]:
        """Select a random template from the category."""
        templates = EMAIL_TEMPLATES.get(category, [])
        if not templates:
            # Fallback to newsletters if category not found
            templates = EMAIL_TEMPLATES.get("newsletters", [])
        
        return random.choice(templates)
    
    def _select_sender(self, template: Dict, recipient: Dict) -> Dict[str, str]:
        """Select appropriate sender based on template and recipient."""
        sender_type = template.get("sender_type", "internal_users")
        
        # IMPORTANT: Spam templates MUST always use spam senders
        # Never override spam sender type with internal users
        if sender_type == "external_spam":
            return self._get_spam_sender()
        
        # Override with distribution-based selection sometimes (but not for spam)
        distribution = self.settings.get("sender_distribution", {})
        if distribution and random.random() < 0.3:  # 30% chance to use distribution
            # Exclude external_spam from random selection to prevent internal users sending spam
            non_spam_distribution = {k: v for k, v in distribution.items() if k != "external_spam"}
            if non_spam_distribution:
                sender_type = random.choices(
                    list(non_spam_distribution.keys()),
                    weights=list(non_spam_distribution.values())
                )[0]
        
        if sender_type == "external_business":
            return self._get_external_business_sender()
        elif sender_type == "external":
            return self._get_external_sender(template)
        elif sender_type == "internal_system":
            return self._get_system_sender(recipient)
        else:
            return self._get_internal_user_sender(recipient)
    
    def _get_internal_user_sender(self, recipient: Dict) -> Dict[str, str]:
        """Get an internal user as sender."""
        # Get users from config
        users = self.config.get("users", [])
        
        # Filter out the recipient
        other_users = [u for u in users if u.get("upn") != recipient.get("upn")]
        
        if other_users:
            sender_user = random.choice(other_users)
            upn = sender_user.get("upn", "")
            name_part = upn.split("@")[0]
            display_name = " ".join(p.capitalize() for p in name_part.replace(".", " ").split())
            
            return {
                "email": upn,
                "name": display_name,
                "title": sender_user.get("job_title", "Employee"),
                "department": sender_user.get("department", "General"),
                "first_name": display_name.split()[0] if display_name else "Colleague",
            }
        
        # Fallback to generic sender
        domain = recipient.get("upn", "").split("@")[-1]
        return {
            "email": f"colleague@{domain}",
            "name": "A Colleague",
            "title": "Employee",
            "department": "General",
            "first_name": "Colleague",
        }
    
    def _get_system_sender(self, recipient: Dict) -> Dict[str, str]:
        """Get a system/department email as sender."""
        department = recipient.get("department", "General")
        dept_config = self.departments.get(department, {})
        
        domain = recipient.get("upn", "").split("@")[-1]
        system_email = dept_config.get("system_email", f"notifications@{domain}")
        system_email = system_email.replace("{domain}", domain)
        
        # Map department to display name
        dept_names = {
            "Human Resources": ("HR Team", "Human Resources"),
            "IT Department": ("IT Support", "IT Department"),
            "Finance Department": ("Finance Team", "Finance"),
            "Executive Leadership": ("Executive Office", "Leadership"),
            "Marketing Department": ("Marketing Team", "Marketing"),
            "Sales Department": ("Sales Team", "Sales"),
            "Legal & Compliance": ("Legal Team", "Legal"),
            "Operations Department": ("Operations Team", "Operations"),
            "Customer Service": ("Support Team", "Customer Service"),
            "Claims Department": ("Claims Team", "Claims"),
        }
        
        name, dept = dept_names.get(department, ("Notifications", "System"))
        
        return {
            "email": system_email,
            "name": name,
            "title": f"{dept} Notifications",
            "department": department,
            "first_name": name.split()[0],
        }
    
    def _get_external_sender(self, template: Dict) -> Dict[str, str]:
        """Get an external sender."""
        category = template.get("category", "newsletters")
        
        # Get external senders from config
        if category == "newsletters":
            senders = self.external_senders.get("newsletters", [])
        else:
            senders = self.external_senders.get("vendors", [])
        
        if senders:
            sender = random.choice(senders)
            return {
                "email": sender.get("email", "newsletter@external.com"),
                "name": sender.get("name", "External Sender"),
                "title": sender.get("type", "Newsletter"),
                "department": "External",
                "first_name": sender.get("name", "Team").split()[0],
            }
        
        # Fallback external senders
        fallback_senders = [
            {"email": "newsletter@techindustryweekly.com", "name": "Tech Industry Weekly"},
            {"email": "updates@businessinsights.com", "name": "Business Insights"},
            {"email": "news@industrynews.com", "name": "Industry News Daily"},
        ]
        sender = random.choice(fallback_senders)
        return {
            "email": sender["email"],
            "name": sender["name"],
            "title": "Newsletter",
            "department": "External",
            "first_name": "Team",
        }
    
    def _get_spam_sender(self) -> Dict[str, str]:
        """Get a spam/junk email sender."""
        from .templates.spam_templates import SPAM_SENDER_DOMAINS, SPAM_SENDER_NAMES
        
        domain = random.choice(SPAM_SENDER_DOMAINS)
        name = random.choice(SPAM_SENDER_NAMES)
        
        # Generate a random-looking email prefix
        prefixes = ["info", "noreply", "support", "team", "admin", "service", "contact", "help"]
        prefix = random.choice(prefixes)
        
        return {
            "email": f"{prefix}@{domain}",
            "name": name,
            "title": "External",
            "department": "External",
            "first_name": name.split()[0],
        }
    
    def _get_external_business_sender(self) -> Dict[str, str]:
        """Get a legitimate external business sender for realistic business communications."""
        from .templates.external_business_templates import (
            EXTERNAL_BUSINESS_DOMAINS,
            EXTERNAL_SENDER_PROFILES,
        )
        
        # Select a random domain and sender profile
        domain = random.choice(EXTERNAL_BUSINESS_DOMAINS)
        # Profile is a tuple: (first_name, last_name, title)
        profile = random.choice(EXTERNAL_SENDER_PROFILES)
        first_name, last_name, title = profile
        
        # Various email formats used by businesses
        email_formats = [
            f"{first_name.lower()}.{last_name.lower()}@{domain}",
            f"{first_name.lower()[0]}{last_name.lower()}@{domain}",
            f"{first_name.lower()}@{domain}",
            f"{first_name.lower()}{last_name.lower()[0]}@{domain}",
        ]
        email = random.choice(email_formats)
        
        # Extract company name from domain (capitalize and clean up)
        company_parts = domain.split(".")[0].replace("-", " ").replace("_", " ")
        company_name = " ".join(word.capitalize() for word in company_parts.split())
        
        return {
            "email": email,
            "name": f"{first_name} {last_name}",
            "title": title,
            "department": "External",
            "first_name": first_name,
            "company": company_name,
        }
    
    def _generate_date(self) -> datetime:
        """Generate a realistic backdated timestamp with good distribution.
        
        Distribution aims for realistic mailbox activity:
        - ~40% of emails from the last 30 days (recent activity)
        - ~30% of emails from 1-3 months ago
        - ~20% of emails from 3-6 months ago
        - ~10% of emails from 6-12 months ago
        """
        date_settings = self.settings.get("date_settings", {})
        months_back = self.settings.get("date_range_months", 12)
        
        # Calculate date range
        end_date = datetime.now()
        
        # Use weighted distribution for more realistic email age spread
        # This ensures we have both recent and older emails
        distribution_choice = random.random()
        
        if distribution_choice < 0.40:
            # 40% - Last 30 days (recent emails)
            days_back = random.randint(0, 30)
        elif distribution_choice < 0.70:
            # 30% - 1-3 months ago
            days_back = random.randint(31, 90)
        elif distribution_choice < 0.90:
            # 20% - 3-6 months ago
            days_back = random.randint(91, 180)
        else:
            # 10% - 6-12 months ago (or configured max)
            max_days = months_back * 30
            days_back = random.randint(181, max(181, max_days))
        
        date = end_date - timedelta(days=days_back)
        
        # Apply weekday bias
        weekday_pct = date_settings.get("weekday_percentage", 90) / 100
        if random.random() < weekday_pct:
            # Ensure it's a weekday (Mon=0, Sun=6)
            while date.weekday() >= 5:
                date -= timedelta(days=1)
        
        # Apply business hours bias
        business_hours_pct = date_settings.get("business_hours_percentage", 85) / 100
        if random.random() < business_hours_pct:
            hour = random.randint(8, 18)
        else:
            hour = random.randint(0, 23)
        
        minute = random.randint(0, 59)
        second = random.randint(0, 59)
        
        return date.replace(hour=hour, minute=minute, second=second)
    
    def _generate_subject(self, template: Dict, recipient: Dict, sender: Dict) -> str:
        """Generate email subject from template."""
        subject_templates = template.get("subject_templates", ["Update"])
        subject = random.choice(subject_templates)
        
        return self._fill_placeholders(subject, recipient, sender)
    
    def _generate_body(self, template: Dict, recipient: Dict, sender: Dict) -> str:
        """Generate email body with dynamic content."""
        body = template.get("body_template", "<p>Email content</p>")
        
        # Fill basic placeholders
        body = self._fill_placeholders(body, recipient, sender)
        
        # Generate dynamic content sections
        body = self._fill_dynamic_content(body, template, recipient, sender)
        
        return body
    
    def _fill_placeholders(self, text: str, recipient: Dict, sender: Dict) -> str:
        """Replace placeholders with actual values."""
        now = datetime.now()
        domain = recipient.get("upn", "company.com").split("@")[-1]
        
        # Get company name from domain
        company_name = domain.split(".")[0].replace("-", " ").title()
        if "onmicrosoft" in domain:
            company_name = domain.split(".")[0].replace("-", " ").title()
        
        # Get SharePoint site info
        department = recipient.get("department", "General")
        site_name = self._get_site_for_department(department)
        site_display = site_name.replace("-", " ").title() if site_name else department
        
        replacements = {
            # Recipient placeholders
            "{recipient_name}": recipient.get("display_name", "Colleague"),
            "{recipient_first_name}": recipient.get("display_name", "Colleague").split()[0],
            "{recipient_email}": recipient.get("upn", ""),
            "{recipient_department}": department,
            
            # Sender placeholders
            "{sender_name}": sender.get("name", "Sender"),
            "{sender_first_name}": sender.get("first_name", "Sender"),
            "{sender_email}": sender.get("email", ""),
            "{sender_title}": sender.get("title", ""),
            "{sender_department}": sender.get("department", ""),
            "{sender_role}": sender.get("title", ""),
            "{sender_phone}": self._generate_phone(),
            "{sender_company}": sender.get("company", "External Company"),
            
            # External business link placeholders
            "{company_website}": f"https://www.{sender.get('email', 'company.com').split('@')[-1] if '@' in sender.get('email', '') else 'company.com'}",
            "{calendar_link}": f"https://calendly.com/{sender.get('first_name', 'contact').lower()}-{secrets.token_hex(4)}",
            "{portal_link}": f"https://portal.{sender.get('email', 'company.com').split('@')[-1] if '@' in sender.get('email', '') else 'company.com'}/client/{secrets.token_urlsafe(8)}",
            "{invoice_link}": f"https://billing.{sender.get('email', 'company.com').split('@')[-1] if '@' in sender.get('email', '') else 'company.com'}/invoice/{secrets.token_hex(8).upper()}",
            "{contract_link}": f"https://docs.{sender.get('email', 'company.com').split('@')[-1] if '@' in sender.get('email', '') else 'company.com'}/contract/{secrets.token_hex(8)}",
            "{meeting_link}": f"https://teams.microsoft.com/l/meetup-join/{secrets.token_urlsafe(32)}",
            "{zoom_link}": f"https://zoom.us/j/{random.randint(10000000000, 99999999999)}?pwd={secrets.token_urlsafe(16)}",
            
            # Company/Domain placeholders
            "{company_name}": company_name,
            "{domain}": domain,
            
            # Date placeholders
            "{date}": now.strftime("%B %d, %Y"),
            "{month}": now.strftime("%B"),
            "{year}": str(now.year),
            "{quarter}": str((now.month - 1) // 3 + 1),
            "{week_number}": str(now.isocalendar()[1]),
            
            # SharePoint placeholders
            "{site_name}": site_display,
            "{site_url}": f"https://{domain.split('.')[0]}.sharepoint.com/sites/{site_name}",
            "{sharepoint_url}": f"https://{domain.split('.')[0]}.sharepoint.com/sites/{site_name}",
            
            # Dynamic content placeholders
            "{department}": department,
            "{industry}": random.choice(["Technology", "Business", "Finance", "Healthcare"]),
            "{newsletter_name}": random.choice(["Industry Insights", "Weekly Digest", "News Roundup", "Harvard Business Review", "Forbes Daily", "TechCrunch Weekly"]),
            "{tagline}": "Your weekly source for industry news and insights",
            
            # Newsletter link placeholders
            "{sender_domain}": sender.get("email", "newsletter.com").split("@")[-1] if "@" in sender.get("email", "") else "newsletter.com",
            "{unsubscribe_id}": secrets.token_urlsafe(16),
            "{newsletter_id}": f"nl-{now.strftime('%Y%m%d')}-{secrets.token_hex(4)}",
            
            # Document placeholders
            "{document_name}": self._generate_document_name(department),
            "{document_type}": random.choice(["Report", "Presentation", "Document", "Spreadsheet"]),
            "{document_icon}": random.choice(["📄", "📊", "📑", "📋"]),
            "{folder_path}": random.choice(["Documents", "Reports", "Shared", "Team Files"]),
            "{file_size}": f"{random.randint(50, 500)} KB",
            "{file_type}": random.choice(["Word Document", "Excel Spreadsheet", "PDF", "PowerPoint"]),
            "{file_icon}": random.choice(["📄", "📊", "📑", "📋"]),
            "{share_date}": now.strftime("%B %d, %Y"),
            
            # Report placeholders
            "{report_type}": random.choice(["Financial", "Status", "Performance", "Analytics", "Budget"]),
            "{period}": f"Q{(now.month - 1) // 3 + 1} {now.year}",
            "{attachment_name}": self._generate_document_name(department),
            
            # Meeting placeholders
            "{meeting_topic}": self._generate_meeting_topic(department),
            "{duration}": random.choice(["30 minutes", "45 minutes", "1 hour"]),
            "{proposed_times}": self._generate_proposed_times(),
            
            # Project placeholders
            "{project_name}": self._generate_project_name(department),
            "{milestone}": random.choice(["Phase 1 Complete", "Review", "Launch", "Planning"]),
            "{update_type}": random.choice(["Progress Update", "Status Change", "Milestone"]),
            
            # Announcement placeholders
            "{announcement_title}": self._generate_announcement_title(),
            "{announcement_type}": random.choice(["Company Update", "Policy Change", "News", "Announcement"]),
            "{effective_date}": (now + timedelta(days=random.randint(7, 30))).strftime("%B %d, %Y"),
            "{deadline}": (now + timedelta(days=random.randint(3, 14))).strftime("%B %d, %Y"),
            
            # Policy placeholders
            "{policy_name}": random.choice(["Remote Work Policy", "Travel Policy", "Expense Policy", "Leave Policy"]),
            
            # Sensitivity
            "{sensitivity_label}": "INTERNAL",
            
            # Misc
            "{topic}": self._generate_topic(department),
            "{message_topic}": self._generate_topic(department),
            "{timeline}": random.choice(["This week", "Next week", "By end of month", "ASAP"]),
            
            # Security placeholders
            "{temp_password}": self._generate_temp_password(),
            "{expiry_hours}": str(random.choice([4, 8, 12, 24])),
            "{expiry_date}": (now + timedelta(hours=random.choice([4, 8, 12, 24]))).strftime("%B %d, %Y at %I:%M %p"),
            "{reset_link}": f"https://passwordreset.{domain}/reset?token={secrets.token_urlsafe(32)}",
            "{request_id}": f"REQ-{secrets.token_hex(4).upper()}",
            "{incident_id}": f"INC-{now.strftime('%Y%m%d')}-{secrets.token_hex(3).upper()}",
            "{ticket_id}": f"TKT-{secrets.token_hex(4).upper()}",
            "{alert_id}": f"ALT-{secrets.token_hex(4).upper()}",
            "{block_reason}": random.choice([
                "Multiple failed login attempts detected",
                "Suspicious login from unrecognized location",
                "Password expired - account locked for security",
                "Potential unauthorized access attempt",
                "Security policy violation detected",
            ]),
            "{detection_time}": (now - timedelta(hours=random.randint(1, 24))).strftime("%B %d, %Y at %I:%M %p"),
            "{detection_location}": random.choice([
                "Unknown location (IP: 185.xxx.xxx.xxx)",
                "Moscow, Russia",
                "Beijing, China",
                "Lagos, Nigeria",
                "Unknown VPN connection",
            ]),
            "{support_phone}": self._generate_phone(),
            "{unlock_time}": now.strftime("%B %d, %Y at %I:%M %p"),
            "{unlocked_by}": "IT Service Desk",
            "{unlock_reason}": random.choice([
                "User verified identity via phone",
                "Password reset completed successfully",
                "Security review completed - no threat detected",
                "Account lockout period expired",
            ]),
            "{signin_time}": (now - timedelta(hours=random.randint(1, 48))).strftime("%B %d, %Y at %I:%M %p"),
            "{signin_location}": random.choice([
                "London, United Kingdom",
                "New York, United States",
                "Sydney, Australia",
                "Unknown location",
            ]),
            "{ip_address}": f"{random.randint(1,255)}.{random.randint(1,255)}.{random.randint(1,255)}.{random.randint(1,255)}",
            "{device_info}": random.choice([
                "Windows 11 PC",
                "MacBook Pro",
                "iPhone 15",
                "Android Device",
                "Unknown Device",
            ]),
            "{browser_info}": random.choice([
                "Chrome 120.0",
                "Safari 17.0",
                "Firefox 121.0",
                "Edge 120.0",
                "Unknown Browser",
            ]),
            "{confirm_link}": f"https://security.{domain}/confirm?token={secrets.token_urlsafe(32)}",
            "{deny_link}": f"https://security.{domain}/deny?token={secrets.token_urlsafe(32)}",
            "{mfa_setup_link}": f"https://security.{domain}/mfa/setup?user={recipient.get('upn', '')}",
            "{help_link}": f"https://support.{domain}/mfa-guide",
        }
        
        for placeholder, value in replacements.items():
            text = text.replace(placeholder, str(value))
        
        return text
    
    def _fill_dynamic_content(self, body: str, template: Dict, recipient: Dict, sender: Dict) -> str:
        """Fill dynamic content sections in the email body."""
        department = recipient.get("department", "General")
        
        # Company news
        if "{company_news}" in body:
            body = body.replace("{company_news}", self._generate_company_news())
        
        # Upcoming events
        if "{upcoming_events}" in body:
            body = body.replace("{upcoming_events}", self._generate_upcoming_events())
        
        # Employee spotlight
        if "{employee_spotlight}" in body:
            body = body.replace("{employee_spotlight}", self._generate_employee_spotlight())
        
        # Fun fact
        if "{fun_fact}" in body:
            body = body.replace("{fun_fact}", self._generate_fun_fact())
        
        # Department updates
        if "{department_updates}" in body:
            body = body.replace("{department_updates}", self._generate_department_updates())
        
        # Articles (for newsletters)
        if "{articles}" in body:
            body = body.replace("{articles}", self._generate_articles())
        
        # Activity items
        if "{activity_items}" in body:
            body = body.replace("{activity_items}", self._generate_activity_items())
        
        # Key points
        if "{key_points}" in body:
            body = body.replace("{key_points}", self._generate_key_points(department))
        
        # Executive summary
        if "{executive_summary}" in body:
            body = body.replace("{executive_summary}", self._generate_executive_summary())
        
        # Action items
        if "{action_items}" in body:
            body = body.replace("{action_items}", self._generate_action_items())
        
        # Personal message
        if "{personal_message}" in body:
            body = body.replace("{personal_message}", self._generate_personal_message())
        
        # Action required
        if "{action_required}" in body:
            body = body.replace("{action_required}", self._generate_action_required())
        
        # Greeting
        if "{greeting}" in body:
            body = body.replace("{greeting}", self._generate_greeting())
        
        # Main announcement
        if "{main_announcement}" in body:
            body = body.replace("{main_announcement}", self._generate_main_announcement())
        
        # Details
        if "{details}" in body:
            body = body.replace("{details}", self._generate_details())
        
        # Contact info
        if "{contact_info}" in body:
            body = body.replace("{contact_info}", self._generate_contact_info(department))
        
        # Project status items
        if "{completed_items}" in body:
            body = body.replace("{completed_items}", self._generate_status_items("completed"))
        if "{in_progress_items}" in body:
            body = body.replace("{in_progress_items}", self._generate_status_items("in_progress"))
        if "{upcoming_items}" in body:
            body = body.replace("{upcoming_items}", self._generate_status_items("upcoming"))
        if "{next_steps}" in body:
            body = body.replace("{next_steps}", self._generate_next_steps())
        
        # Meeting content
        if "{meeting_context}" in body:
            body = body.replace("{meeting_context}", self._generate_meeting_context())
        if "{agenda_items}" in body:
            body = body.replace("{agenda_items}", self._generate_agenda_items())
        
        # Status report content
        if "{metrics}" in body:
            body = body.replace("{metrics}", self._generate_metrics())
        if "{accomplishments}" in body:
            body = body.replace("{accomplishments}", self._generate_accomplishments())
        if "{focus_areas}" in body:
            body = body.replace("{focus_areas}", self._generate_focus_areas())
        if "{blockers}" in body:
            body = body.replace("{blockers}", self._generate_blockers())
        
        # Collaboration content
        if "{request_context}" in body:
            body = body.replace("{request_context}", self._generate_request_context())
        if "{request_details}" in body:
            body = body.replace("{request_details}", self._generate_request_details())
        if "{background}" in body:
            body = body.replace("{background}", self._generate_background())
        
        # Policy content
        if "{policy_overview}" in body:
            body = body.replace("{policy_overview}", self._generate_policy_overview())
        if "{key_changes}" in body:
            body = body.replace("{key_changes}", self._generate_key_changes())
        
        # Document review content
        if "{context}" in body:
            body = body.replace("{context}", self._generate_document_context())
        if "{feedback_areas}" in body:
            body = body.replace("{feedback_areas}", self._generate_feedback_areas())
        
        # Leadership message
        if "{message_body}" in body:
            body = body.replace("{message_body}", self._generate_leadership_message())
        
        # Quoted thread (for threading)
        if "{quoted_thread}" in body:
            body = body.replace("{quoted_thread}", "")  # Will be filled by ThreadManager
        
        # External business content placeholders
        if "{followup_intro}" in body:
            body = body.replace("{followup_intro}", self._generate_followup_intro())
        if "{followup_body}" in body:
            body = body.replace("{followup_body}", self._generate_followup_body())
        if "{followup_action}" in body:
            body = body.replace("{followup_action}", self._generate_followup_action())
        if "{sender_company}" in body:
            body = body.replace("{sender_company}", sender.get("company", "Our Company"))
        if "{sender_phone}" in body:
            body = body.replace("{sender_phone}", self._generate_phone())
        if "{proposal_summary}" in body:
            body = body.replace("{proposal_summary}", self._generate_proposal_summary())
        if "{proposal_highlights}" in body:
            body = body.replace("{proposal_highlights}", self._generate_proposal_highlights())
        if "{proposal_next_steps}" in body:
            body = body.replace("{proposal_next_steps}", self._generate_proposal_next_steps())
        if "{meeting_intro}" in body:
            body = body.replace("{meeting_intro}", self._generate_meeting_intro())
        if "{meeting_duration}" in body:
            body = body.replace("{meeting_duration}", random.choice(["15-minute", "30-minute", "45-minute", "1-hour"]))
        if "{meeting_agenda}" in body:
            body = body.replace("{meeting_agenda}", self._generate_external_meeting_agenda())
        if "{meeting_times}" in body:
            body = body.replace("{meeting_times}", self._generate_meeting_times())
        if "{project_summary}" in body:
            body = body.replace("{project_summary}", self._generate_project_summary())
        if "{project_notes}" in body:
            body = body.replace("{project_notes}", self._generate_project_notes())
        if "{invoice_number}" in body:
            body = body.replace("{invoice_number}", f"{random.randint(10000, 99999)}")
        if "{month}" in body:
            body = body.replace("{month}", random.choice(["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]))
        if "{invoice_details}" in body:
            body = body.replace("{invoice_details}", self._generate_invoice_details())
        if "{payment_terms}" in body:
            body = body.replace("{payment_terms}", self._generate_payment_terms())
        if "{contract_summary}" in body:
            body = body.replace("{contract_summary}", self._generate_contract_summary())
        if "{contract_terms}" in body:
            body = body.replace("{contract_terms}", self._generate_contract_terms())
        if "{contract_next_steps}" in body:
            body = body.replace("{contract_next_steps}", self._generate_contract_next_steps())
        if "{introduction_context}" in body:
            body = body.replace("{introduction_context}", self._generate_introduction_context())
        if "{introduction_body}" in body:
            body = body.replace("{introduction_body}", self._generate_introduction_body())
        if "{introduction_cta}" in body:
            body = body.replace("{introduction_cta}", self._generate_introduction_cta())
        if "{thank_you_context}" in body:
            body = body.replace("{thank_you_context}", self._generate_thank_you_context())
        if "{thank_you_body}" in body:
            body = body.replace("{thank_you_body}", self._generate_thank_you_body())
        if "{support_context}" in body:
            body = body.replace("{support_context}", self._generate_support_context())
        if "{support_details}" in body:
            body = body.replace("{support_details}", self._generate_support_details())
        if "{support_resolution}" in body:
            body = body.replace("{support_resolution}", self._generate_support_resolution())
        if "{event_details}" in body:
            body = body.replace("{event_details}", self._generate_event_details())
        if "{event_agenda}" in body:
            body = body.replace("{event_agenda}", self._generate_event_agenda())
        if "{event_registration}" in body:
            body = body.replace("{event_registration}", self._generate_event_registration())
        if "{product_intro}" in body:
            body = body.replace("{product_intro}", self._generate_product_intro())
        if "{product_features}" in body:
            body = body.replace("{product_features}", self._generate_product_features())
        if "{product_cta}" in body:
            body = body.replace("{product_cta}", self._generate_product_cta())
        if "{feedback_intro}" in body:
            body = body.replace("{feedback_intro}", self._generate_feedback_intro())
        if "{feedback_questions}" in body:
            body = body.replace("{feedback_questions}", self._generate_feedback_questions())
        if "{feedback_cta}" in body:
            body = body.replace("{feedback_cta}", self._generate_feedback_cta())
        
        return body
    
    def _select_sensitivity(self, category: str, recipient: Dict) -> str:
        """Select sensitivity label based on category and recipient."""
        if not self.settings.get("include_sensitivity_labels", True):
            return "general"
        
        sensitivity_config = self.config.get("sensitivity_labels", {})
        department = recipient.get("department", "General")
        
        # Check category-specific sensitivity
        for label_name, label_config in sensitivity_config.items():
            applies_to = label_config.get("applies_to", [])
            departments = label_config.get("departments", [])
            
            if category in applies_to:
                if not departments or department in departments:
                    # Use percentage to determine if this label applies
                    if random.random() * 100 < label_config.get("percentage", 0):
                        return label_config.get("name", label_name)
        
        return "General"
    
    def _select_attachment_type(self, recipient: Dict) -> str:
        """Select attachment type based on recipient's department."""
        department = recipient.get("department", "General")
        
        # Department-specific attachment types
        dept_attachments = {
            "Finance Department": ["xlsx", "pdf"],
            "Human Resources": ["docx", "pdf"],
            "IT Department": ["pptx", "pdf", "docx"],
            "Marketing Department": ["pptx", "pdf"],
            "Sales Department": ["pptx", "xlsx"],
            "Executive Leadership": ["pptx", "pdf"],
            "Legal & Compliance": ["docx", "pdf"],
            "Claims Department": ["docx", "pdf", "xlsx"],
        }
        
        types = dept_attachments.get(department, ["docx", "xlsx", "pptx", "pdf"])
        return random.choice(types)
    
    def _generate_importance(self, category: str, template: Dict) -> str:
        """
        Generate email importance level based on category and template.
        
        Returns:
            'high', 'normal', or 'low'
        """
        # Categories that tend to have high importance
        high_importance_categories = ["security", "organisational"]
        
        # Check template for urgency indicators
        subject_templates = template.get("subject_templates", [])
        has_urgent = any(
            "urgent" in s.lower() or "important" in s.lower() or "action required" in s.lower()
            for s in subject_templates
        )
        
        if has_urgent:
            return random.choices(["high", "normal"], weights=[0.7, 0.3])[0]
        
        if category in high_importance_categories:
            return random.choices(["high", "normal", "low"], weights=[0.3, 0.6, 0.1])[0]
        
        # Default distribution: mostly normal
        return random.choices(["high", "normal", "low"], weights=[0.1, 0.8, 0.1])[0]
    
    def _generate_flag_status(self, email_date: datetime, category: str) -> Dict[str, Any]:
        """
        Generate follow-up flag status for the email.
        
        Returns:
            Dictionary with flag information:
            - flagged: bool - whether email is flagged
            - flag_type: str - 'followUp', 'reply', 'review', etc.
            - due_date: datetime or None - when follow-up is due
            - completed: bool - whether flag is completed
        """
        # Only flag some emails (about 15%)
        if random.random() > 0.15:
            return {"flagged": False}
        
        # Categories more likely to be flagged
        flag_weights = {
            "interdepartmental": 0.25,
            "organisational": 0.20,
            "external_business": 0.20,
            "security": 0.15,
            "attachments": 0.15,
            "links": 0.10,
            "newsletters": 0.05,
            "spam": 0.02,
        }
        
        # Check if this category should be flagged
        if random.random() > flag_weights.get(category, 0.10):
            return {"flagged": False}
        
        # Flag types
        flag_types = ["followUp", "reply", "review", "forYourInformation", "noResponseNecessary"]
        flag_type = random.choices(
            flag_types,
            weights=[0.4, 0.25, 0.2, 0.1, 0.05]
        )[0]
        
        # Due date (1-14 days from email date)
        due_days = random.randint(1, 14)
        due_date = email_date + timedelta(days=due_days)
        
        # Older emails might have completed flags
        now = datetime.now()
        if email_date.tzinfo:
            now = now.replace(tzinfo=email_date.tzinfo)
        
        days_old = (now - email_date).days
        completed = days_old > due_days and random.random() > 0.3
        
        return {
            "flagged": True,
            "flag_type": flag_type,
            "due_date": due_date,
            "completed": completed,
        }
    
    def _generate_categories(self, category: str, recipient: Dict) -> List[str]:
        """
        Generate Outlook categories (color tags) for the email.
        
        Returns:
            List of category names (Outlook color categories)
        """
        # Only categorize some emails (about 20%)
        if random.random() > 0.20:
            return []
        
        # Map email categories to Outlook color categories
        category_mapping = {
            "interdepartmental": ["Blue Category", "Green Category"],
            "organisational": ["Purple Category", "Red Category"],
            "security": ["Red Category", "Orange Category"],
            "external_business": ["Yellow Category", "Green Category"],
            "attachments": ["Blue Category", "Yellow Category"],
            "links": ["Green Category"],
            "newsletters": ["Purple Category"],
            "spam": [],  # Don't categorize spam
        }
        
        # Department-based categories
        department = recipient.get("department", "")
        dept_categories = {
            "Finance Department": ["Yellow Category"],
            "Human Resources": ["Purple Category"],
            "IT Department": ["Blue Category"],
            "Marketing Department": ["Orange Category"],
            "Sales Department": ["Green Category"],
            "Executive Leadership": ["Red Category"],
            "Legal & Compliance": ["Red Category"],
            "Claims Department": ["Yellow Category"],
        }
        
        available = category_mapping.get(category, ["Blue Category"])
        dept_cats = dept_categories.get(department, [])
        
        # Combine and pick 1-2 categories
        all_cats = list(set(available + dept_cats))
        if not all_cats:
            return []
        
        num_cats = random.choices([1, 2], weights=[0.8, 0.2])[0]
        return random.sample(all_cats, min(num_cats, len(all_cats)))
    
    def _generate_threading_info(
        self,
        category: str,
        template: Dict,
        sender: Dict,
        recipient: Dict
    ) -> Dict[str, Any]:
        """
        Generate email threading information (reply/forward chains).
        
        Returns:
            Dictionary with threading info:
            - is_reply: bool - whether this is a reply
            - is_forward: bool - whether this is a forward
            - is_reply_all: bool - whether this is a reply-all
            - thread_id: str - conversation thread ID
            - in_reply_to: str - message ID being replied to
            - references: list - chain of message IDs
            - original_sender: dict - original sender for forwards
        """
        # Check if template supports threading - default to True for most categories
        # to ensure we get a good mix of threaded emails
        supports_threading = template.get("supports_threading", True)
        
        # Categories that should never have threading
        no_threading_categories = ["spam", "newsletters"]
        if category in no_threading_categories:
            supports_threading = False
        
        if not supports_threading:
            return {"is_reply": False, "is_forward": False, "is_reply_all": False}
        
        # Threading probability by category - increased for more realistic mailbox
        # Real mailboxes have many reply chains and forwards
        thread_weights = {
            "interdepartmental": 0.55,  # Increased from 0.45 - lots of internal back-and-forth
            "external_business": 0.50,  # Increased from 0.35 - client conversations
            "organisational": 0.35,     # Increased from 0.20 - HR/policy discussions
            "attachments": 0.40,        # Increased from 0.25 - document reviews
            "links": 0.25,              # Increased from 0.15 - shared link discussions
            "security": 0.20,           # Increased from 0.10 - security follow-ups
            "newsletters": 0.0,         # Keep at 0 - newsletters aren't replied to
            "spam": 0.0,                # Keep at 0 - spam isn't replied to
        }
        
        if random.random() > thread_weights.get(category, 0.30):
            return {"is_reply": False, "is_forward": False, "is_reply_all": False}
        
        # Determine thread type
        thread_type = random.choices(
            ["reply", "reply_all", "forward"],
            weights=[0.5, 0.25, 0.25]
        )[0]
        
        # Generate thread ID
        import hashlib
        thread_seed = f"{sender.get('email', '')}-{recipient.get('email', '')}-{random.randint(1000, 9999)}"
        thread_id = hashlib.md5(thread_seed.encode()).hexdigest()[:16]
        
        # Generate message ID for in-reply-to
        original_msg_id = f"<{thread_id}.{random.randint(1, 100)}@{sender.get('email', 'example.com').split('@')[-1]}>"
        
        # Build references chain (1-3 previous messages)
        num_refs = random.randint(1, 3)
        references = [
            f"<{thread_id}.{i}@{sender.get('email', 'example.com').split('@')[-1]}>"
            for i in range(num_refs)
        ]
        
        result = {
            "is_reply": thread_type == "reply",
            "is_forward": thread_type == "forward",
            "is_reply_all": thread_type == "reply_all",
            "thread_id": thread_id,
            "in_reply_to": original_msg_id,
            "references": references,
        }
        
        # For forwards, include original sender info
        if thread_type == "forward":
            # Generate a plausible original sender
            original_senders = [
                {"name": "John Smith", "email": "john.smith@external.com"},
                {"name": "Sarah Johnson", "email": "sarah.j@partner.org"},
                {"name": "Mike Wilson", "email": "m.wilson@vendor.net"},
                {"name": "Emily Brown", "email": "emily.brown@client.com"},
                {"name": "David Lee", "email": "d.lee@supplier.io"},
            ]
            result["original_sender"] = random.choice(original_senders)
        
        return result
    
    def _get_site_for_department(self, department: str) -> str:
        """Get SharePoint site name for a department."""
        dept_config = self.departments.get(department, {})
        site_name = dept_config.get("sharepoint_site")
        
        if site_name and site_name in self.site_lookup:
            return site_name
        
        # Try to find a matching site
        dept_lower = department.lower().replace(" ", "-").replace("&", "")
        for site_name in self.site_lookup:
            if dept_lower in site_name.lower():
                return site_name
        
        # Return first available site or default
        if self.site_lookup:
            return list(self.site_lookup.keys())[0]
        return "company-intranet"
    
    # ==========================================================================
    # Content Generation Methods
    # ==========================================================================
    
    def _generate_phone(self) -> str:
        """Generate a realistic phone number with international variety."""
        formats = [
            f"+1 ({random.randint(200, 999)}) {random.randint(200, 999)}-{random.randint(1000, 9999)}",
            f"+44 {random.randint(20, 79)} {random.randint(1000, 9999)} {random.randint(1000, 9999)}",
            f"+1-{random.randint(200, 999)}-{random.randint(200, 999)}-{random.randint(1000, 9999)}",
        ]
        return random.choice(formats)
    
    def _generate_document_name(self, department: str) -> str:
        """Generate a realistic document name using variations module."""
        # Use the variations module for more diverse document names
        return variations.get_random_document_name(department)
    
    def _generate_meeting_topic(self, department: str) -> str:
        """Generate a meeting topic using variations module."""
        # Use the variations module for more diverse meeting topics
        return variations.get_random_meeting_topic(department)
    
    def _generate_proposed_times(self) -> str:
        """Generate proposed meeting times."""
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        times = ["9:00 AM", "10:00 AM", "10:30 AM", "11:00 AM", "1:00 PM", "2:00 PM", "2:30 PM", "3:00 PM", "3:30 PM", "4:00 PM"]
        
        day1 = random.choice(days)
        day2 = random.choice([d for d in days if d != day1])
        time1 = random.choice(times)
        time2 = random.choice(times)
        
        # Add more variation in format
        formats = [
            f"{day1} at {time1} or {day2} at {time2}",
            f"Either {day1} ({time1}) or {day2} ({time2})",
            f"{day1} @ {time1}, alternatively {day2} @ {time2}",
            f"Option 1: {day1} {time1}\nOption 2: {day2} {time2}",
        ]
        
        return random.choice(formats)
    
    def _generate_temp_password(self) -> str:
        """Generate a realistic temporary password.
        
        Creates passwords that look like real temporary passwords:
        - Mix of uppercase, lowercase, numbers, and special characters
        - 12-16 characters long
        - Easy to read (avoids ambiguous characters like 0/O, 1/l)
        """
        # Character sets (avoiding ambiguous characters)
        uppercase = "ABCDEFGHJKLMNPQRSTUVWXYZ"  # No I, O
        lowercase = "abcdefghjkmnpqrstuvwxyz"   # No i, l, o
        digits = "23456789"                      # No 0, 1
        special = "!@#$%&*"
        
        # Generate password with guaranteed character types
        password_chars = [
            secrets.choice(uppercase),
            secrets.choice(uppercase),
            secrets.choice(lowercase),
            secrets.choice(lowercase),
            secrets.choice(lowercase),
            secrets.choice(digits),
            secrets.choice(digits),
            secrets.choice(special),
        ]
        
        # Add more random characters to reach 12-16 length
        all_chars = uppercase + lowercase + digits + special
        remaining_length = random.randint(4, 8)
        password_chars.extend(secrets.choice(all_chars) for _ in range(remaining_length))
        
        # Shuffle to randomize position
        random.shuffle(password_chars)
        
        return "".join(password_chars)
    
    def _generate_project_name(self, department: str) -> str:
        """Generate a project name using variations module."""
        return variations.get_random_project_name()
    
    def _generate_announcement_title(self) -> str:
        """Generate an announcement title with more variety."""
        titles = [
            "New Office Guidelines",
            "Upcoming Company Event",
            "System Maintenance Notice",
            "Policy Update",
            "Team Restructuring",
            "New Benefits Program",
            "Holiday Schedule",
            "Training Opportunity",
            "Sustainability Initiative",
            "Workplace Improvements",
            "Important: Organizational Changes",
            "New Partnership Announcement",
            "Office Relocation Update",
            "Technology Upgrade Notice",
            "Employee Recognition Program",
            "Health & Safety Update",
            "Remote Work Policy Changes",
            "Quarterly Business Update",
            "New Product Launch",
            "Customer Success Story",
        ]
        return random.choice(titles)
    
    def _generate_topic(self, department: str) -> str:
        """Generate a general topic with department-specific jargon."""
        # Use variations module for department-specific content
        base_topics = [
            f"Q{variations.get_current_quarter()} Planning",
            "Process Improvement",
            "Resource Allocation",
            "Timeline Review",
            "Budget Discussion",
            "Team Collaboration",
            "Project Requirements",
            "Strategy Alignment",
            variations.get_department_priority(department),
            f"{variations.get_department_jargon(department)} discussion",
        ]
        return random.choice(base_topics)
    
    def _generate_company_news(self) -> str:
        """Generate company news content with extensive variations."""
        quarter = variations.get_current_quarter()
        year = variations.get_current_fiscal_year()
        
        news_templates = [
            # Product/Service launches
            "<p>We're excited to announce the launch of our new customer portal, which will streamline how we interact with our clients.</p>",
            "<p>Our new mobile app is now available for download! This marks a significant milestone in our digital transformation journey.</p>",
            "<p>We've just released version 2.0 of our flagship product, featuring enhanced security and improved user experience.</p>",
            "<p>The beta testing phase for our new platform has concluded successfully, with overwhelmingly positive feedback.</p>",
            
            # Team achievements
            f"<p>Congratulations to the Sales team for exceeding Q{quarter} targets by {random.randint(10, 25)}%! This achievement reflects the hard work and dedication of everyone involved.</p>",
            f"<p>Our Marketing team's latest campaign generated {random.randint(40, 80)}% more leads than projected. Outstanding work!</p>",
            f"<p>The IT department successfully completed the cloud migration ahead of schedule, with zero downtime.</p>",
            f"<p>Our customer support team achieved a {random.randint(95, 99)}% satisfaction rating this quarter!</p>",
            
            # Sustainability and CSR
            f"<p>Our sustainability initiative has reached a major milestone - we've reduced our carbon footprint by {random.randint(15, 30)}% this year.</p>",
            f"<p>Through our community outreach program, employees have volunteered over {random.randint(500, 1500)} hours this quarter.</p>",
            "<p>We're proud to announce our commitment to becoming carbon neutral by 2030.</p>",
            "<p>Our recycling program has diverted over 10 tons of waste from landfills this year.</p>",
            
            # Company growth
            f"<p>We're pleased to welcome {random.randint(10, 30)} new team members who joined us this month across various departments.</p>",
            f"<p>Our company has grown by {random.randint(15, 40)}% this year, reflecting strong market demand for our services.</p>",
            "<p>We've opened a new office in the downtown area to accommodate our growing team.</p>",
            f"<p>Customer base has expanded to over {random.randint(5, 20)},000 active users worldwide.</p>",
            
            # Awards and recognition
            f"<p>The company has been recognized as one of the top employers in our industry for the {random.choice(['second', 'third', 'fourth', 'fifth'])} consecutive year.</p>",
            "<p>We've been awarded the Industry Excellence Award for innovation in customer service.</p>",
            f"<p>Our CEO has been named among the Top {random.randint(25, 100)} Business Leaders of {year}.</p>",
            "<p>We've earned the Great Place to Work certification for our commitment to employee wellbeing.</p>",
            
            # Partnerships and collaborations
            "<p>We're thrilled to announce a strategic partnership with a leading technology provider to enhance our offerings.</p>",
            "<p>Our collaboration with the local university has resulted in three new research initiatives.</p>",
            "<p>We've joined forces with industry leaders to establish new standards for data security.</p>",
        ]
        return random.choice(news_templates)
    
    def _generate_upcoming_events(self) -> str:
        """Generate upcoming events content with extensive variations."""
        now = datetime.now()
        
        # Event types with times
        event_types = [
            ("Town Hall Meeting", ["2:00 PM", "3:00 PM", "10:00 AM"]),
            ("All-Hands Meeting", ["9:00 AM", "2:00 PM", "4:00 PM"]),
            ("Team Building Event", ["1:00 PM", "10:00 AM", "All Day"]),
            ("Training Session", ["10:00 AM", "2:00 PM", "9:30 AM"]),
            ("Quarterly Review", ["10:00 AM", "2:00 PM", "3:00 PM"]),
            ("Lunch & Learn", ["12:00 PM", "12:30 PM", "1:00 PM"]),
            ("Wellness Workshop", ["11:00 AM", "3:00 PM", "4:00 PM"]),
            ("Leadership Forum", ["9:00 AM", "2:00 PM", "10:00 AM"]),
            ("Innovation Day", ["9:00 AM", "All Day", "10:00 AM"]),
            ("Customer Appreciation Event", ["5:00 PM", "6:00 PM", "4:00 PM"]),
            ("Annual Company Picnic", ["11:00 AM", "12:00 PM", "All Day"]),
            ("Holiday Party", ["5:00 PM", "6:00 PM", "7:00 PM"]),
            ("Professional Development Day", ["9:00 AM", "All Day", "10:00 AM"]),
            ("Diversity & Inclusion Workshop", ["2:00 PM", "10:00 AM", "3:00 PM"]),
            ("New Employee Orientation", ["9:00 AM", "10:00 AM", "1:00 PM"]),
            ("Product Demo Day", ["2:00 PM", "3:00 PM", "10:00 AM"]),
            ("Strategy Planning Session", ["9:00 AM", "10:00 AM", "2:00 PM"]),
            ("Team Retrospective", ["3:00 PM", "4:00 PM", "2:00 PM"]),
        ]
        
        # Generate 1-3 events
        num_events = random.randint(1, 3)
        selected_events = random.sample(event_types, min(num_events, len(event_types)))
        
        events_html = []
        for event_name, times in selected_events:
            days_ahead = random.randint(3, 30)
            event_date = now + timedelta(days=days_ahead)
            time = random.choice(times)
            
            # Vary the format
            formats = [
                f"<p>📅 <strong>{event_name}</strong> - {event_date.strftime('%B %d')} at {time}</p>",
                f"<p>📅 <strong>{event_name}</strong> - {event_date.strftime('%A, %B %d')} @ {time}</p>",
                f"<p>🗓️ <strong>{event_name}</strong>: {event_date.strftime('%B %d, %Y')} ({time})</p>",
                f"<p>📌 <strong>{event_name}</strong> | {event_date.strftime('%b %d')} | {time}</p>",
            ]
            events_html.append(random.choice(formats))
        
        return "".join(events_html)
    
    def _generate_employee_spotlight(self) -> str:
        """Generate employee spotlight content with extensive variations."""
        first_names = [
            "Sarah", "Michael", "Emily", "James", "Jessica", "David", "Amanda", "Christopher",
            "Ashley", "Matthew", "Jennifer", "Daniel", "Elizabeth", "Andrew", "Stephanie",
            "Joshua", "Nicole", "Ryan", "Megan", "Brandon", "Rachel", "Kevin", "Lauren",
            "Priya", "Wei", "Carlos", "Fatima", "Yuki", "Olga", "Ahmed", "Maria",
        ]
        last_names = [
            "Johnson", "Chen", "Davis", "Wilson", "Martinez", "Anderson", "Taylor", "Thomas",
            "Garcia", "Brown", "Miller", "Jones", "Williams", "Smith", "Lee", "Patel",
            "Kim", "Nguyen", "Rodriguez", "Singh", "O'Brien", "Murphy", "Schmidt", "Müller",
        ]
        
        departments = [
            "Marketing", "IT", "HR", "Sales", "Finance", "Operations", "Engineering",
            "Customer Success", "Product", "Legal", "Research", "Design", "Support",
        ]
        
        achievements = [
            "outstanding work on the {project} project",
            "completing their {certification} certification",
            "organizing the successful {event}",
            "closing our biggest deal of the quarter",
            "leading the {initiative} initiative",
            "exceptional customer service ratings",
            "mentoring {count} new team members",
            "innovative solution to {problem}",
            "going above and beyond during {situation}",
            "receiving the {award} award",
            "celebrating {years} years with the company",
            "successful launch of {product}",
            "outstanding leadership during {project}",
            "achieving {metric}% improvement in {area}",
        ]
        
        projects = ["brand refresh", "system migration", "digital transformation", "customer portal", "mobile app", "data analytics"]
        certifications = ["AWS", "Azure", "PMP", "Scrum Master", "Six Sigma", "CISSP", "Google Cloud"]
        events = ["wellness week", "team building day", "charity drive", "hackathon", "training program"]
        initiatives = ["sustainability", "diversity", "innovation", "process improvement", "customer experience"]
        problems = ["the integration challenge", "the scalability issue", "the performance bottleneck"]
        situations = ["the product launch", "the system outage", "the busy season", "the transition period"]
        awards = ["Employee of the Month", "Innovation", "Customer Champion", "Team Player", "Rising Star"]
        products = ["the new dashboard", "the mobile app", "the API platform", "the analytics tool"]
        
        name = f"{random.choice(first_names)} {random.choice(last_names)}"
        dept = random.choice(departments)
        achievement = random.choice(achievements)
        
        # Fill in placeholders
        achievement = achievement.replace("{project}", random.choice(projects))
        achievement = achievement.replace("{certification}", random.choice(certifications))
        achievement = achievement.replace("{event}", random.choice(events))
        achievement = achievement.replace("{initiative}", random.choice(initiatives))
        achievement = achievement.replace("{problem}", random.choice(problems))
        achievement = achievement.replace("{situation}", random.choice(situations))
        achievement = achievement.replace("{award}", random.choice(awards))
        achievement = achievement.replace("{product}", random.choice(products))
        achievement = achievement.replace("{count}", str(random.randint(2, 5)))
        achievement = achievement.replace("{years}", str(random.randint(5, 15)))
        achievement = achievement.replace("{metric}", str(random.randint(15, 50)))
        achievement = achievement.replace("{area}", random.choice(["efficiency", "productivity", "satisfaction", "performance"]))
        
        templates = [
            f"<p>This week we're celebrating <strong>{name}</strong> from the {dept} team for {achievement}!</p>",
            f"<p>Congratulations to <strong>{name}</strong> from {dept} for {achievement}!</p>",
            f"<p>A big thank you to <strong>{name}</strong> from {dept} for {achievement}!</p>",
            f"<p>Kudos to <strong>{name}</strong> ({dept}) for {achievement}!</p>",
            f"<p>🌟 <strong>Employee Spotlight:</strong> {name} from {dept} - recognized for {achievement}!</p>",
            f"<p>👏 Shoutout to <strong>{name}</strong> in {dept} for {achievement}!</p>",
        ]
        
        return random.choice(templates)
    
    def _generate_fun_fact(self) -> str:
        """Generate a fun fact with extensive variations."""
        year = variations.get_current_fiscal_year()
        
        facts = [
            # Company stats
            f"Our company has employees in {random.randint(8, 25)} different countries!",
            f"We've served over {random.randint(5, 50)},000 customers this year.",
            f"The average tenure of our employees is {random.choice(['3.5', '4', '4.5', '5', '5.5'])} years.",
            f"Our team has collectively volunteered over {random.randint(300, 1500)} hours this quarter.",
            f"We've reduced paper usage by {random.randint(30, 60)}% since going digital.",
            
            # Fun workplace facts
            f"Our office coffee machines serve approximately {random.randint(500, 2000)} cups per week!",
            f"The most popular lunch spot among employees is the {random.choice(['Italian place', 'sushi restaurant', 'food truck', 'salad bar'])} down the street.",
            f"Our company Slack has over {random.randint(50, 200)} custom emojis created by employees.",
            f"The company book club has read {random.randint(12, 36)} books since it started.",
            f"Our pet-friendly policy means we have {random.randint(5, 20)} regular office dogs!",
            
            # Achievement facts
            f"We've shipped {random.randint(50, 200)} product updates this year.",
            f"Our support team has resolved over {random.randint(10, 50)},000 tickets with a {random.randint(95, 99)}% satisfaction rate.",
            f"Employees have completed {random.randint(500, 2000)} training courses this year.",
            f"Our innovation program has generated {random.randint(50, 200)} new ideas this quarter.",
            
            # Historical facts
            f"The company was founded in {year - random.randint(5, 25)} in a small garage.",
            f"Our first customer is still with us after {random.randint(5, 15)} years!",
            f"The company name was chosen from {random.randint(50, 200)} suggestions.",
            
            # Sustainability facts
            f"Our solar panels generate {random.randint(20, 50)}% of our office energy needs.",
            f"We've planted {random.randint(500, 5000)} trees through our environmental program.",
            f"Our recycling program has diverted {random.randint(5, 20)} tons of waste from landfills.",
        ]
        
        return random.choice(facts)
    
    def _generate_department_updates(self) -> str:
        """Generate department updates content with extensive variations."""
        quarter = variations.get_current_quarter()
        
        department_updates = {
            "IT": [
                "New security protocols will be implemented next week.",
                "System maintenance scheduled for this weekend.",
                "New collaboration tools rolling out next month.",
                "Cybersecurity training is now mandatory for all staff.",
                "Help desk response times improved by 25%.",
                "Cloud migration Phase 2 begins next quarter.",
            ],
            "HR": [
                "Open enrollment for benefits begins next month.",
                "New wellness program launching soon.",
                "Performance review cycle starts next week.",
                "Updated PTO policy now in effect.",
                "Leadership development program accepting applications.",
                "Employee satisfaction survey results are in!",
            ],
            "Marketing": [
                f"New campaign launching next week.",
                "Brand guidelines have been updated.",
                "Social media engagement up 35% this month.",
                "Website redesign project on track.",
                "Customer testimonial video series in production.",
                "Trade show preparations underway.",
            ],
            "Sales": [
                f"Q{quarter} targets exceeded by {random.randint(8, 20)}%.",
                "New CRM features now available.",
                "Sales training workshop next Tuesday.",
                "Territory realignment complete.",
                "New pricing structure effective next month.",
                "Customer retention rate at all-time high.",
            ],
            "Finance": [
                f"Budget planning for FY{variations.get_current_fiscal_year() + 1} has begun.",
                "Expense report deadline is end of month.",
                "New travel policy now in effect.",
                "Quarterly financial review scheduled.",
                "Vendor payment terms updated.",
                "Cost optimization initiative showing results.",
            ],
            "Operations": [
                "New efficiency measures showing positive results.",
                "Supply chain improvements reducing lead times.",
                "Quality metrics at record levels.",
                "Process automation saving 20 hours weekly.",
                "Vendor scorecard reviews complete.",
                "Inventory management system upgrade complete.",
            ],
        }
        
        # Select 2-3 random departments
        selected_depts = random.sample(list(department_updates.keys()), random.randint(2, 3))
        
        updates_html = []
        for dept in selected_depts:
            update = random.choice(department_updates[dept])
            updates_html.append(f"<p><strong>{dept}:</strong> {update}</p>")
        
        return "".join(updates_html)
    
    def _generate_articles(self) -> str:
        """Generate newsletter articles with extensive variations and realistic links."""
        year = variations.get_current_fiscal_year()
        
        # Realistic external article URLs from reputable business/tech sources
        article_urls = [
            "https://hbr.org/article/",
            "https://www.forbes.com/sites/",
            "https://www.mckinsey.com/insights/",
            "https://www.gartner.com/en/articles/",
            "https://www.forrester.com/report/",
            "https://www.techcrunch.com/",
            "https://www.wired.com/story/",
            "https://www.fastcompany.com/",
            "https://www.inc.com/",
            "https://www.entrepreneur.com/article/",
            "https://www.businessinsider.com/",
            "https://www.cnbc.com/",
            "https://www.reuters.com/business/",
            "https://www.bloomberg.com/news/",
            "https://www.wsj.com/articles/",
            "https://www.ft.com/content/",
            "https://www.economist.com/",
            "https://sloanreview.mit.edu/article/",
            "https://www.strategy-business.com/article/",
            "https://www.bain.com/insights/",
        ]
        
        article_pool = [
            # Industry trends
            {
                "title": f"Industry Trends: What to Watch in {year}",
                "time": f"{random.randint(4, 8)} min read",
                "summary": "The latest developments shaping our industry and what they mean for businesses like ours.",
                "url": f"{random.choice(article_urls)}industry-trends-{year}-{random.randint(1000, 9999)}"
            },
            {
                "title": "Market Analysis: Emerging Opportunities",
                "time": f"{random.randint(5, 10)} min read",
                "summary": "A deep dive into new market segments and growth potential for the coming year.",
                "url": f"{random.choice(article_urls)}market-analysis-opportunities-{random.randint(1000, 9999)}"
            },
            {
                "title": f"Economic Outlook for {year}",
                "time": f"{random.randint(6, 12)} min read",
                "summary": "Expert predictions and analysis of economic factors affecting our business.",
                "url": f"{random.choice(article_urls)}economic-outlook-{year}-{random.randint(1000, 9999)}"
            },
            # Technology
            {
                "title": "Technology Spotlight: AI in the Workplace",
                "time": f"{random.randint(3, 6)} min read",
                "summary": "How artificial intelligence is transforming how we work.",
                "url": f"{random.choice(article_urls)}ai-workplace-transformation-{random.randint(1000, 9999)}"
            },
            {
                "title": "Digital Transformation: Success Stories",
                "time": f"{random.randint(4, 7)} min read",
                "summary": "Real-world examples of companies thriving through digital innovation.",
                "url": f"{random.choice(article_urls)}digital-transformation-success-{random.randint(1000, 9999)}"
            },
            {
                "title": "Cybersecurity Best Practices for Teams",
                "time": f"{random.randint(3, 5)} min read",
                "summary": "Essential security tips every employee should know.",
                "url": f"{random.choice(article_urls)}cybersecurity-best-practices-{random.randint(1000, 9999)}"
            },
            {
                "title": "The Future of Cloud Computing",
                "time": f"{random.randint(5, 8)} min read",
                "summary": "Exploring the next generation of cloud technologies and their business impact.",
                "url": f"{random.choice(article_urls)}future-cloud-computing-{random.randint(1000, 9999)}"
            },
            # Workplace
            {
                "title": "Best Practices for Remote Collaboration",
                "time": f"{random.randint(3, 5)} min read",
                "summary": "Tips and tools for effective teamwork in a distributed environment.",
                "url": f"{random.choice(article_urls)}remote-collaboration-tips-{random.randint(1000, 9999)}"
            },
            {
                "title": "Building a Culture of Innovation",
                "time": f"{random.randint(4, 6)} min read",
                "summary": "How leading companies foster creativity and continuous improvement.",
                "url": f"{random.choice(article_urls)}innovation-culture-building-{random.randint(1000, 9999)}"
            },
            {
                "title": "Work-Life Balance in the Modern Era",
                "time": f"{random.randint(3, 5)} min read",
                "summary": "Strategies for maintaining wellbeing while staying productive.",
                "url": f"{random.choice(article_urls)}work-life-balance-strategies-{random.randint(1000, 9999)}"
            },
            # Leadership
            {
                "title": "Leadership Lessons from Top CEOs",
                "time": f"{random.randint(5, 8)} min read",
                "summary": "Insights and advice from industry leaders on effective management.",
                "url": f"{random.choice(article_urls)}ceo-leadership-lessons-{random.randint(1000, 9999)}"
            },
            {
                "title": "Developing Your Leadership Style",
                "time": f"{random.randint(4, 7)} min read",
                "summary": "A guide to understanding and improving your leadership approach.",
                "url": f"{random.choice(article_urls)}leadership-style-development-{random.randint(1000, 9999)}"
            },
            # Professional development
            {
                "title": f"Skills That Will Matter in {year}",
                "time": f"{random.randint(4, 6)} min read",
                "summary": "The competencies employers are looking for in the evolving job market.",
                "url": f"{random.choice(article_urls)}skills-matter-{year}-{random.randint(1000, 9999)}"
            },
            {
                "title": "Continuous Learning: A Career Imperative",
                "time": f"{random.randint(3, 5)} min read",
                "summary": "Why lifelong learning is essential for professional growth.",
                "url": f"{random.choice(article_urls)}continuous-learning-career-{random.randint(1000, 9999)}"
            },
        ]
        
        # Select 2-4 random articles
        selected_articles = random.sample(article_pool, random.randint(2, 4))
        
        articles_html = []
        for article in selected_articles:
            articles_html.append(f"""
        <div class="article">
            <div class="article-title">{article['title']}</div>
            <div class="article-meta">{article['time']}</div>
            <p class="article-summary">{article['summary']}</p>
            <a href="{article['url']}" class="read-more">Read More →</a>
        </div>""")
        
        return "".join(articles_html)
    
    def _generate_activity_items(self) -> str:
        """Generate SharePoint activity items with extensive variations."""
        quarter = variations.get_current_quarter()
        year = variations.get_current_fiscal_year()
        
        first_names = ["John", "Sarah", "Michael", "Emily", "David", "Jessica", "Chris", "Amanda", "James", "Rachel",
                       "Kevin", "Lauren", "Brian", "Michelle", "Andrew", "Nicole", "Daniel", "Ashley", "Matthew", "Jennifer"]
        last_names = ["Smith", "Johnson", "Chen", "Davis", "Wilson", "Martinez", "Brown", "Taylor", "Lee", "Garcia",
                      "Anderson", "Thomas", "Jackson", "White", "Harris", "Martin", "Thompson", "Moore", "Young", "Allen"]
        
        document_names = [
            f"Q{quarter} Report.xlsx", "Project Plan.docx", "Budget Analysis.xlsx",
            "Meeting Notes.docx", "Presentation.pptx", "Requirements.docx",
            f"FY{year} Strategy.pptx", "Team Roster.xlsx", "Process Guide.pdf",
            "Training Materials.pptx", "Policy Update.docx", "Vendor Contract.pdf",
            "Marketing Brief.docx", "Sales Forecast.xlsx", "Technical Spec.docx",
            "Status Report.docx", "Risk Assessment.xlsx", "Timeline.xlsx",
            "Proposal Draft.docx", "Client Feedback.docx", "Action Items.docx",
        ]
        
        folder_names = [
            f"Archive {year}", f"Q{quarter} Documents", "Shared Resources",
            "Team Files", "Project Assets", "Templates", "Reports",
            "Meeting Materials", "Training", "Policies", "Contracts",
            "Marketing Assets", "Sales Materials", "HR Documents",
        ]
        
        time_references = [
            "Just now", "5 minutes ago", "15 minutes ago", "30 minutes ago",
            "1 hour ago", "2 hours ago", "3 hours ago", "Earlier today",
            "Yesterday", "2 days ago", "3 days ago", "Last week",
        ]
        
        activity_types = [
            ("📄", "New document uploaded", document_names),
            ("✏️", "Document edited", document_names),
            ("📁", "New folder created", folder_names),
            ("🔗", "Document shared", document_names),
            ("💬", "Comment added to", document_names),
            ("📥", "Document downloaded", document_names),
            ("🗑️", "Document moved to archive", document_names),
            ("📋", "Document copied", document_names),
            ("👁️", "Document viewed", document_names),
            ("✅", "Document approved", document_names),
        ]
        
        # Generate 3-5 activity items
        num_activities = random.randint(3, 5)
        activities_html = []
        
        for _ in range(num_activities):
            icon, action, items = random.choice(activity_types)
            item = random.choice(items)
            name = f"{random.choice(first_names)} {random.choice(last_names)}"
            time_ref = random.choice(time_references)
            
            activities_html.append(f"""
        <div class="activity-item">
            <span class="activity-icon">{icon}</span>
            <span class="activity-text">{action}: {item}</span>
            <div class="activity-meta">{time_ref} by {name}</div>
        </div>""")
        
        return "".join(activities_html)
    
    def _generate_key_points(self, department: str) -> str:
        """Generate key points for reports with extensive variations."""
        points = {
            "Finance Department": [
                f"<li>Revenue increased by {random.randint(5, 15)}% compared to last quarter</li>",
                f"<li>Operating costs reduced by {random.randint(3, 10)}%</li>",
                "<li>Cash flow remains strong with healthy reserves</li>",
                f"<li>Budget utilization at {random.randint(85, 98)}%</li>",
                f"<li>Accounts receivable improved by {random.randint(5, 15)}%</li>",
                "<li>Audit findings addressed and resolved</li>",
                f"<li>Cost savings of ${random.randint(50, 500)}K achieved</li>",
            ],
            "Human Resources": [
                f"<li>Employee satisfaction score: {random.choice(['4.0', '4.1', '4.2', '4.3', '4.4', '4.5'])}/5</li>",
                f"<li>Turnover rate decreased by {random.randint(2, 8)}%</li>",
                f"<li>{random.randint(10, 30)} new hires onboarded this month</li>",
                f"<li>Training completion rate: {random.randint(90, 99)}%</li>",
                f"<li>Time-to-hire reduced by {random.randint(10, 25)}%</li>",
                "<li>Benefits enrollment completed successfully</li>",
                f"<li>Employee engagement up {random.randint(5, 15)}%</li>",
            ],
            "IT Department": [
                f"<li>System uptime: {random.choice(['99.9', '99.95', '99.99'])}%</li>",
                f"<li>Security incidents: {random.randint(0, 2)}</li>",
                f"<li>Help desk tickets resolved: {random.randint(92, 99)}%</li>",
                f"<li>Average response time: {random.randint(15, 45)} minutes</li>",
                f"<li>Cloud migration: {random.randint(60, 95)}% complete</li>",
                "<li>Zero critical vulnerabilities detected</li>",
                f"<li>Infrastructure costs reduced by {random.randint(10, 25)}%</li>",
            ],
            "Marketing Department": [
                f"<li>Website traffic increased by {random.randint(15, 40)}%</li>",
                f"<li>Social media engagement up {random.randint(25, 60)}%</li>",
                "<li>Lead generation exceeded targets</li>",
                f"<li>Email open rate: {random.randint(25, 45)}%</li>",
                f"<li>Campaign ROI: {random.randint(150, 300)}%</li>",
                f"<li>Brand awareness increased by {random.randint(10, 30)}%</li>",
                f"<li>Content production up {random.randint(20, 50)}%</li>",
            ],
            "Sales Department": [
                f"<li>Quota attainment: {random.randint(95, 120)}%</li>",
                f"<li>Pipeline value increased by {random.randint(15, 40)}%</li>",
                f"<li>Win rate improved to {random.randint(25, 45)}%</li>",
                f"<li>Average deal size up {random.randint(10, 30)}%</li>",
                f"<li>Customer retention: {random.randint(90, 98)}%</li>",
                f"<li>New accounts: {random.randint(15, 50)} this quarter</li>",
                f"<li>Sales cycle reduced by {random.randint(10, 25)}%</li>",
            ],
            "Operations Department": [
                f"<li>Process efficiency improved by {random.randint(10, 25)}%</li>",
                f"<li>Quality metrics at {random.randint(95, 99)}%</li>",
                f"<li>On-time delivery: {random.randint(92, 99)}%</li>",
                f"<li>Inventory accuracy: {random.choice(['97.5', '98', '98.5', '99', '99.5'])}%</li>",
                f"<li>Cost per unit reduced by {random.randint(5, 15)}%</li>",
                "<li>Zero safety incidents this period</li>",
                f"<li>Vendor performance score: {random.randint(85, 98)}%</li>",
            ],
            "Claims Department": [
                f"<li>Claims processed: {random.randint(150, 500)} this period</li>",
                f"<li>Average processing time reduced by {random.randint(10, 25)}%</li>",
                f"<li>Settlement accuracy: {random.choice(['97.5', '98', '98.5', '99', '99.5'])}%</li>",
                f"<li>Customer satisfaction: {random.choice(['4.2', '4.3', '4.4', '4.5', '4.6'])}/5</li>",
                f"<li>Fraud detection rate improved by {random.randint(15, 35)}%</li>",
                f"<li>Reserve accuracy: {random.randint(95, 99)}%</li>",
                f"<li>Subrogation recovery: ${random.randint(50, 200)}K this quarter</li>",
            ],
        }
        
        dept_points = points.get(department, [
            f"<li>Key metric 1: {random.choice(['On track', 'Exceeding targets', 'Meeting goals'])}</li>",
            f"<li>Key metric 2: Improved by {random.randint(5, 20)}%</li>",
            f"<li>Key metric 3: {random.choice(['Meeting targets', 'Above expectations', 'Strong performance'])}</li>",
        ])
        
        # Select 3-4 random points from the department's list
        selected_points = random.sample(dept_points, min(random.randint(3, 4), len(dept_points)))
        return "".join(selected_points)
    
    def _generate_executive_summary(self) -> str:
        """Generate executive summary with extensive variations."""
        quarter = variations.get_current_quarter()
        
        summaries = [
            f"<p><strong>Executive Summary:</strong> Overall performance remains strong with key metrics trending positively. Focus areas for Q{quarter + 1 if quarter < 4 else 1} include operational efficiency and customer satisfaction improvements.</p>",
            "<p><strong>Executive Summary:</strong> This quarter showed significant progress across all departments. Strategic initiatives are on track and budget utilization is within expected parameters.</p>",
            "<p><strong>Executive Summary:</strong> Results exceeded expectations in most areas. Continued investment in technology and talent development is recommended.</p>",
            "<p><strong>Executive Summary:</strong> Strong performance across key metrics demonstrates the effectiveness of our strategic initiatives. We remain well-positioned for continued growth.</p>",
            "<p><strong>Executive Summary:</strong> The organization has made substantial progress toward annual objectives. Cross-functional collaboration has been a key driver of success.</p>",
            "<p><strong>Executive Summary:</strong> Market conditions remain favorable and our competitive position has strengthened. Focus on innovation and customer experience continues to yield positive results.</p>",
            "<p><strong>Executive Summary:</strong> Operational excellence initiatives are delivering measurable improvements. Employee engagement and customer satisfaction scores are at record levels.</p>",
            "<p><strong>Executive Summary:</strong> Financial performance is ahead of plan with strong revenue growth and improved margins. Strategic investments are on track to deliver expected returns.</p>",
        ]
        return random.choice(summaries)
    
    def _generate_action_items(self) -> str:
        """Generate action items with extensive variations."""
        deadline = variations.get_random_deadline()
        
        action_sets = [
            [
                "Review the attached document and provide feedback",
                "Update your team on the changes",
                f"Complete any required training by {deadline}",
            ],
            [
                "Familiarize yourself with the new procedures",
                "Attend the upcoming information session",
                "Contact your manager with any questions",
            ],
            [
                "Update your records as needed",
                "Share this information with relevant team members",
                "Implement changes by the effective date",
            ],
            [
                f"Submit your response by {deadline}",
                "Review the updated guidelines",
                "Confirm your attendance at the briefing session",
            ],
            [
                "Complete the required acknowledgment form",
                "Discuss implications with your team",
                "Identify any potential concerns or questions",
            ],
            [
                "Review the attached materials thoroughly",
                "Prepare questions for the Q&A session",
                "Share feedback through the designated channel",
            ],
            [
                "Update your project plans accordingly",
                "Communicate changes to stakeholders",
                "Schedule follow-up meetings as needed",
            ],
            [
                "Ensure compliance with new requirements",
                "Document any exceptions or special cases",
                "Report progress at the next team meeting",
            ],
        ]
        
        selected_actions = random.choice(action_sets)
        return "".join([f"<li>{action}</li>" for action in selected_actions])
    
    def _generate_personal_message(self) -> str:
        """Generate a personal message for document sharing with extensive variations."""
        messages = [
            "I thought you might find this useful for the upcoming project. Let me know if you have any questions.",
            "As discussed in our meeting, here's the document for your review. Please share your thoughts when you get a chance.",
            "This contains the information you requested. Feel free to reach out if you need any clarification.",
            "Please review this at your earliest convenience. I'd appreciate your feedback by end of week.",
            "Here's the document we talked about. I've highlighted the key sections for your attention.",
            "I've attached the latest version with the updates we discussed. Let me know if anything needs adjustment.",
            "This should have everything you need. Happy to walk through it together if that would be helpful.",
            "As promised, here's the document. I think you'll find section 3 particularly relevant to your work.",
            "I wanted to share this with you before our next meeting. Would love to hear your perspective.",
            "Here's the draft for your review. I'm open to suggestions and feedback.",
            "This is the final version incorporating everyone's input. Please confirm it looks good to you.",
            "I've put together this document based on our discussion. Let me know if I've captured everything correctly.",
            "Sharing this for your reference. No immediate action needed, but thought you'd want to be in the loop.",
            "Here's the analysis you requested. I can schedule time to discuss the findings if you'd like.",
        ]
        return random.choice(messages)
    
    def _generate_action_required(self) -> str:
        """Generate action required section with extensive variations."""
        deadline = variations.get_random_deadline()
        
        actions = [
            f"<p style='background: #fde7e9; padding: 10px; border-radius: 4px;'><strong>⚠️ Action Required:</strong> Please review and provide your feedback by {deadline}.</p>",
            "<p style='background: #fff4ce; padding: 10px; border-radius: 4px;'><strong>📋 Next Steps:</strong> Review the document and confirm receipt.</p>",
            f"<p style='background: #e7f3fe; padding: 10px; border-radius: 4px;'><strong>📅 Deadline:</strong> Please complete your review by {deadline}.</p>",
            "<p style='background: #e8f5e9; padding: 10px; border-radius: 4px;'><strong>✅ For Your Approval:</strong> Please review and approve at your earliest convenience.</p>",
            f"<p style='background: #fce4ec; padding: 10px; border-radius: 4px;'><strong>🔔 Reminder:</strong> Your input is needed by {deadline}.</p>",
            "<p style='background: #f3e5f5; padding: 10px; border-radius: 4px;'><strong>💬 Feedback Requested:</strong> Please share your thoughts on the attached document.</p>",
            "",  # Sometimes no action required
            "",  # Increase probability of no action
        ]
        return random.choice(actions)
    
    def _generate_greeting(self) -> str:
        """Generate a greeting with extensive variations."""
        greetings = [
            "I hope this message finds you well.",
            "I wanted to share an important update with you.",
            "Thank you for your continued dedication to our organization.",
            "I'm pleased to share the following information with you.",
            "I hope you're having a great week.",
            "Thank you for your hard work and commitment.",
            "I'm writing to share some exciting news.",
            "I wanted to take a moment to connect with you.",
            "As we continue to make progress, I wanted to update you.",
            "I'm delighted to share the following with you.",
            "Thank you for being such a valued member of our team.",
            "I hope this update finds you in good spirits.",
            "As we move forward together, I wanted to share some thoughts.",
            "I appreciate your ongoing contributions to our success.",
        ]
        return random.choice(greetings)
    
    def _generate_main_announcement(self) -> str:
        """Generate main announcement content with extensive variations."""
        announcements = [
            "We are implementing new workplace guidelines to enhance our collaborative environment and support flexible working arrangements.",
            "Following careful consideration, we are pleased to announce updates to our benefits program that will take effect next quarter.",
            "As part of our commitment to continuous improvement, we are introducing new tools and processes to streamline our operations.",
            "We are excited to share that our company has achieved a significant milestone in our sustainability journey.",
            "I'm pleased to announce the launch of our new employee development program, designed to support your professional growth.",
            "We are making important changes to our organizational structure to better serve our customers and support our teams.",
            "Our company is embarking on an exciting new initiative that will transform how we work together.",
            "We are proud to announce a strategic partnership that will expand our capabilities and market reach.",
            "Following extensive feedback, we are implementing changes to improve work-life balance across the organization.",
            "We are investing in new technology to enhance productivity and collaboration across all teams.",
            "Our commitment to diversity and inclusion is being strengthened with new programs and resources.",
            "We are pleased to share updates to our compensation and recognition programs.",
        ]
        return random.choice(announcements)
    
    def _generate_details(self) -> str:
        """Generate details section with extensive variations."""
        details = [
            "<p>This initiative is part of our broader strategy to improve employee experience and operational efficiency. We have carefully considered feedback from across the organization in developing these changes.</p>",
            "<p>These updates reflect our commitment to staying competitive and ensuring our team members have the support they need to succeed. Implementation will be phased to minimize disruption.</p>",
            "<p>We believe these changes will have a positive impact on our day-to-day operations. Training and resources will be provided to ensure a smooth transition.</p>",
            "<p>This decision was made after extensive consultation with stakeholders across the organization. We are confident it will deliver significant benefits for everyone.</p>",
            "<p>The changes outlined above are designed to address feedback we've received and position us for continued success. Your input has been invaluable in shaping these initiatives.</p>",
            "<p>Implementation will begin next month, with full rollout expected by end of quarter. Detailed guidance and support resources will be shared in the coming weeks.</p>",
            "<p>We recognize that change can be challenging, and we are committed to supporting you through this transition. Please don't hesitate to reach out with questions or concerns.</p>",
            "<p>These improvements are part of our ongoing commitment to excellence. We will continue to gather feedback and make adjustments as needed.</p>",
        ]
        return random.choice(details)
    
    def _generate_contact_info(self, department: str) -> str:
        """Generate contact information with extensive variations."""
        contacts = {
            "Human Resources": [
                "Please contact HR at hr@company.com or visit the HR SharePoint site for more information.",
                "For HR-related questions, reach out to your HR Business Partner or email hr-support@company.com.",
                "Visit the HR portal for FAQs, or contact the HR team at extension 2000.",
                "Questions? Email hr@company.com or schedule time with your HR representative.",
            ],
            "IT Department": [
                "For technical questions, please submit a ticket through the IT Help Desk or email it-support@company.com.",
                "Need help? Open a ticket at helpdesk.company.com or call the IT Support line at extension 4000.",
                "Contact IT Support via the self-service portal or email techsupport@company.com.",
                "For urgent issues, call the IT hotline. For non-urgent requests, submit a ticket online.",
            ],
            "Finance Department": [
                "For finance-related queries, please contact the Finance team at finance@company.com.",
                "Questions about expenses or budgets? Email finance-support@company.com or visit the Finance SharePoint.",
                "Contact your Finance Business Partner or email accounts@company.com for assistance.",
                "For billing inquiries, email invoices@company.com. For budget questions, contact your department's finance liaison.",
            ],
            "Marketing Department": [
                "For marketing requests, submit a brief through the Marketing portal or email marketing@company.com.",
                "Contact the Marketing team at marketing-requests@company.com for campaign support.",
                "Visit the Brand Center for assets and guidelines, or email brand@company.com for questions.",
            ],
            "Sales Department": [
                "For sales support, contact your regional sales manager or email sales-ops@company.com.",
                "Questions about pricing or proposals? Email sales-support@company.com.",
                "Visit the Sales Hub for resources, or contact the Sales Operations team.",
            ],
            "Legal & Compliance": [
                "For legal questions, email legal@company.com. For compliance matters, contact compliance@company.com.",
                "Submit contract requests through the Legal portal. For urgent matters, call extension 3000.",
                "Contact the Legal team for contract reviews or compliance guidance.",
            ],
            "Claims Department": [
                "For claims inquiries, email claims@company.com or call the Claims hotline at extension 5000.",
                "Submit new claims through the Claims portal. For urgent matters, contact your assigned adjuster directly.",
                "Visit the Claims SharePoint site for forms and guidelines, or email claims-support@company.com.",
                "For claim status updates, check the Claims portal or contact your claims representative.",
            ],
        }
        
        dept_contacts = contacts.get(department, [
            "Please contact your manager or the relevant department for more information.",
            "For questions, reach out to your team lead or department head.",
            "Visit the company intranet for contact information, or email info@company.com.",
        ])
        
        return random.choice(dept_contacts)
    
    def _generate_status_items(self, status_type: str) -> str:
        """Generate status items for project updates with extensive variations."""
        completed_items = [
            "Requirements gathering and analysis",
            "Initial design review completed",
            "Stakeholder approval obtained",
            "Technical architecture finalized",
            "Development environment setup",
            "Sprint 1 deliverables completed",
            "Security review passed",
            "Integration testing completed",
            "Documentation draft completed",
            "Team onboarding finished",
        ]
        
        in_progress_items = [
            f"Development phase - {random.randint(40, 80)}% complete",
            "Testing and quality assurance",
            "Documentation updates",
            "User interface refinements",
            "Performance optimization",
            "Bug fixes and improvements",
            "Integration with external systems",
            "Data migration in progress",
            "Training material preparation",
            "Stakeholder review sessions",
        ]
        
        upcoming_items = [
            "User acceptance testing",
            "Training session preparation",
            "Production deployment",
            "Go-live readiness review",
            "Post-launch monitoring setup",
            "User feedback collection",
            "Performance benchmarking",
            "Documentation finalization",
            "Handover to operations team",
            "Project retrospective",
        ]
        
        items_map = {
            "completed": completed_items,
            "in_progress": in_progress_items,
            "upcoming": upcoming_items,
        }
        
        item_list = items_map.get(status_type, ["Item pending"])
        selected_items = random.sample(item_list, min(random.randint(2, 4), len(item_list)))
        
        return "".join([f'<div class="status-item">{item}</div>' for item in selected_items])
    
    def _generate_next_steps(self) -> str:
        """Generate next steps with extensive variations."""
        deadline = variations.get_random_deadline()
        
        step_sets = [
            [
                f"Complete remaining development tasks by {deadline}",
                "Schedule review meeting with stakeholders",
                "Prepare deployment checklist",
            ],
            [
                "Finalize documentation",
                "Conduct team training session",
                "Begin user acceptance testing",
            ],
            [
                "Review feedback and incorporate changes",
                "Update project timeline",
                "Communicate progress to leadership",
            ],
            [
                "Complete code review process",
                "Run final regression tests",
                "Prepare release notes",
            ],
            [
                "Gather stakeholder sign-off",
                "Schedule deployment window",
                "Notify affected teams",
            ],
            [
                "Address outstanding issues",
                "Update risk register",
                "Plan for contingencies",
            ],
            [
                "Conduct knowledge transfer sessions",
                "Update operational runbooks",
                "Establish support procedures",
            ],
        ]
        
        selected_steps = random.choice(step_sets)
        return "".join([f"<li>{step}</li>" for step in selected_steps])
    
    def _generate_meeting_context(self) -> str:
        """Generate meeting context with extensive variations."""
        contexts = [
            "I'd like to schedule some time to discuss the upcoming project and align on next steps.",
            "Following up on our previous conversation, I think it would be helpful to have a quick sync.",
            "I have some updates to share and would value your input on a few decisions we need to make.",
            "As we approach the deadline, I wanted to ensure we're aligned on deliverables and timeline.",
            "I'd appreciate the opportunity to get your perspective on a few items before we proceed.",
            "There are some developments I'd like to discuss with you at your earliest convenience.",
            "I think a brief meeting would help us clarify expectations and address any concerns.",
            "I'd like to touch base on the progress we've made and discuss the path forward.",
            "Given recent changes, I believe it would be valuable to sync up and realign our approach.",
            "I have some ideas I'd like to run by you and get your feedback on.",
            "As we enter the next phase, I wanted to ensure we have a shared understanding of priorities.",
            "I'd like to discuss some opportunities I've identified that could benefit our project.",
        ]
        return random.choice(contexts)
    
    def _generate_agenda_items(self) -> str:
        """Generate meeting agenda items with extensive variations."""
        agenda_sets = [
            [
                "Review current status and progress",
                "Discuss blockers and challenges",
                "Align on next steps and timeline",
            ],
            [
                "Project overview and objectives",
                "Resource requirements",
                "Q&A and open discussion",
            ],
            [
                "Review action items from last meeting",
                "Updates from each team",
                "Planning for next phase",
            ],
            [
                "Key decisions to be made",
                "Risk assessment and mitigation",
                "Timeline review",
            ],
            [
                "Progress against milestones",
                "Budget and resource update",
                "Stakeholder feedback",
            ],
            [
                "Technical deep dive",
                "Integration requirements",
                "Testing strategy",
            ],
            [
                "Customer feedback review",
                "Feature prioritization",
                "Release planning",
            ],
            [
                "Team capacity planning",
                "Dependency management",
                "Communication plan",
            ],
        ]
        
        selected_agenda = random.choice(agenda_sets)
        return "".join([f"<li>{item}</li>" for item in selected_agenda])
    
    def _generate_metrics(self) -> str:
        """Generate metrics for status reports with extensive variations."""
        metric_sets = [
            [
                ("Tasks Completed", str(random.randint(10, 30))),
                ("On Track", f"{random.randint(80, 100)}%"),
                ("Open Items", str(random.randint(3, 10))),
            ],
            [
                ("Sprint Progress", f"{random.randint(60, 95)}%"),
                ("Bugs Fixed", str(random.randint(5, 20))),
                ("Tests Passing", f"{random.randint(90, 100)}%"),
            ],
            [
                ("Milestones Hit", f"{random.randint(3, 8)}/{random.randint(8, 12)}"),
                ("Team Velocity", str(random.randint(20, 50))),
                ("Blockers", str(random.randint(0, 3))),
            ],
            [
                ("Features Delivered", str(random.randint(5, 15))),
                ("Code Coverage", f"{random.randint(75, 95)}%"),
                ("Days to Launch", str(random.randint(5, 30))),
            ],
        ]
        
        selected_metrics = random.choice(metric_sets)
        
        metrics_html = []
        for label, value in selected_metrics:
            metrics_html.append(f"""
        <div class="metric">
            <div class="metric-value">{value}</div>
            <div class="metric-label">{label}</div>
        </div>""")
        
        return "".join(metrics_html)
    
    def _generate_accomplishments(self) -> str:
        """Generate accomplishments list with extensive variations."""
        accomplishment_pool = [
            "Completed phase 1 of the project ahead of schedule",
            "Resolved critical bug affecting customer experience",
            f"Onboarded {random.randint(2, 5)} new team members",
            "Launched new feature to positive feedback",
            f"Reduced processing time by {random.randint(15, 35)}%",
            "Completed quarterly compliance review",
            "Successfully migrated to new system",
            "Achieved 100% SLA compliance",
            f"Delivered training to {random.randint(30, 100)}+ employees",
            f"Closed {random.randint(10, 30)} support tickets",
            "Implemented automated testing framework",
            f"Improved system performance by {random.randint(20, 50)}%",
            "Completed security audit with no critical findings",
            "Launched customer feedback program",
            f"Reduced costs by ${random.randint(10, 100)}K",
            "Achieved record customer satisfaction scores",
            "Successfully completed vendor evaluation",
            "Implemented new monitoring and alerting system",
        ]
        
        selected = random.sample(accomplishment_pool, random.randint(3, 5))
        return "".join([f"<li>{item}</li>" for item in selected])
    
    def _generate_focus_areas(self) -> str:
        """Generate focus areas list with extensive variations."""
        quarter = variations.get_current_quarter()
        next_quarter = quarter + 1 if quarter < 4 else 1
        
        focus_pool = [
            f"Complete remaining deliverables for Q{quarter}",
            "Improve team collaboration processes",
            "Address technical debt",
            "Customer satisfaction improvements",
            "Process automation initiatives",
            "Team skill development",
            "Quality assurance enhancements",
            "Documentation updates",
            "Stakeholder communication",
            f"Prepare for Q{next_quarter} planning",
            "Performance optimization",
            "Security hardening",
            "User experience improvements",
            "Integration enhancements",
            "Scalability improvements",
            "Cost optimization",
            "Knowledge sharing initiatives",
            "Cross-team collaboration",
        ]
        
        selected = random.sample(focus_pool, random.randint(3, 4))
        return "".join([f"<li>{item}</li>" for item in selected])
    
    def _generate_blockers(self) -> str:
        """Generate blockers list with extensive variations."""
        blocker_pool = [
            "Awaiting approval from legal team",
            "Resource constraints for testing phase",
            "Dependency on external vendor delivery",
            "Budget approval pending",
            "Waiting for stakeholder feedback",
            "Technical dependency on another team",
            "Infrastructure provisioning delayed",
            "Third-party API issues",
            "Scope clarification needed",
            "Resource availability constraints",
            "Pending security review",
            "Waiting for data access permissions",
        ]
        
        no_blockers = [
            "No significant blockers at this time",
            "All blockers have been resolved",
            "No current impediments to progress",
            "Team is operating without major blockers",
        ]
        
        # 40% chance of no blockers
        if random.random() < 0.4:
            return f"<li>{random.choice(no_blockers)}</li>"
        else:
            selected = random.sample(blocker_pool, random.randint(1, 3))
            return "".join([f"<li>{item}</li>" for item in selected])
    
    def _generate_request_context(self) -> str:
        """Generate collaboration request context with extensive variations."""
        contexts = [
            "we're working on a cross-functional initiative that could benefit from your team's expertise",
            "I'm exploring ways to improve our processes and believe your department's input would be valuable",
            "we have an upcoming project that requires collaboration across teams",
            "I've identified an opportunity where our teams could work together effectively",
            "we're facing a challenge that I think your team's experience could help address",
            "there's a strategic initiative that would benefit from your department's perspective",
            "I'm looking to leverage your team's expertise on a new project",
            "we're planning an initiative that aligns with your team's capabilities",
            "I believe there's potential for synergy between our departments on this matter",
            "we're exploring options that could benefit from cross-departmental collaboration",
        ]
        return random.choice(contexts)
    
    def _generate_request_details(self) -> str:
        """Generate collaboration request details with extensive variations."""
        details = [
            "We would appreciate your team's input on the technical requirements and feasibility assessment.",
            "Could you share your insights on best practices and lessons learned from similar initiatives?",
            "We're looking for guidance on compliance requirements and approval processes.",
            "Your expertise in this area would be invaluable as we define our approach.",
            "We'd like to understand how this might impact your team's workflows and priorities.",
            "Could you help us identify potential risks and mitigation strategies?",
            "We're seeking your input on resource requirements and timeline estimates.",
            "Your perspective on stakeholder management would be greatly appreciated.",
            "We'd value your feedback on our proposed approach before we proceed.",
            "Could you advise on integration points and dependencies with your systems?",
        ]
        return random.choice(details)
    
    def _generate_background(self) -> str:
        """Generate background information with extensive variations."""
        quarter = variations.get_current_quarter()
        
        backgrounds = [
            f"This initiative is part of our strategic plan for the year and has executive sponsorship. We're aiming to complete the first phase by end of Q{quarter}.",
            "The project was initiated based on customer feedback and aligns with our goal of improving operational efficiency.",
            "This effort supports our digital transformation objectives and will impact multiple departments across the organization.",
            "Leadership has identified this as a priority initiative with significant potential for business impact.",
            "This project emerged from our annual planning process and addresses key strategic objectives.",
            "The initiative is designed to address gaps identified in our recent assessment and improve overall performance.",
            "This work builds on previous successful projects and extends our capabilities in this area.",
            "The project has been approved by the steering committee and is aligned with our long-term vision.",
            "This initiative responds to market changes and positions us competitively for the future.",
            "The effort is part of our continuous improvement program and has broad organizational support.",
        ]
        return random.choice(backgrounds)
    
    def _generate_policy_overview(self) -> str:
        """Generate policy overview with extensive variations."""
        overviews = [
            "<p>This policy establishes guidelines for employees regarding workplace conduct and expectations. It applies to all full-time and part-time employees.</p>",
            "<p>The updated policy reflects changes in regulations and best practices. It aims to provide clarity and consistency across the organization.</p>",
            "<p>This policy outlines procedures and requirements that all employees must follow. It has been reviewed and approved by leadership.</p>",
            "<p>This document provides comprehensive guidance on organizational standards and expectations. All team members are expected to familiarize themselves with its contents.</p>",
            "<p>The policy has been developed in consultation with stakeholders across the organization. It represents our commitment to maintaining high standards.</p>",
            "<p>This policy supersedes all previous versions and takes effect immediately. Please review the key changes outlined below.</p>",
            "<p>The guidelines contained in this policy are designed to ensure consistency and fairness across all departments and locations.</p>",
            "<p>This policy has been updated to align with current industry standards and regulatory requirements. Your compliance is essential.</p>",
        ]
        return random.choice(overviews)
    
    def _generate_key_changes(self) -> str:
        """Generate key policy changes with extensive variations."""
        change_pool = [
            "Updated eligibility criteria",
            "New approval workflow",
            "Revised documentation requirements",
            "Extended coverage options",
            "Simplified request process",
            "Updated compliance requirements",
            "New reporting procedures",
            "Updated timelines",
            "Additional support resources",
            "Clarified roles and responsibilities",
            "Enhanced security measures",
            "Streamlined escalation process",
            "Updated contact information",
            "New self-service options",
            "Revised exception handling",
            "Updated training requirements",
        ]
        
        selected = random.sample(change_pool, random.randint(3, 5))
        return "".join([f"<li>{change}</li>" for change in selected])
    
    def _generate_document_context(self) -> str:
        """Generate document review context with extensive variations."""
        contexts = [
            "This document outlines our proposed approach for the upcoming initiative. I've incorporated feedback from our initial discussions.",
            "I've drafted this based on our requirements gathering sessions. It's ready for your review before we proceed to the next phase.",
            "This is the latest version incorporating changes from the last review cycle. Please focus on sections 3 and 4 which have been updated.",
            "I've put together this document to capture our discussions and proposed next steps. Your input would be valuable.",
            "This draft reflects the current state of our planning. I'd appreciate your review before we finalize.",
            "I've updated the document based on stakeholder feedback. Please review the highlighted sections.",
            "This version includes the revisions we discussed. Let me know if anything needs further adjustment.",
            "I've compiled this based on input from multiple teams. Your perspective would help ensure completeness.",
            "This document is ready for final review. Please flag any concerns before we proceed to approval.",
            "I've structured this to address the key questions raised in our last meeting. Your feedback is welcome.",
        ]
        return random.choice(contexts)
    
    def _generate_feedback_areas(self) -> str:
        """Generate feedback areas with extensive variations."""
        area_pool = [
            "Overall approach and methodology",
            "Timeline and resource estimates",
            "Risk assessment and mitigation strategies",
            "Technical accuracy",
            "Completeness of requirements",
            "Alignment with business objectives",
            "Clarity and readability",
            "Feasibility of proposed solutions",
            "Any gaps or missing information",
            "Stakeholder impact assessment",
            "Budget and cost considerations",
            "Implementation complexity",
            "Dependencies and assumptions",
            "Success criteria and metrics",
            "Communication and change management",
        ]
        
        selected = random.sample(area_pool, random.randint(3, 4))
        return "".join([f"<li>{area}</li>" for area in selected])
    
    def _generate_leadership_message(self) -> str:
        """Generate leadership message body with extensive variations."""
        quarter = variations.get_current_quarter()
        year = variations.get_current_fiscal_year()
        
        messages = [
            # Quarterly reflection
            f"""<p>As we approach the end of another successful quarter, I wanted to take a moment to reflect on our achievements and share my thoughts on the road ahead.</p>
            <p>Our team has demonstrated remarkable resilience and adaptability in the face of challenges. The results speak for themselves - we've exceeded our targets and delivered exceptional value to our customers.</p>
            <p>Looking forward, I'm excited about the opportunities that lie ahead. We will continue to invest in our people, our technology, and our processes to ensure we remain at the forefront of our industry.</p>
            <p>Thank you for your continued dedication and hard work. Together, we will achieve great things.</p>""",
            
            # Strategic direction
            """<p>I'm writing to share some exciting news about our company's direction and to thank each of you for your contributions.</p>
            <p>Over the past months, we've made significant progress on our strategic initiatives. Our commitment to innovation and excellence has positioned us well for future growth.</p>
            <p>I encourage everyone to continue embracing our values of collaboration, integrity, and customer focus. These principles guide everything we do and are the foundation of our success.</p>
            <p>Please don't hesitate to reach out if you have questions or ideas to share. Your input is always valued.</p>""",
            
            # Gratitude and unity
            """<p>As we navigate through these dynamic times, I wanted to connect with you directly to share updates and express my gratitude.</p>
            <p>Our organization has shown incredible strength and unity. The way our teams have come together to overcome challenges and deliver results is truly inspiring.</p>
            <p>We remain committed to our mission and to supporting each of you in your professional growth. New opportunities and initiatives will be announced in the coming weeks.</p>
            <p>Thank you for being part of this journey. I'm proud of what we've accomplished together and optimistic about our future.</p>""",
            
            # Year-end reflection
            f"""<p>As we reflect on {year}, I am filled with gratitude for the incredible work our teams have accomplished.</p>
            <p>This year has been transformative for our organization. We've launched new products, expanded into new markets, and strengthened our position as an industry leader.</p>
            <p>None of this would have been possible without your dedication, creativity, and commitment to excellence. Each of you has played a vital role in our success.</p>
            <p>As we look ahead to {year + 1}, I am confident that together we will continue to achieve remarkable things. Thank you for your contributions.</p>""",
            
            # Change and transformation
            """<p>Change is a constant in our industry, and I'm proud of how our organization continues to adapt and thrive.</p>
            <p>The initiatives we've undertaken this year are positioning us for long-term success. Our investments in technology, talent, and customer experience are already showing results.</p>
            <p>I want to acknowledge the effort and flexibility you've shown during this period of transformation. Your willingness to embrace new ways of working has been essential.</p>
            <p>We have an exciting roadmap ahead, and I look forward to sharing more details in the coming weeks. Thank you for your continued commitment.</p>""",
            
            # Team appreciation
            """<p>I wanted to take a moment to recognize the outstanding work happening across our organization.</p>
            <p>Every day, I see examples of teamwork, innovation, and dedication that make me proud to lead this company. Your commitment to our customers and to each other is what sets us apart.</p>
            <p>We've achieved significant milestones this quarter, and I want to ensure you know how much your contributions are valued. Success is a team effort, and we have an exceptional team.</p>
            <p>Please continue to bring your best every day. Together, there's no limit to what we can accomplish.</p>""",
            
            # Vision and values
            """<p>Our company's success is built on a foundation of strong values and a clear vision for the future.</p>
            <p>As we continue to grow and evolve, it's important that we stay true to the principles that have guided us: integrity, innovation, and a relentless focus on customer success.</p>
            <p>I'm inspired by how our teams embody these values every day. Whether it's going the extra mile for a customer or supporting a colleague, these actions define who we are.</p>
            <p>Thank you for being ambassadors of our culture. Your commitment to our values is what makes this organization special.</p>""",
            
            # Growth and opportunity
            f"""<p>Q{quarter} has been a period of significant growth and opportunity for our organization.</p>
            <p>We've welcomed new team members, launched exciting initiatives, and strengthened our market position. The momentum we've built is a testament to your hard work and dedication.</p>
            <p>Looking ahead, we have ambitious goals and I'm confident in our ability to achieve them. We will continue to invest in the tools, training, and support you need to succeed.</p>
            <p>I'm grateful for your contributions and excited about what we'll accomplish together in the months ahead.</p>""",
            
            # Customer focus
            """<p>Our customers are at the heart of everything we do, and I'm proud of the exceptional experiences we deliver every day.</p>
            <p>The feedback we receive consistently highlights the professionalism, expertise, and care that our teams provide. This is a direct reflection of your commitment to excellence.</p>
            <p>As we continue to grow, let's never lose sight of what matters most: creating value for our customers and building lasting relationships.</p>
            <p>Thank you for your dedication to our customers and to our mission. Your work makes a real difference.</p>""",
            
            # Innovation and future
            """<p>Innovation has always been a cornerstone of our success, and I'm excited about the creative solutions emerging from our teams.</p>
            <p>The ideas and initiatives you're driving are shaping the future of our company and our industry. Your willingness to challenge the status quo and think differently is invaluable.</p>
            <p>We will continue to create an environment where innovation thrives, where new ideas are welcomed, and where calculated risks are encouraged.</p>
            <p>Keep pushing boundaries and exploring new possibilities. The future we're building together is bright.</p>""",
        ]
        return random.choice(messages)
    
    # =========================================================================
    # External Business Content Generators
    # =========================================================================
    
    def _generate_followup_intro(self) -> str:
        """Generate follow-up email introduction."""
        intros = [
            "I wanted to follow up on our recent conversation and see how things are progressing on your end.",
            "Thank you for taking the time to meet with me last week. I've been thinking about our discussion.",
            "I hope this email finds you well. I wanted to touch base regarding our recent discussions.",
            "Following up on our call, I wanted to share some additional thoughts and next steps.",
            "It was great connecting with you recently. I wanted to continue our conversation.",
            "I hope you had a chance to review the materials I sent over. I wanted to check in.",
            "Thank you for your time during our meeting. I wanted to follow up on a few items.",
            "I wanted to reach out and see if you had any questions about what we discussed.",
            "Following our productive conversation, I wanted to provide some additional information.",
            "I hope you're doing well. I wanted to circle back on our previous discussion.",
        ]
        return random.choice(intros)
    
    def _generate_followup_body(self) -> str:
        """Generate follow-up email body content."""
        bodies = [
            "Based on our discussion, I believe there's a strong opportunity for us to work together. I've outlined some initial thoughts on how we might proceed.",
            "I've had a chance to review the requirements you shared, and I'm confident we can deliver a solution that meets your needs.",
            "After reflecting on our conversation, I wanted to highlight a few key points that I think are particularly relevant to your situation.",
            "I've discussed your requirements with our team, and we're excited about the possibility of partnering with you on this initiative.",
            "I wanted to share some additional resources that might be helpful as you evaluate your options.",
            "Our team has been working on some ideas based on your feedback, and I'd love to share them with you.",
            "I've put together some preliminary recommendations that I think could address the challenges you mentioned.",
            "After our meeting, I did some additional research and found some interesting insights that might be valuable.",
            "I wanted to provide an update on the items we discussed and outline potential next steps.",
            "Based on your feedback, I've refined our approach and would like to walk you through the updated proposal.",
        ]
        return random.choice(bodies)
    
    def _generate_followup_action(self) -> str:
        """Generate follow-up email call to action."""
        actions = [
            "Would you be available for a follow-up call this week to discuss further?",
            "I'd love to schedule a brief call to answer any questions you might have.",
            "Please let me know if you'd like me to send over more detailed information.",
            "I'm happy to arrange a demo or presentation at your convenience.",
            "Would it be helpful to set up a meeting with our technical team?",
            "Let me know if you'd like to proceed with the next steps we discussed.",
            "I'm available to discuss this further whenever works best for your schedule.",
            "Please feel free to reach out if you need any additional information.",
            "I'd be happy to provide references or case studies if that would be helpful.",
            "Let me know how you'd like to proceed, and I'll make the necessary arrangements.",
        ]
        return random.choice(actions)
    
    def _generate_proposal_summary(self) -> str:
        """Generate proposal summary content."""
        summaries = [
            "Our proposal outlines a comprehensive solution designed to address your specific requirements while maximizing value and minimizing risk.",
            "We've developed a tailored approach that leverages our expertise and proven methodologies to deliver measurable results.",
            "This proposal presents a strategic partnership opportunity that aligns with your business objectives and growth plans.",
            "Our solution combines industry best practices with innovative approaches to help you achieve your goals efficiently.",
            "We've designed a flexible engagement model that can scale with your needs and adapt to changing requirements.",
            "The proposed solution addresses the key challenges you've identified while providing a foundation for future growth.",
            "Our approach focuses on delivering quick wins while building toward long-term strategic objectives.",
            "This proposal outlines a phased implementation plan that minimizes disruption while maximizing impact.",
        ]
        return random.choice(summaries)
    
    def _generate_proposal_highlights(self) -> str:
        """Generate proposal highlights as list items."""
        highlight_pool = [
            "Proven track record with similar implementations",
            "Dedicated project team with relevant expertise",
            "Flexible pricing options to fit your budget",
            "Comprehensive training and support included",
            "Clear milestones and deliverables",
            "Risk mitigation strategies built into the approach",
            "Scalable solution that grows with your needs",
            "Integration with your existing systems",
            "Ongoing support and maintenance options",
            "Measurable ROI within the first year",
            "Industry-leading security and compliance",
            "Customizable to your specific requirements",
            "Accelerated timeline with parallel workstreams",
            "Knowledge transfer to your internal team",
        ]
        selected = random.sample(highlight_pool, random.randint(3, 5))
        return "".join([f"<li>{highlight}</li>" for highlight in selected])
    
    def _generate_proposal_next_steps(self) -> str:
        """Generate proposal next steps content."""
        next_steps = [
            "Once you've had a chance to review, I'd suggest we schedule a call to discuss any questions and finalize the scope.",
            "The next step would be to arrange a meeting with key stakeholders to align on priorities and timeline.",
            "I recommend we set up a discovery session to dive deeper into your requirements before finalizing the proposal.",
            "Please review the proposal and let me know if you'd like any modifications before we proceed.",
            "I'm available to present this proposal to your team and address any questions they might have.",
            "Once approved, we can begin the onboarding process and kick off the project within two weeks.",
            "I suggest we schedule a follow-up meeting to discuss the proposal in detail and address any concerns.",
            "Please share this with your team, and let me know when you'd like to discuss next steps.",
        ]
        return random.choice(next_steps)
    
    def _generate_meeting_intro(self) -> str:
        """Generate external meeting request introduction."""
        intros = [
            "I hope this message finds you well. I'm reaching out to schedule some time to discuss a potential collaboration.",
            "I wanted to connect with you regarding an opportunity that I believe could be mutually beneficial.",
            "Following up on our previous correspondence, I'd like to schedule a call to explore this further.",
            "I've been following your company's work and would love the opportunity to discuss how we might work together.",
            "I'm reaching out to see if you'd be available for a brief conversation about your upcoming initiatives.",
            "I wanted to introduce myself and explore whether there might be an opportunity for us to collaborate.",
            "Based on our mutual connections, I thought it would be valuable to connect and share ideas.",
            "I'm reaching out because I believe there's a strong alignment between our organizations.",
        ]
        return random.choice(intros)
    
    def _generate_external_meeting_agenda(self) -> str:
        """Generate external meeting agenda items."""
        agenda_pool = [
            "Introductions and company overviews",
            "Discussion of your current challenges and priorities",
            "Overview of our relevant capabilities and experience",
            "Exploration of potential collaboration opportunities",
            "Review of timeline and next steps",
            "Q&A and open discussion",
            "Partnership structure and engagement models",
            "Case studies and success stories",
            "Technical deep-dive on specific solutions",
            "Budget and resource considerations",
            "Implementation approach and methodology",
            "Support and ongoing relationship",
        ]
        selected = random.sample(agenda_pool, random.randint(3, 5))
        return "".join([f"<li>{item}</li>" for item in selected])
    
    def _generate_meeting_times(self) -> str:
        """Generate proposed meeting times."""
        from datetime import datetime, timedelta
        
        now = datetime.now()
        times = []
        
        # Generate 3 proposed times in the next 2 weeks
        for i in range(3):
            days_ahead = random.randint(2, 10)
            proposed_date = now + timedelta(days=days_ahead)
            # Skip weekends
            while proposed_date.weekday() >= 5:
                proposed_date += timedelta(days=1)
            
            hour = random.choice([9, 10, 11, 14, 15, 16])
            time_str = proposed_date.strftime(f"%A, %B %d at {hour}:00 AM" if hour < 12 else f"%A, %B %d at {hour-12 if hour > 12 else hour}:00 PM")
            times.append(f"<li>{time_str}</li>")
        
        return "".join(times)
    
    def _generate_project_summary(self) -> str:
        """Generate project update summary."""
        summaries = [
            "The project is progressing well and we're on track to meet our key milestones. The team has been working diligently to address the priorities we identified.",
            "We've made significant progress this period and are pleased to report that we're ahead of schedule on several deliverables.",
            "The project continues to move forward as planned. We've successfully completed the current phase and are preparing for the next stage.",
            "Overall, the project is in good shape. We've encountered some minor challenges but have implemented solutions to keep us on track.",
            "We're pleased to report strong progress across all workstreams. The team's collaboration has been excellent.",
            "The project is proceeding according to plan with all major milestones on schedule. Quality metrics remain strong.",
            "We've achieved several important milestones this period and are well-positioned for the upcoming phase.",
            "The project team has been highly productive, and we're confident in our ability to deliver on our commitments.",
        ]
        return random.choice(summaries)
    
    def _generate_project_notes(self) -> str:
        """Generate project notes/additional information."""
        notes = [
            "Please note that we may need to adjust the timeline slightly based on resource availability. I'll keep you informed of any changes.",
            "I'd like to schedule a brief call to discuss some items that require your input before we proceed.",
            "If you have any questions or concerns about the progress, please don't hesitate to reach out.",
            "We're planning a stakeholder review session next week and would appreciate your participation.",
            "The team has identified some opportunities for optimization that we'd like to discuss with you.",
            "Please review the attached documentation and let me know if you need any clarification.",
            "We're on track for the upcoming milestone and will provide a detailed update upon completion.",
            "I'll be sending a more detailed report by end of week with supporting documentation.",
        ]
        return random.choice(notes)
    
    def _generate_invoice_details(self) -> str:
        """Generate invoice details content."""
        services = [
            "Professional Services - Project Implementation",
            "Consulting Services - Strategic Advisory",
            "Software License - Annual Subscription",
            "Support Services - Premium Support Package",
            "Training Services - Team Enablement Program",
            "Development Services - Custom Solution Development",
            "Managed Services - Monthly Retainer",
            "Integration Services - System Integration",
        ]
        
        service = random.choice(services)
        amount = random.randint(5, 50) * 1000
        
        return f"""<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
        <tr style="background-color: #f5f5f5;">
            <th style="padding: 10px; text-align: left; border: 1px solid #ddd;">Description</th>
            <th style="padding: 10px; text-align: right; border: 1px solid #ddd;">Amount</th>
        </tr>
        <tr>
            <td style="padding: 10px; border: 1px solid #ddd;">{service}</td>
            <td style="padding: 10px; text-align: right; border: 1px solid #ddd;">${amount:,}.00</td>
        </tr>
        <tr style="font-weight: bold;">
            <td style="padding: 10px; border: 1px solid #ddd;">Total Due</td>
            <td style="padding: 10px; text-align: right; border: 1px solid #ddd;">${amount:,}.00</td>
        </tr>
        </table>"""
    
    def _generate_payment_terms(self) -> str:
        """Generate payment terms content."""
        terms = [
            "Payment is due within 30 days of invoice date. Please reference the invoice number when making payment.",
            "Net 30 terms apply. Early payment discounts available - contact us for details.",
            "Payment due upon receipt. Please ensure payment is made within 15 business days.",
            "Standard payment terms of Net 45 apply. Wire transfer details are included below.",
            "Payment is due within 30 days. Multiple payment options are available for your convenience.",
            "Please process payment within 30 days. Contact our accounts team if you have any questions.",
        ]
        return random.choice(terms)
    
    def _generate_contract_summary(self) -> str:
        """Generate contract summary content."""
        summaries = [
            "Please find attached the contract documents for your review. This agreement outlines the terms of our engagement and the scope of services to be provided.",
            "I'm pleased to share the finalized contract for our partnership. The document reflects the terms we discussed and agreed upon.",
            "Attached is the service agreement for your signature. Please review the terms carefully and let me know if you have any questions.",
            "The enclosed contract formalizes our business relationship and sets forth the mutual obligations of both parties.",
            "Please review the attached agreement which details the scope, timeline, and commercial terms of our engagement.",
            "I've attached the contract documents for your review and signature. This represents the culmination of our negotiations.",
        ]
        return random.choice(summaries)
    
    def _generate_contract_terms(self) -> str:
        """Generate contract key terms as list items."""
        terms_pool = [
            "Initial term of 12 months with automatic renewal",
            "30-day notice period for termination",
            "Quarterly business reviews included",
            "Service level agreements with defined metrics",
            "Confidentiality and data protection provisions",
            "Intellectual property rights clearly defined",
            "Liability limitations and indemnification",
            "Dispute resolution procedures",
            "Change management process",
            "Payment terms and invoicing schedule",
            "Warranty and support commitments",
            "Compliance with applicable regulations",
        ]
        selected = random.sample(terms_pool, random.randint(4, 6))
        return "".join([f"<li>{term}</li>" for term in selected])
    
    def _generate_contract_next_steps(self) -> str:
        """Generate contract next steps content."""
        next_steps = [
            "Please review the contract and return a signed copy at your earliest convenience. Once executed, we can begin the onboarding process.",
            "If you have any questions or require modifications, please let me know. Otherwise, please sign and return the agreement.",
            "I'm available to discuss any aspects of the contract. Once signed, we'll schedule a kickoff meeting to begin our engagement.",
            "Please have your legal team review the document. I'm happy to address any questions or concerns they may have.",
            "Once you've reviewed and signed the agreement, please return it via email. We'll then proceed with the next steps.",
            "Let me know if you need any clarification on the terms. We're excited to formalize our partnership.",
        ]
        return random.choice(next_steps)
    
    def _generate_introduction_context(self) -> str:
        """Generate introduction/networking email context."""
        contexts = [
            "I came across your profile and was impressed by your work in the industry. I thought it would be valuable to connect.",
            "We met briefly at the conference last month, and I wanted to follow up and introduce myself properly.",
            "A mutual colleague suggested I reach out to you given our shared interests in this space.",
            "I've been following your company's growth and would love the opportunity to learn more about your work.",
            "I'm reaching out because I believe there could be some interesting synergies between our organizations.",
            "Your recent article caught my attention, and I wanted to connect to discuss some of the ideas you shared.",
            "I was referred to you by a colleague who thought we might benefit from connecting.",
            "I noticed we share several connections and thought it would be worthwhile to introduce myself.",
        ]
        return random.choice(contexts)
    
    def _generate_introduction_body(self) -> str:
        """Generate introduction email body content."""
        bodies = [
            "I lead business development at our company, where we specialize in helping organizations like yours achieve their strategic objectives. I'd love to learn more about your current priorities and explore whether there might be opportunities to collaborate.",
            "Our company has been working with several organizations in your industry, and I thought there might be some valuable insights we could share. I'm always interested in connecting with forward-thinking professionals.",
            "I'm passionate about building meaningful professional relationships and believe that connecting with people like yourself can lead to mutually beneficial opportunities.",
            "We've recently launched some initiatives that I think could be relevant to your work. I'd welcome the chance to share more and hear about what you're working on.",
            "I'm always looking to expand my network with talented professionals. I believe that great things happen when the right people connect.",
            "Our paths seem to cross in several areas, and I thought it would be valuable to have a conversation and explore potential areas of collaboration.",
        ]
        return random.choice(bodies)
    
    def _generate_introduction_cta(self) -> str:
        """Generate introduction email call to action."""
        ctas = [
            "Would you be open to a brief call to introduce ourselves and explore potential synergies?",
            "I'd love to buy you a coffee (virtual or in-person) and learn more about your work.",
            "Would you have 15 minutes for a quick introductory call sometime this week or next?",
            "I'd welcome the opportunity to connect and share ideas. Let me know if you'd be interested.",
            "Please feel free to reach out if you'd like to connect. I'm always happy to make time for a conversation.",
            "I'd be grateful for the opportunity to learn from your experience. Would you be open to a brief chat?",
        ]
        return random.choice(ctas)
    
    def _generate_thank_you_context(self) -> str:
        """Generate thank you email context."""
        contexts = [
            "I wanted to take a moment to express my sincere gratitude for your time and support.",
            "Thank you so much for meeting with me yesterday. I really appreciated the opportunity to connect.",
            "I wanted to follow up with a note of thanks for your assistance with our recent project.",
            "I'm writing to express my appreciation for your partnership and collaboration.",
            "Thank you for your continued support and for being such a valued partner.",
            "I wanted to reach out personally to thank you for everything you've done.",
            "Your support has been invaluable, and I wanted to make sure you know how much it's appreciated.",
            "I'm grateful for the opportunity to work with you and wanted to express my thanks.",
        ]
        return random.choice(contexts)
    
    def _generate_thank_you_body(self) -> str:
        """Generate thank you email body content."""
        bodies = [
            "Your insights and guidance have been incredibly valuable. The perspective you shared has helped shape our approach and will undoubtedly contribute to our success.",
            "The time you invested in our discussion was truly appreciated. Your expertise and willingness to share your knowledge made a real difference.",
            "Working with you has been a pleasure. Your professionalism and dedication to excellence are evident in everything you do.",
            "Your contribution to this project has been outstanding. The results we've achieved are a direct reflection of your hard work and commitment.",
            "I'm continually impressed by your expertise and the value you bring to our partnership. Thank you for being such a reliable partner.",
            "Your support during this initiative has been exceptional. We couldn't have achieved these results without your involvement.",
        ]
        return random.choice(bodies)
    
    def _generate_support_context(self) -> str:
        """Generate support/service email context."""
        contexts = [
            "Thank you for contacting our support team. I'm writing to provide an update on your recent inquiry.",
            "I wanted to follow up on the issue you reported and share the steps we've taken to address it.",
            "Thank you for bringing this matter to our attention. We take all customer feedback seriously.",
            "I'm reaching out regarding your recent support request. I have some good news to share.",
            "Thank you for your patience while we investigated the issue you reported.",
            "I wanted to personally follow up on your case and ensure you have all the information you need.",
            "Following up on our recent conversation, I wanted to provide a comprehensive update.",
            "Thank you for being a valued customer. I'm writing to address the concerns you raised.",
        ]
        return random.choice(contexts)
    
    def _generate_support_details(self) -> str:
        """Generate support details content."""
        details = [
            "Our technical team has thoroughly investigated the issue and identified the root cause. We've implemented a fix that should resolve the problem.",
            "After reviewing your account, we've made the necessary adjustments. You should see the changes reflected within 24-48 hours.",
            "We've escalated your request to our specialized team, who are working on a solution. We expect to have this resolved shortly.",
            "The issue you experienced was related to a system update. We've rolled back the changes and confirmed that everything is now working correctly.",
            "We've reviewed your feedback and have taken steps to improve our processes. Your input helps us serve you better.",
            "Our team has completed the requested changes. Please review and let us know if everything meets your expectations.",
        ]
        return random.choice(details)
    
    def _generate_support_resolution(self) -> str:
        """Generate support resolution content."""
        resolutions = [
            "The issue has been fully resolved, and you should now have full access to all features. Please let us know if you experience any further difficulties.",
            "We've credited your account as a gesture of goodwill for any inconvenience caused. Thank you for your understanding.",
            "Your request has been completed successfully. If you have any questions, please don't hesitate to reach out.",
            "We've implemented a permanent fix to prevent this issue from recurring. Thank you for your patience during this process.",
            "The matter has been resolved to your satisfaction, we hope. Please contact us if you need any further assistance.",
            "We've taken steps to ensure this doesn't happen again. Your feedback has been valuable in improving our service.",
        ]
        return random.choice(resolutions)
    
    def _generate_event_details(self) -> str:
        """Generate event invitation details."""
        from datetime import datetime, timedelta
        
        event_types = [
            ("Annual Industry Conference", "Conference Center"),
            ("Executive Roundtable", "Private Dining Room"),
            ("Product Launch Event", "Innovation Hub"),
            ("Networking Reception", "Rooftop Lounge"),
            ("Workshop: Best Practices", "Training Center"),
            ("Customer Appreciation Event", "Grand Ballroom"),
            ("Thought Leadership Summit", "Convention Center"),
            ("Partner Kickoff Meeting", "Corporate Headquarters"),
        ]
        
        event_name, venue = random.choice(event_types)
        event_date = datetime.now() + timedelta(days=random.randint(14, 60))
        
        return f"""<p><strong>Event:</strong> {event_name}</p>
        <p><strong>Date:</strong> {event_date.strftime("%A, %B %d, %Y")}</p>
        <p><strong>Time:</strong> {random.choice(["9:00 AM - 5:00 PM", "2:00 PM - 6:00 PM", "6:00 PM - 9:00 PM"])}</p>
        <p><strong>Venue:</strong> {venue}</p>"""
    
    def _generate_event_agenda(self) -> str:
        """Generate event agenda items."""
        agenda_pool = [
            "Welcome and opening remarks",
            "Keynote presentation",
            "Panel discussion with industry experts",
            "Networking break",
            "Breakout sessions",
            "Product demonstrations",
            "Q&A session",
            "Closing remarks and next steps",
            "Cocktail reception",
            "Awards ceremony",
            "Workshop sessions",
            "Executive fireside chat",
        ]
        selected = random.sample(agenda_pool, random.randint(4, 6))
        return "".join([f"<li>{item}</li>" for item in selected])
    
    def _generate_event_registration(self) -> str:
        """Generate event registration call to action."""
        registrations = [
            "Space is limited, so please RSVP by clicking the link below. We look forward to seeing you there!",
            "Please confirm your attendance by responding to this email. We'll send detailed logistics closer to the date.",
            "Register now to secure your spot. Early registration includes exclusive benefits.",
            "Click here to register and add the event to your calendar. Don't miss this opportunity!",
            "Please let us know if you'll be attending so we can finalize arrangements. We hope to see you there!",
            "RSVP required. Please respond by the end of the week to confirm your participation.",
        ]
        return random.choice(registrations)
    
    def _generate_product_intro(self) -> str:
        """Generate product announcement introduction."""
        intros = [
            "We're excited to announce the launch of our latest innovation, designed to help you achieve more.",
            "I'm thrilled to share some exciting news about a new solution that we've been working on.",
            "We've been listening to your feedback, and today we're proud to introduce something special.",
            "After months of development, we're ready to unveil our newest offering to valued partners like you.",
            "I wanted to personally share some exciting news about our latest product release.",
            "We're pleased to announce a significant enhancement to our product lineup.",
        ]
        return random.choice(intros)
    
    def _generate_product_features(self) -> str:
        """Generate product features as list items."""
        feature_pool = [
            "Enhanced performance and reliability",
            "Intuitive user interface",
            "Advanced analytics and reporting",
            "Seamless integration capabilities",
            "Enterprise-grade security",
            "Scalable architecture",
            "24/7 customer support",
            "Mobile-friendly design",
            "Customizable workflows",
            "Real-time collaboration features",
            "Automated processes",
            "Comprehensive documentation",
        ]
        selected = random.sample(feature_pool, random.randint(4, 6))
        return "".join([f"<li>{feature}</li>" for feature in selected])
    
    def _generate_product_cta(self) -> str:
        """Generate product announcement call to action."""
        ctas = [
            "Schedule a demo today to see how this can benefit your organization.",
            "Contact us to learn more and take advantage of our early adopter pricing.",
            "Visit our website to explore the full feature set and request a trial.",
            "Reply to this email to schedule a personalized walkthrough.",
            "Click here to access exclusive resources and get started today.",
            "Let's schedule a call to discuss how this can address your specific needs.",
        ]
        return random.choice(ctas)
    
    def _generate_feedback_intro(self) -> str:
        """Generate feedback request introduction."""
        intros = [
            "We value your opinion and would love to hear about your experience working with us.",
            "Your feedback is important to us, and we'd appreciate a few minutes of your time.",
            "We're always looking for ways to improve, and your input would be invaluable.",
            "As a valued partner, your perspective matters to us. We'd love to hear your thoughts.",
            "We're committed to continuous improvement and would appreciate your honest feedback.",
            "Your satisfaction is our priority, and we'd like to know how we're doing.",
        ]
        return random.choice(intros)
    
    def _generate_feedback_questions(self) -> str:
        """Generate feedback questions as list items."""
        question_pool = [
            "How would you rate your overall experience?",
            "What aspects of our service do you find most valuable?",
            "Are there areas where we could improve?",
            "Would you recommend us to colleagues?",
            "How responsive have we been to your needs?",
            "What additional services would you find helpful?",
            "How does our solution compare to alternatives?",
            "What would make your experience even better?",
        ]
        selected = random.sample(question_pool, random.randint(3, 4))
        return "".join([f"<li>{question}</li>" for question in selected])
    
    def _generate_feedback_cta(self) -> str:
        """Generate feedback request call to action."""
        ctas = [
            "Please take a moment to complete our brief survey. Your input helps us serve you better.",
            "Click here to share your feedback. It only takes a few minutes.",
            "Reply to this email with your thoughts, or schedule a call if you'd prefer to discuss in person.",
            "We'd love to hear from you. Please share your feedback at your convenience.",
            "Your response would be greatly appreciated. Thank you for helping us improve.",
            "Please let us know your thoughts. Every piece of feedback helps us grow.",
        ]
        return random.choice(ctas)