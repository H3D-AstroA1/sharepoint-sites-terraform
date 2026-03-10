"""
Email content generation with realistic organisational content.

Generates dynamic email content based on templates, recipient context,
and SharePoint site integration.
"""

import random
import string
import secrets
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any, Tuple

from .templates import EMAIL_TEMPLATES, THREADABLE_TEMPLATES, ATTACHMENT_TEMPLATES, PASSWORD_TEMPLATES


class EmailContentGenerator:
    """Generates realistic email content based on templates and context."""
    
    def __init__(self, config: Dict[str, Any], sites: Dict[str, Any]):
        """
        Initialize the content generator.
        
        Args:
            config: Mailbox configuration dictionary.
            sites: SharePoint sites configuration dictionary.
        """
        self.config = config
        self.sites = sites
        self.settings = config.get("settings", {})
        self.departments = config.get("departments", {})
        self.external_senders = config.get("external_senders", {})
        
        # Build site lookup
        self.site_lookup = self._build_site_lookup()
        
    def _build_site_lookup(self) -> Dict[str, Dict]:
        """Build a lookup dictionary for SharePoint sites."""
        lookup = {}
        for site in self.sites.get("sites", []):
            name = site.get("name", "")
            lookup[name] = site
        return lookup
    
    def generate_email(self, recipient: Dict[str, Any]) -> Dict[str, Any]:
        """
        Generate a complete email for a recipient.
        
        Args:
            recipient: Recipient user configuration.
            
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
        }
        
        # Add attachment type if needed
        if email["has_attachment"]:
            email["attachment_type"] = self._select_attachment_type(recipient)
            email["attachment_department"] = recipient.get("department", "General")
        
        return email
    
    def _select_category(self) -> str:
        """Select email category based on configured distribution."""
        distribution = self.settings.get("email_distribution", {
            "newsletters": 15,
            "links": 20,
            "attachments": 15,
            "organisational": 15,
            "interdepartmental": 20,
            "security": 15,
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
        
        # Override with distribution-based selection sometimes
        distribution = self.settings.get("sender_distribution", {})
        if distribution and random.random() < 0.3:  # 30% chance to use distribution
            sender_type = random.choices(
                list(distribution.keys()),
                weights=list(distribution.values())
            )[0]
        
        if sender_type == "external":
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
                "title": sender_user.get("role", "Employee"),
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
    
    def _generate_date(self) -> datetime:
        """Generate a realistic backdated timestamp."""
        date_settings = self.settings.get("date_settings", {})
        months_back = self.settings.get("date_range_months", 12)
        
        # Calculate date range
        end_date = datetime.now()
        start_date = end_date - timedelta(days=months_back * 30)
        days_range = (end_date - start_date).days
        
        # Apply recent bias if configured
        if date_settings.get("recent_bias", True):
            # Use exponential distribution for recent bias
            random_days = int(days_range * (1 - random.random() ** 2))
        else:
            random_days = random.randint(0, days_range)
        
        date = end_date - timedelta(days=random_days)
        
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
            "{newsletter_name}": random.choice(["Industry Insights", "Weekly Digest", "News Roundup"]),
            "{tagline}": "Your weekly source for industry news and insights",
            
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
        }
        
        types = dept_attachments.get(department, ["docx", "xlsx", "pptx", "pdf"])
        return random.choice(types)
    
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
        """Generate a realistic phone number."""
        return f"+1 ({random.randint(200, 999)}) {random.randint(200, 999)}-{random.randint(1000, 9999)}"
    
    def _generate_document_name(self, department: str) -> str:
        """Generate a realistic document name."""
        now = datetime.now()
        
        doc_templates = {
            "Finance Department": [
                f"Budget_Report_Q{(now.month-1)//3+1}_{now.year}.xlsx",
                f"Financial_Statement_{now.strftime('%B')}_{now.year}.pdf",
                f"Expense_Analysis_{now.year}.xlsx",
            ],
            "Human Resources": [
                f"Employee_Handbook_v{random.randint(1,5)}.{random.randint(0,9)}.pdf",
                f"Benefits_Summary_{now.year}.pdf",
                f"Onboarding_Guide_{now.year}.docx",
            ],
            "IT Department": [
                f"System_Architecture_v{random.randint(1,3)}.{random.randint(0,9)}.pptx",
                f"Security_Policy_{now.year}.pdf",
                f"IT_Roadmap_{now.year}.pptx",
            ],
            "Marketing Department": [
                f"Campaign_Results_Q{(now.month-1)//3+1}.pptx",
                f"Brand_Guidelines_v{random.randint(1,3)}.pdf",
                f"Marketing_Plan_{now.year}.pptx",
            ],
            "Executive Leadership": [
                f"Board_Presentation_Q{(now.month-1)//3+1}.pptx",
                f"Strategic_Plan_{now.year}-{now.year+1}.pdf",
                f"Executive_Summary_{now.strftime('%B')}.docx",
            ],
        }
        
        templates = doc_templates.get(department, [
            f"Report_{now.strftime('%Y-%m-%d')}.pdf",
            f"Document_{now.strftime('%B')}_{now.year}.docx",
            f"Presentation_{now.year}.pptx",
        ])
        
        return random.choice(templates)
    
    def _generate_meeting_topic(self, department: str) -> str:
        """Generate a meeting topic."""
        topics = {
            "Finance Department": ["Budget Review", "Q{} Financial Planning", "Cost Analysis", "Audit Preparation"],
            "Human Resources": ["Performance Review Process", "Benefits Update", "Recruitment Strategy", "Training Program"],
            "IT Department": ["System Upgrade", "Security Review", "Infrastructure Planning", "Project Status"],
            "Marketing Department": ["Campaign Planning", "Brand Strategy", "Content Calendar", "Analytics Review"],
            "Sales Department": ["Pipeline Review", "Territory Planning", "Sales Strategy", "Client Feedback"],
            "Executive Leadership": ["Strategic Planning", "Board Preparation", "Quarterly Review", "Leadership Sync"],
        }
        
        dept_topics = topics.get(department, ["Project Update", "Team Sync", "Planning Session", "Review Meeting"])
        topic = random.choice(dept_topics)
        
        if "{}" in topic:
            topic = topic.format((datetime.now().month - 1) // 3 + 1)
        
        return topic
    
    def _generate_proposed_times(self) -> str:
        """Generate proposed meeting times."""
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        times = ["9:00 AM", "10:00 AM", "11:00 AM", "2:00 PM", "3:00 PM", "4:00 PM"]
        
        day1 = random.choice(days)
        day2 = random.choice([d for d in days if d != day1])
        time1 = random.choice(times)
        time2 = random.choice(times)
        
        return f"{day1} at {time1} or {day2} at {time2}"
    
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
        """Generate a project name."""
        prefixes = ["Project", "Initiative", "Program"]
        
        dept_names = {
            "IT Department": ["Digital Transformation", "Cloud Migration", "Security Enhancement", "System Modernization"],
            "Marketing Department": ["Brand Refresh", "Campaign Launch", "Digital Marketing", "Customer Engagement"],
            "Finance Department": ["Cost Optimization", "Financial Systems", "Budget Planning", "Compliance"],
            "Human Resources": ["Talent Development", "Employee Experience", "HR Transformation", "Wellness"],
            "Sales Department": ["Sales Enablement", "CRM Implementation", "Territory Expansion", "Partner Program"],
        }
        
        names = dept_names.get(department, ["Business Improvement", "Process Optimization", "Strategic Initiative"])
        
        return f"{random.choice(prefixes)} {random.choice(names)}"
    
    def _generate_announcement_title(self) -> str:
        """Generate an announcement title."""
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
        ]
        return random.choice(titles)
    
    def _generate_topic(self, department: str) -> str:
        """Generate a general topic."""
        topics = [
            "Q{} Planning".format((datetime.now().month - 1) // 3 + 1),
            "Process Improvement",
            "Resource Allocation",
            "Timeline Review",
            "Budget Discussion",
            "Team Collaboration",
            "Project Requirements",
            "Strategy Alignment",
        ]
        return random.choice(topics)
    
    def _generate_company_news(self) -> str:
        """Generate company news content."""
        news_items = [
            "<p>We're excited to announce the launch of our new customer portal, which will streamline how we interact with our clients.</p>",
            "<p>Congratulations to the Sales team for exceeding Q{} targets by 15%! This achievement reflects the hard work and dedication of everyone involved.</p>".format((datetime.now().month - 1) // 3 + 1),
            "<p>Our sustainability initiative has reached a major milestone - we've reduced our carbon footprint by 20% this year.</p>",
            "<p>We're pleased to welcome 15 new team members who joined us this month across various departments.</p>",
            "<p>The company has been recognized as one of the top employers in our industry for the third consecutive year.</p>",
        ]
        return random.choice(news_items)
    
    def _generate_upcoming_events(self) -> str:
        """Generate upcoming events content."""
        now = datetime.now()
        events = [
            f"<p>📅 <strong>Town Hall Meeting</strong> - {(now + timedelta(days=random.randint(5, 15))).strftime('%B %d')} at 2:00 PM</p>",
            f"<p>📅 <strong>Team Building Event</strong> - {(now + timedelta(days=random.randint(10, 20))).strftime('%B %d')}</p>",
            f"<p>📅 <strong>Training Session</strong> - {(now + timedelta(days=random.randint(3, 10))).strftime('%B %d')} at 10:00 AM</p>",
            f"<p>📅 <strong>Quarterly Review</strong> - {(now + timedelta(days=random.randint(15, 30))).strftime('%B %d')}</p>",
        ]
        return random.choice(events)
    
    def _generate_employee_spotlight(self) -> str:
        """Generate employee spotlight content."""
        spotlights = [
            "<p>This week we're celebrating <strong>Sarah Johnson</strong> from the Marketing team for her outstanding work on the brand refresh project!</p>",
            "<p>Congratulations to <strong>Michael Chen</strong> from IT for completing his AWS certification!</p>",
            "<p>A big thank you to <strong>Emily Davis</strong> from HR for organizing the successful wellness week!</p>",
            "<p>Kudos to <strong>James Wilson</strong> from Sales for closing our biggest deal of the quarter!</p>",
        ]
        return random.choice(spotlights)
    
    def _generate_fun_fact(self) -> str:
        """Generate a fun fact."""
        facts = [
            "Our company has employees in 12 different countries!",
            "We've served over 10,000 customers this year.",
            "The average tenure of our employees is 4.5 years.",
            "Our team has collectively volunteered over 500 hours this quarter.",
            "We've reduced paper usage by 40% since going digital.",
        ]
        return random.choice(facts)
    
    def _generate_department_updates(self) -> str:
        """Generate department updates content."""
        updates = [
            "<p><strong>IT:</strong> New security protocols will be implemented next week.</p><p><strong>HR:</strong> Open enrollment for benefits begins next month.</p>",
            "<p><strong>Marketing:</strong> New campaign launching next week.</p><p><strong>Sales:</strong> Q{} targets exceeded by 10%.</p>".format((datetime.now().month - 1) // 3 + 1),
            "<p><strong>Finance:</strong> Budget planning for next year has begun.</p><p><strong>Operations:</strong> New efficiency measures showing positive results.</p>",
        ]
        return random.choice(updates)
    
    def _generate_articles(self) -> str:
        """Generate newsletter articles."""
        articles = """
        <div class="article">
            <div class="article-title">Industry Trends: What to Watch in {}</div>
            <div class="article-meta">5 min read</div>
            <p class="article-summary">The latest developments shaping our industry and what they mean for businesses like ours.</p>
            <a href="#" class="read-more">Read More →</a>
        </div>
        <div class="article">
            <div class="article-title">Best Practices for Remote Collaboration</div>
            <div class="article-meta">3 min read</div>
            <p class="article-summary">Tips and tools for effective teamwork in a distributed environment.</p>
            <a href="#" class="read-more">Read More →</a>
        </div>
        <div class="article">
            <div class="article-title">Technology Spotlight: AI in the Workplace</div>
            <div class="article-meta">4 min read</div>
            <p class="article-summary">How artificial intelligence is transforming how we work.</p>
            <a href="#" class="read-more">Read More →</a>
        </div>
        """.format(datetime.now().year)
        return articles
    
    def _generate_activity_items(self) -> str:
        """Generate SharePoint activity items."""
        activities = """
        <div class="activity-item">
            <span class="activity-icon">📄</span>
            <span class="activity-text">New document uploaded: Q{} Report.xlsx</span>
            <div class="activity-meta">2 hours ago by John Smith</div>
        </div>
        <div class="activity-item">
            <span class="activity-icon">✏️</span>
            <span class="activity-text">Document edited: Project Plan.docx</span>
            <div class="activity-meta">Yesterday by Sarah Johnson</div>
        </div>
        <div class="activity-item">
            <span class="activity-icon">📁</span>
            <span class="activity-text">New folder created: Archive 2024</span>
            <div class="activity-meta">3 days ago by Admin</div>
        </div>
        """.format((datetime.now().month - 1) // 3 + 1)
        return activities
    
    def _generate_key_points(self, department: str) -> str:
        """Generate key points for reports."""
        points = {
            "Finance Department": [
                "<li>Revenue increased by 8% compared to last quarter</li>",
                "<li>Operating costs reduced by 5%</li>",
                "<li>Cash flow remains strong with healthy reserves</li>",
            ],
            "Human Resources": [
                "<li>Employee satisfaction score: 4.2/5</li>",
                "<li>Turnover rate decreased by 3%</li>",
                "<li>15 new hires onboarded this month</li>",
            ],
            "IT Department": [
                "<li>System uptime: 99.9%</li>",
                "<li>Security incidents: 0</li>",
                "<li>Help desk tickets resolved: 95%</li>",
            ],
            "Marketing Department": [
                "<li>Website traffic increased by 25%</li>",
                "<li>Social media engagement up 40%</li>",
                "<li>Lead generation exceeded targets</li>",
            ],
        }
        dept_points = points.get(department, [
            "<li>Key metric 1: On track</li>",
            "<li>Key metric 2: Improved</li>",
            "<li>Key metric 3: Meeting targets</li>",
        ])
        return "".join(dept_points)
    
    def _generate_executive_summary(self) -> str:
        """Generate executive summary."""
        summaries = [
            "<p><strong>Executive Summary:</strong> Overall performance remains strong with key metrics trending positively. Focus areas for next period include operational efficiency and customer satisfaction improvements.</p>",
            "<p><strong>Executive Summary:</strong> This quarter showed significant progress across all departments. Strategic initiatives are on track and budget utilization is within expected parameters.</p>",
            "<p><strong>Executive Summary:</strong> Results exceeded expectations in most areas. Continued investment in technology and talent development is recommended.</p>",
        ]
        return random.choice(summaries)
    
    def _generate_action_items(self) -> str:
        """Generate action items."""
        items = [
            "<li>Review the attached document and provide feedback</li><li>Update your team on the changes</li><li>Complete any required training by the deadline</li>",
            "<li>Familiarize yourself with the new procedures</li><li>Attend the upcoming information session</li><li>Contact your manager with any questions</li>",
            "<li>Update your records as needed</li><li>Share this information with relevant team members</li><li>Implement changes by the effective date</li>",
        ]
        return random.choice(items)
    
    def _generate_personal_message(self) -> str:
        """Generate a personal message for document sharing."""
        messages = [
            "I thought you might find this useful for the upcoming project. Let me know if you have any questions.",
            "As discussed in our meeting, here's the document for your review. Please share your thoughts when you get a chance.",
            "This contains the information you requested. Feel free to reach out if you need any clarification.",
            "Please review this at your earliest convenience. I'd appreciate your feedback by end of week.",
        ]
        return random.choice(messages)
    
    def _generate_action_required(self) -> str:
        """Generate action required section."""
        actions = [
            "<p style='background: #fde7e9; padding: 10px; border-radius: 4px;'><strong>⚠️ Action Required:</strong> Please review and provide your feedback by the deadline.</p>",
            "<p style='background: #fff4ce; padding: 10px; border-radius: 4px;'><strong>📋 Next Steps:</strong> Review the document and confirm receipt.</p>",
            "",  # Sometimes no action required
        ]
        return random.choice(actions)
    
    def _generate_greeting(self) -> str:
        """Generate a greeting."""
        greetings = [
            "I hope this message finds you well.",
            "I wanted to share an important update with you.",
            "Thank you for your continued dedication to our organization.",
            "I'm pleased to share the following information with you.",
        ]
        return random.choice(greetings)
    
    def _generate_main_announcement(self) -> str:
        """Generate main announcement content."""
        announcements = [
            "We are implementing new workplace guidelines to enhance our collaborative environment and support flexible working arrangements.",
            "Following careful consideration, we are pleased to announce updates to our benefits program that will take effect next quarter.",
            "As part of our commitment to continuous improvement, we are introducing new tools and processes to streamline our operations.",
            "We are excited to share that our company has achieved a significant milestone in our sustainability journey.",
        ]
        return random.choice(announcements)
    
    def _generate_details(self) -> str:
        """Generate details section."""
        details = [
            "<p>This initiative is part of our broader strategy to improve employee experience and operational efficiency. We have carefully considered feedback from across the organization in developing these changes.</p>",
            "<p>These updates reflect our commitment to staying competitive and ensuring our team members have the support they need to succeed. Implementation will be phased to minimize disruption.</p>",
            "<p>We believe these changes will have a positive impact on our day-to-day operations. Training and resources will be provided to ensure a smooth transition.</p>",
        ]
        return random.choice(details)
    
    def _generate_contact_info(self, department: str) -> str:
        """Generate contact information."""
        contacts = {
            "Human Resources": "Please contact HR at hr@company.com or visit the HR SharePoint site for more information.",
            "IT Department": "For technical questions, please submit a ticket through the IT Help Desk or email it-support@company.com.",
            "Finance Department": "For finance-related queries, please contact the Finance team at finance@company.com.",
        }
        return contacts.get(department, "Please contact your manager or the relevant department for more information.")
    
    def _generate_status_items(self, status_type: str) -> str:
        """Generate status items for project updates."""
        items = {
            "completed": [
                '<div class="status-item">Requirements gathering and analysis</div>',
                '<div class="status-item">Initial design review completed</div>',
                '<div class="status-item">Stakeholder approval obtained</div>',
            ],
            "in_progress": [
                '<div class="status-item">Development phase - 60% complete</div>',
                '<div class="status-item">Testing and quality assurance</div>',
                '<div class="status-item">Documentation updates</div>',
            ],
            "upcoming": [
                '<div class="status-item">User acceptance testing</div>',
                '<div class="status-item">Training session preparation</div>',
                '<div class="status-item">Production deployment</div>',
            ],
        }
        return "".join(items.get(status_type, ['<div class="status-item">Item pending</div>']))
    
    def _generate_next_steps(self) -> str:
        """Generate next steps."""
        steps = [
            "<li>Complete remaining development tasks by Friday</li><li>Schedule review meeting with stakeholders</li><li>Prepare deployment checklist</li>",
            "<li>Finalize documentation</li><li>Conduct team training session</li><li>Begin user acceptance testing</li>",
            "<li>Review feedback and incorporate changes</li><li>Update project timeline</li><li>Communicate progress to leadership</li>",
        ]
        return random.choice(steps)
    
    def _generate_meeting_context(self) -> str:
        """Generate meeting context."""
        contexts = [
            "I'd like to schedule some time to discuss the upcoming project and align on next steps.",
            "Following up on our previous conversation, I think it would be helpful to have a quick sync.",
            "I have some updates to share and would value your input on a few decisions we need to make.",
            "As we approach the deadline, I wanted to ensure we're aligned on deliverables and timeline.",
        ]
        return random.choice(contexts)
    
    def _generate_agenda_items(self) -> str:
        """Generate meeting agenda items."""
        agendas = [
            "<li>Review current status and progress</li><li>Discuss blockers and challenges</li><li>Align on next steps and timeline</li>",
            "<li>Project overview and objectives</li><li>Resource requirements</li><li>Q&A and open discussion</li>",
            "<li>Review action items from last meeting</li><li>Updates from each team</li><li>Planning for next phase</li>",
        ]
        return random.choice(agendas)
    
    def _generate_metrics(self) -> str:
        """Generate metrics for status reports."""
        return """
        <div class="metric">
            <div class="metric-value">{}</div>
            <div class="metric-label">Tasks Completed</div>
        </div>
        <div class="metric">
            <div class="metric-value">{}%</div>
            <div class="metric-label">On Track</div>
        </div>
        <div class="metric">
            <div class="metric-value">{}</div>
            <div class="metric-label">Open Items</div>
        </div>
        """.format(random.randint(10, 30), random.randint(80, 100), random.randint(3, 10))
    
    def _generate_accomplishments(self) -> str:
        """Generate accomplishments list."""
        accomplishments = [
            "<li>Completed phase 1 of the project ahead of schedule</li><li>Resolved critical bug affecting customer experience</li><li>Onboarded 3 new team members</li>",
            "<li>Launched new feature to positive feedback</li><li>Reduced processing time by 20%</li><li>Completed quarterly compliance review</li>",
            "<li>Successfully migrated to new system</li><li>Achieved 100% SLA compliance</li><li>Delivered training to 50+ employees</li>",
        ]
        return random.choice(accomplishments)
    
    def _generate_focus_areas(self) -> str:
        """Generate focus areas list."""
        areas = [
            "<li>Complete remaining deliverables for Q{}</li><li>Improve team collaboration processes</li><li>Address technical debt</li>".format((datetime.now().month - 1) // 3 + 1),
            "<li>Customer satisfaction improvements</li><li>Process automation initiatives</li><li>Team skill development</li>",
            "<li>Quality assurance enhancements</li><li>Documentation updates</li><li>Stakeholder communication</li>",
        ]
        return random.choice(areas)
    
    def _generate_blockers(self) -> str:
        """Generate blockers list."""
        blockers = [
            "<li>Awaiting approval from legal team</li><li>Resource constraints for testing phase</li>",
            "<li>Dependency on external vendor delivery</li><li>Budget approval pending</li>",
            "<li>No significant blockers at this time</li>",
        ]
        return random.choice(blockers)
    
    def _generate_request_context(self) -> str:
        """Generate collaboration request context."""
        contexts = [
            "we're working on a cross-functional initiative that could benefit from your team's expertise",
            "I'm exploring ways to improve our processes and believe your department's input would be valuable",
            "we have an upcoming project that requires collaboration across teams",
        ]
        return random.choice(contexts)
    
    def _generate_request_details(self) -> str:
        """Generate collaboration request details."""
        details = [
            "We would appreciate your team's input on the technical requirements and feasibility assessment.",
            "Could you share your insights on best practices and lessons learned from similar initiatives?",
            "We're looking for guidance on compliance requirements and approval processes.",
        ]
        return random.choice(details)
    
    def _generate_background(self) -> str:
        """Generate background information."""
        backgrounds = [
            "This initiative is part of our strategic plan for the year and has executive sponsorship. We're aiming to complete the first phase by end of quarter.",
            "The project was initiated based on customer feedback and aligns with our goal of improving operational efficiency.",
            "This effort supports our digital transformation objectives and will impact multiple departments across the organization.",
        ]
        return random.choice(backgrounds)
    
    def _generate_policy_overview(self) -> str:
        """Generate policy overview."""
        overviews = [
            "<p>This policy establishes guidelines for employees regarding workplace conduct and expectations. It applies to all full-time and part-time employees.</p>",
            "<p>The updated policy reflects changes in regulations and best practices. It aims to provide clarity and consistency across the organization.</p>",
            "<p>This policy outlines procedures and requirements that all employees must follow. It has been reviewed and approved by leadership.</p>",
        ]
        return random.choice(overviews)
    
    def _generate_key_changes(self) -> str:
        """Generate key policy changes."""
        changes = [
            "<li>Updated eligibility criteria</li><li>New approval workflow</li><li>Revised documentation requirements</li>",
            "<li>Extended coverage options</li><li>Simplified request process</li><li>Updated compliance requirements</li>",
            "<li>New reporting procedures</li><li>Updated timelines</li><li>Additional support resources</li>",
        ]
        return random.choice(changes)
    
    def _generate_document_context(self) -> str:
        """Generate document review context."""
        contexts = [
            "This document outlines our proposed approach for the upcoming initiative. I've incorporated feedback from our initial discussions.",
            "I've drafted this based on our requirements gathering sessions. It's ready for your review before we proceed to the next phase.",
            "This is the latest version incorporating changes from the last review cycle. Please focus on sections 3 and 4 which have been updated.",
        ]
        return random.choice(contexts)
    
    def _generate_feedback_areas(self) -> str:
        """Generate feedback areas."""
        areas = [
            "<li>Overall approach and methodology</li><li>Timeline and resource estimates</li><li>Risk assessment and mitigation strategies</li>",
            "<li>Technical accuracy</li><li>Completeness of requirements</li><li>Alignment with business objectives</li>",
            "<li>Clarity and readability</li><li>Feasibility of proposed solutions</li><li>Any gaps or missing information</li>",
        ]
        return random.choice(areas)
    
    def _generate_leadership_message(self) -> str:
        """Generate leadership message body."""
        messages = [
            """<p>As we approach the end of another successful quarter, I wanted to take a moment to reflect on our achievements and share my thoughts on the road ahead.</p>
            <p>Our team has demonstrated remarkable resilience and adaptability in the face of challenges. The results speak for themselves - we've exceeded our targets and delivered exceptional value to our customers.</p>
            <p>Looking forward, I'm excited about the opportunities that lie ahead. We will continue to invest in our people, our technology, and our processes to ensure we remain at the forefront of our industry.</p>
            <p>Thank you for your continued dedication and hard work. Together, we will achieve great things.</p>""",
            
            """<p>I'm writing to share some exciting news about our company's direction and to thank each of you for your contributions.</p>
            <p>Over the past months, we've made significant progress on our strategic initiatives. Our commitment to innovation and excellence has positioned us well for future growth.</p>
            <p>I encourage everyone to continue embracing our values of collaboration, integrity, and customer focus. These principles guide everything we do and are the foundation of our success.</p>
            <p>Please don't hesitate to reach out if you have questions or ideas to share. Your input is always valued.</p>""",
            
            """<p>As we navigate through these dynamic times, I wanted to connect with you directly to share updates and express my gratitude.</p>
            <p>Our organization has shown incredible strength and unity. The way our teams have come together to overcome challenges and deliver results is truly inspiring.</p>
            <p>We remain committed to our mission and to supporting each of you in your professional growth. New opportunities and initiatives will be announced in the coming weeks.</p>
            <p>Thank you for being part of this journey. I'm proud of what we've accomplished together and optimistic about our future.</p>""",
        ]
        return random.choice(messages)