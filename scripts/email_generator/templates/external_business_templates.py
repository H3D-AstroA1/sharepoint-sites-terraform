"""
External Business Communication Templates.

This module contains templates for legitimate external business emails
such as client communications, vendor correspondence, partner updates,
and professional networking emails.
"""

from typing import Any, Dict, List

# Realistic external business domains
EXTERNAL_BUSINESS_DOMAINS = [
    # Consulting firms
    "accenture.com",
    "deloitte.com",
    "pwc.com",
    "kpmg.com",
    "mckinsey.com",
    "bcg.com",
    "ey.com",
    
    # Technology companies
    "microsoft.com",
    "google.com",
    "amazon.com",
    "salesforce.com",
    "oracle.com",
    "sap.com",
    "ibm.com",
    "cisco.com",
    "adobe.com",
    "servicenow.com",
    "workday.com",
    "zendesk.com",
    "hubspot.com",
    "slack.com",
    "zoom.us",
    "dropbox.com",
    "atlassian.com",
    
    # Financial services
    "jpmorgan.com",
    "goldmansachs.com",
    "morganstanley.com",
    "barclays.com",
    "hsbc.com",
    "citi.com",
    "ubs.com",
    
    # Professional services
    "linkedin.com",
    "indeed.com",
    "glassdoor.com",
    
    # Industry-specific
    "gartner.com",
    "forrester.com",
    "idc.com",
    
    # Generic business domains
    "globaltech-solutions.com",
    "enterprise-partners.com",
    "business-consulting.com",
    "professional-services.net",
    "corporate-solutions.com",
    "strategic-advisors.com",
    "innovation-labs.com",
    "digital-ventures.com",
    "growth-partners.com",
    "market-insights.com",
    "industry-experts.com",
    "premier-consulting.com",
    "executive-search.com",
    "talent-solutions.com",
    "supply-chain-partners.com",
    "logistics-global.com",
    "manufacturing-solutions.com",
    "quality-assurance.com",
    "compliance-experts.com",
    "legal-advisors.com",
    "financial-planning.com",
    "investment-group.com",
    "capital-partners.com",
    "venture-associates.com",
    "tech-innovations.com",
    "cloud-services.com",
    "data-analytics.com",
    "cyber-security.com",
    "it-solutions.com",
    "software-development.com",
    "web-services.com",
    "mobile-solutions.com",
    "ai-innovations.com",
    "machine-learning.com",
]

# External sender names (first name, last name, title)
EXTERNAL_SENDER_PROFILES = [
    ("James", "Mitchell", "Account Manager"),
    ("Sarah", "Thompson", "Client Success Manager"),
    ("Michael", "Chen", "Senior Consultant"),
    ("Emily", "Rodriguez", "Business Development Manager"),
    ("David", "Williams", "Project Manager"),
    ("Jennifer", "Brown", "Sales Director"),
    ("Robert", "Taylor", "Partnership Manager"),
    ("Lisa", "Anderson", "Customer Success Lead"),
    ("Christopher", "Martinez", "Solutions Architect"),
    ("Amanda", "Garcia", "Engagement Manager"),
    ("Daniel", "Lee", "Technical Account Manager"),
    ("Michelle", "Wilson", "Regional Sales Manager"),
    ("Kevin", "Johnson", "Strategic Accounts Director"),
    ("Rachel", "Davis", "Client Relations Manager"),
    ("Andrew", "Miller", "Business Analyst"),
    ("Stephanie", "Moore", "Marketing Manager"),
    ("Brian", "Jackson", "Operations Manager"),
    ("Nicole", "White", "Product Specialist"),
    ("Matthew", "Harris", "Implementation Consultant"),
    ("Lauren", "Martin", "Account Executive"),
    ("Ryan", "Thompson", "Delivery Manager"),
    ("Jessica", "Clark", "Program Manager"),
    ("Brandon", "Lewis", "Technical Consultant"),
    ("Megan", "Walker", "Customer Experience Manager"),
    ("Justin", "Hall", "Sales Engineer"),
    ("Ashley", "Allen", "Relationship Manager"),
    ("Tyler", "Young", "Business Development Rep"),
    ("Samantha", "King", "Client Advisor"),
    ("Patrick", "Wright", "Solutions Consultant"),
    ("Heather", "Scott", "Account Director"),
]

# External business email templates
EXTERNAL_BUSINESS_TEMPLATES: List[Dict[str, Any]] = [
    # Client/Vendor Follow-up
    {
        "name": "client_followup",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "Following up on our conversation",
            "Re: Our recent discussion",
            "Quick follow-up from our meeting",
            "Checking in - {topic}",
            "Following up on {topic}",
            "Re: Next steps for {project_name}",
            "Touching base on our partnership",
            "Following up from {company_name}",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Hi {recipient_first_name},</p>

<p>{followup_intro}</p>

<p>{followup_body}</p>

<p>{followup_action}</p>

<p>Please let me know if you have any questions or if there's anything else I can help with. You can also <a href="{calendar_link}">schedule a call directly on my calendar</a>.</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company} | <a href="{company_website}">Website</a><br>
{sender_phone}<br>
<a href="mailto:{sender_email}">{sender_email}</a></p>
</body>
</html>""",
        "has_attachment": False,
        "supports_threading": True,
    },
    
    # Proposal/Quote
    {
        "name": "proposal_quote",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "Proposal for {project_name}",
            "Quote Request - {company_name}",
            "Pricing proposal for your review",
            "Updated proposal - {topic}",
            "Re: RFP Response - {project_name}",
            "Service proposal for {company_name}",
            "Partnership proposal",
            "Quotation for requested services",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Dear {recipient_first_name},</p>

<p>Thank you for the opportunity to submit this proposal. Based on our discussions, I'm pleased to provide the following for your consideration.</p>

<p>{proposal_summary}</p>

<p><strong>Key Highlights:</strong></p>
<ul>
{proposal_highlights}
</ul>

<p><strong>Next Steps:</strong></p>
<p>{proposal_next_steps}</p>

<p>I've attached the detailed proposal document for your review. Please don't hesitate to reach out if you have any questions or would like to schedule a call to discuss further.</p>

<p>Looking forward to the opportunity to work together.</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company}<br>
{sender_phone}<br>
{sender_email}</p>
</body>
</html>""",
        "has_attachment": True,
        "supports_threading": True,
    },
    
    # Meeting Request
    {
        "name": "external_meeting_request",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "Meeting request - {topic}",
            "Can we schedule a call?",
            "Request for meeting - {company_name}",
            "Let's connect - {topic}",
            "Scheduling a discussion",
            "Meeting to discuss partnership",
            "Call request - {project_name}",
            "Would you have time for a quick call?",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Hi {recipient_first_name},</p>

<p>{meeting_intro}</p>

<p>I'd like to schedule a {meeting_duration} call to discuss:</p>
<ul>
{meeting_agenda}
</ul>

<p>Would any of the following times work for you?</p>
<ul>
{meeting_times}
</ul>

<p>Alternatively, you can <a href="{calendar_link}">book a time directly on my calendar</a> or join via <a href="{zoom_link}">Zoom</a>.</p>

<p>Please let me know what works best for your schedule, or feel free to suggest alternative times.</p>

<p>Looking forward to connecting.</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company} | <a href="{company_website}">Website</a><br>
{sender_phone}<br>
<a href="mailto:{sender_email}">{sender_email}</a></p>
</body>
</html>""",
        "has_attachment": False,
        "supports_threading": True,
    },
    
    # Project Update
    {
        "name": "external_project_update",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "Project Update - {project_name}",
            "Status update: {project_name}",
            "Weekly update - {topic}",
            "Progress report - {project_name}",
            "Milestone achieved - {project_name}",
            "Project status: {topic}",
            "Update on deliverables",
            "Implementation progress report",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Hi {recipient_first_name},</p>

<p>I wanted to provide you with an update on the progress of {project_name}.</p>

<p><strong>Summary:</strong></p>
<p>{project_summary}</p>

<p><strong>Completed This Period:</strong></p>
<ul>
{completed_items}
</ul>

<p><strong>In Progress:</strong></p>
<ul>
{in_progress_items}
</ul>

<p><strong>Upcoming:</strong></p>
<ul>
{upcoming_items}
</ul>

<p>{project_notes}</p>

<p>Please let me know if you have any questions or concerns.</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company}<br>
{sender_phone}<br>
{sender_email}</p>
</body>
</html>""",
        "has_attachment": False,
        "supports_threading": True,
    },
    
    # Invoice/Billing
    {
        "name": "invoice_billing",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "Invoice #{invoice_number} - {company_name}",
            "Monthly invoice - {month}",
            "Payment reminder - Invoice #{invoice_number}",
            "Billing statement - {company_name}",
            "Invoice for services rendered",
            "Account statement - {month}",
            "Payment confirmation request",
            "Updated invoice - {company_name}",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Dear {recipient_first_name},</p>

<p>Please find attached the invoice for services provided during {billing_period}.</p>

<p><strong>Invoice Details:</strong></p>
<table style="border-collapse: collapse; margin: 15px 0;">
<tr><td style="padding: 5px 15px 5px 0;"><strong>Invoice Number:</strong></td><td>#{invoice_number}</td></tr>
<tr><td style="padding: 5px 15px 5px 0;"><strong>Invoice Date:</strong></td><td>{invoice_date}</td></tr>
<tr><td style="padding: 5px 15px 5px 0;"><strong>Due Date:</strong></td><td>{due_date}</td></tr>
<tr><td style="padding: 5px 15px 5px 0;"><strong>Amount Due:</strong></td><td>{invoice_amount}</td></tr>
</table>

<p><strong>Payment Instructions:</strong></p>
<p>{payment_instructions}</p>

<p>If you have any questions regarding this invoice, please don't hesitate to contact me.</p>

<p>Thank you for your business.</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company}<br>
{sender_phone}<br>
{sender_email}</p>
</body>
</html>""",
        "has_attachment": True,
        "supports_threading": True,
    },
    
    # Contract/Agreement
    {
        "name": "contract_agreement",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "Contract for review - {project_name}",
            "Service agreement - {company_name}",
            "NDA for signature",
            "Partnership agreement - review required",
            "Contract renewal - {company_name}",
            "Updated terms and conditions",
            "Agreement for your signature",
            "MSA for review - {company_name}",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Dear {recipient_first_name},</p>

<p>Please find attached the {contract_type} for your review and signature.</p>

<p><strong>Document Summary:</strong></p>
<p>{contract_summary}</p>

<p><strong>Key Terms:</strong></p>
<ul>
{contract_terms}
</ul>

<p><strong>Action Required:</strong></p>
<p>Please review the attached document and return a signed copy by {signature_deadline}. If you have any questions or require any modifications, please let me know.</p>

<p>Thank you for your attention to this matter.</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company}<br>
{sender_phone}<br>
{sender_email}</p>
</body>
</html>""",
        "has_attachment": True,
        "supports_threading": True,
    },
    
    # Introduction/Networking
    {
        "name": "introduction_networking",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "Introduction - {sender_name} from {sender_company}",
            "Connecting with you",
            "Introduction via {referral_name}",
            "Nice to meet you at {event_name}",
            "Following up from {event_name}",
            "Reaching out - potential collaboration",
            "Introduction and partnership opportunity",
            "Connecting regarding {topic}",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Hi {recipient_first_name},</p>

<p>{intro_context}</p>

<p>{intro_about}</p>

<p>{intro_value_prop}</p>

<p>{intro_call_to_action}</p>

<p>I'd love to learn more about your work and explore potential synergies.</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company}<br>
{sender_phone}<br>
{sender_email}</p>
</body>
</html>""",
        "has_attachment": False,
        "supports_threading": True,
    },
    
    # Thank You/Appreciation
    {
        "name": "thank_you_appreciation",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "Thank you for your business",
            "Appreciation for our partnership",
            "Thank you - {project_name}",
            "Grateful for the opportunity",
            "Thank you for your trust",
            "Appreciation note",
            "Thank you for choosing {sender_company}",
            "Grateful for our collaboration",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Dear {recipient_first_name},</p>

<p>{thank_you_intro}</p>

<p>{thank_you_body}</p>

<p>{thank_you_future}</p>

<p>Please don't hesitate to reach out if there's anything we can do to support you.</p>

<p>With appreciation,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company}<br>
{sender_phone}<br>
{sender_email}</p>
</body>
</html>""",
        "has_attachment": False,
        "supports_threading": True,
    },
    
    # Service/Support
    {
        "name": "service_support",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "Re: Support ticket #{ticket_id}",
            "Your request has been received",
            "Update on your inquiry",
            "Resolution for ticket #{ticket_id}",
            "Support follow-up - {topic}",
            "Your case has been resolved",
            "Service update - {company_name}",
            "Response to your inquiry",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Hi {recipient_first_name},</p>

<p>Thank you for contacting {sender_company} support.</p>

<p><strong>Ticket Reference:</strong> #{ticket_id}</p>

<p>{support_update}</p>

<p><strong>Resolution/Next Steps:</strong></p>
<p>{support_resolution}</p>

<p>If you have any further questions or need additional assistance, please don't hesitate to reply to this email or contact our support team.</p>

<p>Thank you for your patience.</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company}<br>
Support: support@{sender_domain}</p>
</body>
</html>""",
        "has_attachment": False,
        "supports_threading": True,
    },
    
    # Event/Webinar Invitation
    {
        "name": "event_invitation",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "You're invited: {event_name}",
            "Exclusive invitation - {event_name}",
            "Join us for {event_name}",
            "Webinar invitation: {topic}",
            "Save the date - {event_name}",
            "Conference invitation - {event_name}",
            "Workshop invitation: {topic}",
            "Upcoming event - {event_name}",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Dear {recipient_first_name},</p>

<p>We're excited to invite you to {event_name}!</p>

<p><strong>Event Details:</strong></p>
<table style="border-collapse: collapse; margin: 15px 0;">
<tr><td style="padding: 5px 15px 5px 0;"><strong>Date:</strong></td><td>{event_date}</td></tr>
<tr><td style="padding: 5px 15px 5px 0;"><strong>Time:</strong></td><td>{event_time}</td></tr>
<tr><td style="padding: 5px 15px 5px 0;"><strong>Location:</strong></td><td>{event_location}</td></tr>
</table>

<p><strong>What You'll Learn:</strong></p>
<ul>
{event_highlights}
</ul>

<p>{event_speakers}</p>

<p><strong>Register Now:</strong></p>
<p>Space is limited. Please confirm your attendance by {registration_deadline}.</p>

<p>We look forward to seeing you there!</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company}<br>
{sender_email}</p>
</body>
</html>""",
        "has_attachment": False,
        "supports_threading": False,
    },
    
    # Product/Service Announcement
    {
        "name": "product_announcement",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "Introducing: {product_name}",
            "New feature announcement",
            "Exciting news from {sender_company}",
            "Product update - {product_name}",
            "New capabilities now available",
            "Important update from {sender_company}",
            "Announcing {product_name}",
            "What's new at {sender_company}",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Dear {recipient_first_name},</p>

<p>We're thrilled to share some exciting news with you!</p>

<p>{announcement_intro}</p>

<p><strong>Key Features:</strong></p>
<ul>
{announcement_features}
</ul>

<p><strong>Benefits for You:</strong></p>
<p>{announcement_benefits}</p>

<p>{announcement_cta}</p>

<p>If you have any questions, please don't hesitate to reach out.</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company}<br>
{sender_email}</p>
</body>
</html>""",
        "has_attachment": False,
        "supports_threading": False,
    },
    
    # Feedback Request
    {
        "name": "feedback_request",
        "category": "external_business",
        "sender_type": "external_business",
        "subject_templates": [
            "We'd love your feedback",
            "How was your experience?",
            "Quick survey - your opinion matters",
            "Feedback request - {project_name}",
            "Help us improve",
            "Your thoughts on our service",
            "Customer satisfaction survey",
            "Share your experience with us",
        ],
        "body_template": """<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
<p>Hi {recipient_first_name},</p>

<p>{feedback_intro}</p>

<p>Your feedback helps us improve our services and better serve customers like you.</p>

<p><strong>We'd appreciate your thoughts on:</strong></p>
<ul>
{feedback_areas}
</ul>

<p>{feedback_cta}</p>

<p>Thank you for taking the time to share your experience. Your input is invaluable to us.</p>

<p>Best regards,</p>
<p><strong>{sender_name}</strong><br>
{sender_title}<br>
{sender_company}<br>
{sender_email}</p>
</body>
</html>""",
        "has_attachment": False,
        "supports_threading": False,
    },
]

# Export all templates
__all__ = [
    "EXTERNAL_BUSINESS_TEMPLATES",
    "EXTERNAL_BUSINESS_DOMAINS",
    "EXTERNAL_SENDER_PROFILES",
]
