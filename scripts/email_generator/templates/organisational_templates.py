"""
Organisational communication email templates for M365 email population.

Contains templates for:
- Company announcements
- HR policy updates
- Leadership messages
"""

from typing import Dict, Any

COMPANY_ANNOUNCEMENT: Dict[str, Any] = {
    "category": "organisational",
    "sensitivity": "internal",
    "sender_type": "internal_system",
    "subject_templates": [
        "📢 Important Announcement: {announcement_title}",
        "Company Update: {announcement_title}",
        "All Staff: {announcement_title}",
        "Message from {sender_role}: {announcement_title}",
        "🔔 {announcement_type}: {announcement_title}",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; }}
        .banner {{ background: linear-gradient(135deg, #0078d4, #00bcf2); color: white; 
                  padding: 30px; text-align: center; }}
        .banner h1 {{ margin: 0; font-size: 24px; }}
        .banner .type {{ opacity: 0.9; font-size: 14px; margin-top: 5px; }}
        .content {{ padding: 30px; }}
        .highlight {{ background: #e6f3ff; border-left: 4px solid #0078d4; padding: 15px; margin: 20px 0; }}
        .action-items {{ background: #f3f2f1; padding: 20px; border-radius: 8px; margin: 20px 0; }}
        .effective-date {{ background: #dff6dd; border: 1px solid #107c10; padding: 15px; 
                          border-radius: 8px; margin: 20px 0; text-align: center; }}
        .contact-box {{ background: #f3f2f1; padding: 15px; border-radius: 8px; margin: 20px 0; }}
        .footer {{ padding: 20px; background: #f3f2f1; text-align: center; font-size: 12px; }}
        .sensitivity-label {{ display: inline-block; background: #ffc107; color: #333; 
                             padding: 3px 10px; border-radius: 3px; font-size: 11px; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="banner">
            <div class="type">{announcement_type}</div>
            <h1>{announcement_title}</h1>
        </div>
        
        <div class="content">
            <p>Dear Colleagues,</p>
            
            <p>{greeting}</p>
            
            <div class="highlight">
                {main_announcement}
            </div>
            
            {details}
            
            <div class="action-items">
                <strong>📋 What This Means For You:</strong>
                <ul>
                    {action_items}
                </ul>
            </div>
            
            <div class="effective-date">
                <strong>📅 Effective Date:</strong> {effective_date}
            </div>
            
            <div class="contact-box">
                <strong>❓ Questions?</strong><br>
                {contact_info}
            </div>
            
            <p>Thank you for your attention to this matter.</p>
            
            <p>Best regards,</p>
            <p><strong>{sender_name}</strong><br>
            {sender_title}</p>
        </div>
        
        <div class="footer">
            <span class="sensitivity-label">SENSITIVITY: {sensitivity_label}</span>
            <p style="margin-top: 10px;">{company_name} | Internal Communication</p>
        </div>
    </div>
</body>
</html>""",
}

HR_POLICY_UPDATE: Dict[str, Any] = {
    "category": "organisational",
    "sensitivity": "internal",
    "sender_type": "internal_system",
    "department": "Human Resources",
    "subject_templates": [
        "HR Policy Update: {policy_name}",
        "📋 Updated Policy: {policy_name}",
        "Important: Changes to {policy_name}",
        "HR Notice: {policy_name} - Effective {date}",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .header {{ background: #5c2d91; color: white; padding: 20px; border-radius: 8px 8px 0 0; }}
        .content {{ background: #f9f9f9; padding: 20px; border-radius: 0 0 8px 8px; }}
        .policy-box {{ background: white; border: 1px solid #ddd; padding: 20px; margin: 20px 0; border-radius: 8px; }}
        .changes {{ background: #fff4ce; padding: 15px; border-radius: 8px; margin: 20px 0; }}
        .action-required {{ background: #fde7e9; border-left: 4px solid #d13438; padding: 15px; margin: 20px 0; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>📋 HR Policy Update</h2>
            <p>{policy_name}</p>
        </div>
        
        <div class="content">
            <p>Dear {recipient_first_name},</p>
            
            <p>We are writing to inform you of updates to the <strong>{policy_name}</strong>.</p>
            
            <div class="policy-box">
                <h3>Policy Overview</h3>
                {policy_overview}
            </div>
            
            <div class="changes">
                <strong>📝 Key Changes:</strong>
                <ul>
                    {key_changes}
                </ul>
            </div>
            
            <div class="action-required">
                <strong>⚠️ Action Required:</strong><br>
                {action_required}
            </div>
            
            <p>The updated policy is effective from <strong>{effective_date}</strong>.</p>
            
            <p>For questions, please contact HR at <a href="mailto:hr@{domain}">hr@{domain}</a>.</p>
            
            <p>Best regards,<br>
            <strong>Human Resources Team</strong></p>
        </div>
    </div>
</body>
</html>""",
}

LEADERSHIP_MESSAGE: Dict[str, Any] = {
    "category": "organisational",
    "sensitivity": "internal",
    "sender_type": "internal_users",
    "subject_templates": [
        "A Message from {sender_role}",
        "{sender_name}: {message_topic}",
        "From the {sender_role}'s Desk: {message_topic}",
        "Leadership Update: {message_topic}",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.8; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 30px; }}
        .greeting {{ font-size: 18px; margin-bottom: 20px; }}
        .message {{ margin: 20px 0; }}
        .signature {{ margin-top: 40px; padding-top: 20px; border-top: 2px solid #0078d4; }}
        .signature-name {{ font-size: 18px; font-weight: bold; color: #0078d4; }}
    </style>
</head>
<body>
    <div class="container">
        <p class="greeting">Dear Team,</p>
        
        <div class="message">
            {message_body}
        </div>
        
        <div class="signature">
            <p class="signature-name">{sender_name}</p>
            <p style="color: #666; margin: 5px 0;">{sender_title}</p>
            <p style="color: #666; margin: 5px 0;">{company_name}</p>
        </div>
    </div>
</body>
</html>""",
}

# Export list of organisational templates
ORGANISATIONAL_TEMPLATES = [COMPANY_ANNOUNCEMENT, HR_POLICY_UPDATE, LEADERSHIP_MESSAGE]
