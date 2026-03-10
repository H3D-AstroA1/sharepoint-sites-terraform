"""
Email templates for M365 email population.

Contains HTML templates and subject line templates for different email categories:
- Newsletters (company and industry)
- SharePoint links and notifications
- Emails with attachments
- Organisational communications
- Inter-departmental emails
- Security notifications (account blocked, password reset, MFA)
"""

from typing import Dict, List, Any

# =============================================================================
# NEWSLETTER TEMPLATES
# =============================================================================

COMPANY_NEWSLETTER = {
    "category": "newsletters",
    "sensitivity": "general",
    "sender_type": "internal_system",
    "subject_templates": [
        "{company_name} Weekly Update - Week {week_number}",
        "{company_name} Monthly Newsletter - {month} {year}",
        "📢 Company News & Updates - {date}",
        "This Week at {company_name}",
        "{company_name} Digest - {month} {year}",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .header {{ background: linear-gradient(135deg, #0078d4, #00bcf2); color: white; padding: 30px; text-align: center; }}
        .content {{ padding: 30px; max-width: 600px; margin: 0 auto; }}
        .section {{ margin-bottom: 25px; }}
        .section-title {{ color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 5px; margin-bottom: 15px; }}
        .highlight-box {{ background: #f3f2f1; padding: 15px; border-left: 4px solid #0078d4; margin: 15px 0; }}
        .footer {{ background: #f3f2f1; padding: 20px; text-align: center; font-size: 12px; color: #666; }}
        a {{ color: #0078d4; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📰 {company_name} Newsletter</h1>
        <p>{date}</p>
    </div>
    
    <div class="content">
        <p>Dear {recipient_first_name},</p>
        
        <p>Welcome to this week's company newsletter! Here's what's happening across the organisation.</p>
        
        <div class="section">
            <h2 class="section-title">📢 Company News</h2>
            {company_news}
        </div>
        
        <div class="section">
            <h2 class="section-title">📅 Upcoming Events</h2>
            {upcoming_events}
        </div>
        
        <div class="section">
            <h2 class="section-title">🏆 Employee Spotlight</h2>
            {employee_spotlight}
        </div>
        
        <div class="highlight-box">
            <strong>💡 Did You Know?</strong><br>
            {fun_fact}
        </div>
        
        <div class="section">
            <h2 class="section-title">📊 Department Updates</h2>
            {department_updates}
        </div>
        
        <p>Have news to share? Contact <a href="mailto:communications@{domain}">Corporate Communications</a>.</p>
        
        <p>Best regards,<br>
        <strong>Corporate Communications Team</strong></p>
    </div>
    
    <div class="footer">
        <p>{company_name} | Internal Communications</p>
        <p>This email is intended for internal use only.</p>
    </div>
</body>
</html>""",
}

INDUSTRY_NEWSLETTER = {
    "category": "newsletters",
    "sensitivity": "general",
    "sender_type": "external",
    "subject_templates": [
        "{industry} Weekly Digest - {date}",
        "Your {industry} News Roundup",
        "📈 {industry} Insights - {month} {year}",
        "Top {industry} Stories This Week",
        "{newsletter_name}: {month} Edition",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; background: #f5f5f5; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; background: white; }}
        .header {{ background: #1a1a2e; color: white; padding: 25px; text-align: center; }}
        .content {{ padding: 25px; }}
        .article {{ border-bottom: 1px solid #eee; padding: 15px 0; }}
        .article:last-child {{ border-bottom: none; }}
        .article-title {{ color: #1a1a2e; font-size: 18px; margin-bottom: 8px; font-weight: bold; }}
        .article-meta {{ color: #666; font-size: 12px; margin-bottom: 10px; }}
        .article-summary {{ color: #444; }}
        .read-more {{ color: #0066cc; text-decoration: none; font-weight: bold; }}
        .footer {{ background: #f5f5f5; padding: 20px; text-align: center; font-size: 11px; color: #666; }}
        .footer a {{ color: #0066cc; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>{newsletter_name}</h1>
            <p>{tagline}</p>
        </div>
        
        <div class="content">
            <p>Hi {recipient_first_name},</p>
            <p>Here's your weekly roundup of the latest {industry} news and insights.</p>
            
            {articles}
            
            <p style="margin-top: 25px;">Stay informed,<br>
            <strong>The {newsletter_name} Team</strong></p>
        </div>
        
        <div class="footer">
            <p>You're receiving this because you subscribed to {newsletter_name}.</p>
            <p><a href="#">Unsubscribe</a> | <a href="#">Manage Preferences</a> | <a href="#">View Online</a></p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# SHAREPOINT LINK TEMPLATES
# =============================================================================

SHAREPOINT_DOCUMENT_SHARED = {
    "category": "links",
    "sensitivity": "internal",
    "sender_type": "internal_users",
    "subject_templates": [
        "New Document: {document_name}",
        "📄 Document Shared: {document_name}",
        "Please Review: {document_name} on {site_name}",
        "Updated: {document_name} - Action Required",
        "{sender_name} shared a document with you",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .document-card {{ background: #f3f2f1; border-radius: 8px; padding: 20px; margin: 20px 0; }}
        .document-icon {{ font-size: 48px; margin-bottom: 10px; }}
        .document-name {{ font-size: 18px; font-weight: bold; color: #0078d4; }}
        .document-meta {{ color: #666; font-size: 14px; margin: 10px 0; }}
        .btn {{ display: inline-block; background: #0078d4; color: white; padding: 12px 24px; 
               text-decoration: none; border-radius: 4px; margin-top: 15px; }}
        .btn:hover {{ background: #106ebe; }}
        .message-box {{ background: #fff4ce; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; }}
        .footer {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; font-size: 12px; color: #666; }}
    </style>
</head>
<body>
    <div class="container">
        <p>Hi {recipient_first_name},</p>
        
        <p>{sender_name} has shared a document with you from the <strong>{site_name}</strong> SharePoint site.</p>
        
        <div class="document-card">
            <div class="document-icon">{document_icon}</div>
            <div class="document-name">{document_name}</div>
            <div class="document-meta">
                📁 Location: {site_name} / {folder_path}<br>
                👤 Shared by: {sender_name}<br>
                📅 Date: {share_date}<br>
                📊 Size: {file_size}
            </div>
            <a href="{sharepoint_url}" class="btn">Open Document</a>
        </div>
        
        <div class="message-box">
            <strong>Message from {sender_first_name}:</strong><br>
            {personal_message}
        </div>
        
        {action_required}
        
        <p>Best regards,<br>
        {sender_name}<br>
        <span style="color: #666;">{sender_title} | {sender_department}</span></p>
        
        <div class="footer">
            <p>This is a notification from SharePoint.</p>
            <p><a href="{site_url}">Visit {site_name}</a></p>
        </div>
    </div>
</body>
</html>""",
}

SHAREPOINT_SITE_ACTIVITY = {
    "category": "links",
    "sensitivity": "general",
    "sender_type": "internal_system",
    "subject_templates": [
        "Activity on {site_name}",
        "📊 Weekly Activity Summary - {site_name}",
        "New content on {site_name}",
        "{site_name} - Recent Updates",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .header {{ background: #0078d4; color: white; padding: 20px; border-radius: 8px 8px 0 0; }}
        .content {{ background: #f9f9f9; padding: 20px; border-radius: 0 0 8px 8px; }}
        .activity-item {{ background: white; padding: 15px; margin: 10px 0; border-radius: 4px; border-left: 4px solid #0078d4; }}
        .activity-icon {{ font-size: 20px; margin-right: 10px; }}
        .activity-text {{ color: #333; }}
        .activity-meta {{ color: #666; font-size: 12px; margin-top: 5px; }}
        .btn {{ display: inline-block; background: #0078d4; color: white; padding: 10px 20px; 
               text-decoration: none; border-radius: 4px; margin-top: 15px; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>📊 {site_name}</h2>
            <p>Activity Summary</p>
        </div>
        
        <div class="content">
            <p>Hi {recipient_first_name},</p>
            <p>Here's what's been happening on the {site_name} SharePoint site:</p>
            
            {activity_items}
            
            <a href="{site_url}" class="btn">View Site</a>
            
            <p style="margin-top: 20px; font-size: 12px; color: #666;">
                You're receiving this because you follow {site_name}.
            </p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# ATTACHMENT EMAIL TEMPLATES
# =============================================================================

REPORT_WITH_ATTACHMENT = {
    "category": "attachments",
    "sensitivity": "confidential",
    "sender_type": "internal_users",
    "has_attachment": True,
    "subject_templates": [
        "{report_type} Report - {period}",
        "📊 {report_type} for Your Review",
        "{department} {report_type} - {date}",
        "Attached: {document_name}",
        "Q{quarter} {report_type} - Please Review",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .attachment-box {{ background: #fff4ce; border: 1px solid #ffc107; border-radius: 8px; 
                          padding: 15px; margin: 20px 0; display: flex; align-items: center; }}
        .attachment-icon {{ font-size: 36px; margin-right: 15px; }}
        .attachment-info {{ flex: 1; }}
        .attachment-name {{ font-weight: bold; color: #333; }}
        .attachment-size {{ color: #666; font-size: 12px; }}
        .key-points {{ background: #f3f2f1; padding: 15px; border-radius: 8px; margin: 20px 0; }}
        .key-points ul {{ margin: 10px 0; padding-left: 20px; }}
        .deadline {{ background: #fde7e9; border-left: 4px solid #d13438; padding: 10px 15px; margin: 20px 0; }}
        .signature {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; }}
    </style>
</head>
<body>
    <div class="container">
        <p>Dear {recipient_name},</p>
        
        <p>Please find attached the <strong>{report_type}</strong> for {period}.</p>
        
        <div class="attachment-box">
            <div class="attachment-icon">{file_icon}</div>
            <div class="attachment-info">
                <div class="attachment-name">📎 {attachment_name}</div>
                <div class="attachment-size">{file_type} • {file_size}</div>
            </div>
        </div>
        
        <div class="key-points">
            <strong>📋 Key Highlights:</strong>
            <ul>
                {key_points}
            </ul>
        </div>
        
        {executive_summary}
        
        <div class="deadline">
            <strong>⏰ Review Deadline:</strong> {deadline}<br>
            <span style="font-size: 13px;">Please provide your feedback by this date.</span>
        </div>
        
        <p>If you have any questions or need clarification on any items, please don't hesitate to reach out.</p>
        
        <div class="signature">
            <p>Best regards,</p>
            <p><strong>{sender_name}</strong><br>
            {sender_title}<br>
            {sender_department}</p>
            <p style="font-size: 12px; color: #666;">
                📞 {sender_phone}<br>
                ✉️ {sender_email}
            </p>
        </div>
    </div>
</body>
</html>""",
}

DOCUMENT_FOR_REVIEW = {
    "category": "attachments",
    "sensitivity": "internal",
    "sender_type": "internal_users",
    "has_attachment": True,
    "subject_templates": [
        "For Review: {document_name}",
        "📝 Please Review: {document_name}",
        "Draft {document_type} - Feedback Requested",
        "{document_name} - Your Input Needed",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .attachment-card {{ background: #e6f3ff; border: 1px solid #0078d4; border-radius: 8px; padding: 15px; margin: 20px 0; }}
        .context-box {{ background: #f3f2f1; padding: 15px; border-radius: 8px; margin: 20px 0; }}
        .feedback-request {{ background: #dff6dd; border-left: 4px solid #107c10; padding: 15px; margin: 20px 0; }}
    </style>
</head>
<body>
    <div class="container">
        <p>Hi {recipient_first_name},</p>
        
        <p>I'd appreciate your feedback on the attached {document_type}.</p>
        
        <div class="attachment-card">
            <strong>📎 {attachment_name}</strong><br>
            <span style="color: #666; font-size: 13px;">{file_type} • {file_size}</span>
        </div>
        
        <div class="context-box">
            <strong>📋 Context:</strong><br>
            {context}
        </div>
        
        <div class="feedback-request">
            <strong>🎯 Specifically, I'd like your input on:</strong>
            <ul>
                {feedback_areas}
            </ul>
        </div>
        
        <p>Please share your thoughts by <strong>{deadline}</strong> if possible.</p>
        
        <p>Thanks in advance!</p>
        
        <p>{sender_first_name}</p>
    </div>
</body>
</html>""",
}

# =============================================================================
# ORGANISATIONAL COMMUNICATION TEMPLATES
# =============================================================================

COMPANY_ANNOUNCEMENT = {
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

HR_POLICY_UPDATE = {
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

LEADERSHIP_MESSAGE = {
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

# =============================================================================
# INTER-DEPARTMENTAL EMAIL TEMPLATES
# =============================================================================

PROJECT_UPDATE = {
    "category": "interdepartmental",
    "sensitivity": "internal",
    "sender_type": "internal_users",
    "supports_threading": True,
    "subject_templates": [
        "Re: {project_name} - Status Update",
        "{project_name} - Weekly Progress",
        "Update: {project_name} - {milestone}",
        "Re: {project_name} Discussion",
        "{project_name} - {update_type}",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .status-section {{ margin: 20px 0; }}
        .status-header {{ font-weight: bold; margin-bottom: 10px; padding: 8px; border-radius: 4px; }}
        .completed {{ background: #dff6dd; color: #107c10; }}
        .in-progress {{ background: #e6f3ff; color: #0078d4; }}
        .upcoming {{ background: #f3e6ff; color: #8764b8; }}
        .status-item {{ padding: 5px 0 5px 20px; position: relative; }}
        .status-item::before {{ content: '•'; position: absolute; left: 5px; }}
        .next-steps {{ background: #f3f2f1; padding: 15px; border-radius: 8px; margin: 20px 0; }}
        .quoted-email {{ border-left: 3px solid #ccc; padding-left: 15px; margin: 20px 0; 
                        color: #666; font-size: 13px; }}
    </style>
</head>
<body>
    <div class="container">
        <p>Hi {recipient_first_name},</p>
        
        <p>{context_message}</p>
        
        <p>Here's the latest update on <strong>{project_name}</strong>:</p>
        
        <div class="status-section">
            <div class="status-header completed">✅ Completed</div>
            {completed_items}
        </div>
        
        <div class="status-section">
            <div class="status-header in-progress">🔄 In Progress</div>
            {in_progress_items}
        </div>
        
        <div class="status-section">
            <div class="status-header upcoming">⏳ Upcoming</div>
            {upcoming_items}
        </div>
        
        <div class="next-steps">
            <strong>📋 Next Steps:</strong>
            <ol>
                {next_steps}
            </ol>
        </div>
        
        <p>Let me know if you have any questions or need additional information.</p>
        
        <p>Cheers,<br>
        {sender_first_name}</p>
        
        {quoted_thread}
    </div>
</body>
</html>""",
}

MEETING_REQUEST = {
    "category": "interdepartmental",
    "sensitivity": "general",
    "sender_type": "internal_users",
    "subject_templates": [
        "Meeting Request: {meeting_topic}",
        "Can we schedule a call about {meeting_topic}?",
        "📅 {meeting_topic} - Meeting Invite",
        "Quick sync on {meeting_topic}?",
        "Let's discuss: {meeting_topic}",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .meeting-card {{ background: #e6f3ff; border: 1px solid #0078d4; border-radius: 8px; padding: 20px; margin: 20px 0; }}
        .meeting-details {{ margin: 15px 0; }}
        .detail-row {{ display: flex; margin: 8px 0; }}
        .detail-label {{ font-weight: bold; width: 100px; }}
        .agenda {{ background: #f3f2f1; padding: 15px; border-radius: 8px; margin: 20px 0; }}
    </style>
</head>
<body>
    <div class="container">
        <p>Hi {recipient_first_name},</p>
        
        <p>{meeting_context}</p>
        
        <div class="meeting-card">
            <h3>📅 Meeting Request</h3>
            <div class="meeting-details">
                <div class="detail-row">
                    <span class="detail-label">Topic:</span>
                    <span>{meeting_topic}</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Duration:</span>
                    <span>{duration}</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Proposed:</span>
                    <span>{proposed_times}</span>
                </div>
            </div>
        </div>
        
        <div class="agenda">
            <strong>📋 Agenda:</strong>
            <ul>
                {agenda_items}
            </ul>
        </div>
        
        <p>Please let me know which time works best for you, or suggest an alternative.</p>
        
        <p>Thanks,<br>
        {sender_first_name}</p>
    </div>
</body>
</html>""",
}

STATUS_REPORT = {
    "category": "interdepartmental",
    "sensitivity": "internal",
    "sender_type": "internal_users",
    "subject_templates": [
        "Weekly Status Report - {department}",
        "{department} Update - Week of {date}",
        "Status Update: {project_name}",
        "📊 Weekly Summary - {department}",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .header {{ background: #0078d4; color: white; padding: 15px; border-radius: 8px 8px 0 0; }}
        .content {{ background: #f9f9f9; padding: 20px; border-radius: 0 0 8px 8px; }}
        .metric {{ display: inline-block; background: white; padding: 15px; margin: 5px; border-radius: 8px; text-align: center; min-width: 100px; }}
        .metric-value {{ font-size: 24px; font-weight: bold; color: #0078d4; }}
        .metric-label {{ font-size: 12px; color: #666; }}
        .section {{ margin: 20px 0; }}
        .section-title {{ font-weight: bold; color: #0078d4; margin-bottom: 10px; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>📊 Weekly Status Report</h2>
            <p>{department} - Week of {date}</p>
        </div>
        
        <div class="content">
            <p>Hi {recipient_first_name},</p>
            
            <p>Here's the weekly status update from {department}:</p>
            
            <div style="text-align: center; margin: 20px 0;">
                {metrics}
            </div>
            
            <div class="section">
                <div class="section-title">✅ Accomplishments</div>
                <ul>{accomplishments}</ul>
            </div>
            
            <div class="section">
                <div class="section-title">🎯 Focus Areas</div>
                <ul>{focus_areas}</ul>
            </div>
            
            <div class="section">
                <div class="section-title">⚠️ Blockers/Risks</div>
                <ul>{blockers}</ul>
            </div>
            
            <p>Let me know if you have any questions.</p>
            
            <p>Best,<br>
            {sender_first_name}</p>
        </div>
    </div>
</body>
</html>""",
}

COLLABORATION_REQUEST = {
    "category": "interdepartmental",
    "sensitivity": "general",
    "sender_type": "internal_users",
    "subject_templates": [
        "Collaboration Request: {topic}",
        "Need your input on {topic}",
        "Cross-team collaboration: {topic}",
        "Request for assistance: {topic}",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .request-box {{ background: #fff4ce; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; }}
        .details {{ background: #f3f2f1; padding: 15px; border-radius: 8px; margin: 20px 0; }}
    </style>
</head>
<body>
    <div class="container">
        <p>Hi {recipient_first_name},</p>
        
        <p>I hope this email finds you well. I'm reaching out because {request_context}.</p>
        
        <div class="request-box">
            <strong>🤝 Request:</strong><br>
            {request_details}
        </div>
        
        <div class="details">
            <strong>📋 Background:</strong><br>
            {background}
        </div>
        
        <p><strong>Timeline:</strong> {timeline}</p>
        
        <p>Would you be available to discuss this? I'm happy to set up a quick call at your convenience.</p>
        
        <p>Thanks in advance for your help!</p>
        
        <p>Best regards,<br>
        {sender_name}<br>
        <span style="color: #666;">{sender_title} | {sender_department}</span></p>
    </div>
</body>
</html>""",
}

# =============================================================================
# TEMPLATE COLLECTIONS
# =============================================================================

# Import security templates
from .security_templates import (
    ACCOUNT_BLOCKED,
    PASSWORD_RESET_WITH_TEMP,
    PASSWORD_RESET_LINK,
    ACCOUNT_UNLOCKED,
    SUSPICIOUS_ACTIVITY,
    SECURITY_TEMPLATES,
)

# Import spam templates
from .spam_templates import (
    SPAM_TEMPLATES,
    SPAM_SENDER_DOMAINS,
    SPAM_SENDER_NAMES,
)

# All templates organized by category
EMAIL_TEMPLATES = {
    "newsletters": [COMPANY_NEWSLETTER, INDUSTRY_NEWSLETTER],
    "links": [SHAREPOINT_DOCUMENT_SHARED, SHAREPOINT_SITE_ACTIVITY],
    "attachments": [REPORT_WITH_ATTACHMENT, DOCUMENT_FOR_REVIEW],
    "organisational": [COMPANY_ANNOUNCEMENT, HR_POLICY_UPDATE, LEADERSHIP_MESSAGE],
    "interdepartmental": [PROJECT_UPDATE, MEETING_REQUEST, STATUS_REPORT, COLLABORATION_REQUEST],
    "security": SECURITY_TEMPLATES,
    "spam": SPAM_TEMPLATES,
}

# Templates that support threading
THREADABLE_TEMPLATES = [PROJECT_UPDATE, MEETING_REQUEST, COLLABORATION_REQUEST]

# Templates that require attachments
ATTACHMENT_TEMPLATES = [REPORT_WITH_ATTACHMENT, DOCUMENT_FOR_REVIEW]

# Security templates with temporary passwords
PASSWORD_TEMPLATES = [PASSWORD_RESET_WITH_TEMP]