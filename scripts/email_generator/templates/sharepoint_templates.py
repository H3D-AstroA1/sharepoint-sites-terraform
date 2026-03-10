"""
SharePoint link email templates for M365 email population.

Contains templates for:
- Document sharing notifications
- Site activity summaries
"""

from typing import Dict, Any

SHAREPOINT_DOCUMENT_SHARED: Dict[str, Any] = {
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

SHAREPOINT_SITE_ACTIVITY: Dict[str, Any] = {
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

# Export list of SharePoint templates
SHAREPOINT_TEMPLATES = [SHAREPOINT_DOCUMENT_SHARED, SHAREPOINT_SITE_ACTIVITY]
