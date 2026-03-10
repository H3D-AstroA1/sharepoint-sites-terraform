"""
Attachment email templates for M365 email population.

Contains templates for:
- Reports with attachments
- Documents for review
"""

from typing import Dict, Any

REPORT_WITH_ATTACHMENT: Dict[str, Any] = {
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

DOCUMENT_FOR_REVIEW: Dict[str, Any] = {
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

# Export list of attachment templates
ATTACHMENT_TEMPLATES = [REPORT_WITH_ATTACHMENT, DOCUMENT_FOR_REVIEW]
