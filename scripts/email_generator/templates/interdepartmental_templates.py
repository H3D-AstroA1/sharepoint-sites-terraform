"""
Inter-departmental email templates for M365 email population.

Contains templates for:
- Project updates
- Meeting requests
- Status reports
- Collaboration requests
"""

from typing import Dict, Any

PROJECT_UPDATE: Dict[str, Any] = {
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

MEETING_REQUEST: Dict[str, Any] = {
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

STATUS_REPORT: Dict[str, Any] = {
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

COLLABORATION_REQUEST: Dict[str, Any] = {
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

# Export list of interdepartmental templates
INTERDEPARTMENTAL_TEMPLATES = [PROJECT_UPDATE, MEETING_REQUEST, STATUS_REPORT, COLLABORATION_REQUEST]

# Templates that support threading
THREADABLE_TEMPLATES = [PROJECT_UPDATE, MEETING_REQUEST, COLLABORATION_REQUEST]
