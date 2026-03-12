"""
Newsletter email templates for M365 email population.

Contains templates for:
- Company newsletters (internal)
- Industry newsletters (external)
"""

from typing import Dict, Any

COMPANY_NEWSLETTER: Dict[str, Any] = {
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

INDUSTRY_NEWSLETTER: Dict[str, Any] = {
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
            <p><a href="https://newsletters.{sender_domain}/unsubscribe?id={unsubscribe_id}">Unsubscribe</a> | <a href="https://newsletters.{sender_domain}/preferences?id={unsubscribe_id}">Manage Preferences</a> | <a href="https://newsletters.{sender_domain}/view/{newsletter_id}">View Online</a></p>
        </div>
    </div>
</body>
</html>""",
}

# Export list of newsletter templates
NEWSLETTER_TEMPLATES = [COMPANY_NEWSLETTER, INDUSTRY_NEWSLETTER]
