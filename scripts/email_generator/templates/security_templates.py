"""
Security-related email templates for M365 email population.

Contains templates for:
- Account blocked notifications
- Password reset emails (with temporary password)
- Password reset emails (with link only)
- Account unlock notifications
- Suspicious activity alerts
"""

from typing import Dict, List, Any

# =============================================================================
# ACCOUNT BLOCKED TEMPLATE
# =============================================================================

ACCOUNT_BLOCKED: Dict[str, Any] = {
    "category": "security",
    "sensitivity": "confidential",
    "sender_type": "internal_system",
    "department": "IT Security",
    "subject_templates": [
        "🔒 Account Security Alert: Your Account Has Been Blocked",
        "Security Notice: Account Access Suspended",
        "⚠️ Urgent: Your {company_name} Account Has Been Locked",
        "Account Security: Immediate Action Required",
        "🚫 Account Blocked - Security Violation Detected",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background: #f5f5f5; }}
        .container {{ max-width: 600px; margin: 0 auto; background: white; }}
        .header {{ background: linear-gradient(135deg, #d13438, #a4262c); color: white; padding: 30px; text-align: center; }}
        .header h1 {{ margin: 0; font-size: 24px; }}
        .header .icon {{ font-size: 48px; margin-bottom: 15px; }}
        .content {{ padding: 30px; }}
        .alert-box {{ background: #fde7e9; border: 1px solid #d13438; border-radius: 8px; padding: 20px; margin: 20px 0; }}
        .alert-box h3 {{ color: #d13438; margin-top: 0; }}
        .reason-box {{ background: #f3f2f1; padding: 20px; border-radius: 8px; margin: 20px 0; }}
        .action-steps {{ background: #e6f3ff; border-left: 4px solid #0078d4; padding: 20px; margin: 20px 0; }}
        .action-steps ol {{ margin: 10px 0; padding-left: 20px; }}
        .action-steps li {{ margin: 10px 0; }}
        .contact-box {{ background: #dff6dd; border: 1px solid #107c10; padding: 15px; border-radius: 8px; margin: 20px 0; }}
        .warning {{ color: #d13438; font-weight: bold; }}
        .footer {{ background: #f3f2f1; padding: 20px; text-align: center; font-size: 12px; color: #666; }}
        .security-badge {{ display: inline-block; background: #d13438; color: white; padding: 5px 15px; border-radius: 20px; font-size: 12px; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="icon">🔒</div>
            <h1>Account Security Alert</h1>
            <p>Your account has been temporarily blocked</p>
        </div>
        
        <div class="content">
            <p>Dear {recipient_first_name},</p>
            
            <div class="alert-box">
                <h3>⚠️ Your Account Has Been Blocked</h3>
                <p>Your {company_name} account (<strong>{recipient_email}</strong>) has been temporarily blocked due to a security concern.</p>
            </div>
            
            <div class="reason-box">
                <strong>📋 Reason for Block:</strong>
                <p>{block_reason}</p>
                <p><strong>Detected:</strong> {detection_time}</p>
                <p><strong>Location:</strong> {detection_location}</p>
            </div>
            
            <div class="action-steps">
                <strong>🔧 To Restore Access:</strong>
                <ol>
                    <li>Contact the IT Service Desk immediately</li>
                    <li>Verify your identity with security questions</li>
                    <li>Complete the account recovery process</li>
                    <li>Reset your password using a secure device</li>
                </ol>
            </div>
            
            <div class="contact-box">
                <strong>📞 IT Service Desk Contact:</strong><br>
                Email: <a href="mailto:servicedesk@{domain}">servicedesk@{domain}</a><br>
                Phone: {support_phone}<br>
                Hours: Monday - Friday, 8:00 AM - 6:00 PM
            </div>
            
            <p class="warning">⚠️ Do not attempt to log in repeatedly as this may extend the block duration.</p>
            
            <p>If you did not trigger this security alert, please contact IT Security immediately as your account may have been compromised.</p>
            
            <p>Best regards,<br>
            <strong>IT Security Team</strong><br>
            {company_name}</p>
        </div>
        
        <div class="footer">
            <span class="security-badge">SECURITY NOTICE</span>
            <p style="margin-top: 15px;">This is an automated security notification from {company_name}.<br>
            Please do not reply to this email.</p>
            <p>Incident Reference: {incident_id}</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# PASSWORD RESET WITH TEMPORARY PASSWORD
# =============================================================================

PASSWORD_RESET_WITH_TEMP: Dict[str, Any] = {
    "category": "security",
    "sensitivity": "highly_confidential",
    "sender_type": "internal_system",
    "department": "IT Security",
    "subject_templates": [
        "🔑 Your Temporary Password - Action Required",
        "Password Reset: Your New Temporary Password",
        "Account Recovery: Temporary Password Enclosed",
        "🔐 {company_name} Password Reset - Temporary Credentials",
        "Your Password Has Been Reset - Temporary Password Inside",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background: #f5f5f5; }}
        .container {{ max-width: 600px; margin: 0 auto; background: white; }}
        .header {{ background: linear-gradient(135deg, #0078d4, #00bcf2); color: white; padding: 30px; text-align: center; }}
        .header h1 {{ margin: 0; font-size: 24px; }}
        .header .icon {{ font-size: 48px; margin-bottom: 15px; }}
        .content {{ padding: 30px; }}
        .password-box {{ background: #1a1a2e; color: #00ff88; padding: 25px; border-radius: 8px; margin: 20px 0; text-align: center; font-family: 'Consolas', 'Courier New', monospace; }}
        .password-box .label {{ color: #aaa; font-size: 12px; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 2px; }}
        .password-box .password {{ font-size: 28px; letter-spacing: 3px; font-weight: bold; padding: 15px; background: #2a2a3e; border-radius: 5px; display: inline-block; }}
        .warning-box {{ background: #fff4ce; border: 1px solid #ffc107; border-radius: 8px; padding: 20px; margin: 20px 0; }}
        .warning-box h3 {{ color: #856404; margin-top: 0; }}
        .expiry-notice {{ background: #fde7e9; border-left: 4px solid #d13438; padding: 15px; margin: 20px 0; }}
        .steps {{ background: #f3f2f1; padding: 20px; border-radius: 8px; margin: 20px 0; }}
        .steps ol {{ margin: 10px 0; padding-left: 20px; }}
        .steps li {{ margin: 10px 0; }}
        .footer {{ background: #f3f2f1; padding: 20px; text-align: center; font-size: 12px; color: #666; }}
        .confidential-badge {{ display: inline-block; background: #d13438; color: white; padding: 5px 15px; border-radius: 20px; font-size: 12px; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="icon">🔑</div>
            <h1>Password Reset</h1>
            <p>Your temporary password is ready</p>
        </div>
        
        <div class="content">
            <p>Dear {recipient_first_name},</p>
            
            <p>Your password for <strong>{recipient_email}</strong> has been reset. Please use the temporary password below to log in:</p>
            
            <div class="password-box">
                <div class="label">Your Temporary Password</div>
                <div class="password">{temp_password}</div>
            </div>
            
            <div class="expiry-notice">
                <strong>⏰ This password expires in {expiry_hours} hours</strong><br>
                You must change your password immediately after logging in.
            </div>
            
            <div class="warning-box">
                <h3>⚠️ Important Security Notice</h3>
                <ul>
                    <li>This password is for one-time use only</li>
                    <li>Do not share this password with anyone</li>
                    <li>Delete this email after changing your password</li>
                    <li>IT will never ask for your password</li>
                </ul>
            </div>
            
            <div class="steps">
                <strong>📋 Next Steps:</strong>
                <ol>
                    <li>Go to <a href="https://login.{domain}">https://login.{domain}</a></li>
                    <li>Enter your email: <strong>{recipient_email}</strong></li>
                    <li>Enter the temporary password above</li>
                    <li>Create a new strong password (minimum 12 characters)</li>
                    <li>Complete any additional verification steps</li>
                </ol>
            </div>
            
            <p>If you did not request this password reset, please contact IT Security immediately at <a href="mailto:security@{domain}">security@{domain}</a>.</p>
            
            <p>Best regards,<br>
            <strong>IT Service Desk</strong><br>
            {company_name}</p>
        </div>
        
        <div class="footer">
            <span class="confidential-badge">HIGHLY CONFIDENTIAL</span>
            <p style="margin-top: 15px;">This email contains sensitive security information.<br>
            Do not forward or share this email.</p>
            <p>Request ID: {request_id}</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# PASSWORD RESET WITH LINK (NO PASSWORD)
# =============================================================================

PASSWORD_RESET_LINK: Dict[str, Any] = {
    "category": "security",
    "sensitivity": "confidential",
    "sender_type": "internal_system",
    "department": "IT Security",
    "subject_templates": [
        "🔗 Password Reset Link - Action Required",
        "Reset Your {company_name} Password",
        "Password Reset Request Received",
        "🔐 Complete Your Password Reset",
        "Action Required: Reset Your Password",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background: #f5f5f5; }}
        .container {{ max-width: 600px; margin: 0 auto; background: white; }}
        .header {{ background: linear-gradient(135deg, #0078d4, #00bcf2); color: white; padding: 30px; text-align: center; }}
        .header h1 {{ margin: 0; font-size: 24px; }}
        .header .icon {{ font-size: 48px; margin-bottom: 15px; }}
        .content {{ padding: 30px; }}
        .reset-button {{ display: block; background: linear-gradient(135deg, #0078d4, #00bcf2); color: white; text-decoration: none; padding: 18px 40px; border-radius: 8px; text-align: center; font-size: 18px; font-weight: bold; margin: 30px auto; max-width: 300px; }}
        .link-box {{ background: #f3f2f1; padding: 15px; border-radius: 8px; margin: 20px 0; word-break: break-all; font-family: 'Consolas', monospace; font-size: 12px; }}
        .expiry-notice {{ background: #fff4ce; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; }}
        .security-tips {{ background: #e6f3ff; padding: 20px; border-radius: 8px; margin: 20px 0; }}
        .security-tips ul {{ margin: 10px 0; padding-left: 20px; }}
        .footer {{ background: #f3f2f1; padding: 20px; text-align: center; font-size: 12px; color: #666; }}
        .security-badge {{ display: inline-block; background: #0078d4; color: white; padding: 5px 15px; border-radius: 20px; font-size: 12px; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="icon">🔗</div>
            <h1>Password Reset Request</h1>
            <p>Click below to reset your password</p>
        </div>
        
        <div class="content">
            <p>Dear {recipient_first_name},</p>
            
            <p>We received a request to reset the password for your {company_name} account (<strong>{recipient_email}</strong>).</p>
            
            <p>Click the button below to create a new password:</p>
            
            <a href="{reset_link}" class="reset-button">Reset My Password</a>
            
            <div class="expiry-notice">
                <strong>⏰ This link expires in {expiry_hours} hours</strong><br>
                For security reasons, this password reset link will expire on {expiry_date}.
            </div>
            
            <p>If the button doesn't work, copy and paste this link into your browser:</p>
            <div class="link-box">{reset_link}</div>
            
            <div class="security-tips">
                <strong>🔒 Password Security Tips:</strong>
                <ul>
                    <li>Use at least 12 characters</li>
                    <li>Include uppercase, lowercase, numbers, and symbols</li>
                    <li>Don't reuse passwords from other accounts</li>
                    <li>Consider using a password manager</li>
                </ul>
            </div>
            
            <p><strong>Didn't request this?</strong><br>
            If you didn't request a password reset, you can safely ignore this email. Your password will remain unchanged.</p>
            
            <p>If you're concerned about your account security, please contact IT Security at <a href="mailto:security@{domain}">security@{domain}</a>.</p>
            
            <p>Best regards,<br>
            <strong>IT Service Desk</strong><br>
            {company_name}</p>
        </div>
        
        <div class="footer">
            <span class="security-badge">SECURITY NOTICE</span>
            <p style="margin-top: 15px;">This is an automated message from {company_name}.<br>
            Please do not reply to this email.</p>
            <p>Request ID: {request_id}</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# ACCOUNT UNLOCKED NOTIFICATION
# =============================================================================

ACCOUNT_UNLOCKED: Dict[str, Any] = {
    "category": "security",
    "sensitivity": "internal",
    "sender_type": "internal_system",
    "department": "IT Security",
    "subject_templates": [
        "✅ Your Account Has Been Unlocked",
        "Account Access Restored - {company_name}",
        "🔓 Account Unlocked Successfully",
        "Good News: Your Account Is Now Active",
        "Account Recovery Complete",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background: #f5f5f5; }}
        .container {{ max-width: 600px; margin: 0 auto; background: white; }}
        .header {{ background: linear-gradient(135deg, #107c10, #00a651); color: white; padding: 30px; text-align: center; }}
        .header h1 {{ margin: 0; font-size: 24px; }}
        .header .icon {{ font-size: 48px; margin-bottom: 15px; }}
        .content {{ padding: 30px; }}
        .success-box {{ background: #dff6dd; border: 1px solid #107c10; border-radius: 8px; padding: 20px; margin: 20px 0; text-align: center; }}
        .success-box h3 {{ color: #107c10; margin-top: 0; }}
        .details-box {{ background: #f3f2f1; padding: 20px; border-radius: 8px; margin: 20px 0; }}
        .recommendations {{ background: #e6f3ff; border-left: 4px solid #0078d4; padding: 20px; margin: 20px 0; }}
        .recommendations ul {{ margin: 10px 0; padding-left: 20px; }}
        .login-button {{ display: block; background: linear-gradient(135deg, #107c10, #00a651); color: white; text-decoration: none; padding: 15px 30px; border-radius: 8px; text-align: center; font-size: 16px; font-weight: bold; margin: 20px auto; max-width: 200px; }}
        .footer {{ background: #f3f2f1; padding: 20px; text-align: center; font-size: 12px; color: #666; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="icon">✅</div>
            <h1>Account Unlocked</h1>
            <p>Your access has been restored</p>
        </div>
        
        <div class="content">
            <p>Dear {recipient_first_name},</p>
            
            <div class="success-box">
                <h3>🎉 Your Account Is Now Active</h3>
                <p>Your {company_name} account (<strong>{recipient_email}</strong>) has been successfully unlocked and is ready to use.</p>
            </div>
            
            <div class="details-box">
                <strong>📋 Unlock Details:</strong>
                <p><strong>Account:</strong> {recipient_email}</p>
                <p><strong>Unlocked:</strong> {unlock_time}</p>
                <p><strong>Unlocked by:</strong> {unlocked_by}</p>
                <p><strong>Reason:</strong> {unlock_reason}</p>
            </div>
            
            <a href="https://login.{domain}" class="login-button">Log In Now</a>
            
            <div class="recommendations">
                <strong>🔒 Security Recommendations:</strong>
                <ul>
                    <li>Consider changing your password if you haven't recently</li>
                    <li>Review your recent account activity</li>
                    <li>Enable multi-factor authentication if not already active</li>
                    <li>Report any suspicious activity to IT Security</li>
                </ul>
            </div>
            
            <p>If you continue to experience issues accessing your account, please contact the IT Service Desk.</p>
            
            <p>Best regards,<br>
            <strong>IT Service Desk</strong><br>
            {company_name}</p>
        </div>
        
        <div class="footer">
            <p>This is an automated notification from {company_name}.</p>
            <p>Ticket Reference: {ticket_id}</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# SUSPICIOUS ACTIVITY ALERT
# =============================================================================

SUSPICIOUS_ACTIVITY: Dict[str, Any] = {
    "category": "security",
    "sensitivity": "confidential",
    "sender_type": "internal_system",
    "department": "IT Security",
    "subject_templates": [
        "⚠️ Suspicious Activity Detected on Your Account",
        "Security Alert: Unusual Sign-In Activity",
        "🚨 Action Required: Verify Recent Account Activity",
        "Security Notice: Unrecognized Login Attempt",
        "Account Security Alert - Please Review",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background: #f5f5f5; }}
        .container {{ max-width: 600px; margin: 0 auto; background: white; }}
        .header {{ background: linear-gradient(135deg, #ff8c00, #ffc107); color: #333; padding: 30px; text-align: center; }}
        .header h1 {{ margin: 0; font-size: 24px; color: #333; }}
        .header .icon {{ font-size: 48px; margin-bottom: 15px; }}
        .content {{ padding: 30px; }}
        .alert-box {{ background: #fff4ce; border: 1px solid #ffc107; border-radius: 8px; padding: 20px; margin: 20px 0; }}
        .activity-details {{ background: #f3f2f1; padding: 20px; border-radius: 8px; margin: 20px 0; }}
        .activity-details table {{ width: 100%; border-collapse: collapse; }}
        .activity-details td {{ padding: 8px 0; border-bottom: 1px solid #ddd; }}
        .activity-details td:first-child {{ font-weight: bold; width: 40%; }}
        .action-buttons {{ text-align: center; margin: 30px 0; }}
        .btn {{ display: inline-block; padding: 12px 25px; border-radius: 5px; text-decoration: none; font-weight: bold; margin: 5px; }}
        .btn-yes {{ background: #107c10; color: white; }}
        .btn-no {{ background: #d13438; color: white; }}
        .security-tips {{ background: #e6f3ff; padding: 20px; border-radius: 8px; margin: 20px 0; }}
        .footer {{ background: #f3f2f1; padding: 20px; text-align: center; font-size: 12px; color: #666; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="icon">⚠️</div>
            <h1>Suspicious Activity Detected</h1>
            <p>Please review this sign-in attempt</p>
        </div>
        
        <div class="content">
            <p>Dear {recipient_first_name},</p>
            
            <div class="alert-box">
                <strong>We detected an unusual sign-in attempt on your account.</strong>
                <p>Please review the details below and let us know if this was you.</p>
            </div>
            
            <div class="activity-details">
                <strong>📋 Sign-In Details:</strong>
                <table>
                    <tr><td>Account:</td><td>{recipient_email}</td></tr>
                    <tr><td>Date & Time:</td><td>{signin_time}</td></tr>
                    <tr><td>Location:</td><td>{signin_location}</td></tr>
                    <tr><td>IP Address:</td><td>{ip_address}</td></tr>
                    <tr><td>Device:</td><td>{device_info}</td></tr>
                    <tr><td>Browser:</td><td>{browser_info}</td></tr>
                </table>
            </div>
            
            <div class="action-buttons">
                <p><strong>Was this you?</strong></p>
                <a href="{confirm_link}" class="btn btn-yes">✓ Yes, this was me</a>
                <a href="{deny_link}" class="btn btn-no">✗ No, secure my account</a>
            </div>
            
            <div class="security-tips">
                <strong>🔒 If this wasn't you:</strong>
                <ul>
                    <li>Click "No, secure my account" above</li>
                    <li>Change your password immediately</li>
                    <li>Review your recent account activity</li>
                    <li>Contact IT Security if you need assistance</li>
                </ul>
            </div>
            
            <p>If you have any concerns, please contact IT Security at <a href="mailto:security@{domain}">security@{domain}</a>.</p>
            
            <p>Best regards,<br>
            <strong>IT Security Team</strong><br>
            {company_name}</p>
        </div>
        
        <div class="footer">
            <p>This is an automated security notification from {company_name}.</p>
            <p>Alert ID: {alert_id}</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# COLLECTION OF ALL SECURITY TEMPLATES
# =============================================================================

SECURITY_TEMPLATES = [
    ACCOUNT_BLOCKED,
    PASSWORD_RESET_WITH_TEMP,
    PASSWORD_RESET_LINK,
    ACCOUNT_UNLOCKED,
    SUSPICIOUS_ACTIVITY,
]

# Templates with temporary passwords (for special handling)
PASSWORD_TEMPLATES = [PASSWORD_RESET_WITH_TEMP]
