"""
Spam/Junk email templates for M365 email population.

Contains realistic spam templates that would typically end up in junk folder:
- Promotional offers
- Phishing attempts (safe/simulated)
- Newsletter spam
- Lottery/prize scams
- Fake invoices
"""

from typing import Dict, List, Any

# =============================================================================
# PROMOTIONAL SPAM
# =============================================================================

PROMOTIONAL_SPAM = {
    "category": "spam",
    "sensitivity": "normal",
    "sender_type": "external_spam",
    "department": None,
    "subject_templates": [
        "🔥 FLASH SALE - 90% OFF Everything! Limited Time!",
        "You Won't Believe These Deals! Act Now!",
        "💰 Exclusive Offer Just For You - Don't Miss Out!",
        "URGENT: Your Special Discount Expires Tonight!",
        "🎁 FREE Gift Inside - Open Now!",
        "Last Chance! Prices Slashed - Today Only!",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 0; padding: 0; background: #ff6b6b; }
        .container { max-width: 600px; margin: 0 auto; background: white; }
        .header { background: linear-gradient(135deg, #ff6b6b, #feca57); padding: 30px; text-align: center; color: white; }
        .header h1 { margin: 0; font-size: 32px; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); }
        .content { padding: 20px; text-align: center; }
        .big-button { display: inline-block; background: #ff6b6b; color: white; padding: 20px 40px; font-size: 24px; text-decoration: none; border-radius: 10px; margin: 20px 0; }
        .countdown { background: #1a1a2e; color: #00ff88; padding: 15px; font-size: 24px; font-family: monospace; }
        .price { font-size: 48px; color: #ff6b6b; font-weight: bold; }
        .original-price { text-decoration: line-through; color: #999; font-size: 24px; }
        .footer { background: #333; color: #999; padding: 15px; font-size: 10px; text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔥 MEGA SALE 🔥</h1>
            <p>The Biggest Discount Event of the Year!</p>
        </div>
        
        <div class="content">
            <div class="countdown">
                ⏰ OFFER EXPIRES IN: 23:59:59
            </div>
            
            <h2>EVERYTHING MUST GO!</h2>
            
            <p class="original-price">Was: $999.99</p>
            <p class="price">NOW: $49.99</p>
            
            <a href="#" class="big-button">CLAIM YOUR DISCOUNT NOW!</a>
            
            <p>✅ Free Shipping<br>
            ✅ 30-Day Returns<br>
            ✅ Satisfaction Guaranteed</p>
            
            <p><strong>Don't wait! This offer won't last!</strong></p>
        </div>
        
        <div class="footer">
            <p>You received this email because you signed up for our newsletter. 
            To unsubscribe, click <a href="#">here</a>.</p>
            <p>SuperDeals Inc. | 123 Marketing Street | Spam City, SP 12345</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# PHISHING ATTEMPT (SIMULATED - SAFE)
# =============================================================================

PHISHING_BANK = {
    "category": "spam",
    "sensitivity": "normal",
    "sender_type": "external_spam",
    "department": None,
    "subject_templates": [
        "⚠️ Urgent: Your Account Has Been Compromised",
        "Action Required: Verify Your Account Information",
        "Security Alert: Unusual Activity Detected",
        "Your Account Will Be Suspended - Verify Now",
        "Important: Update Your Payment Information",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; border: 1px solid #ddd; }
        .header { background: #003087; padding: 20px; text-align: center; }
        .header img { height: 40px; }
        .content { padding: 30px; }
        .alert-box { background: #fff3cd; border: 1px solid #ffc107; padding: 15px; margin: 20px 0; border-radius: 5px; }
        .button { display: inline-block; background: #003087; color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; margin: 20px 0; }
        .footer { background: #f5f5f5; padding: 20px; font-size: 11px; color: #666; text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2 style="color: white; margin: 0;">Security Center</h2>
        </div>
        
        <div class="content">
            <p>Dear Valued Customer,</p>
            
            <div class="alert-box">
                <strong>⚠️ Security Alert</strong><br>
                We detected unusual activity on your account. Your account access has been temporarily limited.
            </div>
            
            <p>We noticed the following suspicious activity:</p>
            <ul>
                <li>Login attempt from unrecognized device</li>
                <li>Location: {detection_location}</li>
                <li>Time: {detection_time}</li>
            </ul>
            
            <p>To restore full access to your account, please verify your information immediately:</p>
            
            <center>
                <a href="#" class="button">Verify My Account</a>
            </center>
            
            <p><strong>If you don't verify within 24 hours, your account will be permanently suspended.</strong></p>
            
            <p>Thank you for your cooperation.</p>
            
            <p>Security Team</p>
        </div>
        
        <div class="footer">
            <p>This is an automated message. Please do not reply.</p>
            <p>© 2024 Financial Services. All rights reserved.</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# LOTTERY/PRIZE SCAM
# =============================================================================

LOTTERY_SCAM = {
    "category": "spam",
    "sensitivity": "normal",
    "sender_type": "external_spam",
    "department": None,
    "subject_templates": [
        "🎉 CONGRATULATIONS! You've Won $1,000,000!",
        "You Are Our Lucky Winner! Claim Your Prize!",
        "WINNER NOTIFICATION: $500,000 Awaits You!",
        "🏆 You've Been Selected! Claim Your Reward!",
        "URGENT: Your Lottery Winnings Are Ready!",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: 'Times New Roman', serif; margin: 0; padding: 20px; background: #fffef0; }
        .container { max-width: 600px; margin: 0 auto; background: white; border: 3px solid gold; padding: 30px; }
        .header { text-align: center; border-bottom: 2px solid gold; padding-bottom: 20px; }
        .header h1 { color: #b8860b; margin: 0; }
        .content { padding: 20px 0; }
        .prize-box { background: linear-gradient(135deg, #ffd700, #ffec8b); padding: 20px; text-align: center; margin: 20px 0; border-radius: 10px; }
        .prize-amount { font-size: 48px; color: #006400; font-weight: bold; }
        .ref-number { background: #f0f0f0; padding: 10px; font-family: monospace; margin: 10px 0; }
        .footer { font-size: 10px; color: #999; margin-top: 30px; border-top: 1px solid #ddd; padding-top: 15px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🎰 INTERNATIONAL LOTTERY COMMISSION 🎰</h1>
            <p>Official Winner Notification</p>
        </div>
        
        <div class="content">
            <p>Dear Lucky Winner,</p>
            
            <p>We are pleased to inform you that your email address was selected in our annual international lottery draw!</p>
            
            <div class="prize-box">
                <p>YOUR WINNING AMOUNT:</p>
                <p class="prize-amount">$1,000,000.00</p>
                <p>ONE MILLION US DOLLARS</p>
            </div>
            
            <p class="ref-number">
                Reference Number: ILC/{ticket_id}<br>
                Batch Number: 2024/INTL/WIN
            </p>
            
            <p>To claim your prize, please contact our claims agent immediately:</p>
            
            <p>
                <strong>Claims Agent:</strong> Mr. James Williams<br>
                <strong>Email:</strong> claims@lottery-intl-winner.com<br>
                <strong>Phone:</strong> +44 700 123 4567
            </p>
            
            <p>Please provide the following information:</p>
            <ul>
                <li>Full Name</li>
                <li>Address</li>
                <li>Phone Number</li>
                <li>Reference Number (above)</li>
            </ul>
            
            <p><strong>Note:</strong> A processing fee of $500 is required to release your winnings.</p>
            
            <p>Congratulations once again!</p>
            
            <p>Yours faithfully,<br>
            <strong>Dr. Robert Johnson</strong><br>
            Director, International Lottery Commission</p>
        </div>
        
        <div class="footer">
            <p>This lottery is sponsored by major international corporations. All rights reserved.</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# FAKE INVOICE
# =============================================================================

FAKE_INVOICE = {
    "category": "spam",
    "sensitivity": "normal",
    "sender_type": "external_spam",
    "department": None,
    "subject_templates": [
        "Invoice #INV-{ticket_id} - Payment Due",
        "URGENT: Outstanding Invoice Requires Immediate Payment",
        "Your Invoice from Online Services - Due Today",
        "Payment Reminder: Invoice #{ticket_id}",
        "Final Notice: Invoice Payment Required",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
        .container { max-width: 600px; margin: 0 auto; background: white; }
        .header { background: #2c3e50; color: white; padding: 20px; }
        .header h1 { margin: 0; font-size: 24px; }
        .content { padding: 30px; }
        .invoice-details { background: #f9f9f9; padding: 20px; margin: 20px 0; }
        .invoice-table { width: 100%; border-collapse: collapse; }
        .invoice-table th, .invoice-table td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        .total { font-size: 24px; color: #e74c3c; font-weight: bold; }
        .button { display: inline-block; background: #27ae60; color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; }
        .warning { background: #fff3cd; border: 1px solid #ffc107; padding: 15px; margin: 20px 0; }
        .footer { font-size: 11px; color: #999; padding: 20px; border-top: 1px solid #ddd; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>INVOICE</h1>
        </div>
        
        <div class="content">
            <p>Dear Customer,</p>
            
            <p>Please find attached your invoice for recent services.</p>
            
            <div class="invoice-details">
                <table class="invoice-table">
                    <tr>
                        <th>Invoice Number:</th>
                        <td>INV-{ticket_id}</td>
                    </tr>
                    <tr>
                        <th>Date:</th>
                        <td>{current_date}</td>
                    </tr>
                    <tr>
                        <th>Due Date:</th>
                        <td>{deadline}</td>
                    </tr>
                </table>
            </div>
            
            <table class="invoice-table">
                <tr>
                    <th>Description</th>
                    <th>Amount</th>
                </tr>
                <tr>
                    <td>Annual Subscription Renewal</td>
                    <td>$299.99</td>
                </tr>
                <tr>
                    <td>Premium Support Package</td>
                    <td>$149.99</td>
                </tr>
                <tr>
                    <td>Processing Fee</td>
                    <td>$29.99</td>
                </tr>
                <tr>
                    <td><strong>TOTAL DUE:</strong></td>
                    <td class="total">$479.97</td>
                </tr>
            </table>
            
            <div class="warning">
                <strong>⚠️ Payment Required Within 24 Hours</strong><br>
                Failure to pay may result in service interruption and additional late fees.
            </div>
            
            <center>
                <a href="#" class="button">Pay Now</a>
            </center>
            
            <p>If you believe this invoice was sent in error, please contact our billing department.</p>
        </div>
        
        <div class="footer">
            <p>Online Services Inc. | Billing Department | billing@online-services-invoice.com</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# NEWSLETTER SPAM
# =============================================================================

NEWSLETTER_SPAM = {
    "category": "spam",
    "sensitivity": "normal",
    "sender_type": "external_spam",
    "department": None,
    "subject_templates": [
        "📧 Your Daily Digest - Top Stories You Can't Miss!",
        "Weekly Roundup: The News Everyone's Talking About",
        "🌟 Trending Now: Must-Read Articles Inside",
        "Don't Miss Out! This Week's Hottest Content",
        "Your Personalized News Feed - Updated!",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Georgia, serif; margin: 0; padding: 0; background: #f0f0f0; }
        .container { max-width: 600px; margin: 0 auto; background: white; }
        .header { background: #1a1a2e; color: white; padding: 30px; text-align: center; }
        .header h1 { margin: 0; font-size: 28px; }
        .content { padding: 20px; }
        .article { border-bottom: 1px solid #eee; padding: 20px 0; }
        .article h3 { margin: 0 0 10px 0; color: #1a1a2e; }
        .article p { color: #666; margin: 0; }
        .read-more { color: #e74c3c; text-decoration: none; font-weight: bold; }
        .ad-box { background: #fff3cd; padding: 20px; margin: 20px 0; text-align: center; border: 2px dashed #ffc107; }
        .footer { background: #1a1a2e; color: #999; padding: 20px; text-align: center; font-size: 11px; }
        .social { margin: 20px 0; }
        .social a { margin: 0 10px; color: white; text-decoration: none; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📰 DAILY DIGEST</h1>
            <p>Your personalized news roundup</p>
        </div>
        
        <div class="content">
            <div class="article">
                <h3>10 Secrets Successful People Never Tell You</h3>
                <p>Discover the hidden habits of millionaires that could change your life forever...</p>
                <a href="#" class="read-more">Read More →</a>
            </div>
            
            <div class="article">
                <h3>Scientists Discover Breakthrough That Will Shock You</h3>
                <p>New research reveals something incredible about everyday habits...</p>
                <a href="#" class="read-more">Read More →</a>
            </div>
            
            <div class="ad-box">
                <strong>🎁 SPONSORED</strong><br>
                Get 50% off your first order! Use code: NEWSLETTER50
            </div>
            
            <div class="article">
                <h3>Why Everyone Is Talking About This New Trend</h3>
                <p>The latest craze sweeping the nation - find out what it's all about...</p>
                <a href="#" class="read-more">Read More →</a>
            </div>
            
            <div class="article">
                <h3>You Won't Believe What Happened Next</h3>
                <p>An ordinary day turned extraordinary when this happened...</p>
                <a href="#" class="read-more">Read More →</a>
            </div>
        </div>
        
        <div class="footer">
            <div class="social">
                <a href="#">Facebook</a> | <a href="#">Twitter</a> | <a href="#">Instagram</a>
            </div>
            <p>You're receiving this because you subscribed to our newsletter.</p>
            <p><a href="#" style="color: #999;">Unsubscribe</a> | <a href="#" style="color: #999;">Update Preferences</a></p>
            <p>Daily Digest Media | 456 Content Ave | Newsletter City, NL 67890</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# EXPORT ALL SPAM TEMPLATES
# =============================================================================

SPAM_TEMPLATES = [
    PROMOTIONAL_SPAM,
    PHISHING_BANK,
    LOTTERY_SCAM,
    FAKE_INVOICE,
    NEWSLETTER_SPAM,
]

# Spam sender domains for realistic external addresses
SPAM_SENDER_DOMAINS = [
    "promo-deals-now.com",
    "super-savings-alert.net",
    "winner-notification.org",
    "secure-verify-account.com",
    "invoice-services-online.net",
    "daily-digest-news.com",
    "exclusive-offers-today.com",
    "urgent-action-required.net",
    "prize-claim-center.org",
    "billing-department-notice.com",
]

SPAM_SENDER_NAMES = [
    "Deals Team",
    "Security Alert",
    "Prize Department",
    "Billing Services",
    "Newsletter Team",
    "Customer Support",
    "Account Services",
    "Notification Center",
    "Rewards Program",
    "Special Offers",
]
