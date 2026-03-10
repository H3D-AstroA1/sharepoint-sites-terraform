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

PROMOTIONAL_SPAM: Dict[str, Any] = {
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
        .content { padding: 20px; text-align: center; }
        .big-button { display: inline-block; background: #ff6b6b; color: white; padding: 20px 40px; font-size: 24px; text-decoration: none; border-radius: 10px; margin: 20px 0; }
        .price { font-size: 48px; color: #ff6b6b; font-weight: bold; }
        .footer { background: #333; color: #999; padding: 15px; font-size: 10px; text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>MEGA SALE</h1>
            <p>The Biggest Discount Event of the Year!</p>
        </div>
        <div class="content">
            <h2>EVERYTHING MUST GO!</h2>
            <p class="price">NOW: $49.99</p>
            <a href="#" class="big-button">CLAIM YOUR DISCOUNT NOW!</a>
        </div>
        <div class="footer">
            <p>To unsubscribe, click here.</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# PHISHING ATTEMPT (SIMULATED - SAFE)
# =============================================================================

PHISHING_BANK: Dict[str, Any] = {
    "category": "spam",
    "sensitivity": "normal",
    "sender_type": "external_spam",
    "department": None,
    "subject_templates": [
        "Urgent: Your Account Has Been Compromised",
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
        .content { padding: 30px; }
        .alert-box { background: #fff3cd; border: 1px solid #ffc107; padding: 15px; margin: 20px 0; }
        .button { display: inline-block; background: #003087; color: white; padding: 15px 30px; text-decoration: none; }
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
                <strong>Security Alert</strong><br>
                We detected unusual activity on your account.
            </div>
            <p>To restore full access, please verify your information:</p>
            <center><a href="#" class="button">Verify My Account</a></center>
            <p>Security Team</p>
        </div>
        <div class="footer">
            <p>This is an automated message.</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# LOTTERY/PRIZE SCAM
# =============================================================================

LOTTERY_SCAM: Dict[str, Any] = {
    "category": "spam",
    "sensitivity": "normal",
    "sender_type": "external_spam",
    "department": None,
    "subject_templates": [
        "CONGRATULATIONS! You've Won $1,000,000!",
        "You Are Our Lucky Winner! Claim Your Prize!",
        "WINNER NOTIFICATION: $500,000 Awaits You!",
        "You've Been Selected! Claim Your Reward!",
        "URGENT: Your Lottery Winnings Are Ready!",
    ],
    "body_template": """<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: 'Times New Roman', serif; margin: 0; padding: 20px; background: #fffef0; }
        .container { max-width: 600px; margin: 0 auto; background: white; border: 3px solid gold; padding: 30px; }
        .header { text-align: center; border-bottom: 2px solid gold; padding-bottom: 20px; }
        .prize-box { background: linear-gradient(135deg, #ffd700, #ffec8b); padding: 20px; text-align: center; margin: 20px 0; }
        .prize-amount { font-size: 48px; color: #006400; font-weight: bold; }
        .footer { font-size: 10px; color: #999; margin-top: 30px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>INTERNATIONAL LOTTERY COMMISSION</h1>
            <p>Official Winner Notification</p>
        </div>
        <div class="content">
            <p>Dear Lucky Winner,</p>
            <p>Your email address was selected in our annual lottery draw!</p>
            <div class="prize-box">
                <p>YOUR WINNING AMOUNT:</p>
                <p class="prize-amount">$1,000,000.00</p>
            </div>
            <p>To claim your prize, contact our claims agent.</p>
            <p>Congratulations!</p>
        </div>
        <div class="footer">
            <p>This lottery is sponsored by major international corporations.</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# FAKE INVOICE
# =============================================================================

FAKE_INVOICE: Dict[str, Any] = {
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
        .content { padding: 30px; }
        .invoice-table { width: 100%; border-collapse: collapse; }
        .invoice-table th, .invoice-table td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        .total { font-size: 24px; color: #e74c3c; font-weight: bold; }
        .button { display: inline-block; background: #27ae60; color: white; padding: 15px 30px; text-decoration: none; }
        .warning { background: #fff3cd; border: 1px solid #ffc107; padding: 15px; margin: 20px 0; }
        .footer { font-size: 11px; color: #999; padding: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>INVOICE</h1>
        </div>
        <div class="content">
            <p>Dear Customer,</p>
            <p>Please find your invoice for recent services.</p>
            <table class="invoice-table">
                <tr><th>Description</th><th>Amount</th></tr>
                <tr><td>Annual Subscription</td><td>$299.99</td></tr>
                <tr><td>Premium Support</td><td>$149.99</td></tr>
                <tr><td><strong>TOTAL DUE:</strong></td><td class="total">$449.98</td></tr>
            </table>
            <div class="warning">
                <strong>Payment Required Within 24 Hours</strong>
            </div>
            <center><a href="#" class="button">Pay Now</a></center>
        </div>
        <div class="footer">
            <p>Online Services Inc. | Billing Department</p>
        </div>
    </div>
</body>
</html>""",
}

# =============================================================================
# NEWSLETTER SPAM
# =============================================================================

NEWSLETTER_SPAM: Dict[str, Any] = {
    "category": "spam",
    "sensitivity": "normal",
    "sender_type": "external_spam",
    "department": None,
    "subject_templates": [
        "Your Daily Digest - Top Stories You Can't Miss!",
        "Weekly Roundup: The News Everyone's Talking About",
        "Trending Now: Must-Read Articles Inside",
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
        .content { padding: 20px; }
        .article { border-bottom: 1px solid #eee; padding: 20px 0; }
        .article h3 { margin: 0 0 10px 0; color: #1a1a2e; }
        .read-more { color: #e74c3c; text-decoration: none; font-weight: bold; }
        .ad-box { background: #fff3cd; padding: 20px; margin: 20px 0; text-align: center; border: 2px dashed #ffc107; }
        .footer { background: #1a1a2e; color: #999; padding: 20px; text-align: center; font-size: 11px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>DAILY DIGEST</h1>
            <p>Your personalized news roundup</p>
        </div>
        <div class="content">
            <div class="article">
                <h3>10 Secrets Successful People Never Tell You</h3>
                <p>Discover the hidden habits of millionaires...</p>
                <a href="#" class="read-more">Read More</a>
            </div>
            <div class="ad-box">
                <strong>SPONSORED</strong><br>
                Get 50% off your first order! Use code: NEWSLETTER50
            </div>
            <div class="article">
                <h3>Why Everyone Is Talking About This New Trend</h3>
                <p>The latest craze sweeping the nation...</p>
                <a href="#" class="read-more">Read More</a>
            </div>
        </div>
        <div class="footer">
            <p>You're receiving this because you subscribed.</p>
            <p><a href="#" style="color: #999;">Unsubscribe</a></p>
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
