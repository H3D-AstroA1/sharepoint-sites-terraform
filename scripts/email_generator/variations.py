"""
Content variation pools for M365 email population.

This module provides extensive pools of randomized content to make
generated emails more realistic and varied. Each pool contains
multiple options that are randomly selected during email generation.
"""

import random
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional

# =============================================================================
# GREETING VARIATIONS
# =============================================================================

FORMAL_GREETINGS = [
    "Dear {name}",
    "Dear {name},",
    "Dear Mr./Ms. {last_name}",
    "Good morning {name}",
    "Good afternoon {name}",
]

INFORMAL_GREETINGS = [
    "Hi {name}",
    "Hi {name},",
    "Hello {name}",
    "Hello {name},",
    "Hey {name}",
    "Hey {name},",
    "{name},",
    "{name} -",
]

TEAM_GREETINGS = [
    "Hi Team",
    "Hello Team",
    "Hi Everyone",
    "Hello Everyone",
    "Hi All",
    "Hello All",
    "Team,",
    "All,",
    "Dear Team",
    "Dear Colleagues",
    "Dear All",
]

# =============================================================================
# CLOSING VARIATIONS
# =============================================================================

FORMAL_CLOSINGS = [
    "Best regards",
    "Kind regards",
    "Regards",
    "Sincerely",
    "With best regards",
    "Respectfully",
    "With kind regards",
    "Yours sincerely",
]

INFORMAL_CLOSINGS = [
    "Best",
    "Thanks",
    "Thank you",
    "Cheers",
    "Many thanks",
    "Thanks!",
    "Talk soon",
    "Speak soon",
    "All the best",
    "Take care",
]

PROFESSIONAL_CLOSINGS = [
    "Best regards",
    "Kind regards",
    "Thank you",
    "Thanks",
    "Best",
    "Regards",
]

# =============================================================================
# PROJECT NAME VARIATIONS
# =============================================================================

PROJECT_PREFIXES = [
    "Project", "Initiative", "Program", "Operation", "Campaign",
    "Endeavor", "Venture", "Mission", "Strategy", "Plan",
]

PROJECT_CODENAMES = [
    "Phoenix", "Atlas", "Titan", "Apollo", "Mercury", "Neptune",
    "Orion", "Horizon", "Summit", "Pinnacle", "Catalyst", "Momentum",
    "Velocity", "Quantum", "Fusion", "Synergy", "Nexus", "Apex",
    "Zenith", "Eclipse", "Aurora", "Nova", "Stellar", "Vanguard",
    "Pioneer", "Frontier", "Odyssey", "Genesis", "Evolution", "Transform",
]

PROJECT_DESCRIPTORS = [
    "Digital Transformation", "Cloud Migration", "System Upgrade",
    "Process Improvement", "Customer Experience", "Data Analytics",
    "Security Enhancement", "Infrastructure Modernization",
    "Workflow Automation", "Platform Integration", "Mobile First",
    "AI Implementation", "Cost Optimization", "Quality Assurance",
    "Compliance Update", "Market Expansion", "Product Launch",
    "Brand Refresh", "Talent Development", "Sustainability",
]

# =============================================================================
# DOCUMENT NAME VARIATIONS
# =============================================================================

DOCUMENT_TYPES = [
    "Report", "Analysis", "Summary", "Overview", "Review",
    "Assessment", "Proposal", "Plan", "Strategy", "Guidelines",
    "Policy", "Procedure", "Manual", "Handbook", "Template",
    "Checklist", "Framework", "Roadmap", "Presentation", "Brief",
]

DOCUMENT_PREFIXES = [
    "Q{quarter}", "FY{year}", "{month}", "Weekly", "Monthly",
    "Annual", "Quarterly", "Bi-weekly", "Draft", "Final",
    "Updated", "Revised", "v{version}", "Internal", "External",
]

DOCUMENT_TOPICS = {
    "Human Resources": [
        "Employee Handbook", "Onboarding Guide", "Benefits Summary",
        "Performance Review", "Training Materials", "Policy Update",
        "Recruitment Plan", "Compensation Analysis", "Exit Interview",
        "Diversity Report", "Wellness Program", "Team Building",
    ],
    "Finance Department": [
        "Budget Report", "Financial Statement", "Expense Analysis",
        "Revenue Forecast", "Cash Flow", "Audit Findings",
        "Investment Summary", "Cost Analysis", "Tax Planning",
        "Procurement Report", "Vendor Analysis", "ROI Assessment",
    ],
    "IT Department": [
        "System Architecture", "Security Assessment", "Network Diagram",
        "Disaster Recovery", "Change Request", "Incident Report",
        "Capacity Planning", "Migration Plan", "Technical Specification",
        "User Guide", "API Documentation", "Release Notes",
    ],
    "Marketing Department": [
        "Campaign Brief", "Market Analysis", "Brand Guidelines",
        "Content Calendar", "Social Media Report", "SEO Analysis",
        "Competitor Analysis", "Customer Insights", "Event Plan",
        "Press Release", "Product Launch", "Marketing ROI",
    ],
    "Sales Department": [
        "Sales Forecast", "Pipeline Report", "Territory Analysis",
        "Customer Proposal", "Contract Draft", "Pricing Strategy",
        "Win/Loss Analysis", "Account Plan", "Commission Report",
        "Lead Generation", "Sales Training", "CRM Update",
    ],
    "Executive Leadership": [
        "Board Presentation", "Strategic Plan", "Executive Summary",
        "Investor Update", "Annual Report", "Vision Statement",
        "Organizational Chart", "Leadership Update", "Town Hall",
        "Stakeholder Report", "Risk Assessment", "Growth Strategy",
    ],
    "Legal & Compliance": [
        "Contract Review", "Compliance Checklist", "Legal Brief",
        "NDA Template", "Privacy Policy", "Terms of Service",
        "Regulatory Update", "Risk Assessment", "Audit Report",
        "Litigation Summary", "IP Portfolio", "Due Diligence",
    ],
    "Operations Department": [
        "Process Map", "SOP Document", "Quality Report",
        "Inventory Analysis", "Logistics Plan", "Vendor Scorecard",
        "Capacity Report", "Efficiency Analysis", "Safety Report",
        "Maintenance Schedule", "Resource Allocation", "KPI Dashboard",
    ],
    "Claims Department": [
        "Claim Assessment", "Case Summary", "Settlement Report",
        "Investigation Findings", "Liability Analysis", "Damage Report",
        "Claim Status Update", "Adjuster Notes", "Coverage Review",
        "Subrogation Report", "Reserve Analysis", "Claim Audit",
    ],
}

# =============================================================================
# MEETING TOPIC VARIATIONS
# =============================================================================

MEETING_TYPES = [
    "Sync", "Check-in", "Review", "Planning", "Brainstorm",
    "Workshop", "Training", "Presentation", "Discussion", "Update",
    "Kickoff", "Retrospective", "Stand-up", "Deep Dive", "Strategy",
]

MEETING_TOPICS = [
    "Q{quarter} Planning", "Project Status", "Budget Review",
    "Team Alignment", "Process Improvement", "Customer Feedback",
    "Product Roadmap", "Technical Architecture", "Resource Planning",
    "Risk Assessment", "Performance Review", "Goal Setting",
    "Sprint Planning", "Release Planning", "Stakeholder Update",
    "Cross-functional Collaboration", "Innovation Workshop",
    "Best Practices", "Lessons Learned", "Knowledge Sharing",
]

MEETING_DURATIONS = [
    "15 minutes", "30 minutes", "45 minutes", "1 hour",
    "1.5 hours", "2 hours", "Half day", "Full day",
]

# =============================================================================
# TIME-BASED CONTENT
# =============================================================================

def get_current_quarter() -> int:
    """Get current fiscal quarter."""
    month = datetime.now().month
    return (month - 1) // 3 + 1

def get_current_fiscal_year() -> int:
    """Get current fiscal year."""
    return datetime.now().year

def get_random_quarter() -> str:
    """Get a random quarter string."""
    return f"Q{random.randint(1, 4)}"

def get_random_fiscal_year() -> str:
    """Get a random fiscal year string."""
    current_year = datetime.now().year
    year = random.choice([current_year - 1, current_year, current_year + 1])
    return f"FY{year}"

def get_random_month() -> str:
    """Get a random month name."""
    months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    return random.choice(months)

def get_random_deadline() -> str:
    """Get a random deadline string."""
    days = random.randint(1, 14)
    deadline_date = datetime.now() + timedelta(days=days)
    formats = [
        deadline_date.strftime("%B %d, %Y"),
        deadline_date.strftime("%d/%m/%Y"),
        deadline_date.strftime("%A, %B %d"),
        f"End of day {deadline_date.strftime('%A')}",
        f"COB {deadline_date.strftime('%A')}",
        f"By {deadline_date.strftime('%B %d')}",
    ]
    return random.choice(formats)

def get_random_date_reference() -> str:
    """Get a random date reference for content."""
    references = [
        f"last {random.choice(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'])}",
        f"this {random.choice(['week', 'month', 'quarter'])}",
        f"next {random.choice(['week', 'month', 'quarter'])}",
        "yesterday",
        "earlier today",
        "this morning",
        "this afternoon",
        f"on {get_random_month()} {random.randint(1, 28)}",
    ]
    return random.choice(references)

# =============================================================================
# TONE VARIATIONS
# =============================================================================

URGENCY_PHRASES = {
    "high": [
        "URGENT: ", "ACTION REQUIRED: ", "IMMEDIATE: ",
        "TIME SENSITIVE: ", "PRIORITY: ", "CRITICAL: ",
    ],
    "medium": [
        "Important: ", "Please review: ", "Attention: ",
        "FYI - Action needed: ", "Reminder: ",
    ],
    "low": [
        "", "FYI: ", "For your information: ",
        "When you have a moment: ", "No rush: ",
    ],
}

POLITE_REQUESTS = [
    "Could you please",
    "Would you mind",
    "I'd appreciate if you could",
    "When you have a chance, please",
    "If possible, could you",
    "Would it be possible to",
    "I was wondering if you could",
    "Please",
    "Kindly",
]

ACKNOWLEDGMENTS = [
    "Thank you for your help with this.",
    "I appreciate your assistance.",
    "Thanks in advance for your support.",
    "Your help is greatly appreciated.",
    "Thank you for taking the time to review this.",
    "I really appreciate your input on this.",
    "Thanks for your attention to this matter.",
]

FOLLOW_UP_PHRASES = [
    "Please let me know if you have any questions.",
    "Feel free to reach out if you need any clarification.",
    "Don't hesitate to contact me if you need more information.",
    "I'm happy to discuss this further if needed.",
    "Let me know if you'd like to schedule a call to discuss.",
    "Please reach out if you have any concerns.",
    "I'm available to chat if you want to discuss this.",
]

# =============================================================================
# DEPARTMENT-SPECIFIC CONTENT
# =============================================================================

DEPARTMENT_JARGON = {
    "Human Resources": [
        "employee engagement", "talent acquisition", "performance management",
        "onboarding process", "retention strategy", "workforce planning",
        "compensation and benefits", "learning and development", "HRIS",
        "employee experience", "culture initiatives", "DEI programs",
    ],
    "Finance Department": [
        "EBITDA", "cash flow", "P&L statement", "balance sheet",
        "accounts receivable", "accounts payable", "budget variance",
        "financial forecast", "ROI analysis", "cost center",
        "capital expenditure", "operating expenses", "fiscal year",
    ],
    "IT Department": [
        "system integration", "API endpoint", "cloud infrastructure",
        "cybersecurity", "data migration", "DevOps pipeline",
        "technical debt", "scalability", "uptime SLA",
        "incident response", "change management", "disaster recovery",
    ],
    "Marketing Department": [
        "brand awareness", "lead generation", "conversion rate",
        "customer journey", "content strategy", "SEO optimization",
        "social engagement", "campaign ROI", "market segmentation",
        "brand positioning", "competitive analysis", "go-to-market",
    ],
    "Sales Department": [
        "sales pipeline", "quota attainment", "deal velocity",
        "customer acquisition", "upselling", "cross-selling",
        "sales enablement", "territory management", "win rate",
        "customer lifetime value", "churn rate", "ARR/MRR",
    ],
    "Executive Leadership": [
        "strategic initiative", "organizational alignment", "stakeholder management",
        "board governance", "executive sponsorship", "change leadership",
        "vision and mission", "corporate strategy", "market positioning",
        "competitive advantage", "growth trajectory", "shareholder value",
    ],
    "Legal & Compliance": [
        "regulatory compliance", "contractual obligations", "due diligence",
        "intellectual property", "data privacy", "GDPR compliance",
        "risk mitigation", "legal liability", "corporate governance",
        "audit trail", "policy enforcement", "legal review",
    ],
    "Operations Department": [
        "operational efficiency", "process optimization", "supply chain",
        "quality assurance", "lean methodology", "continuous improvement",
        "resource utilization", "capacity planning", "KPI tracking",
        "vendor management", "logistics coordination", "inventory control",
    ],
    "Claims Department": [
        "claim adjudication", "loss assessment", "coverage determination",
        "subrogation recovery", "reserve estimation", "claim lifecycle",
        "first notice of loss", "settlement negotiation", "liability evaluation",
        "fraud detection", "claim triage", "total loss determination",
    ],
}

DEPARTMENT_PRIORITIES = {
    "Human Resources": [
        "improving employee satisfaction scores",
        "reducing time-to-hire metrics",
        "enhancing the onboarding experience",
        "developing leadership training programs",
        "implementing new HRIS features",
    ],
    "Finance Department": [
        "closing the books on time",
        "improving forecast accuracy",
        "reducing operational costs",
        "streamlining the approval process",
        "enhancing financial reporting",
    ],
    "IT Department": [
        "maintaining system uptime",
        "improving security posture",
        "completing the cloud migration",
        "reducing technical debt",
        "enhancing user experience",
    ],
    "Marketing Department": [
        "increasing brand awareness",
        "improving lead quality",
        "launching the new campaign",
        "enhancing customer engagement",
        "optimizing marketing spend",
    ],
    "Sales Department": [
        "hitting quarterly targets",
        "expanding into new markets",
        "improving win rates",
        "reducing sales cycle time",
        "enhancing customer relationships",
    ],
    "Executive Leadership": [
        "driving organizational growth",
        "improving operational efficiency",
        "enhancing stakeholder communication",
        "executing the strategic plan",
        "fostering innovation culture",
    ],
    "Legal & Compliance": [
        "ensuring regulatory compliance",
        "reducing legal risk exposure",
        "streamlining contract processes",
        "updating privacy policies",
        "completing the audit successfully",
    ],
    "Operations Department": [
        "improving process efficiency",
        "reducing operational costs",
        "enhancing quality metrics",
        "optimizing resource allocation",
        "streamlining vendor management",
    ],
    "Claims Department": [
        "reducing claim processing time",
        "improving settlement accuracy",
        "enhancing fraud detection rates",
        "increasing customer satisfaction scores",
        "optimizing reserve management",
    ],
}

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def get_random_greeting(tone: str = "professional", name: str = "") -> str:
    """Get a random greeting based on tone."""
    if tone == "formal":
        greeting = random.choice(FORMAL_GREETINGS)
    elif tone == "informal":
        greeting = random.choice(INFORMAL_GREETINGS)
    elif tone == "team":
        return random.choice(TEAM_GREETINGS)
    else:  # professional
        greeting = random.choice(PROFESSIONAL_CLOSINGS[:3] + INFORMAL_GREETINGS[:4])
        greeting = random.choice(["Hi {name}", "Hello {name}", "Dear {name}"])
    
    if name:
        first_name = name.split()[0] if name else "there"
        last_name = name.split()[-1] if len(name.split()) > 1 else name
        greeting = greeting.replace("{name}", first_name)
        greeting = greeting.replace("{last_name}", last_name)
    else:
        greeting = greeting.replace("{name}", "there")
        greeting = greeting.replace("{last_name}", "")
    
    return greeting

def get_random_closing(tone: str = "professional") -> str:
    """Get a random closing based on tone."""
    if tone == "formal":
        return random.choice(FORMAL_CLOSINGS)
    elif tone == "informal":
        return random.choice(INFORMAL_CLOSINGS)
    else:  # professional
        return random.choice(PROFESSIONAL_CLOSINGS)

def get_random_project_name() -> str:
    """Generate a random project name."""
    style = random.choice(["codename", "descriptor", "combined"])
    
    if style == "codename":
        return f"{random.choice(PROJECT_PREFIXES)} {random.choice(PROJECT_CODENAMES)}"
    elif style == "descriptor":
        return random.choice(PROJECT_DESCRIPTORS)
    else:
        return f"{random.choice(PROJECT_CODENAMES)} - {random.choice(PROJECT_DESCRIPTORS)}"

def get_random_document_name(department: str = "General") -> str:
    """Generate a random document name for a department."""
    topics = DOCUMENT_TOPICS.get(department, DOCUMENT_TOPICS.get("Executive Leadership", []))
    topic = random.choice(topics) if topics else "Document"
    
    prefix = random.choice(DOCUMENT_PREFIXES)
    prefix = prefix.replace("{quarter}", str(get_current_quarter()))
    prefix = prefix.replace("{year}", str(get_current_fiscal_year()))
    prefix = prefix.replace("{month}", get_random_month())
    prefix = prefix.replace("{version}", f"{random.randint(1, 5)}.{random.randint(0, 9)}")
    
    if random.random() < 0.5:
        return f"{prefix} {topic}"
    else:
        return topic

def get_random_meeting_topic(department: str = "General") -> str:
    """Generate a random meeting topic."""
    topic = random.choice(MEETING_TOPICS)
    topic = topic.replace("{quarter}", str(get_current_quarter()))
    
    if random.random() < 0.3:
        meeting_type = random.choice(MEETING_TYPES)
        return f"{topic} {meeting_type}"
    
    return topic

def get_random_meeting_duration() -> str:
    """Get a random meeting duration."""
    return random.choice(MEETING_DURATIONS)

def get_department_jargon(department: str) -> str:
    """Get random department-specific jargon."""
    jargon = DEPARTMENT_JARGON.get(department, DEPARTMENT_JARGON.get("Executive Leadership", []))
    return random.choice(jargon) if jargon else "business operations"

def get_department_priority(department: str) -> str:
    """Get a random department priority."""
    priorities = DEPARTMENT_PRIORITIES.get(department, DEPARTMENT_PRIORITIES.get("Executive Leadership", []))
    return random.choice(priorities) if priorities else "achieving our goals"

def get_random_urgency() -> str:
    """Get a random urgency level."""
    return random.choices(
        ["high", "medium", "low"],
        weights=[10, 30, 60]
    )[0]

def get_urgency_prefix(urgency: str = "low") -> str:
    """Get an urgency prefix for subject lines."""
    return random.choice(URGENCY_PHRASES.get(urgency, URGENCY_PHRASES["low"]))

def get_random_polite_request() -> str:
    """Get a random polite request phrase."""
    return random.choice(POLITE_REQUESTS)

def get_random_acknowledgment() -> str:
    """Get a random acknowledgment phrase."""
    return random.choice(ACKNOWLEDGMENTS)

def get_random_follow_up() -> str:
    """Get a random follow-up phrase."""
    return random.choice(FOLLOW_UP_PHRASES)

def get_random_tone() -> str:
    """Get a random email tone."""
    return random.choices(
        ["formal", "professional", "informal"],
        weights=[20, 60, 20]
    )[0]

# =============================================================================
# CONTENT VARIATION POOLS
# =============================================================================

STATUS_UPDATE_INTROS = [
    "I wanted to give you a quick update on",
    "Here's the latest status on",
    "Just a brief update regarding",
    "I'm writing to update you on",
    "Following up on our discussion about",
    "As promised, here's an update on",
    "I wanted to share some progress on",
    "Quick update for you on",
]

REQUEST_INTROS = [
    "I was hoping you could help me with",
    "I need your assistance with",
    "Could you please take a look at",
    "I'd like to request your input on",
    "Would you be able to help with",
    "I'm reaching out because I need",
    "I was wondering if you could assist with",
]

MEETING_INTROS = [
    "I'd like to schedule some time to discuss",
    "Can we find time to meet about",
    "I was hoping we could connect regarding",
    "Would you be available to discuss",
    "I'd like to set up a meeting to talk about",
    "Can we schedule a quick call about",
]

POSITIVE_UPDATES = [
    "I'm pleased to report that",
    "Great news -",
    "I'm happy to share that",
    "We've made excellent progress on",
    "Things are going well with",
    "I'm excited to announce that",
]

CONCERN_UPDATES = [
    "I wanted to flag a potential issue with",
    "We've encountered a challenge with",
    "I need to bring to your attention",
    "There's a concern I'd like to discuss regarding",
    "We may need to revisit our approach to",
]

def get_random_intro(intro_type: str = "status") -> str:
    """Get a random introduction phrase."""
    intros = {
        "status": STATUS_UPDATE_INTROS,
        "request": REQUEST_INTROS,
        "meeting": MEETING_INTROS,
        "positive": POSITIVE_UPDATES,
        "concern": CONCERN_UPDATES,
    }
    return random.choice(intros.get(intro_type, STATUS_UPDATE_INTROS))
