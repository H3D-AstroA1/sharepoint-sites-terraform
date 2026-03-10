# Azure AD Auto-Discovery Design for Email Population

## Overview

This document outlines the design for Azure AD auto-discovery integration with the M365 Email Population Tool. The goal is to automatically discover users, groups, and distribution lists from Azure AD to create realistic email scenarios.

**IMPORTANT: This is an ENHANCEMENT to the existing implementation, NOT a replacement.**

### What Stays the Same
- ✅ Existing `mailboxes.yaml` manual configuration still works
- ✅ All current email templates and content generation
- ✅ Existing folder support (inbox, sent, drafts, deleted, junk)
- ✅ Current attachment and threading functionality
- ✅ Existing Graph API client and authentication
- ✅ Menu integration and CLI options
- ✅ Cleanup functionality

### What Gets Enhanced
- 🆕 **Optional** Azure AD auto-discovery as alternative to manual YAML
- 🆕 User pool from Azure AD for senders/recipients
- 🆕 CC/BCC support in email generation
- 🆕 Groups and distribution lists as recipients
- 🆕 Organizational hierarchy awareness
- 🆕 **Menu integration** for easy user selection

## Menu Integration Design

### Updated Main Menu
```
╔══════════════════════════════════════════════════════════════════════╗
║                    M365 SHAREPOINT & EMAIL TOOL                       ║
╠══════════════════════════════════════════════════════════════════════╣
║                                                                       ║
║  SHAREPOINT OPERATIONS                                                ║
║  ─────────────────────                                                ║
║  [1] Create SharePoint Sites                                          ║
║  [2] Populate Files in Sites                                          ║
║  [3] List SharePoint Sites                                            ║
║  [4] List Files in Sites                                              ║
║                                                                       ║
║  EMAIL OPERATIONS                                                     ║
║  ────────────────                                                     ║
║  [5] Populate Emails in Mailboxes                                     ║
║  [6] Cleanup Emails from Mailboxes                                    ║
║  [7] List Mailboxes                                                   ║
║  [8] Discover Users from Azure AD  ← NEW                              ║
║                                                                       ║
║  CONFIGURATION                                                        ║
║  ─────────────                                                        ║
║  [C] Edit Configuration Files                                         ║
║  [A] Manage App Registration                                          ║
║  [P] Check Prerequisites                                              ║
║                                                                       ║
║  [H] Help    [Q] Quit                                                 ║
╚══════════════════════════════════════════════════════════════════════╝
```

### Option [8] Discover Users from Azure AD - Submenu
```
╔══════════════════════════════════════════════════════════════════════╗
║                    AZURE AD USER DISCOVERY                            ║
╠══════════════════════════════════════════════════════════════════════╣
║                                                                       ║
║  [1] Discover All Users                                               ║
║      Discover all users from Azure AD and show mailbox status         ║
║                                                                       ║
║  [2] Discover Users with Mailboxes Only                               ║
║      Find users with Exchange Online licenses                         ║
║                                                                       ║
║  [3] Discover Groups & Distribution Lists                             ║
║      Find M365 groups, security groups, and distribution lists        ║
║                                                                       ║
║  [4] Filter by Department                                             ║
║      Discover users from specific departments                         ║
║                                                                       ║
║  [5] Save Discovery to Cache                                          ║
║      Cache discovered users for faster email population               ║
║                                                                       ║
║  [6] View Cached Users                                                ║
║      Show currently cached users and groups                           ║
║                                                                       ║
║  [7] Clear Cache                                                      ║
║      Remove cached user data                                          ║
║                                                                       ║
║  [B] Back to Main Menu                                                ║
╚══════════════════════════════════════════════════════════════════════╝
```

### Enhanced Option [5] Populate Emails - Updated Flow
```
╔══════════════════════════════════════════════════════════════════════╗
║                    POPULATE EMAILS IN MAILBOXES                       ║
╠══════════════════════════════════════════════════════════════════════╣
║                                                                       ║
║  Select user source:                                                  ║
║                                                                       ║
║  [1] Use mailboxes.yaml configuration                                 ║
║      Use manually configured users from config file                   ║
║                                                                       ║
║  [2] Auto-discover from Azure AD  ← NEW                               ║
║      Automatically discover users with mailboxes                      ║
║                                                                       ║
║  [3] Use cached Azure AD users  ← NEW                                 ║
║      Use previously discovered and cached users                       ║
║                                                                       ║
║  [B] Back to Main Menu                                                ║
╚══════════════════════════════════════════════════════════════════════╝
```

### Auto-Discovery Interactive Flow
```
╔══════════════════════════════════════════════════════════════════════╗
║                    AUTO-DISCOVER USERS                                ║
╠══════════════════════════════════════════════════════════════════════╣
║                                                                       ║
║  🔍 Discovering users from Azure AD...                                ║
║                                                                       ║
║  ┌────────────────────────────────────────────────────────────────┐  ║
║  │  Discovery Results                                              │  ║
║  ├────────────────────────────────────────────────────────────────┤  ║
║  │  Total users found:           5,200                             │  ║
║  │  Users with mailboxes:        20                                │  ║
║  │  Users without mailboxes:     5,180                             │  ║
║  │  M365 Groups:                 45                                │  ║
║  │  Distribution Lists:          12                                │  ║
║  └────────────────────────────────────────────────────────────────┘  ║
║                                                                       ║
║  Mailbox Users (will receive emails):                                 ║
║  ─────────────────────────────────────                                ║
║  #   Name                    Department          Email                ║
║  1   John Smith              Human Resources     john.smith@...       ║
║  2   Jane Doe                Finance             jane.doe@...         ║
║  3   Bob Wilson              IT Department       bob.wilson@...       ║
║  ... (17 more)                                                        ║
║                                                                       ║
║  [A] Select All mailbox users (20)                                    ║
║  [S] Select specific users                                            ║
║  [F] Filter by department                                             ║
║  [C] Cache these results                                              ║
║  [P] Proceed with population                                          ║
║  [B] Back                                                             ║
╚══════════════════════════════════════════════════════════════════════╝
```

### Department Filter Selection
```
╔══════════════════════════════════════════════════════════════════════╗
║                    FILTER BY DEPARTMENT                               ║
╠══════════════════════════════════════════════════════════════════════╣
║                                                                       ║
║  Select departments to include:                                       ║
║                                                                       ║
║  [ ] Human Resources (3 mailbox users, 450 total)                     ║
║  [x] IT Department (5 mailbox users, 120 total)                       ║
║  [x] Finance (4 mailbox users, 200 total)                             ║
║  [ ] Marketing (2 mailbox users, 180 total)                           ║
║  [ ] Sales (3 mailbox users, 350 total)                               ║
║  [ ] Operations (2 mailbox users, 400 total)                          ║
║  [ ] Executive (1 mailbox user, 10 total)                             ║
║                                                                       ║
║  Selected: 2 departments, 9 mailbox users, 320 total users            ║
║                                                                       ║
║  [Enter] Apply filter    [A] Select all    [N] Select none            ║
║  [B] Back                                                             ║
╚══════════════════════════════════════════════════════════════════════╝
```

## Requirements Summary

1. **Auto-discover users from Azure AD** - No manual YAML maintenance
2. **Distinguish mailbox vs non-mailbox users** - Configurable limits
3. **Support all users as senders/recipients** - Even without mailboxes
4. **CC/BCC support** - Multiple recipients from user pool
5. **Groups and Distribution Lists** - Future-proof design
6. **Configurable limits** - Adaptable to changing user counts
7. **Menu Integration** - Easy-to-use interactive menus
8. **Backward Compatible** - Existing YAML config continues to work

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                        Azure AD Tenant                               │
├─────────────────────────────────────────────────────────────────────┤
│  ┌──────────────┐  ┌──────────────┐  ┌────────────────────────┐    │
│  │    Users     │  │   Groups     │  │  Distribution Lists    │    │
│  │   ~5200      │  │   M365/Sec   │  │    Mail-enabled        │    │
│  └──────────────┘  └──────────────┘  └────────────────────────┘    │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    Azure AD Discovery Service                        │
├─────────────────────────────────────────────────────────────────────┤
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │                    Graph API Queries                          │  │
│  │  • GET /users - All users with attributes                     │  │
│  │  • GET /groups - M365 groups, security groups                 │  │
│  │  • GET /groups/{id}/members - Group membership                │  │
│  │  • Check mailbox existence via /users/{id}/mailFolders        │  │
│  └──────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│                      User Pool Manager                               │
├─────────────────────────────────────────────────────────────────────┤
│  ┌────────────────┐  ┌────────────────┐  ┌────────────────────┐    │
│  │ Mailbox Users  │  │ Non-Mailbox    │  │ Groups/DLs         │    │
│  │ ~20 users      │  │ Users ~5180    │  │ Variable           │    │
│  │ Recipients     │  │ Sender Pool    │  │ CC/BCC Pool        │    │
│  └────────────────┘  └────────────────┘  └────────────────────┘    │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    Email Generation Engine                           │
├─────────────────────────────────────────────────────────────────────┤
│  Folder      │ From              │ To                │ CC/BCC       │
│  ────────────┼───────────────────┼───────────────────┼──────────────│
│  Inbox       │ Any user/group    │ Mailbox owner     │ Any users    │
│  Sent Items  │ Mailbox owner     │ Any user/group    │ Any users    │
│  Drafts      │ Mailbox owner     │ Any user/group    │ Any users    │
│  Deleted     │ Mixed             │ Mixed             │ Any users    │
│  Junk        │ External/Spam     │ Mailbox owner     │ -            │
└─────────────────────────────────────────────────────────────────────┘
```

## Data Model

### User Categories

```python
class UserCategory(Enum):
    MAILBOX_USER = "mailbox"           # Has Exchange license/mailbox
    NON_MAILBOX_USER = "non_mailbox"   # Azure AD user, no mailbox
    EXTERNAL_USER = "external"          # Guest users
    SERVICE_ACCOUNT = "service"         # Service/app accounts
```

### Recipient Types

```python
class RecipientType(Enum):
    USER = "user"                       # Individual user
    M365_GROUP = "m365_group"           # Microsoft 365 Group
    SECURITY_GROUP = "security_group"   # Security Group
    DISTRIBUTION_LIST = "distribution"  # Distribution List
    MAIL_CONTACT = "mail_contact"       # External mail contact
```

### User Pool Structure

```python
@dataclass
class UserPool:
    mailbox_users: List[User]           # Users with mailboxes (recipients)
    non_mailbox_users: List[User]       # Users without mailboxes (senders)
    groups: List[Group]                 # M365 and Security groups
    distribution_lists: List[DistList]  # Distribution lists
    
    def get_random_sender(self) -> User:
        """Get random user from all users pool"""
        
    def get_random_recipients(self, count: int) -> List[User]:
        """Get random recipients for To/CC/BCC"""
        
    def get_random_group(self) -> Group:
        """Get random group for CC/BCC"""
```

## Configuration Schema

### Updated mailboxes.yaml

```yaml
# =============================================================================
# EMAIL POPULATION CONFIGURATION
# =============================================================================

settings:
  # User source: 'yaml' (manual) or 'azure_ad' (auto-discovery)
  user_source: azure_ad
  
  # Email generation settings
  default_email_count: 100
  date_range_months: 12
  business_hours_bias: 0.8
  thread_probability: 0.4
  
  # CC/BCC settings
  cc_probability: 0.3              # 30% of emails have CC
  bcc_probability: 0.1             # 10% of emails have BCC
  max_cc_recipients: 5             # Max CC recipients per email
  max_bcc_recipients: 3            # Max BCC recipients per email

# =============================================================================
# AZURE AD AUTO-DISCOVERY SETTINGS
# =============================================================================

azure_ad:
  enabled: true
  
  # -----------------------------------------------------------------------------
  # User Discovery
  # -----------------------------------------------------------------------------
  users:
    # Discover all users (no limit by default, set to limit)
    max_users: null                 # null = no limit, or set number
    
    # Filter options (all optional)
    include_departments: []         # Empty = all departments
    exclude_departments: ["Contractors", "Vendors"]
    include_job_titles: []          # Empty = all titles
    exclude_service_accounts: true
    exclude_external_users: false   # Include guest users?
    
    # Mailbox detection
    validate_mailbox_exists: true   # Check if user has mailbox
    
  # -----------------------------------------------------------------------------
  # Group Discovery
  # -----------------------------------------------------------------------------
  groups:
    enabled: true
    max_groups: null                # null = no limit
    
    # Group types to include
    include_m365_groups: true       # Microsoft 365 Groups
    include_security_groups: true   # Security Groups (mail-enabled)
    include_distribution_lists: true # Distribution Lists
    
    # Filter options
    exclude_patterns: ["Test*", "Dev*", "Temp*"]
    min_members: 2                  # Minimum members to include
    
  # -----------------------------------------------------------------------------
  # Caching
  # -----------------------------------------------------------------------------
  cache:
    enabled: true
    ttl_minutes: 60                 # Cache TTL
    cache_file: ".azure_ad_cache.json"

# =============================================================================
# MANUAL USER CONFIGURATION (when user_source: yaml)
# =============================================================================

users:
  - upn: john.smith@contoso.com
    role: HR Manager
    department: Human Resources
    email_volume: high
    has_mailbox: true

  - upn: jane.doe@contoso.com
    role: Financial Analyst
    department: Finance
    email_volume: medium
    has_mailbox: true
```

## Graph API Queries

### 1. Get All Users

```http
GET /users
?$select=id,userPrincipalName,displayName,mail,department,jobTitle,accountEnabled
&$filter=accountEnabled eq true
&$top=999
```

### 2. Check Mailbox Existence

```http
GET /users/{user-id}/mailFolders/inbox
# Returns 200 if mailbox exists, 404 if not
```

### 3. Get Users with Exchange License

```http
GET /users
?$filter=assignedLicenses/any(l:l/skuId eq '{EXCHANGE_ONLINE_SKU}')
&$select=id,userPrincipalName,displayName,mail,department,jobTitle
```

### 4. Get M365 Groups

```http
GET /groups
?$filter=groupTypes/any(c:c eq 'Unified')
&$select=id,displayName,mail,description
```

### 5. Get Distribution Lists

```http
GET /groups
?$filter=mailEnabled eq true and securityEnabled eq false
&$select=id,displayName,mail,description
```

### 6. Get Group Members

```http
GET /groups/{group-id}/members
?$select=id,userPrincipalName,displayName,mail
```

## Email Generation Logic

### Folder-Specific Rules

```python
def generate_email_for_folder(folder: str, mailbox_owner: User, user_pool: UserPool):
    if folder == "inbox":
        # Received emails - from anyone to mailbox owner
        sender = user_pool.get_random_sender()
        to = [mailbox_owner]
        cc = user_pool.get_random_recipients(random.randint(0, 3)) if random.random() < 0.3 else []
        
    elif folder == "sentitems":
        # Sent emails - from mailbox owner to anyone
        sender = mailbox_owner
        to = user_pool.get_random_recipients(random.randint(1, 3))
        cc = user_pool.get_random_recipients(random.randint(0, 3)) if random.random() < 0.3 else []
        
    elif folder == "drafts":
        # Draft emails - from mailbox owner to anyone (unsent)
        sender = mailbox_owner
        to = user_pool.get_random_recipients(random.randint(1, 5))
        cc = user_pool.get_random_recipients(random.randint(0, 5)) if random.random() < 0.4 else []
        
    elif folder == "deleteditems":
        # Deleted emails - mix of sent and received
        if random.random() < 0.5:
            # Deleted received email
            sender = user_pool.get_random_sender()
            to = [mailbox_owner]
        else:
            # Deleted sent email
            sender = mailbox_owner
            to = user_pool.get_random_recipients(random.randint(1, 3))
            
    elif folder == "junkemail":
        # Spam - external senders only
        sender = generate_spam_sender()
        to = [mailbox_owner]
        cc = []
        
    return Email(sender=sender, to=to, cc=cc, bcc=bcc)
```

### CC/BCC Logic

```python
def add_cc_bcc(email: Email, user_pool: UserPool, config: dict):
    # CC - visible to all recipients
    if random.random() < config.get('cc_probability', 0.3):
        cc_count = random.randint(1, config.get('max_cc_recipients', 5))
        
        # Mix of users and groups
        if random.random() < 0.2 and user_pool.groups:
            # Add a group to CC
            email.cc.append(user_pool.get_random_group())
            cc_count -= 1
            
        email.cc.extend(user_pool.get_random_recipients(cc_count))
    
    # BCC - hidden from other recipients
    if random.random() < config.get('bcc_probability', 0.1):
        bcc_count = random.randint(1, config.get('max_bcc_recipients', 3))
        email.bcc = user_pool.get_random_recipients(bcc_count)
```

## Implementation Plan

### Phase 1: Azure AD Discovery Service
- [ ] Create `azure_ad_discovery.py` module
- [ ] Implement user discovery with pagination
- [ ] Implement mailbox validation
- [ ] Add caching layer

### Phase 2: User Pool Manager
- [ ] Create `user_pool.py` module
- [ ] Implement user categorization (mailbox/non-mailbox)
- [ ] Add random selection methods
- [ ] Support filtering by department/role

### Phase 3: Group and Distribution List Support
- [ ] Add group discovery
- [ ] Add distribution list discovery
- [ ] Implement group member resolution
- [ ] Add groups to recipient pool

### Phase 4: CC/BCC Integration
- [ ] Update email generation to support CC/BCC
- [ ] Update Graph API payload builder
- [ ] Add configuration options

### Phase 5: Configuration Updates
- [ ] Update `mailboxes.yaml` schema
- [ ] Update `config.py` loader
- [ ] Add CLI options for auto-discovery

### Phase 6: Testing and Documentation
- [ ] Test with large user pools
- [ ] Performance optimization
- [ ] Update documentation

## API Permissions Required

| Permission | Type | Purpose |
|------------|------|---------|
| `User.Read.All` | Application | Read all user profiles |
| `Group.Read.All` | Application | Read all groups |
| `GroupMember.Read.All` | Application | Read group memberships |
| `Mail.ReadWrite` | Application | Create emails in mailboxes |
| `Directory.Read.All` | Application | Read directory data |

## Performance Considerations

1. **Pagination**: Handle large user lists with pagination (999 per page)
2. **Caching**: Cache user/group data to avoid repeated API calls
3. **Batch Requests**: Use Graph API batching for mailbox validation
4. **Rate Limiting**: Implement exponential backoff for throttling

## Security Considerations

1. **Least Privilege**: Only request necessary permissions
2. **Data Handling**: Don't store sensitive user data unnecessarily
3. **Audit Logging**: Log all operations for compliance
4. **Consent**: Ensure admin consent for application permissions

## Future Enhancements

1. **Dynamic Groups**: Support Azure AD dynamic groups
2. **Shared Mailboxes**: Support shared/resource mailboxes
3. **Room/Equipment**: Support room and equipment mailboxes
4. **Cross-Tenant**: Support multi-tenant scenarios
5. **Incremental Sync**: Delta queries for user changes

## Additional Enhancements

### 1. Organizational Hierarchy Awareness
```python
# Use manager/direct reports relationships for realistic email patterns
class OrgHierarchy:
    def get_manager(self, user: User) -> Optional[User]
    def get_direct_reports(self, user: User) -> List[User]
    def get_peers(self, user: User) -> List[User]  # Same manager
    def get_department_head(self, department: str) -> Optional[User]
```

**Benefits:**
- Emails from managers to direct reports look realistic
- Status reports go "up" the chain
- Team emails include actual team members

### 2. Smart Sender Selection Based on Email Type
```python
EMAIL_TYPE_SENDER_RULES = {
    "newsletter": ["hr", "marketing", "communications"],
    "security_alert": ["it", "security"],
    "policy_update": ["hr", "legal", "compliance"],
    "project_update": ["same_department", "cross_functional"],
    "meeting_request": ["manager", "peer", "direct_report"],
    "leadership_message": ["executive", "department_head"],
}
```

### 3. Realistic Reply Chains
```python
# When generating thread replies, use actual participants
def generate_reply_chain(original_email: Email, user_pool: UserPool):
    participants = [original_email.sender] + original_email.to + original_email.cc
    # Replies come from actual participants, not random users
```

### 4. Department-Aware CC Selection
```python
# CC recipients often from same or related departments
RELATED_DEPARTMENTS = {
    "Finance": ["Accounting", "Procurement", "Executive"],
    "IT": ["Security", "Operations", "Engineering"],
    "HR": ["Legal", "Compliance", "Executive"],
    "Marketing": ["Sales", "Product", "Communications"],
}
```

### 5. Time-Zone Aware Email Generation
```python
# Users in different time zones send emails at appropriate times
def get_user_timezone(user: User) -> str:
    # From Azure AD officeLocation or usageLocation
    return user.timezone or "UTC"

def generate_email_time(sender: User) -> datetime:
    # Business hours in sender's timezone
    tz = get_user_timezone(sender)
    return random_business_hours(tz)
```

### 6. Shared Mailbox Support
```python
class MailboxType(Enum):
    USER = "user"
    SHARED = "shared"
    ROOM = "room"
    EQUIPMENT = "equipment"

# Shared mailboxes can also receive emails
# Room/Equipment mailboxes for meeting requests
```

### 7. Email Conversation Patterns
```python
CONVERSATION_PATTERNS = {
    "quick_exchange": {"replies": 2-3, "participants": 2},
    "team_discussion": {"replies": 5-10, "participants": 3-6},
    "approval_chain": {"replies": 3-4, "participants": 2-3},
    "broadcast": {"replies": 0, "participants": "many"},
}
```

### 8. Intelligent Attachment Distribution
```python
# Certain departments send certain attachment types more often
DEPARTMENT_ATTACHMENT_BIAS = {
    "Finance": {"xlsx": 0.6, "pdf": 0.3, "docx": 0.1},
    "Marketing": {"pptx": 0.4, "pdf": 0.3, "docx": 0.2, "xlsx": 0.1},
    "Legal": {"pdf": 0.5, "docx": 0.4, "xlsx": 0.1},
    "Engineering": {"docx": 0.4, "pdf": 0.3, "xlsx": 0.2, "pptx": 0.1},
}
```

### 9. Progressive Loading for Large Tenants
```python
# For 5000+ users, load in batches to avoid memory issues
async def discover_users_progressive(batch_size: int = 500):
    async for batch in paginate_users(batch_size):
        yield process_batch(batch)
        # Allow other operations between batches
```

### 10. Dry Run Mode with Statistics
```bash
python populate_emails.py --auto-discover --dry-run

# Output:
# ╔══════════════════════════════════════════════════════════════╗
# ║                    DRY RUN SUMMARY                            ║
# ╠══════════════════════════════════════════════════════════════╣
# ║  Users discovered:        5,200                               ║
# ║  Users with mailboxes:    20                                  ║
# ║  Groups discovered:       45                                  ║
# ║  Distribution lists:      12                                  ║
# ║                                                               ║
# ║  Emails to generate:      2,000 (100 per mailbox)            ║
# ║  Estimated time:          ~15 minutes                         ║
# ║  API calls required:      ~2,500                              ║
# ╚══════════════════════════════════════════════════════════════╝
```

### 11. Resume Capability
```python
# Save progress to resume if interrupted
class ProgressTracker:
    def save_checkpoint(self, mailbox: str, emails_created: int)
    def load_checkpoint(self) -> Optional[Checkpoint]
    def clear_checkpoint(self)

# Resume from last checkpoint
python populate_emails.py --resume
```

### 12. Parallel Processing
```python
# Process multiple mailboxes concurrently
async def populate_mailboxes_parallel(
    mailboxes: List[str],
    max_concurrent: int = 5
):
    semaphore = asyncio.Semaphore(max_concurrent)
    tasks = [populate_with_semaphore(mb, semaphore) for mb in mailboxes]
    await asyncio.gather(*tasks)
```

## Performance Optimizations

### 1. Batch Graph API Requests
```python
# Batch up to 20 requests per call
async def batch_create_emails(emails: List[Email], batch_size: int = 20):
    for batch in chunks(emails, batch_size):
        await graph_batch_request(batch)
```

### 2. Connection Pooling
```python
# Reuse HTTP connections
session = aiohttp.ClientSession(
    connector=aiohttp.TCPConnector(limit=100)
)
```

### 3. Smart Caching Strategy
```python
CACHE_STRATEGY = {
    "users": {"ttl": 3600, "refresh": "background"},      # 1 hour
    "groups": {"ttl": 3600, "refresh": "background"},     # 1 hour
    "mailbox_status": {"ttl": 86400, "refresh": "lazy"},  # 24 hours
}
```

## Summary

This design provides:
- ✅ Auto-discovery of all Azure AD users (5200+)
- ✅ Automatic mailbox detection (20 mailbox users)
- ✅ All users available as senders/recipients
- ✅ CC/BCC support with configurable limits
- ✅ Groups and distribution lists support
- ✅ Configurable limits (no hardcoded values)
- ✅ Future-proof for changing user counts
- ✅ Caching for performance
- ✅ Extensible architecture

**Additional Enhancements:**
- ✅ Organizational hierarchy awareness
- ✅ Smart sender selection by email type
- ✅ Realistic reply chains with actual participants
- ✅ Department-aware CC selection
- ✅ Time-zone aware email generation
- ✅ Shared/Room/Equipment mailbox support
- ✅ Email conversation patterns
- ✅ Intelligent attachment distribution
- ✅ Progressive loading for large tenants
- ✅ Dry run mode with statistics
- ✅ Resume capability for interrupted runs
- ✅ Parallel processing for speed

**Performance Optimizations:**
- ✅ Batch Graph API requests
- ✅ Connection pooling
- ✅ Smart caching strategy

Ready to proceed with implementation?
