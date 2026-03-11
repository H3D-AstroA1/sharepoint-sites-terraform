"""
Exchange Web Services (EWS) client for email operations.

This module provides EWS-based email creation with full control over
message properties including timestamps and draft status, which are
read-only in the Microsoft Graph API.

Requires: exchangelib (pip install exchangelib)

Authentication:
    EWS requires OAuth2 client credentials flow with the following:
    - App registration with EWS.full_access_as_app permission
    - Client ID and Client Secret
    - Tenant ID
    
    The app config is loaded from .app_config.json which is created
    by the menu.py app registration process.
"""

from datetime import datetime
from typing import Any, Dict, List, Optional
from pathlib import Path
import json
import logging
import urllib.request
import urllib.parse

# Try to import exchangelib - it's optional
try:
    from exchangelib import (
        Account,
        Configuration,
        Credentials,
        DELEGATE,
        IMPERSONATION,
        Message,
        Mailbox,
        HTMLBody,
        FileAttachment,
        OAuth2Credentials,
        OAuth2AuthorizationCodeCredentials,
        Identity,
        Build,
        Version,
    )
    from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
    from exchangelib.folders import Inbox, SentItems, DeletedItems, JunkEmail, Drafts
    EXCHANGELIB_AVAILABLE = True
except ImportError:
    EXCHANGELIB_AVAILABLE = False

# Path to app config file (created by menu.py app registration)
APP_CONFIG_FILE = Path(__file__).parent.parent / ".app_config.json"


class Colors:
    """ANSI color codes for terminal output."""
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    NC = '\033[0m'


class EWSClient:
    """
    Exchange Web Services client for creating emails with full property control.
    
    Unlike the Graph API, EWS allows setting:
    - datetime_received (receivedDateTime)
    - datetime_sent (sentDateTime)
    - is_draft status
    - All MAPI properties
    
    Authentication:
        Uses OAuth2 client credentials flow with app registration.
        Requires EWS.full_access_as_app permission on the app.
    """
    
    # EWS endpoint for Office 365
    EWS_URL = "https://outlook.office365.com/EWS/Exchange.asmx"
    
    # EWS OAuth2 scope
    EWS_SCOPE = "https://outlook.office365.com/.default"
    
    def __init__(self, rate_limit_config: Optional[Dict[str, Any]] = None):
        """
        Initialize the EWS client.
        
        Args:
            rate_limit_config: Optional rate limiting configuration.
        """
        self.account: Optional[Any] = None
        self.credentials: Optional[Any] = None
        self.config: Optional[Any] = None
        self._current_mailbox: Optional[str] = None
        self._app_config: Optional[Dict[str, Any]] = None
        self._access_token: Optional[str] = None
        self._tenant_id: Optional[str] = None
        
        # Rate limiting - EWS can handle faster requests than Graph API
        # Default to 50ms delay (vs 100ms for Graph API)
        config = rate_limit_config or {}
        self.request_delay_ms = config.get("request_delay_ms", 50)
        
    @staticmethod
    def is_available() -> bool:
        """Check if exchangelib is installed."""
        return EXCHANGELIB_AVAILABLE
    
    def _load_app_config(self) -> Optional[Dict[str, Any]]:
        """Load app configuration from .app_config.json."""
        if APP_CONFIG_FILE.exists():
            try:
                with open(APP_CONFIG_FILE, 'r') as f:
                    return json.load(f)
            except Exception:
                pass
        return None
    
    def _get_ews_token(self, app_config: Dict[str, Any]) -> Optional[str]:
        """
        Get EWS access token using client credentials flow.
        
        Args:
            app_config: App configuration with client_id, client_secret, tenant_id.
            
        Returns:
            Access token string or None if failed.
        """
        app_id = app_config.get("app_id")
        client_secret = app_config.get("client_secret")
        tenant_id = app_config.get("tenant_id")
        
        if not all([app_id, client_secret, tenant_id]):
            print("  ⚠ Missing app credentials for EWS authentication")
            return None
        
        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        
        data = urllib.parse.urlencode({
            "client_id": app_id,
            "client_secret": client_secret,
            "scope": self.EWS_SCOPE,
            "grant_type": "client_credentials"
        }).encode('utf-8')
        
        try:
            req = urllib.request.Request(token_url, data=data)
            req.add_header("Content-Type", "application/x-www-form-urlencoded")
            
            with urllib.request.urlopen(req, timeout=30) as response:
                result = json.loads(response.read().decode('utf-8'))
                return result.get("access_token")
        except Exception as e:
            print(f"  ✗ Failed to get EWS token: {e}")
            return None
    
    def authenticate(self, app_config: Optional[Dict[str, Any]] = None) -> bool:
        """
        Authenticate using OAuth2 client credentials flow.
        
        Args:
            app_config: Optional app configuration. If not provided, loads from .app_config.json.
            
        Returns:
            True if authentication successful.
        """
        if not EXCHANGELIB_AVAILABLE:
            print("  ⚠ exchangelib not installed. Run: pip install exchangelib")
            return False
        
        # Load app config if not provided
        if app_config is None:
            app_config = self._load_app_config()
        
        if not app_config:
            print("  ⚠ No app configuration found. Run 'Manage App Registration' from the menu.")
            return False
        
        self._app_config = app_config
        self._tenant_id = app_config.get("tenant_id")
        
        # Get EWS access token using client credentials
        access_token = self._get_ews_token(app_config)
        if not access_token:
            return False
        
        self._access_token = access_token
        
        try:
            # Create OAuth2 credentials with the access token
            self.credentials = OAuth2Credentials(
                client_id=app_config.get("app_id"),
                client_secret=app_config.get("client_secret"),
                tenant_id=self._tenant_id,
                access_token={
                    'access_token': access_token,
                    'token_type': 'Bearer',
                }
            )
            print(f"  {Colors.GREEN}✓{Colors.NC} EWS authentication successful")
            return True
        except Exception as e:
            print(f"  ✗ EWS authentication failed: {e}")
            return False
    
    def _get_account(self, mailbox: str) -> Optional[Any]:
        """
        Get or create an Account object for the specified mailbox.
        
        When using client credentials flow with full_access_as_app permission,
        we need to configure the credentials with the identity to impersonate.
        
        Args:
            mailbox: User's email address (UPN).
            
        Returns:
            Account object or None if failed.
        """
        if not EXCHANGELIB_AVAILABLE or not self.credentials:
            return None
        
        try:
            # For full_access_as_app with client credentials, we need to
            # set the identity on the credentials object to specify which
            # mailbox to impersonate
            self.credentials.identity = Identity(primary_smtp_address=mailbox)
            
            # Create configuration for EWS with explicit service endpoint
            # Note: Only provide service_endpoint, not server (they are mutually exclusive)
            config = Configuration(
                credentials=self.credentials,
                service_endpoint='https://outlook.office365.com/EWS/Exchange.asmx',
                auth_type='OAuth 2.0',
            )
            
            # Create account with impersonation access
            account = Account(
                primary_smtp_address=mailbox,
                config=config,
                autodiscover=False,
                access_type=IMPERSONATION,
            )
            
            return account
        except Exception as e:
            print(f"  ✗ Failed to access mailbox {mailbox}: {e}")
            return None
    
    def _get_folder(self, account: Any, folder_name: str) -> Optional[Any]:
        """
        Get the folder object for the specified folder name.
        
        Args:
            account: EWS Account object.
            folder_name: Folder name (inbox, sentitems, deleteditems, junkemail, drafts).
            
        Returns:
            Folder object or None.
        """
        folder_map = {
            'inbox': account.inbox,
            'sentitems': account.sent,
            'deleteditems': account.trash,
            'junkemail': account.junk,
            'drafts': account.drafts,
        }
        
        return folder_map.get(folder_name.lower(), account.inbox)
    
    def _get_folder_id(self, folder_name: str) -> str:
        """
        Get the EWS DistinguishedFolderId for a folder name.
        
        Args:
            folder_name: Folder name (inbox, sentitems, deleteditems, junkemail, drafts).
            
        Returns:
            EWS DistinguishedFolderId string.
        """
        folder_id_map = {
            'inbox': 'inbox',
            'sentitems': 'sentitems',
            'deleteditems': 'deleteditems',
            'junkemail': 'junkemail',
            'drafts': 'drafts',
        }
        return folder_id_map.get(folder_name.lower(), 'inbox')
    
    def create_email(self, mailbox: str, email: Dict[str, Any], folder: str = "inbox") -> bool:
        """
        Create an email in the specified mailbox folder using EWS MIME upload.
        
        This method uses EWS MIME upload to create emails with proper timestamps
        and without the [Draft] prefix. MIME upload is the only way to create
        "received" emails with backdated timestamps in Exchange.
        
        Args:
            mailbox: User's email address (UPN).
            email: Email data dictionary containing:
                - subject: Email subject
                - body: Email body (HTML)
                - sender: Dict with 'name' and 'email'
                - recipient: Dict with 'name' and 'email'
                - date: datetime for the email timestamp
                - cc_recipients: Optional list of CC recipients
                - bcc_recipients: Optional list of BCC recipients
                - attachments: Optional list of attachments
            folder: Target folder (inbox, sentitems, deleteditems, junkemail, drafts).
            
        Returns:
            True if email created successfully, False otherwise.
        """
        import time
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.base import MIMEBase
        from email import encoders
        from email.utils import formatdate, formataddr
        import re
        
        if not EXCHANGELIB_AVAILABLE:
            return False
        
        # Apply rate limiting
        if self.request_delay_ms > 0:
            time.sleep(self.request_delay_ms / 1000.0)
        
        try:
            account = self._get_account(mailbox)
            if not account:
                return False
            
            target_folder = self._get_folder(account, folder)
            if not target_folder:
                return False
            
            # Extract email data
            sender = email.get("sender", {})
            recipient = email.get("recipient", {})
            email_date = email.get("date", datetime.now())
            
            # EWS requires timezone-aware datetimes
            if email_date.tzinfo is None:
                from exchangelib import UTC
                email_date = email_date.replace(tzinfo=UTC)
            
            # Determine From/To based on folder
            if folder.lower() in ("sentitems", "drafts"):
                from_email = mailbox
                from_name = recipient.get("name", "")
                to_email = sender.get("email", "recipient@external.com")
                to_name = sender.get("name", "Recipient")
            else:
                from_email = sender.get("email", f"sender@{mailbox.split('@')[-1]}")
                from_name = sender.get("name", "Sender")
                to_email = mailbox
                to_name = recipient.get("name", "")
            
            # Build MIME message
            mime_msg = MIMEMultipart('alternative')
            mime_msg['Subject'] = email.get("subject", "No Subject")
            mime_msg['From'] = formataddr((from_name, from_email))
            mime_msg['To'] = formataddr((to_name, to_email))
            mime_msg['Date'] = formatdate(email_date.timestamp(), localtime=False)
            mime_msg['Message-ID'] = f"<{email_date.timestamp()}.{id(email)}@{from_email.split('@')[-1]}>"
            
            # Add CC recipients
            if email.get("cc_recipients"):
                cc_list = [formataddr((cc.get("name", ""), cc.get("email", "")))
                          for cc in email["cc_recipients"]]
                mime_msg['Cc'] = ", ".join(cc_list)
            
            # Add body - both plain text and HTML
            body_html = email.get("body", "")
            # Create plain text version by stripping HTML tags
            body_text = re.sub(r'<[^>]+>', '', body_html)
            body_text = body_text.replace('&nbsp;', ' ').replace('&amp;', '&')
            
            # Attach plain text and HTML parts
            part_text = MIMEText(body_text, 'plain', 'utf-8')
            part_html = MIMEText(body_html, 'html', 'utf-8')
            mime_msg.attach(part_text)
            mime_msg.attach(part_html)
            
            # Add attachments
            if email.get("attachments"):
                # Convert to mixed multipart for attachments
                mime_msg_with_attachments = MIMEMultipart('mixed')
                mime_msg_with_attachments['Subject'] = mime_msg['Subject']
                mime_msg_with_attachments['From'] = mime_msg['From']
                mime_msg_with_attachments['To'] = mime_msg['To']
                mime_msg_with_attachments['Date'] = mime_msg['Date']
                mime_msg_with_attachments['Message-ID'] = mime_msg['Message-ID']
                if mime_msg.get('Cc'):
                    mime_msg_with_attachments['Cc'] = mime_msg['Cc']
                
                # Add the alternative part (text/html body)
                mime_msg_with_attachments.attach(mime_msg)
                
                # Add each attachment
                for attachment in email["attachments"]:
                    att_name = attachment.get("name", "attachment.txt")
                    att_content = attachment.get("content", b"")
                    att_type = attachment.get("contentType", "application/octet-stream")
                    
                    if isinstance(att_content, str):
                        att_content = att_content.encode('utf-8')
                    
                    maintype, subtype = att_type.split('/', 1) if '/' in att_type else ('application', 'octet-stream')
                    part = MIMEBase(maintype, subtype)
                    part.set_payload(att_content)
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename=att_name)
                    mime_msg_with_attachments.attach(part)
                
                mime_msg = mime_msg_with_attachments
            
            # Convert MIME message to bytes
            mime_content = mime_msg.as_bytes()
            
            # Use CreateItem with extended properties to set message flags
            # PR_MESSAGE_FLAGS (0x0E07) with MSGFLAG_READ (0x0001) marks as read, not draft
            import base64
            
            # Base64 encode the MIME content for SOAP
            mime_b64 = base64.b64encode(mime_content).decode('ascii')
            
            # Get the folder ID - need to map to EWS distinguished folder names
            folder_id = self._get_folder_id(folder)
            
            # Determine if message should be read
            is_read_val = "true" if self._should_be_read(email_date) else "false"
            
            # Format the date for EWS (ISO 8601 format)
            # EWS expects: 2024-01-15T10:30:00Z
            ews_date = email_date.strftime('%Y-%m-%dT%H:%M:%SZ')
            
            # Build the CreateItem SOAP request with MIME content
            # Using MimeContent allows setting the Date header
            # Setting extended properties for timestamps:
            # - PR_MESSAGE_DELIVERY_TIME (0x0E06) - when message was delivered
            # - PR_CLIENT_SUBMIT_TIME (0x0039) - when message was sent
            # - PR_MESSAGE_FLAGS (0x0E07) - message flags (1 = read)
            soap_body = f'''<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2016"/>
    <t:ExchangeImpersonation>
      <t:ConnectingSID>
        <t:SmtpAddress>{mailbox}</t:SmtpAddress>
      </t:ConnectingSID>
    </t:ExchangeImpersonation>
  </soap:Header>
  <soap:Body>
    <m:CreateItem MessageDisposition="SaveOnly">
      <m:SavedItemFolderId>
        <t:DistinguishedFolderId Id="{folder_id}"/>
      </m:SavedItemFolderId>
      <m:Items>
        <t:Message>
          <t:MimeContent CharacterSet="UTF-8">{mime_b64}</t:MimeContent>
          <t:IsRead>{is_read_val}</t:IsRead>
          <t:ExtendedProperty>
            <t:ExtendedFieldURI PropertyTag="0x0E07" PropertyType="Integer"/>
            <t:Value>1</t:Value>
          </t:ExtendedProperty>
          <t:ExtendedProperty>
            <t:ExtendedFieldURI PropertyTag="0x0E06" PropertyType="SystemTime"/>
            <t:Value>{ews_date}</t:Value>
          </t:ExtendedProperty>
          <t:ExtendedProperty>
            <t:ExtendedFieldURI PropertyTag="0x0039" PropertyType="SystemTime"/>
            <t:Value>{ews_date}</t:Value>
          </t:ExtendedProperty>
        </t:Message>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>'''
            
            # Make the SOAP request with retry logic
            import urllib.request
            import urllib.error
            
            headers = {
                'Content-Type': 'text/xml; charset=utf-8',
                'Authorization': f'Bearer {self._access_token}',
            }
            
            # Retry configuration
            max_retries = 3
            retry_delay = 2  # seconds
            timeout = 90  # Increased timeout for large MIME content
            
            for attempt in range(max_retries):
                try:
                    req = urllib.request.Request(
                        'https://outlook.office365.com/EWS/Exchange.asmx',
                        data=soap_body.encode('utf-8'),
                        headers=headers,
                        method='POST'
                    )
                    
                    with urllib.request.urlopen(req, timeout=timeout) as response:
                        response_data = response.read().decode('utf-8')
                        
                        # Check for success
                        if 'NoError' in response_data:
                            return True
                        else:
                            # Extract error message
                            error_match = re.search(r'<m:MessageText>([^<]+)</m:MessageText>', response_data)
                            if error_match:
                                print(f"  ✗ EWS CreateItem failed: {error_match.group(1)}")
                            else:
                                # Check for response code
                                code_match = re.search(r'<m:ResponseCode>([^<]+)</m:ResponseCode>', response_data)
                                if code_match and code_match.group(1) != 'NoError':
                                    print(f"  ✗ EWS CreateItem failed: {code_match.group(1)}")
                            return False
                            
                except (urllib.error.URLError, TimeoutError, ConnectionError) as e:
                    if attempt < max_retries - 1:
                        # Retry with exponential backoff
                        time.sleep(retry_delay * (attempt + 1))
                        continue
                    else:
                        raise  # Re-raise on final attempt
            
            # If we get here without returning, something went wrong
            return False
            
        except Exception as e:
            print(f"  ✗ EWS create_email failed: {e}")
            return False
    
    def _should_be_read(self, email_date: datetime) -> bool:
        """
        Determine if an email should be marked as read based on its age.
        
        Emails older than 7 days are marked as read.
        
        Args:
            email_date: The date of the email.
            
        Returns:
            True if the email should be marked as read.
        """
        # Use timezone-aware datetime for comparison with EWS dates
        from exchangelib import UTC
        now = datetime.now(UTC)
        
        # Ensure email_date is also timezone-aware
        if email_date.tzinfo is None:
            email_date = email_date.replace(tzinfo=UTC)
        
        age_days = (now - email_date).days
        return age_days > 7
    
    def verify_connection(self, mailbox: str) -> bool:
        """
        Verify that we can connect to the specified mailbox.
        
        Args:
            mailbox: User's email address (UPN).
            
        Returns:
            True if connection successful.
        """
        if not EXCHANGELIB_AVAILABLE:
            return False
        
        try:
            account = self._get_account(mailbox)
            if account:
                # Try to access inbox to verify connection
                _ = account.inbox.total_count
                return True
            return False
        except Exception as e:
            print(f"  ✗ EWS connection verification failed: {e}")
            return False


def check_exchangelib_installed() -> tuple:
    """
    Check if exchangelib is installed and return version info.
    
    Returns:
        Tuple of (installed: bool, version: str or None)
    """
    try:
        import exchangelib
        return True, getattr(exchangelib, '__version__', 'unknown')
    except ImportError:
        return False, None


def install_exchangelib() -> bool:
    """
    Attempt to install exchangelib using pip.
    
    Returns:
        True if installation successful.
    """
    import subprocess
    import sys
    
    try:
        print("  Installing exchangelib...")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "exchangelib"],
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode == 0:
            print("  ✓ exchangelib installed successfully")
            return True
        else:
            print(f"  ✗ Installation failed: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"  ✗ Installation failed: {e}")
        return False
