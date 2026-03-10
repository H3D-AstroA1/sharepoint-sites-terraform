"""
Microsoft Graph API client for email operations.

Handles authentication and email creation in M365 mailboxes
using the Microsoft Graph API.
"""

import json
import os
import platform
import random
import shutil
import subprocess
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Any


# Default Azure CLI installation paths on Windows
AZURE_CLI_PATHS = [
    r"C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
    r"C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
]

# App config file path
SCRIPT_DIR = Path(__file__).parent.resolve()
APP_CONFIG_FILE = SCRIPT_DIR.parent / ".app_config.json"


class GraphClient:
    """Client for Microsoft Graph API email operations."""
    
    GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
    
    def __init__(self):
        """Initialize the Graph client."""
        self.access_token: Optional[str] = None
        self._az_path: Optional[str] = None
    
    def authenticate(self, app_config: Optional[Dict] = None) -> bool:
        """
        Authenticate with Microsoft Graph API.
        
        Tries app credentials first, then falls back to Azure CLI.
        
        Args:
            app_config: Optional app configuration with client credentials.
            
        Returns:
            True if authentication successful, False otherwise.
        """
        # Try app credentials first
        if app_config and app_config.get("client_secret"):
            self.access_token = self._get_token_client_credentials(app_config)
            if self.access_token:
                return True
        
        # Try loading app config from file
        file_config = self._load_app_config()
        if file_config and file_config.get("client_secret"):
            self.access_token = self._get_token_client_credentials(file_config)
            if self.access_token:
                return True
        
        # Fall back to Azure CLI
        self.access_token = self._get_token_azure_cli()
        return self.access_token is not None
    
    def create_email(self, mailbox: str, email: Dict[str, Any], folder: str = "inbox") -> bool:
        """
        Create an email in a user's mailbox.
        
        Uses the messages endpoint to create emails directly in the specified folder
        with backdated timestamps.
        
        Args:
            mailbox: User's email address (UPN).
            email: Email data dictionary.
            folder: Target folder (inbox, drafts, sentitems). Default is inbox.
            
        Returns:
            True if email created successfully, False otherwise.
        """
        if not self.access_token:
            return False
        
        # Create in specific folder to avoid Drafts
        # Using mailFolders/{folder}/messages puts the email directly in that folder
        url = f"{self.GRAPH_BASE_URL}/users/{mailbox}/mailFolders/{folder}/messages"
        
        # Build message payload
        message = self._build_message_payload(mailbox, email)
        
        try:
            # Create the message
            response = self._make_request("POST", url, message)
            message_id = response.get("id")
            
            if not message_id:
                return False
            
            # Upload attachments if present
            if email.get("attachments"):
                for attachment in email["attachments"]:
                    self._upload_attachment(mailbox, message_id, attachment)
            
            return True
            
        except urllib.error.HTTPError as e:
            if e.code == 403:
                print(f"  Permission denied for mailbox: {mailbox}")
            elif e.code == 404:
                print(f"  Mailbox not found: {mailbox}")
            return False
        except Exception as e:
            print(f"  Error creating email: {e}")
            return False
    
    def create_email_batch(
        self,
        mailbox: str,
        emails: List[Dict[str, Any]],
        batch_size: int = 20
    ) -> Dict[str, int]:
        """
        Create multiple emails using batch requests.
        
        Args:
            mailbox: User's email address (UPN).
            emails: List of email data dictionaries.
            batch_size: Number of emails per batch (max 20).
            
        Returns:
            Dictionary with success and failure counts.
        """
        results = {"success": 0, "failed": 0}
        
        # Process in batches
        for i in range(0, len(emails), batch_size):
            batch = emails[i:i + batch_size]
            batch_results = self._process_batch(mailbox, batch)
            results["success"] += batch_results["success"]
            results["failed"] += batch_results["failed"]
        
        return results
    
    def verify_permissions(self) -> Dict[str, Any]:
        """
        Verify that required Graph API permissions are available.
        
        Returns:
            Dictionary with permission status.
        """
        result = {
            "has_permissions": False,
            "can_read_users": False,
            "can_write_mail": False,
            "error": None,
        }
        
        if not self.access_token:
            result["error"] = "Not authenticated"
            return result
        
        # Test user read permission
        try:
            url = f"{self.GRAPH_BASE_URL}/users?$top=1"
            self._make_request("GET", url)
            result["can_read_users"] = True
        except urllib.error.HTTPError as e:
            if e.code == 403:
                result["error"] = "Missing User.Read.All permission"
        except Exception:
            pass
        
        # Test mail write permission (by checking a specific user's messages)
        # This is a lightweight check - actual write will be tested during email creation
        result["can_write_mail"] = True  # Assume true, will fail on actual write if not
        
        result["has_permissions"] = result["can_read_users"]
        
        return result
    
    def get_user_info(self, upn: str) -> Optional[Dict[str, Any]]:
        """
        Get user information from Azure AD.
        
        Args:
            upn: User Principal Name.
            
        Returns:
            User information dictionary or None.
        """
        if not self.access_token:
            return None
        
        try:
            url = f"{self.GRAPH_BASE_URL}/users/{upn}"
            return self._make_request("GET", url)
        except Exception:
            return None
    
    def verify_mailbox(self, upn: str) -> Dict[str, Any]:
        """
        Verify that a mailbox exists and is accessible.
        
        Args:
            upn: User Principal Name (email address).
            
        Returns:
            Dictionary with verification status:
            - exists: True if user exists in Azure AD
            - has_mailbox: True if user has a mailbox
            - display_name: User's display name if found
            - error: Error message if any
        """
        result = {
            "exists": False,
            "has_mailbox": False,
            "display_name": None,
            "error": None,
        }
        
        if not self.access_token:
            result["error"] = "Not authenticated"
            return result
        
        try:
            # Check if user exists
            url = f"{self.GRAPH_BASE_URL}/users/{upn}"
            user_info = self._make_request("GET", url)
            
            if user_info:
                result["exists"] = True
                result["display_name"] = user_info.get("displayName", upn)
                
                # Check if user has a mailbox by trying to access their mailbox settings
                try:
                    mailbox_url = f"{self.GRAPH_BASE_URL}/users/{upn}/mailboxSettings"
                    self._make_request("GET", mailbox_url)
                    result["has_mailbox"] = True
                except urllib.error.HTTPError as e:
                    if e.code == 404:
                        result["error"] = "User exists but has no mailbox"
                    elif e.code == 403:
                        # Permission denied might mean mailbox exists but we can't access settings
                        # Try a simpler check
                        result["has_mailbox"] = True  # Assume mailbox exists
                except Exception:
                    # Assume mailbox exists if we can't check settings
                    result["has_mailbox"] = True
                    
        except urllib.error.HTTPError as e:
            if e.code == 404:
                result["error"] = f"User not found: {upn}"
            elif e.code == 403:
                result["error"] = f"Permission denied to access user: {upn}"
            else:
                result["error"] = f"HTTP error {e.code}: {e.reason}"
        except Exception as e:
            result["error"] = str(e)
        
        return result
    
    def verify_mailboxes(self, upns: List[str]) -> Dict[str, Dict[str, Any]]:
        """
        Verify multiple mailboxes at once.
        
        Args:
            upns: List of User Principal Names.
            
        Returns:
            Dictionary mapping UPN to verification result.
        """
        results = {}
        for upn in upns:
            results[upn] = self.verify_mailbox(upn)
        return results
    
    # =========================================================================
    # EMAIL DELETION METHODS
    # =========================================================================
    
    def get_emails(
        self,
        mailbox: str,
        folder: str = "inbox",
        top: int = 100,
        filter_query: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Get emails from a mailbox folder.
        
        Args:
            mailbox: User's email address (UPN).
            folder: Folder name (inbox, sentitems, deleteditems, etc.).
            top: Maximum number of emails to retrieve.
            filter_query: Optional OData filter query.
            
        Returns:
            List of email dictionaries.
        """
        if not self.access_token:
            return []
        
        emails = []
        url = f"{self.GRAPH_BASE_URL}/users/{mailbox}/mailFolders/{folder}/messages"
        url += f"?$top={min(top, 999)}&$select=id,subject,receivedDateTime,from,isRead"
        
        if filter_query:
            url += f"&$filter={urllib.parse.quote(filter_query)}"
        
        try:
            while url and len(emails) < top:
                response = self._make_request("GET", url)
                emails.extend(response.get("value", []))
                url = response.get("@odata.nextLink")
            
            return emails[:top]
        except Exception as e:
            print(f"  Error getting emails: {e}")
            return []
    
    def get_email_count(self, mailbox: str, folder: str = "inbox") -> int:
        """
        Get the count of emails in a mailbox folder.
        
        Args:
            mailbox: User's email address (UPN).
            folder: Folder name.
            
        Returns:
            Number of emails in the folder.
        """
        if not self.access_token:
            return 0
        
        try:
            url = f"{self.GRAPH_BASE_URL}/users/{mailbox}/mailFolders/{folder}/messages/$count"
            req = urllib.request.Request(url)
            req.add_header("Authorization", f"Bearer {self.access_token}")
            req.add_header("ConsistencyLevel", "eventual")
            
            with urllib.request.urlopen(req, timeout=30) as response:
                return int(response.read().decode())
        except Exception:
            # Fallback: get messages and count
            emails = self.get_emails(mailbox, folder, top=1000)
            return len(emails)
    
    def delete_email(self, mailbox: str, message_id: str, permanent: bool = False) -> bool:
        """
        Delete a single email.
        
        Args:
            mailbox: User's email address (UPN).
            message_id: The message ID to delete.
            permanent: If True, permanently delete. If False, move to Deleted Items.
            
        Returns:
            True if successful, False otherwise.
        """
        if not self.access_token:
            return False
        
        try:
            if permanent:
                # Permanently delete
                url = f"{self.GRAPH_BASE_URL}/users/{mailbox}/messages/{message_id}"
                self._make_request("DELETE", url)
            else:
                # Move to Deleted Items
                url = f"{self.GRAPH_BASE_URL}/users/{mailbox}/messages/{message_id}/move"
                self._make_request("POST", url, {"destinationId": "deleteditems"})
            
            return True
        except Exception as e:
            return False
    
    def delete_emails_batch(
        self,
        mailbox: str,
        message_ids: List[str],
        permanent: bool = False,
        batch_size: int = 20
    ) -> Dict[str, int]:
        """
        Delete multiple emails in batches.
        
        Args:
            mailbox: User's email address (UPN).
            message_ids: List of message IDs to delete.
            permanent: If True, permanently delete.
            batch_size: Number of deletions per batch.
            
        Returns:
            Dictionary with success and failed counts.
        """
        results = {"success": 0, "failed": 0}
        
        for i in range(0, len(message_ids), batch_size):
            batch = message_ids[i:i + batch_size]
            
            for msg_id in batch:
                if self.delete_email(mailbox, msg_id, permanent):
                    results["success"] += 1
                else:
                    results["failed"] += 1
        
        return results
    
    def delete_all_emails(
        self,
        mailbox: str,
        folder: str = "inbox",
        permanent: bool = False,
        max_emails: int = 10000
    ) -> Dict[str, int]:
        """
        Delete all emails from a mailbox folder.
        
        Args:
            mailbox: User's email address (UPN).
            folder: Folder name to clear.
            permanent: If True, permanently delete.
            max_emails: Maximum number of emails to delete.
            
        Returns:
            Dictionary with success and failed counts.
        """
        results = {"success": 0, "failed": 0}
        
        # Get emails in batches and delete
        while results["success"] + results["failed"] < max_emails:
            emails = self.get_emails(mailbox, folder, top=100)
            
            if not emails:
                break
            
            message_ids = [e.get("id") for e in emails if e.get("id")]
            batch_results = self.delete_emails_batch(mailbox, message_ids, permanent)
            
            results["success"] += batch_results["success"]
            results["failed"] += batch_results["failed"]
            
            # If we couldn't delete any, stop to avoid infinite loop
            if batch_results["success"] == 0:
                break
        
        return results
    
    def empty_deleted_items(self, mailbox: str) -> Dict[str, int]:
        """
        Permanently delete all emails in the Deleted Items folder.
        
        Args:
            mailbox: User's email address (UPN).
            
        Returns:
            Dictionary with success and failed counts.
        """
        return self.delete_all_emails(mailbox, "deleteditems", permanent=True)
    
    def _build_message_payload(self, mailbox: str, email: Dict[str, Any]) -> Dict[str, Any]:
        """
        Build the message payload for Graph API.
        
        Args:
            mailbox: Recipient mailbox.
            email: Email data dictionary.
            
        Returns:
            Message payload dictionary.
        """
        sender = email.get("sender", {})
        recipient = email.get("recipient", {})
        email_date = email.get("date", datetime.now())
        
        message = {
            "subject": email.get("subject", "No Subject"),
            "body": {
                "contentType": "HTML",
                "content": email.get("body", ""),
            },
            "from": {
                "emailAddress": {
                    "name": sender.get("name", "Sender"),
                    "address": sender.get("email", f"sender@{mailbox.split('@')[-1]}"),
                }
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "name": recipient.get("name", ""),
                        "address": mailbox,
                    }
                }
            ],
            "receivedDateTime": email_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "sentDateTime": email_date.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "isRead": self._should_be_read(email_date),
            "isDraft": False,
        }
        
        # Add CC recipients if present
        if email.get("cc_recipients"):
            message["ccRecipients"] = [
                {
                    "emailAddress": {
                        "name": cc.get("name", ""),
                        "address": cc.get("email", ""),
                    }
                }
                for cc in email["cc_recipients"]
            ]
        
        # Add importance based on category
        category = email.get("category", "")
        if category == "organisational":
            message["importance"] = random.choice(["normal", "high", "high"])  # 66% high for org emails
        elif category == "security":
            message["importance"] = "high"  # Security emails always high importance
        else:
            # Small chance of high importance for other emails
            message["importance"] = "high" if random.random() < 0.1 else "normal"
        
        # Add categories/labels - combine sensitivity with color categories
        categories = []
        sensitivity = email.get("sensitivity_label", "")
        if sensitivity and sensitivity.lower() != "general":
            categories.append(sensitivity)
        
        # Add color categories randomly (like Outlook categories)
        color_categories = self._generate_color_categories(category)
        categories.extend(color_categories)
        
        if categories:
            message["categories"] = categories
        
        # Add flag status - some emails should be flagged for follow-up
        message["flag"] = self._generate_flag_status(email_date, category)
        
        # Add @mentions in body if applicable
        if email.get("has_mention") and "@" in email.get("body", ""):
            message["mentionsPreview"] = {"isMentioned": True}
        
        # Add inference classification (focused vs other)
        message["inferenceClassification"] = self._get_inference_classification(sender, category)
        
        # Add internet message headers for external emails
        sender_email = sender.get("email", "")
        if sender_email and not sender_email.endswith(mailbox.split("@")[-1]):
            message["internetMessageHeaders"] = [
                {
                    "name": "X-MS-Exchange-Organization-SCL",
                    "value": str(random.choice([0, 0, 0, 1, 2]))  # Spam confidence level
                },
                {
                    "name": "X-MS-Exchange-Organization-AuthSource",
                    "value": "external.mail.protection.outlook.com"
                }
            ]
        
        # Add conversation ID for threading
        if email.get("conversation_id"):
            message["conversationId"] = email["conversation_id"]
        
        # Add reply/forward indicators
        subject = email.get("subject", "")
        if subject.startswith("RE:") or subject.startswith("Re:"):
            message["isReply"] = True
        elif subject.startswith("FW:") or subject.startswith("Fwd:"):
            message["isForward"] = True
        
        return message
    
    def _upload_attachment(
        self,
        mailbox: str,
        message_id: str,
        attachment: Dict[str, Any]
    ) -> bool:
        """
        Upload an attachment to a message.
        
        Args:
            mailbox: User's email address.
            message_id: Message ID to attach to.
            attachment: Attachment data dictionary.
            
        Returns:
            True if successful, False otherwise.
        """
        url = f"{self.GRAPH_BASE_URL}/users/{mailbox}/messages/{message_id}/attachments"
        
        attachment_payload = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": attachment.get("name", "attachment"),
            "contentType": attachment.get("content_type", "application/octet-stream"),
            "contentBytes": attachment.get("content_base64", ""),
        }
        
        try:
            self._make_request("POST", url, attachment_payload)
            return True
        except Exception:
            return False
    
    def _process_batch(
        self,
        mailbox: str,
        emails: List[Dict[str, Any]]
    ) -> Dict[str, int]:
        """
        Process a batch of emails using Graph API batch endpoint.
        
        Args:
            mailbox: User's email address.
            emails: List of email data dictionaries.
            
        Returns:
            Dictionary with success and failure counts.
        """
        results = {"success": 0, "failed": 0}
        
        # Build batch request
        batch_requests = []
        for i, email in enumerate(emails):
            message = self._build_message_payload(mailbox, email)
            batch_requests.append({
                "id": str(i),
                "method": "POST",
                "url": f"/users/{mailbox}/messages",
                "headers": {"Content-Type": "application/json"},
                "body": message,
            })
        
        batch_payload = {"requests": batch_requests}
        
        try:
            url = f"{self.GRAPH_BASE_URL}/$batch"
            response = self._make_request("POST", url, batch_payload)
            
            # Process responses
            for resp in response.get("responses", []):
                if resp.get("status") in [200, 201]:
                    results["success"] += 1
                    
                    # Handle attachments for successful messages
                    idx = int(resp.get("id", 0))
                    if idx < len(emails) and emails[idx].get("attachments"):
                        message_id = resp.get("body", {}).get("id")
                        if message_id:
                            for attachment in emails[idx]["attachments"]:
                                self._upload_attachment(mailbox, message_id, attachment)
                else:
                    results["failed"] += 1
                    
        except Exception as e:
            # Fall back to individual requests
            for email in emails:
                if self.create_email(mailbox, email):
                    results["success"] += 1
                else:
                    results["failed"] += 1
        
        return results
    
    def _make_request(
        self,
        method: str,
        url: str,
        data: Optional[Dict] = None,
        max_retries: int = 3
    ) -> Dict[str, Any]:
        """
        Make an HTTP request to the Graph API with retry logic.
        
        Implements exponential backoff for handling rate limiting (429 errors)
        and transient failures.
        
        Args:
            method: HTTP method (GET, POST, etc.).
            url: Request URL.
            data: Optional request body data.
            max_retries: Maximum number of retry attempts.
            
        Returns:
            Response data dictionary.
            
        Raises:
            urllib.error.HTTPError: On HTTP errors after all retries exhausted.
        """
        import time
        
        last_exception = None
        
        for attempt in range(max_retries + 1):
            try:
                req = urllib.request.Request(url, method=method)
                req.add_header("Authorization", f"Bearer {self.access_token}")
                req.add_header("Content-Type", "application/json")
                
                if data:
                    req.data = json.dumps(data).encode("utf-8")
                
                with urllib.request.urlopen(req, timeout=60) as response:
                    if response.status == 204:  # No content
                        return {}
                    return json.loads(response.read().decode("utf-8"))
                    
            except urllib.error.HTTPError as e:
                last_exception = e
                
                # Handle rate limiting (429 Too Many Requests)
                if e.code == 429:
                    # Get Retry-After header if available
                    retry_after = e.headers.get("Retry-After", None)
                    if retry_after:
                        try:
                            wait_time = int(retry_after)
                        except ValueError:
                            wait_time = 2 ** attempt  # Exponential backoff
                    else:
                        wait_time = 2 ** attempt  # Exponential backoff: 1, 2, 4 seconds
                    
                    if attempt < max_retries:
                        time.sleep(wait_time)
                        continue
                
                # Handle transient server errors (500, 502, 503, 504)
                elif e.code in [500, 502, 503, 504]:
                    if attempt < max_retries:
                        wait_time = 2 ** attempt
                        time.sleep(wait_time)
                        continue
                
                # For other errors, raise immediately
                raise
                
            except (urllib.error.URLError, TimeoutError) as e:
                last_exception = e
                if attempt < max_retries:
                    wait_time = 2 ** attempt
                    time.sleep(wait_time)
                    continue
                raise
        
        # If we've exhausted all retries, raise the last exception
        if last_exception:
            raise last_exception
        
        return {}
    
    def _should_be_read(self, email_date: datetime) -> bool:
        """
        Determine if email should be marked as read based on age.
        
        Args:
            email_date: Email date.
            
        Returns:
            True if email should be marked as read.
        """
        days_old = (datetime.now() - email_date).days
        
        # Older emails more likely to be read
        if days_old > 30:
            return random.random() < 0.95
        elif days_old > 7:
            return random.random() < 0.80
        elif days_old > 1:
            return random.random() < 0.60
        else:
            return random.random() < 0.30
    
    def _generate_flag_status(self, email_date: datetime, category: str) -> Dict[str, Any]:
        """
        Generate flag status for an email.
        
        Args:
            email_date: Email date.
            category: Email category.
            
        Returns:
            Flag status dictionary for Graph API.
        """
        days_old = (datetime.now() - email_date).days
        
        # Determine if email should be flagged
        # Recent emails more likely to be flagged
        # Certain categories more likely to be flagged
        flag_probability = 0.05  # Base 5% chance
        
        if category in ["organisational", "security"]:
            flag_probability = 0.15  # 15% for important categories
        elif category == "interdepartmental":
            flag_probability = 0.10  # 10% for inter-departmental
        
        # Recent emails more likely to be flagged
        if days_old < 7:
            flag_probability *= 2
        elif days_old > 30:
            flag_probability *= 0.5
        
        if random.random() < flag_probability:
            # Determine flag status
            if days_old > 14:
                # Older flagged emails might be completed
                if random.random() < 0.6:
                    return {"flagStatus": "complete"}
                else:
                    return {"flagStatus": "flagged"}
            else:
                # Recent flagged emails are still active
                if random.random() < 0.8:
                    return {"flagStatus": "flagged"}
                else:
                    # Some have due dates
                    due_date = datetime.now() + timedelta(days=random.randint(1, 7))
                    return {
                        "flagStatus": "flagged",
                        "dueDateTime": {
                            "dateTime": due_date.strftime("%Y-%m-%dT%H:%M:%S"),
                            "timeZone": "UTC"
                        }
                    }
        else:
            return {"flagStatus": "notFlagged"}
    
    def _generate_color_categories(self, category: str) -> List[str]:
        """
        Generate Outlook color categories for an email.
        
        Args:
            category: Email category.
            
        Returns:
            List of color category names.
        """
        # Outlook preset categories
        color_categories = {
            "organisational": ["Important", "Company", "Action Required"],
            "security": ["Urgent", "Security", "Action Required"],
            "interdepartmental": ["Project", "Team", "Collaboration"],
            "newsletters": ["Newsletter", "FYI", "Reading"],
            "attachments": ["Documents", "Review", "Files"],
            "links": ["SharePoint", "Documents", "Links"],
        }
        
        # Only 15% of emails get color categories
        if random.random() > 0.15:
            return []
        
        # Get relevant categories for this email type
        relevant = color_categories.get(category, ["General"])
        
        # Pick 1-2 categories
        num_categories = random.choice([1, 1, 1, 2])  # Usually just 1
        return random.sample(relevant, min(num_categories, len(relevant)))
    
    def _get_inference_classification(self, sender: Dict[str, str], category: str) -> str:
        """
        Determine if email should be in Focused or Other inbox.
        
        Args:
            sender: Sender information.
            category: Email category.
            
        Returns:
            "focused" or "other"
        """
        # Newsletters and external senders more likely to be in "Other"
        if category == "newsletters":
            return "other" if random.random() < 0.8 else "focused"
        
        # Security and organisational emails always focused
        if category in ["security", "organisational"]:
            return "focused"
        
        # Check if external sender
        sender_email = sender.get("email", "")
        if "newsletter" in sender_email.lower() or "noreply" in sender_email.lower():
            return "other"
        
        # Most internal emails are focused
        return "focused" if random.random() < 0.85 else "other"
    
    def _get_token_client_credentials(self, app_config: Dict[str, Any]) -> Optional[str]:
        """
        Get token using client credentials flow.
        
        Args:
            app_config: App configuration with credentials.
            
        Returns:
            Access token or None.
        """
        tenant_id = app_config.get("tenant_id")
        client_id = app_config.get("app_id")
        client_secret = app_config.get("client_secret")
        
        if not all([tenant_id, client_id, client_secret]):
            return None
        
        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        
        data = urllib.parse.urlencode({
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        }).encode("utf-8")
        
        try:
            req = urllib.request.Request(token_url, data=data)
            req.add_header("Content-Type", "application/x-www-form-urlencoded")
            
            with urllib.request.urlopen(req, timeout=30) as response:
                result = json.loads(response.read().decode("utf-8"))
                return result.get("access_token")
        except Exception:
            return None
    
    def _get_token_azure_cli(self) -> Optional[str]:
        """
        Get token using Azure CLI.
        
        Returns:
            Access token or None.
        """
        az_path = self._find_azure_cli_path()
        if not az_path:
            return None
        
        try:
            result = subprocess.run(
                [az_path, "account", "get-access-token",
                 "--resource", "https://graph.microsoft.com",
                 "--query", "accessToken", "-o", "tsv"],
                capture_output=True,
                text=True,
                timeout=30,
            )
            
            if result.returncode == 0:
                return result.stdout.strip()
        except Exception:
            pass
        
        return None
    
    def _find_azure_cli_path(self) -> Optional[str]:
        """Find Azure CLI path."""
        if self._az_path:
            return self._az_path
        
        # Check if 'az' is in PATH
        az_path = shutil.which("az")
        if az_path:
            self._az_path = az_path
            return az_path
        
        # On Windows, check .cmd extension and default paths
        if platform.system().lower() == "windows":
            az_cmd_path = shutil.which("az.cmd")
            if az_cmd_path:
                self._az_path = az_cmd_path
                return az_cmd_path
            
            for path in AZURE_CLI_PATHS:
                if os.path.exists(path):
                    self._az_path = path
                    return path
        
        return None
    
    def _load_app_config(self) -> Optional[Dict[str, Any]]:
        """Load app configuration from file."""
        if APP_CONFIG_FILE.exists():
            try:
                with open(APP_CONFIG_FILE, 'r') as f:
                    return json.load(f)
            except Exception:
                pass
        return None
