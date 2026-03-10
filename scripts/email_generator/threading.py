"""
Email thread management for realistic conversation chains.

Creates reply chains, forward chains, and reply-all threads
to simulate realistic email conversations.
"""

import random
import uuid
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any


class ThreadManager:
    """Manages email threading for realistic conversations."""
    
    def __init__(self, config: Dict[str, Any]):
        """
        Initialize the thread manager.
        
        Args:
            config: Mailbox configuration dictionary.
        """
        self.config = config
        self.settings = config.get("settings", {})
        self.threading_config = self.settings.get("threading", {})
        
        # Track active threads for potential continuation
        self.active_threads: Dict[str, List[Dict]] = {}
    
    def should_create_thread(self) -> bool:
        """
        Determine if this email should be part of a thread.
        
        Returns:
            True if email should be threaded, False for standalone.
        """
        if not self.threading_config.get("enabled", True):
            return False
        
        single_percentage = self.threading_config.get("single_email_percentage", 60)
        return random.random() > (single_percentage / 100)
    
    def get_thread_type(self) -> str:
        """
        Select the type of thread to create.
        
        Returns:
            Thread type: 'reply', 'forward', or 'reply_all'.
        """
        reply_pct = self.threading_config.get("reply_chain_percentage", 25)
        forward_pct = self.threading_config.get("forward_chain_percentage", 10)
        reply_all_pct = self.threading_config.get("reply_all_percentage", 5)
        
        total = reply_pct + forward_pct + reply_all_pct
        rand = random.random() * total
        
        if rand < reply_pct:
            return "reply"
        elif rand < reply_pct + forward_pct:
            return "forward"
        else:
            return "reply_all"
    
    def create_thread(self, email: Dict[str, Any], recipient: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create or continue an email thread.
        
        Args:
            email: Base email dictionary.
            recipient: Recipient user configuration.
            
        Returns:
            Modified email with thread information.
        """
        thread_type = self.get_thread_type()
        
        if thread_type == "reply":
            return self._create_reply_chain(email, recipient)
        elif thread_type == "forward":
            return self._create_forward_chain(email, recipient)
        else:
            return self._create_reply_all_chain(email, recipient)
    
    def _create_reply_chain(self, email: Dict[str, Any], recipient: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create a reply chain with 2-5 emails.
        
        Args:
            email: Base email dictionary.
            recipient: Recipient user configuration.
            
        Returns:
            Email with reply chain content.
        """
        max_depth = self.threading_config.get("max_thread_depth", 5)
        chain_length = random.randint(2, min(max_depth, 5))
        
        # Modify subject to indicate reply
        original_subject = email.get("subject", "")
        if not original_subject.startswith("Re:"):
            email["subject"] = f"Re: {original_subject}"
        
        # Generate conversation ID
        email["conversation_id"] = self._generate_conversation_id()
        email["thread_type"] = "reply"
        email["thread_depth"] = chain_length
        
        # Add quoted content to body
        quoted_content = self._generate_quoted_chain(email, recipient, chain_length)
        email["body"] = self._add_quoted_content(email["body"], quoted_content)
        
        return email
    
    def _create_forward_chain(self, email: Dict[str, Any], recipient: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create a forward chain.
        
        Args:
            email: Base email dictionary.
            recipient: Recipient user configuration.
            
        Returns:
            Email with forward chain content.
        """
        # Modify subject to indicate forward
        original_subject = email.get("subject", "")
        if not original_subject.startswith("Fwd:") and not original_subject.startswith("FW:"):
            email["subject"] = f"Fwd: {original_subject}"
        
        # Generate conversation ID
        email["conversation_id"] = self._generate_conversation_id()
        email["thread_type"] = "forward"
        email["thread_depth"] = 1
        
        # Add forwarded content
        forwarded_content = self._generate_forwarded_content(email, recipient)
        email["body"] = self._add_forwarded_content(email["body"], forwarded_content)
        
        return email
    
    def _create_reply_all_chain(self, email: Dict[str, Any], recipient: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create a reply-all chain with multiple recipients.
        
        Args:
            email: Base email dictionary.
            recipient: Recipient user configuration.
            
        Returns:
            Email with reply-all chain content.
        """
        # Similar to reply chain but with CC recipients
        email = self._create_reply_chain(email, recipient)
        email["thread_type"] = "reply_all"
        
        # Add CC recipients (other users from config)
        users = self.config.get("users", [])
        other_users = [u for u in users if u.get("upn") != recipient.get("upn")]
        
        if other_users:
            # Add 1-3 CC recipients
            num_cc = min(random.randint(1, 3), len(other_users))
            cc_users = random.sample(other_users, num_cc)
            email["cc_recipients"] = [
                {
                    "email": u.get("upn"),
                    "name": self._get_display_name(u),
                }
                for u in cc_users
            ]
        
        return email
    
    def _generate_quoted_chain(
        self,
        email: Dict[str, Any],
        recipient: Dict[str, Any],
        depth: int
    ) -> List[Dict[str, Any]]:
        """
        Generate quoted email chain content.
        
        Args:
            email: Current email dictionary.
            recipient: Recipient user configuration.
            depth: Number of emails in the chain.
            
        Returns:
            List of quoted email dictionaries.
        """
        quoted_emails = []
        current_date = email.get("date", datetime.now())
        
        # Get participants
        sender = email.get("sender", {})
        participants = [
            {"name": sender.get("name", "Sender"), "email": sender.get("email", "")},
            {"name": recipient.get("display_name", "Recipient"), "email": recipient.get("upn", "")},
        ]
        
        for i in range(depth - 1):
            # Alternate between participants
            prev_sender = participants[i % 2]
            
            # Generate previous email date (earlier)
            prev_date = current_date - timedelta(
                hours=random.randint(1, 48),
                minutes=random.randint(0, 59)
            )
            
            # Generate response content
            response = self._generate_thread_response(i, depth)
            
            quoted_emails.append({
                "sender_name": prev_sender["name"],
                "sender_email": prev_sender["email"],
                "date": prev_date,
                "content": response,
            })
            
            current_date = prev_date
        
        return quoted_emails
    
    def _generate_thread_response(self, position: int, total: int) -> str:
        """
        Generate appropriate response content for thread position.
        
        Args:
            position: Position in the thread (0 = most recent previous).
            total: Total number of emails in thread.
            
        Returns:
            Response content string.
        """
        # Different responses based on position in thread
        if position == 0:
            # Most recent previous email
            responses = [
                "Thanks for the update. I'll review and get back to you.",
                "Good progress! Let's discuss this in our next meeting.",
                "I've reviewed the attached document. A few comments below.",
                "Agreed. I'll coordinate with the team on next steps.",
                "Thanks for flagging this. I'll follow up with the relevant stakeholders.",
                "Sounds good. Let me know if you need anything else from my end.",
                "I'll take a look at this today and share my thoughts.",
            ]
        elif position == total - 2:
            # Original email in thread
            responses = [
                "Hi team, I wanted to share an update on the project status.",
                "Following up on our discussion from last week.",
                "I've put together some thoughts on this topic for your review.",
                "As discussed, here's the information you requested.",
                "I'd like to get your input on the following items.",
            ]
        else:
            # Middle of thread
            responses = [
                "Thanks for sharing. I have a few questions.",
                "This looks good overall. One suggestion:",
                "I've looped in the team for their input.",
                "Good point. Let me check on that.",
                "Noted. I'll update the document accordingly.",
            ]
        
        return random.choice(responses)
    
    def _generate_forwarded_content(
        self,
        email: Dict[str, Any],
        recipient: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Generate forwarded email content.
        
        Args:
            email: Current email dictionary.
            recipient: Recipient user configuration.
            
        Returns:
            Forwarded email content dictionary.
        """
        # Generate original sender (different from current sender)
        original_senders = [
            {"name": "John Smith", "email": "john.smith@company.com"},
            {"name": "Sarah Johnson", "email": "sarah.johnson@company.com"},
            {"name": "Michael Chen", "email": "michael.chen@company.com"},
            {"name": "Emily Davis", "email": "emily.davis@company.com"},
        ]
        original_sender = random.choice(original_senders)
        
        # Generate original date
        current_date = email.get("date", datetime.now())
        original_date = current_date - timedelta(
            days=random.randint(1, 7),
            hours=random.randint(0, 23)
        )
        
        # Generate original content
        original_contents = [
            "Please see the attached document for your review. Let me know if you have any questions.",
            "I wanted to share this information with you. It may be relevant to our upcoming project.",
            "Here's the update I mentioned in our meeting. Please review when you get a chance.",
            "FYI - this came through today and I thought you should be aware.",
        ]
        
        return {
            "sender_name": original_sender["name"],
            "sender_email": original_sender["email"],
            "date": original_date,
            "content": random.choice(original_contents),
        }
    
    def _add_quoted_content(self, body: str, quoted_emails: List[Dict]) -> str:
        """
        Add quoted email content to the body.
        
        Args:
            body: Current email body HTML.
            quoted_emails: List of quoted email dictionaries.
            
        Returns:
            Body with quoted content appended.
        """
        if not quoted_emails:
            return body
        
        quoted_html = ""
        for quoted in quoted_emails:
            date_str = quoted["date"].strftime("%B %d, %Y at %I:%M %p")
            quoted_html += f'''
<div style="border-left: 3px solid #ccc; padding-left: 15px; margin: 20px 0; color: #666;">
    <p style="margin: 0 0 10px 0;"><strong>On {date_str}, {quoted["sender_name"]} &lt;{quoted["sender_email"]}&gt; wrote:</strong></p>
    <p style="margin: 0;">{quoted["content"]}</p>
</div>
'''
        
        # Insert before closing body/html tags or append
        if "</div>" in body and body.rstrip().endswith("</body></html>"):
            # Find last content div and insert after
            insert_pos = body.rfind("</div>")
            if insert_pos > 0:
                return body[:insert_pos + 6] + quoted_html + body[insert_pos + 6:]
        
        # Append to body
        if "</body>" in body:
            return body.replace("</body>", quoted_html + "</body>")
        
        return body + quoted_html
    
    def _add_forwarded_content(self, body: str, forwarded: Dict) -> str:
        """
        Add forwarded email content to the body.
        
        Args:
            body: Current email body HTML.
            forwarded: Forwarded email dictionary.
            
        Returns:
            Body with forwarded content appended.
        """
        date_str = forwarded["date"].strftime("%A, %B %d, %Y %I:%M %p")
        
        forwarded_html = f'''
<div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ccc;">
    <p style="color: #666; margin-bottom: 15px;"><strong>---------- Forwarded message ---------</strong></p>
    <p style="color: #666; margin: 5px 0;"><strong>From:</strong> {forwarded["sender_name"]} &lt;{forwarded["sender_email"]}&gt;</p>
    <p style="color: #666; margin: 5px 0;"><strong>Date:</strong> {date_str}</p>
    <p style="color: #666; margin: 5px 0;"><strong>Subject:</strong> {forwarded.get("subject", "Original Message")}</p>
    <div style="margin-top: 15px;">
        <p>{forwarded["content"]}</p>
    </div>
</div>
'''
        
        # Insert before closing body/html tags
        if "</body>" in body:
            return body.replace("</body>", forwarded_html + "</body>")
        
        return body + forwarded_html
    
    def _generate_conversation_id(self) -> str:
        """Generate a unique conversation ID."""
        return str(uuid.uuid4())
    
    def _get_display_name(self, user: Dict[str, Any]) -> str:
        """Get display name for a user."""
        if "display_name" in user:
            return user["display_name"]
        
        upn = user.get("upn", "")
        name_part = upn.split("@")[0]
        parts = name_part.replace(".", " ").replace("_", " ").replace("-", " ").split()
        return " ".join(part.capitalize() for part in parts)
