"""
Mail Operations Module

Provides functions for interacting with Microsoft Graph Mail API:
- Send email messages
- Create draft emails
- Manage messages

@author: Generated for outlook_automation repository (Python Graph implementation)
"""

import requests
import logging
from typing import List, Dict, Any, Optional
from .auth import GraphClient


logger = logging.getLogger("outlook_automation")


class MailClient:
    """
    Client for Microsoft Graph Mail API operations.
    """

    def __init__(self, graph_client: GraphClient):
        """
        Initialize mail client.

        Args:
            graph_client: Authenticated GraphClient instance
        """
        self.graph_client = graph_client
        self.base_url = f"{graph_client.base_url}/me"

    def send_mail(
        self,
        to: str,
        subject: str,
        body: str,
        cc: Optional[List[str]] = None,
        bcc: Optional[List[str]] = None,
        content_type: str = "Text"
    ) -> bool:
        """
        Send an email message.

        Args:
            to: Recipient email address
            subject: Email subject
            body: Email body content
            cc: List of CC email addresses (optional)
            bcc: List of BCC email addresses (optional)
            content_type: Content type - "Text" or "HTML" (default: "Text")

        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Sending email to: {to}")

            # Build message structure
            message = {
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": content_type,
                        "content": body
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": to
                            }
                        }
                    ]
                }
            }

            # Add CC recipients if provided
            if cc:
                message["message"]["ccRecipients"] = [
                    {"emailAddress": {"address": addr}} for addr in cc
                ]

            # Add BCC recipients if provided
            if bcc:
                message["message"]["bccRecipients"] = [
                    {"emailAddress": {"address": addr}} for addr in bcc
                ]

            # Send mail
            url = f"{self.base_url}/sendMail"
            headers = self.graph_client.get_headers()
            response = requests.post(url, headers=headers, json=message)
            response.raise_for_status()

            logger.info(f"Email sent successfully to: {to}")
            return True

        except requests.exceptions.HTTPError as e:
            logger.error(f"HTTP error sending email: {e}")
            logger.error(f"Response: {e.response.text if e.response else 'No response'}")
            return False
        except Exception as e:
            logger.error(f"Error sending email: {e}")
            return False

    def create_draft(
        self,
        to: str,
        subject: str,
        body: str,
        cc: Optional[List[str]] = None,
        bcc: Optional[List[str]] = None,
        content_type: str = "Text"
    ) -> Optional[Dict[str, Any]]:
        """
        Create a draft email in Outlook (not sent).

        The draft is saved to the Drafts folder for the user to review and send manually.

        Args:
            to: Recipient email address
            subject: Email subject
            body: Email body content
            cc: List of CC email addresses (optional)
            bcc: List of BCC email addresses (optional)
            content_type: Content type - "Text" or "HTML" (default: "Text")

        Returns:
            Created message dictionary, or None if failed
        """
        try:
            logger.info(f"Creating draft email to: {to}")

            # Build message structure
            message = {
                "subject": subject,
                "body": {
                    "contentType": content_type,
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to
                        }
                    }
                ]
            }

            # Add CC recipients if provided
            if cc:
                message["ccRecipients"] = [
                    {"emailAddress": {"address": addr}} for addr in cc
                ]

            # Add BCC recipients if provided
            if bcc:
                message["bccRecipients"] = [
                    {"emailAddress": {"address": addr}} for addr in bcc
                ]

            # Create draft (POST to /me/messages without sending)
            url = f"{self.base_url}/messages"
            headers = self.graph_client.get_headers()
            response = requests.post(url, headers=headers, json=message)
            response.raise_for_status()

            draft = response.json()
            logger.info(f"Draft email created successfully (ID: {draft.get('id', 'unknown')})")
            return draft

        except requests.exceptions.HTTPError as e:
            logger.error(f"HTTP error creating draft email: {e}")
            logger.error(f"Response: {e.response.text if e.response else 'No response'}")
            return None
        except Exception as e:
            logger.error(f"Error creating draft email: {e}")
            return None

    def get_message(self, message_id: str) -> Optional[Dict[str, Any]]:
        """
        Get a single email message by ID.

        Args:
            message_id: Message ID from Microsoft Graph

        Returns:
            Message dictionary, or None if not found
        """
        try:
            url = f"{self.base_url}/messages/{message_id}"
            headers = self.graph_client.get_headers()
            response = requests.get(url, headers=headers)

            if response.status_code == 404:
                logger.warning(f"Message not found: {message_id}")
                return None

            response.raise_for_status()
            return response.json()

        except Exception as e:
            logger.error(f"Error retrieving message {message_id}: {e}")
            return None

    def delete_message(self, message_id: str) -> bool:
        """
        Delete an email message.

        Args:
            message_id: Message ID from Microsoft Graph

        Returns:
            True if successful, False otherwise
        """
        try:
            url = f"{self.base_url}/messages/{message_id}"
            headers = self.graph_client.get_headers()
            response = requests.delete(url, headers=headers)
            response.raise_for_status()

            logger.info(f"Deleted message: {message_id}")
            return True

        except Exception as e:
            logger.error(f"Error deleting message {message_id}: {e}")
            return False

    def send_draft(self, message_id: str) -> bool:
        """
        Send a draft message.

        Args:
            message_id: Draft message ID from Microsoft Graph

        Returns:
            True if successful, False otherwise
        """
        try:
            url = f"{self.base_url}/messages/{message_id}/send"
            headers = self.graph_client.get_headers()
            response = requests.post(url, headers=headers)
            response.raise_for_status()

            logger.info(f"Sent draft message: {message_id}")
            return True

        except Exception as e:
            logger.error(f"Error sending draft message {message_id}: {e}")
            return False
