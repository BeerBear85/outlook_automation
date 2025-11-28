"""
Microsoft Graph Authentication Module

Handles authentication to Microsoft Graph using MSAL (Microsoft Authentication Library).
Supports device code flow and interactive browser flow with automatic token caching.

@author: Generated for outlook_automation repository (Python Graph implementation)
"""

import os
import json
import msal
from typing import Optional, Dict, Any
from pathlib import Path


class GraphAuthenticator:
    """
    Handles Microsoft Graph authentication with token caching.

    Uses device code flow by default (user-friendly for CLI).
    Automatically caches and refreshes tokens.
    """

    def __init__(self, client_id: str, tenant_id: str = "common", scopes: list = None):
        """
        Initialize the authenticator.

        Args:
            client_id: Azure AD application (client) ID
            tenant_id: Azure AD tenant ID or "common" for multi-tenant
            scopes: List of permission scopes (e.g., ["Calendars.ReadWrite"])
        """
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.scopes = scopes or [
            "Calendars.ReadWrite",
            "Mail.ReadWrite",
            "User.Read"
        ]

        # Authority URL
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"

        # Token cache file location
        cache_dir = Path.home() / ".outlook_automation"
        cache_dir.mkdir(exist_ok=True)
        self.cache_file = cache_dir / "token_cache.json"

        # Initialize token cache
        self.token_cache = msal.SerializableTokenCache()
        if self.cache_file.exists():
            self.token_cache.deserialize(self.cache_file.read_text())

        # Create public client application
        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=self.token_cache
        )

    def _save_cache(self):
        """Save token cache to disk if it has changed."""
        if self.token_cache.has_state_changed:
            self.cache_file.write_text(self.token_cache.serialize())

    def get_access_token(self, use_device_flow: bool = True) -> Optional[str]:
        """
        Get a valid access token, using cached token if available.

        Args:
            use_device_flow: If True, use device code flow. If False, use interactive browser flow.

        Returns:
            Access token string, or None if authentication failed.
        """
        # First, try to get token from cache
        accounts = self.app.get_accounts()
        if accounts:
            # Try silent authentication with cached token
            result = self.app.acquire_token_silent(self.scopes, account=accounts[0])
            if result and "access_token" in result:
                print(f"✓ Using cached token for account: {accounts[0]['username']}")
                return result["access_token"]

        # No cached token, need interactive authentication
        if use_device_flow:
            return self._authenticate_device_flow()
        else:
            return self._authenticate_interactive()

    def _authenticate_device_flow(self) -> Optional[str]:
        """
        Authenticate using device code flow.
        User will be prompted to visit a URL and enter a code.

        Returns:
            Access token string, or None if authentication failed.
        """
        print("\n" + "="*70)
        print("Microsoft Graph Authentication (Device Code Flow)")
        print("="*70)

        # Initiate device flow
        flow = self.app.initiate_device_flow(scopes=self.scopes)

        if "user_code" not in flow:
            print(f"✗ Failed to create device flow: {flow.get('error_description', 'Unknown error')}")
            return None

        # Display instructions to user
        print(f"\n{flow['message']}\n")
        print("Waiting for authentication...")

        # Wait for user to authenticate
        result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            self._save_cache()
            print(f"✓ Authentication successful!")
            print(f"✓ Signed in as: {result.get('id_token_claims', {}).get('preferred_username', 'Unknown')}")
            print("="*70 + "\n")
            return result["access_token"]
        else:
            error = result.get("error_description", result.get("error", "Unknown error"))
            print(f"✗ Authentication failed: {error}")
            return None

    def _authenticate_interactive(self) -> Optional[str]:
        """
        Authenticate using interactive browser flow.
        Opens browser for user to sign in.

        Returns:
            Access token string, or None if authentication failed.
        """
        print("\n" + "="*70)
        print("Microsoft Graph Authentication (Interactive Browser Flow)")
        print("="*70)
        print("\nA browser window will open for authentication...")

        result = self.app.acquire_token_interactive(scopes=self.scopes)

        if "access_token" in result:
            self._save_cache()
            print(f"✓ Authentication successful!")
            print(f"✓ Signed in as: {result.get('id_token_claims', {}).get('preferred_username', 'Unknown')}")
            print("="*70 + "\n")
            return result["access_token"]
        else:
            error = result.get("error_description", result.get("error", "Unknown error"))
            print(f"✗ Authentication failed: {error}")
            return None

    def get_account_info(self) -> Optional[Dict[str, Any]]:
        """
        Get information about the currently authenticated account.

        Returns:
            Dictionary with account info, or None if not authenticated.
        """
        accounts = self.app.get_accounts()
        if accounts:
            return accounts[0]
        return None

    def clear_cache(self):
        """Clear the token cache (force re-authentication on next call)."""
        if self.cache_file.exists():
            self.cache_file.unlink()
        self.token_cache = msal.SerializableTokenCache()
        print("✓ Token cache cleared")


class GraphClient:
    """
    HTTP client for Microsoft Graph API with automatic authentication.
    """

    def __init__(self, authenticator: GraphAuthenticator):
        """
        Initialize Graph API client.

        Args:
            authenticator: GraphAuthenticator instance
        """
        self.authenticator = authenticator
        self.base_url = "https://graph.microsoft.com/v1.0"
        self._access_token = None

    def _ensure_token(self):
        """Ensure we have a valid access token."""
        if not self._access_token:
            self._access_token = self.authenticator.get_access_token()
            if not self._access_token:
                raise Exception("Failed to acquire access token")

    def get_headers(self) -> Dict[str, str]:
        """
        Get HTTP headers with authorization.

        Returns:
            Dictionary of HTTP headers
        """
        self._ensure_token()
        return {
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }

    def refresh_token(self):
        """Force token refresh on next API call."""
        self._access_token = None


def create_authenticator_from_config(config: Dict[str, Any]) -> GraphAuthenticator:
    """
    Create a GraphAuthenticator from configuration dictionary.

    Args:
        config: Configuration dictionary with client_id, tenant_id, scopes

    Returns:
        GraphAuthenticator instance
    """
    return GraphAuthenticator(
        client_id=config.get("client_id"),
        tenant_id=config.get("tenant_id", "common"),
        scopes=config.get("scopes", ["Calendars.ReadWrite", "Mail.ReadWrite", "User.Read"])
    )
