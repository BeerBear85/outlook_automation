#!/usr/bin/env python3
"""
Connect to Microsoft Graph for Outlook Automation

Interactive authentication script for Microsoft Graph API using MSAL.

Equivalent to: Connect-Graph.ps1 (PowerShell version)

@author: Generated for outlook_automation repository (Python Graph implementation)
"""

import sys
from pathlib import Path

# Add src directory to path
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))

from outlook_graph import get_config, create_authenticator_from_config, GraphClient
import requests


def main():
    """
    Main function for authentication to Microsoft Graph.
    """
    print("=" * 70)
    print("Microsoft Graph Connection Script (Python)")
    print("=" * 70)
    print()

    # Step 1: Check prerequisites
    print("[1/4] Checking prerequisites...")
    try:
        import msal
        print(f"  ✓ Found msal version {msal.__version__}")
    except ImportError:
        print("  ✗ ERROR: msal library not found!")
        print()
        print("Please install required packages with:")
        print("  pip install -r requirements.txt")
        print()
        return 1
    print()

    # Step 2: Load configuration
    print("[2/4] Loading configuration...")
    try:
        config = get_config()
        config.validate()
        print(f"  ✓ Configuration loaded")
        print(f"    Tenant ID: {config.tenant_id}")
        print(f"    Client ID: {config.client_id[:8]}...")
        print(f"    Scopes: {', '.join(config.scopes)}")
    except Exception as e:
        print(f"  ✗ ERROR: Configuration error: {e}")
        print()
        print("Please ensure you have:")
        print("  1. Created a config.json file with your client_id, OR")
        print("  2. Set GRAPH_CLIENT_ID environment variable")
        print()
        print("See README.md for setup instructions.")
        return 1
    print()

    # Step 3: Authenticate to Microsoft Graph
    print("[3/4] Connecting to Microsoft Graph...")
    print("  A browser window or device code prompt will appear.")
    print("  Please sign in with your Microsoft account.")
    print()

    try:
        # Create authenticator
        authenticator = create_authenticator_from_config(config.to_dict())

        # Get access token (will prompt for authentication)
        token = authenticator.get_access_token(use_device_flow=True)

        if not token:
            print("  ✗ ERROR: Failed to authenticate")
            return 1

        print("  ✓ Successfully connected!")
        print()

    except Exception as e:
        print(f"  ✗ ERROR: Failed to connect to Microsoft Graph")
        print(f"  {e}")
        print()
        return 1

    # Step 4: Validate connection
    print("[4/4] Validating connection...")
    try:
        # Create Graph client
        graph_client = GraphClient(authenticator)

        # Test connection by getting user profile
        url = f"{config.graph_endpoint}/me"
        headers = graph_client.get_headers()
        response = requests.get(url, headers=headers)
        response.raise_for_status()

        user = response.json()

        print("  ✓ Connection validated successfully!")
        print()
        print("Connection Details:")
        print("=" * 70)
        print(f"  Display Name: {user.get('displayName', 'Unknown')}")
        print(f"  Email:        {user.get('userPrincipalName', 'Unknown')}")
        print(f"  User ID:      {user.get('id', 'Unknown')}")
        print()

        # Get account info from cache
        account = authenticator.get_account_info()
        if account:
            print(f"  Cached Account: {account.get('username', 'Unknown')}")
        print()

        print("SUCCESS: Ready to run Graph-based automation scripts!")
        print()
        print("You can now run:")
        print("  python scripts/show_meeting_summary.py")
        print("  python scripts/test_connection.py")
        print()

        return 0

    except Exception as e:
        print("  ✗ ERROR: Connection validation failed")
        print(f"  {e}")
        print()
        print("Please try running this script again.")
        print()
        return 1


if __name__ == "__main__":
    sys.exit(main())
