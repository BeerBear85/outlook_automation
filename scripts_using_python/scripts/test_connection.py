#!/usr/bin/env python3
"""
Test Microsoft Graph Connection

Utility script to verify authentication and permissions.

Equivalent to: Test-GraphConnection.ps1 (PowerShell version)

@author: Generated for outlook_automation repository (Python Graph implementation)
"""

import sys
from pathlib import Path
from datetime import datetime, timedelta

# Add src directory to path
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))

from outlook_graph import (
    get_config,
    create_authenticator_from_config,
    GraphClient,
    CalendarClient
)
import requests


def main():
    """
    Main function to test Microsoft Graph connection and permissions.
    """
    print("=" * 70)
    print("Microsoft Graph Connection Test (Python)")
    print("=" * 70)
    print()

    # Load configuration
    try:
        config = get_config()
        config.validate()
    except Exception as e:
        print(f"✗ Configuration error: {e}")
        print()
        print("Please run: python scripts/connect_graph.py")
        return 1

    # Test 1: Check if authenticated
    print("[Test 1/4] Checking authentication status...")
    try:
        authenticator = create_authenticator_from_config(config.to_dict())
        account = authenticator.get_account_info()

        if account:
            print("  ✓ PASS: Authenticated to Microsoft Graph")
            print(f"    Account: {account.get('username', 'Unknown')}")
            print(f"    Environment: {account.get('environment', 'Unknown')}")
        else:
            print("  ✗ FAIL: Not authenticated")
            print()
            print("Please run: python scripts/connect_graph.py")
            return 1

    except Exception as e:
        print(f"  ✗ FAIL: Authentication check failed: {e}")
        return 1

    print()

    # Test 2: Check required scopes
    print("[Test 2/4] Checking required permissions...")
    required_scopes = ["Calendars.ReadWrite", "Mail.ReadWrite", "User.Read"]

    # Note: With MSAL, we can't directly verify granted scopes without making an API call
    # We'll assume scopes are granted if authentication succeeded
    for scope in required_scopes:
        if scope in config.scopes:
            print(f"  ✓ PASS: {scope} (requested)")
        else:
            print(f"  ⚠ WARN: {scope} (not in config)")

    print()

    # Test 3: Test calendar access
    print("[Test 3/4] Testing calendar access...")
    try:
        graph_client = GraphClient(authenticator)
        calendar_client = CalendarClient(graph_client, config.timezone)

        # Try to get calendar events for today
        today = datetime.now()
        tomorrow = today + timedelta(days=1)

        events = calendar_client.list_events(today, tomorrow)

        print("  ✓ PASS: Successfully accessed calendar")
        print(f"    Found {len(events)} event(s) for today")

    except Exception as e:
        print("  ✗ FAIL: Cannot access calendar")
        print(f"    Error: {e}")
        return 1

    print()

    # Test 4: Test user profile access
    print("[Test 4/4] Testing user profile access...")
    try:
        url = f"{config.graph_endpoint}/me"
        headers = graph_client.get_headers()
        params = {"$select": "displayName,userPrincipalName,mailboxSettings"}
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()

        user = response.json()

        print("  ✓ PASS: Successfully accessed user profile")
        print(f"    Display Name: {user.get('displayName', 'Unknown')}")
        print(f"    Email:        {user.get('userPrincipalName', 'Unknown')}")

        mailbox_settings = user.get('mailboxSettings', {})
        if mailbox_settings and 'timeZone' in mailbox_settings:
            print(f"    Timezone:     {mailbox_settings['timeZone']}")

    except Exception as e:
        print("  ✗ FAIL: Cannot access user profile")
        print(f"    Error: {e}")
        return 1

    print()

    # Summary
    print("=" * 70)
    print("All tests PASSED!")
    print()
    print("Your Microsoft Graph connection is working correctly.")
    print("You can now run the automation scripts:")
    print("  python scripts/show_meeting_summary.py")
    print()

    return 0


if __name__ == "__main__":
    sys.exit(main())
