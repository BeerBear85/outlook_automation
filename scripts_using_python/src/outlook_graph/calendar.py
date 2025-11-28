"""
Calendar Operations Module

Provides functions for interacting with Microsoft Graph Calendar API:
- List calendar events
- Get single event
- Update event
- Create event
- Filter and process events

@author: Generated for outlook_automation repository (Python Graph implementation)
"""

import requests
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional
from .auth import GraphClient
from .utils import to_iso8601, from_graph_datetime


logger = logging.getLogger("outlook_automation")


class CalendarClient:
    """
    Client for Microsoft Graph Calendar API operations.
    """

    def __init__(self, graph_client: GraphClient, timezone: str = "Europe/Copenhagen"):
        """
        Initialize calendar client.

        Args:
            graph_client: Authenticated GraphClient instance
            timezone: Timezone for date/time conversions
        """
        self.graph_client = graph_client
        self.timezone = timezone
        self.base_url = f"{graph_client.base_url}/me"

    def list_events(self, start_date: datetime, end_date: datetime) -> List[Dict[str, Any]]:
        """
        Retrieve calendar events from Microsoft Graph for a date range.

        Uses /me/calendarView endpoint to get events in a specific time range.

        Args:
            start_date: Start of date range (timezone-aware)
            end_date: End of date range (timezone-aware)

        Returns:
            List of event dictionaries from Graph API
        """
        try:
            # Format dates for Graph API (ISO 8601)
            start_str = to_iso8601(start_date)
            end_str = to_iso8601(end_date)

            logger.info(f"Retrieving calendar events from {start_str} to {end_str}")

            # Build request URL
            url = f"{self.base_url}/calendarView"
            params = {
                "startDateTime": start_str,
                "endDateTime": end_str,
                "$top": 1000  # Maximum results per page
            }

            # Make request
            headers = self.graph_client.get_headers()
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()

            data = response.json()
            events = data.get("value", [])

            # Handle pagination if needed
            while "@odata.nextLink" in data:
                logger.info(f"Fetching next page of events...")
                response = requests.get(data["@odata.nextLink"], headers=headers)
                response.raise_for_status()
                data = response.json()
                events.extend(data.get("value", []))

            logger.info(f"Retrieved {len(events)} calendar events from Microsoft Graph")
            return events

        except requests.exceptions.HTTPError as e:
            logger.error(f"HTTP error retrieving calendar events: {e}")
            logger.error(f"Response: {e.response.text if e.response else 'No response'}")
            raise Exception(f"Failed to retrieve calendar events: {e}")
        except Exception as e:
            logger.error(f"Error retrieving calendar events: {e}")
            raise Exception(f"Failed to retrieve calendar events: {e}")

    def get_event(self, event_id: str) -> Optional[Dict[str, Any]]:
        """
        Get a single calendar event by ID.

        Args:
            event_id: Event ID from Microsoft Graph

        Returns:
            Event dictionary, or None if not found
        """
        try:
            url = f"{self.base_url}/events/{event_id}"
            headers = self.graph_client.get_headers()
            response = requests.get(url, headers=headers)

            if response.status_code == 404:
                logger.warning(f"Event not found: {event_id}")
                return None

            response.raise_for_status()
            return response.json()

        except Exception as e:
            logger.error(f"Error retrieving event {event_id}: {e}")
            return None

    def update_event(self, event_id: str, updates: Dict[str, Any]) -> bool:
        """
        Update a calendar event.

        Args:
            event_id: Event ID from Microsoft Graph
            updates: Dictionary of fields to update

        Returns:
            True if successful, False otherwise
        """
        try:
            url = f"{self.base_url}/events/{event_id}"
            headers = self.graph_client.get_headers()
            response = requests.patch(url, headers=headers, json=updates)
            response.raise_for_status()

            logger.info(f"Updated event: {event_id}")
            return True

        except Exception as e:
            logger.error(f"Error updating event {event_id}: {e}")
            return False

    def create_event(self, event_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """
        Create a new calendar event.

        Args:
            event_data: Event data dictionary (Graph API format)

        Returns:
            Created event dictionary, or None if failed
        """
        try:
            url = f"{self.base_url}/events"
            headers = self.graph_client.get_headers()
            response = requests.post(url, headers=headers, json=event_data)
            response.raise_for_status()

            created_event = response.json()
            logger.info(f"Created event: {created_event.get('subject', 'Untitled')}")
            return created_event

        except Exception as e:
            logger.error(f"Error creating event: {e}")
            return None

    def get_event_identifier(self, event: Dict[str, Any]) -> str:
        """
        Get a stable identifier for a Graph calendar event.
        Uses iCalUId (equivalent to GlobalAppointmentID) or falls back to Id.

        Args:
            event: Event dictionary from Graph API

        Returns:
            Event identifier string
        """
        # Try iCalUId first (stable across updates)
        if "iCalUId" in event and event["iCalUId"]:
            return event["iCalUId"]

        # Fall back to Graph Id
        if "id" in event and event["id"]:
            return event["id"]

        # If both fail, return empty string
        return ""


def filter_event(event: Dict[str, Any], ignore_patterns: List[str], timezone: str = "Europe/Copenhagen") -> Optional[str]:
    """
    Check if an event should be filtered out.

    Returns None if event should be included, otherwise returns reason for exclusion.

    Args:
        event: Event dictionary from Graph API
        ignore_patterns: List of regex patterns to ignore
        timezone: Timezone for date conversion

    Returns:
        None if included, or string reason if excluded
    """
    from .utils import should_ignore_appointment

    subject = event.get("subject", "")

    # Check if matches ignore pattern
    if should_ignore_appointment(subject, ignore_patterns):
        return "Matches ignore pattern"

    # Skip all-day events
    if event.get("isAllDay", False):
        return "All-day event"

    # Skip cancelled meetings
    if event.get("isCancelled", False):
        return "Meeting cancelled"

    # Skip declined meetings
    response_status = event.get("responseStatus", {})
    if response_status.get("response") == "declined":
        return "Meeting declined"

    # Skip private events
    if event.get("sensitivity") == "private":
        return "Private event"

    # Skip Out of Office events
    if event.get("showAs") == "oof":
        return "Out of Office"

    # Event should be included
    return None


def get_meeting_hours_summary(
    events: List[Dict[str, Any]],
    start_date: datetime,
    end_date: datetime,
    ignore_patterns: List[str],
    timezone: str = "Europe/Copenhagen",
    period_name: str = ""
) -> Dict[str, Any]:
    """
    Calculate total meeting hours for a time period.

    Args:
        events: List of event dictionaries from Graph API
        start_date: Start of time period
        end_date: End of time period
        ignore_patterns: List of regex patterns to ignore
        timezone: Timezone for conversions
        period_name: Name of period for logging

    Returns:
        Dictionary with 'hours' (float) and 'count' (int)
    """
    from .utils import get_appointment_duration

    if period_name:
        logger.info(f"Processing period: {period_name} ({start_date.date()} to {end_date.date()})")

    total_hours = 0.0
    appointment_count = 0
    excluded_count = 0

    for event in events:
        # Convert Graph DateTimeTimeZone to local DateTime
        event_start = from_graph_datetime(event.get("start", {}), timezone)
        event_end = from_graph_datetime(event.get("end", {}), timezone)

        # Check if event falls within the time period
        if event_start >= start_date and event_start < end_date:
            subject = event.get("subject", "Untitled")
            duration = get_appointment_duration(event_start, event_end)

            logger.info(f"  Found event: '{subject}' | Start: {event_start.strftime('%Y-%m-%d %H:%M')} | Duration: {duration} hours")

            # Check if event should be filtered
            exclusion_reason = filter_event(event, ignore_patterns, timezone)

            if exclusion_reason:
                logger.info(f"    -> EXCLUDED: {exclusion_reason}")
                excluded_count += 1
                continue

            # Include this event
            total_hours += duration
            appointment_count += 1
            logger.info(f"    -> INCLUDED in time estimate")

    summary_hours = round(total_hours, 2)
    logger.info(f"Summary for {period_name} - Total: {summary_hours} hours | Included: {appointment_count} events | Excluded: {excluded_count}")
    logger.info("")

    return {
        "hours": summary_hours,
        "count": appointment_count
    }
