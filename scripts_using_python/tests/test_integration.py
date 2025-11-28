"""
Integration tests for Outlook Graph Automation

Tests end-to-end workflows with mocked Graph API responses.
This ensures the complete functionality works correctly without requiring actual
authentication or live API calls.

@author: Generated for outlook_automation repository
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime, timedelta
from pathlib import Path
import sys
import json

# Add src to path
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))

from outlook_graph import (
    CalendarClient,
    MailClient,
    get_meeting_hours_summary,
    filter_event,
    from_graph_datetime,
    get_weekday_bounds,
    get_next_working_day,
)

try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo


# Sample Graph API event data for testing
def create_sample_event(subject, start_hour, duration_hours, is_cancelled=False, is_all_day=False):
    """Create a sample Graph API event object."""
    start_dt = datetime(2025, 1, 15, start_hour, 0, tzinfo=ZoneInfo("UTC"))
    end_dt = start_dt + timedelta(hours=duration_hours)

    return {
        "id": f"event-{subject.replace(' ', '-')}",
        "subject": subject,
        "start": {
            "dateTime": start_dt.strftime("%Y-%m-%dT%H:%M:%S.0000000"),
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": end_dt.strftime("%Y-%m-%dT%H:%M:%S.0000000"),
            "timeZone": "UTC"
        },
        "isCancelled": is_cancelled,
        "isAllDay": is_all_day,
        "sensitivity": "normal",
        "showAs": "busy",
        "responseStatus": {
            "response": "accepted"
        },
        "organizer": {
            "emailAddress": {
                "name": "Test Organizer",
                "address": "organizer@example.com"
            }
        },
        "iCalUId": f"ical-{subject.replace(' ', '-')}"
    }


class TestCalendarOperationsIntegration:
    """Integration tests for calendar operations."""

    @patch('requests.get')
    def test_list_events_with_pagination(self, mock_get):
        """Test listing events with pagination."""
        # Mock first page response
        page1_response = Mock()
        page1_response.status_code = 200
        page1_response.json.return_value = {
            "value": [
                create_sample_event("Meeting 1", 9, 1),
                create_sample_event("Meeting 2", 10, 1)
            ],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendarView?$skip=2"
        }

        # Mock second page response
        page2_response = Mock()
        page2_response.status_code = 200
        page2_response.json.return_value = {
            "value": [
                create_sample_event("Meeting 3", 11, 1)
            ]
        }

        # Configure mock to return different responses for different calls
        mock_get.side_effect = [page1_response, page2_response]

        # Create mock client
        graph_client = Mock()
        graph_client.get_headers.return_value = {"Authorization": "Bearer test-token"}
        graph_client.base_url = "https://graph.microsoft.com/v1.0"

        calendar_client = CalendarClient(graph_client)

        # Test
        start = datetime(2025, 1, 15, 0, 0, tzinfo=ZoneInfo("UTC"))
        end = datetime(2025, 1, 16, 0, 0, tzinfo=ZoneInfo("UTC"))

        events = calendar_client.list_events(start, end)

        # Verify
        assert len(events) == 3
        assert events[0]["subject"] == "Meeting 1"
        assert events[1]["subject"] == "Meeting 2"
        assert events[2]["subject"] == "Meeting 3"
        assert mock_get.call_count == 2  # Two API calls due to pagination

    @patch('requests.get')
    def test_list_events_error_handling(self, mock_get):
        """Test error handling when listing events fails."""
        # Mock error response
        error_response = Mock()
        error_response.status_code = 401
        error_response.text = "Unauthorized"
        error_response.raise_for_status.side_effect = Exception("401 Unauthorized")

        mock_get.return_value = error_response

        # Create mock client
        graph_client = Mock()
        graph_client.get_headers.return_value = {"Authorization": "Bearer invalid-token"}
        graph_client.base_url = "https://graph.microsoft.com/v1.0"

        calendar_client = CalendarClient(graph_client)

        # Test
        start = datetime(2025, 1, 15, 0, 0, tzinfo=ZoneInfo("UTC"))
        end = datetime(2025, 1, 16, 0, 0, tzinfo=ZoneInfo("UTC"))

        with pytest.raises(Exception, match="Failed to retrieve calendar events"):
            calendar_client.list_events(start, end)


class TestMailOperationsIntegration:
    """Integration tests for mail operations."""

    @patch('requests.post')
    def test_create_draft_success(self, mock_post):
        """Test creating draft email successfully."""
        # Mock successful response
        success_response = Mock()
        success_response.status_code = 201
        success_response.json.return_value = {
            "id": "draft-message-id",
            "subject": "Test Subject"
        }
        success_response.raise_for_status = Mock()

        mock_post.return_value = success_response

        # Create mock client
        graph_client = Mock()
        graph_client.get_headers.return_value = {"Authorization": "Bearer test-token"}
        graph_client.base_url = "https://graph.microsoft.com/v1.0"

        mail_client = MailClient(graph_client)

        # Test
        draft = mail_client.create_draft(
            to="recipient@example.com",
            subject="Test Subject",
            body="Test body content"
        )

        # Verify
        assert draft is not None
        assert draft["id"] == "draft-message-id"
        assert draft["subject"] == "Test Subject"
        mock_post.assert_called_once()

    @patch('requests.post')
    def test_create_draft_with_cc_bcc(self, mock_post):
        """Test creating draft with CC and BCC recipients."""
        success_response = Mock()
        success_response.status_code = 201
        success_response.json.return_value = {"id": "draft-id"}
        success_response.raise_for_status = Mock()

        mock_post.return_value = success_response

        graph_client = Mock()
        graph_client.get_headers.return_value = {"Authorization": "Bearer test-token"}
        graph_client.base_url = "https://graph.microsoft.com/v1.0"

        mail_client = MailClient(graph_client)

        # Test
        draft = mail_client.create_draft(
            to="to@example.com",
            subject="Test",
            body="Body",
            cc=["cc1@example.com", "cc2@example.com"],
            bcc=["bcc@example.com"]
        )

        # Verify
        assert draft is not None
        call_args = mock_post.call_args
        request_body = call_args[1]["json"]

        assert "ccRecipients" in request_body
        assert len(request_body["ccRecipients"]) == 2
        assert "bccRecipients" in request_body
        assert len(request_body["bccRecipients"]) == 1


class TestMeetingHoursSummary:
    """Integration tests for meeting hours calculation."""

    def test_calculate_meeting_hours_simple(self):
        """Test calculating meeting hours with simple event list."""
        events = [
            create_sample_event("Meeting 1", 9, 1),   # 9:00-10:00 (1 hour)
            create_sample_event("Meeting 2", 10, 1.5), # 10:00-11:30 (1.5 hours)
            create_sample_event("Meeting 3", 13, 2),   # 13:00-15:00 (2 hours)
        ]

        tz = ZoneInfo("Europe/Copenhagen")
        start_date = datetime(2025, 1, 15, 0, 0, tzinfo=tz)
        end_date = datetime(2025, 1, 16, 0, 0, tzinfo=tz)

        summary = get_meeting_hours_summary(
            events,
            start_date,
            end_date,
            ignore_patterns=[],
            timezone="Europe/Copenhagen"
        )

        # Should include all 3 meetings
        assert summary["count"] == 3
        # Total: 1 + 1.5 + 2 = 4.5 hours
        assert summary["hours"] == 4.5

    def test_calculate_meeting_hours_with_filtering(self):
        """Test calculating meeting hours with filtered events."""
        events = [
            create_sample_event("Team Meeting", 9, 1),
            create_sample_event("Lunch Break", 12, 1),  # Should be filtered
            create_sample_event("Project Review", 14, 1),
            create_sample_event("All Day Event", 0, 8, is_all_day=True),  # Filtered
            create_sample_event("Cancelled Meeting", 16, 1, is_cancelled=True),  # Filtered
        ]

        tz = ZoneInfo("Europe/Copenhagen")
        start_date = datetime(2025, 1, 15, 0, 0, tzinfo=tz)
        end_date = datetime(2025, 1, 16, 0, 0, tzinfo=tz)

        summary = get_meeting_hours_summary(
            events,
            start_date,
            end_date,
            ignore_patterns=[r"^Lunch.*"],
            timezone="Europe/Copenhagen"
        )

        # Should include only "Team Meeting" and "Project Review"
        assert summary["count"] == 2
        assert summary["hours"] == 2.0  # 1 + 1

    def test_filter_event_all_day(self):
        """Test that all-day events are filtered."""
        event = create_sample_event("All Day", 0, 24, is_all_day=True)

        reason = filter_event(event, [], "Europe/Copenhagen")

        assert reason == "All-day event"

    def test_filter_event_cancelled(self):
        """Test that cancelled events are filtered."""
        event = create_sample_event("Cancelled", 9, 1, is_cancelled=True)

        reason = filter_event(event, [], "Europe/Copenhagen")

        assert reason == "Meeting cancelled"

    def test_filter_event_ignored_pattern(self):
        """Test that events matching ignore patterns are filtered."""
        event = create_sample_event("Lunch with team", 12, 1)

        reason = filter_event(event, [r"^Lunch.*"], "Europe/Copenhagen")

        assert reason == "Matches ignore pattern"

    def test_filter_event_not_filtered(self):
        """Test that normal events are not filtered."""
        event = create_sample_event("Team Meeting", 9, 1)

        reason = filter_event(event, [], "Europe/Copenhagen")

        assert reason is None  # Not filtered


class TestEndToEndWorkflow:
    """End-to-end integration tests simulating full workflows."""

    def test_full_meeting_summary_workflow(self):
        """Test complete workflow: fetch events, calculate hours, filter."""
        # Create sample events for a week
        # Note: Using hours within 0-23 range, representing different days
        events = []

        # Monday (Jan 13)
        events.append(create_sample_event("Monday Standup", 9, 0.5))
        events.append(create_sample_event("Monday Project Review", 10, 1))

        # Tuesday (Jan 14) - same times as Monday for simplicity
        events.append(create_sample_event("Tuesday Team Meeting", 9, 1))

        # Wednesday (Jan 15) - all-day event (should be filtered)
        events.append(create_sample_event("Wednesday Conference", 0, 8, is_all_day=True))

        # Thursday (Jan 16)
        events.append(create_sample_event("Thursday Sprint Planning", 9, 2))

        # Friday (Jan 17)
        events.append(create_sample_event("Friday Retrospective", 9, 1.5))
        events.append(create_sample_event("Lunch Friday", 12, 1))  # Should be filtered

        tz = ZoneInfo("Europe/Copenhagen")
        week_start = datetime(2025, 1, 13, 0, 0, tzinfo=tz)  # Monday
        week_end = datetime(2025, 1, 18, 0, 0, tzinfo=tz)    # Saturday

        summary = get_meeting_hours_summary(
            events,
            week_start,
            week_end,
            ignore_patterns=[r"^Lunch.*"],
            timezone="Europe/Copenhagen"
        )

        # Should include: Mon (1.5h) + Tue (1h) + Thu (2h) + Fri (1.5h) = 6 hours
        # Excluded: Wed all-day, Fri lunch
        assert summary["count"] == 5
        assert summary["hours"] == 6.0

    def test_weekday_bounds_and_next_working_day_integration(self):
        """Test date/time utilities work together correctly."""
        # Friday
        friday = datetime(2025, 1, 17, 14, 0, tzinfo=ZoneInfo("Europe/Copenhagen"))

        # Get this week's bounds (Mon-Fri)
        monday, friday_end = get_weekday_bounds(friday)

        assert monday.weekday() == 0  # Monday
        assert monday.day == 13  # Jan 13
        assert friday_end.weekday() == 4  # Friday
        assert friday_end.day == 17  # Jan 17

        # Get next working day from Friday (should be Monday)
        next_day = get_next_working_day(friday)

        assert next_day.weekday() == 0  # Monday
        assert next_day.day == 20  # Jan 20 (next week)

    @patch('requests.post')
    def test_reschedule_email_workflow(self, mock_post):
        """Test complete reschedule workflow: detect full-hour meeting, create draft."""
        # Mock draft creation success
        success_response = Mock()
        success_response.status_code = 201
        success_response.json.return_value = {"id": "draft-id"}
        success_response.raise_for_status = Mock()
        mock_post.return_value = success_response

        # Create full-hour meeting (starts at 9:00)
        full_hour_meeting = create_sample_event("Full Hour Meeting", 9, 1)

        # Verify it starts at :00
        start_time = from_graph_datetime(
            full_hour_meeting["start"],
            "Europe/Copenhagen"
        )

        assert start_time.minute == 0
        assert start_time.second == 0

        # Create draft email
        graph_client = Mock()
        graph_client.get_headers.return_value = {"Authorization": "Bearer test-token"}
        graph_client.base_url = "https://graph.microsoft.com/v1.0"

        mail_client = MailClient(graph_client)

        organizer_email = full_hour_meeting["organizer"]["emailAddress"]["address"]
        subject = "Request to shift meeting start time to :05"
        body = f"Meeting: {full_hour_meeting['subject']}\nPlease consider moving to :05"

        draft = mail_client.create_draft(organizer_email, subject, body)

        # Verify
        assert draft is not None
        assert draft["id"] == "draft-id"
        mock_post.assert_called_once()


class TestRegressionTests:
    """Regression tests to ensure bugs don't reoccur."""

    def test_timezone_conversion_consistency(self):
        """
        Regression test: Ensure timezone conversions are consistent.
        Bug: Previous versions had inconsistent UTC/local conversions.
        """
        # Create event in UTC
        utc_event = {
            "dateTime": "2025-01-15T09:00:00.0000000",
            "timeZone": "UTC"
        }

        # Convert to Copenhagen time
        cph_time = from_graph_datetime(utc_event, "Europe/Copenhagen")

        # In January, Copenhagen is UTC+1 (CET)
        assert cph_time.hour == 10  # 9 UTC = 10 CET
        assert cph_time.tzinfo.key == "Europe/Copenhagen"

    def test_weekend_handling_saturday_to_monday(self):
        """
        Regression test: Ensure weekends are properly skipped.
        Bug: Previous versions incorrectly handled Saturday -> Monday transition.
        """
        # Saturday
        saturday = datetime(2025, 1, 18, 10, 0, tzinfo=ZoneInfo("Europe/Copenhagen"))

        # Next working day should be Monday
        next_day = get_next_working_day(saturday)

        assert next_day.weekday() == 0  # Monday
        assert next_day.day == 20  # Jan 20

    def test_weekend_handling_sunday_to_monday(self):
        """
        Regression test: Ensure Sunday -> Monday works correctly.
        """
        # Sunday
        sunday = datetime(2025, 1, 19, 10, 0, tzinfo=ZoneInfo("Europe/Copenhagen"))

        # Next working day should be Monday
        next_day = get_next_working_day(sunday)

        assert next_day.weekday() == 0  # Monday
        assert next_day.day == 20  # Jan 20

    def test_empty_event_list_doesnt_crash(self):
        """
        Regression test: Ensure empty event list is handled gracefully.
        Bug: Previous versions crashed with empty lists.
        """
        tz = ZoneInfo("Europe/Copenhagen")
        start = datetime(2025, 1, 15, 0, 0, tzinfo=tz)
        end = datetime(2025, 1, 16, 0, 0, tzinfo=tz)

        summary = get_meeting_hours_summary(
            [],  # Empty event list
            start,
            end,
            ignore_patterns=[],
            timezone="Europe/Copenhagen"
        )

        assert summary["hours"] == 0.0
        assert summary["count"] == 0

    def test_overlapping_meetings_counted_separately(self):
        """
        Regression test: Ensure overlapping meetings are counted separately.
        Bug: Previous versions might have merged overlapping meetings.
        """
        events = [
            create_sample_event("Meeting A", 9, 2),    # 9:00-11:00
            create_sample_event("Meeting B", 10, 1),   # 10:00-11:00 (overlaps)
        ]

        tz = ZoneInfo("Europe/Copenhagen")
        start = datetime(2025, 1, 15, 0, 0, tzinfo=tz)
        end = datetime(2025, 1, 16, 0, 0, tzinfo=tz)

        summary = get_meeting_hours_summary(
            events,
            start,
            end,
            ignore_patterns=[],
            timezone="Europe/Copenhagen"
        )

        # Should count both meetings
        assert summary["count"] == 2
        # Total hours: 2 + 1 = 3 (even though they overlap)
        assert summary["hours"] == 3.0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
