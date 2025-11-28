"""
Unit tests for utils module

Tests date/time utilities, filtering, template rendering, and logging functions.

@author: Generated for outlook_automation repository
"""

import pytest
from datetime import datetime, timedelta
from pathlib import Path
import tempfile
import sys

# Add src to path
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))

from outlook_graph.utils import (
    get_weekday_bounds,
    get_next_working_day,
    get_appointment_duration,
    to_iso8601,
    from_graph_datetime,
    should_ignore_appointment,
    load_ignore_patterns,
    render_template,
    load_email_template,
)

try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo


class TestDateTimeFunctions:
    """Test date/time utility functions."""

    def test_get_weekday_bounds_monday(self):
        """Test weekday bounds for a Monday."""
        # Monday, January 15, 2025
        monday = datetime(2025, 1, 13, 10, 30, tzinfo=ZoneInfo("Europe/Copenhagen"))
        start, end = get_weekday_bounds(monday)

        assert start.weekday() == 0  # Monday
        assert end.weekday() == 4  # Friday
        assert start.hour == 0
        assert start.minute == 0
        assert end.hour == 23
        assert end.minute == 59

    def test_get_weekday_bounds_friday(self):
        """Test weekday bounds for a Friday."""
        # Friday, January 17, 2025
        friday = datetime(2025, 1, 17, 14, 0, tzinfo=ZoneInfo("Europe/Copenhagen"))
        start, end = get_weekday_bounds(friday)

        assert start.weekday() == 0  # Monday (same week)
        assert end.weekday() == 4  # Friday
        assert (end - start).days == 4

    def test_get_weekday_bounds_saturday(self):
        """Test weekday bounds for a Saturday (should return next week)."""
        # Saturday, January 18, 2025
        saturday = datetime(2025, 1, 18, 10, 0, tzinfo=ZoneInfo("Europe/Copenhagen"))
        start, end = get_weekday_bounds(saturday)

        assert start.weekday() == 0  # Monday
        assert start.day == 20  # Next Monday

    def test_get_weekday_bounds_sunday(self):
        """Test weekday bounds for a Sunday (should return next week)."""
        # Sunday, January 19, 2025
        sunday = datetime(2025, 1, 19, 10, 0, tzinfo=ZoneInfo("Europe/Copenhagen"))
        start, end = get_weekday_bounds(sunday)

        assert start.weekday() == 0  # Monday
        assert start.day == 20  # Next Monday

    def test_get_next_working_day_monday(self):
        """Test next working day from Monday (should be Tuesday)."""
        monday = datetime(2025, 1, 13, 10, 0, tzinfo=ZoneInfo("Europe/Copenhagen"))
        next_day = get_next_working_day(monday)

        assert next_day.weekday() == 1  # Tuesday
        assert next_day.day == 14

    def test_get_next_working_day_friday(self):
        """Test next working day from Friday (should skip weekend to Monday)."""
        friday = datetime(2025, 1, 17, 10, 0, tzinfo=ZoneInfo("Europe/Copenhagen"))
        next_day = get_next_working_day(friday)

        assert next_day.weekday() == 0  # Monday
        assert next_day.day == 20

    def test_get_next_working_day_saturday(self):
        """Test next working day from Saturday (should be Monday)."""
        saturday = datetime(2025, 1, 18, 10, 0, tzinfo=ZoneInfo("Europe/Copenhagen"))
        next_day = get_next_working_day(saturday)

        assert next_day.weekday() == 0  # Monday
        assert next_day.day == 20

    def test_get_appointment_duration(self):
        """Test appointment duration calculation."""
        start = datetime(2025, 1, 15, 9, 0)
        end = datetime(2025, 1, 15, 10, 30)
        duration = get_appointment_duration(start, end)

        assert duration == 1.5  # 1.5 hours

    def test_get_appointment_duration_full_day(self):
        """Test appointment duration for a full day."""
        start = datetime(2025, 1, 15, 9, 0)
        end = datetime(2025, 1, 15, 17, 0)
        duration = get_appointment_duration(start, end)

        assert duration == 8.0  # 8 hours

    def test_to_iso8601(self):
        """Test conversion to ISO8601 format."""
        dt = datetime(2025, 1, 15, 9, 30, 45, tzinfo=ZoneInfo("Europe/Copenhagen"))
        iso_str = to_iso8601(dt)

        assert iso_str == "2025-01-15T09:30:45"
        assert "T" in iso_str
        assert "Z" not in iso_str  # Should not include timezone

    def test_from_graph_datetime_utc(self):
        """Test conversion from Graph DateTimeTimeZone object (UTC)."""
        graph_obj = {
            "dateTime": "2025-01-15T09:00:00.0000000",
            "timeZone": "UTC"
        }

        dt = from_graph_datetime(graph_obj, "Europe/Copenhagen")

        assert isinstance(dt, datetime)
        assert dt.tzinfo is not None
        assert dt.hour == 10  # UTC+1 in January (CET)

    def test_from_graph_datetime_local(self):
        """Test conversion from Graph DateTimeTimeZone object (local timezone)."""
        graph_obj = {
            "dateTime": "2025-01-15T09:00:00",
            "timeZone": "Europe/Copenhagen"
        }

        dt = from_graph_datetime(graph_obj, "Europe/Copenhagen")

        assert dt.hour == 9
        assert dt.minute == 0


class TestFilteringFunctions:
    """Test filtering utility functions."""

    def test_should_ignore_appointment_match(self):
        """Test that appointments matching patterns are ignored."""
        patterns = [r"^Lunch.*", r"^Personal.*"]

        assert should_ignore_appointment("Lunch with team", patterns) is True
        assert should_ignore_appointment("Personal time", patterns) is True

    def test_should_ignore_appointment_no_match(self):
        """Test that appointments not matching patterns are not ignored."""
        patterns = [r"^Lunch.*", r"^Personal.*"]

        assert should_ignore_appointment("Team meeting", patterns) is False
        assert should_ignore_appointment("Project review", patterns) is False

    def test_should_ignore_appointment_case_insensitive(self):
        """Test that pattern matching is case-insensitive by default."""
        patterns = [r"^lunch.*"]

        assert should_ignore_appointment("Lunch break", patterns) is True
        assert should_ignore_appointment("LUNCH BREAK", patterns) is True

    def test_should_ignore_appointment_complex_pattern(self):
        """Test complex regex patterns."""
        patterns = [r".*\[IGNORE\].*", r"^(Lunch|Break|Personal).*"]

        assert should_ignore_appointment("Meeting [IGNORE] internal", patterns) is True
        assert should_ignore_appointment("Break time", patterns) is True
        assert should_ignore_appointment("Regular meeting", patterns) is False

    def test_load_ignore_patterns(self):
        """Test loading ignore patterns from file."""
        # Create temporary config directory with patterns file
        with tempfile.TemporaryDirectory() as tmpdir:
            config_dir = Path(tmpdir)
            patterns_file = config_dir / "ignore_appointments.txt"

            # Write test patterns
            patterns_file.write_text("""# Test patterns
^Lunch.*
^Personal.*
# Another comment
^Break.*
""")

            patterns = load_ignore_patterns(config_dir)

            assert len(patterns) == 3
            assert "^Lunch.*" in patterns
            assert "^Personal.*" in patterns
            assert "^Break.*" in patterns
            assert "# Test patterns" not in patterns  # Comments excluded

    def test_load_ignore_patterns_empty_file(self):
        """Test loading from empty patterns file."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_dir = Path(tmpdir)
            patterns_file = config_dir / "ignore_appointments.txt"
            patterns_file.write_text("")

            patterns = load_ignore_patterns(config_dir)

            assert len(patterns) == 0

    def test_load_ignore_patterns_missing_file(self):
        """Test loading when patterns file doesn't exist."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_dir = Path(tmpdir)
            patterns = load_ignore_patterns(config_dir)

            assert len(patterns) == 0


class TestTemplateRendering:
    """Test template rendering functions."""

    def test_render_template_basic(self):
        """Test basic template rendering."""
        template = "Hello {NAME}, your meeting is at {TIME}"
        result = render_template(template, name="John", time="9:00 AM")

        assert result == "Hello John, your meeting is at 9:00 AM"

    def test_render_template_multiple_placeholders(self):
        """Test template with multiple placeholders."""
        template = "Meeting: {SUBJECT}\nOrganizer: {ORGANIZER}\nTime: {START_TIME}"
        result = render_template(
            template,
            subject="Team Sync",
            organizer="Alice",
            start_time="2025-01-15 09:00"
        )

        assert "Team Sync" in result
        assert "Alice" in result
        assert "2025-01-15 09:00" in result

    def test_render_template_unused_placeholders(self):
        """Test that unused placeholders remain unchanged."""
        template = "Hello {NAME}, your code is {CODE}"
        result = render_template(template, name="John")

        assert "John" in result
        assert "{CODE}" in result  # Unused placeholder remains

    def test_load_email_template(self):
        """Test loading email template from file."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_dir = Path(tmpdir)
            template_file = config_dir / "meeting_change_request_template.txt"

            template_content = """Subject: Test Request

Dear {ORGANIZER},

Meeting: {SUBJECT}
Time: {START_TIME}
"""
            template_file.write_text(template_content)

            loaded_template = load_email_template(config_dir)

            assert "Dear {ORGANIZER}" in loaded_template
            assert "{SUBJECT}" in loaded_template
            assert "{START_TIME}" in loaded_template

    def test_load_email_template_default(self):
        """Test loading default template when file doesn't exist."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_dir = Path(tmpdir)
            template = load_email_template(config_dir)

            # Should return default template
            assert "Dear {ORGANIZER}" in template
            assert "{SUBJECT}" in template
            assert "{START_TIME}" in template
            assert "{NEW_START_TIME}" in template


class TestEdgeCases:
    """Test edge cases and error handling."""

    def test_get_appointment_duration_zero(self):
        """Test zero-duration appointment."""
        start = datetime(2025, 1, 15, 9, 0)
        end = datetime(2025, 1, 15, 9, 0)
        duration = get_appointment_duration(start, end)

        assert duration == 0.0

    def test_get_appointment_duration_negative(self):
        """Test negative duration (end before start)."""
        start = datetime(2025, 1, 15, 10, 0)
        end = datetime(2025, 1, 15, 9, 0)
        duration = get_appointment_duration(start, end)

        assert duration == -1.0

    def test_should_ignore_appointment_empty_patterns(self):
        """Test filtering with empty pattern list."""
        patterns = []
        assert should_ignore_appointment("Any meeting", patterns) is False

    def test_should_ignore_appointment_invalid_regex(self):
        """Test that invalid regex patterns are handled gracefully."""
        patterns = [r"[invalid(regex"]  # Invalid pattern

        # Should not crash, should skip invalid pattern
        result = should_ignore_appointment("Test meeting", patterns)
        assert result is False

    def test_to_iso8601_naive_datetime(self):
        """Test ISO8601 conversion with naive datetime."""
        dt = datetime(2025, 1, 15, 9, 30)  # Naive (no timezone)
        iso_str = to_iso8601(dt)

        assert iso_str == "2025-01-15T09:30:00"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
