"""
Utility Functions Module

Provides shared helper functions for:
- Date/time calculations and conversions
- Appointment filtering
- Configuration file loading
- Logging
- Template rendering

@author: Generated for outlook_automation repository (Python Graph implementation)
"""

import re
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Tuple, Dict, Any, Optional
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except ImportError:
    from backports.zoneinfo import ZoneInfo  # Fallback for older Python


# -----------------------------------------------------------------------------
# Date/Time Helper Functions
# -----------------------------------------------------------------------------

def get_weekday_bounds(reference_date: datetime, timezone: str = "Europe/Copenhagen") -> Tuple[datetime, datetime]:
    """
    Get Monday and Friday bounds for a given week.
    For weekends, returns the upcoming work week (Monday-Friday).

    Args:
        reference_date: Reference date (timezone-aware or naive)
        timezone: Timezone string (IANA format)

    Returns:
        Tuple of (monday, friday) as timezone-aware datetimes
    """
    # Ensure timezone-aware
    if reference_date.tzinfo is None:
        tz = ZoneInfo(timezone)
        reference_date = reference_date.replace(tzinfo=tz)

    # Get day of week (0 = Monday, 6 = Sunday)
    day_of_week = reference_date.weekday()

    # For weekends, use the upcoming Monday
    # For weekdays, use the Monday of the current week
    if day_of_week == 6:  # Sunday
        days_to_monday = 1
    elif day_of_week == 5:  # Saturday
        days_to_monday = 2
    else:  # Weekday
        days_to_monday = -day_of_week

    monday = (reference_date + timedelta(days=days_to_monday)).replace(hour=0, minute=0, second=0, microsecond=0)
    friday = (monday + timedelta(days=4)).replace(hour=23, minute=59, second=59, microsecond=999999)

    return monday, friday


def get_next_working_day(reference_date: datetime, timezone: str = "Europe/Copenhagen") -> datetime:
    """
    Get the next working day (Monday-Friday) from a given reference date.
    Skips weekends.

    Args:
        reference_date: Reference date
        timezone: Timezone string

    Returns:
        Next working day as timezone-aware datetime
    """
    # Ensure timezone-aware
    if reference_date.tzinfo is None:
        tz = ZoneInfo(timezone)
        reference_date = reference_date.replace(tzinfo=tz)

    next_day = (reference_date + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)

    # If next day is Saturday (5), move to Monday (+2 days)
    if next_day.weekday() == 5:
        return next_day + timedelta(days=2)
    # If next day is Sunday (6), move to Monday (+1 day)
    elif next_day.weekday() == 6:
        return next_day + timedelta(days=1)
    # Otherwise, next day is already a working day
    else:
        return next_day


def get_appointment_duration(start: datetime, end: datetime) -> float:
    """
    Calculate the duration of an appointment in hours.

    Args:
        start: Start datetime
        end: End datetime

    Returns:
        Duration in hours (rounded to 2 decimal places)
    """
    duration = end - start
    hours = duration.total_seconds() / 3600
    return round(hours, 2)


def to_iso8601(dt: datetime) -> str:
    """
    Convert datetime to ISO8601 string for Microsoft Graph API.

    Args:
        dt: Datetime object (timezone-aware or naive)

    Returns:
        ISO8601 formatted string (e.g., "2025-01-15T09:00:00")
    """
    # Remove timezone info for Graph API (Graph expects local time in ISO format)
    if dt.tzinfo is not None:
        dt = dt.replace(tzinfo=None)
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def from_graph_datetime(dt_tz_obj: Dict[str, str], timezone: str = "Europe/Copenhagen") -> datetime:
    """
    Convert Microsoft Graph DateTimeTimeZone object to Python datetime.

    Graph returns: {"dateTime": "2025-01-15T09:00:00.0000000", "timeZone": "UTC"}

    Args:
        dt_tz_obj: Dictionary with "dateTime" and "timeZone" keys
        timezone: Target timezone for conversion

    Returns:
        Timezone-aware datetime object
    """
    dt_string = dt_tz_obj.get("dateTime", "")
    tz_string = dt_tz_obj.get("timeZone", "UTC")

    # Parse datetime string (handle fractional seconds)
    # Graph format: "2025-01-15T09:00:00.0000000"
    dt_string = dt_string.split('.')[0]  # Remove fractional seconds
    dt = datetime.fromisoformat(dt_string)

    # Apply source timezone
    if tz_string == "UTC":
        dt = dt.replace(tzinfo=ZoneInfo("UTC"))
    else:
        try:
            dt = dt.replace(tzinfo=ZoneInfo(tz_string))
        except:
            # Fallback to UTC if timezone is invalid
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))

    # Convert to target timezone
    target_tz = ZoneInfo(timezone)
    dt = dt.astimezone(target_tz)

    return dt


# -----------------------------------------------------------------------------
# Filtering Functions
# -----------------------------------------------------------------------------

def load_ignore_patterns(config_dir: Path) -> List[str]:
    """
    Load regex patterns from ignore_appointments.txt file.

    Args:
        config_dir: Path to config directory

    Returns:
        List of regex pattern strings
    """
    ignore_file = config_dir / "ignore_appointments.txt"
    patterns = []

    if ignore_file.exists():
        try:
            with open(ignore_file, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    # Skip empty lines and comments
                    if line and not line.startswith('#'):
                        patterns.append(line)
        except Exception as e:
            logging.warning(f"Failed to load ignore patterns: {e}")

    return patterns


def should_ignore_appointment(subject: str, patterns: List[str]) -> bool:
    """
    Test if an appointment should be ignored based on regex patterns.

    Args:
        subject: Appointment subject
        patterns: List of regex patterns

    Returns:
        True if appointment should be ignored, False otherwise
    """
    for pattern in patterns:
        try:
            # Use IGNORECASE by default unless pattern includes (?-i)
            if re.search(pattern, subject, re.IGNORECASE):
                return True
        except re.error:
            logging.warning(f"Invalid regex pattern: {pattern}")
            continue

    return False


def load_ignored_appointment_ids(config_dir: Path) -> List[str]:
    """
    Load appointment identifiers from ignored_full_hour_appointments.txt file.

    Args:
        config_dir: Path to config directory

    Returns:
        List of appointment identifier strings
    """
    ignore_file = config_dir / "ignored_full_hour_appointments.txt"
    identifiers = []

    if ignore_file.exists():
        try:
            with open(ignore_file, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    # Skip empty lines and comments
                    if line and not line.startswith('#'):
                        identifiers.append(line)
        except Exception as e:
            logging.warning(f"Failed to load ignored appointment IDs: {e}")

    return identifiers


def save_ignored_appointment(identifier: str, subject: str, start_time: datetime, config_dir: Path):
    """
    Add an appointment identifier to the ignored_full_hour_appointments.txt file.

    Args:
        identifier: Appointment identifier (iCalUId or id)
        subject: Appointment subject
        start_time: Appointment start time
        config_dir: Path to config directory
    """
    ignore_file = config_dir / "ignored_full_hour_appointments.txt"

    # Create file with header if it doesn't exist
    if not ignore_file.exists():
        header = """# Ignored Full-Hour Appointments
# This file contains appointment identifiers that should not trigger the reschedule popup.
# Each line contains an iCalUId or event Id from Microsoft Graph.
# Lines starting with # are comments and will be ignored.
# You can manually edit this file to add or remove entries.
#
"""
        ignore_file.write_text(header, encoding='utf-8')

    # Add comment with meeting details and identifier
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    start_str = start_time.strftime("%Y-%m-%d %H:%M")
    comment = f"# Added: {timestamp} | Subject: {subject} | Start: {start_str}\n"

    with open(ignore_file, 'a', encoding='utf-8') as f:
        f.write(comment)
        f.write(identifier + '\n')

    logging.info(f"Added to ignore list: {subject} (ID: {identifier})")


# -----------------------------------------------------------------------------
# Template Functions
# -----------------------------------------------------------------------------

def load_email_template(config_dir: Path) -> str:
    """
    Load email template from meeting_change_request_template.txt file.

    Args:
        config_dir: Path to config directory

    Returns:
        Template string with placeholders
    """
    template_file = config_dir / "meeting_change_request_template.txt"

    if template_file.exists():
        try:
            return template_file.read_text(encoding='utf-8')
        except Exception as e:
            logging.warning(f"Failed to load email template: {e}")

    # Return default template if file doesn't exist
    return """Subject: Request to shift meeting start time to :05

Dear {ORGANIZER},

I hope this message finds you well. I'm reaching out regarding our upcoming meeting:

Meeting: {SUBJECT}
Current Start Time: {START_TIME}

Would it be possible to shift the meeting start time by 5 minutes to {NEW_START_TIME}? This small adjustment would help create a buffer between back-to-back meetings and allow for better preparation time.

If this change works for you and other attendees, I would greatly appreciate it. If the current time is critical, please feel free to keep it as scheduled.

Thank you for considering this request.

Best regards
"""


def render_template(template: str, **kwargs) -> str:
    """
    Replace placeholders in template with provided values.

    Args:
        template: Template string with {PLACEHOLDER} markers
        **kwargs: Placeholder values

    Returns:
        Rendered template string
    """
    result = template
    for key, value in kwargs.items():
        placeholder = f"{{{key.upper()}}}"
        result = result.replace(placeholder, str(value))
    return result


# -----------------------------------------------------------------------------
# Logging Functions
# -----------------------------------------------------------------------------

def setup_logging(log_file: Path, level: int = logging.INFO) -> logging.Logger:
    """
    Configure Python logging to file and console.

    Args:
        log_file: Path to log file
        level: Logging level (default: INFO)

    Returns:
        Configured logger instance
    """
    # Create logger
    logger = logging.getLogger("outlook_automation")
    logger.setLevel(level)

    # Remove existing handlers
    logger.handlers = []

    # File handler
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(level)
    file_formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s', datefmt='%H:%M:%S')
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)

    # Console handler (less verbose)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)  # Only warnings and errors to console
    console_formatter = logging.Formatter('[%(levelname)s] %(message)s')
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    return logger


def initialize_log_file(log_dir: Path, script_name: str = "Meeting Hour Summary") -> Path:
    """
    Create or clear the log file and write a header.

    Args:
        log_dir: Directory for log file
        script_name: Name of the script for header

    Returns:
        Path to log file
    """
    log_file = log_dir / "log.txt"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    header = f"""================================================================================
{script_name} Script Log (Microsoft Graph Python Version)
Generated: {timestamp}
================================================================================

"""

    log_file.write_text(header, encoding='utf-8')
    return log_file
