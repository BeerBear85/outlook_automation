"""
Outlook Graph Automation Package

Python package for Microsoft Outlook automation using Microsoft Graph REST API.

Provides modules for:
- Authentication (auth.py)
- Configuration management (config.py)
- Calendar operations (calendar.py)
- Mail operations (mail.py)
- Utility functions (utils.py)

@author: Generated for outlook_automation repository (Python Graph implementation)
"""

from .auth import GraphAuthenticator, GraphClient, create_authenticator_from_config
from .config import Config, get_config, reload_config
from .calendar import CalendarClient, filter_event, get_meeting_hours_summary
from .mail import MailClient
from .utils import (
    get_weekday_bounds,
    get_next_working_day,
    get_appointment_duration,
    to_iso8601,
    from_graph_datetime,
    load_ignore_patterns,
    should_ignore_appointment,
    load_ignored_appointment_ids,
    save_ignored_appointment,
    load_email_template,
    render_template,
    setup_logging,
    initialize_log_file
)

__version__ = "1.0.0"
__author__ = "Generated for outlook_automation repository"

__all__ = [
    # Auth
    "GraphAuthenticator",
    "GraphClient",
    "create_authenticator_from_config",
    # Config
    "Config",
    "get_config",
    "reload_config",
    # Calendar
    "CalendarClient",
    "filter_event",
    "get_meeting_hours_summary",
    # Mail
    "MailClient",
    # Utils
    "get_weekday_bounds",
    "get_next_working_day",
    "get_appointment_duration",
    "to_iso8601",
    "from_graph_datetime",
    "load_ignore_patterns",
    "should_ignore_appointment",
    "load_ignored_appointment_ids",
    "save_ignored_appointment",
    "load_email_template",
    "render_template",
    "setup_logging",
    "initialize_log_file",
]
