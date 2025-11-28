#!/usr/bin/env python3
"""
Meeting Hour Summary Script (Microsoft Graph Python Version)

Displays daily and weekly meeting hour summaries via tkinter GUI popup.
Uses Microsoft Graph REST API instead of COM or PowerShell.

Equivalent to: Show-MeetingHourSummary.ps1 (PowerShell version)

@author: Generated for outlook_automation repository (Python Graph implementation)
"""

import sys
from pathlib import Path
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox, Canvas, Button, Label, Frame

# Add src directory to path
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))

from outlook_graph import (
    get_config,
    create_authenticator_from_config,
    GraphClient,
    CalendarClient,
    MailClient,
    get_weekday_bounds,
    get_next_working_day,
    get_meeting_hours_summary,
    from_graph_datetime,
    load_ignore_patterns,
    load_ignored_appointment_ids,
    save_ignored_appointment,
    load_email_template,
    render_template,
    initialize_log_file,
    setup_logging,
    filter_event
)
import logging


logger = None


def get_daily_meeting_hours(events, start_date, working_day_count, ignore_patterns, timezone):
    """
    Calculate total meeting hours for each of the next N working days (Monday-Friday).
    Skips weekends.

    Returns:
        List of dictionaries with date, hours, count, and day_of_week
    """
    daily_hours = []
    current_date = start_date
    working_days_found = 0

    while working_days_found < working_day_count:
        # Skip weekends (Saturday = 5, Sunday = 6)
        if current_date.weekday() not in [5, 6]:
            day_start = current_date.replace(hour=0, minute=0, second=0, microsecond=0)
            day_end = current_date.replace(hour=23, minute=59, second=59, microsecond=999999)

            result = get_meeting_hours_summary(
                events,
                day_start,
                day_end,
                ignore_patterns,
                timezone,
                f"Working Day {working_days_found + 1} - {day_start.strftime('%Y-%m-%d')}"
            )

            daily_hours.append({
                'date': day_start,
                'hours': result['hours'],
                'count': result['count'],
                'day_of_week': day_start.strftime("%a")
            })

            working_days_found += 1

        current_date += timedelta(days=1)

    return daily_hours


def get_full_hour_meetings(events, start_date, end_date, ignore_patterns, ignored_ids, timezone, max_count=10):
    """
    Find meetings starting exactly on the full hour in the date range.
    Excludes all-day events, private items, OOO entries, and previously ignored meetings.

    Returns at most max_count meetings, starting with the earliest.
    """
    logger.info(f"Scanning for full-hour meetings (starting at :00) in range")
    if ignored_ids:
        logger.info(f"  Loaded {len(ignored_ids)} previously ignored appointment(s)")

    full_hour_meetings = []
    now = datetime.now(start_date.tzinfo)
    skipped_reasons = {
        'Ignored': 0,
        'AllDay': 0,
        'Private': 0,
        'OutOfOffice': 0,
        'AlreadyStarted': 0,
        'NotFullHour': 0,
        'Cancelled': 0,
        'Declined': 0,
        'PreviouslyIgnored': 0
    }

    for event in events:
        event_start = from_graph_datetime(event.get("start", {}), timezone)

        # Check if event is in the date range
        if event_start >= start_date and event_start < end_date:
            # Check exclusion reason
            exclusion_reason = filter_event(event, ignore_patterns, timezone)
            if exclusion_reason:
                if exclusion_reason == "Matches ignore pattern":
                    skipped_reasons['Ignored'] += 1
                elif exclusion_reason == "All-day event":
                    skipped_reasons['AllDay'] += 1
                elif exclusion_reason == "Private event":
                    skipped_reasons['Private'] += 1
                elif exclusion_reason == "Out of Office":
                    skipped_reasons['OutOfOffice'] += 1
                elif exclusion_reason == "Meeting cancelled":
                    skipped_reasons['Cancelled'] += 1
                elif exclusion_reason == "Meeting declined":
                    skipped_reasons['Declined'] += 1
                continue

            # Skip meetings that have already started
            if event_start < now:
                skipped_reasons['AlreadyStarted'] += 1
                continue

            # Check if start time is exactly on the hour (minute = 0, second = 0)
            if event_start.minute == 0 and event_start.second == 0:
                # Check if this event was previously ignored
                event_id = event.get("iCalUId") or event.get("id", "")
                if event_id and event_id in ignored_ids:
                    skipped_reasons['PreviouslyIgnored'] += 1
                    logger.info(f"  Skipped previously ignored meeting: '{event.get('subject', 'Untitled')}' | Start: {event_start.strftime('%Y-%m-%d %H:%M')}")
                    continue

                full_hour_meetings.append(event)
                organizer_name = event.get('organizer', {}).get('emailAddress', {}).get('name', 'Unknown')
                logger.info(f"  Found full-hour meeting: '{event.get('subject', 'Untitled')}' | Start: {event_start.strftime('%Y-%m-%d %H:%M')} | Organizer: {organizer_name}")
            else:
                skipped_reasons['NotFullHour'] += 1

    # Sort by start time and take the first max_count
    full_hour_meetings.sort(key=lambda e: from_graph_datetime(e.get("start", {}), timezone))
    full_hour_meetings = full_hour_meetings[:max_count]

    logger.info(f"Full-hour meeting scan complete: Found {len(full_hour_meetings)} meetings")
    logger.info(f"  Skipped: {skipped_reasons['Ignored']} (ignored pattern), {skipped_reasons['AllDay']} (all-day), {skipped_reasons['Private']} (private), {skipped_reasons['OutOfOffice']} (OOO), {skipped_reasons['AlreadyStarted']} (already started), {skipped_reasons['Cancelled']} (cancelled), {skipped_reasons['Declined']} (declined), {skipped_reasons['PreviouslyIgnored']} (previously ignored)")
    logger.info("")

    return full_hour_meetings


def show_reschedule_dialog(subject, start_time, organizer):
    """
    Show a popup dialog asking if user wants to draft a reschedule email.

    Returns: "CreateDraft", "Skip", or "NeverAskAgain"
    """
    dialog = tk.Toplevel()
    dialog.title("Reschedule Meeting?")
    dialog.geometry("550x320")
    dialog.resizable(False, False)
    dialog.attributes('-topmost', True)

    result = {"choice": "Skip"}

    # Title
    title_label = Label(dialog, text="Meeting starts at full hour", font=("Segoe UI", 12, "bold"))
    title_label.place(x=20, y=20, width=510, height=25)

    # Message
    message_text = "The following meeting starts exactly on the hour. Would you like to draft an email requesting it be moved to :05?"
    message_label = Label(dialog, text=message_text, font=("Segoe UI", 9), wraplength=510, justify="left")
    message_label.place(x=20, y=55, width=510, height=60)

    # Details
    details_text = f"Subject: {subject}\nStart Time: {start_time.strftime('%A, %B %d, %Y %H:%M')}\nOrganizer: {organizer}"
    details_label = Label(dialog, text=details_text, font=("Segoe UI", 9), fg="darkblue", justify="left")
    details_label.place(x=20, y=120, width=510, height=80)

    def on_create_draft():
        result["choice"] = "CreateDraft"
        dialog.destroy()

    def on_skip():
        result["choice"] = "Skip"
        dialog.destroy()

    def on_never_ask():
        result["choice"] = "NeverAskAgain"
        dialog.destroy()

    # Buttons
    create_button = Button(dialog, text="Create Draft Email", font=("Segoe UI", 9), command=on_create_draft)
    create_button.place(x=30, y=220, width=150, height=35)

    skip_button = Button(dialog, text="Skip for Now", font=("Segoe UI", 9), command=on_skip)
    skip_button.place(x=200, y=220, width=150, height=35)

    never_button = Button(dialog, text="Never Ask Again", font=("Segoe UI", 9), fg="darkred", command=on_never_ask)
    never_button.place(x=370, y=220, width=150, height=35)

    # Center on screen
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
    y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
    dialog.geometry(f"+{x}+{y}")

    dialog.wait_window()
    return result["choice"]


def create_reschedule_draft(event, template, mail_client, timezone, config_dir):
    """
    Create a draft email requesting meeting reschedule.

    Returns: True if successful, False otherwise
    """
    try:
        # Get organizer email
        organizer_email = event.get('organizer', {}).get('emailAddress', {}).get('address', '')
        organizer_name = event.get('organizer', {}).get('emailAddress', {}).get('name', 'Unknown')

        if not organizer_email:
            raise Exception("Cannot determine organizer email address")

        # Get event start time
        current_start = from_graph_datetime(event.get("start", {}), timezone)

        # Calculate new start time (add 5 minutes)
        new_start = current_start + timedelta(minutes=5)

        # Replace placeholders in template
        email_content = render_template(
            template,
            ORGANIZER=organizer_name,
            SUBJECT=event.get('subject', 'Untitled'),
            START_TIME=current_start.strftime('%A, %B %d, %Y %H:%M'),
            NEW_START_TIME=new_start.strftime('%H:%M')
        )

        # Extract subject line from template (first line after "Subject:")
        subject_line = "Request to shift meeting start time to :05"
        if "Subject:" in email_content:
            lines = email_content.split('\n')
            for i, line in enumerate(lines):
                if line.startswith("Subject:"):
                    subject_line = line.replace("Subject:", "").strip()
                    # Remove subject line from body
                    email_content = '\n'.join(lines[i+1:]).strip()
                    # Remove leading blank line
                    if email_content.startswith('\n'):
                        email_content = email_content[1:]
                    break

        # Create draft email
        draft = mail_client.create_draft(organizer_email, subject_line, email_content)

        return draft is not None

    except Exception as e:
        logger.error(f"Failed to create draft email: {e}")
        return False


def show_meeting_summary_gui(today_hours, next_day_hours, this_week_hours, next_week_hours, daily_hours, today, next_working_day, this_week_bounds, next_week_bounds, ignore_pattern_count):
    """
    Show Windows-style popup with meeting hour summaries and 5-day bar chart.
    """
    root = tk.Tk()
    root.title("Meeting Hour Summary")
    root.geometry("450x650")
    root.resizable(False, False)

    # Title
    title_label = Label(root, text="Meeting Hour Summary", font=("Segoe UI", 16, "bold"))
    title_label.place(x=20, y=20, width=410, height=30)

    # Subtitle
    now = datetime.now()
    subtitle_text = f"Generated on: {now.strftime('%A, %B %d, %Y %H:%M')} | Microsoft Graph (Python)"
    subtitle_label = Label(root, text=subtitle_text, font=("Segoe UI", 9, "italic"), fg="gray")
    subtitle_label.place(x=20, y=55, width=410, height=20)

    y_pos = 90

    def add_summary_row(y, label_text, hours, count, date_range):
        # Period label
        period_label = Label(root, text=label_text, font=("Segoe UI", 10, "bold"), anchor="w")
        period_label.place(x=30, y=y, width=150, height=20)

        # Hours value
        hours_label = Label(root, text=f"{hours} hours", font=("Segoe UI", 10), anchor="w")
        hours_label.place(x=190, y=y, width=100, height=20)

        # Meeting count
        count_label = Label(root, text=f"({count} meetings)", font=("Segoe UI", 9), fg="gray", anchor="w")
        count_label.place(x=300, y=y, width=120, height=20)

        # Date range
        date_label = Label(root, text=date_range, font=("Segoe UI", 8), fg="darkgray", anchor="w")
        date_label.place(x=30, y=y+22, width=390, height=16)

        return y + 50

    # Add summary rows
    y_pos = add_summary_row(y_pos, "Today:", today_hours['hours'], today_hours['count'], today.strftime("%A, %B %d"))
    y_pos = add_summary_row(y_pos, "Next Working Day:", next_day_hours['hours'], next_day_hours['count'], next_working_day.strftime("%A, %B %d"))
    y_pos = add_summary_row(y_pos, "This Week:", this_week_hours['hours'], this_week_hours['count'], f"{this_week_bounds[0].strftime('%b %d')} - {this_week_bounds[1].date().strftime('%b %d')}")
    y_pos = add_summary_row(y_pos, "Next Week:", next_week_hours['hours'], next_week_hours['count'], f"{next_week_bounds[0].strftime('%b %d')} - {next_week_bounds[1].date().strftime('%b %d')}")

    y_pos += 10

    # Chart title
    chart_title = Label(root, text="Next 5 Working Days Overview", font=("Segoe UI", 10, "bold"))
    chart_title.place(x=20, y=y_pos, width=410, height=20)

    y_pos += 25

    # Create canvas for bar chart
    canvas = Canvas(root, width=410, height=180, bg="white", relief="solid", borderwidth=1)
    canvas.place(x=20, y=y_pos)

    # Draw bar chart
    if daily_hours:
        chart_width = 390
        chart_height = 120
        bar_count = len(daily_hours)
        bar_width = chart_width // (bar_count * 1.5)
        bar_spacing = bar_width // 2
        max_hours = 8
        start_x = 10
        start_y = 10

        for i, day_data in enumerate(daily_hours):
            hours = day_data['hours']

            # Determine bar color
            if hours <= 3:
                bar_color = "#228B22"  # Green
            elif hours <= 4:
                bar_color = "#FFC107"  # Yellow
            else:
                bar_color = "#DC3545"  # Red

            # Calculate bar height
            bar_height = min((hours / max_hours) * chart_height, chart_height)

            # Calculate bar position
            bar_x = start_x + (i * (bar_width + bar_spacing))
            bar_y = start_y + chart_height - bar_height

            # Draw bar
            canvas.create_rectangle(bar_x, bar_y, bar_x + bar_width, bar_y + bar_height, fill=bar_color, outline="black")

            # Draw hours value on top
            hours_text = f"{hours}h"
            text_y = max(bar_y - 12, 2)
            canvas.create_text(bar_x + bar_width/2, text_y, text=hours_text, font=("Segoe UI", 8, "bold"))

            # Draw day label below
            day_label = f"{day_data['day_of_week']}\n{day_data['date'].strftime('%m/%d')}"
            canvas.create_text(bar_x + bar_width/2, start_y + chart_height + 20, text=day_label, font=("Segoe UI", 7))

    y_pos += 185

    # Footer note
    footer_text = f"Note: {ignore_pattern_count} ignore pattern(s) applied | Powered by Microsoft Graph (Python)"
    footer_label = Label(root, text=footer_text, font=("Segoe UI", 8), fg="darkgray")
    footer_label.place(x=20, y=y_pos + 15, width=410, height=30)

    # OK button
    ok_button = Button(root, text="OK", font=("Segoe UI", 10), command=root.destroy)
    ok_button.place(x=170, y=y_pos + 50, width=100, height=30)

    # Center on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")

    root.mainloop()


def main():
    """
    Main function for meeting hour summary script.
    """
    global logger

    print("Meeting Hour Summary Script (Microsoft Graph Python Version)")

    # Initialize
    script_dir = Path(__file__).parent.parent
    config_dir = script_dir / "config"
    config_dir.mkdir(exist_ok=True)

    # Initialize logging
    log_file = initialize_log_file(script_dir, "Meeting Hour Summary")
    logger = setup_logging(log_file)

    print(f"Logging to: {log_file}")
    print()

    logger.info("Script execution started")
    logger.info(f"Script directory: {script_dir}")
    logger.info("Using Microsoft Graph REST API (Python)")

    # Load configuration
    try:
        config = get_config()
        config.validate()
    except Exception as e:
        error_msg = f"Configuration error: {e}\n\nPlease run: python scripts/connect_graph.py"
        messagebox.showerror("Configuration Error", error_msg)
        logger.error(f"Configuration error: {e}")
        return 1

    # Authenticate
    try:
        authenticator = create_authenticator_from_config(config.to_dict())
        account = authenticator.get_account_info()

        if not account:
            error_msg = "Not authenticated to Microsoft Graph.\n\nPlease run: python scripts/connect_graph.py"
            messagebox.showwarning("Authentication Required", error_msg)
            logger.error("Not authenticated")
            return 1

        graph_client = GraphClient(authenticator)
        calendar_client = CalendarClient(graph_client, config.timezone)
        mail_client = MailClient(graph_client)

    except Exception as e:
        error_msg = f"Authentication error: {e}\n\nPlease run: python scripts/connect_graph.py"
        messagebox.showerror("Authentication Error", error_msg)
        logger.error(f"Authentication error: {e}")
        return 1

    # Load ignore patterns
    ignore_patterns = load_ignore_patterns(config_dir)
    logger.info(f"Loaded {len(ignore_patterns)} ignore patterns from config/ignore_appointments.txt")
    for pattern in ignore_patterns:
        logger.info(f"  - Pattern: {pattern}")
    logger.info("")

    # Define time periods
    try:
        from zoneinfo import ZoneInfo
    except ImportError:
        from backports.zoneinfo import ZoneInfo

    tz = ZoneInfo(config.timezone)
    now = datetime.now(tz)
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    next_working_day = get_next_working_day(now, config.timezone)
    day_after_next = get_next_working_day(next_working_day, config.timezone)

    # Get week bounds
    this_week_bounds = get_weekday_bounds(now, config.timezone)
    next_week_monday = this_week_bounds[0] + timedelta(days=7)
    next_week_bounds = get_weekday_bounds(next_week_monday, config.timezone)

    # Fetch calendar items
    fetch_start = today
    fourteen_days_later = today + timedelta(days=14)
    fetch_end = max(next_week_bounds[1], fourteen_days_later)

    logger.info(f"Connecting to Microsoft Graph...")
    logger.info(f"Fetching calendar events from {fetch_start.date()} to {fetch_end.date()}")
    logger.info("")

    try:
        events = calendar_client.list_events(fetch_start, fetch_end + timedelta(days=1))
        logger.info("")
        logger.info("=" * 70)
        logger.info("ALL CALENDAR EVENTS RETRIEVED FROM GRAPH")
        logger.info("=" * 70)
        logger.info("")

        # Sort and log all events
        sorted_events = sorted(events, key=lambda e: from_graph_datetime(e.get("start", {}), config.timezone))
        for event in sorted_events:
            event_start = from_graph_datetime(event.get("start", {}), config.timezone)
            logger.info(f"  {event_start.strftime('%Y-%m-%d %H:%M')} | {event.get('subject', 'Untitled')}")
        logger.info("")

    except Exception as e:
        error_msg = f"Failed to retrieve calendar events.\n\nError: {e}\n\nPlease ensure you are authenticated."
        messagebox.showerror("Calendar Error", error_msg)
        logger.error(f"Failed to retrieve calendar events: {e}")
        return 1

    # Calculate meeting hours
    logger.info("=" * 70)
    logger.info("CALCULATING MEETING HOURS")
    logger.info("=" * 70)
    logger.info("")

    today_hours = get_meeting_hours_summary(events, today, next_working_day, ignore_patterns, config.timezone, "Today")
    next_day_hours = get_meeting_hours_summary(events, next_working_day, day_after_next, ignore_patterns, config.timezone, "Next Working Day")
    this_week_hours = get_meeting_hours_summary(events, this_week_bounds[0], this_week_bounds[1], ignore_patterns, config.timezone, "This Week")
    next_week_hours = get_meeting_hours_summary(events, next_week_bounds[0], next_week_bounds[1], ignore_patterns, config.timezone, "Next Week")

    # Calculate daily hours for bar chart
    logger.info("=" * 70)
    logger.info("CALCULATING DAILY HOURS FOR BAR CHART (NEXT 5 WORKING DAYS)")
    logger.info("=" * 70)
    logger.info("")

    daily_hours = get_daily_meeting_hours(events, today, 5, ignore_patterns, config.timezone)

    # Process full-hour meetings
    email_template = load_email_template(config_dir)
    ignored_ids = load_ignored_appointment_ids(config_dir)
    logger.info(f"Loaded {len(ignored_ids)} ignored full-hour appointment(s)")
    logger.info("")

    logger.info("=" * 70)
    logger.info("SCANNING FOR FULL-HOUR MEETINGS")
    logger.info("=" * 70)
    logger.info("")

    full_hour_meetings = get_full_hour_meetings(events, now, fourteen_days_later, ignore_patterns, ignored_ids, config.timezone, max_count=10)

    # Process each full-hour meeting
    for meeting in full_hour_meetings:
        meeting_start = from_graph_datetime(meeting.get("start", {}), config.timezone)
        organizer_name = meeting.get('organizer', {}).get('emailAddress', {}).get('name', 'Unknown')

        # Show dialog
        user_choice = show_reschedule_dialog(meeting.get('subject', 'Untitled'), meeting_start, organizer_name)

        if user_choice == "CreateDraft":
            success = create_reschedule_draft(meeting, email_template, mail_client, config.timezone, config_dir)
            if success:
                logger.info(f"Draft email created for meeting: '{meeting.get('subject', 'Untitled')}'")
                messagebox.showinfo("Draft Created", "Draft email created successfully and saved to your Drafts folder.")
            else:
                logger.error(f"Failed to create draft email for meeting: '{meeting.get('subject', 'Untitled')}'")
                messagebox.showerror("Error", "Failed to create draft email. Please check the log file.")

        elif user_choice == "NeverAskAgain":
            event_id = meeting.get("iCalUId") or meeting.get("id", "")
            if event_id:
                save_ignored_appointment(event_id, meeting.get('subject', 'Untitled'), meeting_start, config_dir)
                messagebox.showinfo("Added to Ignore List", "This meeting has been added to the ignore list and will not be shown again.\n\nYou can manually edit config/ignored_full_hour_appointments.txt to remove it if needed.")
            else:
                logger.warning(f"Could not get identifier for meeting: '{meeting.get('subject', 'Untitled')}'")
                messagebox.showwarning("Warning", "Unable to get a stable identifier for this meeting. It cannot be added to the ignore list.")

        else:
            logger.info(f"User skipped meeting: '{meeting.get('subject', 'Untitled')}'")

    # Log final summary
    logger.info("=" * 70)
    logger.info("EXECUTION SUMMARY")
    logger.info("=" * 70)
    logger.info(f"Today: {today_hours['hours']} hours ({today_hours['count']} meetings)")
    logger.info(f"Next Working Day ({next_working_day.strftime('%Y-%m-%d')}): {next_day_hours['hours']} hours ({next_day_hours['count']} meetings)")
    logger.info(f"This Week: {this_week_hours['hours']} hours ({this_week_hours['count']} meetings)")
    logger.info(f"Next Week: {next_week_hours['hours']} hours ({next_week_hours['count']} meetings)")
    logger.info(f"Full-hour meetings found: {len(full_hour_meetings)}")
    logger.info("")
    logger.info("Script execution completed successfully")
    logger.info("=" * 70)

    print(f"Processing complete. Detailed log saved to: {log_file}")
    print("Check the log file to see all events found and filtering decisions.")
    print()

    # Show GUI
    show_meeting_summary_gui(
        today_hours,
        next_day_hours,
        this_week_hours,
        next_week_hours,
        daily_hours,
        today,
        next_working_day,
        this_week_bounds,
        next_week_bounds,
        len(ignore_patterns)
    )

    print(f"For detailed information about which events were included/excluded, see: {log_file}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
