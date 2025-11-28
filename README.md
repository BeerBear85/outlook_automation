# Outlook Meeting Hour Summary

Scripts that display meeting hour summaries from your Outlook calendar with a visual bar chart, and help optimize meeting schedules by suggesting time shifts for full-hour meetings.

## Three Implementations Available

This repository contains **three parallel implementations** of the same core functionality:

### 1. Python + Microsoft Graph (Newest)
**Location:** `scripts_using_python/`

- ✅ **Works with new Outlook** (Outlook for Windows)
- ✅ Works with classic Outlook
- ✅ Cloud-based (no local Outlook installation required)
- ✅ **Cross-platform** (Windows, macOS, Linux)
- ✅ Modern Python implementation with full test coverage
- 📖 See [scripts_using_python/README.md](scripts_using_python/README.md) for detailed documentation

**Use this if:** You prefer Python, need cross-platform support, or want the most modern implementation.

### 2. PowerShell + Microsoft Graph (Recommended for Windows)
**Location:** `scripts_using_graph/`

- ✅ **Works with new Outlook** (Outlook for Windows)
- ✅ Works with classic Outlook
- ✅ Cloud-based (no local Outlook installation required)
- ✅ Future-proof solution
- ✅ Native PowerShell experience
- 📖 See [scripts_using_graph/README.md](scripts_using_graph/README.md) for detailed documentation

**Use this if:** You prefer PowerShell and have new Outlook, or want a modern, cloud-first solution.

### 3. COM-based Implementation (Legacy)
**Location:** `scripts_using_com/`

- ⚠️ **Does NOT work with new Outlook**
- ✅ Works with classic Outlook only
- ✅ Faster (uses local Outlook cache)
- ✅ Works offline
- 📖 Documentation below applies to this implementation

**Use this if:** You have classic Outlook and don't plan to migrate to new Outlook.

---

**This README documents the COM-based implementation.**
- For PowerShell Graph implementation, see [scripts_using_graph/README.md](scripts_using_graph/README.md)
- For Python Graph implementation, see [scripts_using_python/README.md](scripts_using_python/README.md)

---

## Features

### Meeting Hour Summary with Visual Bar Chart
- Displays total meeting hours for:
  - Today
  - Next Working Day (skips weekends)
  - This week (Monday-Friday)
  - Next week (Monday-Friday)
- **5-Day Bar Chart** showing the next 5 working days:
  - Color-coded bars: Green (0-3h), Yellow (3-4h), Red (4+h)
  - Visual overview of upcoming meeting load
  - Helps identify overloaded days at a glance
  - Automatically skips weekends
- Excludes appointments matching configurable regex patterns
- Shows results in a clean Windows Forms popup

### Full-Hour Meeting Optimizer
- Scans calendar for meetings starting exactly on the hour (e.g., 10:00, 11:00)
- Suggests shifting meetings to :05 (e.g., 10:00 → 10:05) for better work-life balance
- Creates draft emails to organizers with customizable templates
- **"Never Ask Again"** option to permanently skip specific meetings
- Processes up to 10 meetings at a time
- Excludes all-day events, private meetings, and Out of Office entries
- Emails are saved as drafts (never sent automatically)

## Prerequisites

- Windows operating system
- Microsoft Outlook installed and configured
- PowerShell 5.1 or higher
- Outlook must be accessible via COM (not supported when running as Administrator)

## Installation

1. Clone or download this repository
2. Ensure Outlook is installed and configured with your email account
3. (Optional) Configure ignore patterns and email template (see Configuration section)

## Usage

### Running the Script

Navigate to the `scripts_using_com` directory and run:

```powershell
.\Show-MeetingHourSummary.ps1
```

Or from the repository root:

```powershell
.\scripts_using_com\Show-MeetingHourSummary.ps1
```

### What Happens When You Run the Script

1. **Full-Hour Meeting Check** (if any found):
   - Popup appears for each meeting starting on the hour
   - Shows meeting subject, start time, and organizer
   - Three options:
     - **Create Draft Email** - Create a draft requesting time shift to :05
     - **Skip for Now** - Skip this meeting this time
     - **Never Ask Again** - Permanently ignore this meeting
   - Draft emails are saved to your Outlook Drafts folder

2. **Meeting Hour Summary with Bar Chart**:
   - Main summary window appears showing:
     - Total hours and meeting counts for today, next working day, this week, and next week
     - **5-day color-coded bar chart** showing upcoming meeting load
     - Green bars (0-3 hours), Yellow bars (3-4 hours), Red bars (4+ hours)
   - Click **OK** to close

## Configuration

### Ignore Patterns (ignore_appointments.txt)

Create `scripts_using_com/ignore_appointments.txt` to exclude certain appointments from calculations:

```text
# This file contains regex patterns for appointments to ignore
# One pattern per line, lines starting with # are comments

# Ignore lunch breaks
^Lunch$

# Ignore personal appointments
^Personal.*

# Ignore cancelled meetings
.*\[CANCELLED\].*

# Ignore specific recurring meetings
^Daily Standup$
```

**Pattern syntax:**
- Uses regex (case-sensitive by default)
- Use `(?i)` prefix for case-insensitive matching: `(?i)^lunch$`
- Use `.*` for wildcards: `^Personal.*` matches "Personal: Doctor", "Personal Time", etc.

### Email Template (meeting_change_request_template.txt)

Customize the draft email template in `scripts_using_com/meeting_change_request_template.txt`:

```text
Subject: Request to shift meeting start time to :05

Dear {ORGANIZER},

I hope this message finds you well. I'm reaching out regarding our upcoming meeting:

Meeting: {SUBJECT}
Current Start Time: {START_TIME}

Would it be possible to shift the meeting start time by 5 minutes to {NEW_START_TIME}? This small adjustment would help create a buffer between back-to-back meetings and allow for better preparation time.

If this change works for you and other attendees, I would greatly appreciate it. If the current time is critical, please feel free to keep it as scheduled.

Thank you for considering this request.

Best regards
```

**Available placeholders:**
- `{ORGANIZER}` - Meeting organizer's name
- `{SUBJECT}` - Meeting subject
- `{START_TIME}` - Current meeting start time
- `{NEW_START_TIME}` - Proposed new time (:05 past the hour)

## Testing

The project includes comprehensive Pester tests for all functions.

### Running Tests

From the `scripts_using_com` directory:

```powershell
# Install Pester if not already installed
Install-Module -Name Pester -Force -SkipPublisherCheck

# Run all tests
.\Run-Tests.ps1

# Or run tests directly with Pester
Invoke-Pester -Path .\Show-MeetingHourSummary.Tests.ps1
```

### Test Coverage

- Date calculations (week bounds, DST handling, leap years)
- Appointment duration calculations
- Ignore pattern matching
- Full-hour meeting detection
- Email template loading
- Edge cases and boundary conditions

## Troubleshooting

### "This script cannot be run with administrative privileges"

**Solution:** Close PowerShell and reopen without "Run as Administrator". Outlook COM access doesn't work with elevated privileges.

### "Failed to connect to Outlook"

**Possible causes:**
- Outlook is not installed
- Outlook is not configured with an email account
- Outlook is not running (the script will attempt to start it)

**Solution:** Open Outlook manually and ensure it's working properly.

### No full-hour meetings detected

**Possible reasons:**
- No meetings in the next 14 days starting exactly on the hour
- Meetings are private or marked as Out of Office
- Meetings match patterns in `ignore_appointments.txt`
- All qualifying meetings have already started

### Draft emails not appearing in Drafts folder

**Check:**
1. Look in Outlook's Drafts folder
2. If using multiple email accounts, check the correct account's Drafts folder
3. Ensure Outlook synchronized properly (try sending/receiving)

## File Structure

```
outlook_automation/
├── README.md                                    # This file (overview of all implementations)
├── scripts_using_com/                           # COM-based implementation (legacy)
│   ├── Show-MeetingHourSummary.ps1             # Main script
│   ├── Show-MeetingHourSummary.Tests.ps1       # Pester tests
│   ├── Run-Tests.ps1                            # Test runner
│   ├── README_TESTS.md                          # Test documentation
│   ├── ignore_appointments.txt                  # Ignore patterns (optional)
│   ├── ignored_full_hour_appointments.txt       # Never Ask Again list (auto-generated)
│   ├── meeting_change_request_template.txt      # Email template (customizable)
│   └── log.txt                                  # Execution log (auto-generated)
├── scripts_using_graph/                         # PowerShell Graph implementation
│   ├── README.md                                # Graph PowerShell documentation
│   ├── Connect-Graph.ps1                        # Authentication script
│   ├── Show-MeetingHourSummary.ps1             # Main script (Graph version)
│   ├── Test-GraphConnection.ps1                 # Connection validation
│   ├── OutlookGraphAutomation.psm1             # Shared module
│   └── config/                                  # Configuration files
│       ├── ignore_appointments.txt
│       ├── meeting_change_request_template.txt
│       └── ignored_full_hour_appointments.txt
├── scripts_using_python/                        # Python Graph implementation (newest)
│   ├── README.md                                # Python implementation documentation
│   ├── requirements.txt                         # Python dependencies
│   ├── pytest.ini                               # Test configuration
│   ├── run_tests.py                             # Test runner
│   ├── config.example.json                      # Example configuration
│   ├── .env.example                             # Example environment variables
│   ├── src/outlook_graph/                       # Python package
│   │   ├── auth.py                              # Authentication (MSAL)
│   │   ├── config.py                            # Configuration management
│   │   ├── calendar.py                          # Calendar operations
│   │   ├── mail.py                              # Mail operations
│   │   └── utils.py                             # Utility functions
│   ├── scripts/                                 # Entry point scripts
│   │   ├── connect_graph.py                     # Authentication
│   │   ├── test_connection.py                   # Connection testing
│   │   └── show_meeting_summary.py              # Main script
│   ├── tests/                                   # Automated tests (62 tests)
│   │   ├── test_utils.py                        # Unit tests
│   │   ├── test_config.py                       # Config tests
│   │   └── test_integration.py                  # Integration tests
│   └── config/                                  # Configuration files
│       ├── ignore_appointments.txt
│       ├── meeting_change_request_template.txt
│       └── ignored_full_hour_appointments.txt
└── example/
    └── update_calendar_busytype.ps1             # Other utilities
```

## How It Works

### Meeting Hour Summary with Bar Chart
1. Connects to Outlook via COM objects
2. Retrieves calendar items from today through next week
3. Filters appointments based on ignore patterns
4. Excludes all-day events, cancelled meetings, and declined meetings
5. Calculates total hours for each time period:
   - Today
   - Next Working Day (automatically skips weekends)
   - This Week (Monday-Friday)
   - Next Week (Monday-Friday)
6. **Calculates daily hours for the next 5 working days**:
   - Automatically skips weekends (Saturday and Sunday)
   - Determines color coding based on hours per day
7. Displays results in a Windows Forms popup with color-coded bar chart

### Full-Hour Meeting Optimizer
1. Scans calendar items for the next 14 days
2. Identifies meetings starting exactly at :00 (minute = 0, second = 0)
3. Applies filters (all-day, private, OOO, ignore patterns, cancelled, declined)
4. Excludes previously ignored meetings (from ignored_full_hour_appointments.txt)
5. Sorts by earliest start time
6. Limits to 10 meetings per run
7. Shows confirmation dialog for each meeting with three options
8. Creates draft email using template with placeholders replaced
9. Saves to Drafts folder (never sends automatically)
10. Saves "Never Ask Again" selections to ignored_full_hour_appointments.txt

## Notes

- The script works with your default Outlook calendar
- All draft emails must be manually reviewed and sent by you
- Private meetings and Out of Office are never included
- The script respects your existing ignore patterns for both features
- Line endings: The script uses CRLF (Windows) line endings

## Contributing

When making changes:
1. Update tests in `Show-MeetingHourSummary.Tests.ps1`
2. Run all tests to ensure nothing breaks: `.\Run-Tests.ps1`
3. Update this README if adding new features

## License

Generated for outlook_automation repository
