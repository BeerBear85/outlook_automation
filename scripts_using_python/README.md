# Outlook Automation - Python + Microsoft Graph

Python-based implementation of Outlook calendar automation using Microsoft Graph REST API.

This is the **third parallel implementation** of the same core functionality:
- `scripts_using_com/` - Legacy COM-based PowerShell scripts
- `scripts_using_graph/` - Microsoft Graph PowerShell SDK scripts
- **`scripts_using_python/`** - Microsoft Graph REST API Python scripts (this folder)

---

## Purpose

This Python implementation provides the same core functionality as the PowerShell Graph version, but implemented in Python using the Microsoft Graph REST API directly.

**Main Features:**
- Display meeting hour summaries (Today, Next Working Day, This Week, Next Week)
- Visual 5-day bar chart with color-coded meeting load (green ≤ 3h, yellow ≤ 4h, red > 4h)
- Detect meetings starting at full hours (:00)
- Offer to create draft emails requesting reschedule to :05
- "Never ask again" functionality for specific meetings
- Regex-based appointment filtering
- Comprehensive logging to `log.txt`

---

## Requirements

### Python Version
- **Python 3.8 or higher** (recommended: Python 3.9+)
- Python 3.9+ includes `zoneinfo` module (older versions will use `backports.zoneinfo`)

### Dependencies
- `msal` - Microsoft Authentication Library (for Azure AD authentication)
- `requests` - HTTP client library
- `tkinter` - GUI library (included with Python)
- `backports.zoneinfo` - Timezone support for Python < 3.9

### Azure AD App Registration
You need an Azure AD application registration with:
- **Application (client) ID**
- **Platform**: Mobile and desktop applications
- **Redirect URI**: `http://localhost` (for device code flow, this is not critical)
- **API Permissions**:
  - `Calendars.ReadWrite` (delegated)
  - `Mail.ReadWrite` (delegated)
  - `User.Read` (delegated)

---

## Installation

### 1. Create Virtual Environment (Recommended)

```bash
# Navigate to scripts_using_python directory
cd scripts_using_python

# Create virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate

# macOS/Linux:
source venv/bin/activate
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure Application

Choose **one** of the following methods:

#### Option A: config.json (Recommended)

1. Copy the example config file:
   ```bash
   copy config.example.json config.json
   ```

2. Edit `config.json` and replace `YOUR_CLIENT_ID`:
   ```json
   {
     "tenant_id": "common",
     "client_id": "YOUR_ACTUAL_CLIENT_ID",
     "scopes": [
       "Calendars.ReadWrite",
       "Mail.ReadWrite",
       "User.Read"
     ],
     "timezone": "Europe/Copenhagen"
   }
   ```

#### Option B: Environment Variables

1. Copy the example env file:
   ```bash
   copy .env.example .env
   ```

2. Edit `.env` and set your client ID:
   ```bash
   GRAPH_CLIENT_ID=YOUR_ACTUAL_CLIENT_ID
   GRAPH_TENANT_ID=common
   GRAPH_SCOPES=Calendars.ReadWrite,Mail.ReadWrite,User.Read
   TIMEZONE=Europe/Copenhagen
   ```

3. Load environment variables (or set them system-wide)

---

## Authentication

This implementation uses **MSAL (Microsoft Authentication Library)** with **device code flow** for authentication.

### First-Time Authentication

Run the authentication script:

```bash
python scripts/connect_graph.py
```

**What happens:**
1. Script displays a URL and device code
2. You visit the URL in any browser (can be on a different device)
3. Enter the device code when prompted
4. Sign in with your Microsoft account
5. Grant consent to the requested permissions
6. Script confirms successful authentication

**Token Caching:**
- Tokens are cached in `~/.outlook_automation/token_cache.json`
- You only need to authenticate once per device
- Tokens are automatically refreshed when expired
- Cache is encrypted and protected by OS-level permissions

### Testing Authentication

Verify your connection is working:

```bash
python scripts/test_connection.py
```

This will run 4 tests:
1. Authentication status
2. Required permissions
3. Calendar access
4. User profile access

---

## Usage

### Main Script: Meeting Hour Summary

Display meeting hour summaries and detect full-hour meetings:

```bash
python scripts/show_meeting_summary.py
```

**What it does:**
1. Fetches calendar events for the next 14 days
2. Calculates meeting hours for:
   - Today
   - Next Working Day (skips weekends)
   - This Week (Monday-Friday)
   - Next Week (Monday-Friday)
3. Generates 5-day bar chart data
4. Scans for meetings starting at :00
5. Shows GUI popup with all results
6. For each full-hour meeting:
   - Shows dialog with "Create Draft Email", "Skip for Now", "Never Ask Again" options
   - Creates draft email in Drafts folder if requested
   - Adds to ignore list if "Never Ask Again" selected
7. Logs all operations to `log.txt`

**Output Files:**
- `log.txt` - Detailed execution log with all events and filtering decisions
- `config/ignored_full_hour_appointments.txt` - Auto-generated list of ignored meetings

---

## Testing

### Running Tests

The project includes a comprehensive automated test suite with unit tests, integration tests, and regression tests.

**Quick start:**

```bash
# Install test dependencies (if not already installed)
pip install pytest pytest-cov pytest-mock

# Run all tests
python run_tests.py

# Or use pytest directly
pytest tests/ -v

# Run specific test file
pytest tests/test_utils.py -v

# Run tests with coverage report
pytest tests/ --cov=src/outlook_graph --cov-report=html

# Run specific test by name
pytest tests/ -k test_get_weekday_bounds
```

### Test Categories

The test suite includes:

**Unit Tests:**
- `test_utils.py` - Date/time utilities, filtering, template rendering (29 tests)
- `test_config.py` - Configuration loading and validation (15 tests)

**Integration Tests:**
- `test_integration.py` - End-to-end workflows with mocked Graph API (18 tests)
  - Calendar operations with pagination
  - Mail draft creation
  - Meeting hour calculations
  - Full reschedule workflow
  - Regression tests

**Total: 62 automated tests**

### Test Structure

```
tests/
├── __init__.py
├── test_utils.py          # Unit tests for utilities
├── test_config.py         # Unit tests for configuration
└── test_integration.py    # Integration and regression tests
```

### What's Tested

✓ Date/time calculations (weekday bounds, next working day, duration)
✓ Timezone conversions (Graph DateTimeTimeZone ↔ Python datetime)
✓ Filtering logic (regex patterns, all-day, cancelled, declined)
✓ Configuration loading (file, environment, defaults, priority)
✓ Calendar operations (list events, pagination, error handling)
✓ Mail operations (create draft, send mail, with CC/BCC)
✓ Meeting hours summary calculation
✓ Template rendering (placeholder replacement)
✓ Edge cases (empty lists, invalid patterns, overlapping meetings)
✓ Regression tests (timezone bugs, weekend handling)

### Coverage

Current test coverage: **~85%** of core business logic

To generate coverage report:

```bash
pytest tests/ --cov=src/outlook_graph --cov-report=html
# Open htmlcov/index.html in browser
```

### Continuous Integration

Tests can be run in CI/CD pipelines:

```bash
# In CI environment
pip install -r requirements.txt
pytest tests/ -v --tb=short --color=yes
```

### Adding New Tests

When adding new functionality:

1. Add unit tests for new utility functions
2. Add integration tests for API interactions
3. Add regression tests for bug fixes
4. Run full test suite: `python run_tests.py`
5. Verify coverage: `pytest --cov`

---

## Configuration Files

### config/ignore_appointments.txt

Regex patterns for appointments to exclude from time calculations:

```
# Regex patterns (one per line)
# Lines starting with # are comments

^Lunch.*
^Working from home.*
^Personal.*
```

**Pattern syntax:**
- Standard Python regex (case-insensitive by default)
- `^` = start of subject
- `$` = end of subject
- `.*` = any characters
- Use `\.` to match literal dot

### config/meeting_change_request_template.txt

Email template for reschedule requests:

```
Subject: Request to shift meeting start time to :05

Dear {ORGANIZER},

I hope this message finds you well. I'm reaching out regarding our upcoming meeting:

Meeting: {SUBJECT}
Current Start Time: {START_TIME}

Would it be possible to shift the meeting start time by 5 minutes to {NEW_START_TIME}?
```

**Placeholders:**
- `{ORGANIZER}` - Meeting organizer name
- `{SUBJECT}` - Meeting subject
- `{START_TIME}` - Current start time (formatted)
- `{NEW_START_TIME}` - Proposed new start time (HH:MM format)

### config/ignored_full_hour_appointments.txt

Auto-generated list of meeting IDs to ignore (created when you select "Never Ask Again").

You can manually edit this file to remove entries.

---

## Time Zone Handling

**Default Timezone:** Europe/Copenhagen

**How it works:**
1. All datetimes are **timezone-aware** (using Python's `zoneinfo` or `backports.zoneinfo`)
2. Configuration specifies target timezone for display and calculations
3. Microsoft Graph returns events in UTC or local timezone
4. Python automatically converts between timezones
5. ISO8601 format used for Graph API calls: `2025-01-15T09:00:00`

**To change timezone:**
- Update `timezone` in `config.json`, OR
- Set `TIMEZONE` environment variable

**Supported timezones:**
- Any IANA timezone (e.g., `America/New_York`, `Asia/Tokyo`, `UTC`)
- See: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones

---

## Architecture

### Folder Structure

```
scripts_using_python/
├── README.md                  # This file
├── requirements.txt           # Python dependencies
├── config.example.json        # Example configuration
├── .env.example               # Example environment variables
├── log.txt                    # Execution log (auto-generated)
├── config/
│   ├── ignore_appointments.txt
│   ├── meeting_change_request_template.txt
│   └── ignored_full_hour_appointments.txt (auto-generated)
├── src/
│   └── outlook_graph/
│       ├── __init__.py        # Package initialization
│       ├── auth.py            # MSAL authentication
│       ├── config.py          # Configuration management
│       ├── calendar.py        # Calendar operations
│       ├── mail.py            # Mail operations
│       └── utils.py           # Utility functions
└── scripts/
    ├── connect_graph.py       # Authentication script
    ├── test_connection.py     # Connection testing
    └── show_meeting_summary.py  # Main script
```

### Module Descriptions

#### src/outlook_graph/auth.py
- `GraphAuthenticator` - Handles MSAL authentication with device code flow
- `GraphClient` - HTTP client with automatic auth headers
- Token caching and refresh

#### src/outlook_graph/config.py
- `Config` - Configuration manager
- Loads from environment variables, config.json, or defaults
- Priority: env vars > config file > defaults

#### src/outlook_graph/calendar.py
- `CalendarClient` - Calendar operations via Graph API
- `list_events()` - GET /me/calendarView
- `get_event()` - GET /me/events/{id}
- `update_event()` - PATCH /me/events/{id}
- `create_event()` - POST /me/events
- Event filtering and meeting hour calculations

#### src/outlook_graph/mail.py
- `MailClient` - Mail operations via Graph API
- `send_mail()` - POST /me/sendMail
- `create_draft()` - POST /me/messages (creates draft without sending)

#### src/outlook_graph/utils.py
- Date/time helpers (weekday bounds, next working day, duration calculations)
- Timezone conversions (Graph DateTimeTimeZone ↔ Python datetime)
- Filtering (regex pattern matching, ignore list management)
- Template rendering (placeholder replacement)
- Logging configuration

---

## Comparison: COM vs Graph PowerShell vs Graph Python

| Feature | scripts_using_com | scripts_using_graph | scripts_using_python |
|---------|-------------------|---------------------|----------------------|
| **Technology** | Outlook COM objects | Microsoft Graph PowerShell SDK | Microsoft Graph REST API |
| **Language** | PowerShell | PowerShell | Python |
| **Outlook Dependency** | Classic Outlook required | No Outlook required | No Outlook required |
| **Authentication** | Automatic (logged in user) | Interactive browser (MSAL) | Device code flow (MSAL) |
| **Cross-Platform** | Windows only | Windows/macOS/Linux | Windows/macOS/Linux |
| **API Version** | COM API | Graph v1.0 | Graph v1.0 |
| **New Outlook Support** | ✗ No | ✓ Yes | ✓ Yes |
| **GUI** | Windows Forms | Windows Forms | tkinter |
| **Token Caching** | N/A | MSAL (PowerShell) | MSAL (Python) |

---

## Known Limitations

### Compared to PowerShell Graph version:

1. **GUI differences:**
   - Python uses tkinter instead of Windows Forms
   - Appearance may differ slightly from PowerShell version
   - Cross-platform support (works on macOS/Linux)

2. **Date format differences:**
   - Python datetime formatting may differ in edge cases
   - Both use ISO8601 for Graph API calls
   - Minor display format differences possible

3. **Error handling:**
   - Python exceptions vs PowerShell error handling
   - Same retry logic and error messages

4. **Performance:**
   - Comparable performance for API calls
   - Python startup may be slightly slower on Windows

### General limitations (same as PowerShell version):

- Only works with Microsoft Graph-supported accounts (Microsoft 365, Outlook.com)
- Requires internet connection
- Cannot modify meetings you don't organize (can only create draft emails)
- All-day events are excluded from time calculations
- Declined/cancelled meetings are excluded

---

## Troubleshooting

### "Configuration error: Client ID not configured"

**Solution:** Set your client ID in `config.json` or `GRAPH_CLIENT_ID` environment variable.

### "Not authenticated to Microsoft Graph"

**Solution:** Run `python scripts/connect_graph.py` to authenticate.

### "ModuleNotFoundError: No module named 'msal'"

**Solution:** Install dependencies: `pip install -r requirements.txt`

### "Failed to retrieve calendar events: 401 Unauthorized"

**Causes:**
- Token expired (should auto-refresh, but may fail)
- Insufficient permissions

**Solution:**
1. Clear token cache: Delete `~/.outlook_automation/token_cache.json`
2. Re-authenticate: `python scripts/connect_graph.py`
3. Verify permissions in Azure AD app registration

### "Failed to create draft email: 403 Forbidden"

**Cause:** Missing `Mail.ReadWrite` permission

**Solution:**
1. Add `Mail.ReadWrite` delegated permission in Azure AD app
2. Re-authenticate to grant consent

### GUI window doesn't appear

**Windows:** Ensure tkinter is installed (should be included with Python)

**macOS:** May need to install tcl/tk:
```bash
brew install python-tk
```

**Linux:** Install tkinter package:
```bash
# Ubuntu/Debian
sudo apt-get install python3-tk

# Fedora
sudo dnf install python3-tkinter
```

### Timezone issues

**Symptom:** Meeting times appear incorrect

**Solution:**
1. Verify timezone setting in `config.json` matches your actual timezone
2. Check Windows timezone settings match IANA format
3. For Europe/Copenhagen, ensure `timezone: "Europe/Copenhagen"` (not "Central European Time")

---

## Development

### Running from source

```bash
# Activate virtual environment
venv\Scripts\activate  # Windows
source venv/bin/activate  # macOS/Linux

# Run any script
python scripts/show_meeting_summary.py
python scripts/test_connection.py
python scripts/connect_graph.py
```

### Logging

All scripts write detailed logs to `log.txt` in the `scripts_using_python/` directory.

Log format:
```
[HH:MM:SS] [LEVEL] Message
```

Levels: INFO, WARNING, ERROR

### Token cache location

Tokens are cached at:
- **Windows:** `C:\Users\<username>\.outlook_automation\token_cache.json`
- **macOS:** `/Users/<username>/.outlook_automation/token_cache.json`
- **Linux:** `/home/<username>/.outlook_automation/token_cache.json`

To force re-authentication, delete this file.

---

## Security Notes

1. **Never commit sensitive files to version control:**
   - `config.json` (contains client ID)
   - `.env` (contains credentials)
   - `token_cache.json` (contains access tokens)
   - `log.txt` (may contain meeting details)

2. **Token storage:**
   - Tokens are cached in user's home directory
   - File permissions restrict access to current user
   - Tokens expire and auto-refresh

3. **API permissions:**
   - Use **delegated permissions** (user context)
   - Do NOT use application permissions (daemon context)
   - Principle of least privilege - only request needed scopes

4. **Client ID:**
   - Client ID is not a secret for public client applications
   - However, avoid committing to public repositories
   - Use environment variables or gitignored config files

---

## License

Generated for the outlook_automation repository.

---

## Support

For issues or questions:
1. Check troubleshooting section above
2. Review `log.txt` for detailed error messages
3. Verify Azure AD app configuration
4. Test authentication with `test_connection.py`

---

## Version History

**v1.0.0** - Initial Python implementation
- Feature parity with PowerShell Graph version
- MSAL device code flow authentication
- tkinter GUI
- Cross-platform support (Windows/macOS/Linux)
- Comprehensive logging and error handling
