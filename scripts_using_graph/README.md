# Outlook Automation - Microsoft Graph Implementation

This folder contains the **Microsoft Graph PowerShell** implementation of Outlook automation scripts, migrated from the legacy COM-based implementation.

**Why this migration?** The new Outlook (Outlook for Windows, formerly "Project Monarch") **does not support COM automation**. Microsoft Graph is the modern, cloud-first API that works with both classic and new Outlook.

---

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Authentication](#authentication)
- [Usage](#usage)
- [Configuration](#configuration)
- [COM to Graph Migration Notes](#com-to-graph-migration-notes)
- [Folder Structure](#folder-structure)
- [Troubleshooting](#troubleshooting)
- [Differences from COM Implementation](#differences-from-com-implementation)

---

## Features

### Show-MeetingHourSummary.ps1

The main automation script provides:

1. **Meeting Hour Summaries:**
   - Today's total hours
   - Next working day hours (automatically skips weekends)
   - This week hours (Monday-Friday)
   - Next week hours (Monday-Friday)

2. **5-Day Bar Chart:**
   - Visual representation of meeting load for next 5 working days
   - Color-coded: Green (≤3h), Yellow (≤4h), Red (>4h)
   - Helps identify overloaded days at a glance

3. **Full-Hour Meeting Reschedule Assistant:**
   - Scans for meetings starting exactly at :00 (full hour)
   - Offers to create draft email requesting start time shift to :05
   - Provides buffer time between back-to-back meetings
   - Supports "Never Ask Again" functionality

4. **Intelligent Filtering:**
   - Regex pattern-based appointment exclusion
   - Excludes all-day events, private items, cancelled meetings
   - Excludes declined meetings and Out of Office blocks
   - Fully customizable via configuration files

5. **Comprehensive Logging:**
   - Detailed log.txt file with all processing steps
   - Helps troubleshoot filtering decisions
   - Tracks all calendar events retrieved

---

## Prerequisites

### Required Software

1. **PowerShell 5.1 or later**
   - Windows 10/11 includes this by default
   - Check version: `$PSVersionTable.PSVersion`

2. **Microsoft.Graph PowerShell Module**
   - Version 2.0 or later recommended
   - Contains cmdlets for interacting with Microsoft Graph API

3. **Microsoft Account with Calendar Access**
   - Personal Microsoft account or
   - Work/School account (Microsoft 365)

### Permissions

The scripts require the following Microsoft Graph permissions (delegated):
- `Calendars.ReadWrite` - Read and write calendar events
- `Mail.ReadWrite` - Create draft emails
- `User.Read` - Read user profile information

These are **user-delegated permissions** - you authenticate as yourself and the scripts act on your behalf. No admin consent required.

---

## Installation

### Step 1: Install Microsoft Graph PowerShell Module

Open PowerShell (as regular user, **not administrator**) and run:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

This may take a few minutes as it installs several sub-modules.

**Verify installation:**

```powershell
Get-Module Microsoft.Graph.* -ListAvailable
```

You should see multiple modules like `Microsoft.Graph.Authentication`, `Microsoft.Graph.Calendar`, `Microsoft.Graph.Mail`, etc.

### Step 2: Clone or Download Repository

If you're reading this, you likely already have the repository. If not:

```powershell
git clone https://github.com/yourusername/outlook_automation.git
cd outlook_automation\scripts_using_graph
```

---

## Authentication

### Interactive Authentication (One-Time Per Session)

Before running any automation scripts, you must authenticate to Microsoft Graph:

```powershell
cd scripts_using_graph
.\Connect-Graph.ps1
```

**What happens:**
1. A browser window opens
2. Sign in with your Microsoft account
3. Consent to the requested permissions (first time only)
4. PowerShell confirms successful connection

**Output example:**

```
Microsoft Graph Connection Script
=================================

[1/4] Checking for Microsoft.Graph module...
  Found Microsoft.Graph version 2.10.0

[2/4] Checking existing connection...
  No existing connection

[3/4] Connecting to Microsoft Graph...
  A browser window will open for authentication.
  Please sign in with your Microsoft account.

  Successfully connected!

[4/4] Validating connection...
  Connection validated successfully!

Connection Details:
===================
  Account:     user@example.com
  Tenant ID:   xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
  Scopes:      Calendars.ReadWrite, Mail.ReadWrite, User.Read

Testing calendar access...
  Signed in as: John Doe (user@example.com)

SUCCESS: Ready to run Graph-based automation scripts!

You can now run:
  .\Show-MeetingHourSummary.ps1
  .\Test-GraphConnection.ps1
```

### Connection Duration

The authentication token is valid for **1 hour** (default). After that, you'll need to reconnect:

```powershell
.\Connect-Graph.ps1
```

---

## Usage

### Show Meeting Hour Summary

After authenticating, run the main script:

```powershell
.\Show-MeetingHourSummary.ps1
```

**What it does:**
1. Retrieves calendar events from Microsoft Graph
2. Calculates meeting hours for various time periods
3. Applies filtering rules from configuration files
4. Checks for full-hour meetings (optional reschedule prompts)
5. Displays Windows Forms popup with results and bar chart
6. Generates detailed log file (log.txt)

**Popup example:**

```
=================================
Meeting Hour Summary
Generated on: Thursday, November 14, 2025 15:30 | Microsoft Graph
=================================

Today:             3.5 hours  (4 meetings)
                   Thursday, November 14

Next Working Day:  4.0 hours  (5 meetings)
                   Friday, November 15

This Week:         15.5 hours (18 meetings)
                   Nov 11 - Nov 15

Next Week:         12.0 hours (14 meetings)
                   Nov 18 - Nov 22

[5-Day Bar Chart Visualization]

Note: 4 ignore pattern(s) applied | Powered by Microsoft Graph
```

### Test Connection

To verify your authentication is working:

```powershell
.\Test-GraphConnection.ps1
```

This runs 4 tests:
1. Authentication status
2. Required permissions/scopes
3. Calendar access
4. User profile access

**All tests should PASS** before running automation scripts.

---

## Configuration

### ignore_appointments.txt

Location: `config/ignore_appointments.txt`

Define regex patterns (one per line) for appointments to exclude from hour calculations:

```
# Regex patterns for appointments to ignore
# Lines starting with # are comments

^Lunch.*
^Break$
^Personal.*
.*\[IGNORE\].*
^Working from home.*
```

**Pattern Examples:**

| Pattern | Matches | Doesn't Match |
|---------|---------|---------------|
| `^Lunch$` | "Lunch" (exact) | "Lunch with Client" |
| `^Lunch.*` | "Lunch", "Lunch Break", "Lunch Meeting" | "Team Lunch" |
| `.*standup.*` | "Daily standup", "Team standup meeting" | "Stand-up comedy" |
| `(?i)^ooo -.*` | "OOO - Vacation" (case-insensitive) | "Out of office" |

**Case Sensitivity:**
- By default, patterns are **case-sensitive**
- Use `(?i)` prefix for case-insensitive: `(?i)^lunch$` matches "Lunch", "LUNCH", "lunch"

### meeting_change_request_template.txt

Location: `config/meeting_change_request_template.txt`

Customize the email template for reschedule requests. Supports placeholders:

- `{ORGANIZER}` - Meeting organizer name
- `{SUBJECT}` - Meeting subject
- `{START_TIME}` - Current start time (formatted)
- `{NEW_START_TIME}` - Proposed new start time (:05 past the hour)

**Default template:**

```
Subject: Request to shift meeting start time to :05

Dear {ORGANIZER},

I hope this message finds you well. I'm reaching out regarding our upcoming meeting:

Meeting: {SUBJECT}
Current Start Time: {START_TIME}

Would it be possible to shift the meeting start time by 5 minutes to {NEW_START_TIME}?
This small adjustment would help create a buffer between back-to-back meetings and
allow for better preparation time.

If this change works for you and other attendees, I would greatly appreciate it.
If the current time is critical, please feel free to keep it as scheduled.

Thank you for considering this request.

Best regards
```

### ignored_full_hour_appointments.txt

Location: `config/ignored_full_hour_appointments.txt`

**Automatically managed** - stores identifiers of meetings you've chosen "Never Ask Again" for.

You can manually edit this file to remove entries if you want to be prompted again.

Format:
```
# Ignored Full-Hour Appointments
# Each line contains an ICalUId or event Id from Microsoft Graph

# Added: 2025-11-14 15:30:00 | Subject: Team Standup | Start: 2025-11-15 09:00
040000008200E00074C5B7101A82E00800000000B0...

# Added: 2025-11-14 16:00:00 | Subject: Weekly Review | Start: 2025-11-18 10:00
040000008200E00074C5B7101A82E00800000000C1...
```

---

## COM to Graph Migration Notes

### Key API Differences

| Aspect | COM (Old) | Microsoft Graph (New) |
|--------|-----------|----------------------|
| **Connection** | `New-Object -ComObject Outlook.Application` | `Connect-MgGraph -Scopes ...` |
| **Calendar Access** | `$namespace.GetDefaultFolder(olFolderCalendar)` | `Get-MgUserCalendarView -UserId "me"` |
| **Item Retrieval** | `$calendar.Items.Restrict($filter)` | `-StartDateTime` / `-EndDateTime` parameters |
| **Date Filtering** | Outlook filter strings | ISO 8601 date format |
| **Draft Email** | `$outlook.CreateItem(olMailItem)` | `New-MgUserMessage -UserId "me"` |
| **Properties** | Direct: `$item.Subject`, `$item.Start` | Nested: `$event.Subject`, `$event.Start.DateTime` |
| **Enumerations** | `olPrivate`, `olOutOfOffice` | String values: "private", "oof" |
| **Identifiers** | `GlobalAppointmentID`, `EntryID` | `ICalUId`, `Id` |
| **Authentication** | Windows integrated (automatic) | OAuth 2.0 (explicit, browser-based) |
| **Cleanup** | `Marshal.ReleaseComObject()` | Not needed (HTTP/REST) |

### Property Mapping

**Event Properties:**

| COM Property | Graph Property | Notes |
|--------------|----------------|-------|
| `Start` | `Start.DateTime` | Graph uses nested DateTimeTimeZone object |
| `End` | `End.DateTime` | Graph uses nested DateTimeTimeZone object |
| `Subject` | `Subject` | Same |
| `AllDayEvent` | `IsAllDay` | Property name change |
| `Sensitivity = olPrivate` | `Sensitivity = "private"` | Enum → string |
| `BusyStatus = olOutOfOffice` | `ShowAs = "oof"` | Property renamed, enum → string |
| `MeetingStatus = olMeetingCanceled` | `IsCancelled` | Boolean property |
| `ResponseStatus = olResponseDeclined` | `ResponseStatus.Response = "declined"` | Nested object |
| `Organizer` | `Organizer.EmailAddress.Name` | Nested object |
| `GetOrganizer().Address` | `Organizer.EmailAddress.Address` | Direct property access |
| `GlobalAppointmentID` | `ICalUId` | Stable identifier |
| `EntryID` | `Id` | Graph's unique ID |

**Email Properties:**

| COM Property | Graph Property | Notes |
|--------------|----------------|-------|
| `To = "email"` | `ToRecipients = @(@{ EmailAddress = @{ Address = "email" }})` | Complex nested structure |
| `Subject = "text"` | `Subject = "text"` | Same |
| `Body = "text"` | `Body = @{ ContentType = "Text"; Content = "text" }` | Nested object |
| `Save()` | `New-MgUserMessage ... ` (auto-saved) | Implicitly saved |

### Behavior Differences

1. **No Local Outlook Required:**
   - COM: Requires Outlook installed and configured
   - Graph: Works anywhere (cloud-based API)

2. **Recurring Events:**
   - COM: Returns occurrences based on date filter
   - Graph: May return master event + occurrences separately (use `CalendarView` to get occurrences)

3. **Timezones:**
   - COM: Uses local timezone automatically
   - Graph: Returns UTC by default, requires conversion

4. **Performance:**
   - COM: Faster for large date ranges (local cache)
   - Graph: Network latency, but more reliable

5. **Offline Mode:**
   - COM: Works offline (cached data)
   - Graph: Requires internet connection

---

## Folder Structure

```
scripts_using_graph/
│
├── OutlookGraphAutomation.psm1      # Shared PowerShell module
│   ├── Date/time utilities
│   ├── Graph API wrappers
│   ├── Configuration loaders
│   ├── Filtering functions
│   └── Logging utilities
│
├── Connect-Graph.ps1                 # Authentication script
│   └── Interactive OAuth 2.0 login
│
├── Show-MeetingHourSummary.ps1       # Main feature script
│   ├── Meeting hour calculations
│   ├── Full-hour meeting detection
│   ├── Windows Forms UI
│   └── Draft email creation
│
├── Test-GraphConnection.ps1          # Connection validation utility
│   └── 4 diagnostic tests
│
├── README.md                         # This file
│
├── config/                           # Configuration files
│   ├── ignore_appointments.txt           # Regex patterns for exclusions
│   ├── meeting_change_request_template.txt # Email template
│   └── ignored_full_hour_appointments.txt  # Persisted ignore list
│
└── tests/                            # Unit tests (planned)
    ├── OutlookGraphAutomation.Tests.ps1
    ├── Show-MeetingHourSummary.Tests.ps1
    └── Run-Tests.ps1
```

---

## Troubleshooting

### "Not authenticated to Microsoft Graph"

**Cause:** You haven't run `Connect-Graph.ps1` or your token expired.

**Solution:**

```powershell
.\Connect-Graph.ps1
```

### "Failed to retrieve calendar events"

**Possible causes:**
1. Not authenticated (see above)
2. Missing permissions/scopes
3. Network connectivity issues
4. Calendar not accessible

**Diagnosis:**

```powershell
.\Test-GraphConnection.ps1
```

Check which test fails and follow the error message.

### "Insufficient privileges to complete the operation"

**Cause:** Your token doesn't have the required scopes.

**Solution:** Disconnect and reconnect with proper scopes:

```powershell
Disconnect-MgGraph
.\Connect-Graph.ps1
```

When prompted, **consent to all requested permissions**.

### "Module 'Microsoft.Graph' not found"

**Cause:** Microsoft.Graph module not installed.

**Solution:**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

### Browser doesn't open during authentication

**Cause:** Default browser not configured or pop-up blocker.

**Solution:**
1. Manually open the URL shown in the PowerShell window
2. Complete authentication in the browser
3. Return to PowerShell

### Events are in wrong timezone

**Cause:** Graph returns dates in UTC, conversion may be needed.

**Solution:** The script automatically converts to local time using `ConvertTo-LocalDateTime` function. If you see incorrect times, check your system's timezone setting:

```powershell
Get-TimeZone
```

### Log file shows different events than expected

**Cause:** Filtering rules or date range issues.

**Check:**
1. Review `config/ignore_appointments.txt` patterns
2. Check log.txt for filtering decisions
3. Verify your date/time expectations

**Debug:** Temporarily clear all ignore patterns and re-run to see all events.

---

## Differences from COM Implementation

### What's the Same

✓ **Functionality** - 100% feature parity with COM version
✓ **User Interface** - Identical Windows Forms popup
✓ **Configuration Files** - Same format and behavior
✓ **Filtering Logic** - Same regex patterns and exclusions
✓ **Date Calculations** - Same weekday/working day logic
✓ **Logging** - Same detailed log.txt format

### What's Different

❌ **Authentication** - Explicit OAuth 2.0 login required (no automatic Windows auth)
❌ **Local Outlook** - Not required (cloud-based API)
❌ **Performance** - Slight network latency vs. local COM
❌ **Property Access** - Nested objects instead of flat properties
❌ **Offline Mode** - Not supported (requires internet)

### What's Better

✅ **Works with New Outlook** - Future-proof solution
✅ **Cross-Platform** - Can run on macOS/Linux with PowerShell Core (future)
✅ **No COM Cleanup** - No memory management issues
✅ **Better Error Messages** - HTTP status codes provide clear diagnostics
✅ **No Admin Rights** - Never requires elevation

---

## Additional Resources

- **Microsoft Graph Documentation:** https://learn.microsoft.com/en-us/graph/
- **Microsoft Graph PowerShell SDK:** https://learn.microsoft.com/en-us/powershell/microsoftgraph/
- **Calendar API Reference:** https://learn.microsoft.com/en-us/graph/api/resources/calendar
- **Mail API Reference:** https://learn.microsoft.com/en-us/graph/api/resources/message

---

## Support & Feedback

For issues, questions, or contributions:

1. Check this README and troubleshooting section
2. Review log.txt for detailed error information
3. Run `.\Test-GraphConnection.ps1` for diagnostics
4. Open an issue on GitHub with:
   - Error message
   - Log file excerpt
   - Output of `Test-GraphConnection.ps1`

---

## Migration Roadmap

**Completed:**
- ✅ Core Graph authentication
- ✅ Calendar event retrieval
- ✅ Meeting hour calculations
- ✅ Full-hour meeting detection
- ✅ Draft email creation
- ✅ Windows Forms UI
- ✅ Configuration file support
- ✅ Comprehensive logging

**Planned:**
- ⏳ Pester unit tests for Graph functions
- ⏳ PowerShell Core (7.x) compatibility testing
- ⏳ Additional Graph-based features

---

## License

Same as parent repository.

---

**Generated for outlook_automation repository (Microsoft Graph migration)**
