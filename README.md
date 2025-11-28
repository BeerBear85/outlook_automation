# Outlook Meeting Hour Summary

Automated tools to help you manage your meeting schedule by displaying meeting hour summaries with visual charts and suggesting optimal meeting times.

---

## Overview

This repository contains **three parallel implementations** of the same core functionality, each optimized for different use cases and platforms.

All implementations provide:
- **Meeting Hour Summaries** - View total hours for Today, Next Working Day, This Week, and Next Week
- **5-Day Visual Bar Chart** - Color-coded overview (Green ≤3h, Yellow ≤4h, Red >4h)
- **Full-Hour Meeting Optimizer** - Detect meetings at :00 and suggest shifting to :05
- **Draft Email Generation** - Automated polite requests to meeting organizers
- **Customizable Filtering** - Regex-based patterns to exclude specific appointments

---

## Choose Your Implementation

### 1. Python + Microsoft Graph (Newest) 🐍

**📁 Location:** [`scripts_using_python/`](scripts_using_python/)

**Best for:** Cross-platform use, Python developers, modern cloud-based solution

**Features:**
- ✅ Works with **new Outlook** (Outlook for Windows)
- ✅ Works with classic Outlook
- ✅ **Cross-platform** (Windows, macOS, Linux)
- ✅ No Outlook installation required (cloud-based)
- ✅ Modern Python implementation with **62 automated tests**
- ✅ MSAL authentication with device code flow
- ✅ ~85% test coverage

**Quick Start:**
```bash
cd scripts_using_python
python -m venv venv
venv\Scripts\activate  # Windows: venv\Scripts\activate | macOS/Linux: source venv/bin/activate
pip install -r requirements.txt
python scripts/connect_graph.py
python scripts/show_meeting_summary.py
```

**📖 [Full Python Documentation →](scripts_using_python/README.md)**

---

### 2. PowerShell + Microsoft Graph (Recommended for Windows) 💻

**📁 Location:** [`scripts_using_graph/`](scripts_using_graph/)

**Best for:** Windows users, PowerShell enthusiasts, modern cloud-based solution

**Features:**
- ✅ Works with **new Outlook** (Outlook for Windows)
- ✅ Works with classic Outlook
- ✅ No Outlook installation required (cloud-based)
- ✅ Native PowerShell experience
- ✅ Microsoft.Graph PowerShell SDK
- ✅ Future-proof solution
- ✅ Interactive browser authentication

**Quick Start:**
```powershell
cd scripts_using_graph
.\Connect-Graph.ps1
.\Show-MeetingHourSummary.ps1
```

**📖 [Full PowerShell Graph Documentation →](scripts_using_graph/README.md)**

---

### 3. COM-based PowerShell (Legacy) 🕰️

**📁 Location:** [`scripts_using_com/`](scripts_using_com/)

**Best for:** Classic Outlook users, offline usage, fastest performance

**Features:**
- ✅ Works with **classic Outlook only**
- ✅ Faster (uses local Outlook cache)
- ✅ Works offline
- ✅ No authentication required
- ⚠️ **Does NOT work with new Outlook**
- ⚠️ Windows-only
- ⚠️ Requires Outlook installed

**Quick Start:**
```powershell
cd scripts_using_com
.\Show-MeetingHourSummary.ps1
```

**📖 [Full COM Documentation →](scripts_using_com/README.md)**

---

## Comparison Table

| Feature | Python Graph | PowerShell Graph | COM (Legacy) |
|---------|-------------|------------------|--------------|
| **New Outlook Support** | ✅ Yes | ✅ Yes | ❌ No |
| **Classic Outlook Support** | ✅ Yes | ✅ Yes | ✅ Yes |
| **Platform** | Windows, macOS, Linux | Windows, macOS, Linux | Windows only |
| **Requires Outlook Installed** | ❌ No | ❌ No | ✅ Yes |
| **Authentication** | Device code (MSAL) | Browser (MSAL) | None (uses logged-in user) |
| **Offline Support** | ❌ No | ❌ No | ✅ Yes |
| **Performance** | Cloud | Cloud | Fast (local) |
| **Future-Proof** | ✅ Yes | ✅ Yes | ⚠️ Limited |
| **Automated Tests** | 62 tests | None | Pester tests |
| **Language** | Python 3.8+ | PowerShell 5.1+ | PowerShell 5.1+ |

---

## Repository Structure

```
outlook_automation/
├── README.md                      # This file (overview)
│
├── scripts_using_python/          # Python + Microsoft Graph (newest)
│   ├── README.md                  # Python documentation
│   ├── src/outlook_graph/         # Python package
│   ├── scripts/                   # Entry point scripts
│   ├── tests/                     # 62 automated tests
│   └── config/                    # Configuration files
│
├── scripts_using_graph/           # PowerShell + Microsoft Graph
│   ├── README.md                  # PowerShell Graph documentation
│   ├── Connect-Graph.ps1
│   ├── Show-MeetingHourSummary.ps1
│   ├── OutlookGraphAutomation.psm1
│   └── config/                    # Configuration files
│
└── scripts_using_com/             # COM-based PowerShell (legacy)
    ├── README.md                  # COM documentation
    ├── Show-MeetingHourSummary.ps1
    ├── Show-MeetingHourSummary.Tests.ps1
    └── ...                        # Configuration files
```

---

## Which Implementation Should I Use?

### Use **Python + Microsoft Graph** if:
- ✅ You prefer Python over PowerShell
- ✅ You need cross-platform support (macOS, Linux)
- ✅ You want the most modern implementation with comprehensive tests
- ✅ You're building automation pipelines
- ✅ You have new Outlook (or plan to migrate)

### Use **PowerShell + Microsoft Graph** if:
- ✅ You prefer PowerShell
- ✅ You have new Outlook (or plan to migrate)
- ✅ You want a cloud-based solution
- ✅ You don't need offline support
- ✅ You want a future-proof solution

### Use **COM-based PowerShell** if:
- ✅ You have **classic Outlook** and don't plan to migrate
- ✅ You need offline support
- ✅ You want the fastest performance (local cache)
- ✅ You're on Windows only
- ⚠️ You understand this is legacy and won't work with new Outlook

---

## Migration Path

### From COM to Graph (Recommended)

If you're currently using the COM implementation and want to migrate:

1. **Choose your preferred Graph implementation:**
   - [Python](scripts_using_python/README.md) - Cross-platform
   - [PowerShell](scripts_using_graph/README.md) - Windows native

2. **Copy your configuration:**
   - `ignore_appointments.txt` → new implementation's `config/` folder
   - `meeting_change_request_template.txt` → new implementation's `config/` folder
   - `ignored_full_hour_appointments.txt` → new implementation's `config/` folder

3. **Authenticate once:**
   - Python: Run `python scripts/connect_graph.py`
   - PowerShell: Run `.\Connect-Graph.ps1`

4. **Run the new implementation:**
   - Python: `python scripts/show_meeting_summary.py`
   - PowerShell: `.\Show-MeetingHourSummary.ps1`

---

## Features (All Implementations)

### Meeting Hour Summary
- Calculate total meeting hours for customizable time periods
- Visual 5-day bar chart with color coding:
  - **Green** (0-3 hours) - Healthy meeting load
  - **Yellow** (3-4 hours) - Moderate meeting load
  - **Red** (4+ hours) - Heavy meeting load
- Automatically skips weekends
- Meeting counts for each period

### Full-Hour Meeting Optimizer
- Detects meetings starting exactly at :00 (10:00, 11:00, etc.)
- Suggests shifting to :05 for better work-life balance
- Creates draft emails (never sent automatically)
- Customizable email templates
- "Never Ask Again" option to permanently ignore specific meetings

### Intelligent Filtering
- Regex-based appointment filtering
- Excludes all-day events automatically
- Excludes cancelled and declined meetings
- Excludes private appointments and Out of Office
- Customizable ignore patterns

---

## Getting Started

1. **Choose your implementation** (see comparison above)
2. **Navigate to the implementation folder**
3. **Read the implementation-specific README**
4. **Follow the installation and setup instructions**
5. **Run the scripts**

---

## Documentation Links

- **[Python + Microsoft Graph README](scripts_using_python/README.md)** - Complete Python documentation
- **[PowerShell + Microsoft Graph README](scripts_using_graph/README.md)** - Complete PowerShell Graph documentation
- **[COM-based PowerShell README](scripts_using_com/README.md)** - Complete COM documentation

---

## Testing

### Python Implementation
- **62 automated tests** (unit, integration, regression)
- **~85% code coverage**
- Fast execution (< 1 second)
- Run: `python run_tests.py`

### COM Implementation
- Comprehensive Pester tests
- Tests for date calculations, filtering, edge cases
- Run: `.\Run-Tests.ps1`

### PowerShell Graph Implementation
- Manual testing recommended
- Use `Test-GraphConnection.ps1` to verify setup

---

## Contributing

When contributing to this repository:

1. Choose the implementation you want to modify
2. Read the implementation-specific README
3. Make your changes
4. Test thoroughly:
   - Python: Run `python run_tests.py`
   - COM: Run `.\Run-Tests.ps1`
   - PowerShell Graph: Test manually
5. Update the relevant README
6. Submit a pull request

---

## Support

For implementation-specific questions:
- **Python**: See [scripts_using_python/README.md](scripts_using_python/README.md)
- **PowerShell Graph**: See [scripts_using_graph/README.md](scripts_using_graph/README.md)
- **COM**: See [scripts_using_com/README.md](scripts_using_com/README.md)

---

## License

Generated for outlook_automation repository

---

## Quick Links

- 🐍 [Python Implementation →](scripts_using_python/)
- 💻 [PowerShell Graph Implementation →](scripts_using_graph/)
- 🕰️ [COM Implementation (Legacy) →](scripts_using_com/)
