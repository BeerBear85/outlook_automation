# OutlookGraphAutomation PowerShell Module
# Shared helpers and wrappers for Microsoft Graph-based Outlook automation
#
# @author: Generated for outlook_automation repository (Graph migration)
# -----------------------------------------------------------------------------

<#
.SYNOPSIS
    Shared PowerShell module for Microsoft Graph Outlook automation.

.DESCRIPTION
    This module provides helper functions and Graph API wrappers for:
    - Date/time calculations (working days, week bounds)
    - Calendar event retrieval and filtering
    - Draft email creation
    - Configuration file loading
    - Logging utilities

    Used by all scripts in scripts_using_graph folder.
#>

# -----------------------------------------------------------------------------
# Date/Time Helper Functions
# -----------------------------------------------------------------------------

function Get-WeekdayBounds {
    <#
    .SYNOPSIS
        Gets the Monday and Friday bounds for a given week.
        For weekends, returns the upcoming work week (Monday-Friday).
    .OUTPUTS
        Hashtable with Monday and Friday DateTime objects.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [DateTime]$ReferenceDate
    )

    # Get the day of week (0 = Sunday, 1 = Monday, etc.)
    $dayOfWeek = [int]$ReferenceDate.DayOfWeek

    # For weekends, use the upcoming Monday
    # For weekdays, use the Monday of the current week
    if ($dayOfWeek -eq 0) {
        # Sunday - go forward 1 day to Monday
        $daysToMonday = 1
    } elseif ($dayOfWeek -eq 6) {
        # Saturday - go forward 2 days to Monday
        $daysToMonday = 2
    } else {
        # Weekday - go back to Monday of current week
        $daysToMonday = 1 - $dayOfWeek
    }

    $monday = $ReferenceDate.Date.AddDays($daysToMonday)
    $friday = $monday.AddDays(4)

    return @{
        Monday = $monday
        Friday = $friday.AddHours(23).AddMinutes(59).AddSeconds(59)
    }
}

function Get-NextWorkingDay {
    <#
    .SYNOPSIS
        Gets the next working day (Monday-Friday) from a given reference date.
        Skips weekends.
    .OUTPUTS
        DateTime object representing the next working day.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [DateTime]$ReferenceDate
    )

    $nextDay = $ReferenceDate.Date.AddDays(1)

    # If next day is Saturday (6), move to Monday (+2 days)
    if ($nextDay.DayOfWeek -eq [DayOfWeek]::Saturday) {
        return $nextDay.AddDays(2)
    }
    # If next day is Sunday (0), move to Monday (+1 day)
    elseif ($nextDay.DayOfWeek -eq [DayOfWeek]::Sunday) {
        return $nextDay.AddDays(1)
    }
    # Otherwise, next day is already a working day
    else {
        return $nextDay
    }
}

function Get-AppointmentDuration {
    <#
    .SYNOPSIS
        Calculates the duration of an appointment in hours.
    .OUTPUTS
        Decimal hours rounded to 2 decimal places.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [DateTime]$Start,
        [Parameter(Mandatory=$true)]
        [DateTime]$End
    )

    $duration = $End - $Start
    return [Math]::Round($duration.TotalHours, 2)
}

function ConvertTo-LocalDateTime {
    <#
    .SYNOPSIS
        Converts a Graph API DateTimeTimeZone object to local DateTime.
    .DESCRIPTION
        Graph API returns dates in a structured format with DateTime and TimeZone properties.
        This function extracts and converts to local time if needed.
    .OUTPUTS
        DateTime object in local timezone.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [Object]$DateTimeTimeZone
    )

    try {
        # Graph returns DateTimeTimeZone with DateTime and TimeZone properties
        $dateTimeString = $DateTimeTimeZone.DateTime
        $timeZone = $DateTimeTimeZone.TimeZone

        # Parse the datetime string
        $dateTime = [DateTime]::Parse($dateTimeString)

        # If timezone is UTC, convert to local
        if ($timeZone -eq "UTC") {
            $dateTime = $dateTime.ToLocalTime()
        }

        return $dateTime
    }
    catch {
        Write-Error "Failed to convert DateTimeTimeZone: $_"
        return $null
    }
}

# -----------------------------------------------------------------------------
# Configuration Loading Functions
# -----------------------------------------------------------------------------

function Get-ConfigDirectory {
    <#
    .SYNOPSIS
        Gets the config directory path relative to the module.
    .OUTPUTS
        String path to config directory.
    #>
    $moduleDir = Split-Path -Parent $PSCommandPath
    return Join-Path $moduleDir "config"
}

function Load-IgnorePatterns {
    <#
    .SYNOPSIS
        Loads regex patterns from ignore_appointments.txt file.
    .OUTPUTS
        Array of regex patterns to ignore.
    #>
    param (
        [string]$ConfigDir = (Get-ConfigDirectory)
    )

    $ignoreFile = Join-Path $ConfigDir "ignore_appointments.txt"
    $patterns = @()

    if (Test-Path $ignoreFile) {
        Get-Content $ignoreFile | ForEach-Object {
            $line = $_.Trim()
            # Skip empty lines and comments
            if ($line -and -not $line.StartsWith("#")) {
                $patterns += $line
            }
        }
    }

    return $patterns
}

function Load-EmailTemplate {
    <#
    .SYNOPSIS
        Loads email template from meeting_change_request_template.txt file.
    .OUTPUTS
        String containing the template text.
    #>
    param (
        [string]$ConfigDir = (Get-ConfigDirectory)
    )

    $templateFile = Join-Path $ConfigDir "meeting_change_request_template.txt"

    if (Test-Path $templateFile) {
        return Get-Content $templateFile -Raw
    } else {
        # Return a default template if file doesn't exist
        return @"
Subject: Request to shift meeting start time to :05

Dear {ORGANIZER},

I hope this message finds you well. I'm reaching out regarding our upcoming meeting:

Meeting: {SUBJECT}
Current Start Time: {START_TIME}

Would it be possible to shift the meeting start time by 5 minutes to {NEW_START_TIME}? This small adjustment would help create a buffer between back-to-back meetings and allow for better preparation time.

If this change works for you and other attendees, I would greatly appreciate it. If the current time is critical, please feel free to keep it as scheduled.

Thank you for considering this request.

Best regards
"@
    }
}

function Load-IgnoredFullHourAppointments {
    <#
    .SYNOPSIS
        Loads appointment identifiers from ignored_full_hour_appointments.txt file.
    .OUTPUTS
        Array of appointment identifiers to ignore.
    #>
    param (
        [string]$ConfigDir = (Get-ConfigDirectory)
    )

    $ignoreFile = Join-Path $ConfigDir "ignored_full_hour_appointments.txt"
    $identifiers = @()

    if (Test-Path $ignoreFile) {
        Get-Content $ignoreFile | ForEach-Object {
            $line = $_.Trim()
            # Skip empty lines and comments
            if ($line -and -not $line.StartsWith("#")) {
                $identifiers += $line
            }
        }
    }

    return $identifiers
}

function Save-IgnoredFullHourAppointment {
    <#
    .SYNOPSIS
        Adds an appointment identifier to the ignored_full_hour_appointments.txt file.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$Identifier,
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        [Parameter(Mandatory=$true)]
        [DateTime]$StartTime,
        [string]$ConfigDir = (Get-ConfigDirectory)
    )

    $ignoreFile = Join-Path $ConfigDir "ignored_full_hour_appointments.txt"

    # Create file with header if it doesn't exist
    if (-not (Test-Path $ignoreFile)) {
        $header = @"
# Ignored Full-Hour Appointments
# This file contains appointment identifiers that should not trigger the reschedule popup.
# Each line contains an ICalUId or event Id from Microsoft Graph.
# Lines starting with # are comments and will be ignored.
# You can manually edit this file to add or remove entries.
#
"@
        $header | Out-File -FilePath $ignoreFile -Encoding utf8
    }

    # Add comment with meeting details and identifier
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $comment = "# Added: $timestamp | Subject: $Subject | Start: $($StartTime.ToString('yyyy-MM-dd HH:mm'))"
    $comment | Out-File -FilePath $ignoreFile -Append -Encoding utf8
    $Identifier | Out-File -FilePath $ignoreFile -Append -Encoding utf8
}

# -----------------------------------------------------------------------------
# Filtering Functions
# -----------------------------------------------------------------------------

function Test-ShouldIgnoreAppointment {
    <#
    .SYNOPSIS
        Tests if an appointment should be ignored based on regex patterns.
    .OUTPUTS
        Boolean indicating whether to ignore the appointment.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        [Parameter(Mandatory=$true)]
        [string[]]$IgnorePatterns
    )

    foreach ($pattern in $IgnorePatterns) {
        # Use -cmatch for case-sensitive matching by default
        # Users can use (?i) in their patterns for case-insensitive matching
        if ($Subject -cmatch $pattern) {
            return $true
        }
    }

    return $false
}

# -----------------------------------------------------------------------------
# Logging Functions
# -----------------------------------------------------------------------------

function Initialize-LogFile {
    <#
    .SYNOPSIS
        Creates or clears the log file and writes a header.
    .OUTPUTS
        String path to the log file.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$LogDirectory
    )

    $logFile = Join-Path $LogDirectory "log.txt"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $header = @"
================================================================================
Meeting Hour Summary Script Log (Microsoft Graph Version)
Generated: $timestamp
================================================================================

"@

    $header | Out-File -FilePath $logFile -Encoding utf8
    return $logFile
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a message to the log file with error handling and console fallback.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$LogFile,
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    try {
        # Try to write to log file
        $logEntry | Out-File -FilePath $LogFile -Append -Encoding utf8 -ErrorAction Stop
    }
    catch {
        # If file logging fails, write to console as fallback
        Write-Host "[LOG ERROR] Failed to write to log file: $_" -ForegroundColor Yellow
        Write-Host $logEntry -ForegroundColor Cyan
    }
}

# -----------------------------------------------------------------------------
# Microsoft Graph API Functions
# -----------------------------------------------------------------------------

function Test-GraphConnection {
    <#
    .SYNOPSIS
        Checks if user is authenticated to Microsoft Graph.
    .OUTPUTS
        Boolean indicating if authenticated.
    #>
    try {
        $context = Get-MgContext
        if ($null -eq $context) {
            return $false
        }
        return $true
    }
    catch {
        return $false
    }
}

function Get-GraphCalendarEvents {
    <#
    .SYNOPSIS
        Retrieves calendar events from Microsoft Graph for a date range.
    .DESCRIPTION
        Wrapper around Get-MgUserCalendarView with error handling and logging.
    .OUTPUTS
        Array of calendar event objects.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [DateTime]$StartDate,
        [Parameter(Mandatory=$true)]
        [DateTime]$EndDate,
        [string]$LogFile = $null
    )

    if (-not (Test-GraphConnection)) {
        throw "Not authenticated to Microsoft Graph. Please run Connect-Graph.ps1 first."
    }

    try {
        # Format dates for Graph API (ISO 8601)
        $startDateTime = $StartDate.ToString("yyyy-MM-ddTHH:mm:ss")
        $endDateTime = $EndDate.ToString("yyyy-MM-ddTHH:mm:ss")

        if ($LogFile) {
            Write-Log -LogFile $LogFile -Message "Retrieving calendar events from $startDateTime to $endDateTime"
        }

        # Get calendar view with all events in the date range
        $events = Get-MgUserCalendarView -UserId "me" `
            -StartDateTime $startDateTime `
            -EndDateTime $endDateTime `
            -All `
            -ErrorAction Stop

        if ($LogFile) {
            Write-Log -LogFile $LogFile -Message "Retrieved $($events.Count) calendar events from Microsoft Graph"
        }

        return $events
    }
    catch {
        $errorMessage = "Failed to retrieve calendar events: $_"
        if ($LogFile) {
            Write-Log -LogFile $LogFile -Message $errorMessage -Level "ERROR"
        }
        throw $errorMessage
    }
}

function New-GraphDraftEmail {
    <#
    .SYNOPSIS
        Creates a draft email in Microsoft Graph (Outlook).
    .DESCRIPTION
        Creates a draft email with the specified recipient, subject, and body.
        The email is saved to the Drafts folder but not sent.
    .OUTPUTS
        Boolean indicating success.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$ToAddress,
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        [Parameter(Mandatory=$true)]
        [string]$Body,
        [string]$LogFile = $null
    )

    if (-not (Test-GraphConnection)) {
        throw "Not authenticated to Microsoft Graph. Please run Connect-Graph.ps1 first."
    }

    try {
        if ($LogFile) {
            Write-Log -LogFile $LogFile -Message "Creating draft email to: $ToAddress"
        }

        # Create message body structure
        $messageBody = @{
            Subject = $Subject
            Body = @{
                ContentType = "Text"
                Content = $Body
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = $ToAddress
                    }
                }
            )
        }

        # Create draft (Save parameter means don't send)
        $draft = New-MgUserMessage -UserId "me" -BodyParameter $messageBody -ErrorAction Stop

        if ($LogFile) {
            Write-Log -LogFile $LogFile -Message "Draft email created successfully (ID: $($draft.Id))"
        }

        return $true
    }
    catch {
        $errorMessage = "Failed to create draft email: $_"
        if ($LogFile) {
            Write-Log -LogFile $LogFile -Message $errorMessage -Level "ERROR"
        }
        Write-Error $errorMessage
        return $false
    }
}

function Get-EventIdentifier {
    <#
    .SYNOPSIS
        Gets a stable identifier for a Graph calendar event.
        Uses ICalUId (equivalent to GlobalAppointmentID) or falls back to Id.
    .OUTPUTS
        String containing the event identifier.
    #>
    param (
        [Parameter(Mandatory=$true)]
        [Object]$Event
    )

    # Try ICalUId first (stable across updates, equivalent to GlobalAppointmentID)
    if ($Event.ICalUId) {
        return $Event.ICalUId
    }

    # Fall back to Graph Id
    if ($Event.Id) {
        return $Event.Id
    }

    # If both fail, return empty string
    return ""
}

# -----------------------------------------------------------------------------
# Export Module Members
# -----------------------------------------------------------------------------

Export-ModuleMember -Function @(
    'Get-WeekdayBounds',
    'Get-NextWorkingDay',
    'Get-AppointmentDuration',
    'ConvertTo-LocalDateTime',
    'Load-IgnorePatterns',
    'Load-EmailTemplate',
    'Load-IgnoredFullHourAppointments',
    'Save-IgnoredFullHourAppointment',
    'Test-ShouldIgnoreAppointment',
    'Initialize-LogFile',
    'Write-Log',
    'Test-GraphConnection',
    'Get-GraphCalendarEvents',
    'New-GraphDraftEmail',
    'Get-EventIdentifier'
)
