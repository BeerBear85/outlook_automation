# Meeting Hour Summary Script
# Displays daily and weekly meeting hour summaries via Windows Forms popup
# Reads Outlook calendar via COM objects and filters appointments based on regex patterns
#
# @author: Generated for outlook_automation repository
# -----------------------------------------------------------------------------

<#
.SYNOPSIS
    Shows a popup with meeting hour summaries for Today, Tomorrow, This Week, and Next Week.

.DESCRIPTION
    This script reads your Outlook calendar and calculates total meeting hours for:
    - Today
    - Tomorrow
    - This week (Monday to Friday)
    - Next week (Monday to Friday)

    Appointments matching regex patterns in ignore_appointments.txt will be excluded.
    Results are displayed in a Windows Forms popup window.

.EXAMPLE
    .\Show-MeetingHourSummary.ps1
    Displays the meeting hour summary popup.
#>

# -----------------------------------------------------------------------------
# Helper Functions
# -----------------------------------------------------------------------------

function Get-ScriptDirectory {
    <#
    .SYNOPSIS
        Gets the directory where this script is located.
    #>
    return $PSScriptRoot
}

function Load-IgnorePatterns {
    <#
    .SYNOPSIS
        Loads regex patterns from ignore_appointments.txt file.
    .OUTPUTS
        Array of regex patterns to ignore.
    #>
    param (
        [string]$ScriptDir
    )

    $ignoreFile = Join-Path $ScriptDir "ignore_appointments.txt"
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

function Test-ShouldIgnoreAppointment {
    <#
    .SYNOPSIS
        Tests if an appointment should be ignored based on regex patterns.
    #>
    param (
        [string]$Subject,
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

function Load-EmailTemplate {
    <#
    .SYNOPSIS
        Loads email template from meeting_change_request_template.txt file.
    .OUTPUTS
        String containing the template text.
    #>
    param (
        [string]$ScriptDir
    )

    $templateFile = Join-Path $ScriptDir "meeting_change_request_template.txt"

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

function Get-WeekdayBounds {
    <#
    .SYNOPSIS
        Gets the Monday and Friday bounds for a given week.
        For weekends, returns the upcoming work week (Monday-Friday).
    #>
    param (
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
    #>
    param (
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
    #>
    param (
        [DateTime]$Start,
        [DateTime]$End
    )

    $duration = $End - $Start
    return [Math]::Round($duration.TotalHours, 2)
}

function Get-FullHourMeetings {
    <#
    .SYNOPSIS
        Finds meetings starting exactly on the full hour in the next 14 days.
        Excludes all-day events, private items, and Out of Office entries.
        Returns at most 10 meetings, starting with the earliest.
    #>
    param (
        [Object]$Items,
        [DateTime]$StartDate,
        [DateTime]$EndDate,
        [string[]]$IgnorePatterns,
        [int]$MaxCount = 10,
        [string]$LogFile = $null
    )

    if ($LogFile) {
        Write-Log -LogFile $LogFile -Message "Scanning for full-hour meetings (starting at :00) in the next 14 days"
    }

    $fullHourMeetings = @()
    $now = Get-Date
    $skippedReasons = @{
        'Ignored' = 0
        'AllDay' = 0
        'Private' = 0
        'OutOfOffice' = 0
        'AlreadyStarted' = 0
        'NotFullHour' = 0
        'Cancelled' = 0
        'Declined' = 0
    }

    foreach ($item in $Items) {
        # Allow both real Outlook AppointmentItems and test mock objects (PSCustomObject)
        if ($item -is [Microsoft.Office.Interop.Outlook.AppointmentItem] -or $item -is [PSCustomObject]) {
            $appointmentStart = $item.Start

            # Check if appointment is in the next 14 days
            if ($appointmentStart -ge $StartDate -and $appointmentStart -lt $EndDate) {
                # Skip if matches ignore pattern
                if (Test-ShouldIgnoreAppointment -Subject $item.Subject -IgnorePatterns $IgnorePatterns) {
                    $skippedReasons['Ignored']++
                    continue
                }

                # Skip all-day events
                if ($item.AllDayEvent) {
                    $skippedReasons['AllDay']++
                    continue
                }

                # Skip private items
                if ($item.Sensitivity -eq [Microsoft.Office.Interop.Outlook.OlSensitivity]::olPrivate) {
                    $skippedReasons['Private']++
                    continue
                }

                # Skip Out of Office (BusyStatus = olOutOfOffice)
                if ($item.BusyStatus -eq [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olOutOfOffice) {
                    $skippedReasons['OutOfOffice']++
                    continue
                }

                # Skip meetings that have already started
                if ($appointmentStart -lt $now) {
                    $skippedReasons['AlreadyStarted']++
                    continue
                }

                # Skip cancelled meetings (safe property check for both COM objects and test mocks)
                try {
                    if ($item.MeetingStatus -eq [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeetingCanceled) {
                        $skippedReasons['Cancelled']++
                        continue
                    }
                } catch {
                    # Property doesn't exist (test mock), continue processing
                }

                # Skip declined meetings (safe property check for both COM objects and test mocks)
                try {
                    if ($item.ResponseStatus -eq [Microsoft.Office.Interop.Outlook.OlResponseStatus]::olResponseDeclined) {
                        $skippedReasons['Declined']++
                        continue
                    }
                } catch {
                    # Property doesn't exist (test mock), continue processing
                }

                # Check if start time is exactly on the hour (minute = 0, second = 0)
                if ($appointmentStart.Minute -eq 0 -and $appointmentStart.Second -eq 0) {
                    $fullHourMeetings += $item
                    if ($LogFile) {
                        Write-Log -LogFile $LogFile -Message "  Found full-hour meeting: '$($item.Subject)' | Start: $($appointmentStart.ToString('yyyy-MM-dd HH:mm')) | Organizer: $($item.Organizer)"
                    }
                } else {
                    $skippedReasons['NotFullHour']++
                }
            }
        }
    }

    # Sort by start time and take the first $MaxCount
    $fullHourMeetings = $fullHourMeetings | Sort-Object Start | Select-Object -First $MaxCount

    if ($LogFile) {
        Write-Log -LogFile $LogFile -Message "Full-hour meeting scan complete: Found $($fullHourMeetings.Count) meetings"
        Write-Log -LogFile $LogFile -Message "  Skipped: $($skippedReasons['Ignored']) (ignored pattern), $($skippedReasons['AllDay']) (all-day), $($skippedReasons['Private']) (private), $($skippedReasons['OutOfOffice']) (OOO), $($skippedReasons['AlreadyStarted']) (already started), $($skippedReasons['Cancelled']) (cancelled), $($skippedReasons['Declined']) (declined)"
        Write-Log -LogFile $LogFile -Message ""
    }

    return $fullHourMeetings
}

function Show-MeetingRescheduleDialog {
    <#
    .SYNOPSIS
        Shows a popup dialog asking if user wants to draft a reschedule email.
    .OUTPUTS
        Boolean indicating if user confirmed (True) or declined (False).
    #>
    param (
        [string]$Subject,
        [DateTime]$StartTime,
        [string]$Organizer
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Create form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Reschedule Meeting?"
    $form.Size = New-Object System.Drawing.Size(500, 280)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    # Title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 25)
    $titleLabel.Text = "Meeting starts at full hour"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($titleLabel)

    # Message label
    $messageLabel = New-Object System.Windows.Forms.Label
    $messageLabel.Location = New-Object System.Drawing.Point(20, 55)
    $messageLabel.Size = New-Object System.Drawing.Size(460, 60)
    $messageLabel.Text = "The following meeting starts exactly on the hour. Would you like to draft an email requesting it be moved to :05?"
    $messageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($messageLabel)

    # Details label
    $detailsLabel = New-Object System.Windows.Forms.Label
    $detailsLabel.Location = New-Object System.Drawing.Point(20, 120)
    $detailsLabel.Size = New-Object System.Drawing.Size(460, 80)
    $detailsLabel.Text = "Subject: $Subject`r`nStart Time: $($StartTime.ToString('dddd, MMMM dd, yyyy HH:mm'))`r`nOrganizer: $Organizer"
    $detailsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $detailsLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Controls.Add($detailsLabel)

    # Yes button
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(150, 210)
    $yesButton.Size = New-Object System.Drawing.Size(90, 30)
    $yesButton.Text = "Yes"
    $yesButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.Controls.Add($yesButton)

    # No button
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(260, 210)
    $noButton.Size = New-Object System.Drawing.Size(90, 30)
    $noButton.Text = "No"
    $noButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $form.Controls.Add($noButton)

    $form.AcceptButton = $yesButton
    $form.CancelButton = $noButton

    # Show the form and return result
    $result = $form.ShowDialog()
    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

function New-RescheduleDraftEmail {
    <#
    .SYNOPSIS
        Creates a draft email in Outlook requesting meeting reschedule.
    #>
    param (
        [Object]$Outlook,
        [Object]$AppointmentItem,
        [string]$Template
    )

    try {
        # Get organizer email address
        $organizerEmail = $AppointmentItem.Organizer
        if ($AppointmentItem.GetOrganizer()) {
            $organizerRecipient = $AppointmentItem.GetOrganizer()
            $organizerEmail = $organizerRecipient.Address
        }

        # Get organizer name
        $organizerName = $AppointmentItem.Organizer

        # Calculate new start time (add 5 minutes)
        $currentStart = $AppointmentItem.Start
        $newStart = $currentStart.AddMinutes(5)

        # Replace placeholders in template
        $emailContent = $Template -replace '\{ORGANIZER\}', $organizerName `
                                   -replace '\{SUBJECT\}', $AppointmentItem.Subject `
                                   -replace '\{START_TIME\}', $currentStart.ToString('dddd, MMMM dd, yyyy HH:mm') `
                                   -replace '\{NEW_START_TIME\}', $newStart.ToString('HH:mm')

        # Extract subject line from template (first line after "Subject:")
        $subjectLine = "Request to shift meeting start time to :05"
        if ($emailContent -match 'Subject:\s*(.+)') {
            $subjectLine = $matches[1].Trim()
            # Remove the subject line from the body
            $emailContent = $emailContent -replace 'Subject:\s*.+\r?\n\r?\n?', ''
        }

        # Create draft email
        $mail = $Outlook.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olMailItem)
        $mail.To = $organizerEmail
        $mail.Subject = $subjectLine
        $mail.Body = $emailContent.Trim()

        # Save as draft (do not send)
        $mail.Save()

        return $true
    }
    catch {
        Write-Error "Failed to create draft email: $_"
        return $false
    }
}

function Initialize-LogFile {
    <#
    .SYNOPSIS
        Creates or clears the log file and writes a header.
    #>
    param (
        [string]$ScriptDir
    )

    $logFile = Join-Path $ScriptDir "log.txt"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $header = @"
================================================================================
Meeting Hour Summary Script Log
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
        [string]$LogFile,
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

function Get-MeetingHours {
    <#
    .SYNOPSIS
        Calculates total meeting hours for a time period.
    #>
    param (
        [Object]$Items,
        [DateTime]$StartDate,
        [DateTime]$EndDate,
        [string[]]$IgnorePatterns,
        [string]$LogFile = $null,
        [string]$PeriodName = ""
    )

    if ($LogFile) {
        Write-Log -LogFile $LogFile -Message "Processing period: $PeriodName ($($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd')))"
    }

    $totalHours = 0
    $appointmentCount = 0
    $ignoredCount = 0
    $allDayCount = 0

    foreach ($item in $Items) {
        # Allow both real Outlook AppointmentItems and test mock objects (PSCustomObject)
        if ($item -is [Microsoft.Office.Interop.Outlook.AppointmentItem] -or $item -is [PSCustomObject]) {
            $appointmentStart = $item.Start
            $appointmentEnd = $item.End

            # Check if appointment falls within the time period
            if ($appointmentStart -ge $StartDate -and $appointmentStart -lt $EndDate) {
                $duration = Get-AppointmentDuration -Start $appointmentStart -End $appointmentEnd

                if ($LogFile) {
                    Write-Log -LogFile $LogFile -Message "  Found appointment: '$($item.Subject)' | Start: $($appointmentStart.ToString('yyyy-MM-dd HH:mm')) | Duration: $duration hours"

                    # Log detailed properties for debugging
                    try {
                        $meetingStatus = if ($item.MeetingStatus) { $item.MeetingStatus } else { "N/A" }
                        $responseStatus = if ($item.ResponseStatus) { $item.ResponseStatus } else { "N/A" }
                        $isAllDay = $item.AllDayEvent
                        Write-Log -LogFile $LogFile -Message "    Properties: MeetingStatus=$meetingStatus, ResponseStatus=$responseStatus, AllDayEvent=$isAllDay"
                    }
                    catch {
                        Write-Log -LogFile $LogFile -Message "    [Could not read some properties]"
                    }
                }

                # Skip if matches ignore pattern
                if (Test-ShouldIgnoreAppointment -Subject $item.Subject -IgnorePatterns $IgnorePatterns) {
                    if ($LogFile) {
                        Write-Log -LogFile $LogFile -Message "    -> EXCLUDED: Matches ignore pattern" -Level "SKIP"
                    }
                    $ignoredCount++
                    continue
                }

                # Skip all-day events (typically not meetings)
                if ($item.AllDayEvent) {
                    if ($LogFile) {
                        Write-Log -LogFile $LogFile -Message "    -> EXCLUDED: All-day event" -Level "SKIP"
                    }
                    $allDayCount++
                    continue
                }

                # Skip cancelled meetings (safe property check for both COM objects and test mocks)
                try {
                    if ($item.MeetingStatus -eq [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeetingCanceled) {
                        if ($LogFile) {
                            Write-Log -LogFile $LogFile -Message "    -> EXCLUDED: Meeting cancelled" -Level "SKIP"
                        }
                        $ignoredCount++
                        continue
                    }
                } catch {
                    # Property doesn't exist (test mock), continue processing
                }

                # Skip declined meetings (safe property check for both COM objects and test mocks)
                try {
                    if ($item.ResponseStatus -eq [Microsoft.Office.Interop.Outlook.OlResponseStatus]::olResponseDeclined) {
                        if ($LogFile) {
                            Write-Log -LogFile $LogFile -Message "    -> EXCLUDED: Meeting declined" -Level "SKIP"
                        }
                        $ignoredCount++
                        continue
                    }
                } catch {
                    # Property doesn't exist (test mock), continue processing
                }

                $hours = $duration
                $totalHours += $hours
                $appointmentCount++
                if ($LogFile) {
                    Write-Log -LogFile $LogFile -Message "    -> INCLUDED in time estimate" -Level "INCL"
                }
            }
        }
    }

    if ($LogFile) {
        $cancelledCount = 0
        $declinedCount = 0
        # Count cancelled and declined from the total excluded
        $totalExcluded = $ignoredCount + $allDayCount
        Write-Log -LogFile $LogFile -Message "Summary for $PeriodName - Total: $([Math]::Round($totalHours, 2)) hours | Included: $appointmentCount appointments | Excluded: $totalExcluded"
        Write-Log -LogFile $LogFile -Message ""
    }

    return @{
        Hours = [Math]::Round($totalHours, 2)
        Count = $appointmentCount
    }
}

# -----------------------------------------------------------------------------
# Main Script Logic
# -----------------------------------------------------------------------------

# Check if running as admin (not supported for Outlook COM access)
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    [System.Windows.Forms.MessageBox]::Show(
        "This script cannot be run with administrative privileges due to lack of Outlook access.",
        "Administrator Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    Exit 1
}

# Initialize logging
$scriptDir = Get-ScriptDirectory
$logFile = Initialize-LogFile -ScriptDir $scriptDir

Write-Host "Meeting Hour Summary Script - Logging to: $logFile" -ForegroundColor Green
Write-Host ""

Write-Log -LogFile $logFile -Message "Script execution started"
Write-Log -LogFile $logFile -Message "Script directory: $scriptDir"

# Load ignore patterns
$ignorePatterns = Load-IgnorePatterns -ScriptDir $scriptDir
Write-Log -LogFile $logFile -Message "Loaded $($ignorePatterns.Count) ignore patterns from ignore_appointments.txt"
if ($ignorePatterns.Count -gt 0) {
    foreach ($pattern in $ignorePatterns) {
        Write-Log -LogFile $logFile -Message "  - Pattern: $pattern"
    }
}
Write-Log -LogFile $logFile -Message ""

# Load Outlook COM object
Write-Log -LogFile $logFile -Message "Connecting to Outlook..."
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
    Write-Log -LogFile $logFile -Message "Successfully connected to Outlook calendar"
    Write-Log -LogFile $logFile -Message ""
} catch {
    Write-Log -LogFile $logFile -Message "ERROR: Failed to connect to Outlook - $_" -Level "ERROR"
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to connect to Outlook. Please ensure Outlook is installed and configured.`n`nError: $_",
        "Outlook Connection Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    Exit 1
}

# Define time periods
$now = Get-Date
$today = $now.Date
$nextWorkingDay = Get-NextWorkingDay -ReferenceDate $now
$dayAfterNextWorkingDay = Get-NextWorkingDay -ReferenceDate $nextWorkingDay

# Get this week's Monday-Friday bounds
$thisWeekBounds = Get-WeekdayBounds -ReferenceDate $now

# Get next week's Monday-Friday bounds
$nextWeekMonday = $thisWeekBounds.Monday.AddDays(7)
$nextWeekBounds = Get-WeekdayBounds -ReferenceDate $nextWeekMonday

# Fetch calendar items (from today to end of next week or 14 days, whichever is later)
$fetchStartDate = $today
$fourteenDaysLater = $today.AddDays(14)
$fetchEndDate = if ($nextWeekBounds.Friday -gt $fourteenDaysLater) { $nextWeekBounds.Friday } else { $fourteenDaysLater }
$filter = "[Start] >= '" + $fetchStartDate.ToString("g") + "' AND [Start] < '" + $fetchEndDate.AddDays(1).ToString("g") + "'"

try {
    $items = $calendar.Items.Restrict($filter)
    $itemCount = 0
    $allItems = @()
    foreach ($item in $items) {
        if ($item -is [Microsoft.Office.Interop.Outlook.AppointmentItem] -or $item -is [PSCustomObject]) {
            $allItems += [PSCustomObject]@{
                Subject = $item.Subject
                Start = $item.Start
                End = $item.End
            }
            $itemCount++
        }
    }
    Write-Log -LogFile $logFile -Message "Successfully retrieved calendar items from $($fetchStartDate.ToString('yyyy-MM-dd')) to $($fetchEndDate.ToString('yyyy-MM-dd'))"
    Write-Log -LogFile $logFile -Message "Total calendar items found: $itemCount"
    Write-Log -LogFile $logFile -Message ""
    Write-Log -LogFile $logFile -Message "=========================================  "
    Write-Log -LogFile $logFile -Message "ALL CALENDAR ITEMS RETRIEVED FROM OUTLOOK"
    Write-Log -LogFile $logFile -Message "========================================="
    Write-Log -LogFile $logFile -Message ""
    foreach ($item in ($allItems | Sort-Object Start)) {
        Write-Log -LogFile $logFile -Message "  $($item.Start.ToString('yyyy-MM-dd HH:mm')) | $($item.Subject)"
    }
    Write-Log -LogFile $logFile -Message ""
} catch {
    Write-Log -LogFile $logFile -Message "ERROR: Failed to retrieve calendar items - $_" -Level "ERROR"
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to retrieve calendar items.`n`nError: $_",
        "Calendar Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    Exit 1
}

# Calculate meeting hours for each period
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message "CALCULATING MEETING HOURS"
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message ""

$todayHours = Get-MeetingHours -Items $items -StartDate $today -EndDate $nextWorkingDay -IgnorePatterns $ignorePatterns -LogFile $logFile -PeriodName "Today"
$nextWorkingDayHours = Get-MeetingHours -Items $items -StartDate $nextWorkingDay -EndDate $dayAfterNextWorkingDay -IgnorePatterns $ignorePatterns -LogFile $logFile -PeriodName "Next Working Day"
$thisWeekHours = Get-MeetingHours -Items $items -StartDate $thisWeekBounds.Monday -EndDate $thisWeekBounds.Friday -IgnorePatterns $ignorePatterns -LogFile $logFile -PeriodName "This Week"
$nextWeekHours = Get-MeetingHours -Items $items -StartDate $nextWeekBounds.Monday -EndDate $nextWeekBounds.Friday -IgnorePatterns $ignorePatterns -LogFile $logFile -PeriodName "Next Week"

# -----------------------------------------------------------------------------
# Process Full-Hour Meetings for Rescheduling
# -----------------------------------------------------------------------------

# Load email template
$emailTemplate = Load-EmailTemplate -ScriptDir $scriptDir

# Find full-hour meetings in the next 14 days
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message "SCANNING FOR FULL-HOUR MEETINGS"
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message ""

$fullHourMeetings = Get-FullHourMeetings -Items $items -StartDate $now -EndDate $fourteenDaysLater -IgnorePatterns $ignorePatterns -MaxCount 10 -LogFile $logFile

# Process each full-hour meeting
foreach ($meeting in $fullHourMeetings) {
    # Show confirmation dialog
    $shouldCreateDraft = Show-MeetingRescheduleDialog -Subject $meeting.Subject -StartTime $meeting.Start -Organizer $meeting.Organizer

    if ($shouldCreateDraft) {
        # Create draft email
        $success = New-RescheduleDraftEmail -Outlook $outlook -AppointmentItem $meeting -Template $emailTemplate

        if ($success) {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(
                "Draft email created successfully and saved to your Drafts folder.",
                "Draft Created",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } else {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to create draft email. Please check the error message.",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    }
}

# -----------------------------------------------------------------------------
# Build Windows Forms Popup
# -----------------------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Meeting Hour Summary"
$form.Size = New-Object System.Drawing.Size(450, 400)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.TopMost = $true

# Create title label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$titleLabel.Size = New-Object System.Drawing.Size(410, 30)
$titleLabel.Text = "Meeting Hour Summary"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($titleLabel)

# Create subtitle label with current date
$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Location = New-Object System.Drawing.Point(20, 55)
$subtitleLabel.Size = New-Object System.Drawing.Size(410, 20)
$subtitleLabel.Text = "Generated on: $($now.ToString('dddd, MMMM dd, yyyy HH:mm'))"
$subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$subtitleLabel.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($subtitleLabel)

# Y position for content
$yPos = 90

# Helper function to add a summary row
function Add-SummaryRow {
    param (
        [System.Windows.Forms.Form]$Form,
        [int]$YPosition,
        [string]$Label,
        [string]$Hours,
        [int]$Count,
        [string]$DateRange
    )

    # Period label
    $periodLabel = New-Object System.Windows.Forms.Label
    $periodLabel.Location = New-Object System.Drawing.Point(30, $YPosition)
    $periodLabel.Size = New-Object System.Drawing.Size(150, 20)
    $periodLabel.Text = $Label
    $periodLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $Form.Controls.Add($periodLabel)

    # Hours value
    $hoursLabel = New-Object System.Windows.Forms.Label
    $hoursLabel.Location = New-Object System.Drawing.Point(190, $YPosition)
    $hoursLabel.Size = New-Object System.Drawing.Size(100, 20)
    $hoursLabel.Text = "$Hours hours"
    $hoursLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $Form.Controls.Add($hoursLabel)

    # Meeting count
    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Location = New-Object System.Drawing.Point(300, $YPosition)
    $countLabel.Size = New-Object System.Drawing.Size(120, 20)
    $countLabel.Text = "($Count meetings)"
    $countLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $countLabel.ForeColor = [System.Drawing.Color]::Gray
    $Form.Controls.Add($countLabel)

    # Date range (small text below)
    $dateLabel = New-Object System.Windows.Forms.Label
    $dateLabel.Location = New-Object System.Drawing.Point(30, ($YPosition + 22))
    $dateLabel.Size = New-Object System.Drawing.Size(390, 16)
    $dateLabel.Text = $DateRange
    $dateLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $dateLabel.ForeColor = [System.Drawing.Color]::DarkGray
    $Form.Controls.Add($dateLabel)

    return ($YPosition + 50)
}

# Add summary rows
$yPos = Add-SummaryRow -Form $form -YPosition $yPos -Label "Today:" -Hours $todayHours.Hours -Count $todayHours.Count -DateRange $today.ToString("dddd, MMMM dd")
$yPos = Add-SummaryRow -Form $form -YPosition $yPos -Label "Next Working Day:" -Hours $nextWorkingDayHours.Hours -Count $nextWorkingDayHours.Count -DateRange $nextWorkingDay.ToString("dddd, MMMM dd")
$yPos = Add-SummaryRow -Form $form -YPosition $yPos -Label "This Week:" -Hours $thisWeekHours.Hours -Count $thisWeekHours.Count -DateRange "$($thisWeekBounds.Monday.ToString('MMM dd')) - $($thisWeekBounds.Friday.Date.ToString('MMM dd'))"
$yPos = Add-SummaryRow -Form $form -YPosition $yPos -Label "Next Week:" -Hours $nextWeekHours.Hours -Count $nextWeekHours.Count -DateRange "$($nextWeekBounds.Monday.ToString('MMM dd')) - $($nextWeekBounds.Friday.Date.ToString('MMM dd'))"

# Add separator line
$separator = New-Object System.Windows.Forms.Label
$separator.Location = New-Object System.Drawing.Point(20, ($yPos + 5))
$separator.Size = New-Object System.Drawing.Size(410, 2)
$separator.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$form.Controls.Add($separator)

# Add footer note about ignored patterns
$footerNote = New-Object System.Windows.Forms.Label
$footerNote.Location = New-Object System.Drawing.Point(20, ($yPos + 15))
$footerNote.Size = New-Object System.Drawing.Size(410, 30)
if ($ignorePatterns.Count -gt 0) {
    $footerNote.Text = "Note: $($ignorePatterns.Count) ignore pattern(s) applied from ignore_appointments.txt"
} else {
    $footerNote.Text = "Note: No ignore patterns configured (see ignore_appointments.txt)"
}
$footerNote.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$footerNote.ForeColor = [System.Drawing.Color]::DarkGray
$form.Controls.Add($footerNote)

# Add OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(170, ($yPos + 50))
$okButton.Size = New-Object System.Drawing.Size(100, 30)
$okButton.Text = "OK"
$okButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($okButton)
$form.AcceptButton = $okButton

# Adjust form height based on content
$form.Height = $yPos + 130

# Log final summary
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message "EXECUTION SUMMARY"
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message "Today: $($todayHours.Hours) hours ($($todayHours.Count) meetings)"
Write-Log -LogFile $logFile -Message "Next Working Day ($($nextWorkingDay.ToString('yyyy-MM-dd'))): $($nextWorkingDayHours.Hours) hours ($($nextWorkingDayHours.Count) meetings)"
Write-Log -LogFile $logFile -Message "This Week: $($thisWeekHours.Hours) hours ($($thisWeekHours.Count) meetings)"
Write-Log -LogFile $logFile -Message "Next Week: $($nextWeekHours.Hours) hours ($($nextWeekHours.Count) meetings)"
Write-Log -LogFile $logFile -Message "Full-hour meetings found: $($fullHourMeetings.Count)"
Write-Log -LogFile $logFile -Message ""
Write-Log -LogFile $logFile -Message "Script execution completed successfully"
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message ""

Write-Host "Processing complete. Detailed log saved to: $logFile" -ForegroundColor Green
Write-Host "Check the log file to see all appointments found and filtering decisions." -ForegroundColor Cyan
Write-Host ""

# Show the form
$form.ShowDialog() | Out-Null

Write-Host "For detailed information about which appointments were included/excluded, see: $logFile" -ForegroundColor Yellow

# Cleanup COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendar) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
