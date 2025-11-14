# Meeting Hour Summary Script (Microsoft Graph Version)
# Displays daily and weekly meeting hour summaries via Windows Forms popup
# Uses Microsoft Graph PowerShell SDK instead of COM objects
#
# @author: Generated for outlook_automation repository (Graph migration)
# -----------------------------------------------------------------------------

<#
.SYNOPSIS
    Shows a popup with meeting hour summaries for Today, Tomorrow, This Week, and Next Week.

.DESCRIPTION
    This script reads your Outlook calendar via Microsoft Graph and calculates total meeting hours for:
    - Today
    - Next Working Day (skips weekends)
    - This week (Monday to Friday)
    - Next week (Monday to Friday)

    Appointments matching regex patterns in config/ignore_appointments.txt will be excluded.
    Results are displayed in a Windows Forms popup window with a 5-day bar chart.

    This is the Microsoft Graph version that works with New Outlook.

.EXAMPLE
    .\Show-MeetingHourSummary.ps1
    Displays the meeting hour summary popup.

.NOTES
    Prerequisites:
    - Microsoft.Graph PowerShell module installed
    - Authenticated via Connect-Graph.ps1
    - Required scopes: Calendars.ReadWrite, Mail.ReadWrite, User.Read
#>

# -----------------------------------------------------------------------------
# Module Import and Initialization
# -----------------------------------------------------------------------------

# Import shared module
$moduleFile = Join-Path $PSScriptRoot "OutlookGraphAutomation.psm1"
Import-Module $moduleFile -Force

# Check authentication
if (-not (Test-GraphConnection)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Not authenticated to Microsoft Graph.`n`nPlease run Connect-Graph.ps1 first to authenticate.",
        "Authentication Required",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    Exit 1
}

# -----------------------------------------------------------------------------
# Helper Functions (Graph-specific)
# -----------------------------------------------------------------------------

function Get-DailyMeetingHours {
    <#
    .SYNOPSIS
        Calculates total meeting hours for each of the next N working days (Monday-Friday).
        Skips weekends (Saturday and Sunday).
    .OUTPUTS
        Array of hashtables with Date, Hours, Count, and DayOfWeek properties.
    #>
    param (
        [Object[]]$Events,
        [DateTime]$StartDate,
        [int]$WorkingDayCount,
        [string[]]$IgnorePatterns,
        [string]$LogFile = $null
    )

    $dailyHours = @()
    $currentDate = $StartDate
    $workingDaysFound = 0

    while ($workingDaysFound -lt $WorkingDayCount) {
        # Skip weekends (Saturday = 6, Sunday = 0)
        if ($currentDate.DayOfWeek -ne [DayOfWeek]::Saturday -and $currentDate.DayOfWeek -ne [DayOfWeek]::Sunday) {
            $dayStart = $currentDate
            $dayEnd = $dayStart.AddHours(23).AddMinutes(59).AddSeconds(59)

            $dayResult = Get-MeetingHours -Events $Events -StartDate $dayStart -EndDate $dayEnd -IgnorePatterns $IgnorePatterns -LogFile $LogFile -PeriodName "Working Day $($workingDaysFound + 1) - $($dayStart.ToString('yyyy-MM-dd'))"

            $dailyHours += @{
                Date = $dayStart
                Hours = $dayResult.Hours
                Count = $dayResult.Count
                DayOfWeek = $dayStart.ToString("ddd")
            }

            $workingDaysFound++
        }

        $currentDate = $currentDate.AddDays(1)
    }

    return $dailyHours
}

function Get-MeetingHours {
    <#
    .SYNOPSIS
        Calculates total meeting hours for a time period using Graph events.
    #>
    param (
        [Object[]]$Events,
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

    foreach ($event in $Events) {
        # Convert Graph DateTimeTimeZone to local DateTime
        $eventStart = ConvertTo-LocalDateTime -DateTimeTimeZone $event.Start
        $eventEnd = ConvertTo-LocalDateTime -DateTimeTimeZone $event.End

        # Check if event falls within the time period
        if ($eventStart -ge $StartDate -and $eventStart -lt $EndDate) {
            $duration = Get-AppointmentDuration -Start $eventStart -End $eventEnd

            if ($LogFile) {
                Write-Log -LogFile $LogFile -Message "  Found event: '$($event.Subject)' | Start: $($eventStart.ToString('yyyy-MM-dd HH:mm')) | Duration: $duration hours"

                # Log detailed properties for debugging
                Write-Log -LogFile $LogFile -Message "    Properties: IsCancelled=$($event.IsCancelled), ShowAs=$($event.ShowAs), IsAllDay=$($event.IsAllDay), Sensitivity=$($event.Sensitivity)"
            }

            # Skip if matches ignore pattern
            if (Test-ShouldIgnoreAppointment -Subject $event.Subject -IgnorePatterns $IgnorePatterns) {
                if ($LogFile) {
                    Write-Log -LogFile $LogFile -Message "    -> EXCLUDED: Matches ignore pattern" -Level "SKIP"
                }
                $ignoredCount++
                continue
            }

            # Skip all-day events (typically not meetings)
            if ($event.IsAllDay) {
                if ($LogFile) {
                    Write-Log -LogFile $LogFile -Message "    -> EXCLUDED: All-day event" -Level "SKIP"
                }
                $allDayCount++
                continue
            }

            # Skip cancelled meetings
            if ($event.IsCancelled) {
                if ($LogFile) {
                    Write-Log -LogFile $LogFile -Message "    -> EXCLUDED: Meeting cancelled" -Level "SKIP"
                }
                $ignoredCount++
                continue
            }

            # Skip declined meetings
            if ($event.ResponseStatus -and $event.ResponseStatus.Response -eq "declined") {
                if ($LogFile) {
                    Write-Log -LogFile $LogFile -Message "    -> EXCLUDED: Meeting declined" -Level "SKIP"
                }
                $ignoredCount++
                continue
            }

            # Skip private events
            if ($event.Sensitivity -eq "private") {
                if ($LogFile) {
                    Write-Log -LogFile $LogFile -Message "    -> EXCLUDED: Private event" -Level "SKIP"
                }
                $ignoredCount++
                continue
            }

            # Skip Out of Office events
            if ($event.ShowAs -eq "oof") {
                if ($LogFile) {
                    Write-Log -LogFile $LogFile -Message "    -> EXCLUDED: Out of Office" -Level "SKIP"
                }
                $ignoredCount++
                continue
            }

            $hours = $duration
            $totalHours += $hours
            $appointmentCount++
            if ($LogFile) {
                Write-Log -LogFile $LogFile -Message "    -> INCLUDED in time estimate" -Level "INCL"
            }
        }
    }

    if ($LogFile) {
        $totalExcluded = $ignoredCount + $allDayCount
        Write-Log -LogFile $LogFile -Message "Summary for $PeriodName - Total: $([Math]::Round($totalHours, 2)) hours | Included: $appointmentCount events | Excluded: $totalExcluded"
        Write-Log -LogFile $LogFile -Message ""
    }

    return @{
        Hours = [Math]::Round($totalHours, 2)
        Count = $appointmentCount
    }
}

function Get-FullHourMeetings {
    <#
    .SYNOPSIS
        Finds meetings starting exactly on the full hour in the next 14 days.
        Excludes all-day events, private items, Out of Office entries, and previously ignored meetings.
        Returns at most 10 meetings, starting with the earliest.
    #>
    param (
        [Object[]]$Events,
        [DateTime]$StartDate,
        [DateTime]$EndDate,
        [string[]]$IgnorePatterns,
        [string[]]$IgnoredAppointmentIds = @(),
        [int]$MaxCount = 10,
        [string]$LogFile = $null
    )

    if ($LogFile) {
        Write-Log -LogFile $LogFile -Message "Scanning for full-hour meetings (starting at :00) in the next 14 days"
        if ($IgnoredAppointmentIds.Count -gt 0) {
            Write-Log -LogFile $LogFile -Message "  Loaded $($IgnoredAppointmentIds.Count) previously ignored appointment(s)"
        }
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
        'PreviouslyIgnored' = 0
    }

    foreach ($event in $Events) {
        # Convert Graph DateTimeTimeZone to local DateTime
        $eventStart = ConvertTo-LocalDateTime -DateTimeTimeZone $event.Start

        # Check if event is in the next 14 days
        if ($eventStart -ge $StartDate -and $eventStart -lt $EndDate) {
            # Skip if matches ignore pattern
            if (Test-ShouldIgnoreAppointment -Subject $event.Subject -IgnorePatterns $IgnorePatterns) {
                $skippedReasons['Ignored']++
                continue
            }

            # Skip all-day events
            if ($event.IsAllDay) {
                $skippedReasons['AllDay']++
                continue
            }

            # Skip private items
            if ($event.Sensitivity -eq "private") {
                $skippedReasons['Private']++
                continue
            }

            # Skip Out of Office
            if ($event.ShowAs -eq "oof") {
                $skippedReasons['OutOfOffice']++
                continue
            }

            # Skip meetings that have already started
            if ($eventStart -lt $now) {
                $skippedReasons['AlreadyStarted']++
                continue
            }

            # Skip cancelled meetings
            if ($event.IsCancelled) {
                $skippedReasons['Cancelled']++
                continue
            }

            # Skip declined meetings
            if ($event.ResponseStatus -and $event.ResponseStatus.Response -eq "declined") {
                $skippedReasons['Declined']++
                continue
            }

            # Check if start time is exactly on the hour (minute = 0, second = 0)
            if ($eventStart.Minute -eq 0 -and $eventStart.Second -eq 0) {
                # Check if this event was previously ignored
                $eventId = Get-EventIdentifier -Event $event
                if ($eventId -and $IgnoredAppointmentIds -contains $eventId) {
                    $skippedReasons['PreviouslyIgnored']++
                    if ($LogFile) {
                        Write-Log -LogFile $LogFile -Message "  Skipped previously ignored meeting: '$($event.Subject)' | Start: $($eventStart.ToString('yyyy-MM-dd HH:mm'))"
                    }
                    continue
                }

                $fullHourMeetings += $event
                if ($LogFile) {
                    $organizerName = if ($event.Organizer -and $event.Organizer.EmailAddress) { $event.Organizer.EmailAddress.Name } else { "Unknown" }
                    Write-Log -LogFile $LogFile -Message "  Found full-hour meeting: '$($event.Subject)' | Start: $($eventStart.ToString('yyyy-MM-dd HH:mm')) | Organizer: $organizerName"
                }
            } else {
                $skippedReasons['NotFullHour']++
            }
        }
    }

    # Sort by start time and take the first $MaxCount
    $fullHourMeetings = $fullHourMeetings | Sort-Object { (ConvertTo-LocalDateTime -DateTimeTimeZone $_.Start) } | Select-Object -First $MaxCount

    if ($LogFile) {
        Write-Log -LogFile $LogFile -Message "Full-hour meeting scan complete: Found $($fullHourMeetings.Count) meetings"
        Write-Log -LogFile $LogFile -Message "  Skipped: $($skippedReasons['Ignored']) (ignored pattern), $($skippedReasons['AllDay']) (all-day), $($skippedReasons['Private']) (private), $($skippedReasons['OutOfOffice']) (OOO), $($skippedReasons['AlreadyStarted']) (already started), $($skippedReasons['Cancelled']) (cancelled), $($skippedReasons['Declined']) (declined), $($skippedReasons['PreviouslyIgnored']) (previously ignored)"
        Write-Log -LogFile $LogFile -Message ""
    }

    return $fullHourMeetings
}

function Show-MeetingRescheduleDialog {
    <#
    .SYNOPSIS
        Shows a popup dialog asking if user wants to draft a reschedule email.
    .OUTPUTS
        String indicating user choice: "CreateDraft", "Skip", or "NeverAskAgain".
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
    $form.Size = New-Object System.Drawing.Size(550, 320)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    # Title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(510, 25)
    $titleLabel.Text = "Meeting starts at full hour"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($titleLabel)

    # Message label
    $messageLabel = New-Object System.Windows.Forms.Label
    $messageLabel.Location = New-Object System.Drawing.Point(20, 55)
    $messageLabel.Size = New-Object System.Drawing.Size(510, 60)
    $messageLabel.Text = "The following meeting starts exactly on the hour. Would you like to draft an email requesting it be moved to :05?"
    $messageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($messageLabel)

    # Details label
    $detailsLabel = New-Object System.Windows.Forms.Label
    $detailsLabel.Location = New-Object System.Drawing.Point(20, 120)
    $detailsLabel.Size = New-Object System.Drawing.Size(510, 80)
    $detailsLabel.Text = "Subject: $Subject`r`nStart Time: $($StartTime.ToString('dddd, MMMM dd, yyyy HH:mm'))`r`nOrganizer: $Organizer"
    $detailsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $detailsLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Controls.Add($detailsLabel)

    # Store user choice
    $script:userChoice = "Skip"

    # Create Draft button
    $createDraftButton = New-Object System.Windows.Forms.Button
    $createDraftButton.Location = New-Object System.Drawing.Point(30, 220)
    $createDraftButton.Size = New-Object System.Drawing.Size(150, 35)
    $createDraftButton.Text = "Create Draft Email"
    $createDraftButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $createDraftButton.Add_Click({
        $script:userChoice = "CreateDraft"
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($createDraftButton)

    # Skip button
    $skipButton = New-Object System.Windows.Forms.Button
    $skipButton.Location = New-Object System.Drawing.Point(200, 220)
    $skipButton.Size = New-Object System.Drawing.Size(150, 35)
    $skipButton.Text = "Skip for Now"
    $skipButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $skipButton.Add_Click({
        $script:userChoice = "Skip"
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })
    $form.Controls.Add($skipButton)

    # Never Ask Again button
    $neverAskButton = New-Object System.Windows.Forms.Button
    $neverAskButton.Location = New-Object System.Drawing.Point(370, 220)
    $neverAskButton.Size = New-Object System.Drawing.Size(150, 35)
    $neverAskButton.Text = "Never Ask Again"
    $neverAskButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $neverAskButton.ForeColor = [System.Drawing.Color]::DarkRed
    $neverAskButton.Add_Click({
        $script:userChoice = "NeverAskAgain"
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
        $form.Close()
    })
    $form.Controls.Add($neverAskButton)

    $form.AcceptButton = $createDraftButton
    $form.CancelButton = $skipButton

    # Show the form and return user choice
    $form.ShowDialog() | Out-Null
    return $script:userChoice
}

function New-RescheduleDraftEmail {
    <#
    .SYNOPSIS
        Creates a draft email in Outlook requesting meeting reschedule using Graph API.
    #>
    param (
        [Object]$Event,
        [string]$Template,
        [string]$LogFile = $null
    )

    try {
        # Get organizer email address
        $organizerEmail = ""
        $organizerName = ""

        if ($Event.Organizer -and $Event.Organizer.EmailAddress) {
            $organizerEmail = $Event.Organizer.EmailAddress.Address
            $organizerName = $Event.Organizer.EmailAddress.Name
        }

        if (-not $organizerEmail) {
            throw "Cannot determine organizer email address"
        }

        # Get event start time
        $currentStart = ConvertTo-LocalDateTime -DateTimeTimeZone $Event.Start

        # Calculate new start time (add 5 minutes)
        $newStart = $currentStart.AddMinutes(5)

        # Replace placeholders in template
        $emailContent = $Template -replace '\{ORGANIZER\}', $organizerName `
                                   -replace '\{SUBJECT\}', $Event.Subject `
                                   -replace '\{START_TIME\}', $currentStart.ToString('dddd, MMMM dd, yyyy HH:mm') `
                                   -replace '\{NEW_START_TIME\}', $newStart.ToString('HH:mm')

        # Extract subject line from template (first line after "Subject:")
        $subjectLine = "Request to shift meeting start time to :05"
        if ($emailContent -match 'Subject:\s*(.+)') {
            $subjectLine = $matches[1].Trim()
            # Remove the subject line from the body
            $emailContent = $emailContent -replace 'Subject:\s*.+\r?\n\r?\n?', ''
        }

        # Create draft email using Graph API
        $success = New-GraphDraftEmail -ToAddress $organizerEmail -Subject $subjectLine -Body $emailContent.Trim() -LogFile $LogFile

        return $success
    }
    catch {
        if ($LogFile) {
            Write-Log -LogFile $LogFile -Message "Failed to create draft email: $_" -Level "ERROR"
        }
        Write-Error "Failed to create draft email: $_"
        return $false
    }
}

# -----------------------------------------------------------------------------
# Main Script Logic
# -----------------------------------------------------------------------------

# Check if running as admin (not required for Graph, but warn anyway)
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "This script is running with administrative privileges. This is not recommended.`n`nPlease run as a regular user.",
        "Administrator Warning",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
}

# Initialize logging
$scriptDir = $PSScriptRoot
$logFile = Initialize-LogFile -LogDirectory $scriptDir

Write-Host "Meeting Hour Summary Script (Microsoft Graph Version)" -ForegroundColor Green
Write-Host "Logging to: $logFile" -ForegroundColor Green
Write-Host ""

Write-Log -LogFile $logFile -Message "Script execution started"
Write-Log -LogFile $logFile -Message "Script directory: $scriptDir"
Write-Log -LogFile $logFile -Message "Using Microsoft Graph PowerShell SDK"

# Load ignore patterns
$ignorePatterns = Load-IgnorePatterns
Write-Log -LogFile $logFile -Message "Loaded $($ignorePatterns.Count) ignore patterns from config/ignore_appointments.txt"
if ($ignorePatterns.Count -gt 0) {
    foreach ($pattern in $ignorePatterns) {
        Write-Log -LogFile $logFile -Message "  - Pattern: $pattern"
    }
}
Write-Log -LogFile $logFile -Message ""

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

Write-Log -LogFile $logFile -Message "Connecting to Microsoft Graph..."
Write-Log -LogFile $logFile -Message "Fetching calendar events from $($fetchStartDate.ToString('yyyy-MM-dd')) to $($fetchEndDate.ToString('yyyy-MM-dd'))"
Write-Log -LogFile $logFile -Message ""

try {
    $events = Get-GraphCalendarEvents -StartDate $fetchStartDate -EndDate $fetchEndDate.AddDays(1) -LogFile $logFile
    Write-Log -LogFile $logFile -Message ""
    Write-Log -LogFile $logFile -Message "========================================="
    Write-Log -LogFile $logFile -Message "ALL CALENDAR EVENTS RETRIEVED FROM GRAPH"
    Write-Log -LogFile $logFile -Message "========================================="
    Write-Log -LogFile $logFile -Message ""
    foreach ($event in ($events | Sort-Object { (ConvertTo-LocalDateTime -DateTimeTimeZone $_.Start) })) {
        $eventStart = ConvertTo-LocalDateTime -DateTimeTimeZone $event.Start
        Write-Log -LogFile $logFile -Message "  $($eventStart.ToString('yyyy-MM-dd HH:mm')) | $($event.Subject)"
    }
    Write-Log -LogFile $logFile -Message ""
} catch {
    Write-Log -LogFile $logFile -Message "ERROR: Failed to retrieve calendar events - $_" -Level "ERROR"
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to retrieve calendar events from Microsoft Graph.`n`nError: $_`n`nPlease ensure you are authenticated (run Connect-Graph.ps1)",
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

$todayHours = Get-MeetingHours -Events $events -StartDate $today -EndDate $nextWorkingDay -IgnorePatterns $ignorePatterns -LogFile $logFile -PeriodName "Today"
$nextWorkingDayHours = Get-MeetingHours -Events $events -StartDate $nextWorkingDay -EndDate $dayAfterNextWorkingDay -IgnorePatterns $ignorePatterns -LogFile $logFile -PeriodName "Next Working Day"
$thisWeekHours = Get-MeetingHours -Events $events -StartDate $thisWeekBounds.Monday -EndDate $thisWeekBounds.Friday -IgnorePatterns $ignorePatterns -LogFile $logFile -PeriodName "This Week"
$nextWeekHours = Get-MeetingHours -Events $events -StartDate $nextWeekBounds.Monday -EndDate $nextWeekBounds.Friday -IgnorePatterns $ignorePatterns -LogFile $logFile -PeriodName "Next Week"

# Calculate meeting hours for the next 5 working days (for bar chart)
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message "CALCULATING DAILY HOURS FOR BAR CHART (NEXT 5 WORKING DAYS)"
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message ""
$dailyHours = Get-DailyMeetingHours -Events $events -StartDate $today -WorkingDayCount 5 -IgnorePatterns $ignorePatterns -LogFile $logFile

# -----------------------------------------------------------------------------
# Process Full-Hour Meetings for Rescheduling
# -----------------------------------------------------------------------------

# Load email template
$emailTemplate = Load-EmailTemplate

# Load ignored appointments list
$ignoredAppointmentIds = Load-IgnoredFullHourAppointments
Write-Log -LogFile $logFile -Message "Loaded $($ignoredAppointmentIds.Count) ignored full-hour appointment(s)"
Write-Log -LogFile $logFile -Message ""

# Find full-hour meetings in the next 14 days
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message "SCANNING FOR FULL-HOUR MEETINGS"
Write-Log -LogFile $logFile -Message "========================================="
Write-Log -LogFile $logFile -Message ""

$fullHourMeetings = Get-FullHourMeetings -Events $events -StartDate $now -EndDate $fourteenDaysLater -IgnorePatterns $ignorePatterns -IgnoredAppointmentIds $ignoredAppointmentIds -MaxCount 10 -LogFile $logFile

# Process each full-hour meeting
foreach ($meeting in $fullHourMeetings) {
    $meetingStart = ConvertTo-LocalDateTime -DateTimeTimeZone $meeting.Start
    $organizerName = if ($meeting.Organizer -and $meeting.Organizer.EmailAddress) { $meeting.Organizer.EmailAddress.Name } else { "Unknown" }

    # Show confirmation dialog with three options
    $userChoice = Show-MeetingRescheduleDialog -Subject $meeting.Subject -StartTime $meetingStart -Organizer $organizerName

    if ($userChoice -eq "CreateDraft") {
        # Create draft email
        $success = New-RescheduleDraftEmail -Event $meeting -Template $emailTemplate -LogFile $logFile

        if ($success) {
            Write-Log -LogFile $logFile -Message "Draft email created for meeting: '$($meeting.Subject)'"
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(
                "Draft email created successfully and saved to your Drafts folder.",
                "Draft Created",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } else {
            Write-Log -LogFile $logFile -Message "ERROR: Failed to create draft email for meeting: '$($meeting.Subject)'" -Level "ERROR"
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to create draft email. Please check the error message and log file.",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    }
    elseif ($userChoice -eq "NeverAskAgain") {
        # Add event to ignore list
        $eventId = Get-EventIdentifier -Event $meeting
        if ($eventId) {
            Save-IgnoredFullHourAppointment -Identifier $eventId -Subject $meeting.Subject -StartTime $meetingStart
            Write-Log -LogFile $logFile -Message "Added to ignore list: '$($meeting.Subject)' | ID: $eventId"
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(
                "This meeting has been added to the ignore list and will not be shown again.`n`nYou can manually edit config/ignored_full_hour_appointments.txt to remove it if needed.",
                "Added to Ignore List",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } else {
            Write-Log -LogFile $logFile -Message "WARNING: Could not get identifier for meeting: '$($meeting.Subject)'" -Level "WARN"
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(
                "Unable to get a stable identifier for this meeting. It cannot be added to the ignore list.",
                "Warning",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
    }
    else {
        # User chose "Skip" - do nothing, just log it
        Write-Log -LogFile $logFile -Message "User skipped meeting: '$($meeting.Subject)'"
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
$subtitleLabel.Text = "Generated on: $($now.ToString('dddd, MMMM dd, yyyy HH:mm')) | Microsoft Graph"
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

# Add 5-day bar chart
$yPos += 10

# Chart title
$chartTitleLabel = New-Object System.Windows.Forms.Label
$chartTitleLabel.Location = New-Object System.Drawing.Point(20, $yPos)
$chartTitleLabel.Size = New-Object System.Drawing.Size(410, 20)
$chartTitleLabel.Text = "Next 5 Working Days Overview"
$chartTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($chartTitleLabel)

$yPos += 25

# Create panel for bar chart
$chartPanel = New-Object System.Windows.Forms.Panel
$chartPanel.Location = New-Object System.Drawing.Point(20, $yPos)
$chartPanel.Size = New-Object System.Drawing.Size(410, 180)
$chartPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

# Store daily hours data in the panel's Tag property for access in Paint event
$chartPanel.Tag = $dailyHours

# Add Paint event handler for drawing the bar chart
$chartPanel.Add_Paint({
    param($sender, $e)

    $graphics = $e.Graphics
    $data = $sender.Tag

    if ($null -eq $data -or $data.Count -eq 0) {
        return
    }

    # Chart dimensions
    $chartWidth = $sender.Width - 20
    $chartHeight = $sender.Height - 60
    $barWidth = [Math]::Floor($chartWidth / ($data.Count * 1.5))
    $barSpacing = [Math]::Floor($barWidth / 2)
    $maxHours = 8  # Max scale for the chart
    $startX = 10
    $startY = 10

    # Draw each bar
    for ($i = 0; $i -lt $data.Count; $i++) {
        $dayData = $data[$i]
        $hours = $dayData.Hours

        # Determine bar color based on thresholds
        if ($hours -le 3) {
            $barColor = [System.Drawing.Color]::FromArgb(34, 139, 34)  # Green
        } elseif ($hours -le 4) {
            $barColor = [System.Drawing.Color]::FromArgb(255, 193, 7)  # Yellow
        } else {
            $barColor = [System.Drawing.Color]::FromArgb(220, 53, 69)  # Red
        }

        # Calculate bar height (proportional to hours)
        $barHeight = [Math]::Min(($hours / $maxHours) * $chartHeight, $chartHeight)

        # Calculate bar position
        $barX = $startX + ($i * ($barWidth + $barSpacing))
        $barY = $startY + $chartHeight - $barHeight

        # Draw bar
        $brush = New-Object System.Drawing.SolidBrush($barColor)
        $graphics.FillRectangle($brush, $barX, $barY, $barWidth, $barHeight)
        $brush.Dispose()

        # Draw border around bar
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 1)
        $graphics.DrawRectangle($pen, $barX, $barY, $barWidth, $barHeight)
        $pen.Dispose()

        # Draw hours value on top of bar
        $hoursText = "$($hours)h"
        $font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
        $textSize = $graphics.MeasureString($hoursText, $font)
        $textX = $barX + ($barWidth - $textSize.Width) / 2
        $textY = [Math]::Max($barY - $textSize.Height - 2, 2)
        $graphics.DrawString($hoursText, $font, $textBrush, $textX, $textY)
        $font.Dispose()
        $textBrush.Dispose()

        # Draw day label below bar
        $dayLabel = "$($dayData.DayOfWeek)`n$($dayData.Date.ToString('MM/dd'))"
        $labelFont = New-Object System.Drawing.Font("Segoe UI", 7)
        $labelBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
        $labelSize = $graphics.MeasureString($dayLabel, $labelFont)
        $labelX = $barX + ($barWidth - $labelSize.Width) / 2
        $labelY = $startY + $chartHeight + 5
        $graphics.DrawString($dayLabel, $labelFont, $labelBrush, $labelX, $labelY)
        $labelFont.Dispose()
        $labelBrush.Dispose()
    }
})

$form.Controls.Add($chartPanel)
$yPos += 185

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
    $footerNote.Text = "Note: $($ignorePatterns.Count) ignore pattern(s) applied | Powered by Microsoft Graph"
} else {
    $footerNote.Text = "Note: No ignore patterns configured | Powered by Microsoft Graph"
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
Write-Host "Check the log file to see all events found and filtering decisions." -ForegroundColor Cyan
Write-Host ""

# Show the form
$form.ShowDialog() | Out-Null

Write-Host "For detailed information about which events were included/excluded, see: $logFile" -ForegroundColor Yellow
