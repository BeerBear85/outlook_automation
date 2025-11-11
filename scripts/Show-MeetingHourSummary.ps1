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
        [int]$MaxCount = 10
    )

    $fullHourMeetings = @()
    $now = Get-Date

    foreach ($item in $Items) {
        if ($item -is [Microsoft.Office.Interop.Outlook.AppointmentItem]) {
            $appointmentStart = $item.Start

            # Check if appointment is in the next 14 days
            if ($appointmentStart -ge $StartDate -and $appointmentStart -lt $EndDate) {
                # Skip if matches ignore pattern
                if (Test-ShouldIgnoreAppointment -Subject $item.Subject -IgnorePatterns $IgnorePatterns) {
                    continue
                }

                # Skip all-day events
                if ($item.AllDayEvent) {
                    continue
                }

                # Skip private items
                if ($item.Sensitivity -eq [Microsoft.Office.Interop.Outlook.OlSensitivity]::olPrivate) {
                    continue
                }

                # Skip Out of Office (BusyStatus = olOutOfOffice)
                if ($item.BusyStatus -eq [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olOutOfOffice) {
                    continue
                }

                # Skip meetings that have already started
                if ($appointmentStart -lt $now) {
                    continue
                }

                # Check if start time is exactly on the hour (minute = 0, second = 0)
                if ($appointmentStart.Minute -eq 0 -and $appointmentStart.Second -eq 0) {
                    $fullHourMeetings += $item
                }
            }
        }
    }

    # Sort by start time and take the first $MaxCount
    $fullHourMeetings = $fullHourMeetings | Sort-Object Start | Select-Object -First $MaxCount

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

function Get-MeetingHours {
    <#
    .SYNOPSIS
        Calculates total meeting hours for a time period.
    #>
    param (
        [Object]$Items,
        [DateTime]$StartDate,
        [DateTime]$EndDate,
        [string[]]$IgnorePatterns
    )

    $totalHours = 0
    $appointmentCount = 0

    foreach ($item in $Items) {
        if ($item -is [Microsoft.Office.Interop.Outlook.AppointmentItem]) {
            $appointmentStart = $item.Start
            $appointmentEnd = $item.End

            # Check if appointment falls within the time period
            if ($appointmentStart -ge $StartDate -and $appointmentStart -lt $EndDate) {
                # Skip if matches ignore pattern
                if (Test-ShouldIgnoreAppointment -Subject $item.Subject -IgnorePatterns $IgnorePatterns) {
                    continue
                }

                # Skip all-day events (typically not meetings)
                if ($item.AllDayEvent) {
                    continue
                }

                $hours = Get-AppointmentDuration -Start $appointmentStart -End $appointmentEnd
                $totalHours += $hours
                $appointmentCount++
            }
        }
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

# Load ignore patterns
$scriptDir = Get-ScriptDirectory
$ignorePatterns = Load-IgnorePatterns -ScriptDir $scriptDir

# Load Outlook COM object
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
} catch {
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
$tomorrow = $today.AddDays(1)
$dayAfterTomorrow = $tomorrow.AddDays(1)

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
} catch {
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
$todayHours = Get-MeetingHours -Items $items -StartDate $today -EndDate $tomorrow -IgnorePatterns $ignorePatterns
$tomorrowHours = Get-MeetingHours -Items $items -StartDate $tomorrow -EndDate $dayAfterTomorrow -IgnorePatterns $ignorePatterns
$thisWeekHours = Get-MeetingHours -Items $items -StartDate $thisWeekBounds.Monday -EndDate $thisWeekBounds.Friday -IgnorePatterns $ignorePatterns
$nextWeekHours = Get-MeetingHours -Items $items -StartDate $nextWeekBounds.Monday -EndDate $nextWeekBounds.Friday -IgnorePatterns $ignorePatterns

# -----------------------------------------------------------------------------
# Process Full-Hour Meetings for Rescheduling
# -----------------------------------------------------------------------------

# Load email template
$emailTemplate = Load-EmailTemplate -ScriptDir $scriptDir

# Find full-hour meetings in the next 14 days
$fullHourMeetings = Get-FullHourMeetings -Items $items -StartDate $now -EndDate $fourteenDaysLater -IgnorePatterns $ignorePatterns -MaxCount 10

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
$yPos = Add-SummaryRow -Form $form -YPosition $yPos -Label "Tomorrow:" -Hours $tomorrowHours.Hours -Count $tomorrowHours.Count -DateRange $tomorrow.ToString("dddd, MMMM dd")
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

# Show the form
$form.ShowDialog() | Out-Null

# Cleanup COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendar) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
