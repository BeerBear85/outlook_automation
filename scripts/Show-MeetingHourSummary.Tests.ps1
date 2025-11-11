# Regression Tests for Show-MeetingHourSummary.ps1
# Uses Pester testing framework
# Run with: Invoke-Pester -Path .\Show-MeetingHourSummary.Tests.ps1
#
# Compatible with Pester v3+
#
# @author: Generated for outlook_automation repository
# -----------------------------------------------------------------------------

# Define helper functions at script level for Pester v3 compatibility
function Get-WeekdayBounds {
        param (
            [DateTime]$ReferenceDate
        )

        $dayOfWeek = [int]$ReferenceDate.DayOfWeek

        # For weekends (Saturday=6, Sunday=0), use the upcoming Monday
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
        param (
            [DateTime]$Start,
            [DateTime]$End
        )

        $duration = $End - $Start
        return [Math]::Round($duration.TotalHours, 2)
    }

    function Test-ShouldIgnoreAppointment {
        param (
            [string]$Subject,
            [string[]]$IgnorePatterns
        )

        foreach ($pattern in $IgnorePatterns) {
            # Use -cmatch for case-sensitive matching by default
            # Users can use (?i) in their patterns for case-insensitive
            if ($Subject -cmatch $pattern) {
                return $true
            }
        }

        return $false
    }

    function Load-IgnorePatterns {
        param (
            [string]$ScriptDir
        )

        $ignoreFile = Join-Path $ScriptDir "ignore_appointments.txt"
        $patterns = @()

        if (Test-Path $ignoreFile) {
            Get-Content $ignoreFile | ForEach-Object {
                $line = $_.Trim()
                if ($line -and -not $line.StartsWith("#")) {
                    $patterns += $line
                }
            }
        }

        return $patterns
    }

    function Load-EmailTemplate {
        param (
            [string]$ScriptDir
        )

        $templateFile = Join-Path $ScriptDir "meeting_change_request_template.txt"

        if (Test-Path $templateFile) {
            return Get-Content $templateFile -Raw
        } else {
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

    function Get-FullHourMeetings {
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
            $appointmentStart = $item.Start

            if ($appointmentStart -ge $StartDate -and $appointmentStart -lt $EndDate) {
                if (Test-ShouldIgnoreAppointment -Subject $item.Subject -IgnorePatterns $IgnorePatterns) {
                    continue
                }

                if ($item.AllDayEvent) {
                    continue
                }

                if ($item.Sensitivity -eq 2) {  # olPrivate
                    continue
                }

                if ($item.BusyStatus -eq 3) {  # olOutOfOffice
                    continue
                }

                if ($appointmentStart -lt $now) {
                    continue
                }

                if ($appointmentStart.Minute -eq 0 -and $appointmentStart.Second -eq 0) {
                    $fullHourMeetings += $item
                }
            }
        }

        $fullHourMeetings = $fullHourMeetings | Sort-Object Start | Select-Object -First $MaxCount

        return $fullHourMeetings
    }

# Create a temporary directory for tests that need it
$script:tempDir = Join-Path $env:TEMP "MeetingHourSummaryTests_$([Guid]::NewGuid().ToString())"
if (-not (Test-Path $script:tempDir)) {
    New-Item -Path $script:tempDir -ItemType Directory -Force | Out-Null
}

Describe "Get-WeekdayBounds" {
    Context "When given a Monday" {
        It "Should return the same Monday" {
            $testDate = Get-Date "2025-11-10"  # Monday
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $result.Monday.Date | Should Be $testDate.Date
        }

        It "Should return Friday 4 days later" {
            $testDate = Get-Date "2025-11-10"  # Monday
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $expectedFriday = $testDate.AddDays(4).Date
            $result.Friday.Date | Should Be $expectedFriday
        }
    }

    Context "When given a Wednesday" {
        It "Should return the Monday of the same week" {
            $testDate = Get-Date "2025-11-12"  # Wednesday
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $expectedMonday = Get-Date "2025-11-10"  # Monday
            $result.Monday.Date | Should Be $expectedMonday.Date
        }

        It "Should return the Friday of the same week" {
            $testDate = Get-Date "2025-11-12"  # Wednesday
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $expectedFriday = Get-Date "2025-11-14"  # Friday
            $result.Friday.Date | Should Be $expectedFriday.Date
        }
    }

    Context "When given a Sunday" {
        It "Should return the Monday of the following week" {
            $testDate = Get-Date "2025-11-09"  # Sunday
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $expectedMonday = Get-Date "2025-11-10"  # Monday
            $result.Monday.Date | Should Be $expectedMonday.Date
        }
    }

    Context "When given a Saturday" {
        It "Should return the Monday of the following week" {
            $testDate = Get-Date "2025-11-08"  # Saturday
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $expectedMonday = Get-Date "2025-11-10"  # Monday
            $result.Monday.Date | Should Be $expectedMonday.Date
        }
    }

    Context "When crossing month boundaries" {
        It "Should correctly calculate week bounds across month change" {
            $testDate = Get-Date "2025-10-29"  # Wednesday, near end of October
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $expectedMonday = Get-Date "2025-10-27"  # Monday
            $expectedFriday = Get-Date "2025-10-31"  # Friday
            $result.Monday.Date | Should Be $expectedMonday.Date
            $result.Friday.Date | Should Be $expectedFriday.Date
        }
    }

    Context "When crossing year boundaries" {
        It "Should correctly calculate week bounds across year change" {
            $testDate = Get-Date "2026-01-01"  # Thursday, Jan 1
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $expectedMonday = Get-Date "2025-12-29"  # Monday (previous year)
            $expectedFriday = Get-Date "2026-01-02"  # Friday (new year)
            $result.Monday.Date | Should Be $expectedMonday.Date
            $result.Friday.Date | Should Be $expectedFriday.Date
        }
    }

    Context "Friday end time" {
        It "Should set Friday to end of day (23:59:59)" {
            $testDate = Get-Date "2025-11-10"  # Monday
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $result.Friday.Hour | Should Be 23
            $result.Friday.Minute | Should Be 59
            $result.Friday.Second | Should Be 59
        }
    }
}

Describe "Get-AppointmentDuration" {
    Context "When calculating duration" {
        It "Should calculate 1 hour correctly" {
            $start = Get-Date "2025-11-11 09:00"
            $end = Get-Date "2025-11-11 10:00"

            $result = Get-AppointmentDuration -Start $start -End $end
            $result | Should Be 1.0
        }

        It "Should calculate 30 minutes correctly" {
            $start = Get-Date "2025-11-11 09:00"
            $end = Get-Date "2025-11-11 09:30"

            $result = Get-AppointmentDuration -Start $start -End $end
            $result | Should Be 0.5
        }

        It "Should calculate multi-hour meetings correctly" {
            $start = Get-Date "2025-11-11 09:00"
            $end = Get-Date "2025-11-11 12:00"

            $result = Get-AppointmentDuration -Start $start -End $end
            $result | Should Be 3.0
        }

        It "Should handle odd durations with rounding" {
            $start = Get-Date "2025-11-11 09:00"
            $end = Get-Date "2025-11-11 09:45"

            $result = Get-AppointmentDuration -Start $start -End $end
            $result | Should Be 0.75
        }

        It "Should handle meetings spanning midnight" {
            $start = Get-Date "2025-11-11 23:00"
            $end = Get-Date "2025-11-12 01:00"

            $result = Get-AppointmentDuration -Start $start -End $end
            $result | Should Be 2.0
        }

        It "Should round to 2 decimal places" {
            $start = Get-Date "2025-11-11 09:00"
            $end = Get-Date "2025-11-11 09:20"  # 20 minutes = 0.333... hours

            $result = Get-AppointmentDuration -Start $start -End $end
            $result | Should Be 0.33
        }
    }
}

Describe "Test-ShouldIgnoreAppointment" {
    Context "When testing against ignore patterns" {
        It "Should ignore appointments matching exact pattern" {
            $patterns = @("^Lunch$")

            $result = Test-ShouldIgnoreAppointment -Subject "Lunch" -IgnorePatterns $patterns
            $result | Should Be $true
        }

        It "Should not ignore appointments not matching pattern" {
            $patterns = @("^Lunch$")

            $result = Test-ShouldIgnoreAppointment -Subject "Lunch with Client" -IgnorePatterns $patterns
            $result | Should Be $false
        }

        It "Should ignore appointments matching wildcard pattern" {
            $patterns = @("^Personal.*")

            $result = Test-ShouldIgnoreAppointment -Subject "Personal: Doctor Appointment" -IgnorePatterns $patterns
            $result | Should Be $true
        }

        It "Should ignore appointments matching any pattern in list" {
            $patterns = @("^Lunch$", "^Break$", "^Personal.*")

            $result1 = Test-ShouldIgnoreAppointment -Subject "Lunch" -IgnorePatterns $patterns
            $result2 = Test-ShouldIgnoreAppointment -Subject "Break" -IgnorePatterns $patterns
            $result3 = Test-ShouldIgnoreAppointment -Subject "Personal Time" -IgnorePatterns $patterns

            $result1 | Should Be $true
            $result2 | Should Be $true
            $result3 | Should Be $true
        }

        It "Should handle patterns with special regex characters" {
            $patterns = @(".*\[IGNORE\].*")

            $result = Test-ShouldIgnoreAppointment -Subject "Meeting [IGNORE] with notes" -IgnorePatterns $patterns
            $result | Should Be $true
        }

        It "Should be case-sensitive by default" {
            $patterns = @("^lunch$")

            $result = Test-ShouldIgnoreAppointment -Subject "Lunch" -IgnorePatterns $patterns
            $result | Should Be $false
        }

        It "Should support case-insensitive patterns" {
            $patterns = @("(?i)^lunch$")

            $result = Test-ShouldIgnoreAppointment -Subject "Lunch" -IgnorePatterns $patterns
            $result | Should Be $true
        }

        It "Should handle empty pattern list" {
            $patterns = @()

            $result = Test-ShouldIgnoreAppointment -Subject "Any Meeting" -IgnorePatterns $patterns
            $result | Should Be $false
        }

        It "Should ignore appointments matching partial content patterns" {
            $patterns = @(".*standup.*")

            $result = Test-ShouldIgnoreAppointment -Subject "Daily standup meeting" -IgnorePatterns $patterns
            $result | Should Be $true
        }
    }
}

Describe "Load-IgnorePatterns" {
    Context "When loading ignore patterns" {
        It "Should load patterns from file" {
            $testFile = Join-Path $script:tempDir "ignore_appointments.txt"
            @"
^Lunch$
^Break$
^Personal.*
"@ | Out-File -FilePath $testFile -Encoding utf8

            $result = Load-IgnorePatterns -ScriptDir $script:tempDir

            $result.Count | Should Be 3
            $result[0] | Should Be "^Lunch$"
            $result[1] | Should Be "^Break$"
            $result[2] | Should Be "^Personal.*"
        }

        It "Should ignore comment lines" {
            $testFile = Join-Path $script:tempDir "ignore_appointments.txt"
            @"
# This is a comment
^Lunch$
# Another comment
^Break$
"@ | Out-File -FilePath $testFile -Encoding utf8

            $result = Load-IgnorePatterns -ScriptDir $script:tempDir

            $result.Count | Should Be 2
            $result[0] | Should Be "^Lunch$"
            $result[1] | Should Be "^Break$"
        }

        It "Should ignore empty lines" {
            $testFile = Join-Path $script:tempDir "ignore_appointments.txt"
            @"
^Lunch$

^Break$

"@ | Out-File -FilePath $testFile -Encoding utf8

            $result = Load-IgnorePatterns -ScriptDir $script:tempDir

            $result.Count | Should Be 2
        }

        It "Should trim whitespace from patterns" {
            $testFile = Join-Path $script:tempDir "ignore_appointments.txt"
            @"
  ^Lunch$
^Break$
  ^Personal.*
"@ | Out-File -FilePath $testFile -Encoding utf8

            $result = Load-IgnorePatterns -ScriptDir $script:tempDir

            $result[0] | Should Be "^Lunch$"
            $result[2] | Should Be "^Personal.*"
        }

        It "Should return empty array if file doesn't exist" {
            $nonExistentDir = Join-Path $env:TEMP "NonExistent_$([Guid]::NewGuid().ToString())"

            $result = Load-IgnorePatterns -ScriptDir $nonExistentDir

            $result | Should BeNullOrEmpty
        }

        It "Should handle file with only comments and empty lines" {
            $testFile = Join-Path $script:tempDir "ignore_appointments.txt"
            @"
# Comment 1
# Comment 2

# Comment 3
"@ | Out-File -FilePath $testFile -Encoding utf8

            $result = Load-IgnorePatterns -ScriptDir $script:tempDir

            $result | Should BeNullOrEmpty
        }
    }
}

Describe "Time Period Edge Cases" {
    Context "Weekly calculations with special dates" {
        It "Should handle leap year correctly" {
            $testDate = Get-Date "2024-02-29"  # Leap day, Thursday
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $expectedMonday = Get-Date "2024-02-26"
            $expectedFriday = Get-Date "2024-03-01"
            $result.Monday.Date | Should Be $expectedMonday.Date
            $result.Friday.Date | Should Be $expectedFriday.Date
        }

        It "Should handle DST spring forward" {
            # In most US time zones, DST starts second Sunday in March
            $testDate = Get-Date "2025-03-10"  # Monday during DST transition week
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $result.Monday.Date | Should Be $testDate.Date
            $result.Friday.Date | Should Be $testDate.AddDays(4).Date
        }

        It "Should handle DST fall back" {
            # In most US time zones, DST ends first Sunday in November
            $testDate = Get-Date "2025-11-03"  # Monday during DST transition week
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $result.Monday.Date | Should Be $testDate.Date
            $result.Friday.Date | Should Be $testDate.AddDays(4).Date
        }
    }
}

Describe "Integration Scenarios" {
    Context "Realistic appointment filtering scenarios" {
        It "Should filter out multiple ignore patterns correctly" {
            $patterns = @("^Lunch.*", ".*\[Personal\].*", "^OOO -.*")

            $appointments = @(
                "Team Meeting",
                "Lunch with colleagues",
                "Client Call",
                "[Personal] Doctor",
                "OOO - Vacation",
                "Project Planning"
            )

            $filtered = $appointments | Where-Object {
                -not (Test-ShouldIgnoreAppointment -Subject $_ -IgnorePatterns $patterns)
            }

            $filtered.Count | Should Be 3
            $filtered -contains "Team Meeting" | Should Be $true
            $filtered -contains "Client Call" | Should Be $true
            $filtered -contains "Project Planning" | Should Be $true
        }
    }

    Context "Time calculation scenarios" {
        It "Should calculate total hours for multiple meetings" {
            $meetings = @(
                @{ Start = Get-Date "2025-11-11 09:00"; End = Get-Date "2025-11-11 10:00" },  # 1 hour
                @{ Start = Get-Date "2025-11-11 11:00"; End = Get-Date "2025-11-11 11:30" },  # 0.5 hours
                @{ Start = Get-Date "2025-11-11 14:00"; End = Get-Date "2025-11-11 16:00" }   # 2 hours
            )

            $total = ($meetings | ForEach-Object {
                Get-AppointmentDuration -Start $_.Start -End $_.End
            } | Measure-Object -Sum).Sum

            $total | Should Be 3.5
        }

        It "Should handle back-to-back meetings" {
            $meetings = @(
                @{ Start = Get-Date "2025-11-11 09:00"; End = Get-Date "2025-11-11 10:00" },
                @{ Start = Get-Date "2025-11-11 10:00"; End = Get-Date "2025-11-11 11:00" },
                @{ Start = Get-Date "2025-11-11 11:00"; End = Get-Date "2025-11-11 12:00" }
            )

            $total = ($meetings | ForEach-Object {
                Get-AppointmentDuration -Start $_.Start -End $_.End
            } | Measure-Object -Sum).Sum

            $total | Should Be 3.0
        }
    }
}

Describe "Load-EmailTemplate" {
    Context "When loading email template" {
        It "Should load template from file if it exists" {
            $testFile = Join-Path $script:tempDir "meeting_change_request_template.txt"
            $customTemplate = "Custom template with {ORGANIZER} and {SUBJECT}"
            $customTemplate | Out-File -FilePath $testFile -Encoding utf8 -NoNewline

            $result = Load-EmailTemplate -ScriptDir $script:tempDir

            $result | Should Be $customTemplate
        }

        It "Should return default template if file doesn't exist" {
            $nonExistentDir = Join-Path $env:TEMP "NonExistent_$([Guid]::NewGuid().ToString())"

            $result = Load-EmailTemplate -ScriptDir $nonExistentDir

            $result | Should Match "Subject: Request to shift meeting start time"
            $result | Should Match "\{ORGANIZER\}"
            $result | Should Match "\{SUBJECT\}"
            $result | Should Match "\{START_TIME\}"
            $result | Should Match "\{NEW_START_TIME\}"
        }

        It "Should preserve template placeholders" {
            $result = Load-EmailTemplate -ScriptDir "NonExistentPath"

            $result | Should Match "\{ORGANIZER\}"
            $result | Should Match "\{SUBJECT\}"
            $result | Should Match "\{START_TIME\}"
            $result | Should Match "\{NEW_START_TIME\}"
        }

        It "Should handle multi-line templates" {
            $testFile = Join-Path $script:tempDir "meeting_change_request_template.txt"
            $multilineTemplate = @"
Line 1
Line 2
Line 3
"@
            $multilineTemplate | Out-File -FilePath $testFile -Encoding utf8

            $result = Load-EmailTemplate -ScriptDir $script:tempDir

            $result | Should Match "Line 1"
            $result | Should Match "Line 2"
            $result | Should Match "Line 3"
        }
    }
}

Describe "Get-FullHourMeetings" {
    Context "When filtering meetings" {
        It "Should return meetings starting exactly on the hour" {
            $tomorrow = (Get-Date).Date.AddDays(1)
            $meetings = @(
                [PSCustomObject]@{
                    Subject = "Team Meeting"
                    Start = $tomorrow.AddHours(10)  # 10:00
                    AllDayEvent = $false
                    Sensitivity = 0  # olNormal
                    BusyStatus = 2  # olBusy
                },
                [PSCustomObject]@{
                    Subject = "Client Call"
                    Start = $tomorrow.AddHours(14).AddMinutes(30)  # 14:30
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                }
            )

            $startDate = Get-Date
            $endDate = $startDate.AddDays(14)
            $result = Get-FullHourMeetings -Items $meetings -StartDate $startDate -EndDate $endDate -IgnorePatterns @()

            $result.Count | Should Be 1
            $result[0].Subject | Should Be "Team Meeting"
        }

        It "Should exclude all-day events" {
            $tomorrow = (Get-Date).Date.AddDays(1)
            $meetings = @(
                [PSCustomObject]@{
                    Subject = "All Day Event"
                    Start = $tomorrow.AddHours(10)
                    AllDayEvent = $true
                    Sensitivity = 0
                    BusyStatus = 2
                }
            )

            $startDate = Get-Date
            $endDate = $startDate.AddDays(14)
            $result = Get-FullHourMeetings -Items $meetings -StartDate $startDate -EndDate $endDate -IgnorePatterns @()

            $result.Count | Should Be 0
        }

        It "Should exclude private meetings" {
            $tomorrow = (Get-Date).Date.AddDays(1)
            $meetings = @(
                [PSCustomObject]@{
                    Subject = "Private Meeting"
                    Start = $tomorrow.AddHours(10)
                    AllDayEvent = $false
                    Sensitivity = 2  # olPrivate
                    BusyStatus = 2
                }
            )

            $startDate = Get-Date
            $endDate = $startDate.AddDays(14)
            $result = Get-FullHourMeetings -Items $meetings -StartDate $startDate -EndDate $endDate -IgnorePatterns @()

            $result.Count | Should Be 0
        }

        It "Should exclude Out of Office" {
            $tomorrow = (Get-Date).Date.AddDays(1)
            $meetings = @(
                [PSCustomObject]@{
                    Subject = "OOO"
                    Start = $tomorrow.AddHours(10)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 3  # olOutOfOffice
                }
            )

            $startDate = Get-Date
            $endDate = $startDate.AddDays(14)
            $result = Get-FullHourMeetings -Items $meetings -StartDate $startDate -EndDate $endDate -IgnorePatterns @()

            $result.Count | Should Be 0
        }

        It "Should respect ignore patterns" {
            $tomorrow = (Get-Date).Date.AddDays(1)
            $meetings = @(
                [PSCustomObject]@{
                    Subject = "Lunch Meeting"
                    Start = $tomorrow.AddHours(12)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                }
            )

            $patterns = @("^Lunch.*")
            $startDate = Get-Date
            $endDate = $startDate.AddDays(14)
            $result = Get-FullHourMeetings -Items $meetings -StartDate $startDate -EndDate $endDate -IgnorePatterns $patterns

            $result.Count | Should Be 0
        }

        It "Should limit results to MaxCount" {
            $tomorrow = (Get-Date).Date.AddDays(1)
            $meetings = @()
            for ($i = 0; $i -lt 15; $i++) {
                $meetings += [PSCustomObject]@{
                    Subject = "Meeting $i"
                    Start = $tomorrow.AddHours($i)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                }
            }

            $startDate = Get-Date
            $endDate = $startDate.AddDays(14)
            $result = Get-FullHourMeetings -Items $meetings -StartDate $startDate -EndDate $endDate -IgnorePatterns @() -MaxCount 5

            $result.Count | Should Be 5
        }

        It "Should sort by earliest start time" {
            $tomorrow = (Get-Date).Date.AddDays(1)
            $meetings = @(
                [PSCustomObject]@{
                    Subject = "Meeting C"
                    Start = $tomorrow.AddHours(15)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                },
                [PSCustomObject]@{
                    Subject = "Meeting A"
                    Start = $tomorrow.AddHours(9)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                },
                [PSCustomObject]@{
                    Subject = "Meeting B"
                    Start = $tomorrow.AddHours(11)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                }
            )

            $startDate = Get-Date
            $endDate = $startDate.AddDays(14)
            $result = Get-FullHourMeetings -Items $meetings -StartDate $startDate -EndDate $endDate -IgnorePatterns @()

            $result.Count | Should Be 3
            $result[0].Subject | Should Be "Meeting A"
            $result[1].Subject | Should Be "Meeting B"
            $result[2].Subject | Should Be "Meeting C"
        }

        It "Should exclude meetings with non-zero minutes" {
            $tomorrow = (Get-Date).Date.AddDays(1)
            $meetings = @(
                [PSCustomObject]@{
                    Subject = "Meeting at :15"
                    Start = $tomorrow.AddHours(10).AddMinutes(15)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                },
                [PSCustomObject]@{
                    Subject = "Meeting at :00"
                    Start = $tomorrow.AddHours(10)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                }
            )

            $startDate = Get-Date
            $endDate = $startDate.AddDays(14)
            $result = Get-FullHourMeetings -Items $meetings -StartDate $startDate -EndDate $endDate -IgnorePatterns @()

            $result.Count | Should Be 1
            $result[0].Subject | Should Be "Meeting at :00"
        }

        It "Should only include meetings within date range" {
            $today = Get-Date
            $meetings = @(
                [PSCustomObject]@{
                    Subject = "Past Meeting"
                    Start = $today.AddDays(-1).AddHours(10)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                },
                [PSCustomObject]@{
                    Subject = "Future Meeting in range"
                    Start = $today.AddDays(5).AddHours(10)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                },
                [PSCustomObject]@{
                    Subject = "Future Meeting out of range"
                    Start = $today.AddDays(20).AddHours(10)
                    AllDayEvent = $false
                    Sensitivity = 0
                    BusyStatus = 2
                }
            )

            $startDate = $today
            $endDate = $today.AddDays(14)
            $result = Get-FullHourMeetings -Items $meetings -StartDate $startDate -EndDate $endDate -IgnorePatterns @()

            $result.Count | Should Be 1
            $result[0].Subject | Should Be "Future Meeting in range"
        }
    }
}

Describe "Regression Tests" {
    Context "Known edge cases from production" {
        It "Should handle meeting titles with special characters" {
            $patterns = @(".*\[CANCELLED\].*")

            $result = Test-ShouldIgnoreAppointment -Subject "[CANCELLED] Team Sync" -IgnorePatterns $patterns
            $result | Should Be $true
        }

        It "Should handle Unicode characters in meeting titles" {
            $patterns = @()

            $result = Test-ShouldIgnoreAppointment -Subject "Café Meeting ☕" -IgnorePatterns $patterns
            $result | Should Be $false
        }

        It "Should handle very long meeting titles" {
            $longTitle = "A" * 500  # 500 character title
            $patterns = @("^A+$")

            $result = Test-ShouldIgnoreAppointment -Subject $longTitle -IgnorePatterns $patterns
            $result | Should Be $true
        }

        It "Should handle week calculation on first day of year" {
            $testDate = Get-Date "2025-01-01"  # Wednesday
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $expectedMonday = Get-Date "2024-12-30"  # Monday of that week (previous year)
            $result.Monday.Date | Should Be $expectedMonday.Date
        }

        It "Should handle week calculation on last day of year" {
            $testDate = Get-Date "2025-12-31"  # Wednesday
            $result = Get-WeekdayBounds -ReferenceDate $testDate

            $expectedMonday = Get-Date "2025-12-29"  # Monday
            $expectedFriday = Get-Date "2026-01-02"  # Friday (next year)
            $result.Monday.Date | Should Be $expectedMonday.Date
            $result.Friday.Date | Should Be $expectedFriday.Date
        }
    }
}
