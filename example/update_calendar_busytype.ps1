# Outlook Automation that update simple entries with keywords in the calendar to reflect e.g. presence and type.
# This makes it very easy both from computer/phone/Teams to simply add single word entries like "WFH", "OFF", "SICK" or "Drive"
# to the calendar and have them automatically updated to the correct busy status.

# @author: carsten.fjelkstrup-ext@everllence.com
# -----------------------------------------------------------------------------

# Valid presence values are:
# olFree: Indicates that the time slot is free.
# olTentative: Indicates that the time slot is tentative.
# olBusy: Indicates that the time slot is busy.
# olOutOfOffice: Indicates that the time slot is out of office.
# olWorkingElsewhere: Indicates that the time slot is working elsewhere.

# See also
# https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb176631(v=office.12)
# -----------------------------------------------------------------------------

# Use -Install argument to trigger installing this script as a scheduled task
param (
    [switch]$Install
)

# If the Install switch is provided, set up a scheduled task to run this script on workstation lock
if ($Install) {
    # Check if the script is running with administrative privileges
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "The -Install option requires administrative privileges. Please run it as an administrator."
        Exit 1
    }

    # make new folder for schedtask scripts and copy myself there
    Write-Host "Installing scheduled task to run this script on workstation lock"
    New-Item -Path "$env:USERPROFILE\schedtask" -ItemType Directory -Force | Out-Null
    Copy-Item -Path $MyInvocation.MyCommand.Path -Destination "$env:USERPROFILE\schedtask\update_calendar_busytype.ps1" -Force

    try {
        # Register the scheduled task to run as current user
        $taskXml = @"
<?xml version="1.0"?>
<Task xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
    <RegistrationInfo>
        <Date>2025-02-11T14:52:25.7313086</Date>
        <Author>MD-MAN\xogfoz</Author>
        <URI>\EDCD Outlook Automation - Update Calendar Busy Type</URI>
    </RegistrationInfo>
    <Triggers>
        <SessionStateChangeTrigger>
            <Enabled>true</Enabled>
            <StateChange>SessionLock</StateChange>
        </SessionStateChangeTrigger>
    </Triggers>
    <Actions Context="Author">
        <Exec>
            <Command>powershell.exe</Command>
            <Arguments>-File "$env:USERPROFILE\schedtask\update_calendar_busytype.ps1"</Arguments>
        </Exec>
    </Actions>
</Task>
"@
        $xmlPath = "$env:TEMP\update_calendar_busytype_task.xml"
        $taskXml | Out-File -Encoding ascii -FilePath $xmlPath
        schtasks.exe /Create /TN "EDCD Outlook Automation - Update Calendar Busy Type" /XML $xmlPath /F
        Remove-Item $xmlPath
    } catch {
        Write-Error "Failed to create scheduled task: $_"
        Exit 1
    }
    Exit 0
}

# -----------------------------------------------------------------------------
# If the Install switch is not provided, run the script normally

# Normal run cannot run as admin!?
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script cannot be run with administrative privileges due to lack of Outlook access."
    Exit 1
}

# Load the Outlook COM object
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Get the default Calendar folder
$Calendar = $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)

# Get all future items in the Calendar
$currentDate = $(Get-Date).Date # Set to midnight of today
Write-Host "Fetching calendar items from $($currentDate.ToString("g")) onwards"
$Filter = "[Start] >= '" + $currentDate.ToString("g") + "'"
$Items = $Calendar.Items.Restrict($Filter)

# Iterate through items
foreach ($Item in $Items) {
    if ($Item -is [Microsoft.Office.Interop.Outlook.AppointmentItem]) {
        $UpdatedItem = $false

        # WFH
        if ($Item.Subject -like "WFH") {
            $Item.BusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olWorkingElsewhere
            $UpdatedItem = $true
        }

        # OFF
        elseif ($Item.Subject -like "OFF") {
            $Item.BusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olOutOfOffice
            $UpdatedItem = $true
        }

        # SICK
        elseif ($Item.Subject -like "SICK") {
            $Item.BusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olOutOfOffice
            $UpdatedItem = $true
        }
		
        # Drive is always 0900 to 1000 hrs and buddy at 1700 to 1800 hrs
        elseif ($Item.Subject -eq "Drive") { # no wildcard
            # Add "Drive to work" item
            $ItemDriveToWork = $Item
            $ItemDriveToWork.Subject = "Drive to T41"
            $ItemDriveToWork.BusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olOutOfOffice
            $ItemDriveToWork.AllDayEvent = $false
            $ItemDriveToWork.Start = [datetime]::ParseExact("$($ItemDriveToWork.Start.ToString("yyyy-MM-dd")) 09:00", "yyyy-MM-dd HH:mm", $null)
            $ItemDriveToWork.End = [datetime]::ParseExact("$($ItemDriveToWork.Start.ToString("yyyy-MM-dd")) 10:00", "yyyy-MM-dd HH:mm", $null)
            $ItemDriveToWork.ReminderSet = $false
            $ItemDriveToWork.Save()
            Write-Host "Updated $($ItemDriveToWork.EntryID): $($ItemDriveToWork.Subject) starting $($ItemDriveToWork.Start) and ending $($ItemDriveToWork.End) to $($ItemDriveToWork.BusyStatus)"

            # Add "Drive home" item
            $ItemDriveToHome = $Item.Copy()
            $ItemDriveToHome.Subject = "Drive to B10"
            $ItemDriveToHome.BusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olOutOfOffice
            $ItemDriveToHome.AllDayEvent = $false
            $ItemDriveToHome.Start = [datetime]::ParseExact("$($ItemDriveToHome.Start.ToString("yyyy-MM-dd")) 17:00", "yyyy-MM-dd HH:mm", $null)
            $ItemDriveToHome.End = [datetime]::ParseExact("$($ItemDriveToHome.Start.ToString("yyyy-MM-dd")) 18:00", "yyyy-MM-dd HH:mm", $null)
            $ItemDriveToHome.ReminderSet = $false
            $ItemDriveToHome.Save()
            Write-Host "Added $($ItemDriveToHome.EntryID): $($ItemDriveToHome.Subject) starting $($ItemDriveToHome.Start) and ending $($ItemDriveToHome.End) to $($ItemDriveToHome.BusyStatus)"
        }

        # DEMO TIME: Get Things Done slot every morning before 10, where I don't want to be disturbed
        elseif ($Item.Subject -eq "GTD") {
            # Add "Get Things Done" slot
            $Item.Subject = "Get Things Done!"
            $Item.BusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olBusy
            $Item.AllDayEvent = $false
            $Item.Start = [datetime]::ParseExact("$($Item.Start.ToString("yyyy-MM-dd")) 07:00", "yyyy-MM-dd HH:mm", $null)
            $Item.End = [datetime]::ParseExact("$($Item.Start.ToString("yyyy-MM-dd")) 10:00", "yyyy-MM-dd HH:mm", $null)
            $Item.Categories = "Important"
            $UpdatedItem = $true
        }

        # common save logic here
        if ($UpdatedItem) {
            # Always disable the reminder for these entries
            $Item.ReminderSet = $false
            $Item.Save()
            Write-Host "Updated $($Item.EntryID): $($Item.Subject)"
        }
    }
}

Write-Host "Done"
