# Test Microsoft Graph Connection
# Utility script to verify authentication and permissions
#
# @author: Generated for outlook_automation repository (Graph migration)
# -----------------------------------------------------------------------------

<#
.SYNOPSIS
    Tests Microsoft Graph authentication and permissions.

.DESCRIPTION
    This utility script checks if you are properly authenticated to Microsoft Graph
    and validates that you have the necessary permissions to run Outlook automation scripts.

    Use this script to troubleshoot connection issues.

.EXAMPLE
    .\Test-GraphConnection.ps1
    Runs connection and permission tests.
#>

# Import the module
$moduleFile = Join-Path $PSScriptRoot "OutlookGraphAutomation.psm1"
Import-Module $moduleFile -Force

Write-Host "Microsoft Graph Connection Test" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# -----------------------------------------------------------------------------
# Test 1: Check if authenticated
# -----------------------------------------------------------------------------

Write-Host "[Test 1/4] Checking authentication status..." -ForegroundColor Yellow

$isConnected = Test-GraphConnection

if ($isConnected) {
    Write-Host "  PASS: Authenticated to Microsoft Graph" -ForegroundColor Green

    try {
        $context = Get-MgContext
        Write-Host "    Account: $($context.Account)" -ForegroundColor Gray
        Write-Host "    Scopes:  $($context.Scopes -join ', ')" -ForegroundColor Gray
    }
    catch {
        # Context info not critical
    }
} else {
    Write-Host "  FAIL: Not authenticated" -ForegroundColor Red
    Write-Host "    Please run: .\Connect-Graph.ps1" -ForegroundColor Yellow
    Write-Host ""
    Exit 1
}

Write-Host ""

# -----------------------------------------------------------------------------
# Test 2: Check required scopes
# -----------------------------------------------------------------------------

Write-Host "[Test 2/4] Checking required permissions..." -ForegroundColor Yellow

try {
    $context = Get-MgContext
    $currentScopes = $context.Scopes
    $requiredScopes = @("Calendars.ReadWrite", "Mail.ReadWrite", "User.Read")

    $missingScopes = @()
    foreach ($scope in $requiredScopes) {
        if ($currentScopes -contains $scope) {
            Write-Host "  PASS: $scope" -ForegroundColor Green
        } else {
            Write-Host "  FAIL: $scope (missing)" -ForegroundColor Red
            $missingScopes += $scope
        }
    }

    if ($missingScopes.Count -gt 0) {
        Write-Host ""
        Write-Host "  Missing permissions detected!" -ForegroundColor Red
        Write-Host "  Please run: .\Connect-Graph.ps1" -ForegroundColor Yellow
        Write-Host ""
        Exit 1
    }
}
catch {
    Write-Host "  ERROR: Failed to check permissions" -ForegroundColor Red
    Write-Host "    $_" -ForegroundColor Red
    Write-Host ""
    Exit 1
}

Write-Host ""

# -----------------------------------------------------------------------------
# Test 3: Test calendar access
# -----------------------------------------------------------------------------

Write-Host "[Test 3/4] Testing calendar access..." -ForegroundColor Yellow

try {
    # Try to get calendar events for today
    $today = Get-Date
    $tomorrow = $today.AddDays(1)

    $startDateTime = $today.ToString("yyyy-MM-ddTHH:mm:ss")
    $endDateTime = $tomorrow.ToString("yyyy-MM-ddTHH:mm:ss")

    $events = Get-MgUserCalendarView -UserId "me" `
        -StartDateTime $startDateTime `
        -EndDateTime $endDateTime `
        -Top 5 `
        -ErrorAction Stop

    Write-Host "  PASS: Successfully accessed calendar" -ForegroundColor Green
    Write-Host "    Found $($events.Count) event(s) for today" -ForegroundColor Gray
}
catch {
    Write-Host "  FAIL: Cannot access calendar" -ForegroundColor Red
    Write-Host "    $_" -ForegroundColor Red
    Write-Host ""
    Exit 1
}

Write-Host ""

# -----------------------------------------------------------------------------
# Test 4: Test user profile access
# -----------------------------------------------------------------------------

Write-Host "[Test 4/4] Testing user profile access..." -ForegroundColor Yellow

try {
    $user = Get-MgUser -UserId "me" -Property DisplayName,UserPrincipalName,MailboxSettings -ErrorAction Stop

    Write-Host "  PASS: Successfully accessed user profile" -ForegroundColor Green
    Write-Host "    Display Name: $($user.DisplayName)" -ForegroundColor Gray
    Write-Host "    Email:        $($user.UserPrincipalName)" -ForegroundColor Gray

    if ($user.MailboxSettings -and $user.MailboxSettings.TimeZone) {
        Write-Host "    Timezone:     $($user.MailboxSettings.TimeZone)" -ForegroundColor Gray
    }
}
catch {
    Write-Host "  FAIL: Cannot access user profile" -ForegroundColor Red
    Write-Host "    $_" -ForegroundColor Red
    Write-Host ""
    Exit 1
}

Write-Host ""

# -----------------------------------------------------------------------------
# Summary
# -----------------------------------------------------------------------------

Write-Host "================================" -ForegroundColor Cyan
Write-Host "All tests PASSED!" -ForegroundColor Green
Write-Host ""
Write-Host "Your Microsoft Graph connection is working correctly." -ForegroundColor Green
Write-Host "You can now run the automation scripts:" -ForegroundColor Cyan
Write-Host "  .\Show-MeetingHourSummary.ps1" -ForegroundColor White
Write-Host ""
