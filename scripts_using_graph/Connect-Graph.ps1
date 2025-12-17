# Connect to Microsoft Graph for Outlook Automation
# Interactive authentication script for Microsoft Graph PowerShell SDK
#
# @author: Generated for outlook_automation repository (Graph migration)
# -----------------------------------------------------------------------------

<#
.SYNOPSIS
    Authenticates to Microsoft Graph with required permissions.

.DESCRIPTION
    This script performs interactive authentication to Microsoft Graph using
    the Microsoft.Graph PowerShell SDK. It requests the necessary scopes for:
    - Reading and writing calendar events
    - Creating draft emails
    - Reading user profile information

    You must run this script once per PowerShell session before running any
    other Graph-based automation scripts.

.PARAMETER

 None

.EXAMPLE
    .\Connect-Graph.ps1
    Performs interactive authentication and displays connection status.

.NOTES
    Required Module: Microsoft.Graph
    Install with: Install-Module Microsoft.Graph -Scope CurrentUser
#>

# -----------------------------------------------------------------------------
# Check Prerequisites
# -----------------------------------------------------------------------------

Write-Host "Microsoft Graph Connection Script" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Check if Microsoft.Graph module is installed
Write-Host "[1/4] Checking for Microsoft.Graph module..." -ForegroundColor Yellow

$graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication
if (-not $graphModule) {
    Write-Host "ERROR: Microsoft.Graph module not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install the module with:" -ForegroundColor Yellow
    Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor White
    Write-Host ""
    Exit 1
}

Write-Host "  Found Microsoft.Graph version $($graphModule[0].Version)" -ForegroundColor Green
Write-Host ""

# -----------------------------------------------------------------------------
# Disconnect existing connection (if any)
# -----------------------------------------------------------------------------

Write-Host "[2/4] Checking existing connection..." -ForegroundColor Yellow

try {
    $existingContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($null -ne $existingContext) {
        Write-Host "  Existing connection found (User: $($existingContext.Account))" -ForegroundColor Cyan
        Write-Host "  Disconnecting to refresh connection..." -ForegroundColor Cyan
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    } else {
        Write-Host "  No existing connection" -ForegroundColor Gray
    }
}
catch {
    # Ignore errors, just ensure we're starting fresh
}

Write-Host ""

# -----------------------------------------------------------------------------
# Connect to Microsoft Graph
# -----------------------------------------------------------------------------

Write-Host "[3/4] Connecting to Microsoft Graph..." -ForegroundColor Yellow
Write-Host "  A browser window will open for authentication." -ForegroundColor Cyan
Write-Host "  Please sign in with your Microsoft account." -ForegroundColor Cyan
Write-Host ""

try {
    # Define required scopes
    $scopes = @(
        "Calendars.ReadWrite",  # Read and write calendar events
        "Mail.ReadWrite",       # Create draft emails
        "User.Read"             # Read user profile
    )

    # Connect with interactive authentication
    Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop

    # Note: Select-MgProfile is deprecated in Microsoft.Graph v2.0+
    # The module now uses v1.0 API by default (use .Beta cmdlets for beta API)

    Write-Host "  Successfully connected!" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "  ERROR: Failed to connect to Microsoft Graph" -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    Write-Host ""
    Exit 1
}

# -----------------------------------------------------------------------------
# Validate Connection and Display Info
# -----------------------------------------------------------------------------

Write-Host "[4/4] Validating connection..." -ForegroundColor Yellow

try {
    # Get connection context
    $context = Get-MgContext -ErrorAction Stop

    Write-Host "  Connection validated successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Connection Details:" -ForegroundColor Cyan
    Write-Host "===================" -ForegroundColor Cyan
    Write-Host "  Account:     $($context.Account)" -ForegroundColor White
    Write-Host "  App Name:    $($context.AppName)" -ForegroundColor White
    Write-Host "  Tenant ID:   $($context.TenantId)" -ForegroundColor White
    Write-Host "  Scopes:      $($context.Scopes -join ', ')" -ForegroundColor White
    Write-Host ""

    # Get user info from context (doesn't require API call)
    Write-Host "  Signed in as: $($context.Account)" -ForegroundColor Green
    Write-Host ""

    # Try to test calendar access (the actual permission we need)
    Write-Host "Testing calendar access..." -ForegroundColor Yellow
    try {
        # Try to get calendar events as a real connectivity test
        $null = Get-MgUserCalendarView -UserId $context.Account -StartDateTime (Get-Date).ToString("yyyy-MM-ddT00:00:00") -EndDateTime (Get-Date).AddDays(1).ToString("yyyy-MM-ddT23:59:59") -Top 1 -ErrorAction Stop
        Write-Host "  Calendar access confirmed!" -ForegroundColor Green
    }
    catch {
        # If calendar access fails, try user profile as fallback
        Write-Host "  Warning: Could not verify calendar access directly" -ForegroundColor Yellow
        try {
            $user = Get-MgUser -UserId $context.Account -Property DisplayName,UserPrincipalName -ErrorAction Stop
            Write-Host "  User profile access confirmed for: $($user.DisplayName)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Warning: Could not verify user profile access" -ForegroundColor Yellow
            Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Gray
            Write-Host "  You may need to re-consent permissions in your browser." -ForegroundColor Yellow
        }
    }
    Write-Host ""

    Write-Host "SUCCESS: Connected to Microsoft Graph!" -ForegroundColor Green
    Write-Host ""
    Write-Host "You can now run:" -ForegroundColor Cyan
    Write-Host "  .\Show-MeetingHourSummary.ps1" -ForegroundColor White
    Write-Host "  .\Test-GraphConnection.ps1" -ForegroundColor White
    Write-Host ""
    Write-Host "Note: If you see warnings above, the scripts may still work." -ForegroundColor Gray
    Write-Host "      Try running a script and see if it succeeds." -ForegroundColor Gray
    Write-Host ""
}
catch {
    Write-Host "  ERROR: Connection validation failed" -ForegroundColor Red
    Write-Host "  $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please try running this script again." -ForegroundColor Yellow
    Write-Host ""
    Exit 1
}
