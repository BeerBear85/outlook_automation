# Test Runner for Show-MeetingHourSummary.ps1
# Convenient wrapper for running Pester tests
#
# @author: Generated for outlook_automation repository
# -----------------------------------------------------------------------------

<#
.SYNOPSIS
    Runs the regression tests for Show-MeetingHourSummary.ps1

.DESCRIPTION
    This script provides a convenient way to run the Pester test suite for the
    Meeting Hour Summary script. It supports various output formats and options.

.PARAMETER Detailed
    Show detailed test output with timing information

.PARAMETER Quiet
    Show only pass/fail summary without individual test results

.PARAMETER CreateReport
    Generate an XML test report (TestResults.xml)

.PARAMETER TestName
    Run only tests matching the specified name pattern

.EXAMPLE
    .\Run-Tests.ps1
    Runs all tests with standard output

.EXAMPLE
    .\Run-Tests.ps1 -Detailed
    Runs all tests with detailed timing information

.EXAMPLE
    .\Run-Tests.ps1 -Quiet
    Runs all tests showing only the summary

.EXAMPLE
    .\Run-Tests.ps1 -CreateReport
    Runs all tests and generates TestResults.xml

.EXAMPLE
    .\Run-Tests.ps1 -TestName "*WeekdayBounds*"
    Runs only tests matching the pattern
#>

param (
    [switch]$Detailed,
    [switch]$Quiet,
    [switch]$CreateReport,
    [string]$TestName
)

# Ensure we're in the correct directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $scriptDir

try {
    # Check if Pester is available
    $pesterModule = Get-Module -ListAvailable -Name Pester
    if (-not $pesterModule) {
        Write-Error "Pester module not found. Please install it with: Install-Module -Name Pester -Force"
        exit 1
    }

    Write-Host "Using Pester version: $($pesterModule.Version)" -ForegroundColor Cyan
    Write-Host ""

    # Build Pester parameters
    $pesterParams = @{
        Path = ".\Show-MeetingHourSummary.Tests.ps1"
        PassThru = $true
    }

    if ($Detailed) {
        $pesterParams['Verbose'] = $true
    }

    if ($Quiet) {
        $pesterParams['Quiet'] = $true
    }

    if ($CreateReport) {
        $pesterParams['OutputFormat'] = 'NUnitXml'
        $pesterParams['OutputFile'] = 'TestResults.xml'
    }

    if ($TestName) {
        $pesterParams['TestName'] = $TestName
    }

    # Run the tests
    Write-Host "Running tests..." -ForegroundColor Yellow
    Write-Host "================================================================================`n" -ForegroundColor Gray

    $results = Invoke-Pester @pesterParams

    # Display summary
    Write-Host "`n================================================================================" -ForegroundColor Gray
    Write-Host "Test Summary" -ForegroundColor Cyan
    Write-Host "================================================================================`n" -ForegroundColor Gray

    $passColor = if ($results.PassedCount -eq $results.TotalCount) { "Green" } else { "Yellow" }
    $failColor = if ($results.FailedCount -eq 0) { "Green" } else { "Red" }

    Write-Host "Total Tests:   $($results.TotalCount)" -ForegroundColor White
    Write-Host "Passed:        $($results.PassedCount)" -ForegroundColor $passColor
    Write-Host "Failed:        $($results.FailedCount)" -ForegroundColor $failColor
    Write-Host "Skipped:       $($results.SkippedCount)" -ForegroundColor Gray
    Write-Host "Duration:      $($results.Time.ToString())" -ForegroundColor White

    if ($CreateReport) {
        Write-Host "`nReport saved to: TestResults.xml" -ForegroundColor Cyan
    }

    # Exit with appropriate code
    if ($results.FailedCount -gt 0) {
        Write-Host "`nTests FAILED!" -ForegroundColor Red
        exit 1
    } else {
        Write-Host "`nAll tests PASSED!" -ForegroundColor Green
        exit 0
    }
} finally {
    Pop-Location
}
