# Regression Tests for Show-MeetingHourSummary.ps1

## Overview

This directory contains comprehensive regression tests for the Meeting Hour Summary script using the Pester testing framework. The tests ensure the script's core functionality remains stable across changes.

## Test Coverage

### Test File: `Show-MeetingHourSummary.Tests.ps1`

The test suite includes **41 test cases** covering:

#### 1. **Get-WeekdayBounds Function** (9 tests)
- Monday, Wednesday, Sunday, and Saturday date handling
- Month boundary crossing
- Year boundary crossing
- Leap year support
- DST (Daylight Saving Time) transitions
- Weekend handling (returns upcoming work week)

#### 2. **Get-AppointmentDuration Function** (6 tests)
- Standard 1-hour meetings
- 30-minute meetings
- Multi-hour meetings
- Odd durations with rounding
- Meetings spanning midnight
- Decimal precision (2 decimal places)

#### 3. **Test-ShouldIgnoreAppointment Function** (9 tests)
- Exact pattern matching
- Wildcard patterns
- Multiple patterns
- Special regex characters
- Case-sensitive matching (default)
- Case-insensitive patterns (using `(?i)`)
- Empty pattern lists
- Partial content matching

#### 4. **Load-IgnorePatterns Function** (6 tests)
- Loading patterns from file
- Ignoring comment lines (starting with `#`)
- Ignoring empty lines
- Trimming whitespace
- Handling missing files
- Handling files with only comments

#### 5. **Edge Cases** (3 tests)
- Leap year calculations
- DST spring forward
- DST fall back

#### 6. **Integration Scenarios** (3 tests)
- Filtering multiple appointments
- Calculating total hours across meetings
- Back-to-back meeting handling

#### 7. **Regression Tests** (5 tests)
- Special characters in meeting titles
- Unicode characters
- Very long meeting titles (500+ characters)
- First/last day of year calculations

## Prerequisites

### Pester Installation

The tests require the Pester testing framework. Most Windows systems include Pester v3.4.0+ by default.

To check your Pester version:
```powershell
Get-Module -ListAvailable -Name Pester
```

To install/upgrade Pester (optional - tests work with v3+):
```powershell
Install-Module -Name Pester -Force -SkipPublisherCheck
```

## Running the Tests

### Quick Run (All Tests)

From the `scripts` directory:
```powershell
cd C:\path\to\outlook_automation\scripts
Invoke-Pester -Path .\Show-MeetingHourSummary.Tests.ps1
```

### Verbose Output

To see detailed test results:
```powershell
Invoke-Pester -Path .\Show-MeetingHourSummary.Tests.ps1 -Verbose
```

### Generate XML Report

To create an NUnit-compatible XML report for CI/CD:
```powershell
Invoke-Pester -Path .\Show-MeetingHourSummary.Tests.ps1 -OutputFormat NUnitXml -OutputFile TestResults.xml
```

### Run Specific Tests

To run tests for a specific function:
```powershell
Invoke-Pester -Path .\Show-MeetingHourSummary.Tests.ps1 -TestName "*WeekdayBounds*"
```

## Expected Output

Successful test run output:
```
Describing Get-WeekdayBounds
   Context When given a Monday
    [+] Should return the same Monday 740ms
    [+] Should return Friday 4 days later 109ms
   ...

Tests completed in 3.04s
Passed: 41 Failed: 0 Skipped: 0 Pending: 0 Inconclusive: 0
```

## Test Maintenance

### Adding New Tests

When adding new functionality to the main script:

1. Add corresponding tests to `Show-MeetingHourSummary.Tests.ps1`
2. Follow the existing Describe/Context/It structure
3. Ensure tests are compatible with Pester v3+ syntax
4. Run all tests to ensure no regressions

### Test Structure

```powershell
Describe "FunctionName" {
    Context "When scenario description" {
        It "Should expected behavior" {
            # Arrange
            $input = "test data"

            # Act
            $result = FunctionName -Parameter $input

            # Assert
            $result | Should Be "expected value"
        }
    }
}
```

### Pester v3 Compatibility Notes

The tests use Pester v3-compatible syntax:
- `Should Be` instead of `Should -Be`
- `Should BeNullOrEmpty` instead of `Should -BeNullOrEmpty`
- Script-level function definitions instead of `BeforeAll` blocks
- `-contains` operator for array membership tests

## Continuous Integration

### Automated Testing

To integrate with CI/CD pipelines:

```powershell
# Run tests and capture results
$results = Invoke-Pester -Path .\Show-MeetingHourSummary.Tests.ps1 -PassThru

# Exit with error code if tests fail
if ($results.FailedCount -gt 0) {
    Write-Error "$($results.FailedCount) tests failed"
    exit 1
}
```

### Pre-commit Hook

Create `.git/hooks/pre-commit`:
```bash
#!/bin/bash
cd scripts
powershell.exe -Command "Invoke-Pester -Path .\Show-MeetingHourSummary.Tests.ps1 -Quiet"
if [ $? -ne 0 ]; then
    echo "Tests failed. Commit aborted."
    exit 1
fi
```

## Troubleshooting

### Tests Fail with "Should operator not found"

You may be using Pester v5 syntax in a v3 environment. Ensure assertions use:
- `Should Be` (not `Should -Be`)
- `Should BeNullOrEmpty` (not `Should -BeNullOrEmpty`)

### Temporary Directory Errors

Tests create a temporary directory for file operations. If tests fail with directory errors:
```powershell
# Clean up temp directories manually
Remove-Item "$env:TEMP\MeetingHourSummaryTests_*" -Recurse -Force -ErrorAction SilentlyContinue
```

### Date/Time Sensitive Tests

Some tests use specific dates (2025-11-10, etc.). These tests:
- Are timezone-independent
- Test relative date calculations, not absolute dates
- Should pass regardless of when they're run

## Test Coverage Report

Current coverage: **All critical functions tested**

| Function | Tests | Coverage |
|----------|-------|----------|
| Get-WeekdayBounds | 9 | ✅ Full |
| Get-AppointmentDuration | 6 | ✅ Full |
| Test-ShouldIgnoreAppointment | 9 | ✅ Full |
| Load-IgnorePatterns | 6 | ✅ Full |
| Edge Cases | 11 | ✅ Full |

## Known Limitations

- **No Outlook COM object testing**: Tests use mock functions for the core logic. Actual Outlook COM interaction is not tested due to:
  - Dependency on installed Outlook
  - Requirement for configured mailbox
  - Complexity of mocking COM objects

- **No UI testing**: Windows Forms popup is not tested automatically

## Contributing

When contributing changes:

1. Ensure all existing tests pass
2. Add tests for new functionality
3. Maintain Pester v3 compatibility
4. Update this README with new test descriptions
5. Keep test execution time under 5 seconds

## Version History

- **v1.0** - Initial test suite with 41 comprehensive tests
  - Core function testing
  - Edge case coverage
  - Integration scenarios
  - Regression test suite
