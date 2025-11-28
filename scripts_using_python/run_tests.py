#!/usr/bin/env python3
"""
Test Runner Script

Runs the complete test suite for Outlook Graph Automation (Python).

Usage:
    python run_tests.py              # Run all tests
    python run_tests.py -v           # Verbose output
    python run_tests.py -k test_name # Run specific test
    python run_tests.py --cov        # Run with coverage report

@author: Generated for outlook_automation repository
"""

import sys
import pytest


def main():
    """
    Run pytest with custom configuration.
    """
    print("=" * 70)
    print("Outlook Graph Automation - Test Suite")
    print("=" * 70)
    print()

    # Default pytest arguments
    args = [
        "tests/",           # Test directory
        "-v",               # Verbose
        "--tb=short",       # Short traceback format
        "--color=yes",      # Colored output
    ]

    # Add any command-line arguments passed to this script
    if len(sys.argv) > 1:
        args.extend(sys.argv[1:])

    print(f"Running pytest with args: {' '.join(args)}")
    print()

    # Run pytest
    exit_code = pytest.main(args)

    print()
    print("=" * 70)
    if exit_code == 0:
        print("✓ All tests passed!")
    else:
        print(f"✗ Tests failed with exit code: {exit_code}")
    print("=" * 70)

    return exit_code


if __name__ == "__main__":
    sys.exit(main())
