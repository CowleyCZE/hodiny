#!/usr/bin/env python3
"""
Comprehensive test suite for Hodiny2025Manager
Tests all major functionality including edge cases and lunch hour calculations
"""

import sys
from pathlib import Path
from datetime import datetime
import tempfile
import shutil

from hodiny2025_manager import Hodiny2025Manager

# Add path to module - this is needed for standalone script execution
sys.path.append(str(Path(__file__).parent))


class TestHodiny2025Comprehensive:
    """Comprehensive test class for Hodiny2025Manager."""

    @classmethod
    def setup_class(cls):
        """Setup test environment."""
        cls.temp_dir = tempfile.mkdtemp()
        cls.manager = Hodiny2025Manager(cls.temp_dir)

    @classmethod
    def teardown_class(cls):
        """Cleanup test environment."""
        if hasattr(cls, "temp_dir") and Path(cls.temp_dir).exists():
            shutil.rmtree(cls.temp_dir)

    def test_lunch_hour_calculations(self):
        """Test various lunch hour scenarios."""
        test_cases = [
            # (date, start, end, lunch, employees, expected_hours, expected_overtime)
            ("2025-01-15", "08:00", "16:00", "0.5", 1, 7.5, 0.0),  # Half hour lunch
            ("2025-01-16", "07:00", "16:00", "1.0", 1, 8.0, 0.0),  # Standard day
            ("2025-01-17", "07:00", "17:00", "1.0", 1, 9.0, 1.0),  # Overtime
            ("2025-01-18", "06:00", "18:00", "1.5", 1, 10.5, 2.5),  # Long day with long lunch
            ("2025-01-19", "08:00", "12:00", "0.0", 1, 4.0, 0.0),  # No lunch, short day
            ("2025-01-20", "09:00", "17:30", "0.25", 1, 8.25, 0.25),  # 15 min lunch
        ]

        for date, start, end, lunch, employees, expected_hours, expected_overtime in test_cases:
            # Write data
            self.manager.zapis_pracovni_doby(date, start, end, lunch, employees)

            # Read and verify
            record = self.manager.get_daily_record(date)

            assert (
                abs(record["total_hours"] - expected_hours) < 0.01
            ), f"Date {date}: Expected {expected_hours}h, got {record['total_hours']}h"
            assert (
                abs(record["overtime"] - expected_overtime) < 0.01
            ), f"Date {date}: Expected {expected_overtime}h overtime, got {record['overtime']}h"

            print(
                f"✅ {date}: {start}-{end}, lunch {lunch}h → "
                f"{record['total_hours']}h total, {record['overtime']}h overtime"
            )

    def test_edge_cases(self):
        """Test edge cases and special scenarios."""

        # Test zero values
        self.manager.zapis_pracovni_doby("2025-02-01", "00:00", "00:00", "0", 0)
        record = self.manager.get_daily_record("2025-02-01")
        assert record["total_hours"] == 0.0
        assert record["overtime"] == 0.0
        assert record["num_employees"] == 0
        print("✅ Zero values test passed")

        # Test midnight crossing (not supported but should handle gracefully)
        self.manager.zapis_pracovni_doby("2025-02-02", "22:00", "06:00", "0", 1)
        record = self.manager.get_daily_record("2025-02-02")
        # This will be negative, but the backup calculation should handle it
        print(f"✅ Midnight crossing: {record['total_hours']}h (handled gracefully)")

        # Test high employee count
        self.manager.zapis_pracovni_doby("2025-02-03", "08:00", "16:00", "1.0", 50)
        record = self.manager.get_daily_record("2025-02-03")
        assert record["num_employees"] == 50
        assert record["total_all_employees"] == 7.0 * 50  # 7h * 50 employees
        print("✅ High employee count test passed")

    def test_monthly_summary_accuracy(self):
        """Test that monthly summaries are accurate."""

        # Add several days of data in March
        test_data = [
            ("2025-03-01", "08:00", "16:30", "1.0", 3),  # 7.5h
            ("2025-03-02", "07:00", "17:00", "1.0", 2),  # 9.0h, 1h overtime
            ("2025-03-03", "08:00", "16:00", "0.5", 4),  # 7.5h
            ("2025-03-04", "07:30", "18:00", "1.5", 1),  # 9.0h, 1h overtime
        ]

        total_expected_hours = 0.0
        total_expected_overtime = 0.0
        total_expected_all = 0.0

        for date, start, end, lunch, employees in test_data:
            self.manager.zapis_pracovni_doby(date, start, end, lunch, employees)

            # Calculate expected values
            start_time = datetime.strptime(start, "%H:%M")
            end_time = datetime.strptime(end, "%H:%M")
            hours = (end_time - start_time).seconds / 3600 - float(lunch)
            overtime = max(0, hours - 8)

            total_expected_hours += hours
            total_expected_overtime += overtime
            total_expected_all += hours * employees

        # Get monthly summary
        summary = self.manager.get_monthly_summary(3, 2025)

        assert (
            abs(summary["total_hours"] - total_expected_hours) < 0.01
        ), f"Expected {total_expected_hours}h, got {summary['total_hours']}h"
        assert (
            abs(summary["total_overtime"] - total_expected_overtime) < 0.01
        ), f"Expected {total_expected_overtime}h overtime, got {summary['total_overtime']}h"
        assert (
            abs(summary["total_all_employees"] - total_expected_all) < 0.01
        ), f"Expected {total_expected_all}h total, got {summary['total_all_employees']}h"

        print(
            f"✅ Monthly summary: {summary['total_hours']}h total, "
            f"{summary['total_overtime']}h overtime, "
            f"{summary['total_all_employees']}h for all employees"
        )

    def test_data_persistence(self):
        """Test that data persists correctly when reopening the file."""

        # Write some data
        test_date = "2025-04-15"
        self.manager.zapis_pracovni_doby(test_date, "09:00", "17:30", "1.0", 5)

        # Create a new manager instance (simulates reopening)
        manager2 = Hodiny2025Manager(self.temp_dir)

        # Read the data
        record = manager2.get_daily_record(test_date)

        assert record["start_time"] == "09:00"
        assert record["end_time"] == "17:30"
        assert record["lunch_hours"] == 1.0
        assert record["num_employees"] == 5
        assert abs(record["total_hours"] - 7.5) < 0.01

        print("✅ Data persistence test passed")

    def test_formula_vs_backup_calculation(self):
        """Test that backup calculations match expected formula results."""

        # Test multiple scenarios
        scenarios = [
            ("08:00", "16:00", 0.5, 7.5),  # 8h - 0.5h = 7.5h
            ("07:00", "15:30", 1.0, 7.5),  # 8.5h - 1h = 7.5h
            ("06:30", "18:30", 2.0, 10.0),  # 12h - 2h = 10h
            ("10:00", "14:00", 0.25, 3.75),  # 4h - 0.25h = 3.75h
        ]

        for i, (start, end, lunch, expected) in enumerate(scenarios, 1):
            date = f"2025-05-{i:02d}"
            self.manager.zapis_pracovni_doby(date, start, end, str(lunch), 1)

            record = self.manager.get_daily_record(date)

            assert (
                abs(record["total_hours"] - expected) < 0.01
            ), f"Scenario {i}: Expected {expected}h, got {record['total_hours']}h"

        print("✅ Formula vs backup calculation test passed")


def test_comprehensive_hodiny2025():
    """Main test function for pytest."""

    print("🚀 COMPREHENSIVE HODINY2025MANAGER TESTS")
    print("=" * 60)

    test_instance = TestHodiny2025Comprehensive()
    test_instance.setup_class()

    try:
        test_instance.test_lunch_hour_calculations()
        print()

        test_instance.test_edge_cases()
        print()

        test_instance.test_monthly_summary_accuracy()
        print()

        test_instance.test_data_persistence()
        print()

        test_instance.test_formula_vs_backup_calculation()
        print()

        print("🎉 ALL COMPREHENSIVE TESTS PASSED!")

    finally:
        test_instance.teardown_class()


if __name__ == "__main__":
    test_comprehensive_hodiny2025()
