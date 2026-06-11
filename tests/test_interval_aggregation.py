"""Formula contract tests for the WFM interval aggregation engine.

These tests define the expected spreadsheet behavior before the workbook
generator exists. The implementation should mirror these rules in Excel
formulas and any Python verification helpers.
"""

from datetime import datetime
import unittest

from scripts.interval_engine import (
    interval_productive_fraction,
    row_interval_value,
)


PRODUCTIVE_CODES = {"OWD", "OT"}
NONPRODUCTIVE_CODES = {"BRK", "LCH", "MT", "TRN"}
INACTIVE_CODES = {"OFF", "AL", "SL"}


def dt(value: str) -> datetime:
    return datetime.strptime(value, "%Y-%m-%d %H:%M")


class IntervalAggregationTests(unittest.TestCase):
    def test_productive_fraction_prorates_by_overlap_minutes(self) -> None:
        cases = [
            ("2026-06-11 08:15", "2026-06-11 12:00", "2026-06-11 08:00", "2026-06-11 08:30", 0.5),
            ("2026-06-11 08:15", "2026-06-11 12:00", "2026-06-11 08:30", "2026-06-11 09:00", 1.0),
            ("2026-06-11 08:15", "2026-06-11 12:00", "2026-06-11 07:30", "2026-06-11 08:00", 0.0),
            ("2026-06-11 08:15", "2026-06-11 12:00", "2026-06-11 11:30", "2026-06-11 12:00", 1.0),
            ("2026-06-11 08:15", "2026-06-11 12:00", "2026-06-11 12:00", "2026-06-11 12:30", 0.0),
            ("2026-06-11 23:45", "2026-06-12 02:15", "2026-06-11 23:30", "2026-06-12 00:00", 0.5),
            ("2026-06-11 23:45", "2026-06-12 02:15", "2026-06-12 00:00", "2026-06-12 00:30", 1.0),
            ("2026-06-11 23:45", "2026-06-12 02:15", "2026-06-12 02:00", "2026-06-12 02:30", 0.5),
        ]

        for shift_start, shift_end, interval_start, interval_end, expected in cases:
            with self.subTest(interval_start=interval_start):
                self.assertEqual(
                    interval_productive_fraction(
                        dt(shift_start),
                        dt(shift_end),
                        dt(interval_start),
                        dt(interval_end),
                    ),
                    expected,
                )

    def test_inactive_and_nonproductive_codes_count_as_zero(self) -> None:
        for code in ["OFF", "AL", "SL", "Brk", "Lch", "Mt", "Trn", ""]:
            with self.subTest(code=code):
                self.assertEqual(
                    row_interval_value(
                        employee_id="E001",
                        activity_code=code,
                        shift_start=dt("2026-06-11 08:00"),
                        shift_end=dt("2026-06-11 17:00"),
                        interval_start=dt("2026-06-11 09:00"),
                        interval_end=dt("2026-06-11 09:30"),
                        productive_codes=PRODUCTIVE_CODES,
                        nonproductive_codes=NONPRODUCTIVE_CODES,
                        inactive_codes=INACTIVE_CODES,
                    ),
                    0.0,
                )

    def test_blank_or_null_roster_fields_return_zero_without_errors(self) -> None:
        cases = [
            ("", "OWD", dt("2026-06-11 08:00"), dt("2026-06-11 17:00")),
            (None, "OWD", dt("2026-06-11 08:00"), dt("2026-06-11 17:00")),
            ("E001", "", dt("2026-06-11 08:00"), dt("2026-06-11 17:00")),
            ("E001", None, dt("2026-06-11 08:00"), dt("2026-06-11 17:00")),
            ("E001", "OWD", None, dt("2026-06-11 17:00")),
            ("E001", "OWD", dt("2026-06-11 08:00"), None),
        ]

        for employee_id, activity_code, shift_start, shift_end in cases:
            with self.subTest(employee_id=employee_id, activity_code=activity_code):
                self.assertEqual(
                    row_interval_value(
                        employee_id=employee_id,
                        activity_code=activity_code,
                        shift_start=shift_start,
                        shift_end=shift_end,
                        interval_start=dt("2026-06-11 09:00"),
                        interval_end=dt("2026-06-11 09:30"),
                        productive_codes=PRODUCTIVE_CODES,
                        nonproductive_codes=NONPRODUCTIVE_CODES,
                        inactive_codes=INACTIVE_CODES,
                    ),
                    0.0,
                )

    def test_productive_codes_are_case_insensitive_for_schedule_entry(self) -> None:
        self.assertEqual(
            row_interval_value(
                employee_id="E001",
                activity_code="owd",
                shift_start=dt("2026-06-11 08:15"),
                shift_end=dt("2026-06-11 17:00"),
                interval_start=dt("2026-06-11 08:00"),
                interval_end=dt("2026-06-11 08:30"),
                productive_codes=PRODUCTIVE_CODES,
                nonproductive_codes=NONPRODUCTIVE_CODES,
                inactive_codes=INACTIVE_CODES,
            ),
            0.5,
        )


if __name__ == "__main__":
    unittest.main()

