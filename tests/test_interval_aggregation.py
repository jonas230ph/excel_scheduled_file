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
from scripts.build_wfm_workbook import (
    calc_coverage_formula,
    matrix_visible_formula,
    requirement_formula,
    schedule_validation_formula,
    selected_schedule_formula,
    selected_day_variance_formula,
    weekly_scheduled_formula,
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

    def test_generated_coverage_formula_uses_import_safe_staffed_lookup(self) -> None:
        formula = calc_coverage_formula(8, "X")

        self.assertNotIn("XLOOKUP(code,UPPER(", formula)
        self.assertNotIn("SelectedDate", formula)
        self.assertIn("INDEX(tblActivityCodes[CountsAsStaffed]", formula)
        self.assertIn("MATCH(code,tblActivityCodes[Code],0)", formula)
        self.assertIn("Schedule_Matrix!$E$2", formula)

    def test_generated_requirement_formula_distinguishes_missing_from_zero(self) -> None:
        formula = requirement_formula("X", 1)

        self.assertIn("COUNTIFS(tblRequirements[OperationalDate],$E$2", formula)
        self.assertIn('=0,""', formula)
        self.assertIn("SUMIFS(tblRequirements[RequiredHeadcount]", formula)

    def test_generated_schedule_validation_formula_flags_phase4_states(self) -> None:
        formula = schedule_validation_formula(42)

        self.assertIn('COUNTA($A42:$F42)=0,""', formula)
        self.assertIn('$C42="","MissingDate"', formula)
        self.assertIn('$D42=$E42),"InvalidSchedule"', formula)
        self.assertIn('COUNTIF(ActiveActivityCodes,$F42)=0,"InvalidCode"', formula)
        self.assertIn('COUNTIFS(tblScheduleData[EmployeeID],$A42', formula)
        self.assertIn('>1,"Duplicate","Valid"', formula)

    def test_generated_variance_formula_flags_blank_selected_date_first(self) -> None:
        formula = selected_day_variance_formula("X", 260, 261)

        self.assertEqual(
            formula,
            '=IF($E$2="","MissingDate",IF(X$261="","MissingRequirement",ROUND(X$260-X$261,2)))',
        )

    def test_selected_day_schedule_filter_includes_previous_day_overnight_overlap(self) -> None:
        formula = selected_schedule_formula("EmployeeID", 8)

        self.assertIn("Schedule_Data!$A$2:$A$1000", formula)
        self.assertIn("Schedule_Data!$C$2:$C$1000", formula)
        self.assertNotIn("tblScheduleData[OperationalDate]=SelectedDate", formula)
        self.assertNotIn("SelectedDate", formula)
        self.assertNotIn("LET(", formula)
        self.assertIn("Schedule_Matrix!$E$2", formula)
        self.assertIn("(INT(Schedule_Data!$D$2:$D$1000)=0)*Schedule_Data!$C$2:$C$1000", formula)
        self.assertIn("(INT(Schedule_Data!$E$2:$E$1000)=0)*Schedule_Data!$C$2:$C$1000", formula)
        self.assertIn("<Schedule_Matrix!$E$2+1", formula)
        self.assertIn(">Schedule_Matrix!$E$2", formula)
        self.assertIn("Schedule_Data!$A$2:$A$1000<>\"\"", formula)

    def test_visible_and_coverage_formulas_include_break_lunch_and_adhoc_windows(self) -> None:
        visible_formula = matrix_visible_formula(8, "X")
        coverage_formula = calc_coverage_formula(8, "X")

        for field in [
            "Break1Start",
            "Break2Start",
            "LunchStart",
            "Adhoc1Code",
            "Adhoc2Code",
            "Adhoc3Code",
        ]:
            self.assertIn(field, visible_formula)
            self.assertIn(field, coverage_formula)
        self.assertIn('"BRK"', visible_formula)
        self.assertIn('"LCH"', visible_formula)
        self.assertIn("overlayCoverage", coverage_formula)

    def test_generated_exception_formulas_remain_nonvolatile_and_within_excel_limit(self) -> None:
        formulas = [
            matrix_visible_formula(8, "X"),
            calc_coverage_formula(8, "X"),
            selected_schedule_formula("EmployeeID", 8),
            weekly_scheduled_formula("A4"),
        ]

        for formula in formulas:
            with self.subTest(length=len(formula)):
                self.assertLess(len(formula), 8192)
                for volatile_function in ["OFFSET(", "INDIRECT(", "TODAY(", "NOW(", "RAND(", "RANDBETWEEN("]:
                    self.assertNotIn(volatile_function, formula.upper())

    def test_weekly_dashboard_formula_sums_break_lunch_and_adhoc_adjusted_coverage(self) -> None:
        formula = weekly_scheduled_formula("A4")

        self.assertEqual(
            formula,
            '=IF(A4=Schedule_Matrix!$E$2,ROUND(SUM(Schedule_Matrix!$H$260:$BC$260),2),0)',
        )
        self.assertIn("Schedule_Matrix!$H$260:$BC$260", formula)
        self.assertNotIn("SUMPRODUCT(", formula)
        self.assertNotIn("MMULT(", formula)

    def test_selected_schedule_formula_avoids_filter_on_calc_engine(self) -> None:
        formula = selected_schedule_formula("EmployeeID", 8)

        self.assertTrue(formula.startswith("=IFERROR(INDEX(Schedule_Data!$A$2:$A$1000,AGGREGATE("))
        self.assertNotIn("FILTER(", formula)
        self.assertIn("AGGREGATE(15,6,", formula)
        self.assertNotIn("tblScheduleData[", formula)

    def test_google_selected_schedule_formula_uses_filter(self) -> None:
        formula = selected_schedule_formula("EmployeeID", 8, "google")

        self.assertTrue(formula.startswith("=IFERROR(INDEX(FILTER(Schedule_Data!$A$2:$A$1000,"))
        self.assertIn("FILTER(Schedule_Data!$A$2:$A$1000", formula)
        self.assertNotIn("AGGREGATE(", formula)

    def test_generated_exception_formulas_ignore_boundary_precision_noise(self) -> None:
        visible_formula = matrix_visible_formula(8, "X")
        coverage_formula = calc_coverage_formula(8, "X")

        self.assertIn("RawOverlap", visible_formula)
        self.assertIn("RawOverlap", coverage_formula)
        self.assertIn("baseRawOverlap", visible_formula)
        self.assertIn("baseRawOverlap", coverage_formula)
        self.assertIn("TIME(0,0,1)", visible_formula)
        self.assertIn("TIME(0,0,1)", coverage_formula)

    def test_break_lunch_and_adhoc_overlay_expected_staffed_values(self) -> None:
        shift_start = dt("2026-06-11 08:15")
        shift_end = dt("2026-06-11 17:00")
        cases = [
            ("Break 1 subtracts 15 minutes", "2026-06-11 10:00", "2026-06-11 10:30", [("2026-06-11 10:00", "2026-06-11 10:15")], 0.5),
            ("Lunch subtracts full interval", "2026-06-11 12:00", "2026-06-11 12:30", [("2026-06-11 12:00", "2026-06-11 12:30")], 0.0),
            ("Ad hoc meeting subtracts full interval", "2026-06-11 14:00", "2026-06-11 14:30", [("2026-06-11 14:00", "2026-06-11 14:30")], 0.0),
            ("Break 2 subtracts 15 minutes", "2026-06-11 15:00", "2026-06-11 15:30", [("2026-06-11 15:00", "2026-06-11 15:15")], 0.5),
        ]

        for label, interval_start, interval_end, overlays, expected in cases:
            with self.subTest(label=label):
                interval_start_dt = dt(interval_start)
                interval_end_dt = dt(interval_end)
                base = interval_productive_fraction(shift_start, shift_end, interval_start_dt, interval_end_dt)
                overlay = sum(
                    interval_productive_fraction(dt(start), dt(end), interval_start_dt, interval_end_dt)
                    for start, end in overlays
                )
                self.assertEqual(max(0.0, base - overlay), expected)


if __name__ == "__main__":
    unittest.main()
