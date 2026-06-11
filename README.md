# Excel WFM Scheduling Matrix

Editable Microsoft Excel workforce management scheduling workbook for daily interval staffing control. The workbook renders one selected operational day at a time, maps employee schedules across 48 half-hour intervals, and compares scheduled productive headcount against required staffing by interval.

The generated workbook is located at:

```text
excel schedule/wfm_scheduling_matrix.xlsx
```

## Current Scope

Phase 1 through Phase 3 are implemented:

- Workbook tab structure and configurable activity-code tables.
- Normalized staffing requirements by operational date and 30-minute interval.
- Hidden calculation engine for selected-day schedule rows.
- Live `Schedule_Matrix` grid with interval formulas from `00:00` through `23:30`.
- Scheduled productive, required, and over/under summary rows.
- `Summary_Dashboard` source formulas and net staffing variance chart.
- Python verification tests for interval aggregation and workbook structure.

Phase 4 and Phase 5 are still tracked for follow-up:

- Phase 4: visible workbook validation states for invalid code, duplicate row, zero-length shift, missing selected date, and complete conditional-format states.
- Phase 5: workbook-native `TestCases` formulas with an overall pass/fail status.

## Workbook Tabs

| Tab | Purpose |
| --- | --- |
| `Config_Settings` | Activity codes, staffed-count behavior, parameters, and threshold colors. |
| `Staffing_Requirements` | Required headcount by `OperationalDate` and half-hour `IntervalStart`. |
| `Schedule_Data` | Normalized schedule records that feed the daily matrix. |
| `Calc_Engine` | Hidden formula layer for selected-day filtering and interval coverage math. |
| `Schedule_Matrix` | Main planner view with metadata columns and 48 interval columns. |
| `Summary_Dashboard` | Chart-ready interval summary for selected-day net staffing variance. |
| `TestCases` | Workbook test harness placeholder for Phase 5. |

## Key Sheet Anchors

| Area | Anchor |
| --- | --- |
| Selected date | `Schedule_Matrix!E2` |
| Metadata columns | `Schedule_Matrix!A:G` |
| Interval headers | `Schedule_Matrix!H1:BC1` |
| Roster rows | `Schedule_Matrix!A8:BC257` |
| Scheduled productive summary | `Schedule_Matrix!H260:BC260` |
| Required summary | `Schedule_Matrix!H261:BC261` |
| Over/under summary | `Schedule_Matrix!H262:BC262` |
| Activity-code table | `Config_Settings!A1:H40` |
| Requirement table | `Staffing_Requirements!A1:C337` |
| Schedule-data table | `Schedule_Data!A1:K1000` |

## Activity Codes

Activity behavior is configured in `Config_Settings`, not hardcoded into the visible matrix.

| Code | Default Behavior |
| --- | --- |
| `OWD` | Productive, counts as staffed. |
| `OT` | Productive, counts as staffed. |
| `Brk` | Paid nonproductive, visible in grid, does not count as staffed. |
| `Lch` | Meal break, visible in grid, does not count as staffed. |
| `Mt` | Meeting, visible in grid, does not count as staffed. |
| `Trn` | Training, visible in grid, does not count as staffed. |
| `PTO` | Absence, does not count as staffed. |
| `Sick` | Absence, does not count as staffed. |
| `Off` | Off state, does not count as staffed. |
| `UA` | Unapproved absence, does not count as staffed. |
| `AL` | Annual leave, does not count as staffed. |
| `SL` | Sick leave short code, does not count as staffed. |

## Step-by-Step: Use the Workbook

1. Open `excel schedule/wfm_scheduling_matrix.xlsx` in Microsoft Excel or upload it to Google Sheets.
2. Go to `Schedule_Matrix`.
3. Set the selected operational date in `Schedule_Matrix!E2`.
4. Go to `Staffing_Requirements`.
5. Enter one row per required interval using `OperationalDate`, `IntervalStart`, and `RequiredHeadcount`.
6. Go to `Schedule_Data`.
7. Add or edit employee schedule records with employee metadata, operational date, shift start, shift end, and activity code.
8. Return to `Schedule_Matrix`.
9. Review the interval grid from `H:BC`.
10. Review scheduled productive, required, and over/under rows at `H260:BC262`.
11. Open `Summary_Dashboard` to inspect the selected-day net staffing variance chart.

## Step-by-Step: Regenerate the Workbook

Run these commands from the repository root:

```bash
python3 scripts/build_wfm_workbook.py
```

If the default `python3` does not have `openpyxl`, use the bundled Codex runtime:

```bash
/Users/jonasodones/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/build_wfm_workbook.py
```

This rewrites:

```text
excel schedule/wfm_scheduling_matrix.xlsx
```

Use regeneration when changing workbook structure, formulas, activity defaults, charts, conditional formatting, named ranges, or seeded sample data.

## Step-by-Step: Verify the Workbook

Run the unit tests:

```bash
python3 -m unittest
```

Regenerate the workbook:

```bash
python3 scripts/build_wfm_workbook.py
```

Run the structural workbook verifier:

```bash
python3 scripts/verify_wfm_workbook.py
```

If `openpyxl` is missing from the default `python3`, run the same checks with:

```bash
/Users/jonasodones/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m unittest
/Users/jonasodones/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/verify_wfm_workbook.py
```

Expected verifier output:

```text
WFM workbook verification passed
```

## Step-by-Step: Edit Formula Logic Safely

1. Update formula-generation code in `scripts/build_wfm_workbook.py`.
2. Add or update Python test coverage in `tests/test_interval_aggregation.py`.
3. Regenerate the workbook.
4. Run `python3 -m unittest`.
5. Run `python3 scripts/verify_wfm_workbook.py`.
6. Open the workbook and spot-check `Schedule_Matrix`, `Calc_Engine`, and `Summary_Dashboard`.
7. Commit both the script changes and the regenerated `.xlsx` file.

## Step-by-Step: Push Changes

Check the current branch and status:

```bash
git branch --show-current
git status --short
```

Stage documentation, scripts, tests, and workbook output as needed:

```bash
git add README.md scripts tests "excel schedule/wfm_scheduling_matrix.xlsx"
```

Commit:

```bash
git commit -m "docs: add workbook usage guide"
```

Push:

```bash
git push
```

The configured remote is:

```text
git@github.com:jonas230ph/excel_scheduled_file.git
```

## Formula Design Notes

- The visible matrix displays activity codes for intervals where a schedule overlaps the interval.
- The hidden `Calc_Engine` converts each row/interval into numeric staffed coverage.
- Productive headcount is driven by `tblActivityCodes[CountsAsStaffed]`.
- Partial intervals are prorated. For example, a shift starting at `08:15` contributes `0.5` to the `08:00-08:30` interval.
- Overnight shifts are handled by normalizing end datetimes that cross midnight.
- Missing requirements return `MissingRequirement` instead of silently calculating demand as zero.
- Core formulas avoid volatile functions such as `OFFSET` and `INDIRECT`.

## Named Ranges

| Name | Refers To |
| --- | --- |
| `SelectedDate` | `Schedule_Matrix!$E$2` |
| `IntervalHeaders` | `Schedule_Matrix!$H$1:$BC$1` |
| `ActiveActivityCodes` | `Config_Settings!$A$2:$A$13` |
| `tblDailyMatrix` | `Schedule_Matrix!$A$1:$BC$257` |

## Known Follow-Up Work

Complete these before treating the workbook as a full Phase 5 release:

1. Add visible validation states for invalid activity codes.
2. Add visible validation states for zero-length shifts.
3. Add duplicate schedule-row detection.
4. Add missing selected-date handling.
5. Expand conditional formatting for all validation states.
6. Replace placeholder `TestCases` rows with workbook-native expected-vs-actual formulas.
7. Add a `TestCases` overall pass/fail cell.
