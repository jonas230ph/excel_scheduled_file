# Excel WFM Scheduling Matrix

Editable Microsoft Excel workforce management scheduling workbook for daily interval staffing control. The workbook renders one selected operational day at a time, maps employee schedules across 48 half-hour intervals, and compares scheduled productive headcount against required staffing by interval.

The generated workbooks are located at:

```text
excel schedule/wfm_scheduling_matrix.xlsx
excel schedule/wfm_scheduling_matrix_google.xlsx
```

## Current Scope

Phase 1 through Phase 5 are implemented:

- Workbook tab structure and configurable activity-code tables.
- Normalized staffing requirements by operational date and 30-minute interval.
- Hidden calculation engine for selected-day schedule rows.
- Live `Schedule_Matrix` grid with interval formulas from `00:00` through `23:30`.
- Editable break, lunch, and ad hoc windows in `Schedule_Data`, overlaid into the daily matrix.
- Scheduled productive, required, and over/under summary rows.
- `Summary_Dashboard` weekly staffing summary with selected-day intraday staffing chart.
- Visible validation states for invalid activity codes, zero-length shifts, duplicate rows, blank selected date, and missing requirements.
- Conditional formatting for productive, nonproductive, under, exact, over, invalid, duplicate, missing-date, and missing-requirement states.
- Workbook-native `TestCases` formulas with expected-vs-actual checks and an overall pass/fail status.
- Python verification tests for interval aggregation and workbook structure.
- Separate Excel and Google Sheets workbook files so each engine gets formulas it handles cleanly.

Deferred V2 work remains out of scope for this MVP: forecasting, imports, optimization, macros, integrations, and enterprise workflow controls.

## Workbook Tabs

| Tab | Purpose |
| --- | --- |
| `Config_Settings` | Activity codes, staffed-count behavior, parameters, and threshold colors. |
| `Staffing_Requirements` | Required headcount by `OperationalDate` and half-hour `IntervalStart`. |
| `Schedule_Data` | Normalized schedule records that feed the daily matrix. |
| `Calc_Engine` | Hidden formula layer for selected-day filtering and interval coverage math. |
| `Schedule_Matrix` | Main planner view with metadata columns and 48 interval columns. |
| `Summary_Dashboard` | Weekly staffing summary plus selected-day interval charting. |
| `TestCases` | Workbook-native expected-vs-actual checks and overall pass/fail status. |

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
| Hidden schedule helpers | `Schedule_Matrix!BD:BS` |
| Activity-code table | `Config_Settings!A1:H40` |
| Requirement table | `Staffing_Requirements!A1:C337` |
| Schedule-data table | `Schedule_Data!A1:Z1000` |

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

1. Open `excel schedule/wfm_scheduling_matrix.xlsx` in Microsoft Excel.
2. Go to `Schedule_Matrix`.
3. Set the selected operational date in `Schedule_Matrix!E2`.
4. Go to `Staffing_Requirements`.
5. Enter one row per required interval using `OperationalDate`, `IntervalStart`, and `RequiredHeadcount`.
6. Go to `Schedule_Data`.
7. Add or edit employee schedule records with employee metadata, operational date, shift start, shift end, and activity code.
8. Enter optional scheduled exceptions on the same row:
   - `Break1Start` / `Break1End`
   - `Break2Start` / `Break2End`
   - `LunchStart` / `LunchEnd`
   - `Adhoc1Code` / `Adhoc1Start` / `Adhoc1End`
   - `Adhoc2Code` / `Adhoc2Start` / `Adhoc2End`
   - `Adhoc3Code` / `Adhoc3Start` / `Adhoc3End`
9. Return to `Schedule_Matrix`.
10. Review the interval grid from `H:BC`.
11. Review scheduled productive, required, and over/under rows at `H260:BC262`.
12. Open `Summary_Dashboard` to inspect weekly scheduled vs required staffing and the selected-day intraday staffing chart.

For Google Sheets, open `excel schedule/wfm_scheduling_matrix_google.xlsx`.

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
- The selected-day row filter includes any schedule whose normalized shift window overlaps the selected day, including previous-day overnight carry-in.
- The dashboard week starts on Monday based on the selected date in `Schedule_Matrix!E2`.
- The weekly dashboard's selected-date row uses the verified `Schedule_Matrix` daily total, so break, lunch, ad hoc, partial-shift, and overnight rules are reflected in the visible daily productivity number.
- Break, lunch, and ad hoc windows are edited in `Schedule_Data` and carried through hidden helper columns `Schedule_Matrix!BD:BS`.
- `Schedule_Matrix!H:BC` remains the 48-column interval grid. Breaks and lunch display as `BRK` and `LCH`; ad hoc windows display their configured activity code.
- Break/lunch/ad hoc windows replace the base productive coverage for their overlap minutes, so a 15-minute break inside a 30-minute interval subtracts `0.5` staffed coverage.
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

## Deferred V2 Work

- Erlang-C or equivalent forecasting.
- Bulk roster imports and field mapping.
- Solver-based schedule optimization.
- VBA or Office Script scheduling automation.
- Multi-skill, queue, channel, team, site, or location requirements.
- External WFM, HRIS, payroll, or approval workflow integrations.
