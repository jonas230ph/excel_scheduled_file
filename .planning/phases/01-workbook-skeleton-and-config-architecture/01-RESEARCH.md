# Phase 1 Research: Workbook Skeleton and Config Architecture

**Phase:** 1 - Workbook Skeleton and Config Architecture
**Generated:** 2026-06-11
**Method:** Inline research because GSD phase subagents are unavailable in this runtime.

## RESEARCH COMPLETE

## Implementation Facts

- The target artifact should be a standalone Excel workbook at `excel schedule/wfm_scheduling_matrix.xlsx`.
- Phase 1 should create structure, tables, names, and visible validation scaffolding, not the full staffing calculation engine.
- Python `openpyxl` is appropriate for deterministic workbook generation because it can create sheets, tables, named ranges, formulas, fills, widths, freezes, and charts without requiring Excel GUI automation.
- `Config_Settings`, `Staffing_Requirements`, and `Schedule_Matrix` are the required user-facing tabs.
- Supporting tabs should include `Schedule_Data`, `Calc_Engine`, `TestCases`, and optionally `Summary_Dashboard`.

## Technical Approach

Use a single generation script in `scripts/build_wfm_workbook.py` that:

1. Creates the workbook from scratch.
2. Removes the default empty sheet.
3. Adds sheets in stable order.
4. Writes headers, sample rows, and fixed anchors.
5. Creates Excel tables and named ranges.
6. Adds data validation for active activity codes and selected date.
7. Adds interval headers from `00:00` through `23:30` in `Schedule_Matrix!H1:BC1`.
8. Adds summary rows for scheduled, required, and over/under in `H202:BC204`.
9. Adds conditional formatting placeholders for under/exact/over/invalid/missing/nonproductive states.
10. Saves the workbook to `excel schedule/wfm_scheduling_matrix.xlsx`.

## Anchor Contract

| Sheet | Anchor | Purpose |
|-------|--------|---------|
| `Config_Settings` | `A1:H40` | Activity code rules |
| `Config_Settings` | `J1:K30` | Numeric parameters |
| `Config_Settings` | `M1:Q20` | Threshold colors |
| `Staffing_Requirements` | `A1:C337` | Date + interval requirements |
| `Schedule_Data` | `A1:K1000` | Weekly-normalized schedule source |
| `Schedule_Matrix` | `A1:G200` | Employee metadata and selected-day controls |
| `Schedule_Matrix` | `H1:BC1` | 48 interval headers |
| `Schedule_Matrix` | `H202:BC204` | Scheduled, required, over/under rows |
| `Calc_Engine` | `A1:BC250` | Protected helper logic |
| `TestCases` | `A1:H100` | Expected vs actual workbook tests |

## Footguns

- Avoid hardcoding activity code behavior into matrix formulas. Use `Config_Settings`.
- Avoid volatile functions in core formulas.
- Avoid building a full weekly wall chart in Phase 1.
- Treat missing requirements as missing, not zero.
- Make invalid states visible so staffing totals cannot fail silently.

## Validation Architecture

Phase 1 validation should prove the workbook skeleton exists, not the final WFM math:

- Workbook file exists.
- Required sheets exist.
- `Schedule_Matrix!H1:BC1` contains exactly 48 intervals.
- `Config_Settings` contains activity codes with `CountsAsStaffed` values.
- Named ranges exist for activity codes, parameters, requirements, schedule data, matrix, and test cases.
- Summary rows exist at `H202:BC204`.
- Conditional formatting ranges exist for matrix and summary regions.
