# Validation Strategy: Phase 1 Workbook Skeleton and Config Architecture

**Phase:** 1
**Slug:** workbook-skeleton-and-config-architecture
**Date:** 2026-06-11

## Validation Architecture

Phase 1 is valid when the workbook can be regenerated deterministically and the generated artifact contains the agreed sheet/tab architecture, anchors, named ranges, and validation scaffolding.

## Required Checks

1. Run the workbook generation script.
2. Open the generated workbook with a non-GUI Excel parser.
3. Assert all required sheets exist.
4. Assert `Schedule_Matrix!H1:BC1` contains 48 half-hour labels from `00:00` to `23:30`.
5. Assert `Config_Settings!A1:H1`, `J1:K1`, and `M1:Q1` contain the correct table headers.
6. Assert summary row labels exist at `Schedule_Matrix!G260:G262`.
7. Assert named ranges exist for:
   - `tblActivityCodes`
   - `tblParameters`
   - `tblThresholds`
   - `tblRequirements`
   - `tblScheduleData`
   - `tblDailyMatrix`
   - `tblTestCases`
   - `SelectedDate`
   - `ActiveActivityCodes`
   - `IntervalHeaders`
8. Assert workbook file is saved at `excel schedule/wfm_scheduling_matrix.xlsx`.

## Manual Verification

Open the workbook in Excel and confirm:

- Sheet order is clear and navigable.
- `Schedule_Matrix` freezes panes below the header and to the right of metadata.
- Metadata columns A-G and interval columns H-BC are visibly distinct.
- Config tables are readable and editable.
- No macro security prompt appears.
