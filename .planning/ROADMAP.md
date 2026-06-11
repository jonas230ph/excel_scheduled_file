# Roadmap: Excel WFM Scheduling Matrix

**Created:** 2026-06-11
**Granularity:** Standard
**Execution:** Sequential

## Phase 1: Workbook Skeleton and Config Architecture

**Status:** Complete

**Goal:** Create the workbook tab contract, anchors, named range plan, and configurable activity/threshold tables.

**Requirements:** STRUCT-01, STRUCT-02, STRUCT-03, STRUCT-04, CONFIG-01, CONFIG-02, CONFIG-03, CONFIG-04, CONFIG-05

**Deliverables:**
- `Config_Settings` tab with activity code, parameter, and threshold tables.
- `Schedule_Data`, `Calc_Engine`, `Schedule_Matrix`, `Staffing_Requirements`, and `TestCases` tab definitions.
- Named range architecture documented and/or created in the workbook.
- Phase 1 sample codes: `OWD`, `OT`, `Brk`, `Lch`, `Mt`, `Trn`, `PTO`, `Sick`, `Off`, `UA`.

**Acceptance Criteria:**
- Activity behavior can be changed in `Config_Settings` without editing formulas.
- Interval count and interval minutes are defined as parameters.
- Sheet anchors match the project specification.

## Phase 2: Requirements and Calculation Engine

**Status:** Complete

**Goal:** Build the date/interval requirement model and overnight-safe calculation engine.

**Requirements:** REQ-01, REQ-02, REQ-03, REQ-04, CALC-01, CALC-02, CALC-03, CALC-04, CALC-05

**Deliverables:**
- `Staffing_Requirements` normalized date + interval + required headcount table.
- `Calc_Engine` normalized schedule windows.
- Interval overlap logic for same-day and overnight shifts.
- Staffed count logic driven by `ActivityCodes[CountsAsStaffed]`.

**Acceptance Criteria:**
- `22:00-06:00` shifts count on both relevant operational days where appropriate.
- Blank requirements do not silently become zero demand.
- Productive and nonproductive activity states remain visually distinct.

## Phase 3: Live Daily Schedule Matrix and Summary

**Status:** Complete

**Goal:** Render the selected operational day in `Schedule_Matrix` and produce interval staffing variance.

**Requirements:** MATRIX-01, MATRIX-02, MATRIX-03, MATRIX-04, MATRIX-05, MATRIX-06

**Deliverables:**
- `Schedule_Matrix!A:G` metadata fields.
- `Schedule_Matrix!H:BC` 48 interval headers from `00:00` to `23:30`.
- Visible grid states for activity codes.
- Summary rows for scheduled, required, and over/under.
- Net staffing chart source and chart.

**Acceptance Criteria:**
- Changing the selected date refreshes matrix, summary, and chart.
- Editing a schedule row updates affected interval states.
- The chart shows selected-day net variance only.

## Phase 4: Validation, Conditional Formatting, and Performance Hardening

**Status:** In Progress

**Goal:** Make data quality problems visible and keep recalculation responsive.

**Requirements:** CALC-06, VALID-01, VALID-02, VALID-03, VALID-04, VALID-05

**Deliverables:**
- Conditional formatting for under, exact, over, nonproductive, invalid, and missing states.
- Helper checks for missing activity code, blank selected date, duplicate rows, zero-length shifts, and blank requirements.
- Review of formula volatility and selected-day scoping.

**Acceptance Criteria:**
- Invalid records show visible warning states.
- Core formulas avoid volatile functions where practical.
- Workbook remains scoped to the selected day rather than recalculating full-week matrix output.

## Phase 5: Test Harness and Release Readiness

**Status:** Pending

**Goal:** Add in-workbook tests that prove the formula engine works before the workbook is considered ready.

**Requirements:** TEST-01, TEST-02, TEST-03, TEST-04, TEST-05, TEST-06, TEST-07

**Deliverables:**
- `TestCases` sheet with expected vs actual checks.
- Test cases for same-day shifts, overnight shifts, boundaries, activity-code counting, missing data, and variance states.
- Overall pass/fail summary.

**Acceptance Criteria:**
- All test cases pass.
- At least one test fails when `CountsAsStaffed` or overnight logic is intentionally broken.
- Release notes list out-of-scope v2 items.

## Deferred V2+

- Erlang-C forecasting.
- Bulk roster uploads.
- VBA or Office Script scheduling automation.
- Solver-based optimization.
- Multi-skill, channel, queue, team, site, or location requirements.
- External WFM/HRIS integrations.
- Permissions, audit trail, and approval workflows.
