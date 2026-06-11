# State: Excel WFM Scheduling Matrix

## Project Reference

See: `.planning/PROJECT.md` (updated 2026-06-11)

**Core value:** The selected-day staffing variance must be fast, visible, and trustworthy after every schedule edit.
**Current focus:** v1 workbook-engine release checkpoint

## Workflow Preferences

- Mode: Interactive
- Granularity: Standard
- Execution: Sequential
- Git tracking: Yes
- Research: Yes, but full research agents/search providers were unavailable at initialization
- Plan check: Yes
- Verifier: Yes
- PR body sections: none

## Current Decisions

- Use Daily View + Weekly Data Engine.
- Use hidden/protected `Calc_Engine` layer.
- Use `ActivityCodes` rule table rather than hardcoded activity logic.
- Use `TestCases` sheet for workbook formula validation.
- Defer forecasting, imports, optimization, VBA automation, and enterprise workflow features.

## Latest Completion

- Phase 1 workbook skeleton and config architecture implemented in `feature/wfm-workbook-engine`.
- Phase 2 requirements and calculation engine implemented, including Sheets-compatible staffed headcount lookup and missing-requirement handling.
- Phase 3 live daily matrix and summary dashboard implemented, including selected-day net variance chart.
- Phase 4 validation, conditional formatting, and performance hardening implemented, including invalid-code, invalid-schedule, duplicate-row, missing-date, and missing-requirement states.
- Phase 5 workbook-native test harness implemented, including seven expected-vs-actual checks and overall pass/fail status.
- Generated workbook: `excel schedule/wfm_scheduling_matrix.xlsx`
- Structural verifier: `scripts/verify_wfm_workbook.py`
- Latest verification: unit tests, workbook generation, and workbook verifier passing for the local `.xlsx`.

## Next Command

Run `$gsd-verify-work` or commit/push the v1 workbook-engine changes.
