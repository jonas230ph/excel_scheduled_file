# State: Excel WFM Scheduling Matrix

## Project Reference

See: `.planning/PROJECT.md` (updated 2026-06-11)

**Core value:** The selected-day staffing variance must be fast, visible, and trustworthy after every schedule edit.
**Current focus:** Phase 4 validation warnings and Phase 5 workbook-native test harness

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
- Generated workbook: `excel schedule/wfm_scheduling_matrix.xlsx`
- Structural verifier: `scripts/verify_wfm_workbook.py`
- Latest verification: unit tests, workbook generation, workbook verifier, and live Google Sheets readback passing.

## Next Command

Run `$gsd-plan-phase 4` for validation warning states, or `$gsd-plan-phase 5` for workbook-native test cases.
