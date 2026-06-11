# State: Excel WFM Scheduling Matrix

## Project Reference

See: `.planning/PROJECT.md` (updated 2026-06-11)

**Core value:** The selected-day staffing variance must be fast, visible, and trustworthy after every schedule edit.
**Current focus:** Phase 1 complete - verify and prepare Phase 2 calculation engine

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
- Generated workbook: `excel schedule/wfm_scheduling_matrix.xlsx`
- Structural verifier: `scripts/verify_wfm_workbook.py`
- Latest verification: unit tests, workbook generation, and workbook verifier passing.

## Next Command

Run `$gsd-verify-work 1` for a formal Phase 1 audit, or `$gsd-plan-phase 2` to plan the requirements and calculation engine.
