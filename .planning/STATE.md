# State: Excel WFM Scheduling Matrix

## Project Reference

See: `.planning/PROJECT.md` (updated 2026-06-11)

**Core value:** The selected-day staffing variance must be fast, visible, and trustworthy after every schedule edit.
**Current focus:** Phase 1 - Workbook Skeleton and Config Architecture

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

## Next Command

Run `$gsd-plan-phase 1` to create the detailed plan for Phase 1.
