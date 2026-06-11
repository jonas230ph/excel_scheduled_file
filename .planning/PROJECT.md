# Excel WFM Scheduling Matrix

## What This Is

This project is an editable Microsoft Excel workforce management scheduling workbook for daily operational staffing control. It lets a planner choose a target day, edit employee schedule rows, see activity-coded coverage across 48 half-hour intervals, and immediately compare scheduled productive headcount against interval staffing requirements.

The workbook is inspired by enterprise WFM systems like Genesys and NICE IEX, but Phase 1 is deliberately narrower: a responsive daily matrix backed by weekly-normalized schedule data, not a forecasting or optimization platform.

## Core Value

The selected-day staffing variance must be fast, visible, and trustworthy after every schedule edit.

## Requirements

### Validated

(None yet - ship to validate)

### Active

- [ ] Build a workbook tab structure with `Config_Settings`, `Staffing_Requirements`, and `Schedule_Matrix` as the core user-facing tabs.
- [ ] Support a selected operational day that can switch the live daily grid without rebuilding formulas.
- [ ] Store activity behavior in configurable lookup tables rather than hardcoded formula checks.
- [ ] Render 48 half-hour intervals from `00:00` through `23:30` in `Schedule_Matrix!H:BC`.
- [ ] Keep metadata fields in `Schedule_Matrix!A:G`: Employee ID, Name, Contractual Hours, Current Shift, Target Day Select, Shift Start, Shift End.
- [ ] Compute scheduled productive headcount versus required headcount by interval.
- [ ] Show over, exact, under, invalid, missing, and nonproductive states through conditional formatting.
- [ ] Include a hidden or protected calculation layer with named formulas so the visible matrix stays readable.
- [ ] Include a `TestCases` sheet covering same-day shifts, overnight shifts, boundary overlaps, missing codes, blank requirements, and over/under states.

### Out of Scope

- Erlang-C forecasting engines - demand modeling is a later product layer, not needed to prove the matrix engine.
- Bulk roster uploads - import mapping and validation would distract from the core editing loop.
- Complex VBA scheduling algorithms - opaque macros would make the MVP harder to trust and debug.
- Solver-based shift optimization - users should first validate the interval math and daily control-room workflow.
- Multi-channel, queue, team, site, or skill-based requirements - Phase 1 requirements vary only by date and interval.
- External integrations with Genesys, NICE, Verint, HRIS, or payroll systems - the workbook must work standalone first.
- Permissions, audit history, approvals, and workflow routing - enterprise governance is later scope.

## Context

Prior scope reviews settled the core architecture:

- Phase 1 should be `Daily View + Weekly Data Engine`, not a giant full-week wall chart.
- Raw schedule records should be normalized into weekly schedule rows, while the live matrix renders one selected operational day.
- Overnight shifts must use datetime overlap logic, not simple `EndTime > StartTime` checks.
- Activity states should remain visible in the grid, but only productive activity codes should count toward scheduled staffing totals.
- Spreadsheet failures should be visible. Missing activity codes, blank requirements, duplicate rows, invalid times, and unsupported inputs must produce validation states rather than silent formula errors.

## Sheet Tab Architecture

### `Config_Settings`

Purpose: central rule table for workbook behavior, activity code semantics, numeric mapping parameters, and conditional-format thresholds.

Suggested anchors:

| Range | Purpose |
|-------|---------|
| `A1:H1` | Activity code table headers |
| `A2:H40` | Active activity code definitions |
| `J1:K1` | Numeric parameter headers |
| `J2:K30` | Interval, day, and mapping parameters |
| `M1:Q1` | Threshold/color definition headers |
| `M2:Q20` | Conditional-format state definitions |

Activity code table columns:

| Column | Header | Description |
|--------|--------|-------------|
| `A` | `Code` | Short code such as `OWD`, `Brk`, `Lch`, `Mt`, `Trn` |
| `B` | `DisplayLabel` | Friendly label shown in validation/help |
| `C` | `IsActive` | `TRUE/FALSE`, controls dropdown availability |
| `D` | `CountsAsStaffed` | `TRUE/FALSE`, controls productive headcount |
| `E` | `IsPaid` | `TRUE/FALSE`, payroll classification |
| `F` | `Category` | Productive, break, meal, meeting, training, absence, off |
| `G` | `NumericValue` | Optional numeric mapping for formulas/charts |
| `H` | `ColorGroup` | Conditional-format key |

Parameter table examples:

| Cell/Name | Value | Purpose |
|-----------|-------|---------|
| `J2` / `IntervalMinutes` | `30` | Base interval duration |
| `J3` / `IntervalsPerDay` | `48` | Number of columns from H through BC |
| `J4` / `DayStartTime` | `00:00` | First interval |
| `J5` / `ToleranceHeadcount` | `0` | Exact-match tolerance |

Threshold table examples:

| State | MinVariance | MaxVariance | FillColor | FontColor |
|-------|-------------|-------------|-----------|-----------|
| Under | blank | `-1` | red | white |
| Exact | `0` | `0` | green | white |
| Over | `1` | blank | blue | white |
| Invalid | blank | blank | amber | black |
| MissingRequirement | blank | blank | gray | black |
| Nonproductive | blank | blank | light orange | black |

### `Staffing_Requirements`

Purpose: required productive headcount by operational date and 30-minute interval.

Suggested anchors:

| Range | Purpose |
|-------|---------|
| `A1:C1` | Requirement table headers |
| `A2:C337` | One week of date + interval + required headcount, if preloaded |
| `E1:BA1` | Optional wide day-view headers |
| `E2:BA8` | Optional 7-day by 48-interval input layout |

Normalized table columns:

| Column | Header | Description |
|--------|--------|-------------|
| `A` | `OperationalDate` | Date the requirement applies to |
| `B` | `IntervalStart` | Time from `00:00` through `23:30` |
| `C` | `RequiredHeadcount` | Required productive agents |

Phase 1 should treat requirements as date + interval only. No channel, queue, skill, site, team, or location split.

### `Schedule_Matrix`

Purpose: live daily grid that planners use to review one selected operational day.

Suggested anchors:

| Range | Purpose |
|-------|---------|
| `A1:G1` | Metadata headers |
| `H1:BC1` | 48 half-hour interval headers |
| `A2:G200` | Employee schedule metadata rows |
| `H2:BC200` | Visible activity/coverage grid |
| `H202:BC202` | Scheduled productive headcount summary |
| `H203:BC203` | Required headcount summary |
| `H204:BC204` | Over/Under variance summary |
| `H206:BC220` | Net staffing chart source or helper rows |

Metadata columns:

| Column | Header | Description |
|--------|--------|-------------|
| `A` | `EmployeeID` | Unique employee identifier |
| `B` | `Name` | Employee display name |
| `C` | `ContractualHours` | Daily or weekly contract hours |
| `D` | `CurrentShift` | Human-readable shift label/code |
| `E` | `TargetDaySelect` | Selected operational date, preferably linked to a single control cell |
| `F` | `ShiftStart` | Scheduled start datetime/time |
| `G` | `ShiftEnd` | Scheduled end datetime/time |

Interval columns:

| Range | Meaning |
|-------|---------|
| `H1` | `00:00` |
| `I1` | `00:30` |
| `J1` | `01:00` |
| `...` | every 30 minutes |
| `BC1` | `23:30` |

The visible grid should show activity codes or compact state markers, while summary rows count only activity codes where `CountsAsStaffed=TRUE`.

### Supporting Tabs

| Tab | Purpose | Visibility |
|-----|---------|------------|
| `Schedule_Data` | Normalized weekly schedule records powering the matrix | User-editable or protected input |
| `Calc_Engine` | Hidden/protected helper calculations and named formula outputs | Hidden/protected |
| `TestCases` | Formula correctness harness with expected vs actual outputs | Visible to builder/admin |
| `Summary_Dashboard` | Optional cleaner chart/report view sourced from `Schedule_Matrix` summaries | Visible |

## Named Range Architecture

Core tables:

| Name | Refers To |
|------|-----------|
| `tblActivityCodes` | `Config_Settings!$A$1:$H$40` |
| `tblParameters` | `Config_Settings!$J$1:$K$30` |
| `tblThresholds` | `Config_Settings!$M$1:$Q$20` |
| `tblRequirements` | `Staffing_Requirements!$A$1:$C$337` or dynamic table object |
| `tblScheduleData` | `Schedule_Data!$A$1:$K$1000` |
| `tblDailyMatrix` | `Schedule_Matrix!$A$1:$BC$200` |
| `tblTestCases` | `TestCases!$A$1:$H$100` |

Scalar names:

| Name | Refers To |
|------|-----------|
| `SelectedDate` | `Schedule_Matrix!$E$2` or `Control!$B$2` if a control tab is added |
| `IntervalMinutes` | parameter lookup for `IntervalMinutes` |
| `IntervalsPerDay` | parameter lookup for `IntervalsPerDay` |
| `DayStartTime` | parameter lookup for `DayStartTime` |

Dynamic formula names:

| Name | Purpose |
|------|---------|
| `IntervalHeaders` | 48 generated interval starts from `00:00` to `23:30` |
| `ActiveActivityCodes` | active codes for dropdown validation |
| `StaffedCodeFlags` | lookup array mapping codes to staffed/not staffed |
| `SelectedDayRequirements` | required headcount row/array for the selected day |
| `ScheduleWindowStart` | normalized schedule start datetime |
| `ScheduleWindowEnd` | normalized schedule end datetime, with overnight adjustment |
| `IntervalOverlapFlag` | true when an employee schedule overlaps an interval |
| `StaffedIntervalFlag` | overlap flag multiplied by `CountsAsStaffed` |
| `DailyScheduledHeadcount` | scheduled productive headcount by interval |
| `DailyOverUnder` | scheduled minus required by interval |

## Data Flow Map

```text
+-------------------+
| Config_Settings  |
| activity rules   |
| thresholds       |
| parameters       |
+---------+---------+
          |
          v
+-------------------+        +------------------------+
| Schedule_Data     |        | Staffing_Requirements |
| weekly records    |        | date + interval reqs  |
+---------+---------+        +-----------+------------+
          |                              |
          v                              v
       +------------------------------------+
       | Calc_Engine                         |
       | normalize windows                   |
       | apply selected day                  |
       | test interval overlap               |
       | lookup activity behavior            |
       | produce validation states           |
       +----------------+-------------------+
                        |
                        v
       +------------------------------------+
       | Schedule_Matrix                     |
       | A:G metadata                        |
       | H:BC 48 interval grid               |
       | summary rows: scheduled/required/net|
       +----------------+-------------------+
                        |
                        v
       +------------------------------------+
       | Summary_Dashboard / Chart           |
       | net staffing variance by interval   |
       +------------------------------------+
```

## Constraints

- **Performance**: Render one selected day at a time. Do not calculate a full 7-day x employee wall chart in Phase 1.
- **Formula Design**: Prefer table lookups, named formulas, `LET`, and structured references. Avoid volatile functions for core calculations.
- **Trust**: Invalid data must create visible validation states, not silent zeroes.
- **Compatibility**: Phase 1 should work as a standalone workbook without VBA.
- **Scope**: Requirements vary only by operational date and half-hour interval.

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Daily View + Weekly Data Engine | Keeps the workbook responsive while preserving weekly and overnight schedule support | — Pending |
| Hidden/protected calculation layer | Keeps visible matrix readable and formula behavior auditable | — Pending |
| Activity behavior is table-driven | Prevents hardcoded `OWD`/`Brk` formula logic and allows business-rule changes | — Pending |
| Add `TestCases` sheet | Excel formulas need regression coverage for overnight and boundary cases | — Pending |
| Defer forecasting, uploads, and VBA automation | These are enterprise layers, not required to prove the MVP engine | — Pending |

## Evolution

This document evolves at phase transitions and milestone boundaries.

**After each phase transition** (via `$gsd-transition`):
1. Requirements invalidated? -> Move to Out of Scope with reason
2. Requirements validated? -> Move to Validated with phase reference
3. New requirements emerged? -> Add to Active
4. Decisions to log? -> Add to Key Decisions
5. "What This Is" still accurate? -> Update if drifted

**After each milestone** (via `$gsd-complete-milestone`):
1. Full review of all sections
2. Core Value check - still the right priority?
3. Audit Out of Scope - reasons still valid?
4. Update Context with current state

---
*Last updated: 2026-06-11 after initialization*
