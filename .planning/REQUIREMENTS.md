# Requirements: Excel WFM Scheduling Matrix

**Defined:** 2026-06-11
**Core Value:** The selected-day staffing variance must be fast, visible, and trustworthy after every schedule edit.

## v1 Requirements

### Workbook Structure

- [x] **STRUCT-01**: Workbook includes `Config_Settings`, `Staffing_Requirements`, and `Schedule_Matrix` tabs.
- [x] **STRUCT-02**: Workbook includes a hidden or protected `Calc_Engine` layer for schedule normalization, interval overlap, staffed flags, and validation states.
- [x] **STRUCT-03**: Workbook includes a `Schedule_Data` table for normalized weekly schedule records.
- [x] **STRUCT-04**: Workbook includes a `TestCases` sheet with expected vs actual formula checks.
- [x] **STRUCT-05**: Workbook may include `Summary_Dashboard` for a cleaner chart view, sourced only from tested summaries.

### Config Settings

- [x] **CONFIG-01**: `Config_Settings!A1:H40` stores active activity code definitions.
- [x] **CONFIG-02**: Activity code rules include `Code`, `DisplayLabel`, `IsActive`, `CountsAsStaffed`, `IsPaid`, `Category`, `NumericValue`, and `ColorGroup`.
- [x] **CONFIG-03**: `Config_Settings!J1:K30` stores numeric parameters including interval minutes, intervals per day, day start time, and tolerance.
- [x] **CONFIG-04**: `Config_Settings!M1:Q20` stores threshold and color state definitions for conditional formatting.
- [x] **CONFIG-05**: Active activity codes are exposed through a named range for dropdown validation.

### Staffing Requirements

- [x] **REQ-01**: `Staffing_Requirements` stores requirements by `OperationalDate`, `IntervalStart`, and `RequiredHeadcount`.
- [x] **REQ-02**: Requirements support all 48 half-hour intervals per operational day.
- [x] **REQ-03**: Requirements do not split by channel, queue, skill, team, site, or location in Phase 1.
- [x] **REQ-04**: Blank or missing requirements produce a visible missing-requirement state instead of silently treating demand as zero.

### Schedule Matrix

- [x] **MATRIX-01**: `Schedule_Matrix!A:G` stores visible metadata columns: Employee ID, Name, Contractual Hours, Current Shift, Target Day Select, Shift Start, Shift End.
- [x] **MATRIX-02**: `Schedule_Matrix!H:BC` maps exactly 48 half-hour intervals from `00:00` through `23:30`.
- [x] **MATRIX-03**: The matrix renders one selected operational day at a time.
- [x] **MATRIX-04**: The matrix displays visible activity states for productive and nonproductive activity codes.
- [x] **MATRIX-05**: The matrix summary rows include scheduled productive headcount, required headcount, and over/under variance by interval.
- [x] **MATRIX-06**: The matrix supports a clean net staffing chart based on the selected day summary.

### Calculation Engine

- [x] **CALC-01**: Schedule start and end values are normalized into real datetime windows.
- [x] **CALC-02**: Overnight shifts crossing midnight are counted correctly for the selected operational day.
- [x] **CALC-03**: Interval overlap uses `IntervalStart < ScheduleEnd` and `IntervalEnd > ScheduleStart`.
- [x] **CALC-04**: Staffed count uses activity lookup rules rather than hardcoded code checks.
- [x] **CALC-05**: Formula logic uses named ranges or named formulas for key calculations.
- [x] **CALC-06**: Core calculations avoid volatile functions where possible.

### Validation and Conditional Formatting

- [ ] **VALID-01**: Missing activity codes produce a visible invalid-code state.
- [ ] **VALID-02**: Zero-length shifts produce a visible invalid-schedule state.
- [ ] **VALID-03**: Duplicate employee/date/start/end/code rows produce a visible duplicate warning.
- [ ] **VALID-04**: Blank selected date produces a visible missing-date state.
- [ ] **VALID-05**: Conditional formatting distinguishes under, exact, over, nonproductive, invalid, and missing states.

### Test Coverage

- [ ] **TEST-01**: `TestCases` verifies a same-day shift such as `09:00-17:00`.
- [ ] **TEST-02**: `TestCases` verifies an overnight shift such as `22:00-06:00`.
- [ ] **TEST-03**: `TestCases` verifies interval boundary cases at shift start and shift end.
- [ ] **TEST-04**: `TestCases` verifies `OWD` counts as staffed.
- [ ] **TEST-05**: `TestCases` verifies `Brk`, `Lch`, `Mt`, and `Trn` display but do not count as staffed by default.
- [ ] **TEST-06**: `TestCases` verifies missing code, blank requirement, exact, under, and over states.
- [ ] **TEST-07**: `TestCases` exposes an overall pass/fail status for the workbook engine.

## v2 Requirements

### Forecasting and Optimization

- **FORECAST-01**: Erlang-C or equivalent staffing forecast engine.
- **FORECAST-02**: Automated demand generation from historical interval volume.
- **OPT-01**: Solver-based schedule optimization.
- **OPT-02**: VBA or Office Script scheduling automation, if justified by validated usage.

### Data Import and Integrations

- **IMPORT-01**: Bulk roster upload with mapping and validation.
- **IMPORT-02**: Import from HRIS, payroll, or WFM exports.
- **INTEG-01**: Integration with Genesys, NICE, Verint, or similar platforms.

### Enterprise Controls

- **CTRL-01**: User permissions and sheet protection roles.
- **CTRL-02**: Audit history for schedule changes.
- **CTRL-03**: Approval workflow for PTO, absence, or shift changes.
- **CTRL-04**: Requirement splits by skill, queue, channel, team, or site.

## Out of Scope

| Feature | Reason |
|---------|--------|
| Erlang-C forecasting | Separate forecasting layer; not needed to validate schedule matrix mechanics |
| Bulk roster upload | Adds import mapping complexity before core workbook behavior is proven |
| VBA scheduling algorithms | Opaque and harder to debug than worksheet formulas for Phase 1 |
| Solver optimization | Premature until users trust interval variance output |
| Multi-channel requirements | User selected interval-only requirements for Phase 1 |
| External integrations | Standalone workbook behavior must be stable first |
| Permissions and approvals | Enterprise governance, not MVP engine |

## Traceability

| Requirement | Phase | Status |
|-------------|-------|--------|
| STRUCT-01 | Phase 1 | Complete |
| STRUCT-02 | Phase 1 | Complete |
| STRUCT-03 | Phase 1 | Complete |
| STRUCT-04 | Phase 1 | Complete |
| STRUCT-05 | Phase 3 | Complete |
| CONFIG-01 | Phase 1 | Complete |
| CONFIG-02 | Phase 1 | Complete |
| CONFIG-03 | Phase 1 | Complete |
| CONFIG-04 | Phase 1 | Complete |
| CONFIG-05 | Phase 1 | Complete |
| REQ-01 | Phase 2 | Complete |
| REQ-02 | Phase 2 | Complete |
| REQ-03 | Phase 2 | Complete |
| REQ-04 | Phase 2 | Complete |
| MATRIX-01 | Phase 3 | Complete |
| MATRIX-02 | Phase 3 | Complete |
| MATRIX-03 | Phase 3 | Complete |
| MATRIX-04 | Phase 3 | Complete |
| MATRIX-05 | Phase 3 | Complete |
| MATRIX-06 | Phase 3 | Complete |
| CALC-01 | Phase 2 | Complete |
| CALC-02 | Phase 2 | Complete |
| CALC-03 | Phase 2 | Complete |
| CALC-04 | Phase 2 | Complete |
| CALC-05 | Phase 2 | Complete |
| CALC-06 | Phase 4 | Complete |
| VALID-01 | Phase 4 | Pending |
| VALID-02 | Phase 4 | Pending |
| VALID-03 | Phase 4 | Pending |
| VALID-04 | Phase 4 | Pending |
| VALID-05 | Phase 4 | Pending |
| TEST-01 | Phase 5 | Pending |
| TEST-02 | Phase 5 | Pending |
| TEST-03 | Phase 5 | Pending |
| TEST-04 | Phase 5 | Pending |
| TEST-05 | Phase 5 | Pending |
| TEST-06 | Phase 5 | Pending |
| TEST-07 | Phase 5 | Pending |

**Coverage:**
- v1 requirements: 38 total
- Mapped to phases: 38
- Unmapped: 0
- Completed: 26
- Remaining pending: 12

---
*Requirements defined: 2026-06-11*
*Last updated: 2026-06-11 after Google Sheets compatibility and dashboard completion*
