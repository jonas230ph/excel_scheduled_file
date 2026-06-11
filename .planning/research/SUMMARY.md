# Research Summary: Excel WFM Scheduling Matrix

**Generated:** 2026-06-11
**Method:** Lightweight synthesis from office-hours, CEO review, and engineering review. Full GSD research subagents were unavailable in this runtime.

## Stack

- Microsoft Excel workbook.
- Excel Tables for structured inputs.
- Named ranges and named formulas for reusable logic.
- `LET`, lookup formulas, and structured references for readable formulas.
- No VBA in Phase 1.

## Table Stakes

- A daily schedule grid with 48 half-hour intervals.
- Configurable activity codes.
- Requirements by interval.
- Scheduled vs required summary.
- Over/under variance visualization.
- Visible validation states for bad inputs.

## Watch Out For

- Overnight shifts are the highest-risk formula path.
- Hidden zeroes are dangerous. Missing requirements and invalid codes must be visible.
- A full 7-day wall chart risks slow recalculation and wide-sheet editing errors.
- Hardcoded activity-code checks will rot quickly as business rules change.

## Recommendation

Build Phase 1 as a selected-day matrix with weekly-normalized source data, a protected calculation engine, and a visible test harness.
