# Schedule_Matrix Formula Contract

This is the first incremental build contract for the `Schedule_Matrix` tab.
It defines exact anchors and formulas before the workbook generator writes cells.

## Layout

The compact build view uses row 5 for scheduled productive headcount, as requested.
The production-scale summary rows sit below the 250-row roster body at rows
260 through 262.

| Range | Purpose |
|-------|---------|
| `A1:G1` | Metadata headers: `EmployeeID`, `Name`, `ContractualHours`, `CurrentShift`, `TargetDaySelect`, `ShiftStart`, `ShiftEnd` |
| `E2` | Selected operational date; named range `SelectedDate` points to `Schedule_Matrix!$E$2` |
| `H1:BC1` | 48 interval headers from `00:00` through `23:30` |
| `G5` | Label: `Scheduled Productive` |
| `G6` | Label: `Required` |
| `G7` | Label: `Over/Under` |
| `A8:G257` | Employee metadata rows linked from selected-day `Schedule_Data` records |
| `H8:BC257` | Visible interval activity display grid |
| `Calc_Engine!A8:G257` | Selected-day roster extraction from `tblScheduleData` |
| `Calc_Engine!H8:BC257` | Numeric productive coverage fractions used by summaries |
| `G260:G262` | Production summary labels |

## Interval Header Formula

Put this in `Schedule_Matrix!H1` and copy across through `BC1`.

```excel
=TIME(0,0,0)+(COLUMN()-COLUMN($H$1))*TIME(0,30,0)
```

Format `H1:BC1` as `hh:mm`.

## Visible 30-Minute Interval Intersection Cell

Put this in `Schedule_Matrix!H8` and copy across/down through `BC257`.

This formula displays the activity code when the row overlaps the interval. It returns
blank for null roster rows, missing activity codes, missing times, and non-overlapping
intervals. Nonproductive and inactive codes may display here, but they are not counted
by the numeric coverage formula.

```excel
=LET(
  id,$A8,
  code,UPPER(TRIM($D8)),
  rawStart,$F8,
  rawEnd,$G8,
  startDT,IF(INT(rawStart)>0,rawStart,$E$2+MOD(rawStart,1)),
  endBase,IF(INT(rawEnd)>0,rawEnd,$E$2+MOD(rawEnd,1)),
  endDT,IF(endBase<=startDT,endBase+1,endBase),
  intervalStart,$E$2+H$1,
  intervalEnd,intervalStart+TIME(0,30,0),
  overlapDays,MAX(0,MIN(endDT,intervalEnd)-MAX(startDT,intervalStart)),
  IF(OR(id="",code="",rawStart="",rawEnd="",rawStart=rawEnd),"",IF(overlapDays>0,code,""))
)
```

Copy behavior:

- `$A8`, `$D8`, `$F8`, `$G8` keep the source columns fixed while allowing the employee row to change.
- `$E$2` keeps the selected operational date fixed everywhere.
- `H$1` allows the interval column to change while keeping the header row fixed.

## Numeric Productive Coverage Formula

Put this in `Calc_Engine!H8` and copy across/down through `BC257`.

This is the formula the scheduled headcount row should sum. It returns `0`, `0.5`, or
`1` for 30-minute intervals in normal use, while also supporting any partial-minute
overlap if a later schedule import has non-quarter-hour times.

```excel
=LET(
  id,Schedule_Matrix!$A8,
  code,UPPER(TRIM(Schedule_Matrix!$D8)),
  rawStart,Schedule_Matrix!$F8,
  rawEnd,Schedule_Matrix!$G8,
  countsAsStaffed,IFERROR(--XLOOKUP(code,UPPER(tblActivityCodes[Code]),tblActivityCodes[CountsAsStaffed],0),0),
  startDT,IF(INT(rawStart)>0,rawStart,SelectedDate+MOD(rawStart,1)),
  endBase,IF(INT(rawEnd)>0,rawEnd,SelectedDate+MOD(rawEnd,1)),
  endDT,IF(endBase<=startDT,endBase+1,endBase),
  intervalStart,SelectedDate+Schedule_Matrix!H$1,
  intervalEnd,intervalStart+TIME(0,30,0),
  overlapDays,MAX(0,MIN(endDT,intervalEnd)-MAX(startDT,intervalStart)),
  IF(OR(id="",code="",rawStart="",rawEnd="",rawStart=rawEnd,countsAsStaffed=0),0,overlapDays/TIME(0,30,0))
)
```

Expected behavior:

- `08:15-12:00` contributes `0.5` to `08:00`, `1` to `08:30` through `11:30`, and `0` to `12:00`.
- `OFF`, `AL`, `SL`, `Brk`, `Lch`, `Mt`, and `Trn` return `0` when their `CountsAsStaffed` lookup is false.
- Blank/null roster rows return `0`, avoiding `#VALUE!` in summary formulas.

## Scheduled Headcount Formula For Row 5

Put this in `Schedule_Matrix!H5` and copy across through `BC5`.

```excel
=SUM(Calc_Engine!H$8:H$257)
```

The row anchors keep the employee roster window fixed while the interval column changes
from `H` through `BC`.

For production summary rows, put the same formula in `Schedule_Matrix!H260`
and copy across:

```excel
=SUM(Calc_Engine!H$8:H$257)
```

## Required Headcount Formula

Put this in `Schedule_Matrix!H6` and copy across through `BC6`.

```excel
=IFNA(SUMIFS(tblRequirements[RequiredHeadcount],tblRequirements[OperationalDate],$E$2,tblRequirements[IntervalStart],H$1),"")
```

This returns blank when a requirement row is missing, allowing conditional formatting
to show the missing-requirement state instead of silently treating demand as zero.

For production rows, place the same formula in `Schedule_Matrix!H261`.

## Net Staffing Formula

Put this in `Schedule_Matrix!H7` and copy across through `BC7`.

```excel
=IF(H$6="","MissingRequirement",H$5-H$6)
```

For production rows, place this adjusted formula in `Schedule_Matrix!H262`.

```excel
=IF(H$261="","MissingRequirement",H$260-H$261)
```

## Table And Name Dependencies

These formulas assume the workbook generator creates:

| Name | Refers To |
|------|-----------|
| `SelectedDate` | `Schedule_Matrix!$E$2` |
| `tblActivityCodes` | Excel table over `Config_Settings!$A$1:$H$40` |
| `tblRequirements` | Excel table over `Staffing_Requirements!$A$1:$C$337` or equivalent dynamic table |
| `IntervalHeaders` | `Schedule_Matrix!$H$1:$BC$1` |

Required `tblActivityCodes` columns:

| Column | Used By Formula |
|--------|-----------------|
| `Code` | Activity lookup key |
| `CountsAsStaffed` | Numeric productive/nonproductive flag |

Required `tblRequirements` columns:

| Column | Used By Formula |
|--------|-----------------|
| `OperationalDate` | Selected date match |
| `IntervalStart` | Interval header match |
| `RequiredHeadcount` | Required staffing value |
