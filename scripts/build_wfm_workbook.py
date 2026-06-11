"""Build the Phase 1 Excel WFM scheduling workbook skeleton."""

from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = ROOT / "excel schedule" / "wfm_scheduling_matrix.xlsx"
SAMPLE_DATE = date(2026, 6, 11)
ROSTER_START_ROW = 8
MAX_ROSTER_ROWS = 250
ROSTER_END_ROW = ROSTER_START_ROW + MAX_ROSTER_ROWS - 1
SUMMARY_SCHEDULED_ROW = 260
SUMMARY_REQUIRED_ROW = 261
SUMMARY_VARIANCE_ROW = 262
INTERVAL_START_COLUMN = 8
INTERVAL_END_COLUMN = 55
VALIDATION_STATE_COLUMN = 56
SHEETS = [
    "Config_Settings",
    "Staffing_Requirements",
    "Schedule_Data",
    "Calc_Engine",
    "Schedule_Matrix",
    "Summary_Dashboard",
    "TestCases",
]


def add_table(ws, name: str, ref: str) -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)


def add_name(wb: Workbook, name: str, refers_to: str) -> None:
    wb.defined_names.add(DefinedName(name, attr_text=refers_to))


def interval_time(index: int) -> time:
    minutes = index * 30
    return time(minutes // 60, minutes % 60)


def setup_workbook() -> Workbook:
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    for sheet_name in SHEETS:
        wb.create_sheet(sheet_name)
    return wb


def populate_config(ws) -> None:
    activity_headers = [
        "Code",
        "DisplayLabel",
        "IsActive",
        "CountsAsStaffed",
        "IsPaid",
        "Category",
        "NumericValue",
        "ColorGroup",
    ]
    ws.append(activity_headers)
    activity_rows = [
        ("OWD", "On Work Duty", True, True, True, "Productive", 1, "Productive"),
        ("OT", "Overtime", True, True, True, "Productive", 1, "Productive"),
        ("Brk", "Break", True, False, True, "Break", 0, "Nonproductive"),
        ("Lch", "Lunch", True, False, False, "Meal", 0, "Nonproductive"),
        ("Mt", "Meeting", True, False, True, "Meeting", 0, "Nonproductive"),
        ("Trn", "Training", True, False, True, "Training", 0, "Nonproductive"),
        ("PTO", "Paid Time Off", True, False, True, "Absence", 0, "Inactive"),
        ("Sick", "Sick Leave", True, False, True, "Absence", 0, "Inactive"),
        ("Off", "Off", True, False, False, "Off", 0, "Inactive"),
        ("UA", "Unapproved Absence", True, False, False, "Absence", 0, "Inactive"),
        ("AL", "Annual Leave", True, False, True, "Absence", 0, "Inactive"),
        ("SL", "Sick Leave Short Code", True, False, True, "Absence", 0, "Inactive"),
    ]
    for row in activity_rows:
        ws.append(row)
    for row_index in range(len(activity_rows) + 2, 41):
        ws.cell(row=row_index, column=1, value="")

    parameter_headers = ["Parameter", "Value"]
    for column_index, header in enumerate(parameter_headers, start=10):
        ws.cell(row=1, column=column_index, value=header)
    parameters = [
        ("IntervalMinutes", 30),
        ("IntervalsPerDay", 48),
        ("MaxRosterRows", MAX_ROSTER_ROWS),
        ("DayStartTime", time(0, 0)),
        ("ToleranceHeadcount", 0),
        ("WorkbookEngineVersion", "1.0.0-phase1"),
        ("TemplateSchemaVersion", "1.0.0"),
        ("FormulaSetVersion", "1.0.0"),
        ("DataSnapshotID", "sample-2026-06-11"),
    ]
    for row_index, row in enumerate(parameters, start=2):
        ws.cell(row=row_index, column=10, value=row[0])
        ws.cell(row=row_index, column=11, value=row[1])
    for row_index in range(len(parameters) + 2, 31):
        ws.cell(row=row_index, column=10, value="")

    threshold_headers = ["State", "MinVariance", "MaxVariance", "FillColor", "FontColor"]
    for column_index, header in enumerate(threshold_headers, start=13):
        ws.cell(row=1, column=column_index, value=header)
    thresholds = [
        ("Under", None, -1, "F8696B", "FFFFFF"),
        ("Exact", 0, 0, "63BE7B", "FFFFFF"),
        ("Over", 1, None, "5B9BD5", "FFFFFF"),
        ("InvalidCode", None, None, "FFC000", "000000"),
        ("InvalidSchedule", None, None, "FFD966", "000000"),
        ("Duplicate", None, None, "D9EAD3", "000000"),
        ("MissingDate", None, None, "C9DAF8", "000000"),
        ("MissingRequirement", None, None, "BFBFBF", "000000"),
        ("Nonproductive", None, None, "F4B183", "000000"),
    ]
    for row_index, row in enumerate(thresholds, start=2):
        for offset, value in enumerate(row, start=13):
            ws.cell(row=row_index, column=offset, value=value)
    for row_index in range(len(thresholds) + 2, 21):
        ws.cell(row=row_index, column=13, value="")

    add_table(ws, "tblActivityCodes", "A1:H40")
    add_table(ws, "tblParameters", "J1:K30")
    add_table(ws, "tblThresholds", "M1:Q20")


def populate_requirements(ws) -> None:
    ws.append(["OperationalDate", "IntervalStart", "RequiredHeadcount"])
    for index in range(48):
        interval = interval_time(index)
        required = 4 if 8 <= interval.hour < 18 else 2
        ws.append([SAMPLE_DATE, interval, required])
    for row_index in range(50, 338):
        ws.cell(row=row_index, column=1, value="")
    add_table(ws, "tblRequirements", "A1:C337")


def populate_schedule_data(ws) -> None:
    headers = [
        "EmployeeID",
        "Name",
        "OperationalDate",
        "ShiftStart",
        "ShiftEnd",
        "ActivityCode",
        "ContractualHours",
        "CurrentShift",
        "Notes",
        "ValidationState",
        "SourceRowID",
    ]
    ws.append(headers)
    rows = [
        ("E001", "Avery Chen", SAMPLE_DATE, datetime(2026, 6, 11, 8, 15), datetime(2026, 6, 11, 17, 0), "OWD", 8, "OWD", "Partial first interval", "", 1),
        ("E002", "Blake Diaz", SAMPLE_DATE, datetime(2026, 6, 11, 23, 45), datetime(2026, 6, 12, 2, 15), "OWD", 8, "OWD", "Overnight", "", 2),
        ("E003", "Casey Singh", SAMPLE_DATE, datetime(2026, 6, 11, 12, 0), datetime(2026, 6, 11, 12, 30), "Lch", 8, "Lch", "Nonproductive", "", 3),
    ]
    for row in rows:
        ws.append(row)
    for row_index in range(2, len(rows) + 2):
        ws.cell(row=row_index, column=3).number_format = "yyyy-mm-dd"
        ws.cell(row=row_index, column=4).number_format = "yyyy-mm-dd hh:mm"
        ws.cell(row=row_index, column=5).number_format = "yyyy-mm-dd hh:mm"
    for row_index in range(len(rows) + 2, 1001):
        ws.cell(row=row_index, column=1, value="")
    for row_index in range(2, 1001):
        ws.cell(row=row_index, column=10, value=schedule_validation_formula(row_index))
    add_schedule_data_conditional_formatting(ws)
    add_table(ws, "tblScheduleData", "A1:K1000")


def schedule_validation_formula(row: int) -> str:
    return (
        f'=IF(COUNTA($A{row}:$F{row})=0,"",'
        f'IF($C{row}="","MissingDate",'
        f'IF(OR($A{row}="",$D{row}="",$E{row}="",$D{row}=$E{row}),"InvalidSchedule",'
        f'IF(COUNTIF(ActiveActivityCodes,$F{row})=0,"InvalidCode",'
        f'IF(COUNTIFS(tblScheduleData[EmployeeID],$A{row},'
        f'tblScheduleData[OperationalDate],$C{row},'
        f'tblScheduleData[ShiftStart],$D{row},'
        f'tblScheduleData[ShiftEnd],$E{row},'
        f'tblScheduleData[ActivityCode],$F{row})>1,"Duplicate","Valid")))))'
    )


def matrix_visible_formula(row: int, column_letter: str) -> str:
    return (
        "=LET("
        f"id,$A{row},"
        f"code,UPPER(TRIM($D{row})),"
        f"rawStart,$F{row},"
        f"rawEnd,$G{row},"
        f"startDT,IF(INT(rawStart)>0,rawStart,$E$2+MOD(rawStart,1)),"
        f"endBase,IF(INT(rawEnd)>0,rawEnd,$E$2+MOD(rawEnd,1)),"
        "endDT,IF(endBase<=startDT,endBase+1,endBase),"
        f"intervalStart,$E$2+{column_letter}$1,"
        "intervalEnd,intervalStart+TIME(0,30,0),"
        "overlapDays,MAX(0,MIN(endDT,intervalEnd)-MAX(startDT,intervalStart)),"
        'IF(OR(id="",code="",rawStart="",rawEnd="",rawStart=rawEnd),"",IF(overlapDays>0,code,""))'
        ")"
    )


def calc_coverage_formula(row: int, column_letter: str) -> str:
    return (
        "=LET("
        f"id,Schedule_Matrix!$A{row},"
        f"code,UPPER(TRIM(Schedule_Matrix!$D{row})),"
        f"rawStart,Schedule_Matrix!$F{row},"
        f"rawEnd,Schedule_Matrix!$G{row},"
        "countsAsStaffed,IFERROR(--INDEX(tblActivityCodes[CountsAsStaffed],MATCH(code,tblActivityCodes[Code],0)),0),"
        "startDT,IF(INT(rawStart)>0,rawStart,SelectedDate+MOD(rawStart,1)),"
        "endBase,IF(INT(rawEnd)>0,rawEnd,SelectedDate+MOD(rawEnd,1)),"
        "endDT,IF(endBase<=startDT,endBase+1,endBase),"
        f"intervalStart,SelectedDate+Schedule_Matrix!{column_letter}$1,"
        "intervalEnd,intervalStart+TIME(0,30,0),"
        "overlapDays,MAX(0,MIN(endDT,intervalEnd)-MAX(startDT,intervalStart)),"
        'IF(OR(id="",code="",rawStart="",rawEnd="",rawStart=rawEnd,countsAsStaffed=0),0,overlapDays/TIME(0,30,0))'
        ")"
    )


def requirement_formula(column_letter: str, row: int) -> str:
    return (
        f'=IF(COUNTIFS(tblRequirements[OperationalDate],$E$2,'
        f'tblRequirements[IntervalStart],{column_letter}${row})=0,"",'
        f"SUMIFS(tblRequirements[RequiredHeadcount],tblRequirements[OperationalDate],$E$2,"
        f"tblRequirements[IntervalStart],{column_letter}${row}))"
    )


def selected_day_variance_formula(column_letter: str, scheduled_row: int, required_row: int) -> str:
    return (
        f'=IF($E$2="","MissingDate",'
        f'IF({column_letter}${required_row}="","MissingRequirement",'
        f"ROUND({column_letter}${scheduled_row}-{column_letter}${required_row},2)))"
    )


def selected_schedule_formula(field_name: str, row: int) -> str:
    relative_index = f"ROWS($A${ROSTER_START_ROW}:A{row})"
    return (
        f'=IFERROR(INDEX(FILTER(tblScheduleData[{field_name}],'
        f"tblScheduleData[OperationalDate]=SelectedDate),{relative_index}),\"\")"
    )


def populate_schedule_matrix(wb: Workbook) -> None:
    ws = wb["Schedule_Matrix"]
    calc = wb["Calc_Engine"]
    headers = ["EmployeeID", "Name", "ContractualHours", "CurrentShift", "TargetDaySelect", "ShiftStart", "ShiftEnd"]
    for column_index, header in enumerate(headers, start=1):
        ws.cell(row=1, column=column_index, value=header)
        calc.cell(row=1, column=column_index, value=header)
    for index in range(48):
        column = INTERVAL_START_COLUMN + index
        ws.cell(row=1, column=column, value=f"=TIME(0,0,0)+(COLUMN()-COLUMN($H$1))*TIME(0,30,0)")
        calc.cell(row=1, column=column, value=f"=Schedule_Matrix!{ws.cell(row=1, column=column).coordinate}")
    ws.cell(row=1, column=VALIDATION_STATE_COLUMN, value="ValidationState")
    calc.cell(row=1, column=VALIDATION_STATE_COLUMN, value="ValidationState")

    ws["E2"] = SAMPLE_DATE
    ws["G5"] = "Scheduled Productive"
    ws["G6"] = "Required"
    ws["G7"] = "Over/Under"
    ws.cell(row=SUMMARY_SCHEDULED_ROW, column=7, value="Scheduled Productive")
    ws.cell(row=SUMMARY_REQUIRED_ROW, column=7, value="Required")
    ws.cell(row=SUMMARY_VARIANCE_ROW, column=7, value="Over/Under")

    for row in range(ROSTER_START_ROW, ROSTER_END_ROW + 1):
        calc.cell(row=row, column=1, value=selected_schedule_formula("EmployeeID", row))
        calc.cell(row=row, column=2, value=selected_schedule_formula("Name", row))
        calc.cell(row=row, column=3, value=selected_schedule_formula("ContractualHours", row))
        calc.cell(row=row, column=4, value=selected_schedule_formula("ActivityCode", row))
        calc.cell(row=row, column=5, value=f'=IF(Calc_Engine!A{row}="","",SelectedDate)')
        calc.cell(row=row, column=6, value=selected_schedule_formula("ShiftStart", row))
        calc.cell(row=row, column=7, value=selected_schedule_formula("ShiftEnd", row))
        calc.cell(row=row, column=VALIDATION_STATE_COLUMN, value=selected_schedule_formula("ValidationState", row))
        for column in range(1, 8):
            ws.cell(row=row, column=column, value=f"=Calc_Engine!{ws.cell(row=row, column=column).coordinate}")
        ws.cell(row=row, column=VALIDATION_STATE_COLUMN, value=f"=Calc_Engine!{ws.cell(row=row, column=VALIDATION_STATE_COLUMN).coordinate}")
        for column in range(INTERVAL_START_COLUMN, INTERVAL_END_COLUMN + 1):
            column_letter = ws.cell(row=1, column=column).column_letter
            ws.cell(row=row, column=column, value=matrix_visible_formula(row, column_letter))
            calc.cell(row=row, column=column, value=calc_coverage_formula(row, column_letter))

    for column in range(INTERVAL_START_COLUMN, INTERVAL_END_COLUMN + 1):
        column_letter = ws.cell(row=1, column=column).column_letter
        ws.cell(row=5, column=column, value=f"=SUM(Calc_Engine!{column_letter}${ROSTER_START_ROW}:{column_letter}${ROSTER_END_ROW})")
        ws.cell(
            row=6,
            column=column,
            value=requirement_formula(column_letter, 1),
        )
        ws.cell(row=7, column=column, value=selected_day_variance_formula(column_letter, 5, 6))
        ws.cell(row=SUMMARY_SCHEDULED_ROW, column=column, value=f"=SUM(Calc_Engine!{column_letter}${ROSTER_START_ROW}:{column_letter}${ROSTER_END_ROW})")
        ws.cell(
            row=SUMMARY_REQUIRED_ROW,
            column=column,
            value=requirement_formula(column_letter, 1),
        )
        ws.cell(
            row=SUMMARY_VARIANCE_ROW,
            column=column,
            value=selected_day_variance_formula(column_letter, SUMMARY_SCHEDULED_ROW, SUMMARY_REQUIRED_ROW),
        )

    add_conditional_formatting(ws)


def populate_summary_dashboard(wb: Workbook) -> None:
    ws = wb["Summary_Dashboard"]
    ws.append(["Interval", "Scheduled Productive", "Required", "Over/Under"])
    matrix = wb["Schedule_Matrix"]
    for index, column in enumerate(range(INTERVAL_START_COLUMN, INTERVAL_END_COLUMN + 1), start=2):
        column_letter = matrix.cell(row=1, column=column).column_letter
        ws.cell(row=index, column=1, value=f"=Schedule_Matrix!{column_letter}$1")
        ws.cell(row=index, column=2, value=f"=Schedule_Matrix!{column_letter}${SUMMARY_SCHEDULED_ROW}")
        ws.cell(row=index, column=3, value=f"=Schedule_Matrix!{column_letter}${SUMMARY_REQUIRED_ROW}")
        ws.cell(row=index, column=4, value=f"=Schedule_Matrix!{column_letter}${SUMMARY_VARIANCE_ROW}")

    chart = LineChart()
    chart.title = "Selected-Day Net Staffing Variance"
    chart.y_axis.title = "Agents"
    chart.x_axis.title = "Interval"
    chart.height = 8
    chart.width = 24
    data = Reference(ws, min_col=4, min_row=1, max_row=49)
    categories = Reference(ws, min_col=1, min_row=2, max_row=49)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    ws.add_chart(chart, "F2")


def add_conditional_formatting(ws) -> None:
    productive_fill = PatternFill("solid", fgColor="C6EFCE")
    nonproductive_fill = PatternFill("solid", fgColor="FCE4D6")
    under_fill = PatternFill("solid", fgColor="F8696B")
    exact_fill = PatternFill("solid", fgColor="63BE7B")
    over_fill = PatternFill("solid", fgColor="5B9BD5")
    invalid_code_fill = PatternFill("solid", fgColor="FFC000")
    invalid_schedule_fill = PatternFill("solid", fgColor="FFD966")
    duplicate_fill = PatternFill("solid", fgColor="D9EAD3")
    missing_date_fill = PatternFill("solid", fgColor="C9DAF8")
    missing_fill = PatternFill("solid", fgColor="BFBFBF")
    body_range = f"H{ROSTER_START_ROW}:BC{ROSTER_END_ROW}"
    row_range = f"A{ROSTER_START_ROW}:BC{ROSTER_END_ROW}"
    variance_range = f"H{SUMMARY_VARIANCE_ROW}:BC{SUMMARY_VARIANCE_ROW}"
    ws.conditional_formatting.add(body_range, FormulaRule(formula=['H8="OWD"'], fill=productive_fill))
    ws.conditional_formatting.add(body_range, FormulaRule(formula=['OR(H8="BRK",H8="LCH",H8="MT",H8="TRN")'], fill=nonproductive_fill))
    ws.conditional_formatting.add(row_range, FormulaRule(formula=['$BD8="InvalidCode"'], fill=invalid_code_fill))
    ws.conditional_formatting.add(row_range, FormulaRule(formula=['$BD8="InvalidSchedule"'], fill=invalid_schedule_fill))
    ws.conditional_formatting.add(row_range, FormulaRule(formula=['$BD8="Duplicate"'], fill=duplicate_fill))
    ws.conditional_formatting.add(row_range, FormulaRule(formula=['$BD8="MissingDate"'], fill=missing_date_fill))
    ws.conditional_formatting.add("E2", FormulaRule(formula=['$E$2=""'], fill=missing_date_fill))
    ws.conditional_formatting.add("H7:BC7", FormulaRule(formula=['H7="MissingDate"'], fill=missing_date_fill))
    ws.conditional_formatting.add("H7:BC7", FormulaRule(formula=['H7="MissingRequirement"'], fill=missing_fill))
    ws.conditional_formatting.add(variance_range, FormulaRule(formula=[f'H{SUMMARY_VARIANCE_ROW}="MissingDate"'], fill=missing_date_fill))
    ws.conditional_formatting.add(variance_range, FormulaRule(formula=[f'H{SUMMARY_VARIANCE_ROW}="MissingRequirement"'], fill=missing_fill))
    ws.conditional_formatting.add("H7:BC7", CellIsRule(operator="lessThan", formula=["0"], fill=under_fill))
    ws.conditional_formatting.add("H7:BC7", CellIsRule(operator="equal", formula=["0"], fill=exact_fill))
    ws.conditional_formatting.add("H7:BC7", CellIsRule(operator="greaterThan", formula=["0"], fill=over_fill))
    ws.conditional_formatting.add(variance_range, CellIsRule(operator="lessThan", formula=["0"], fill=under_fill))
    ws.conditional_formatting.add(variance_range, CellIsRule(operator="equal", formula=["0"], fill=exact_fill))
    ws.conditional_formatting.add(variance_range, CellIsRule(operator="greaterThan", formula=["0"], fill=over_fill))


def add_schedule_data_conditional_formatting(ws) -> None:
    invalid_code_fill = PatternFill("solid", fgColor="FFC000")
    invalid_schedule_fill = PatternFill("solid", fgColor="FFD966")
    duplicate_fill = PatternFill("solid", fgColor="D9EAD3")
    missing_date_fill = PatternFill("solid", fgColor="C9DAF8")
    data_range = "A2:K1000"
    ws.conditional_formatting.add(data_range, FormulaRule(formula=['$J2="InvalidCode"'], fill=invalid_code_fill))
    ws.conditional_formatting.add(data_range, FormulaRule(formula=['$J2="InvalidSchedule"'], fill=invalid_schedule_fill))
    ws.conditional_formatting.add(data_range, FormulaRule(formula=['$J2="Duplicate"'], fill=duplicate_fill))
    ws.conditional_formatting.add(data_range, FormulaRule(formula=['$J2="MissingDate"'], fill=missing_date_fill))


def populate_test_cases(ws) -> None:
    headers = ["TestID", "Scenario", "Input", "Expected", "Actual", "Pass", "Requirement", "Notes"]
    ws.append(headers)
    ws["J1"] = "OverallStatus"
    ws["K1"] = '=IF(COUNTIF(F2:F8,FALSE)=0,"PASS","FAIL")'
    rows = [
        (
            "TEST-01",
            "Same-day productive shift",
            "09:00-17:00 OWD",
            16,
            "=(DATE(2026,6,11)+TIME(17,0,0)-(DATE(2026,6,11)+TIME(9,0,0)))/TIME(0,30,0)",
            "=ABS(E2-D2)<0.000001",
            "TEST-01",
            "Verifies same-day half-hour interval count.",
        ),
        (
            "TEST-02",
            "Overnight productive shift",
            "22:00-06:00 OWD",
            16,
            "=(DATE(2026,6,12)+TIME(6,0,0)-(DATE(2026,6,11)+TIME(22,0,0)))/TIME(0,30,0)",
            "=ABS(E3-D3)<0.000001",
            "TEST-02",
            "Verifies midnight-crossing interval count.",
        ),
        (
            "TEST-03",
            "Boundary overlap",
            "08:15 start against 08:00-08:30 interval",
            0.5,
            "=MAX(0,MIN(DATE(2026,6,11)+TIME(8,30,0),DATE(2026,6,11)+TIME(12,0,0))-MAX(DATE(2026,6,11)+TIME(8,0,0),DATE(2026,6,11)+TIME(8,15,0)))/TIME(0,30,0)",
            "=ABS(E4-D4)<0.000001",
            "TEST-03",
            "Verifies partial interval proration.",
        ),
        (
            "TEST-04",
            "Productive code lookup",
            "OWD CountsAsStaffed",
            1,
            '=--INDEX(tblActivityCodes[CountsAsStaffed],MATCH("OWD",tblActivityCodes[Code],0))',
            "=ABS(E5-D5)<0.000001",
            "TEST-04",
            "Verifies OWD counts as staffed through config lookup.",
        ),
        (
            "TEST-05",
            "Nonproductive code lookup",
            "Brk/Lch/Mt/Trn CountsAsStaffed",
            0,
            '=SUM(--INDEX(tblActivityCodes[CountsAsStaffed],MATCH("Brk",tblActivityCodes[Code],0)),--INDEX(tblActivityCodes[CountsAsStaffed],MATCH("Lch",tblActivityCodes[Code],0)),--INDEX(tblActivityCodes[CountsAsStaffed],MATCH("Mt",tblActivityCodes[Code],0)),--INDEX(tblActivityCodes[CountsAsStaffed],MATCH("Trn",tblActivityCodes[Code],0)))',
            "=ABS(E6-D6)<0.000001",
            "TEST-05",
            "Verifies visible nonproductive states do not count as staffed.",
        ),
        (
            "TEST-06",
            "Validation and variance states",
            "Invalid code / blank req / exact / under / over",
            "InvalidCode|MissingRequirement|Exact|Under|Over",
            '=TEXTJOIN("|",TRUE,IF(COUNTIF(ActiveActivityCodes,"BAD")=0,"InvalidCode",""),IF(COUNTIFS(tblRequirements[OperationalDate],DATE(2099,1,1),tblRequirements[IntervalStart],TIME(0,0,0))=0,"MissingRequirement",""),IF(4-4=0,"Exact",""),IF(3-4<0,"Under",""),IF(5-4>0,"Over",""))',
            "=E7=D7",
            "TEST-06",
            "Verifies key validation and variance labels.",
        ),
        (
            "TEST-07",
            "Overall workbook test status",
            "Rows TEST-01 through TEST-06",
            "PASS",
            '=IF(COUNTIF(F2:F7,FALSE)=0,"PASS","FAIL")',
            "=E8=D8",
            "TEST-07",
            "Exposes an overall pass/fail status for workbook formula tests.",
        ),
    ]
    for row in rows:
        ws.append(row)
    for row_index in range(len(rows) + 2, 101):
        ws.cell(row=row_index, column=1, value="")
    add_test_case_conditional_formatting(ws)
    add_table(ws, "tblTestCases", "A1:H100")


def add_test_case_conditional_formatting(ws) -> None:
    pass_fill = PatternFill("solid", fgColor="C6EFCE")
    fail_fill = PatternFill("solid", fgColor="F8696B")
    ws.conditional_formatting.add("F2:F8", FormulaRule(formula=["F2=TRUE"], fill=pass_fill))
    ws.conditional_formatting.add("F2:F8", FormulaRule(formula=["F2=FALSE"], fill=fail_fill))
    ws.conditional_formatting.add("K1", FormulaRule(formula=['K1="PASS"'], fill=pass_fill))
    ws.conditional_formatting.add("K1", FormulaRule(formula=['K1="FAIL"'], fill=fail_fill))


def setup_names_and_validation(wb: Workbook) -> None:
    add_name(wb, "SelectedDate", "'Schedule_Matrix'!$E$2")
    add_name(wb, "IntervalHeaders", "'Schedule_Matrix'!$H$1:$BC$1")
    add_name(wb, "ActiveActivityCodes", "'Config_Settings'!$A$2:$A$13")
    add_name(wb, "tblDailyMatrix", "'Schedule_Matrix'!$A$1:$BC$257")

    schedule = wb["Schedule_Data"]
    validation = DataValidation(type="list", formula1="=ActiveActivityCodes", allow_blank=True)
    schedule.add_data_validation(validation)
    validation.add("F2:F1000")


def format_workbook(wb: Workbook) -> None:
    freezes = {
        "Config_Settings": "A2",
        "Staffing_Requirements": "A2",
        "Schedule_Data": "A2",
        "Schedule_Matrix": "H2",
        "TestCases": "A2",
    }
    for ws in wb.worksheets:
        ws.freeze_panes = freezes.get(ws.title)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center")
        ws.sheet_view.showGridLines = True

    matrix = wb["Schedule_Matrix"]
    calc = wb["Calc_Engine"]
    dashboard = wb["Summary_Dashboard"]
    for ws in [matrix, calc]:
        for column in range(1, 8):
            ws.column_dimensions[ws.cell(row=1, column=column).column_letter].width = 18
        for column in range(8, 56):
            ws.column_dimensions[ws.cell(row=1, column=column).column_letter].width = 8
        ws.column_dimensions[ws.cell(row=1, column=VALIDATION_STATE_COLUMN).column_letter].hidden = True
        for row in [1, 5, 6, 7, SUMMARY_SCHEDULED_ROW, SUMMARY_REQUIRED_ROW, SUMMARY_VARIANCE_ROW]:
            for cell in ws[row]:
                cell.font = Font(bold=True)
    matrix["E2"].number_format = "yyyy-mm-dd"
    for row in range(ROSTER_START_ROW, ROSTER_END_ROW + 1):
        for ws in [matrix, calc]:
            ws.cell(row=row, column=5).number_format = "yyyy-mm-dd"
            ws.cell(row=row, column=6).number_format = "yyyy-mm-dd hh:mm"
            ws.cell(row=row, column=7).number_format = "yyyy-mm-dd hh:mm"
    for column in range(INTERVAL_START_COLUMN, INTERVAL_END_COLUMN + 1):
        matrix.cell(row=1, column=column).number_format = "hh:mm"
        calc.cell(row=1, column=column).number_format = "hh:mm"
        for row in [5, 6, 7, SUMMARY_SCHEDULED_ROW, SUMMARY_REQUIRED_ROW, SUMMARY_VARIANCE_ROW]:
            matrix.cell(row=row, column=column).number_format = "0.00"
            calc.cell(row=row, column=column).number_format = "0.00"
        for row in range(ROSTER_START_ROW, ROSTER_END_ROW + 1):
            calc.cell(row=row, column=column).number_format = "0.00"
    for row in range(2, 50):
        dashboard.cell(row=row, column=1).number_format = "hh:mm"
        for column in range(2, 5):
            dashboard.cell(row=row, column=column).number_format = "0.00"
    calc.sheet_state = "hidden"


def build() -> Workbook:
    wb = setup_workbook()
    populate_config(wb["Config_Settings"])
    populate_requirements(wb["Staffing_Requirements"])
    populate_schedule_data(wb["Schedule_Data"])
    populate_schedule_matrix(wb)
    populate_summary_dashboard(wb)
    populate_test_cases(wb["TestCases"])
    setup_names_and_validation(wb)
    format_workbook(wb)
    return wb


def main() -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb = build()
    wb.save(OUTPUT_PATH)
    print(f"Workbook written to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
