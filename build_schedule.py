"""
build_schedule.py — Build Production_Schedule.xlsx from scratch using openpyxl.
Run: python build_schedule.py
Then: python scripts/recalc.py Production_Schedule.xlsx 60
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule, FormulaRule, ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.workbook.defined_name import DefinedName
from datetime import date, timedelta
import random

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

STATIONS = [
    "Profiling", "Saw",
    "Machining 1", "Machining 2",
    "Fabrication 1", "Fabrication 2",
    "Weld 1", "Weld 2", "Weld 3",
    "Blast", "Paint",
    "Assembly 1", "Assembly 2",
]

WORK_CENTRES = ["Profiling", "Saw", "Machining", "Fabrication", "Weld", "Blast", "Paint", "Assembly"]
PRIORITIES = ["Urgent", "High", "Normal", "Low"]
STATUSES = ["Not Started", "In Progress", "On Hold", "Complete", "Waiting Materials"]

STATION_COLOURS = {
    "Profiling":     "2F5496",
    "Saw":           "843C0C",
    "Machining 1":   "7030A0",
    "Machining 2":   "7030A0",
    "Fabrication 1": "375623",
    "Fabrication 2": "375623",
    "Weld 1":        "1F3864",
    "Weld 2":        "1F3864",
    "Weld 3":        "1F3864",
    "Blast":         "595959",
    "Paint":         "0070C0",
    "Assembly 1":    "404040",
    "Assembly 2":    "404040",
}

TAB_COLOURS = {
    "Config":            "808080",
    "Master Schedule":   "1F3864",
    "Dashboard":         "2E5E9E",
    "Queue — All":       "2E75B6",
    "Queue — Profiling": "C00000",
    "Queue — Saw":       "8B4513",
    "Queue — Machining 1":   "7030A0",
    "Queue — Machining 2":   "7030A0",
    "Queue — Fabrication 1": "375623",
    "Queue — Fabrication 2": "375623",
    "Queue — Weld 1":    "1F3864",
    "Queue — Weld 2":    "1F3864",
    "Queue — Weld 3":    "1F3864",
    "Queue — Blast":     "595959",
    "Queue — Paint":     "0070C0",
    "Queue — Assembly 1":"404040",
    "Queue — Assembly 2":"404040",
}

# Priority rank map (lower = higher urgency)
PRIORITY_RANK = {"Urgent": 1, "High": 2, "Normal": 3, "Low": 4}

# ---------------------------------------------------------------------------
# Column layout helpers
# ---------------------------------------------------------------------------

# Identity columns (1-based indices)
ID_COLS = {
    "Job ID":          1,
    "Customer":        2,
    "Product/Description": 3,
    "Order Qty":       4,
    "Drawing Ref":     5,
    "Job Priority":    6,
    "Overall Status":  7,
    "Order Due Date":  8,
    "Date Received":   9,
    "Days Until Due":  10,
    "Priority Rank":   11,
    "Notes":           12,
}

# Station sub-columns: 4 per station starting at col 13
def station_col_start(station_name):
    idx = STATIONS.index(station_name)
    return 13 + idx * 4

def station_cols(station_name):
    s = station_col_start(station_name)
    return {
        "Status":   s,
        "Est Hrs":  s + 1,
        "Operator": s + 2,
        "Notes":    s + 3,
    }

def col_letter(n):
    return get_column_letter(n)

TOTAL_COLS = 12 + 13 * 4  # = 64

# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="000000", size=11, italic=False, strike=False):
    return Font(bold=bold, color=color, size=size, italic=italic, strike=strike)

def side():
    return Side(style="thin", color="BFBFBF")

def thin_border():
    s = side()
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

# ---------------------------------------------------------------------------
# Sample data
# ---------------------------------------------------------------------------

TODAY = date.today()

SAMPLE_JOBS = [
    # (job_id, customer, product, qty, drawing, priority, due_offset_days, received_offset_days, station_statuses)
    # station_statuses: dict of station -> (status, est_hrs, operator)
    ("J-1001", "Brennan Agri Ltd",      "Disc Harrow 3m",           5,  "DH-3000",  "Urgent", -5,  -30, {"Profiling":"Complete","Saw":"Complete","Machining 1":"Complete","Fabrication 1":"In Progress","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1002", "Murphy Farms",          "Muck Spreader 6T",         2,  "MS-6000",  "High",   -2,  -25, {"Profiling":"Complete","Saw":"Complete","Fabrication 1":"Complete","Fabrication 2":"In Progress","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1003", "O'Brien Tillage",       "Subsoiler 5-Leg",          3,  "SS-0500",  "Urgent", -8,  -40, {"Profiling":"Complete","Saw":"Complete","Machining 1":"Complete","Machining 2":"Complete","Fabrication 1":"Complete","Weld 1":"In Progress","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1004", "Kelly Contracting",     "Bale Handler Twin",        4,  "BH-2000",  "High",    2,  -20, {"Profiling":"Complete","Saw":"Complete","Fabrication 1":"Complete","Weld 1":"Complete","Weld 2":"In Progress","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1005", "Walsh Plant Hire",      "Roller Cambridge 4m",      1,  "RC-4000",  "Normal",  7,  -15, {"Profiling":"Complete","Saw":"Complete","Fabrication 1":"In Progress","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1006", "Fitzpatrick Agri",      "Trailer 14T Grain",        2,  "TG-1400",  "Normal", 14,  -10, {"Profiling":"Not Started","Saw":"Not Started","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1007", "Connolly & Sons",       "Disc Harrow 4m",           3,  "DH-4000",  "High",    1,  -18, {"Profiling":"Complete","Saw":"Complete","Machining 1":"In Progress","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1008", "Riordan Farm Supplies", "Muck Spreader 8T",         1,  "MS-8000",  "Low",    21,   -5, {"Profiling":"Not Started","Saw":"Not Started","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1009", "Daly Agri Repairs",     "Subsoiler 7-Leg",          2,  "SS-0700",  "Normal",  3,  -22, {"Profiling":"Complete","Saw":"In Progress","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1010", "Burke Plant",           "Bale Handler Single",      6,  "BH-1000",  "Urgent", -3,  -35, {"Profiling":"Complete","Saw":"Complete","Fabrication 1":"Complete","Weld 1":"Complete","Weld 2":"Complete","Blast":"In Progress","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1011", "Healy Farms",           "Roller Flat 3m",           2,  "RF-3000",  "Normal", 10,  -12, {"Profiling":"Not Started","Saw":"Not Started","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1012", "Sheehan Contracting",   "Trailer 10T Grain",        3,  "TG-1000",  "High",    0,  -28, {"Profiling":"Complete","Saw":"Complete","Fabrication 1":"Complete","Fabrication 2":"Complete","Weld 1":"Complete","Weld 2":"In Progress","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1013", "O'Sullivan Plant Hire", "Disc Harrow 5m",           2,  "DH-5000",  "Low",    30,   -3, {"Profiling":"Not Started","Saw":"Not Started","Machining 1":"Not Started","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1014", "Cronin Agri",           "Muck Spreader 10T",        1,  "MS-A000",  "High",    2,  -20, {"Profiling":"Waiting Materials","Saw":"Not Started","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1015", "McCarthy Farm Equip",   "Subsoiler 3-Leg",          4,  "SS-0300",  "Normal", 18,   -8, {"Profiling":"Not Started","Saw":"Not Started","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1016", "Power Agri Ltd",        "Bale Handler Triple",      2,  "BH-3000",  "Urgent", -1,  -45, {"Profiling":"Complete","Saw":"Complete","Machining 1":"Complete","Machining 2":"Complete","Fabrication 1":"Complete","Fabrication 2":"Complete","Weld 1":"Complete","Weld 2":"Complete","Weld 3":"Complete","Blast":"Complete","Paint":"In Progress","Assembly 1":"Not Started"}),
    ("J-1017", "Nolan Plant",           "Roller Cambridge 6m",      1,  "RC-6000",  "Normal", 25,   -6, {"Profiling":"Not Started","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1018", "Treacy Agri Supplies",  "Trailer 16T Grain",        2,  "TG-1600",  "High",    3,  -32, {"Profiling":"Complete","Saw":"Complete","Fabrication 1":"On Hold","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1019", "Delaney Farms",         "Disc Harrow 6m",           1,  "DH-6000",  "Low",    45,   -2, {"Profiling":"Not Started","Saw":"Not Started","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1020", "Flynn Contracting",     "Muck Spreader 4T",         3,  "MS-4000",  "Normal",  6,  -14, {"Profiling":"Complete","Saw":"In Progress","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1021", "Higgins Farm Mach.",    "Subsoiler 9-Leg",          1,  "SS-0900",  "High",    1,  -26, {"Profiling":"Complete","Saw":"Complete","Machining 1":"Complete","Fabrication 1":"In Progress","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1022", "Egan Plant Hire",       "Bale Handler Twin XL",     2,  "BH-2100",  "Normal", 12,   -9, {"Profiling":"Not Started","Saw":"Not Started","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1023", "Ryan Agri",             "Roller Flat 5m",           2,  "RF-5000",  "Urgent", -6,  -50, {"Profiling":"Complete","Saw":"Complete","Fabrication 1":"Complete","Fabrication 2":"Complete","Weld 1":"Complete","Weld 2":"Complete","Weld 3":"Complete","Blast":"Complete","Paint":"Complete","Assembly 1":"Complete","Assembly 2":"Complete"}),
    ("J-1024", "Dooley Farm Supplies",  "Trailer 8T Silage",        3,  "TS-0800",  "Low",    35,   -4, {"Profiling":"Not Started","Saw":"Not Started","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
    ("J-1025", "Kavanagh Agri Ltd",     "Disc Harrow 2.5m Folding", 4,  "DH-2500",  "High",    2,  -17, {"Profiling":"Complete","Saw":"Complete","Machining 1":"Waiting Materials","Fabrication 1":"Not Started","Weld 1":"Not Started","Blast":"Not Started","Paint":"Not Started","Assembly 1":"Not Started"}),
]

# Realistic est. hours per station type
EST_HRS = {
    "Profiling": 2, "Saw": 1.5, "Machining 1": 4, "Machining 2": 4,
    "Fabrication 1": 6, "Fabrication 2": 6, "Weld 1": 8, "Weld 2": 8, "Weld 3": 8,
    "Blast": 2, "Paint": 3, "Assembly 1": 5, "Assembly 2": 5,
}

OPERATORS = {
    "Profiling": ["Tom B.", "Sean R."],
    "Saw": ["Tom B.", "Liam F."],
    "Machining 1": ["Pat M.", "Kevin D."],
    "Machining 2": ["Pat M.", "Niall H."],
    "Fabrication 1": ["Declan O.", "Brian S."],
    "Fabrication 2": ["Declan O.", "Mark T."],
    "Weld 1": ["Ciaran W.", "Eoin P."],
    "Weld 2": ["Ciaran W.", "Dave L."],
    "Weld 3": ["Shane C.", "Robbie N."],
    "Blast": ["Paul G.", "Alan F."],
    "Paint": ["Paul G.", "Ger M."],
    "Assembly 1": ["Frank D.", "Colm B."],
    "Assembly 2": ["Frank D.", "Noel K."],
}


# ---------------------------------------------------------------------------
# Workbook setup
# ---------------------------------------------------------------------------

def make_workbook():
    wb = openpyxl.Workbook()

    # Remove default sheet
    wb.remove(wb.active)

    # Create sheets in order
    sheet_names = [
        "Config", "Master Schedule", "Dashboard", "Queue — All",
    ] + [f"Queue — {s}" for s in STATIONS]

    for name in sheet_names:
        ws = wb.create_sheet(name)
        ws.sheet_view.showGridLines = False
        if name in TAB_COLOURS:
            ws.sheet_properties.tabColor = TAB_COLOURS[name]

    return wb


# ---------------------------------------------------------------------------
# Config sheet
# ---------------------------------------------------------------------------

def build_config(wb):
    ws = wb["Config"]
    ws.sheet_state = "hidden"

    headers = ["WorkCentres", "Stations", "Priorities", "Statuses"]
    data = [WORK_CENTRES, STATIONS, PRIORITIES, STATUSES]

    for col_idx, (header, col_data) in enumerate(zip(headers, data), start=1):
        ws.cell(1, col_idx, header).font = Font(bold=True)
        for row_idx, val in enumerate(col_data, start=2):
            ws.cell(row_idx, col_idx, val)

    # Define named ranges pointing at Config sheet columns
    max_rows = max(len(d) for d in data) + 1
    scope_map = {
        "WorkCentres": f"Config!$A$2:$A${len(WORK_CENTRES)+1}",
        "Stations":    f"Config!$B$2:$B${len(STATIONS)+1}",
        "Priorities":  f"Config!$C$2:$C${len(PRIORITIES)+1}",
        "Statuses":    f"Config!$D$2:$D${len(STATUSES)+1}",
    }
    for name, ref in scope_map.items():
        dn = DefinedName(name, attr_text=ref)
        wb.defined_names[name] = dn


# ---------------------------------------------------------------------------
# Master Schedule sheet
# ---------------------------------------------------------------------------

def build_master_schedule(wb):
    ws = wb["Master Schedule"]
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.freeze_panes = "A6"

    # ---- Title rows ----
    dark_blue_fill = fill("1F3864")
    white_font = font(bold=True, color="FFFFFF", size=14)

    # Row 1: merged title
    ws.merge_cells("A1:BL1")
    c = ws["A1"]
    c.value = "PRODUCTION SCHEDULE"
    c.fill = dark_blue_fill
    c.font = white_font
    c.alignment = center()
    ws.row_dimensions[1].height = 30

    # Row 2: sub-heading
    ws.merge_cells("A2:BL2")
    c = ws["A2"]
    c.value = "Agricultural Machinery Manufacturer — Live Job Tracker"
    c.fill = fill("2E5E9E")
    c.font = font(bold=False, color="FFFFFF", size=11, italic=True)
    c.alignment = center()
    ws.row_dimensions[2].height = 20

    # Row 3: spacer
    ws.row_dimensions[3].height = 6

    # ---- Column headers (row 4 = group labels, row 5 = field headers) ----
    # Row 4: group labels
    group_label_fill = fill("D9E1F2")
    ws.merge_cells("A4:L4")
    c = ws["A4"]
    c.value = "JOB IDENTITY"
    c.fill = group_label_fill
    c.font = font(bold=True, color="1F3864")
    c.alignment = center()

    for station in STATIONS:
        sc = station_cols(station)
        start = sc["Status"]
        end = sc["Notes"]
        merge_start = col_letter(start) + "4"
        merge_end = col_letter(end) + "4"
        ws.merge_cells(f"{merge_start}:{merge_end}")
        c = ws[merge_start]
        c.value = station.upper()
        c.fill = fill(STATION_COLOURS[station])
        c.font = font(bold=True, color="FFFFFF")
        c.alignment = center()

    ws.row_dimensions[4].height = 18

    # Row 5: field headers
    header_fill = fill("1F3864")
    header_font = font(bold=True, color="FFFFFF", size=10)

    id_headers = [
        "Job ID", "Customer", "Product/Description", "Order Qty",
        "Drawing Ref", "Job Priority", "Overall Status", "Order Due Date",
        "Date Received", "Days Until Due", "Priority Rank", "Notes",
    ]
    for col_idx, h in enumerate(id_headers, start=1):
        c = ws.cell(5, col_idx, h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center()

    for station in STATIONS:
        sc = station_cols(station)
        st_fill = fill(STATION_COLOURS[station])
        for sub, col_idx in sc.items():
            label_map = {
                "Status":   f"{station}\nStatus",
                "Est Hrs":  f"{station}\nEst Hrs",
                "Operator": f"{station}\nOperator",
                "Notes":    f"{station}\nNotes",
            }
            c = ws.cell(5, col_idx, label_map[sub])
            c.fill = st_fill
            c.font = header_font
            c.alignment = center()

    ws.row_dimensions[5].height = 32

    # ---- Column widths ----
    col_widths = {
        1: 16, 2: 22, 3: 28, 4: 8, 5: 14, 6: 10, 7: 14,
        8: 12, 9: 12, 10: 12, 11: 10, 12: 30,
    }
    for col_idx, w in col_widths.items():
        ws.column_dimensions[col_letter(col_idx)].width = w

    for station in STATIONS:
        sc = station_cols(station)
        ws.column_dimensions[col_letter(sc["Status"])].width = 14
        ws.column_dimensions[col_letter(sc["Est Hrs"])].width = 9
        ws.column_dimensions[col_letter(sc["Operator"])].width = 12
        ws.column_dimensions[col_letter(sc["Notes"])].width = 22

    # Hide Priority Rank column (col 11)
    ws.column_dimensions[col_letter(11)].hidden = True

    # ---- Data Validation ----
    priority_dv = DataValidation(
        type="list", formula1='"Urgent,High,Normal,Low"',
        allow_blank=True, showDropDown=False
    )
    priority_dv.sqref = f"{col_letter(ID_COLS['Job Priority'])}6:{col_letter(ID_COLS['Job Priority'])}500"
    ws.add_data_validation(priority_dv)

    for station in STATIONS:
        sc = station_cols(station)
        status_dv = DataValidation(
            type="list",
            formula1='"Not Started,In Progress,On Hold,Complete,Waiting Materials"',
            allow_blank=True, showDropDown=False
        )
        status_dv.sqref = f"{col_letter(sc['Status'])}6:{col_letter(sc['Status'])}500"
        ws.add_data_validation(status_dv)

    # ---- Conditional Formatting ----
    _apply_master_cf(ws)

    # ---- Excel Table ----
    last_col_letter = col_letter(TOTAL_COLS)
    table_ref = f"A5:{last_col_letter}500"
    table = Table(displayName="tblJobs", ref=table_ref)
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    table.tableStyleInfo = style
    ws.add_table(table)

    # ---- Sample Data ----
    _populate_sample_data(ws)


def _apply_master_cf(ws):
    """Apply 12 conditional formatting rules to Master Schedule."""
    status_col = col_letter(ID_COLS["Overall Status"])
    priority_col = col_letter(ID_COLS["Job Priority"])
    due_col = col_letter(ID_COLS["Days Until Due"])
    full_range = f"A6:{col_letter(TOTAL_COLS)}500"
    priority_range = f"{priority_col}6:{priority_col}500"

    # 1. Complete row → grey fill + strikethrough
    ws.conditional_formatting.add(
        full_range,
        FormulaRule(
            formula=[f'${status_col}6="Complete"'],
            stopIfTrue=True,
            fill=fill("F2F2F2"),
            font=Font(color="AAAAAA", strike=True),
        )
    )
    # 2. Overdue (days < 0) and not complete → red
    ws.conditional_formatting.add(
        full_range,
        FormulaRule(
            formula=[f'AND(${due_col}6<0,${status_col}6<>"Complete")'],
            fill=fill("FCE4D6"),
            font=Font(bold=True, color="C00000"),
        )
    )
    # 3. Due 0-3 days and not complete → amber
    ws.conditional_formatting.add(
        full_range,
        FormulaRule(
            formula=[f'AND(${due_col}6>=0,${due_col}6<=3,${status_col}6<>"Complete")'],
            fill=fill("FFF2CC"),
        )
    )
    # 4-7. Priority colour on priority column only
    ws.conditional_formatting.add(
        priority_range,
        FormulaRule(formula=[f'${priority_col}6="Urgent"'], fill=fill("FFCCCC"), font=Font(bold=True, color="C00000"))
    )
    ws.conditional_formatting.add(
        priority_range,
        FormulaRule(formula=[f'${priority_col}6="High"'], fill=fill("FCE4D6"), font=Font(bold=True, color="833C00"))
    )
    ws.conditional_formatting.add(
        priority_range,
        FormulaRule(formula=[f'${priority_col}6="Normal"'], fill=fill("E2EFDA"), font=Font(bold=True, color="375623"))
    )
    ws.conditional_formatting.add(
        priority_range,
        FormulaRule(formula=[f'${priority_col}6="Low"'], fill=fill("D6E4F7"), font=Font(bold=True, color="1F3864"))
    )
    # 8-11. Status cell colours — apply to each station status column individually
    for station in STATIONS:
        sc = station_cols(station)
        st_col = col_letter(sc["Status"])
        st_range = f"{st_col}6:{st_col}500"
        ws.conditional_formatting.add(st_range, FormulaRule(formula=[f'${st_col}6="Waiting Materials"'], fill=fill("FCE4D6")))
        ws.conditional_formatting.add(st_range, FormulaRule(formula=[f'${st_col}6="Complete"'], fill=fill("E2EFDA")))
        ws.conditional_formatting.add(st_range, FormulaRule(formula=[f'${st_col}6="In Progress"'], fill=fill("FFFF99")))
        ws.conditional_formatting.add(st_range, FormulaRule(formula=[f'${st_col}6="On Hold"'], fill=fill("D9D9D9")))


def _populate_sample_data(ws):
    random.seed(42)

    for row_idx, job in enumerate(SAMPLE_JOBS, start=6):
        (job_id, customer, product, qty, drawing, priority,
         due_offset, rcv_offset, station_map) = job

        due_date = TODAY + timedelta(days=due_offset)
        rcv_date = TODAY + timedelta(days=rcv_offset)

        # Determine overall status from station statuses
        all_statuses = list(station_map.values())
        if all(s == "Complete" for s in all_statuses):
            overall = "Complete"
        elif any(s == "In Progress" for s in all_statuses):
            overall = "In Progress"
        elif any(s == "Waiting Materials" for s in all_statuses):
            overall = "Waiting Materials"
        elif any(s == "On Hold" for s in all_statuses):
            overall = "On Hold"
        else:
            overall = "Not Started"

        ws.cell(row_idx, ID_COLS["Job ID"], job_id)
        ws.cell(row_idx, ID_COLS["Customer"], customer)
        ws.cell(row_idx, ID_COLS["Product/Description"], product)
        ws.cell(row_idx, ID_COLS["Order Qty"], qty)
        ws.cell(row_idx, ID_COLS["Drawing Ref"], drawing)
        ws.cell(row_idx, ID_COLS["Job Priority"], priority)

        # Overall Status: formula
        # =IF(COUNTIF(station_status_cols,"Complete")=13,"Complete",IF(COUNTIF(...),"In Progress",...))
        # For simplicity, write as a formula referencing station status cells
        ws.cell(row_idx, ID_COLS["Overall Status"], overall)

        ws.cell(row_idx, ID_COLS["Order Due Date"], due_date).number_format = "DD/MM/YYYY"
        ws.cell(row_idx, ID_COLS["Date Received"], rcv_date).number_format = "DD/MM/YYYY"

        # Days Until Due: formula
        due_cell = col_letter(ID_COLS["Order Due Date"]) + str(row_idx)
        ws.cell(row_idx, ID_COLS["Days Until Due"],
                f'=IFERROR({due_cell}-TODAY(),"")')

        # Priority Rank: formula
        ws.cell(row_idx, ID_COLS["Priority Rank"],
                f'=IFERROR(MATCH({col_letter(ID_COLS["Job Priority"])}{row_idx},Priorities,0),99)')

        # Station data
        for station in STATIONS:
            sc = station_cols(station)
            st_status = station_map.get(station, "Not Started")
            ops = OPERATORS.get(station, ["TBC"])
            hrs = EST_HRS.get(station, 2)

            ws.cell(row_idx, sc["Status"], st_status)
            ws.cell(row_idx, sc["Est Hrs"], hrs)
            if st_status not in ("Not Started",):
                ws.cell(row_idx, sc["Operator"], random.choice(ops))


# ---------------------------------------------------------------------------
# Dashboard sheet
# ---------------------------------------------------------------------------

def build_dashboard(wb):
    ws = wb["Dashboard"]
    ws.freeze_panes = "A8"

    # Title
    ws.merge_cells("A1:L1")
    c = ws["A1"]
    c.value = "PRODUCTION DASHBOARD"
    c.fill = fill("2E5E9E")
    c.font = font(bold=True, color="FFFFFF", size=16)
    c.alignment = center()
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:L2")
    c = ws["A2"]
    c.value = "Manager Overview — Live KPIs"
    c.fill = fill("D9E1F2")
    c.font = font(bold=False, color="1F3864", size=11, italic=True)
    c.alignment = center()
    ws.row_dimensions[2].height = 18

    # ---- KPI Cards (row 4-6) ----
    ms = "'Master Schedule'"
    status_col = col_letter(ID_COLS["Overall Status"])
    due_col = col_letter(ID_COLS["Days Until Due"])
    est_hrs_cols = []
    for station in STATIONS:
        sc = station_cols(station)
        est_hrs_cols.append(f"{ms}!${col_letter(sc['Est Hrs'])}$6:${col_letter(sc['Est Hrs'])}$500")

    kpis = [
        ("Total Active Jobs",
         f'=COUNTIF({ms}!${status_col}$6:${status_col}$500,"<>Complete")',
         "2E5E9E"),
        ("Overdue Jobs",
         f'=SUMPRODUCT(({ms}!${due_col}$6:${due_col}$500<0)*({ms}!${status_col}$6:${status_col}$500<>"Complete")*({ms}!$A$6:$A$500<>""))',
         "C00000"),
        ("Due This Week",
         f'=SUMPRODUCT(({ms}!${due_col}$6:${due_col}$500>=0)*({ms}!${due_col}$6:${due_col}$500<=7)*({ms}!${status_col}$6:${status_col}$500<>"Complete")*({ms}!$A$6:$A$500<>""))',
         "375623"),
        ("Total Remaining Est Hrs",
         f'=SUMPRODUCT(({ms}!${status_col}$6:${status_col}$500<>"Complete")*({ms}!$A$6:$A$500<>"")*(' +
         "+".join(["({ms}!${c}$6:${c}$500)".format(ms=ms, c=col_letter(station_cols(s)["Est Hrs"])) for s in STATIONS]) + "))",
         "1F3864"),
    ]

    kpi_cols = [1, 4, 7, 10]
    for (label, formula, colour), start_col in zip(kpis, kpi_cols):
        # Merge 3 columns
        ws.merge_cells(start_row=4, start_column=start_col, end_row=4, end_column=start_col+2)
        ws.merge_cells(start_row=5, start_column=start_col, end_row=5, end_column=start_col+2)
        ws.merge_cells(start_row=6, start_column=start_col, end_row=6, end_column=start_col+2)

        label_cell = ws.cell(4, start_col, label)
        label_cell.fill = fill(colour)
        label_cell.font = font(bold=True, color="FFFFFF", size=10)
        label_cell.alignment = center()

        val_cell = ws.cell(5, start_col, formula)
        val_cell.fill = fill("FFFFFF")
        val_cell.font = Font(bold=True, size=22, color=colour)
        val_cell.alignment = center()

        # Spacer row
        sp = ws.cell(6, start_col, "")
        sp.fill = fill(colour)

    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 40
    ws.row_dimensions[6].height = 8

    # ---- Backlog Summary Table ----
    # Row 7: spacer
    ws.row_dimensions[7].height = 12

    # Headers row 8
    backlog_headers = [
        "Station", "Jobs Queued", "Remaining Hrs",
        "Urgent", "High", "Normal", "Low",
        "Overdue", "Due This Week", "% Complete"
    ]
    hdr_fill = fill("1F3864")
    hdr_font = font(bold=True, color="FFFFFF", size=10)
    for col_idx, h in enumerate(backlog_headers, start=1):
        c = ws.cell(8, col_idx, h)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = center()

    ws.row_dimensions[8].height = 20
    ws.column_dimensions["A"].width = 20
    for col_idx in range(2, 11):
        ws.column_dimensions[col_letter(col_idx)].width = 14

    # One row per station (rows 9-21)
    for row_idx, station in enumerate(STATIONS, start=9):
        sc = station_cols(station)
        st_status_col = col_letter(sc["Status"])
        st_hrs_col = col_letter(sc["Est Hrs"])
        st_col_ref = f"{ms}!${st_status_col}$6:${st_status_col}$500"
        hrs_col_ref = f"{ms}!${st_hrs_col}$6:${st_hrs_col}$500"
        due_col_ref = f"{ms}!${due_col}$6:${due_col}$500"
        job_col_ref = f"{ms}!$A$6:$A$500"

        ws.cell(row_idx, 1, station)
        # Jobs queued (not complete, has job id)
        ws.cell(row_idx, 2,
            f'=SUMPRODUCT(({st_col_ref}<>"Complete")*({st_col_ref}<>"Not Started")*({job_col_ref}<>""))'
        )
        # Remaining hrs (not complete)
        ws.cell(row_idx, 3,
            f'=SUMPRODUCT(({st_col_ref}<>"Complete")*({job_col_ref}<>"")*{hrs_col_ref})'
        )
        # Urgent/High/Normal/Low queued
        priority_col_ref = f"{ms}!${col_letter(ID_COLS['Job Priority'])}$6:${col_letter(ID_COLS['Job Priority'])}$500"
        for p_idx, priority in enumerate(PRIORITIES, start=4):
            ws.cell(row_idx, p_idx,
                f'=SUMPRODUCT(({st_col_ref}<>"Complete")*({st_col_ref}<>"Not Started")*({priority_col_ref}="{priority}")*({job_col_ref}<>""))'
            )
        # Overdue
        ws.cell(row_idx, 8,
            f'=SUMPRODUCT(({st_col_ref}<>"Complete")*({st_col_ref}<>"Not Started")*({due_col_ref}<0)*({job_col_ref}<>""))'
        )
        # Due this week
        ws.cell(row_idx, 9,
            f'=SUMPRODUCT(({st_col_ref}<>"Complete")*({st_col_ref}<>"Not Started")*({due_col_ref}>=0)*({due_col_ref}<=7)*({job_col_ref}<>""))'
        )
        # % complete
        ws.cell(row_idx, 10,
            f'=IFERROR(SUMPRODUCT(({st_col_ref}="Complete")*({job_col_ref}<>""))/SUMPRODUCT(({job_col_ref}<>"")*1),0)'
        ).number_format = "0%"

        # Alternate row fill
        if row_idx % 2 == 0:
            for c_idx in range(1, 11):
                ws.cell(row_idx, c_idx).fill = fill("EEF2F9")

    # Totals row
    totals_row = 9 + len(STATIONS)
    ws.cell(totals_row, 1, "TOTAL").font = font(bold=True)
    for col_idx in range(2, 10):
        start_row = 9
        end_row = totals_row - 1
        ws.cell(totals_row, col_idx,
            f"=SUM({col_letter(col_idx)}{start_row}:{col_letter(col_idx)}{end_row})"
        ).font = font(bold=True)

    # Conditional: red if Overdue > 0
    overdue_range = f"H9:H{totals_row-1}"
    ws.conditional_formatting.add(
        overdue_range,
        FormulaRule(formula=["$H9>0"], fill=fill("FCE4D6"), font=Font(bold=True, color="C00000"))
    )


# ---------------------------------------------------------------------------
# Queue — All sheet
# ---------------------------------------------------------------------------

def build_queue_all(wb):
    ws = wb["Queue — All"]
    ws.freeze_panes = "A8"

    ms = "'Master Schedule'"

    # Title
    ws.merge_cells("A1:K1")
    c = ws["A1"]
    c.value = "QUEUE — ALL STATIONS"
    c.fill = fill("2E75B6")
    c.font = font(bold=True, color="FFFFFF", size=14)
    c.alignment = center()
    ws.row_dimensions[1].height = 28

    # Instruction
    ws.merge_cells("A2:K2")
    ws["A2"].value = "Use dropdowns to filter by Work Centre and Priority"
    ws["A2"].font = font(italic=True, color="595959")
    ws["A2"].alignment = left()

    # Filter labels row 4
    ws["A4"].value = "Work Centre:"
    ws["A4"].font = font(bold=True)
    ws["C4"].value = "Priority:"
    ws["C4"].font = font(bold=True)

    # Filter dropdowns (B4, D4)
    ws["B4"].value = "All"
    ws["D4"].value = "All"

    wc_dv = DataValidation(type="list", formula1='"All,Profiling,Saw,Machining,Fabrication,Weld,Blast,Paint,Assembly"', allow_blank=False)
    wc_dv.sqref = "B4"
    ws.add_data_validation(wc_dv)

    p_dv = DataValidation(type="list", formula1='"All,Urgent,High,Normal,Low"', allow_blank=False)
    p_dv.sqref = "D4"
    ws.add_data_validation(p_dv)

    # Row 6 spacer
    ws.row_dimensions[6].height = 6

    # Headers row 7
    q_headers = [
        "#", "Job ID", "Customer", "Product/Description", "Drawing Ref",
        "Job Priority", "Overall Status", "Order Due Date",
        "Days Until Due", "Assigned Station", "Notes",
    ]
    hdr_fill = fill("2E75B6")
    for col_idx, h in enumerate(q_headers, start=1):
        c = ws.cell(7, col_idx, h)
        c.fill = hdr_fill
        c.font = font(bold=True, color="FFFFFF", size=10)
        c.alignment = center()
    ws.row_dimensions[7].height = 20

    # Column widths
    widths = [5, 12, 20, 28, 14, 12, 14, 13, 12, 18, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[col_letter(i)].width = w

    # Build the FILTER formula
    # We need to pull from tblJobs: Job ID, Customer, Product, Drawing, Priority, Overall Status, Due Date, Days Until Due, Notes
    # Plus "Assigned Station" — most recent active station for a job.
    # Since FILTER is a dynamic array formula, place in A8.
    # Columns in Master Schedule:
    job_id_col  = col_letter(ID_COLS["Job ID"])
    cust_col    = col_letter(ID_COLS["Customer"])
    prod_col    = col_letter(ID_COLS["Product/Description"])
    draw_col    = col_letter(ID_COLS["Drawing Ref"])
    prio_col    = col_letter(ID_COLS["Job Priority"])
    stat_col    = col_letter(ID_COLS["Overall Status"])
    due_date_col = col_letter(ID_COLS["Order Due Date"])
    days_col    = col_letter(ID_COLS["Days Until Due"])
    rank_col    = col_letter(ID_COLS["Priority Rank"])
    notes_col   = col_letter(ID_COLS["Notes"])

    # Assigned station: for simplicity use Overall Status (can't dynamically calc in formula without helper)
    # We'll include a helper column on Config or just use the overall status column as proxy.
    # The spec says "Assigned Station" = current active station. We'll approximate with a XLOOKUP/IF chain as a formula.
    # For the FILTER formula, construct an HSTACK of columns:

    filter_formula = (
        f'=IFERROR('
        f'SORTBY('
        f'HSTACK('
        f'SEQUENCE(ROWS(FILTER({ms}!${job_id_col}$6:${job_id_col}$500,({ms}!$A$6:$A$500<>"")*'
        f'IF($B$4="All",1,ISNUMBER(SEARCH($B$4,{ms}!$A$6:$A$500)))*'  # WC filter placeholder
        f'IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)'
        f'))),'
        f'FILTER({ms}!${job_id_col}$6:${job_id_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),'
        f'FILTER({ms}!${cust_col}$6:${cust_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),'
        f'FILTER({ms}!${prod_col}$6:${prod_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),'
        f'FILTER({ms}!${draw_col}$6:${draw_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),'
        f'FILTER({ms}!${prio_col}$6:${prio_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),'
        f'FILTER({ms}!${stat_col}$6:${stat_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),'
        f'FILTER({ms}!${due_date_col}$6:${due_date_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),'
        f'FILTER({ms}!${days_col}$6:${days_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),'
        f'FILTER({ms}!${stat_col}$6:${stat_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),'
        f'FILTER({ms}!${notes_col}$6:${notes_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4))'
        f'),'
        f'FILTER({ms}!${rank_col}$6:${rank_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),1,'
        f'FILTER({ms}!${due_date_col}$6:${due_date_col}$500,({ms}!$A$6:$A$500<>"")*IF($D$4="All",1,{ms}!${prio_col}$6:${prio_col}$500=$D$4)),1'
        f'),"No jobs match")'
    )

    ws["A8"].value = filter_formula


# ---------------------------------------------------------------------------
# Individual Station Queue sheets
# ---------------------------------------------------------------------------

def build_station_queues(wb):
    ms = "'Master Schedule'"
    status_col_ms = col_letter(ID_COLS["Overall Status"])
    days_col_ms = col_letter(ID_COLS["Days Until Due"])
    rank_col_ms = col_letter(ID_COLS["Priority Rank"])

    for station in STATIONS:
        sheet_name = f"Queue — {station}"
        ws = wb[sheet_name]
        ws.freeze_panes = "A8"

        tab_colour = STATION_COLOURS[station]
        sc = station_cols(station)
        st_status_col = col_letter(sc["Status"])
        st_hrs_col = col_letter(sc["Est Hrs"])
        st_op_col = col_letter(sc["Operator"])
        st_notes_col = col_letter(sc["Notes"])

        # Title row 1
        ws.merge_cells("A1:L1")
        c = ws["A1"]
        c.value = f"QUEUE — {station.upper()}"
        c.fill = fill(tab_colour)
        c.font = font(bold=True, color="FFFFFF", size=14)
        c.alignment = center()
        ws.row_dimensions[1].height = 28

        # Stats bar rows 2-3
        stats = [
            ("Jobs in Queue",  f'=SUMPRODUCT(({ms}!${st_status_col}$6:${st_status_col}$500<>"Not Started")*({ms}!${st_status_col}$6:${st_status_col}$500<>"Complete")*({ms}!$A$6:$A$500<>""))'),
            ("Remaining Hrs",  f'=SUMPRODUCT(({ms}!${st_status_col}$6:${st_status_col}$500<>"Complete")*({ms}!$A$6:$A$500<>"")*{ms}!${st_hrs_col}$6:${st_hrs_col}$500)'),
            ("Overdue",        f'=SUMPRODUCT(({ms}!${st_status_col}$6:${st_status_col}$500<>"Not Started")*({ms}!${st_status_col}$6:${st_status_col}$500<>"Complete")*({ms}!${days_col_ms}$6:${days_col_ms}$500<0)*({ms}!$A$6:$A$500<>""))'),
            ("Urgent",         f'=SUMPRODUCT(({ms}!${st_status_col}$6:${st_status_col}$500<>"Complete")*({ms}!${col_letter(ID_COLS["Job Priority"])}$6:${col_letter(ID_COLS["Job Priority"])}$500="Urgent")*({ms}!$A$6:$A$500<>""))'),
        ]

        stat_cols_layout = [1, 4, 7, 10]
        for (label, formula), start_col in zip(stats, stat_cols_layout):
            ws.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=start_col+2)
            ws.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=start_col+2)
            lbl = ws.cell(2, start_col, label)
            lbl.fill = fill(tab_colour)
            lbl.font = font(bold=True, color="FFFFFF", size=9)
            lbl.alignment = center()
            val = ws.cell(3, start_col, formula)
            val.fill = fill("FFFFFF")
            val.font = Font(bold=True, size=22, color=tab_colour)
            val.alignment = center()

        ws.row_dimensions[2].height = 16
        ws.row_dimensions[3].height = 36

        # Spacer rows 4-6
        ws.row_dimensions[4].height = 6
        ws.row_dimensions[5].height = 6
        ws.row_dimensions[6].height = 6

        # Headers row 7
        q_headers = [
            "#", "Job ID", "Customer", "Product/Description", "Drawing Ref",
            "Job Priority", "Order Due Date", "Days Until Due",
            f"{station} Status", f"{station} Est Hrs",
            f"{station} Operator", f"{station} Notes",
        ]
        hdr_fill = fill(tab_colour)
        for col_idx, h in enumerate(q_headers, start=1):
            c = ws.cell(7, col_idx, h)
            c.fill = hdr_fill
            c.font = font(bold=True, color="FFFFFF", size=10)
            c.alignment = center()
        ws.row_dimensions[7].height = 20

        # Column widths
        widths = [5, 12, 20, 28, 14, 12, 13, 12, 14, 10, 14, 28]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[col_letter(i)].width = w

        # FILTER formula in A8 — show only jobs where this station status ≠ Not Started AND ≠ Complete
        # unless status = Not Started and job has not yet reached this station
        # Simplified: show all jobs with station status not "Not Started"
        job_id_col_ms  = col_letter(ID_COLS["Job ID"])
        cust_col_ms    = col_letter(ID_COLS["Customer"])
        prod_col_ms    = col_letter(ID_COLS["Product/Description"])
        draw_col_ms    = col_letter(ID_COLS["Drawing Ref"])
        prio_col_ms    = col_letter(ID_COLS["Job Priority"])
        due_date_col_ms = col_letter(ID_COLS["Order Due Date"])

        include_cond = (
            f'({ms}!${st_status_col}$6:${st_status_col}$500<>"Not Started")*'
            f'({ms}!$A$6:$A$500<>"")'
        )

        filter_formula = (
            f'=IFERROR('
            f'SORTBY('
            f'HSTACK('
            f'SEQUENCE(ROWS(FILTER({ms}!${job_id_col_ms}$6:${job_id_col_ms}$500,{include_cond}))),'
            f'FILTER({ms}!${job_id_col_ms}$6:${job_id_col_ms}$500,{include_cond}),'
            f'FILTER({ms}!${cust_col_ms}$6:${cust_col_ms}$500,{include_cond}),'
            f'FILTER({ms}!${prod_col_ms}$6:${prod_col_ms}$500,{include_cond}),'
            f'FILTER({ms}!${draw_col_ms}$6:${draw_col_ms}$500,{include_cond}),'
            f'FILTER({ms}!${prio_col_ms}$6:${prio_col_ms}$500,{include_cond}),'
            f'FILTER({ms}!${due_date_col_ms}$6:${due_date_col_ms}$500,{include_cond}),'
            f'FILTER({ms}!${days_col_ms}$6:${days_col_ms}$500,{include_cond}),'
            f'FILTER({ms}!${st_status_col}$6:${st_status_col}$500,{include_cond}),'
            f'FILTER({ms}!${st_hrs_col}$6:${st_hrs_col}$500,{include_cond}),'
            f'FILTER({ms}!${st_op_col}$6:${st_op_col}$500,{include_cond}),'
            f'FILTER({ms}!${st_notes_col}$6:${st_notes_col}$500,{include_cond})'
            f'),'
            f'FILTER({ms}!${rank_col_ms}$6:${rank_col_ms}$500,{include_cond}),1,'
            f'FILTER({ms}!${due_date_col_ms}$6:${due_date_col_ms}$500,{include_cond}),1'
            f'),"No jobs in queue")'
        )

        ws["A8"].value = filter_formula


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Building Production_Schedule.xlsx ...")
    wb = make_workbook()

    print("  [1/6] Config sheet ...")
    build_config(wb)

    print("  [2/6] Master Schedule sheet ...")
    build_master_schedule(wb)

    print("  [3/6] Dashboard sheet ...")
    build_dashboard(wb)

    print("  [4/6] Queue — All sheet ...")
    build_queue_all(wb)

    print("  [5/6] Station queue sheets ...")
    build_station_queues(wb)

    output_path = "Production_Schedule.xlsx"
    print(f"  [6/6] Saving to {output_path} ...")
    wb.save(output_path)
    print("Done.")


if __name__ == "__main__":
    main()
