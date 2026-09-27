"""Builds the Excel tracker sheet from the ticket list."""

from datetime import datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

STATUS_COLORS = {
    "To Do": "D9D9D9", "In Progress": "BDD7EE", "In Review": "FFE699",
    "Blocked": "F4B084", "Done": "C6EFCE",
}
PRIORITY_COLORS = {
    "Critical": "C00000", "High": "FF7C80", "Medium": "FFD966", "Low": "A9D08E",
}

COLUMNS = [  # (header, ticket key, width)
    ("Ticket ID", "id", 11), ("Title", "title", 38), ("Type", "type", 13),
    ("Priority", "priority", 11), ("Status", "status", 13), ("Assignee", "assignee", 14),
    ("Created", "created", 12), ("Due Date", "due", 12), ("Overdue?", None, 10),
    ("Description", "description", 45),
]


def _header(ws, row, headers):
    for col, text in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=text)
        c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")


def _tickets_sheet(ws, tickets, statuses, priorities):
    ws.title = "Tickets"
    _header(ws, 1, [c[0] for c in COLUMNS])
    for i, (_, _, width) in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    for r, t in enumerate(tickets, 2):
        for c, (_, key, _) in enumerate(COLUMNS, 1):
            if key in ("created", "due") and t.get(key):
                value = datetime.strptime(t[key], "%Y-%m-%d")
            elif key is None:  # Overdue? -> live formula
                value = f'=IF(AND(E{r}<>"Done",H{r}<>"",H{r}<TODAY()),"YES","")'
            else:
                value = t.get(key, "")
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(key in ("title", "description")))
            if key in ("created", "due"):
                cell.number_format = "DD-MMM-YYYY"

    last = max(len(tickets) + 1, 2)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{last}"

    # Dropdowns so the sheet can be edited directly (extra rows for new tickets)
    edit_to = last + 100
    for col, options in (("D", priorities), ("E", statuses)):
        dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{edit_to}")

    # Colour coding that updates when a value changes
    for text, color in STATUS_COLORS.items():
        ws.conditional_formatting.add(f"E2:E{edit_to}", FormulaRule(
            formula=[f'$E2="{text}"'], fill=PatternFill("solid", fgColor=color)))
    for text, color in PRIORITY_COLORS.items():
        font = Font(bold=True, color="FFFFFF" if text == "Critical" else "000000")
        ws.conditional_formatting.add(f"D2:D{edit_to}", FormulaRule(
            formula=[f'$D2="{text}"'], fill=PatternFill("solid", fgColor=color), font=font))
    ws.conditional_formatting.add(f"I2:I{edit_to}", FormulaRule(
        formula=['$I2="YES"'], font=Font(bold=True, color="C00000")))
    return edit_to


def _summary_sheet(wb, tickets, statuses, priorities, rng_end):
    ws = wb.create_sheet("Summary", 0)
    ws.column_dimensions["A"].width = 22
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 13

    ws["A1"] = "Project Ticket Tracker"
    ws["A1"].font = Font(bold=True, size=16, color="1F3864")
    ws["A2"] = "Generated " + datetime.now().strftime("%d %b %Y, %H:%M")
    ws["A2"].font = Font(italic=True, color="808080")

    T = f"Tickets!$A$2:$A${rng_end}"
    S = f"Tickets!$E$2:$E${rng_end}"
    P = f"Tickets!$D$2:$D${rng_end}"
    A = f"Tickets!$F$2:$F${rng_end}"
    O = f"Tickets!$I$2:$I${rng_end}"

    # KPI block
    kpis = [
        ("Total tickets", f"=COUNTA({T})"),
        ("Open (not Done)", f'=COUNTA({T})-COUNTIF({S},"Done")'),
        ("Done", f'=COUNTIF({S},"Done")'),
        ("Overdue", f'=COUNTIF({O},"YES")'),
        ("Completion %", f'=IFERROR(COUNTIF({S},"Done")/COUNTA({T}),0)'),
    ]
    _header(ws, 4, ["Key metric", "Value"])
    for i, (label, formula) in enumerate(kpis, 5):
        ws.cell(row=i, column=1, value=label).border = BORDER
        c = ws.cell(row=i, column=2, value=formula)
        c.border, c.font = BORDER, Font(bold=True)
        if label == "Completion %":
            c.number_format = "0%"

    # By status
    row = 11
    _header(ws, row, ["By status", "Count"])
    status_first = row + 1
    for i, s in enumerate(statuses, status_first):
        ws.cell(row=i, column=1, value=s).border = BORDER
        ws.cell(row=i, column=2, value=f'=COUNTIF({S},A{i})').border = BORDER
    status_last = status_first + len(statuses) - 1

    # By priority
    row = status_last + 2
    _header(ws, row, ["By priority", "Count"])
    prio_first = row + 1
    for i, p in enumerate(priorities, prio_first):
        ws.cell(row=i, column=1, value=p).border = BORDER
        ws.cell(row=i, column=2, value=f'=COUNTIF({P},A{i})').border = BORDER

    # Workload per assignee (status breakdown)
    row = prio_first + len(priorities) + 1
    _header(ws, row, ["Assignee workload"] + statuses + ["Total"])
    assignees = sorted({t["assignee"] for t in tickets})
    for i, name in enumerate(assignees, row + 1):
        ws.cell(row=i, column=1, value=name).border = BORDER
        for j, s in enumerate(statuses, 2):
            ws.cell(row=i, column=j, value=f'=COUNTIFS({A},$A{i},{S},"{s}")').border = BORDER
        tot = ws.cell(row=i, column=len(statuses) + 2,
                      value=f"=SUM(B{i}:{get_column_letter(len(statuses) + 1)}{i})")
        tot.border, tot.font = BORDER, Font(bold=True)

    # Charts
    pie = PieChart()
    pie.title = "Tickets by status"
    pie.add_data(Reference(ws, min_col=2, min_row=status_first - 1, max_row=status_last), titles_from_data=True)
    pie.set_categories(Reference(ws, min_col=1, min_row=status_first, max_row=status_last))
    pie.height, pie.width = 7, 11
    ws.add_chart(pie, "I4")

    bar = BarChart()
    bar.title = "Tickets by priority"
    bar.legend = None
    bar.add_data(Reference(ws, min_col=2, min_row=prio_first - 1, max_row=prio_first + len(priorities) - 1),
                 titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=prio_first, max_row=prio_first + len(priorities) - 1))
    bar.height, bar.width = 7, 11
    ws.add_chart(bar, "I19")


def build_tracker(tickets, out_path, statuses, priorities, is_overdue=None):
    wb = Workbook()
    rng_end = _tickets_sheet(wb.active, tickets, statuses, priorities)
    _summary_sheet(wb, tickets, statuses, priorities, rng_end)
    wb.save(out_path)
    return out_path
