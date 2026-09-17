"""Excel presentation of committed allocation summaries; no database writes."""

from collections import defaultdict
from datetime import date, datetime, time, timedelta
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from config import OUTPUT_DIR


def _date(value):
    if isinstance(value, datetime):
        return value.date()
    return date.fromisoformat(value) if isinstance(value, str) else value


def _time(value):
    if isinstance(value, timedelta):
        return (datetime.min + value).time()
    return time.fromisoformat(value) if isinstance(value, str) else value


def write_seating_report(summaries, output_path=None):
    """Export snapshots returned by generate_arrangements after its commit.

    Counts are student-exam attendances across slots, not unique people.
    Capacity used means occupied seats; unused includes wholly unused rooms.
    Room capacity and unused seats appear only on each room's first group row.
    """
    # Local import avoids a module cycle with the command-line entry point.
    from seating_allocator import compact_ranges

    summaries = list(summaries)
    target = Path(output_path) if output_path is not None else OUTPUT_DIR / "Seating_Arrangement.xlsx"
    eligible = sum(len(e["rolls"]) for _, exams, _, _ in summaries for e in exams)
    allocated = sum(len(a["rolls"]) for _, _, _, rows in summaries for a in rows)
    capacity = sum(r["capacity"] for _, _, rooms, _ in summaries for r in rooms)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Seating Arrangement"
    sheet.append(["EXAM SEATING ARRANGEMENT"])
    sheet.merge_cells("A1:J1")
    for label, value in [("Total Eligible Students", eligible),
                         ("Total Allocated Students", allocated),
                         ("Total Capacity Used", allocated),
                         ("Total Unused Seats", capacity - allocated)]:
        sheet.append([label, value])
    sheet.merge_cells("A6:J6")
    sheet["A6"] = ("Totals count student-exam attendances across slots. Capacity used = occupied seats. "
                   "Capacity and unused seats are shown once per room per slot, including empty rooms.")
    sheet.append([])
    sheet.append(["Exam Date", "Exam Time (Start)", "Exam Time (End)", "Room Number",
                  "Room Capacity", "Class", "Subject", "Allocated Roll Numbers",
                  "Number of Students Allocated", "Unused Seats"])
    for slot, exams, rooms, allocations in summaries:
        day, start, end = slot
        exam_by_id = {e["timetable_id"]: e for e in exams}
        by_room = defaultdict(list)
        for allocation in allocations:
            by_room[allocation["classroom_id"]].append(allocation)
        for room in rooms:
            groups = by_room[room["classroom_id"]]
            unused = room["capacity"] - sum(len(a["rolls"]) for a in groups)
            for index, allocation in enumerate(groups or [None]):
                exam = exam_by_id[allocation["timetable_id"]] if allocation else {}
                rolls = allocation["rolls"] if allocation else []
                sheet.append([_date(day), _time(start), _time(end), str(room["classroom_no"]),
                              room["capacity"] if index == 0 else None,
                              exam.get("class_name", ""), exam.get("subject", ""),
                              compact_ranges(rolls), len(rolls), unused if index == 0 else None])
                sheet.cell(sheet.max_row, 1).number_format = "dd-mmm-yyyy"
                for column in (2, 3):
                    sheet.cell(sheet.max_row, column).number_format = "hh:mm"

    navy = "203864"
    border = Border(*( [Side(style="thin", color="D5DCE4")] * 4))
    for row in sheet:
        for cell in row:
            cell.font = Font(name="Calibri", size=11, color="203040")
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            # Imported identifiers and subjects must remain literal Excel text.
            if isinstance(cell.value, str):
                cell.data_type = "s"
            if cell.row >= 8 or 2 <= cell.row <= 5 and cell.column <= 2:
                cell.border = border
            if cell.row in (1, 8):
                cell.fill = PatternFill("solid", fgColor=navy)
                cell.font = Font(name="Calibri", size=14 if cell.row == 1 else 11,
                                 bold=True, color="FFFFFF")
            elif 2 <= cell.row <= 5 and cell.column <= 2:
                cell.font = Font(name="Calibri", size=11, bold=True)
                cell.fill = PatternFill("solid", fgColor="E8EFF7")
    for column, width in zip("ABCDEFGHIJ", (29, 19, 19, 17, 16, 24, 36, 60, 23, 16)):
        sheet.column_dimensions[column].width = width
    sheet.row_dimensions[1].height = 28
    sheet.row_dimensions[6].height = 32
    sheet.row_dimensions[8].height = 34
    for row in range(9, sheet.max_row + 1):
        lines = max((len(str(sheet.cell(row, col).value or "")) + width - 1) // width
                    for col, width in ((6, 22), (7, 34), (8, 55)))
        sheet.row_dimensions[row].height = max(30, 15 * lines)
    sheet.freeze_panes = "F9"
    sheet.sheet_view.showGridLines = False
    sheet.auto_filter.ref = f"A8:J{sheet.max_row}"
    sheet.print_title_rows = "1:8"
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_options.horizontalCentered = True
    sheet.print_area = f"A1:J{sheet.max_row}"

    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        # Replace only a complete workbook, preserving an older report on failure.
        with NamedTemporaryFile(dir=target.parent, suffix=".xlsx", delete=False) as handle:
            temporary = Path(handle.name)
        workbook.save(temporary)
        temporary.replace(target)
    finally:
        workbook.close()
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return target
