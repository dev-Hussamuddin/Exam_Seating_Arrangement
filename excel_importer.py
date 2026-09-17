"""Validate one Excel workbook and import it in a single MySQL transaction."""

import re
from datetime import date, datetime, time

from openpyxl import load_workbook

from config import ACADEMIC_YEAR, INPUT_DIR
from database import get_connection, migrate_input_schema


HEADERS = {
    "Class Data": ["Class", "Department", "Roll Nos. (Range)",
                   "No. of Students", "Non-included Roll Nos. (Left/Debarred)"],
    "Classroom Data": ["Classroom No.", "Capacity (No. of Benches)"],
    "Timetable": ["Class", "Subject", "Semester", "Exam Date", "Start Time", "End Time"],
}

COLLEGE_HEADERS = {
    "Class Data": ["Class", "Subject", "Roll nos.", "excluding nos.", "No. of Students"],
    "Classroom Data": ["Room No.", "Capacity (No. of benches)"],
    "Timetable": ["Class", "Subject", "Time", "Exam Date"],
}


def parse_roll_list(value, label="Roll nos.", allow_empty=False):
    """Accept individual rolls and comma-separated ranges without filling gaps."""
    if value is None or str(value).strip() == "":
        if allow_empty:
            return []
        raise ValueError(f"{label} is required.")
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    result = set()
    # College sheets use trailing separators ("1 - 63,"), spaces and en dashes.
    # Remove only trailing separators; malformed interior entries still fail.
    value = str(value).replace("\u2013", "-").replace("\u2014", "-")
    value = value.rstrip(" ,;\r\n\t")
    if not value:
        raise ValueError(f"{label} must contain at least one roll number.")
    for segment in re.split(r"[,;\n]", value):
        match = re.fullmatch(r"\s*(\d+)\s*(?:-\s*(\d+))?\s*", segment)
        if not match:
            raise ValueError(f"{label} must contain numbers/ranges, such as 1-4,6,8-10.")
        start = integer(match[1], label)
        end = integer(match[2], label) if match[2] else start
        if start > end:
            raise ValueError(f"{label}: range start must not exceed range end.")
        result.update(range(start, end + 1))
    return sorted(result)


def parse_time_value(value):
    """Preserve source text and resolve omitted end-time AM/PM within the day.

    Unmarked starts retain their written hour (12 means noon). An unmarked
    end before the start advances by 12 hours when that produces a same-day
    interval. Explicit AM/PM and equal endpoints are never silently changed.
    A single Excel/text time remains a start time with no invented duration.
    """
    if isinstance(value, time):
        if value.tzinfo is not None:
            raise ValueError("Time must be a local exam time without a timezone.")
        value = value.replace(microsecond=0).isoformat()
    raw = text(value, "Time", 255)
    parts = re.split(r"\s*(?:-|\u2013|\u2014|\bto\b)\s*", raw, flags=re.I)
    if len(parts) not in (1, 2):
        raise ValueError("Time must be a time or a start-end range, such as 12:00-01:00.")
    def parse(part):
        for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p", "%I %p", "%I%p"):
            try:
                return datetime.strptime(part.strip().upper(), fmt).time()
            except ValueError:
                pass
        return None
    start = parse(parts[0])
    end = parse(parts[1]) if len(parts) == 2 else None
    if start is None or (len(parts) == 2 and end is None):
        raise ValueError("Invalid Time. Use Excel time cells or a range such as 12:00-01:00.")
    if end is not None and end < start and not re.search(r"[AP]M", parts[1], re.I):
        if 1 <= end.hour <= 11:
            end = end.replace(hour=end.hour + 12)
    if start is not None and end is not None and start >= end:
        raise ValueError("End Time must be later than Start Time.")
    return raw, start, end


def text(value, label, limit, required=True):
    value = "" if value is None else str(value).strip()
    if required and not value:
        raise ValueError(f"{label} is required.")
    if len(value) > limit:
        raise ValueError(f"{label} must be at most {limit} characters.")
    return value or None


def integer(value, label, minimum=0):
    if isinstance(value, bool):
        raise ValueError(f"{label} must be a whole number.")
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if not re.fullmatch(r"\d+", str(value).strip()):
        raise ValueError(f"{label} must be a whole number.")
    value = int(value)
    if not minimum <= value <= 2147483647:
        raise ValueError(f"{label} must be between {minimum} and 2147483647.")
    return value


def exam_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(value).strip()):
        raise ValueError("Exam Date must be an Excel date or YYYY-MM-DD.")
    try:
        return datetime.strptime(str(value).strip(), "%Y-%m-%d").date()
    except ValueError:
        raise ValueError("Exam Date must be an Excel date or YYYY-MM-DD.") from None


def exam_time(value):
    if isinstance(value, time) and value.tzinfo is None:
        return value.replace(microsecond=0)
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(str(value).strip(), fmt).time()
        except ValueError:
            pass
    raise ValueError("Start/End Time must be an Excel time or HH:MM (24-hour).")


def normalized_label(value):
    return " ".join(str(value or "").split()).casefold()


def locate_tables(workbook):
    """Find titles and header rows in non-overlapping side-by-side column bands.

    Titles may be centered over their tables. Headers may span nearby rows.
    All titles share one sheet. Workbook cells are never modified.
    """
    candidates = []
    title_names = {normalized_label(name): name for name in HEADERS}
    title_names.update({
        "class table": "Class Data",
        "classroom table": "Classroom Data",
        "time table": "Timetable",
    })
    for sheet in workbook.worksheets:
        rows = [list(row) for row in sheet.iter_rows(values_only=True)]
        titles = {name: [] for name in HEADERS}
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                name = title_names.get(normalized_label(value))
                if name:
                    titles[name].append((r, c))
        if all(titles.values()):
            candidates.append((sheet, rows, titles))
    if len(candidates) != 1:
        raise ValueError("Exactly one worksheet must contain all three table headings: "
                         "Class Data (or Class Table), Classroom Data (or Classroom Table), "
                         "Timetable (or Time table).")
    sheet, rows, titles = candidates[0]
    sheet_name = sheet.title
    if any(len(found) != 1 for found in titles.values()):
        raise ValueError(f"{sheet_name}: each table heading must appear exactly once.")
    ordered = sorted((positions[0][1], positions[0][0], name) for name, positions in titles.items())
    if len({col for col, _, _ in ordered}) != 3:
        raise ValueError("The three tables must be side-by-side in separate columns.")
    # Locate the first header near each title instead of treating the title's
    # column as the table boundary (the real Class title is offset to the right).
    starts = []
    for title_col, title_row, name in ordered:
        anchors = {normalized_label(HEADERS[name][0]), normalized_label(COLLEGE_HEADERS[name][0])}
        found = [(abs(c - title_col), r, c) for r in range(title_row + 1, min(len(rows), title_row + 13))
                 for c, value in enumerate(rows[r]) if normalized_label(value) in anchors]
        if not found:
            raise ValueError(f"{name}: required column headers were not found near its heading.")
        _, _, anchor_col = min(found)
        # Include reordered headers to the left of the first named column.
        known = {normalized_label(h) for h in HEADERS[name] + COLLEGE_HEADERS[name]}
        left = anchor_col
        while left > 0 and any(normalized_label(rows[r][left - 1]) in known
                               for r in range(title_row + 1, min(len(rows), title_row + 13))):
            left -= 1
        starts.append((left, title_row, name))
    starts.sort()
    if len({left for left, _, _ in starts}) != 3:
        raise ValueError("Could not distinguish the three side-by-side table header areas.")
    tables = {}
    for index, (left, title_row, name) in enumerate(starts):
        right = starts[index + 1][0] if index + 1 < len(starts) else max(map(len, rows))
        matches = []
        for layout, headers in (("college", COLLEGE_HEADERS[name]), ("legacy", HEADERS[name])):
            positions = []
            for header in headers:
                found = [(r, c) for r in range(title_row + 1, min(len(rows), title_row + 13))
                         for c in range(left, right) if normalized_label(rows[r][c]) == normalized_label(header)]
                if len(found) > 1:
                    raise ValueError(f"{name}: duplicate column header '{header}'.")
                if not found:
                    break
                positions.append(found[0])
            if len(positions) == len(headers):
                bottom = max(r for r, _ in positions)
                for merged in sheet.merged_cells.ranges:
                    if (merged.min_row - 1, merged.min_col - 1) in positions:
                        bottom = max(bottom, merged.max_row - 1)
                columns = [c for _, c in positions]
                if len(set(columns)) != len(columns):
                    raise ValueError(f"{name}: each header must identify a separate column.")
                matches.append((bottom, columns, layout))
        if len(matches) != 1:
            raise ValueError(f"{name}: expected nearby column headers below its heading containing: "
                             + ", ".join(COLLEGE_HEADERS[name]))
        header_row, columns, layout = matches[0]
        table_rows = [row[:] for row in rows]
        continuation_rows = set()
        for merged in sheet.merged_cells.ranges:
            if (merged.min_col == merged.max_col and merged.min_col - 1 in columns
                    and merged.min_row > header_row + 1):
                col = merged.min_col - 1
                value = rows[merged.min_row - 1][col]
                for row in range(merged.min_row, merged.max_row):
                    table_rows[row][col] = value
                    if name == "Class Data" and col == columns[1]:
                        continuation_rows.add(row)
        # A vertically merged Subject describes one record, not repeated records.
        for row in continuation_rows:
            if all(rows[row][col] is None for col in columns):
                for col in columns:
                    table_rows[row][col] = None
        tables[name] = (table_rows, header_row, columns, layout)
    return tables


def read_workbook(path):
    """Validate all rows before opening a database connection."""
    records = {name: [] for name in HEADERS}
    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        tables = locate_tables(workbook)
        for name in HEADERS:
            rows, header_row, indexes, layout = tables[name]
            seen = set()
            previous_class = None
            for row_number, row in enumerate(rows[header_row + 1:], header_row + 2):
                values = [row[i] if i < len(row) else None for i in indexes]
                if all(v is None or str(v).strip() == "" for v in values):
                    previous_class = None
                    continue
                try:
                    if any(isinstance(v, str) and v.startswith("=") for v in values):
                        raise ValueError("Use values, not formulas, in input cells.")
                    if layout == "college":
                        if name in ("Class Data", "Timetable"):
                            if values[0] is None or str(values[0]).strip() == "":
                                values[0] = previous_class
                            cls = text(values[0], "Class", 50)
                            previous_class = cls
                            subject = text(values[1], "Subject", 150)
                        if name == "Class Data":
                            rolls = parse_roll_list(values[2])
                            omitted = parse_roll_list(values[3], "excluding nos.", allow_empty=True)
                            if not set(omitted) <= set(rolls):
                                raise ValueError("Excluded rolls must belong to this subject's roll list.")
                            excluded = text(",".join(map(str, omitted)), "excluding nos.", 255, required=False)
                            count = integer(values[4], "No. of Students")
                            record = (cls, None, rolls[0], rolls[-1], count, excluded,
                                      subject, ",".join(map(str, rolls)))
                            key = (cls.casefold(), subject.casefold())
                        elif name == "Classroom Data":
                            room = text(values[0], "Room No.", 20)
                            record = (room, integer(values[1], "Capacity", minimum=1))
                            key = room.casefold()
                        else:
                            raw, start, end = parse_time_value(values[2])
                            record = (cls, subject, None, exam_date(values[3]), start, end, raw)
                            key = None
                        if key is not None and key in seen:
                            raise ValueError("Duplicate class/subject or classroom row in this table.")
                        seen.add(key)
                        records[name].append((row_number, record))
                        continue
                    if name == "Class Data":
                        cls = text(values[0], "Class", 50)
                        dept = text(values[1], "Department", 100)
                        match = re.fullmatch(r"\s*(\d+)\s*-\s*(\d+)\s*", str(values[2]))
                        if not match:
                            raise ValueError("Roll Nos. (Range) must look like 1-60.")
                        start, end = [integer(v, "Roll number") for v in match.groups()]
                        if start > end:
                            raise ValueError("Roll start must be less than or equal to roll end.")
                        count = integer(values[3], "No. of Students")
                        excluded = text(values[4], "Non-included rolls", 255, required=False)
                        if excluded:
                            rolls = sorted({integer(v.strip(), "Excluded roll") for v in excluded.split(",")})
                            if any(r < start or r > end for r in rolls):
                                raise ValueError("Excluded rolls must lie within the roll range.")
                            excluded = ",".join(map(str, rolls))
                        record = (cls, dept, start, end, count, excluded)
                        key = cls.casefold()
                    elif name == "Classroom Data":
                        room = text(values[0], "Classroom No.", 20)
                        record = (room, integer(values[1], "Capacity", minimum=1))
                        key = room.casefold()
                    else:
                        cls = text(values[0], "Class", 50)
                        subject = text(values[1], "Subject", 150)
                        semester = text(values[2], "Semester", 50, required=False)
                        day = exam_date(values[3])
                        start, end = exam_time(values[4]), exam_time(values[5])
                        if start >= end:
                            raise ValueError("End Time must be later than Start Time.")
                        record = (cls, subject, semester, day, start, end)
                        key = None
                    if key is not None and key in seen:
                        raise ValueError("Duplicate class or classroom row in this table.")
                    seen.add(key)
                    records[name].append((row_number, record))
                except ValueError as error:
                    raise ValueError(f"{name}, row {row_number}: {error}") from None
    finally:
        workbook.close()
    return records


def import_workbook(path):
    """One batch per class/subject/year; identical timetable rows are skipped.

    Intended for a single local importer at a time. Existing schema is unchanged.
    Student counts are stored as supplied, without subtracting excluded rolls.
    """
    year = text(ACADEMIC_YEAR, "ACADEMIC_YEAR in config.py", 20)
    records = read_workbook(path)
    counts = dict(classes=0, batches=0, classrooms=0, timetable=0)
    notices = set()
    connection = get_connection()
    cursor = None
    try:
        migrate_input_schema(connection)
        connection.start_transaction()
        cursor = connection.cursor(buffered=True)
        for row_number, record in records["Class Data"]:
            cls, dept, start, end, count, excluded = record[:6]
            subject, roll_numbers = record[6:] if len(record) == 8 else (None, None)
            cursor.execute("SELECT class_id, department FROM classes WHERE class_name = %s", (cls,))
            existing = cursor.fetchone()
            if existing:
                class_id, existing_dept = existing
                if dept is not None and existing_dept.casefold() != dept.casefold():
                    raise ValueError(f"Class Data, row {row_number}: {cls} already belongs to {existing_dept}.")
            else:
                # Empty department means unknown; do not infer it from a class name.
                cursor.execute("INSERT INTO classes (class_name, department) VALUES (%s, %s)", (cls, dept or ""))
                class_id = cursor.lastrowid
                counts["classes"] += 1

            cursor.execute("""SELECT batch_id FROM student_batches WHERE class_id = %s AND academic_year = %s
                AND COALESCE(subject, '')=COALESCE(%s, '')""", (class_id, year, subject))
            batches = cursor.fetchall()
            if len(batches) > 1:
                raise ValueError(f"Multiple existing batches for {cls} in {year}; resolve them before importing.")
            if batches:
                cursor.execute("""UPDATE student_batches SET roll_start=%s, roll_end=%s,
                    number_of_students=%s, non_included_rolls=%s, roll_numbers=%s WHERE batch_id=%s""",
                               (start, end, count, excluded, roll_numbers, batches[0][0]))
            else:
                cursor.execute("""INSERT INTO student_batches
                    (class_id, academic_year, roll_start, roll_end, number_of_students, non_included_rolls, subject, roll_numbers)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)""", (class_id, year, start, end, count, excluded, subject, roll_numbers))
            counts["batches"] += 1

        for _, (room, capacity) in records["Classroom Data"]:
            cursor.execute("SELECT classroom_id FROM classrooms WHERE classroom_no=%s", (room,))
            existing = cursor.fetchone()
            if existing:
                cursor.execute("UPDATE classrooms SET capacity=%s WHERE classroom_id=%s", (capacity, existing[0]))
            else:
                cursor.execute("INSERT INTO classrooms (classroom_no, capacity) VALUES (%s, %s)", (room, capacity))
            counts["classrooms"] += 1

        for row_number, record in records["Timetable"]:
            cls, subject, semester, day, start, end = record[:6]
            raw = record[6] if len(record) == 7 else None
            cursor.execute("SELECT class_id FROM classes WHERE class_name=%s", (cls,))
            existing = cursor.fetchone()
            if not existing:
                # Preserve the timetable foreign key without inventing a student batch.
                cursor.execute("INSERT INTO classes (class_name, department) VALUES (%s, %s)", (cls, ""))
                existing = (cursor.lastrowid,)
                counts["classes"] += 1
            cursor.execute("""SELECT batch_id FROM student_batches
                WHERE class_id=%s AND academic_year=%s AND (subject=%s OR subject IS NULL)""",
                           (existing[0], year, subject))
            if not cursor.fetchone():
                notices.add(f"{cls} / {subject}: no student data for {year}; timetable preserved, skipped for seating.")
            values = (existing[0], subject, semester, day, start, end, raw)
            cursor.execute("""SELECT timetable_id FROM timetable WHERE class_id=%s AND subject=%s
                AND COALESCE(semester, '')=COALESCE(%s, '') AND COALESCE(exam_date, '')=COALESCE(%s, '')
                AND COALESCE(start_time, '')=COALESCE(%s, '') AND COALESCE(end_time, '')=COALESCE(%s, '')
                AND COALESCE(time_text, '')=COALESCE(%s, '')""", values)
            if not cursor.fetchone():
                # Match undated records by the actual exam and parsed time slot.
                # Source time formatting may differ without changing the exam.
                cursor.execute("""SELECT timetable_id FROM timetable
                    WHERE class_id=%s AND subject=%s AND exam_date IS NULL
                    AND COALESCE(semester, '')=COALESCE(%s, '')
                    AND COALESCE(start_time, '')=COALESCE(%s, '')
                    AND COALESCE(end_time, '')=COALESCE(%s, '')""",
                               (existing[0], subject, semester, start, end))
                undated = cursor.fetchall()
                if len(undated) > 1:
                    raise ValueError(f"Timetable, row {row_number}: multiple matching undated records for {cls} / {subject}.")
                if undated:
                    cursor.execute("UPDATE timetable SET exam_date=%s, time_text=%s WHERE timetable_id=%s",
                                   (day, raw, undated[0][0]))
                else:
                    cursor.execute("""INSERT INTO timetable
                        (class_id, subject, semester, exam_date, start_time, end_time, time_text)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)""", values)
                counts["timetable"] += 1
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        if cursor is not None:
            cursor.close()
        connection.close()
    for notice in sorted(notices):
        print(notice)
    return counts


def import_from_input():
    """Select exactly one workbook, ignoring Excel's temporary lock files."""
    INPUT_DIR.mkdir(exist_ok=True)
    files = sorted(p for p in INPUT_DIR.iterdir()
                   if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$"))
    if not files:
        print(f"No .xlsx workbook found. Place one Excel workbook in: {INPUT_DIR}")
        return
    if len(files) > 1:
        raise ValueError("Keep exactly one .xlsx workbook in the input folder before importing.")
    counts = import_workbook(files[0])
    print(f"Imported {files[0].name} for academic year {ACADEMIC_YEAR}.")
    print(f"Classes imported: {counts['classes']}")
    print(f"Student batches imported/updated: {counts['batches']}")
    print(f"Classrooms imported/updated: {counts['classrooms']}")
    print(f"Timetable records imported/updated: {counts['timetable']}")
