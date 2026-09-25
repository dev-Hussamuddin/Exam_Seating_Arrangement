"""Optional one-day workbook layout, normalized to the existing importer records."""

SHEETS = ("Exam Configuration", "Student Subject Data", "Classroom Data")
STUDENT_HEADERS = ("Class", "Department", "Subject", "Roll Number Start",
                   "Roll Number End", "Excluded Roll Numbers")
ROOM_HEADERS = ("Room Number", "Seating Capacity")


def read_one_day_workbook(workbook):
    from excel_importer import (exam_date, integer, normalized_label,
                               parse_roll_list, parse_time_value, text)

    if any(name not in workbook.sheetnames for name in SHEETS):
        raise ValueError("One-day input requires sheets: " + ", ".join(SHEETS))
    records = {"Class Data": [], "Classroom Data": [], "Timetable": []}

    def values(sheet):
        for number, row in enumerate(sheet.iter_rows(values_only=True), 1):
            if any(isinstance(v, str) and v.startswith("=") for v in row):
                raise ValueError(f"{sheet.title}, row {number}: use values, not formulas.")
            if any(v is not None and str(v).strip() for v in row):
                yield number, row

    settings = {}
    for number, row in values(workbook[SHEETS[0]]):
        key = normalized_label(row[0])
        if key not in ("exam date", "exam time") or key in settings:
            raise ValueError(f"Exam Configuration, row {number}: unknown or duplicate setting.")
        settings[key] = row[1] if len(row) > 1 else None
    day = exam_date(settings.get("exam date"))
    raw = settings.get("exam time")
    if raw is None or not str(raw).strip():
        raw, start, end = None, None, None
    else:
        raw, start, end = parse_time_value(raw)
        if end is None:
            raise ValueError("Exam Time must be a full start-end interval or blank.")

    def table(name, headers):
        rows = iter(values(workbook[name]))
        first = next(rows, None)
        if first is None:
            raise ValueError(f"{name}: headers are required.")
        labels = [normalized_label(v) for v in first[1]]
        if any(labels.count(normalized_label(h)) != 1 for h in headers):
            raise ValueError(f"{name}: required headers: {', '.join(headers)}.")
        indexes = [labels.index(normalized_label(h)) for h in headers]
        for number, row in rows:
            yield number, [row[i] if i < len(row) else None for i in indexes]

    groups, departments, students = {}, {}, set()
    for number, row in table(SHEETS[1], STUDENT_HEADERS):
        try:
            cls, dept, subject = (text(row[0], "Class", 50), text(row[1], "Department", 100),
                                  text(row[2], "Subject", 150))
            low, high = integer(row[3], "Roll Number Start"), integer(row[4], "Roll Number End")
            if low > high:
                raise ValueError("Roll Number Start must not exceed Roll Number End.")
            omitted = set(parse_roll_list(row[5], "Excluded Roll Numbers", allow_empty=True))
            included = set(range(low, high + 1))
            if not omitted <= included:
                raise ValueError("Excluded rolls must be inside the group's roll range.")
            class_key = cls.casefold()
            if departments.setdefault(class_key, dept.casefold()) != dept.casefold():
                raise ValueError("A class cannot belong to different departments.")
            eligible = included - omitted
            identities = {(class_key, roll) for roll in eligible}
            if students & identities:
                raise ValueError("Overlapping student groups in the same exam slot.")
            students.update(identities)
            key = (class_key, subject.casefold())
            if key in groups:
                group = groups[key]
                if group[4] & included:
                    raise ValueError("Repeated/overlapping ranges for the same class and subject.")
                group[4].update(included)
                group[5].update(omitted)
            else:
                groups[key] = [number, cls, dept, subject, included, omitted]
        except ValueError as error:
            raise ValueError(f"{SHEETS[1]}, row {number}: {error}") from None
    if not groups:
        raise ValueError("Student Subject Data must contain at least one group.")
    for number, cls, dept, subject, included, omitted in groups.values():
        excluded = text(",".join(map(str, sorted(omitted))), "Excluded Roll Numbers", 255, required=False)
        records["Class Data"].append((number, (cls, dept, min(included), max(included),
            len(included - omitted), excluded, subject, ",".join(map(str, sorted(included))))))
        records["Timetable"].append((number, (cls, subject, None, day, start, end, raw)))
    seen = set()
    for number, row in table(SHEETS[2], ROOM_HEADERS):
        room = text(row[0], "Room Number", 20)
        if room.casefold() in seen:
            raise ValueError(f"Classroom Data, row {number}: duplicate room.")
        seen.add(room.casefold())
        records["Classroom Data"].append((number, (room, integer(row[1], "Seating Capacity", minimum=1))))
    if not seen:
        raise ValueError("Classroom Data must contain at least one room.")
    return records
