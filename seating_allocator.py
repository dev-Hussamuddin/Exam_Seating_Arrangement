"""Allocate one student per bench, grouped by exact exam slot, then export Excel."""

from collections import Counter, defaultdict
from datetime import datetime, time, timedelta
from itertools import groupby

from mysql.connector import Error as MySQLError

from config import ACADEMIC_YEAR, INPUT_DIR
from database import get_connection, migrate_seating_schema, migrate_input_schema, migrate_block_schema
from seating_blocks import make_blocks
from seating_report import write_seating_report


def eligible_rolls(start, end, excluded):
    if start < 0 or start > end:
        raise ValueError("Invalid batch roll range.")
    try:
        omitted = {int(v.strip()) for v in (excluded or "").split(",") if v.strip()}
    except ValueError:
        raise ValueError("Excluded rolls must be comma-separated integers.") from None
    if any(r < start or r > end for r in omitted):
        raise ValueError("Excluded rolls must be inside the batch roll range.")
    return [r for r in range(start, end + 1) if r not in omitted]


def compact_ranges(rolls):
    """Format ascending, unique rolls without concealing duplicate input."""
    if list(rolls) != sorted(set(rolls)):
        raise ValueError("Rolls must be ascending and unique.")
    parts = []
    for _, group in groupby(enumerate(rolls), lambda pair: pair[1] - pair[0]):
        values = [pair[1] for pair in group]
        parts.append(str(values[0]) if len(values) == 1 else f"{values[0]}-{values[-1]}")
    return ",".join(parts)


def storage_chunks(rolls):
    """Split only if necessary to fit the authoritative VARCHAR(500)."""
    chunk = []
    for roll in rolls:
        if len(compact_ranges(chunk + [roll])) > 500:
            yield chunk
            chunk = []
        chunk.append(roll)
    if chunk:
        yield chunk


def validate_exam_groups(exams):
    """Distinct timetable groups may share a class, but never a student."""
    ids = set()
    students = {}
    for exam in exams:
        key = exam["timetable_id"]
        if key in ids:
            raise ValueError("Duplicate timetable entry in the same slot.")
        ids.add(key)
        for roll in exam["rolls"]:
            student = (exam["class_id"], roll)
            if student in students:
                raise ValueError(f"Overlapping student groups: {exam['class_name']} roll {roll} "
                                 f"appears in timetable entries {students[student]} and {key} in the same slot.")
            students[student] = key


def validate_allocation(exams, rooms, allocations):
    validate_exam_groups(exams)
    expected = {e["timetable_id"]: e for e in exams}
    capacities = {r["classroom_id"]: r["capacity"] for r in rooms}
    actual = defaultdict(list)
    seats = defaultdict(list)
    used = Counter()
    for allocation in allocations:
        key, room = allocation["timetable_id"], allocation["classroom_id"]
        if key not in expected or room not in capacities:
            raise ValueError("Allocation references an unknown class or classroom.")
        if allocation["class_id"] != expected[key]["class_id"]:
            raise ValueError("Allocation references the wrong exam.")
        actual[key].extend(allocation["rolls"])
        assigned = allocation.get("seats", [])
        if len(assigned) != len(allocation["rolls"]):
            raise ValueError("Every allocated roll must have exactly one seat.")
        seats[room].extend(assigned)
        used[room] += len(allocation["rolls"])
    for key, exam in expected.items():
        # Exact list equality verifies coverage, exclusions, duplicates and order.
        if actual[key] != exam["rolls"]:
            raise ValueError(f"Allocation validation failed for {exam['class_name']}.")
    if any(used[room] > capacity for room, capacity in capacities.items()):
        raise ValueError("Room capacity exceeded.")
    for room, assigned in seats.items():
        if sorted(assigned) != list(range(1, used[room] + 1)):
            raise ValueError("Duplicate, missing or invalid seat numbers in a room.")
    if sum(used.values()) != sum(len(e["rolls"]) for e in exams):
        raise ValueError("Allocated count does not equal eligible count.")


def allocate_slot(exams, rooms):
    """Finish each exam group sequentially before starting the next.

    Preserve caller exam and room order. A group continues into the next room
    when needed; only its final room may share remaining seats with later groups.
    """
    validate_exam_groups(exams)
    if len({r["classroom_id"] for r in rooms}) != len(rooms):
        raise ValueError("Duplicate classroom.")
    if any(r["capacity"] <= 0 for r in rooms):
        raise ValueError("Classroom capacity must be greater than zero.")
    total = sum(len(e["rolls"]) for e in exams)
    capacity = sum(r["capacity"] for r in rooms)
    if total > capacity:
        raise ValueError(f"Insufficient capacity: {total} students require seats but only "
                         f"{capacity} seats are available. Shortage: {total - capacity} seats.")
    positions = {e["timetable_id"]: 0 for e in exams}
    result = []
    for room in rooms:
        seated = {}
        for seat_number in range(1, room["capacity"] + 1):
            available = [e for e in exams if positions[e["timetable_id"]] < len(e["rolls"])]
            if not available:
                break
            exam = available[0]
            key = exam["timetable_id"]
            seated.setdefault(key, {"class_id": exam["class_id"], "timetable_id": key,
                                    "classroom_id": room["classroom_id"], "rolls": [], "seats": []})
            seated[key]["rolls"].append(exam["rolls"][positions[key]])
            seated[key]["seats"].append(seat_number)
            positions[key] += 1
        result.extend(seated.values())
    validate_allocation(exams, rooms, result)
    return result


def load_slot(cursor, slot):
    cursor.execute("""SELECT t.timetable_id, c.class_id, c.class_name, c.department, t.subject
        FROM timetable t JOIN classes c ON c.class_id=t.class_id
        WHERE t.exam_date=%s AND (t.start_time=%s OR (t.start_time IS NULL AND %s IS NULL))
        AND (t.end_time=%s OR (t.end_time IS NULL AND %s IS NULL))
        ORDER BY c.class_id, t.timetable_id""", (slot[0], slot[1], slot[1], slot[2], slot[2]))
    exams = []
    for timetable_id, class_id, name, department, subject in cursor.fetchall():
        cursor.execute("""SELECT roll_start, roll_end, non_included_rolls, roll_numbers FROM student_batches
            WHERE class_id=%s AND academic_year=%s AND subject=%s""", (class_id, ACADEMIC_YEAR, subject))
        batches = cursor.fetchall()
        if not batches:
            cursor.execute("""SELECT roll_start, roll_end, non_included_rolls, roll_numbers FROM student_batches
                WHERE class_id=%s AND academic_year=%s AND subject IS NULL""", (class_id, ACADEMIC_YEAR))
            batches = cursor.fetchall()
        if not batches:
            print(f"{name} / {subject}: no student data for {ACADEMIC_YEAR}; skipped for seating.")
            continue
        if len(batches) != 1:
            raise ValueError(f"{name} needs exactly one student batch for {ACADEMIC_YEAR}.")
        start, end, excluded, explicit = batches[0]
        if explicit is not None:
            included = {int(r) for r in explicit.split(",") if r.strip()}
            omitted = {int(r) for r in (excluded or "").split(",") if r.strip()}
            rolls = sorted(included - omitted)
        else:
            rolls = eligible_rolls(start, end, excluded)
        exams.append(dict(timetable_id=timetable_id, class_id=class_id, class_name=name,
                          department=department, subject=subject, rolls=rolls))
    return exams


def save_slot(cursor, slot, allocations, first_block=1, timetable_ids=None):
    if timetable_ids is not None:
        for timetable_id in timetable_ids:
            cursor.execute("DELETE FROM seating_arrangements WHERE timetable_id=%s", (timetable_id,))
    else:
        cursor.execute("""DELETE FROM seating_arrangements WHERE timetable_id IN
        (SELECT timetable_id FROM timetable WHERE exam_date=%s
        AND (start_time=%s OR (start_time IS NULL AND %s IS NULL))
        AND (end_time=%s OR (end_time IS NULL AND %s IS NULL)))""",
                   (slot[0], slot[1], slot[1], slot[2], slot[2]))
    blocks = make_blocks(allocations, first_block)
    for block in blocks:
        # Keep each exam's foreign key; mixed groups share the supervision serial.
        for group in block["groups"]:
            offset = 0
            for rolls in storage_chunks(group["rolls"]):
                seats = group["seats"][offset:offset + len(rolls)]
                offset += len(rolls)
                cursor.execute("""INSERT INTO seating_arrangements
                    (timetable_id, classroom_id, roll_start, roll_end, roll_numbers, allocated_count,
                     block_number, seat_numbers) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)""",
                               (group["timetable_id"], block["classroom_id"], rolls[0], rolls[-1],
                                compact_ranges(rolls), len(rolls), block["block_number"],
                                ",".join(map(str, seats))))
    return first_block + len(blocks)


def generate_arrangements(connection, exam_day=None, one_day=False):
    """Replace slots atomically for the whole run. Caller owns the connection."""
    cursor = None
    summaries = []
    try:
        connection.start_transaction()
        cursor = connection.cursor(buffered=True)
        cursor.execute("SELECT classroom_id, classroom_no, capacity FROM classrooms ORDER BY classroom_id")
        rooms = [dict(classroom_id=i, classroom_no=n, capacity=c) for i, n, c in cursor.fetchall()]
        query = "SELECT DISTINCT exam_date, start_time, end_time FROM timetable"
        if exam_day is not None:
            from excel_importer import exam_date
            exam_day = exam_date(exam_day)
            cursor.execute(query + " WHERE exam_date=%s ORDER BY start_time, end_time", (exam_day,))
        else:
            cursor.execute(query + " ORDER BY exam_date, start_time, end_time")
        slots = cursor.fetchall()
        if one_day and len({slot[0] for slot in slots}) > 1:
            raise ValueError("Multiple exam dates exist. Select one with --exam-date YYYY-MM-DD; historical data is preserved.")
        if any(day is None or ((start is None) != (end is None)) for day, start, end in slots):
            raise ValueError("Timetable has undated or incomplete exam slots. Input was preserved; "
                             "set actual exam dates and full time intervals before allocation.")
        if any(start is None for _, start, _ in slots) and len(slots) > 1 and one_day:
            raise ValueError("An untimed exam cannot share a day with other slots. Configure its full time interval.")
        next_block = 1
        for slot in slots:
            exams = load_slot(cursor, slot)
            allocations = allocate_slot(exams, rooms)
            next_block = save_slot(cursor, slot, allocations, next_block)
            summaries.append((slot, exams, rooms, allocations))
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        if cursor is not None:
            cursor.close()
    return summaries


def display_time(value):
    if value is None:
        return "Not configured"
    if isinstance(value, timedelta):
        seconds = int(value.total_seconds())
        return f"{seconds // 3600:02d}:{seconds % 3600 // 60:02d}"
    if isinstance(value, time):
        return value.strftime("%H:%M")
    return str(value)[:5]


def print_summary(slot, exams, rooms, allocations):
    day, start, end = slot
    if isinstance(day, str):
        day = datetime.strptime(day, "%Y-%m-%d").date()
    eligible = sum(len(e["rolls"]) for e in exams)
    capacity = sum(r["capacity"] for r in rooms)
    allocated = sum(len(a["rolls"]) for a in allocations)
    print(f"\nExam Slot: {day:%d-%m-%Y} {display_time(start)}-{display_time(end)}")
    print(f"Eligible Students: {eligible}\nAvailable Capacity: {capacity}")
    print(f"Students Allocated: {allocated}\nUnused Seats: {capacity - allocated}")
    names = {e["timetable_id"]: f"{e['class_name']} / {e.get('subject', '')}" for e in exams}
    for room in rooms:
        rows = [a for a in allocations if a["classroom_id"] == room["classroom_id"]]
        used = sum(len(a["rolls"]) for a in rows)
        print(f"\nRoom {room['classroom_no']} - {used}/{room['capacity']}")
        for row in rows:
            print(f"{names[row['timetable_id']]}: {compact_ranges(row['rolls'])}")


def main(exam_day=None, legacy=False, input_dir=None):
    connection = None
    try:
        data = None
        if not legacy:
            from three_file_input import three_file_input_available, read_three_file_input
            from excel_importer import import_records, exam_date
            directory = INPUT_DIR if input_dir is None else input_dir
            if three_file_input_available(directory):
                data = read_three_file_input(directory)
                if exam_day is not None and exam_date(exam_day) != data["exam_info"]["exam_date"]:
                    raise ValueError("--exam-date must match the date in Timetable.xlsx.")
                import_records(data["records"], semester_aware=True)
        connection = get_connection()
        migrate_seating_schema(connection)
        migrate_input_schema(connection)
        migrate_block_schema(connection)
        # Finish any metadata transaction before starting the allocation transaction.
        connection.commit()
        if data is not None:
            from three_file_workflow import generate_three_file_arrangements
            summaries = generate_three_file_arrangements(connection, data)
        else:
            summaries = generate_arrangements(connection, exam_day=exam_day, one_day=True)
        if not summaries:
            print("No timetable records found. Import the Excel workbook first.")
        for summary in summaries:
            print_summary(*summary)
        if summaries:
            try:
                report = (write_seating_report(summaries, exam_info=data["exam_info"])
                          if data is not None else write_seating_report(summaries))
            except (OSError, ValueError, RuntimeError) as error:
                print(f"Allocation succeeded and was committed, but Excel report generation failed: {error}")
                return 1
            print(f"Excel report saved: {report}")
        return 0
    except MySQLError as error:
        print(f"Allocation failed (MySQL error {error.errno}). Check credentials and database setup.")
    except (ValueError, RuntimeError, OSError) as error:
        print(str(error))
    finally:
        if connection is not None:
            connection.close()
    return 1


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate seating for one exam day without changing other dates.")
    parser.add_argument("--exam-date", help="YYYY-MM-DD; optional when the database contains only one exam date.")
    parser.add_argument("--legacy", action="store_true", help="Allocate from existing MySQL timetable instead of the three finalized workbooks.")
    parser.add_argument("--input-dir", help="Directory containing the three finalized workbooks; defaults to input/.")
    args = parser.parse_args()
    raise SystemExit(main(args.exam_date, legacy=args.legacy, input_dir=args.input_dir))
