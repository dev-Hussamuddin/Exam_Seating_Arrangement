"""Allocate only the current finalized timetable, preserving unrelated DB records."""

from seating_allocator import allocate_slot, save_slot


def generate_three_file_arrangements(connection, data):
    """Use validated workbook snapshots and existing DB IDs in one transaction.

    The caller imports the same snapshot first. Historical timetable entries and
    rooms absent from these files cannot enter this run. No live data is deleted
    except the previous allocations of the specifically selected timetable IDs.
    """
    cursor = None
    try:
        connection.start_transaction()
        cursor = connection.cursor(buffered=True)
        info = data["exam_info"]
        slot = (info["exam_date"], info["start_time"], info["end_time"])
        rooms, exams = [], []
        for _, (number, capacity) in data["records"]["Classroom Data"]:
            cursor.execute("SELECT classroom_id FROM classrooms WHERE classroom_no=%s", (number,))
            found = cursor.fetchall()
            if len(found) != 1:
                raise ValueError(f"Room {number}: import the current workbooks before allocation.")
            rooms.append(dict(classroom_id=found[0][0], classroom_no=number, capacity=capacity))
        rooms.sort(key=lambda room: room["classroom_id"])
        for group in data["exams"]:
            cursor.execute("SELECT class_id FROM classes WHERE class_name=%s", (group["class_name"],))
            found = cursor.fetchall()
            if len(found) != 1:
                raise ValueError(f"{group['class_name']}: import the current workbooks before allocation.")
            class_id = found[0][0]
            cursor.execute("""SELECT timetable_id FROM timetable WHERE class_id=%s AND subject=%s
                AND semester=%s AND exam_date=%s
                AND (start_time=%s OR (start_time IS NULL AND %s IS NULL))
                AND (end_time=%s OR (end_time IS NULL AND %s IS NULL))""",
                (class_id, group["subject"], group["semester"], slot[0], slot[1], slot[1], slot[2], slot[2]))
            found = cursor.fetchall()
            if len(found) != 1:
                raise ValueError(f"{group['class_name']} / {group['subject']}: expected exactly one matching imported timetable entry.")
            exams.append(dict(class_id=class_id, timetable_id=found[0][0], class_name=group["class_name"],
                              department=group["department"], subject=group["subject"], semester=group["semester"],
                              rolls=group["rolls"][:], roll_width=group["roll_width"]))
        exams.sort(key=lambda exam: (exam["class_id"], exam["timetable_id"]))
        allocations = allocate_slot(exams, rooms)
        save_slot(cursor, slot, allocations, timetable_ids=[exam["timetable_id"] for exam in exams])
        connection.commit()
        return [(slot, exams, rooms, allocations)]
    except Exception:
        connection.rollback()
        raise
    finally:
        if cursor is not None:
            cursor.close()
