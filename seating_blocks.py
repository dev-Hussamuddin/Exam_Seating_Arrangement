"""Room-level supervision blocks, retaining sequential student/seat assignments."""

TARGET_BLOCK_SIZE = 30


def make_blocks(allocations, first_block=1):
    """Round occupied seats / target to nearest (halves up), then absorb remainder.

    Every nonempty room gets at least one block. Earlier blocks take the target
    number and the last takes the remaining students (63 -> 30 + 33). Blocks
    follow seat order, may mix exam groups, and never cross room boundaries.
    """
    by_room = {}
    for allocation in allocations:
        rolls, seats = allocation["rolls"], allocation["seats"]
        if len(rolls) != len(seats):
            raise ValueError("Every allocated roll must have exactly one seat.")
        students = by_room.setdefault(allocation["classroom_id"], [])
        students.extend({"seat": seat, "roll": roll, "class_id": allocation["class_id"],
                         "timetable_id": allocation["timetable_id"]}
                        for seat, roll in zip(seats, rolls))
    blocks = []
    for room_id, students in by_room.items():
        students.sort(key=lambda student: student["seat"])
        if not students:
            continue
        count = max(1, (len(students) + TARGET_BLOCK_SIZE // 2) // TARGET_BLOCK_SIZE)
        for index in range(count):
            start = index * TARGET_BLOCK_SIZE
            end = start + TARGET_BLOCK_SIZE if index < count - 1 else len(students)
            members = students[start:end]
            groups = {}
            for student in members:
                group = groups.setdefault(student["timetable_id"], {
                    "classroom_id": room_id, "class_id": student["class_id"],
                    "timetable_id": student["timetable_id"], "rolls": [], "seats": []})
                group["rolls"].append(student["roll"])
                group["seats"].append(student["seat"])
            blocks.append({"classroom_id": room_id, "block_number": first_block + len(blocks),
                           "students": members, "groups": list(groups.values())})
    return blocks
