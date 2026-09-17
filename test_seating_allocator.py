"""Offline allocation and transactional persistence tests; no MySQL credentials needed."""

import unittest
from unittest.mock import Mock

from config import ACADEMIC_YEAR
from database import migrate_seating_schema
from seating_allocator import (allocate_slot, compact_ranges, eligible_rolls,
                               generate_arrangements, storage_chunks, validate_allocation)
from test_excel_importer import TestConnection


def exam(i, department, end=6, excluded="", start=1):
    return dict(class_id=i, timetable_id=i, class_name=f"Class{i}",
                department=department, rolls=eligible_rolls(start, end, excluded))


def rooms(*capacities):
    return [dict(classroom_id=i, classroom_no=str(i), capacity=c)
            for i, c in enumerate(capacities, 1)]


class AllocationTests(unittest.TestCase):
    def test_exclusions_and_count(self):
        exams = [exam(1, "CS", 60, "5,12,27")]
        result = allocate_slot(exams, rooms(60))
        self.assertEqual(len(result[0]["rolls"]), 57)
        self.assertEqual(compact_ranges(result[0]["rolls"]), "1-4,6-11,13-26,28-60")

    def test_continuation(self):
        result = allocate_slot([exam(1, "CS", 8, "3")], rooms(4, 4))
        self.assertEqual([a["rolls"] for a in result], [[1, 2, 4, 5], [6, 7, 8]])

    def test_different_departments_preferred(self):
        exams = [exam(1, "CS"), exam(2, "CS"), exam(3, "BAF")]
        result = allocate_slot(exams, rooms(6, 12))
        first_room = [a for a in result if a["classroom_id"] == 1]
        self.assertEqual([a["class_id"] for a in first_room], [1, 3])
        self.assertEqual([len(a["rolls"]) for a in first_room], [3, 3])

    def test_same_department_fallback(self):
        result = allocate_slot([exam(1, "CS", 2), exam(2, "CS", 2)], rooms(4))
        self.assertEqual([a["class_id"] for a in result], [1, 2])

    def test_insufficient_capacity(self):
        with self.assertRaisesRegex(ValueError, "Insufficient capacity: 6 students require seats but only 5 seats are available. Shortage: 1 seats."):
            allocate_slot([exam(1, "CS")], rooms(5))

    def test_duplicate_and_missing_rolls_rejected(self):
        exams, classroom = [exam(1, "CS", 4, "2")], rooms(5)
        for bad_rolls in ([1, 1, 3, 4], [1, 3], [1, 2, 3], [4, 3, 1]):
            result = allocate_slot(exams, classroom)
            result[0]["rolls"] = bad_rolls
            with self.assertRaises(ValueError):
                validate_allocation(exams, classroom, result)

    def test_room_overflow_rejected(self):
        exams = [exam(1, "CS", 4)]
        result = allocate_slot(exams, rooms(4))
        with self.assertRaisesRegex(ValueError, "capacity"):
            validate_allocation(exams, rooms(3), result)

    def test_duplicate_class_exam_rejected(self):
        with self.assertRaisesRegex(ValueError, "Duplicate timetable entry"):
            allocate_slot([exam(1, "CS"), exam(1, "CS")], rooms(20))

    def test_compact_format_and_storage_limit(self):
        self.assertEqual(compact_ranges([1, 2, 3, 4, 6, 7, 8, 10]), "1-4,6-8,10")
        self.assertEqual(compact_ranges([]), "")
        rolls = list(range(1, 1001, 2))
        chunks = list(storage_chunks(rolls))
        self.assertGreater(len(chunks), 1)
        self.assertEqual([r for chunk in chunks for r in chunk], rolls)
        self.assertTrue(all(len(compact_ranges(chunk)) <= 500 for chunk in chunks))

    def test_all_excluded(self):
        self.assertEqual(allocate_slot([exam(1, "CS", 2, "1,2")], []), [])


class PersistenceTests(unittest.TestCase):
    def setUp(self):
        self.connection = TestConnection()
        self.addCleanup(self.connection.db.close)
        db = self.connection.db
        db.executescript("""
            CREATE TABLE seating_arrangements (arrangement_id INTEGER PRIMARY KEY,
                timetable_id INTEGER, classroom_id INTEGER, roll_start INTEGER,
                roll_end INTEGER, roll_numbers VARCHAR(500) NOT NULL, allocated_count INTEGER);
            INSERT INTO classes VALUES (1, 'FYCS', 'CS');
            INSERT INTO classrooms VALUES (1, '308', 4);
            INSERT INTO timetable VALUES (1, 1, 'Maths', '1', '2026-10-09', '09:00:00', '10:00:00', NULL);
            INSERT INTO timetable VALUES (2, 1, 'Science', '1', '2026-10-10', '09:00:00', '10:00:00', NULL);
        """)
        db.execute("INSERT INTO student_batches VALUES (1, 1, ?, 1, 5, 5, '3', NULL, NULL)", (ACADEMIC_YEAR,))
        db.commit()

    def test_rerun_and_separate_slots(self):
        generate_arrangements(self.connection)
        db = self.connection.db
        query = "SELECT timetable_id, classroom_id, roll_numbers, allocated_count FROM seating_arrangements ORDER BY timetable_id"
        first = db.execute(query).fetchall()
        generate_arrangements(self.connection)
        self.assertEqual(db.execute(query).fetchall(), first)
        self.assertEqual(first, [(1, 1, "1-2,4-5", 4), (2, 1, "1-2,4-5", 4)])

    def test_save_replaces_only_selected_slot(self):
        from seating_allocator import save_slot
        generate_arrangements(self.connection)
        db = self.connection.db
        old = db.execute("SELECT * FROM seating_arrangements WHERE timetable_id=2").fetchall()
        self.connection.start_transaction()
        cursor = self.connection.cursor()
        save_slot(cursor, ("2026-10-09", "09:00:00", "10:00:00"), [])
        self.connection.commit()
        cursor.close()
        self.assertEqual(db.execute("SELECT * FROM seating_arrangements WHERE timetable_id=2").fetchall(), old)
        self.assertEqual(db.execute("SELECT COUNT(*) FROM seating_arrangements WHERE timetable_id=1").fetchone()[0], 0)

    def test_failure_in_later_slot_restores_previous_rows(self):
        generate_arrangements(self.connection)
        db = self.connection.db
        db.execute("INSERT INTO timetable VALUES (3, 1, 'Duplicate', '1', '2026-10-10', '09:00:00', '10:00:00', NULL)")
        db.commit()
        before = db.execute("SELECT * FROM seating_arrangements").fetchall()
        with self.assertRaisesRegex(ValueError, "Overlapping student groups"):
            generate_arrangements(self.connection)
        self.assertEqual(db.execute("SELECT * FROM seating_arrangements").fetchall(), before)

    def test_insert_failure_rolls_back_deleted_rows(self):
        generate_arrangements(self.connection)
        db = self.connection.db
        before = db.execute("SELECT * FROM seating_arrangements").fetchall()
        db.execute("CREATE TRIGGER fail_insert BEFORE INSERT ON seating_arrangements BEGIN SELECT RAISE(ABORT, 'test failure'); END")
        db.commit()
        with self.assertRaises(Exception):
            generate_arrangements(self.connection)
        self.assertEqual(db.execute("SELECT * FROM seating_arrangements").fetchall(), before)

    def test_missing_year_batch(self):
        self.connection.db.execute("UPDATE student_batches SET academic_year='old'")
        self.connection.db.commit()
        generate_arrangements(self.connection)
        self.assertEqual(self.connection.db.execute("SELECT COUNT(*) FROM seating_arrangements").fetchone()[0], 0)

    def test_missing_class_skipped_while_valid_class_allocates(self):
        db = self.connection.db
        db.execute("INSERT INTO classes VALUES (2, 'TYCS', '')")
        db.execute("INSERT INTO timetable VALUES (3, 2, 'Sports', NULL, '2026-10-09', '09:00:00', '10:00:00', NULL)")
        db.commit()
        generate_arrangements(self.connection)
        self.assertEqual(db.execute("SELECT timetable_id, allocated_count FROM seating_arrangements ORDER BY timetable_id").fetchall(), [(1, 4), (2, 4)])
        self.assertEqual(db.execute("SELECT COUNT(*) FROM timetable WHERE timetable_id=3").fetchone()[0], 1)


class MigrationTests(unittest.TestCase):
    def test_legacy_backfill(self):
        connection = Mock()
        cursor = connection.cursor.return_value
        cursor.fetchone.return_value = None
        cursor.fetchall.return_value = [(1, 1, 4, 4)]
        migrate_seating_schema(connection)
        cursor.execute.assert_any_call("UPDATE seating_arrangements SET roll_numbers=%s WHERE arrangement_id=%s", ("1-4", 1))

    def test_ambiguous_legacy_preserved(self):
        connection = Mock()
        cursor = connection.cursor.return_value
        cursor.fetchone.return_value = None
        cursor.fetchall.return_value = [(1, 1, 4, 3)]
        with self.assertRaisesRegex(ValueError, "ambiguous"):
            migrate_seating_schema(connection)
        self.assertTrue(all(call.args[0].startswith(("SHOW", "SELECT")) for call in cursor.execute.call_args_list))


if __name__ == "__main__":
    unittest.main()
