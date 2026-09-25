"""MySQL connection and schema setup. Run this file to initialize the database."""

import mysql.connector

from config import BASE_DIR, get_mysql_config


def get_connection():
    """Connect to an existing exam_seating_db database.

    The caller is responsible for closing the connection after use.
    """
    return mysql.connector.connect(**get_mysql_config())


def migrate_seating_schema(connection):
    """Run before allocation transactions: MySQL DDL commits implicitly.

    Only reconstruct legacy rows whose count proves the range is contiguous.
    Ambiguous legacy rows are preserved and require manual clarification.
    """
    cursor = connection.cursor(buffered=True)
    try:
        cursor.execute("SHOW COLUMNS FROM seating_arrangements LIKE 'roll_numbers'")
        column = cursor.fetchone()
        if column and column[2] == "NO":
            return
        condition = " WHERE roll_numbers IS NULL" if column else ""
        cursor.execute("SELECT arrangement_id, roll_start, roll_end, allocated_count "
                       "FROM seating_arrangements" + condition)
        legacy = cursor.fetchall()
        for row_id, start, end, count in legacy:
            if start > end or count != end - start + 1:
                raise ValueError(f"Legacy arrangement {row_id} has ambiguous rolls. "
                                 "Resolve its actual roll numbers before migration; data was preserved.")
        if not column:
            cursor.execute("ALTER TABLE seating_arrangements ADD COLUMN roll_numbers VARCHAR(500) NULL")
        for row_id, start, end, _ in legacy:
            rolls = str(start) if start == end else f"{start}-{end}"
            cursor.execute("UPDATE seating_arrangements SET roll_numbers=%s WHERE arrangement_id=%s",
                           (rolls, row_id))
        connection.commit()
        cursor.execute("ALTER TABLE seating_arrangements MODIFY COLUMN roll_numbers VARCHAR(500) NOT NULL")
    finally:
        cursor.close()


def migrate_input_schema(connection):
    """Preserve subject groups and undated timetable input; run before DML."""
    cursor = connection.cursor(buffered=True)
    try:
        for table, column, definition in (
            ("student_batches", "subject", "VARCHAR(150) NULL"),
            ("student_batches", "roll_numbers", "TEXT NULL"),
            ("timetable", "time_text", "VARCHAR(255) NULL"),
        ):
            cursor.execute(f"SHOW COLUMNS FROM {table} LIKE '{column}'")
            if cursor.fetchone() is None:
                cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
        for column, kind in (("exam_date", "DATE"), ("start_time", "TIME"), ("end_time", "TIME")):
            cursor.execute(f"SHOW COLUMNS FROM timetable LIKE '{column}'")
            existing = cursor.fetchone()
            if existing and existing[2] == "NO":
                cursor.execute(f"ALTER TABLE timetable MODIFY COLUMN {column} {kind} NULL")
        connection.commit()
    finally:
        cursor.close()


def migrate_block_schema(connection):
    """Add nullable metadata without rewriting historical allocations. Run before DML."""
    cursor = connection.cursor(buffered=True)
    try:
        for column, definition in (("block_number", "INT NULL"), ("seat_numbers", "TEXT NULL")):
            cursor.execute(f"SHOW COLUMNS FROM seating_arrangements LIKE '{column}'")
            if cursor.fetchone() is None:
                cursor.execute(f"ALTER TABLE seating_arrangements ADD COLUMN {column} {definition}")
        connection.commit()
    finally:
        cursor.close()


def migrate_student_semester_schema(connection):
    """Preserve legacy batches while distinguishing new class/semester/subject data."""
    cursor = connection.cursor(buffered=True)
    try:
        cursor.execute("SHOW COLUMNS FROM student_batches LIKE 'semester'")
        if cursor.fetchone() is None:
            cursor.execute("ALTER TABLE student_batches ADD COLUMN semester VARCHAR(50) NULL")
        connection.commit()
    finally:
        cursor.close()


def initialize_database():
    """Create the database and tables from schema.sql without deleting data."""
    schema = (BASE_DIR / "schema.sql").read_text(encoding="utf-8")
    server_config = get_mysql_config()
    server_config.pop("database")

    # Connect without selecting a database so first-time setup can create it.
    connection = mysql.connector.connect(**server_config)
    try:
        cursor = connection.cursor()
        try:
            # This schema contains only simple, semicolon-separated statements.
            for statement in schema.split(";"):
                if statement.strip():
                    cursor.execute(statement)
        finally:
            cursor.close()
        migrate_seating_schema(connection)
        migrate_input_schema(connection)
        migrate_block_schema(connection)
        migrate_student_semester_schema(connection)
    finally:
        connection.close()


if __name__ == "__main__":
    initialize_database()
    print("Database exam_seating_db and all five tables are ready.")
