"""Entry point for the offline College Exam Seating Arrangement MVP."""

from zipfile import BadZipFile

from mysql.connector import Error as MySQLError
from openpyxl.utils.exceptions import InvalidFileException

from excel_importer import import_from_input


def main():
    """Import the input workbook into the existing MySQL database."""
    try:
        import_from_input()
        return 0
    except MySQLError as error:
        print(f"Import failed (MySQL error {error.errno}). Check the connection, credentials and database schema.")
    except (ValueError, RuntimeError, OSError, BadZipFile, InvalidFileException) as error:
        print(f"Import failed: {error}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
