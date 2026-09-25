"""Folder paths and MySQL connection settings."""

import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
ACADEMIC_YEAR = "2026-2027"
COLLEGE_NAME = "Royal College of Arts, Science & Commerce"
COLLEGE_SUBTITLE = "(Autonomous)"

MYSQL_CONFIG = {
    "host": os.getenv("MYSQL_HOST", "localhost"),
    "port": int(os.getenv("MYSQL_PORT", "3306")),
    "user": os.getenv("MYSQL_USER", "root"),
    "database": "exam_seating_db",
}


def get_mysql_config():
    """Read the password from the environment without logging or storing it."""
    password = os.getenv("MYSQL_PASSWORD")
    if not password:
        raise RuntimeError("Set MYSQL_PASSWORD in your current session before connecting to MySQL.")
    return {**MYSQL_CONFIG, "password": password}


STUDENTS_PER_BENCH = 1
