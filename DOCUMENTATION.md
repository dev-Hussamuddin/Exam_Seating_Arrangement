# College Exam Seating Arrangement System

> **Compact Project Documentation** — designed for quick faculty/project evaluation.

## 1. Project at a Glance

| Item | Details |
|---|---|
| **Purpose** | Automatically generate college exam seating arrangements from Excel input. |
| **Language** | Python 3 |
| **Database** | MySQL (`exam_seating_db`) |
| **Input** | `input/Tables.xlsx` |
| **Output** | `output/Seating_Arrangement.xlsx` |
| **Main execution** | `main.py` → `seating_allocator.py` |
| **Excel library** | `openpyxl` |
| **MySQL connector** | `mysql-connector-python` |
| **Tests** | Python `unittest` |

### What the system does

The project takes class/student, classroom-capacity, and timetable information from Excel, validates it, stores it in MySQL, calculates eligible students for each exam slot, allocates students to rooms, validates the allocation, and generates a formatted Excel report.

### Main features

- Excel-based input processing
- MySQL-backed data storage
- Subject-specific student/roll handling
- Excluded/debarred/left roll-number handling
- Exam-slot-wise allocation
- Classroom capacity validation
- Department-aware room mixing
- Post-allocation validation
- Transaction + rollback protection
- Automated Excel report generation
- Automated test suite

---

## 2. Project Structure

```text
grp project exam/
│
├── main.py                  # Import stage entry point
├── excel_importer.py        # Excel reading, validation and import
├── seating_allocator.py     # Core seating-allocation algorithm
├── seating_report.py        # Final Excel report generation
├── database.py              # MySQL/schema setup and connection
├── config.py                # Shared configuration
├── schema.sql               # Database/table definition
├── requirements.txt         # Python dependencies
├── README.md                # Original project README
├── DOCUMENTATION.md         # This compact documentation
│
├── input/
│   └── Tables.xlsx          # Project input workbook
│
├── output/
│   └── Seating_Arrangement.xlsx
│
├── test_college_input.py
├── test_excel_importer.py
├── test_seating_allocator.py
└── test_seating_report.py
```

`.venv/` and `__pycache__/` are generated/environment folders and are not part of the application logic.

---

## 3. Requirements

Before running the project, make sure:

- Python 3 is installed.
- MySQL Server is installed and running.
- A MySQL user is available (the default configuration uses `root`).
- `MYSQL_PASSWORD` is set in the PowerShell session.
- `input/Tables.xlsx` is present.

### Current database configuration

Defined in `config.py`:

```text
Host:     localhost
Port:     3306
User:     root
Database: exam_seating_db
```

Optional environment variables supported by the project:

```text
MYSQL_HOST
MYSQL_PORT
MYSQL_USER
```

The password is supplied through `MYSQL_PASSWORD` instead of being hard-coded in source code.

---

# 4. Setup and Run — Windows PowerShell

## A. First-time setup

Run these commands from the project folder.

### 1. Check Python

```powershell
python --version
```

### 2. Create virtual environment

```powershell
python -m venv .venv
```

### 3. Activate virtual environment

```powershell
.\.venv\Scripts\Activate.ps1
```

### 4. Install dependencies

```powershell
pip install -r requirements.txt
```

### 5. Set MySQL password

```powershell
$env:MYSQL_PASSWORD="THEIR_MYSQL_PASSWORD"
```

### 6. Initialize the database (first run only)

```powershell
.\.venv\Scripts\python.exe database.py
```

This creates/prepares `exam_seating_db` and the required tables.

---

## B. Normal project run

After first-time database setup:

### 1. Import the Excel data

```powershell
.\.venv\Scripts\python.exe main.py
```

### 2. Generate the seating arrangement

```powershell
.\.venv\Scripts\python.exe seating_allocator.py
```

### Complete command sequence

```powershell
python --version
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:MYSQL_PASSWORD="THEIR_MYSQL_PASSWORD"
.\.venv\Scripts\python.exe main.py
.\.venv\Scripts\python.exe seating_allocator.py
```

> **Important:** On a fresh database, run `database.py` once before `main.py`.

> **PowerShell note:** `MYSQL_PASSWORD` applies to the current PowerShell session. Set it again in a new terminal.

---

# 5. How the Project Works

```text
              input/Tables.xlsx
                     │
                     ▼
              ┌─────────────┐
              │   main.py   │
              └──────┬──────┘
                     ▼
          ┌─────────────────────┐
          │ excel_importer.py   │
          │ validate + import   │
          └──────────┬──────────┘
                     ▼
               MySQL Database
                     │
                     │ timetable + students + rooms
                     ▼
         ┌────────────────────────┐
         │ seating_allocator.py  │
         │ calculate + validate  │
         └───────────┬────────────┘
                     ▼
           seating_arrangements
                     │
                     ▼
          ┌─────────────────────┐
          │ seating_report.py   │
          └──────────┬──────────┘
                     ▼
        output/Seating_Arrangement.xlsx
```

### Execution stages

**Stage 1 — Import**  
`main.py` calls `excel_importer.py` to locate `Tables.xlsx`, validate the workbook, and store the data in MySQL.

**Stage 2 — Allocation**  
`seating_allocator.py` reads the timetable, eligible students, and classrooms; allocates seats separately for each exact exam date/time slot; validates the result; and stores the allocation.

**Stage 3 — Report**  
`seating_report.py` converts the successful allocation into the final formatted Excel report.

---

# 6. Main Files — Faculty Quick Reference

| File | Responsibility | Runs when |
|---|---|---|
| `main.py` | Starts Excel import | User runs `main.py` |
| `excel_importer.py` | Reads/validates workbook and imports data | Called by `main.py` |
| `seating_allocator.py` | Core seating algorithm + allocation validation | User runs `seating_allocator.py` |
| `seating_report.py` | Creates final Excel report | Called after successful allocation |
| `database.py` | Creates database/tables and manages DB connection/migrations | First-time setup / DB operations |
| `config.py` | Paths, DB settings and shared constants | Imported by other modules |
| `schema.sql` | SQL definition for database/tables | Used by `database.py` |
| `requirements.txt` | Required Python packages | Used during installation |

### Key point

The **main business logic is split into two entry points**:

```text
main.py                = Excel → MySQL
seating_allocator.py   = MySQL → Seating Allocation → Excel
```

---

# 7. Database Design

Database name:

```text
exam_seating_db
```

The project uses five main tables:

| Table | Stores |
|---|---|
| `classes` | Class/department master data |
| `student_batches` | Academic year, roll ranges/lists, exclusions, subjects |
| `classrooms` | Room number and capacity |
| `timetable` | Exam class, subject, date and time |
| `seating_arrangements` | Generated room/roll allocation |

### Important relationships

```text
classes
  │
  ├── student_batches
  │
  └── timetable
         │
         ▼
seating_arrangements ◄── classrooms
```

### Student eligibility

The allocator effectively uses:

```text
Eligible Students = Valid Roll Numbers − Excluded Roll Numbers
```

For a subject, the allocator first looks for a matching **class + academic year + subject** batch. If none exists, it can use the class-level batch.

---

# 8. Excel Input Format

Input file:

```text
input/Tables.xlsx
```

The supplied workbook contains three logical data sections.

### Class / Student section

```text
Class
Subject
Roll nos.
excluding nos.
No. of Students
```

### Classroom section

```text
Room No.
Capacity (No. of benches)
```

### Timetable section

```text
Class
Subject
Time
Exam Date
```

### Roll-number formats supported

```text
1-10
1,3,5,7
1-5,8,10-12
```

The importer also handles common Unicode dash characters and some trailing separators found in Excel data.

### Time formats

The importer accepts Excel time values and common text ranges such as:

```text
09:00-10:00
12:00-01:00
9:00 AM-10:00 AM
```

### Date values

Excel date cells and `YYYY-MM-DD` text are supported.

---

# 9. Seating Allocation Logic

The seating algorithm is implemented in `seating_allocator.py`.

## Algorithm summary

1. Find every distinct **Exam Date + Start Time + End Time** slot.
2. For each slot, load all timetable entries.
3. Determine eligible rolls for every class/subject.
4. Check for duplicate/conflicting exam groups.
5. Calculate total available capacity.
6. Stop if students exceed available capacity.
7. Process classrooms in deterministic order.
8. Place students one by one while tracking department counts in the room.
9. Prefer an available group from a department with fewer students currently seated in that room.
10. Preserve ascending roll-number order within an exam group.
11. Validate the completed allocation.
12. Save the allocation to MySQL.
13. Generate `Seating_Arrangement.xlsx`.

### Simple example

For two departments:

```text
CS       → 1,2,3,4
Science  → 1,2,3,4
```

the room-selection rule can alternate groups based on the department-count balance, while roll numbers remain in ascending order.

### Why this approach is used

It provides deterministic allocation, respects room capacity, preserves roll ordering, and attempts to mix departments within rooms instead of unnecessarily grouping one department together.

---

# 10. Validation and Safety Checks

### Input validation

The importer checks for issues such as:

- Invalid roll-number lists or ranges
- Excluded rolls outside valid roll ranges/lists
- Invalid student counts
- Invalid exam dates/times
- Duplicate/conflicting input structures
- Unsupported formula cells

### Allocation validation

Before committing an allocation, the system checks:

- No missing students
- No extra students
- No duplicate roll allocations
- No room over capacity
- Expected student count = allocated student count
- Timetable and classroom references are valid

### Capacity rule

The allocation stops when:

```text
Total Eligible Students > Total Classroom Capacity
```

and reports the shortage rather than silently creating an invalid arrangement.

---

# 11. Transactions and Data Integrity

The project uses database transactions to avoid partial updates.

### Import

```text
Validate workbook
      ↓
Start transaction
      ↓
Insert/update data
      ↓
Success → COMMIT
Error   → ROLLBACK
```

### Allocation

```text
Read slot
   ↓
Allocate
   ↓
Validate
   ↓
Save
   ↓
All slots successful → COMMIT
Any allocation failure → ROLLBACK
```

Report generation happens after a successful allocation commit. Therefore, a report-writing failure does not undo the already committed database allocation.

---

# 12. Output

Generated file:

```text
output/Seating_Arrangement.xlsx
```

The report contains, by exam slot and room:

- Exam date/time
- Room number and capacity
- Class
- Subject
- Allocated roll numbers
- Number allocated
- Unused seats

It also includes an overall summary of:

```text
Total Eligible Students
Total Allocated Students
Total Capacity Used
Total Unused Seats
```

The report is formatted for practical use/printing and uses compact roll-number ranges such as:

```text
1-4,7,10-12
```

instead of listing every number separately.

---

# 13. Re-running the Project

### Re-run `main.py`

Used to import/refresh data from the current `Tables.xlsx`.

### Re-run `seating_allocator.py`

Used to regenerate the seating arrangement from the current database data.

> Close `output/Seating_Arrangement.xlsx` in Excel before regenerating it, because a file locked by Excel may not be replaceable on Windows.

---

# 14. Testing

The project includes four test modules:

```text
test_college_input.py
test_excel_importer.py
test_seating_allocator.py
test_seating_report.py
```

Run the complete suite with:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -v
```

### Verification result from project review

The included test suite was executed during documentation review and **68 tests passed** in the inspection environment.

---

# 15. Common Problems

| Problem | Check / Fix |
|---|---|
| `MYSQL_PASSWORD` error | Set `$env:MYSQL_PASSWORD="..."` in the current PowerShell session. |
| MySQL connection failure | Make sure MySQL Server is running and host/user/port are correct. |
| Database/tables missing | Run `database.py` once before `main.py`. |
| No `.xlsx` input found | Ensure `input/Tables.xlsx` exists. |
| Multiple input workbooks | Keep only the intended `.xlsx` input workbook in `input/`. |
| Insufficient capacity | Increase room capacity/count or correct the input data. |
| No student data for an exam | Check class, subject, academic year and roll-number input. |
| Report cannot be replaced | Close `Seating_Arrangement.xlsx` in Excel and rerun allocation. |

---

# 16. Security / Good Practice

- Do **not** store the real MySQL password in source code.
- Use `MYSQL_PASSWORD` as an environment variable.
- Do not commit `.venv/` or `__pycache__/` to Git unless specifically required.
- Keep the input workbook structure consistent.
- Back up the database before major schema/data changes.

---

# 17. Project Limitations / Scope

- The current project is configured around MySQL and the supplied Excel input structure.
- Seating is based on the available classroom capacity/bench count.
- The current configuration uses:

```python
STUDENTS_PER_BENCH = 1
```

- Allocation is designed for the data and rules implemented in this version; changes to workbook format or business rules may require code changes.

---

# 18. Future Improvements

Possible extensions include:

- Web-based or GUI interface
- User login/role management
- Live room availability management
- PDF report generation
- Manual seat-lock/override options
- More advanced seat-pattern constraints
- Better configuration through an external settings file

---

# 19. One-Page Faculty Summary

### Problem
Manual examination seating is time-consuming and can cause capacity, duplication, or roll-number errors.

### Solution
A Python + MySQL system automates the process from Excel input to final seating report.

### Input
`input/Tables.xlsx`

### Processing
```text
Excel Validation
      ↓
MySQL Storage
      ↓
Exam-slot Analysis
      ↓
Eligible Student Calculation
      ↓
Room Allocation
      ↓
Allocation Validation
      ↓
Database Storage
```

### Output
`output/Seating_Arrangement.xlsx`

### Core modules
```text
main.py
excel_importer.py
seating_allocator.py
seating_report.py
database.py
config.py
```

### Quality controls
```text
Input validation
Capacity checking
Duplicate/conflict checking
Post-allocation validation
Transactions + rollback
Automated tests
```

### Quick run

```powershell
$env:MYSQL_PASSWORD="THEIR_MYSQL_PASSWORD"
.\.venv\Scripts\python.exe main.py
.\.venv\Scripts\python.exe seating_allocator.py
```

**First time only:**

```powershell
.\.venv\Scripts\python.exe database.py
```

---

## Final Note

This document intentionally keeps the implementation details that demonstrate **how the project works, what files are important, how to run it, how data moves through the system, how seating is allocated, how errors are handled, and how the output is produced**, while removing repeated explanations and low-level detail that are not necessary for a quick project evaluation.
