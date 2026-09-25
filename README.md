# Finalized three-file exam seating workflow

The default workflow uses exactly these three active input workbooks alongside
the preserved legacy `input/Tables.xlsx`:

| Workbook | Purpose | Columns |
|---|---|---|
| `input/Students & Subject Data.xlsx` | Static student/subject groups | Department, Class, Semester, Subject, Roll Nos. (Range / List), Excluded Roll Nos. |
| `input/Classroom Data.xlsx` | Static room capacities | Room No., Capacity (No. of Benches/Seats) |
| `input/Timetable.xlsx` | Dynamic exam-day selection | Class, Semester, Subject |

Enter **Exam Date, Exam Name, Month / Year, and Timing** once at the top of
`Timetable.xlsx`. The program matches Class + Semester + Subject to the static
student data. Use the same class names in both files; different departments with
similarly named classes need distinct class names, as illustrated in the examples.
Numeric roll ranges/lists and mixed lists are supported (`001-010, 015, 020-025`).
Excluded rolls are removed before counting or allocating. Blank or dash-only
exclusion cells mean no exclusions. Leading-zero presentation is retained in the
college report. Overlapping students across different static subjects are allowed;
the same class/roll appearing twice in the selected daily exams is rejected.

The supplied student and timetable rows are **examples for review**, not actual
college records. Room capacities are retained from the original workbook:
201=48, 202=40, 203=45, 204=55, 205=70. The example timetable selects four subjects
(Business Law and Mathematics remain unscheduled), yielding 225 eligible students,
258 available seats, 33 unused seats and 8 supervision blocks.

## Run the finalized workflow

With Python dependencies installed, an initialized MySQL database, and
`MYSQL_PASSWORD` set in the session, run from the project directory:

```powershell
python -B seating_allocator.py
```

This validates and imports the three files, allocates only their selected exams
and listed classrooms, then automatically exports `output/Seating_Arrangement.xlsx`.
`python -B main.py` remains available as an import-only command. An alternative
three-file directory can be supplied with `--input-dir "path/to/folder"` to the
allocator. A partial set of finalized files produces an error instead of silently
falling back to old data. `--exam-date`, if supplied, must match `Timetable.xlsx`.

Only selected timetable IDs have their allocations replaced. Historical and
unselected database records are retained and are not included in the new report.
The importer and allocation stages each retain their transaction/rollback safety.
The new nullable student-batch `semester` column is added before data transactions;
legacy batches are not assigned guessed semesters. Exam header metadata is read
from the current timetable workbook. Live MySQL is not used by the offline tests.

## Finalized output and supervision blocks

The main **Seating Arrangement** sheet matches the college layout, with the college
heading, exam name, month/year, actual date/weekday and timing above exactly seven
columns: **Room No., Block No., Class, Semester, Subject, Roll Nos., Grand Total**.
Room and block labels are merged across their detail rows. Mixed subjects remain
separate rows within the same supervision block. Grand Total contains each row's
student count, with the overall total at the bottom. Supporting Room Blocks,
Allocation Summary and per-room seat maps retain capacity, unused-seat and block
totals. All sheets have print settings and repeated headings.

Allocation finishes each Class + Semester + Subject group in ascending roll order before moving to the next group. Existing class-ID/timetable-ID priority and room-ID order are preserved. Blocks follow actual occupied seat
order within each room. 30 is a target, not a maximum: round occupied seats / 30
to the nearest integer (halves up), at least one block per occupied room, then
absorb the remainder into the final block. Thus **63 in one room = 30 + 33**.
Block serials continue across all rooms for the report. No teacher fields exist.

## Verification and templates

```powershell
python -B -m unittest discover -v
```

The complete offline suite contains **115 tests**, including three-file parsing,
timetable/semester filtering, exclusions, capacity, mixed blocks, 63 → 30 + 33,
continuous numbering, transaction rollback and college-report checks. The sample
output was generated with an in-memory SQLite adapter, without modifying live
MySQL. `create_input_templates.py --directory "empty/folder"` can create another
example set; it refuses to overwrite existing finalized input files.

## Legacy compatibility

The earlier single-workbook formats described below are still supported. To use
the preserved original workbook explicitly:

```powershell
python -B main.py --input "input/Tables.xlsx"
python -B seating_allocator.py --legacy --exam-date 2026-10-05
```

---

# Earlier one-day / block seating workflow (legacy)

The existing importer, transaction handling and legacy workbook format
remain supported. `input/Tables.xlsx` is unchanged. Nothing needs to be deleted
from MySQL to use the additions below.

## Run for one day

```powershell
# Set MYSQL_PASSWORD in this session, then import the existing workbook:
python -B main.py --input "input/Tables.xlsx"
# Select one date; allocations for all other dates are preserved:
python -B seating_allocator.py --legacy --exam-date 2026-10-05
```

The date argument can be omitted when the database has only one exam date.
Multiple stored dates require an explicit selection; the CLI never silently
generates several days. Database initialization remains `python -B database.py`
on first use. Normal import/allocation automatically applies additive migrations.

## New input format (optional; old input still works)

Create a separate workbook with these three worksheets. Import it explicitly
with `python -B main.py --input "path/to/OneDay.xlsx"`; this does not replace
`input/Tables.xlsx`. Keep the existing room capacities: 201=48, 202=40, 203=45,
204=55, 205=70, unless the actual rooms change.

**Exam Configuration** has setting names in column A and values in column B,
without a header row:

| A | B |
|---|---|
| Exam Date | 2026-10-05 |
| Exam Time | 12:00-13:00 |

Enter the date only here. Time may be blank; otherwise supply a complete range.
All student/subject rows in this format sit this one exam slot. No separate
timetable sheet is needed. An untimed slot cannot be allocated alongside other
slots on the same day until its time is configured.

**Student Subject Data** has these headers in its first nonblank row:

| Class | Department | Subject | Roll Number Start | Roll Number End | Excluded Roll Numbers |
|---|---|---|---:|---:|---|
| Example class | Example department | Subject A | 1 | 40 | 3,8-9 |
| Example class | Example department | Subject B | 41 | 70 | 45 |

Eligible counts are calculated as the range minus exclusions (37 and 29 in this
example), stored in MySQL and shown in allocation output. No supplied count is
needed. Rows with the same class and different subjects are supported. Disjoint
ranges for the same class/subject are combined into one exact roll list, retaining
gaps and exclusions. Repeated/overlapping ranges for that class/subject are
rejected. A class/roll cannot take two subjects in the same slot. Department is
required here; a previously unknown department can be filled, but a conflicting
known department is never silently replaced. There is no teacher field.

**Classroom Data** has headers `Room Number` and `Seating Capacity`, followed by
one row per room. Capacity is the number of individual seats/benches.

Legacy side-by-side college/legacy tables continue to work, including their
per-row timetable dates and existing supplied student counts. Disjoint college
rows for the same class/subject now also combine without filling gaps. As before,
imports update matching records rather than deleting records absent from Excel;
changing a dated exam can leave the old timetable entry in place. Review schedule
changes deliberately instead of deleting historical data automatically.

## Blocks and seat maps

- Finish each exam group before moving to the next, using remaining room seats. Room-ID order is unchanged.
- Seats are numbered from 1 within each room, in the actual sequential allocation order.
- A block is a room-level supervision division with a **target of 30 students**,
  not a maximum or a class/subject division. No teacher assignment is stored.
- For each occupied room, round student count / 30 to the nearest whole number
  (halves round up), with at least one block. Earlier blocks take 30 students;
  the last absorbs the remainder: 63 = 30 + 33, 37 = one block of 37,
  48 = 30 + 18, and 93 = 30 + 30 + 33. Empty rooms get no blocks.
- Blocks follow the actual sequential seat order, can contain several classes,
  departments and subjects, and never span rooms. Groups share a block only at a transition between consecutive groups.
- Block numbers start at 1 for the report and continue across rooms and slots;
  they do not restart in each room. A new daily report starts again at 1.
- No room dimensions are supplied, so the map is a numbered bench/seat list,
  not an invented physical row/column plan.
- Student identity is class + roll number. Duplicate students and duplicate seats
  within a slot are rejected. Legacy slots are still grouped by exact date/time;
  partially overlapping time intervals are not automatically reconciled.

`seating_arrangements` gains nullable `block_number` and `seat_numbers` fields.
Existing rows are not backfilled with guessed seat numbers. On regeneration,
only selected slots are replaced within the existing all-or-nothing transaction.
Each new database row describes a subject's portion of a supervision block;
mixed subjects share the same `block_number`. Its comma-separated seat numbers
map positionally to the ascending rolls expanded from `roll_numbers`.

## Excel output

`output/Seating_Arrangement.xlsx` is generated automatically after allocation
commits. The directory is created if necessary. It contains:

- **Room Blocks** (opens first): date/time at the top, all five summary metrics,
  and room/block/class/department/subject/roll/count/unused-seat detail. Several
  subject rows may share a block number; Total Blocks counts supervision blocks,
  not those detail rows.
- **Seating Arrangement**: the original summary layout for compatibility.
- **Seats 001, Seats 002, ...**: one printable numbered seat list per room/slot,
  including empty seats, room number, date/time, block and student identity.

The workbook uses bold headers, borders, room shading, wrapped text, repeated
print headings and landscape print settings. Unknown legacy departments are
shown as "Not specified" in the new staff sheets. Total capacity used means
occupied seats; totals across slots count exam attendances, not unique people.
The old report is replaced only after a complete new workbook has been saved.
An export failure does not undo committed database allocations.

## Tests

```powershell
python -B -m unittest discover -v
```

The original four test modules are retained. `test_one_day_seating.py` adds
one-day parsing, disjoint groups, exclusions, global blocks, seat-order checks,
selected-day preservation, rollback and report coverage. Tests use temporary
workbooks, in-memory SQLite and mocks; they do not change live MySQL data.

---

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
- Sequential class/semester/subject allocation
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
8. Place each group in ascending roll order, continuing into subsequent rooms as needed.
9. Start the next group only after finishing the current group, using remaining room capacity.
10. Preserve ascending roll-number order within an exam group.
11. Validate the completed allocation.
12. Save the allocation to MySQL.
13. Generate `Seating_Arrangement.xlsx`.

### Simple example

For 99 CS students followed by 19 BA students, with room 201 capacity 60:

- Room 201: CS rolls 1-60, blocks 1 and 2 with 30 students each.
- Room 202: CS rolls 61-99, followed by BA rolls 1-19 (58 students).
- Room 202 blocks 3 and 4 contain 30 CS, then 9 CS + 19 BA.

Blocks remain room-level supervision groups. The target is 30, not a maximum:
63 students in one room still form blocks of 30 and 33. Block numbering remains
global and continuous. No teacher fields are added.

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
