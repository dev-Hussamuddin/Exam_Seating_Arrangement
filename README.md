# Exam seating reports

From this project directory in PowerShell, with the existing MySQL connection
settings and `MYSQL_PASSWORD` configured, run:

```powershell
.\.venv\Scripts\python.exe .\seating_allocator.py
```

After the existing allocation transaction commits successfully, the command
automatically creates `output/Seating_Arrangement.xlsx`. The output directory is
created if needed. Close the report in Excel before regenerating it.

The report lists exam date, start and end time, room number, room capacity,
class, subject, exact allocated roll ranges, allocated count and unused seats.
Each allocated timetable group has its own row under its room. Room capacity
and unused seats appear only on the first group row to avoid double counting.
Completely unused rooms are included with zero allocated students.

The top summary counts student-exam attendances across all exam slots (a student
taking two exams is counted twice). Total Capacity Used is occupied seats.
Total Unused Seats includes unused capacity in every room for every slot.

Export uses the successful allocation's in-memory results and does not read or
modify `input/Tables.xlsx`. The importer remains a separate command (`main.py`).
No report is written if allocation fails or there are no timetable slots.
An export failure returns exit code 1 with a message that allocation already
committed; it does not undo database work or replace the previous report.

Run all offline tests (no MySQL connection required):

```powershell
.\.venv\Scripts\python.exe -m unittest discover -v
```
