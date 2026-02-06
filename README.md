# Course-Sorting-Algorithm

High-school course scheduling using Google OR-Tools CP-SAT. Modular and **highly configurable** for different schools (semester blocks, class sizes, rotations, column names).

## Features

- **Data alignment first**: Validates demand (student requests) vs supply (teacher capacity) before solving; refuses to run if courses have no teacher.
- **Section-based SAT model**: No hard "6–8 courses" rule; maximizes assigned course-periods with class size bounds (min/ideal/max; max = room + slack).
- **Grade 8 rotations**: Optional dynamic rotations (e.g. 2-of-3 options per rotation); generic course names; configurable per school.
- **Off-timetable courses**: e.g. Concert Choir—excluded from placement; solver ignores them.
- **Configurable**: Periods, capacities, capacity slack (+5 over room), column names, rotation definitions, solver time.

## Requirements

- Python 3.8+
- See `requirements.txt`: `ortools`, `pandas`, `openpyxl`

## Installation

```bash
git clone https://github.com/Ericzhou-ez/Course-Sorting-Algorithm.git
cd Course-Sorting-Algorithm
pip install -r requirements.txt
```

## Usage

1. **Prepare inputs** (Excel):
   - **Teachers**: `TeacherCourseMapping.xlsx` — columns: Last Name, First Name, Courses, ADST Rotation, Fine Arts Rotation, Classes, Room Capcity.
   - **Students**: `studentCourses.xlsx` — columns: Student Name, Student Number, Grade, Courses, Preferences.  
   Courses can be comma- or period-separated; the loader normalizes (e.g. `CHORAL MUSIC 12. Fine_Arts_rotation` is split correctly).

2. **Run from repo root**:
   ```bash
   python main.py
   ```
   Options:
   - `--teachers PATH`  
   - `--students PATH`  
   - `--out-dir DIR` (directory for all outputs; default: `output`)  
   - `--no-require-alignment` (run even if demand/supply alignment fails)  
   - `--time SECONDS` (solver time limit)

3. **Outputs** (all written into `output/` by default):
   - `output/school_schedule.xlsx`: Period, Course, Teacher, Room, Students, Class Size.
   - `output/student_schedules.xlsx`: Student Name, Student Number, Grade, then one column per period showing each student's course.

## Configuration

Edit `scheduler/config.py` (or extend `SchedulerConfig`) to change:

- **Periods**: `DEFAULT_PERIODS`, `OFF_TIMETABLE_BLOCK`
- **Class size**: `DEFAULT_MIN_CLASS_SIZE`, `DEFAULT_IDEAL_CLASS_SIZE`, `DEFAULT_CAPACITY_SLACK` (+5 over room), `DEFAULT_GLOBAL_MAX_CLASS_SIZE`
- **Teacher**: `DEFAULT_MAX_TEACHER_SECTIONS` (e.g. 7)
- **Off-timetable courses**: `DEFAULT_OFF_TIMETABLE_COURSES` (e.g. choir)
- **Solver**: `DEFAULT_SOLVER_TIME_SECONDS`, `SOLVER_SYMMETRY_BREAK_PER_COURSE` (default False)
- **Rotations**: `DEFAULT_ROTATIONS` — each has `id`, `display_name`, `grade`, `num_slots_per_student` (e.g. 2), `num_options` (e.g. 3), `option_display_names`
- **Excel columns**: `TEACHER_COLUMNS`, `STUDENT_COLUMNS` (for different header names)

## Project Structure

- **`main.py`**: CLI entry; load → validate → solve → export.
- **`scheduler/`**: Core package.
  - **`config.py`**: All tunables; `SchedulerConfig`, `RotationDef`.
  - **`data/`**: `load.py` (teachers, students; course normalization), `validate.py` (demand vs supply), `load_and_validate()`.
  - **`solver/`**: `model.py` (CP-SAT model), `solve.py` (run solver, return schedule).
  - **`rotation.py`**: Assign N-of-M options for rotation sections (e.g. 2 of 3 for G8).
  - **`export.py`**: Write timetable and underloaded-student Excel.
- **`exampleInput/`**: Sample `TeacherCourseMapping.xlsx`, `studentCourses.xlsx`, `student.xlsx`.
- **`courseGeneration.py`**: Generates random `studentCourses.xlsx` from `student.xlsx` (for testing; real deployment uses real requests).
- **`studentDataSummary.py`**: Summarizes enrollments by grade and course.
- **`courseCode.py`**: Course name → code mapping (optional).
- **`Main.py`**: Legacy single-file solver (kept for reference).

## Data alignment

Before solving, the pipeline checks:

- Every course that appears in student requests has at least one qualified teacher.
- For each such course, total teacher section-slots can meet demand at max class size.

If alignment fails, the run exits with a clear message unless `--no-require-alignment` is set. Fix input data (course names, teacher assignments, typos like period-for-comma in Courses) so demand and supply match.

## Contributing

Please don’t bother.
