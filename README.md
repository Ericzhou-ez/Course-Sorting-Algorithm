# Course-Sorting-Algorithm

This Course Sorting Algorithm is tailored for [University Hill Secondary School](https://en.wikipedia.org/wiki/University_Hill_Secondary_School). The project aims to automate the course scheduling process using Google OR-Tools' SAT solver, creating an Intelligent Learning Platform (ILP) that optimizes course schedules while balancing student preferences, teacher availability, and various educational constraints.

## Features

- Automated course scheduling for grades 8-12
- Optimization of student preferences and teacher assignments
- Handling of specialized rotations
- Flexible constraint management for class sizes and teacher workloads
- Outlier and constraint violation reporting

## Requirements

- Python 3.7+
- [Google OR-Tools SAT](https://developers.google.com/optimization)
- Pandas
- Openpyxl
- os
- Random

## Installation

1. Clone the repository:
```
git clone https://github.com/Ericzhou-ez/Course-Sorting-Algorithm.git
cd Course-Sorting-Algorithm
```
4. Install required packages:
```
pip install -r requirements.txt
```

## Usage

1. Prepare your input data:
   
   a. Student data file (`studentCourses.xlsx`):
      - Format: CSV
      - Required columns: Student Name, Student Number, Grade, Courses, Additional Preference Courses
      - Example row:
        ```
        Kiaan Carter,218620,11,"FOCUSED LITERARY STUDIES 11, PRE-CALCULUS 11, SOCIAL STUDIES 11, CHEMISTRY 11, PHYSICS 11, DRAFTING 11, STUDIO ARTS 2D 11 (DRAWING & PAINTING), SPANISH 11","DRAFTING 11, COMPUTER INFORMATION SYSTEMS 11, FOOD STUDIES 11, CHORAL MUSIC 11: CONCERT CHOIR, STUDIO ARTS 2D 11 (DRAWING & PAINTING)"
        ```

   b. Teacher data file (`TeacherCourse.xlsx`):
      - Format: CSV
      - Required columns: Last Name, First Name, Courses
      - Example row:
        ```
        Eric Zhou,"SOCIAL STUDIES 9, SOCIAL STUDIES 10, SOCIAL STUDIES 11", "Law 12"
        ```

2. Configure optimization parameters (optional):
   - Open `main.py` and adjust the following variables:
     - `MIN_CLASS_SIZE`: Minimum number of students per class (default: 24)
     - `MAX_CLASS_SIZE`: Maximum number of students per class (default: 34)
     - `MAX_TEACHER_COURSES`: Maximum number of courses a teacher can teach (default: 7)
   - These parameters allow you to customize the optimal class sizes and teacher workload according to your school's needs.

3. Ensure both input files are in the same directory as the main script.

4. Run ```python main.py```

6. Check the output:
- `school_schedule.xlsx`: Generats schedule.
- `students_with_less_than_8_courses.xlsx`: Reports of any constraint violations or outliers.

Note: Adjusting the optimization parameters may affect the algorithm's ability to find a feasible solution. If you encounter issues, try adjusting these parameters to better fit your school's specific constraints and requirements.

## Project Structure
- `main.py`: Entry point of the application, pre-processing, sorting, and exporting.
    - `def preprocess_teacher_assignments(teachers, students)` checks if student demand for a course is greater than school's offering of that course.
- `courseGeneration.py`: Generates random course selections from `student.xlsx`.
- `studentDataSummary.py`: Determine the number of students in each grade and the number of students in each course (sorted).
- `courseCode.py`: Contains a dictionary mapping the course name to the course code.

## Contributing
Please don't bother.
