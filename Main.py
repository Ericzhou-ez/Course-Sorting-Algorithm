from ortools.sat.python import cp_model
import random
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment

adst_courses = [
    ("FOOD STUDIES 9", 9),
    ("FOOD STUDIES 10 & INTRO 11", 10),
    ("FOOD STUDIES 11", 11),
    ("FOOD STUDIES 12", 12),
    ("INFORMATION & COMMUNICATIONS TECHNOLOGIES 9", 9),
    ("COMPUTER STUDIES 10", 10),
    ("COMPUTER INFORMATION SYSTEMS 11", 11),
    ("COMPUTER INFORMATION SYSTEMS 12", 12),
    ("COMPUTER PROGRAMMING 12", 12),
    ("DRAFTING 11", 11),
    ("DRAFTING 12", 12),
    ("ENGINEERING 11", 11),
    ("ENGINEERING 12", 12)
]
fine_arts_courses = [
    ("MUSIC 9: CONCERT CHOIR", 9),
    ("CHORAL MUSIC 10: CONCERT CHOIR", 10),
    ("CHORAL MUSIC 11: CONCERT CHOIR", 11),
    ("CHORAL MUSIC 12: CONCERT CHOIR", 12),
    ("DRAMA 9 (ACTING)", 9), ("DRAMA 10 (ACTING)", 10),
    ("DRAMA 11 (ACTING)", 11), ("DRAMA 12 (ACTING)", 12),
    ("VISUAL ARTS 9 (DRAWING & PAINTING)", 9),
    ("STUDIO ARTS 2D 10 (DRAWING & PAINTING)", 10),
    ("STUDIO ARTS 2D 11 (DRAWING & PAINTING)", 11),
    ("STUDIO ARTS 2D 12 (DRAWING & PAINTING)", 12)
]

file_path = 'file/path/to/your/studentCourses.xlsx'
df = pd.read_excel(file_path)
students = {}

for index, row in df.iterrows():
    student_name = row['Student Name']
    student_number = row['Student Number']
    grade = row['Grade']

    courses = row['Courses'].split(', ')
    preferences = row['Preferences']

    if pd.isna(preferences):
        preferences = []
    else:
        preferences = preferences.split(', ')

    # preference values
    courses_with_values = {course: 10000000 for course in courses}
    preferences_with_values = {preferences[i]: (1000 if i < 2 else 100) for i in range(len(preferences))}

    # Identify ADST and Fine Arts courses
    for course in courses_with_values:
        if any(course.startswith(adst_course) for adst_course, _ in adst_courses):
            courses_with_values[course] = (courses_with_values[course], "ADST")
        elif any(course.startswith(fine_art_course) for fine_art_course, _ in fine_arts_courses):
            courses_with_values[course] = (courses_with_values[course], "Fine Arts")
        else:
            courses_with_values[course] = (courses_with_values[course], "General")

    # Store in the students dictionary
    students[student_name] = {
        "number": student_number,
        "grade": grade,
        "courses": courses_with_values,
        "preferences": preferences_with_values
    }

"""
Example output for a student
{'number': 218620, 
'grade': 11, 
'courses': {'FOCUSED LITERARY STUDIES 11': (1000000, 'General'), 
'PRE-CALCULUS 11': (1000000, 'General'), 
'SOCIAL STUDIES 11': (1000000, 'General'), 
'PHYSICS 11': (1000000, 'General'), 
'LIFE SCIENCES 11': (1000000, 'General'), 
'ENGINEERING 11': (1000000, 'ADST'), 
'CHORAL MUSIC 11: CONCERT CHOIR': (1000000, 'Fine Arts')}, 

'preferences': {'DRAFTING 11': 1000, 
'COMPUTER INFORMATION SYSTEMS 11': 1000, 
'ENGINEERING 11': 100, 
'DRAMA 11 (ACTING)': 100, 
'CHORAL MUSIC 11: CONCERT CHOIR': 100}
}
"""

def read_teacher_data(file_path):
    try:
        df = pd.read_excel(file_path)
        teachers = {}

        for _, row in df.iterrows():
            last_name = row['Last Name']
            courses = row['Courses'].split(', ') if isinstance(row['Courses'], str) else []
            adst_rotation = row['ADST Rotation'] == 'y'
            fine_arts_rotation = row['Fine Arts Rotation'] == 'y'

            # Add ADST Rotation to courses if marked 'y'
            if adst_rotation and 'ADST Rotation' not in courses:
                courses.append('ADST Rotation')

            # Add Fine Arts Rotation to courses if marked 'y'
            if fine_arts_rotation and 'Fine Arts Rotation' not in courses:
                courses.append('Fine Arts Rotation')

            # Add teacher to the dictionary
            teachers[last_name] = {
                "courses": courses,
                "ADST_rotation": adst_rotation,
                "Fine_Arts_rotation": fine_arts_rotation,
                "available_times": ["S1P1", "S1P2", "S1P3", "S1P4", "S2P1", "S2P2", "S2P3", "S2P4"]
            }

        return teachers
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")
        return None

file_path = '/Users/ez/PycharmProjects/EduCourse/.venv/TeacherCourseMapping.xlsx'
teachers = read_teacher_data(file_path)


def preprocess_teacher_assignments(teachers, students):
    course_demand = {}
    max_class_size = 28

    offTimeTableMusicCourses = ["MUSIC 9: CONCERT CHOIR",
                                "CHORAL MUSIC 10: CONCERT CHOIR",
                                "CHORAL MUSIC 11: CONCERT CHOIR",
                                "CHORAL MUSIC 12: CONCERT CHOIR"]

    # Calculate course demand from the students' selections, ignoring off-time table music courses
    for student, data in students.items():
        for course, (value, category) in data['courses'].items():
            if course not in offTimeTableMusicCourses:  # Ignore off-time table music courses
                if course not in course_demand:
                    course_demand[course] = 0
                course_demand[course] += 1

    # Check teacher availability and calculate how many teachers are missing for each course
    for course, demand in course_demand.items():
        teachers_available = sum(1 for teacher_data in teachers.values() if course in teacher_data['courses'])
        teachers_needed = (demand + max_class_size - 1) // max_class_size  # Calculate number of teachers needed

        # Only handle cases where not enough teachers are available
        if teachers_available < teachers_needed:
            teachers_missing = teachers_needed - teachers_available
            print(f"Warning: {teachers_missing} more teacher(s) needed for {course}")

            # Calculate the average class size with the available teachers
            if teachers_available > 0:
                avg_class_size = demand // teachers_available  # Distribute students to available teachers
                print(f"Average class size for {course} with available teachers: {avg_class_size} students")
            else:
                print(f"No teachers available for {course}, unable to calculate average class size.")

    return course_demand

# from courseGeneration.py
grade_8_required = [("ENGLISH 8", 8),
                    ("MATHEMATICS 8", 8),
                    ("FRENCH 8", 8),
                    ("ADST Rotation", 8),
                    ("Fine Arts Rotation", 8),
                    ("SOCIAL STUDIES 8", 8),
                    ("PHYSICAL AND HEALTH EDUCATION 8", 8),
                    ("SCIENCE 8", 8)
                    ]
grade_9_required = [("ENGLISH 9", 9),
                    ("MATHEMATICS 9", 9),
                    ("SOCIAL STUDIES 9", 9),
                    ("PHYSICAL AND HEALTH EDUCATION 9", 9),
                    ("SCIENCE 9", 9)
                    ]
grade_10_required = [("ENGLISH 10", 10),
                     ("FOUNDATION OF MATHEMATICS & PRE-CALCULUS 10", 10),
                     ("SOCIAL STUDIES 10", 10),
                     ("PHYSICAL AND HEALTH EDUCATION 10", 10),
                     ("SCIENCE 10", 10),
                     ("CAREER AND LIFE EXPLORATION 10 CLE", 10)
                     ]
grade_11_required = [("FOCUSED LITERARY STUDIES 11", 11),
                     ("PRE-CALCULUS 11", 11),
                     ("SOCIAL STUDIES 11", 11),
                     ]
grade_12_required = [("ENGLISH STUDIES 12", 12),
                     ("PRE-CALCULUS 12", 12),
                     ("CAREER & LIFE CONNECTIONS (CLC) & CAPSTONE", 12)
                     ]
language_courses = [("FRENCH 9", 9),
                    ("FRENCH 10", 10),
                    ("FRENCH 11", 11),
                    ("FRENCH 12", 12),
                    ("SPANISH 10", 10),
                    ("SPANISH 11", 11),
                    ("SPANISH 12", 12)
                    ]
adst_courses = [("FOOD STUDIES 9", 9), ("FOOD STUDIES 10 & INTRO 11", 10),
                ("FOOD STUDIES 11", 11),
                ("FOOD STUDIES 12", 12),
                ("INFORMATION & COMMUNICATIONS TECHNOLOGIES 9", 9),
                ("COMPUTER STUDIES 10", 10),
                ("COMPUTER INFORMATION SYSTEMS 11", 11),
                ("COMPUTER INFORMATION SYSTEMS 12", 12),
                ("COMPUTER PROGRAMMING 12", 12),
                ("DRAFTING 11", 11),
                ("DRAFTING 12", 12),
                ("ENGINEERING 11", 11),
                ("ENGINEERING 12", 12)
                ]
fine_arts_courses = [("MUSIC 9: CONCERT CHOIR", 9),
                     ("CHORAL MUSIC 10: CONCERT CHOIR", 10),
                     ("CHORAL MUSIC 11: CONCERT CHOIR", 11),
                     ("CHORAL MUSIC 12: CONCERT CHOIR", 12),
                     ("DRAMA 9 (ACTING)", 9), ("DRAMA 10 (ACTING)", 10),
                     ("DRAMA 11 (ACTING)", 11), ("DRAMA 12 (ACTING)", 12),
                     ("VISUAL ARTS 9 (DRAWING & PAINTING)", 9),
                     ("STUDIO ARTS 2D 10 (DRAWING & PAINTING)", 10),
                     ("STUDIO ARTS 2D 11 (DRAWING & PAINTING)", 11),
                     ("STUDIO ARTS 2D 12 (DRAWING & PAINTING)", 12)
                     ]
science_11_12 = [("LIFE SCIENCES 11", 11),
                 ("ANATOMY AND PHYSIOLOGY 12", 12),
                 ("CHEMISTRY 11", 11),
                 ("CHEMISTRY 12", 12),
                 ("PHYSICS 11", 11),
                 ("PHYSICS 12", 12)
                 ]
grade_12_electives = [("AP MICROECONOMICS 12", 12),
                      ("AP ENGLISH LITERATURE AND COMPOSITION 12", 12),
                      ("AP CALCULUS 12 AB", 12),
                      ("AP STATISTICS 12", 12),
                      ("AP BIOLOGY 12", 12),
                      ("AP CHEMISTRY 12", 12),
                      ("AP PHYSICS 2 HONOURS 12", 12),
                      ("AP HUMAN GEOGRAPHY 12", 12)
                      ]

# main function
def create_schedule(students, teachers):
    model = cp_model.CpModel()
    periods = ['S1P1', 'S1P2', 'S1P3', 'S1P4', 'S2P1', 'S2P2', 'S2P3', 'S2P4']
    min_class_size = 15
    ideal_class_size = 25
    max_class_size = 34
    max_teacher_courses = 7

    # Initialize assigned courses counter for each student
    for student in students:
        students[student]['assigned_courses'] = 0

    # Define ADST and Fine Arts courses
    adst_courses = ['ENGINEERING 11', 'ENGINEERING 12', 'DRAFTING 11', 'DRAFTING 12', 'FOOD STUDIES 9',
                    'FOOD STUDIES 10 & INTRO 11', 'COMPUTER STUDIES 10', 'COMPUTER INFORMATION SYSTEMS 11',
                    'COMPUTER INFORMATION SYSTEMS 12', 'COMPUTER PROGRAMMING 12']

    fine_arts_courses = ['VISUAL ARTS 9', 'STUDIO ARTS 2D 10', 'STUDIO ARTS 2D 11', 'DRAMA 9 (ACTING)',
                         'DRAMA 10 (ACTING)', 'DRAMA 11 (ACTING)', 'DRAMA 12 (ACTING)']

    off_timetable_courses = ['MUSIC 9: CONCERT CHOIR', 'CHORAL MUSIC 10: CONCERT CHOIR',
                             'CHORAL MUSIC 11: CONCERT CHOIR', 'CHORAL MUSIC 12: CONCERT CHOIR']

    # Create variables
    student_assignments = {}
    teacher_assignments = {}
    class_sizes = {}
    off_timetable_assignments = {}

    # Decision Variables: Student Assignments
    for student, data in students.items():
        for course, (value, category) in data['courses'].items():
            if course in off_timetable_courses:
                off_timetable_assignments[(student, course)] = model.NewBoolVar(f'{student}_{course}_off_timetable')
            else:
                for period in periods:
                    student_assignments[(student, course, period)] = model.NewBoolVar(f'{student}_{course}_{period}')

    # Decision Variables: Teacher Assignments
    for teacher, data in teachers.items():
        for course in data['courses']:
            if course in off_timetable_courses:
                off_timetable_assignments[(teacher, course)] = model.NewBoolVar(f'{teacher}_{course}_off_timetable')
            else:
                for period in periods:
                    teacher_assignments[(teacher, course, period)] = model.NewBoolVar(f'{teacher}_{course}_{period}')

    # Decision Variables: Class Sizes
    all_courses = set(course for student_data in students.values() for course in student_data['courses'] if
                      course not in off_timetable_courses)
    for course in all_courses:
        for period in periods:
            class_sizes[(course, period)] = model.NewIntVar(0, len(students), f'size_{course}_{period}')

    # Constraints passed into solver
    # 1. Ensure each student gets between 6 and 8 courses (including off-timetable)
    for student, data in students.items():
        regular_courses = sum(student_assignments[(student, course, period)]
                              for course in students[student]['courses'] if course not in off_timetable_courses
                              for period in periods)
        off_timetable_courses_count = sum(off_timetable_assignments[(student, course)]
                                          for course in students[student]['courses'] if course in off_timetable_courses)

        model.Add(regular_courses + off_timetable_courses_count >= 6)  # Ensure at least 6 courses
        model.Add(regular_courses + off_timetable_courses_count <= 8)  # No more than 8 courses

    # 2. No overlapping classes for students (regular timetable)
    for student in students:
        for period in periods:
            model.Add(sum(student_assignments[(student, course, period)]
                          for course in students[student]['courses'] if course not in off_timetable_courses) <= 1)

    # 3. Teacher assignment constraints (each teacher can have up to 8 courses across all periods)
    for teacher, data in teachers.items():
        model.Add(sum(teacher_assignments[(teacher, course, period)]
                      for course in data['courses'] if course not in off_timetable_courses
                      for period in periods) <= max_teacher_courses)

    # 4. Ensure off-timetable courses are assigned without interfering with regular periods
    for course in off_timetable_courses:
        for teacher in teachers:
            if course in teachers[teacher]['courses']:
                for student in students:
                    if course in students[student]['courses']:
                        model.Add(off_timetable_assignments[(student, course)] == 1)
                        model.Add(off_timetable_assignments[(teacher, course)] == 1)

    # 5. Merge small classes where possible
    for course in all_courses:
        total_students = sum(student_assignments[(student, course, period)]
                             for student in students for period in periods if course in students[student]['courses'])
        class_active = model.NewBoolVar(f'{course}_active')

        # Create boolean to check if a class is active (more than zero students)
        model.Add(total_students > 0).OnlyEnforceIf(class_active)
        model.Add(total_students == 0).OnlyEnforceIf(class_active.Not())

        # Ensure the class size stays within the min/max limits when active
        model.Add(total_students >= min_class_size).OnlyEnforceIf(class_active)
        model.Add(total_students <= max_class_size).OnlyEnforceIf(class_active)

        # Distribute students across fewer periods if total is less than the ideal size
        num_periods_active = model.NewIntVar(1, len(periods), f'{course}_periods_active')
        model.AddDivisionEquality(num_periods_active, total_students, ideal_class_size)

        for period in periods:
            model.Add(class_sizes[(course, period)] <= ideal_class_size).OnlyEnforceIf(class_active)

    # Objective function: Prioritize student preferences and maximize course assignments
    objective_terms = []

    # Maximize student course preferences
    for student, data in students.items():
        for course, preference in data['preferences'].items():
            if course in off_timetable_courses:
                if (student, course) in off_timetable_assignments:
                    objective_terms.append(off_timetable_assignments[(student, course)] * preference)
            else:
                for period in periods:
                    if (student, course, period) in student_assignments:
                        objective_terms.append(student_assignments[(student, course, period)] * preference)

    # Maximize number of courses assigned to each student to avoid incomplete schedules
    for student, data in students.items():
        regular_courses_assigned = sum(student_assignments[(student, course, period)]
                                       for course in students[student]['courses'] if course not in off_timetable_courses
                                       for period in periods)
        off_timetable_courses_assigned = sum(off_timetable_assignments[(student, course)]
                                             for course in students[student]['courses'] if course in off_timetable_courses)
        total_courses_assigned = regular_courses_assigned + off_timetable_courses_assigned
        objective_terms.append(total_courses_assigned)

    # Maximize teacher assignments to their courses (including off-timetable)
    for teacher, data in teachers.items():
        for course in data['courses']:
            if course in off_timetable_courses:
                if (teacher, course) in off_timetable_assignments:
                    objective_terms.append(off_timetable_assignments[(teacher, course)] * 1000)
            else:
                for period in periods:
                    if (teacher, course, period) in teacher_assignments:
                        objective_terms.append(teacher_assignments[(teacher, course, period)] * 1000)

    # Set the objective to maximize the sum of preferences and assignments
    model.Maximize(sum(objective_terms))

    # Solve
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 100.0 # change solver max time as appropriate
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        schedule = process_results(solver, student_assignments, teacher_assignments, class_sizes,
                                   off_timetable_assignments, students, teachers, periods)
        export_schedule(schedule)
        print("Schedule exported successfully.")
        return schedule
    else:
        print(f"No solution found. Solver status: {solver.StatusName()}")
        partial_schedule = process_partial_results(solver, student_assignments, teacher_assignments, class_sizes,
                                                   off_timetable_assignments, students, teachers, periods)
        if partial_schedule:
            export_schedule(partial_schedule)
            print("Partial schedule exported.")
        else:
            print("No partial schedule could be created.")
        return partial_schedule

def process_results(solver, student_assignments, teacher_assignments, class_sizes, off_timetable_assignments, students, teachers, periods):
    schedule = {period: {} for period in periods}
    consolidated_class_sizes = {}  # To store combined class sizes across periods

    # Process regular student assignments
    for (student, course, period), var in student_assignments.items():
        if solver.Value(var):
            if course not in schedule[period]:
                schedule[period][course] = {'students': [], 'teacher': None}
            schedule[period][course]['students'].append(student)

    # Process teacher assignments
    for (teacher, course, period), var in teacher_assignments.items():
        if solver.Value(var):
            if course in schedule[period]:
                schedule[period][course]['teacher'] = teacher

    # Process off-timetable assignments
    for (student, course), var in off_timetable_assignments.items():
        if solver.Value(var):
            # Off-timetable courses won't be part of the regular periods
            if 'off_timetable' not in schedule:
                schedule['off_timetable'] = {}
            if course not in schedule['off_timetable']:
                schedule['off_timetable'][course] = {'students': [], 'teacher': None}
            schedule['off_timetable'][course]['students'].append(student)

    for (teacher, course), var in off_timetable_assignments.items():
        if solver.Value(var):
            if course in schedule['off_timetable']:
                schedule['off_timetable'][course]['teacher'] = teacher

    # Combine class sizes across periods
    for course, period in class_sizes:
        class_size = solver.Value(class_sizes[(course, period)])

        if course not in consolidated_class_sizes:
            consolidated_class_sizes[course] = 0

        consolidated_class_sizes[course] += class_size

    # Filter and print only classes with size >= 15
    for course, total_size in consolidated_class_sizes.items():
        if total_size >= 15:
            print(f"Class size for {course}: {total_size}")
        else:
            # Mark students in this course as having fewer classes if class size is too small
            for period in periods:
                if course in schedule[period]:
                    for student in schedule[period][course]['students']:
                        students[student]['assigned_courses'] -= 1
                    del schedule[period][course]  # Remove the course from the schedule
    return schedule

def process_partial_results(solver, student_assignments, teacher_assignments, class_sizes, off_timetable_assignments, students, teachers, periods):
    schedule = {period: {} for period in periods}

    # Check if a feasible or optimal solution was found
    if solver.StatusName() not in ['OPTIMAL', 'FEASIBLE']:
        print(f"No valid solution found. Solver status: {solver.StatusName()}")
        return None  # Return None if no valid solution exists

    # Process student assignments
    for (student, course, period), var in student_assignments.items():
        if solver.Value(var):  # Check if this variable has a valid assignment
            if course not in schedule[period]:
                schedule[period][course] = {'students': [], 'teacher': None}
            schedule[period][course]['students'].append(student)

    # Process teacher assignments
    for (teacher, course, period), var in teacher_assignments.items():
        if solver.Value(var):
            if course in schedule[period]:
                schedule[period][course]['teacher'] = teacher

    # Process off-timetable assignments
    for (student, course), var in off_timetable_assignments.items():
        if solver.Value(var):
            # Off-timetable courses won't be part of the regular periods
            if 'off_timetable' not in schedule:
                schedule['off_timetable'] = {}
            if course not in schedule['off_timetable']:
                schedule['off_timetable'][course] = {'students': [], 'teacher': None}
            schedule['off_timetable'][course]['students'].append(student)

    for (teacher, course), var in off_timetable_assignments.items():
        if solver.Value(var):
            if course in schedule['off_timetable']:
                schedule['off_timetable'][course]['teacher'] = teacher
    return schedule

departments = {
        'ADST': ['ENGINEERING 11', 
                 'ENGINEERING 12', 
                 'DRAFTING 11', 
                 'DRAFTING 12', 
                 'FOOD STUDIES 9',
                 'FOOD STUDIES 10 & INTRO 11', 
                 'COMPUTER STUDIES 10', 
                 'COMPUTER INFORMATION SYSTEMS 11',
                 'COMPUTER INFORMATION SYSTEMS 12', 
                 'COMPUTER PROGRAMMING 12'],
        'Fine Arts': ['VISUAL ARTS 9', 
                      'STUDIO ARTS 2D 10', 
                      'STUDIO ARTS 2D 11', 
                      'DRAMA 9 (ACTING)', 
                      'DRAMA 10 (ACTING)',
                      'DRAMA 11 (ACTING)', 
                      'DRAMA 12 (ACTING)', 
                      'MUSIC 9: CONCERT CHOIR', 
                      'CHORAL MUSIC 10: CONCERT CHOIR',
                      'CHORAL MUSIC 11: CONCERT CHOIR', 
                      'CHORAL MUSIC 12: CONCERT CHOIR'],
        # Add other departments as needed
}

def export_schedule(schedule):
    # Invert the departments mapping to map courses to departments
    course_to_department = {}
    for department, courses in departments.items():
        for course in courses:
            course_to_department[course] = department
    rows = []
  
    schedule_flat = []
    for period, courses in schedule.items():
        for course, data in courses.items():
            teacher = data['teacher'] if data['teacher'] is not None else 'Unassigned'
            students = ', '.join(data['students']) if data['students'] else 'No students assigned'
            class_size = len(data['students']) if data['students'] else 0
            department = course_to_department.get(course, 'General')  # Default to 'General' if not categorized
            schedule_flat.append({
                'Period': period,
                'Course': course,
                'Department': department,
                'Teacher': teacher,
                'Class Size': class_size,
                'Students': students
            })

    schedule_flat.sort(key=lambda x: (x['Department'], x['Course'], x['Period']))

    current_department = None
    for entry in schedule_flat:
        if entry['Department'] != current_department:
            # Insert a row to indicate the department (for grouping)
            rows.append({'Period': '', 'Course': f"Department: {entry['Department']}", 'Teacher': '', 'Class Size': '', 'Students': ''})
            current_department = entry['Department']

        rows.append({
            'Period': entry['Period'],
            'Course': entry['Course'],
            'Teacher': entry['Teacher'],
            'Class Size': entry['Class Size'],
            'Students': entry['Students']
        })

    # Convert to DataFrame and export to Excel
    df = pd.DataFrame(rows)
    df.to_excel('school_schedule_by_department.xlsx', index=False)
    print("Schedule exported to school_schedule_by_department.xlsx")

def export_student_schedule(students, schedule):
    rows = []
    periods = ['S1P1', 'S1P2', 'S1P3', 'S1P4', 'S2P1', 'S2P2', 'S2P3', 'S2P4']
    students_with_less_than_8_courses = []

    for student_name, data in students.items():
        student_number = data['number']
        student_row = {
            'Student Name': student_name,
            'Student Number': student_number,
        }
        assigned_courses_count = 0
      
        for period in periods:
            assigned_course = None
            for course, course_data in schedule[period].items():
                if student_name in course_data['students']:
                    assigned_course = course
                    assigned_courses_count += 1
                    break  # We found the course for this period

            # If the student is not assigned to a course in this period, leave it blank
            student_row[period] = assigned_course if assigned_course else 'Free'
        rows.append(student_row)

        # Track students with fewer than 8 courses
        if assigned_courses_count < 8:
            students_with_less_than_8_courses.append((student_name, student_number, assigned_courses_count))

    df = pd.DataFrame(rows)
    df.to_excel('student_schedules.xlsx', index=False)
    print("Student schedules exported to student_schedules.xlsx")

    # Export the list of students with fewer than 8 courses
    if students_with_less_than_8_courses:
        df_missing_courses = pd.DataFrame(students_with_less_than_8_courses, columns=['Student Name', 'Student Number', 'Assigned Courses'])
        df_missing_courses.to_excel('students_with_less_than_8_courses.xlsx', index=False)
        print("List of students with fewer than 8 courses exported to students_with_less_than_8_courses.xlsx")

if __name__ == "__main__":
    # Preprocess teacher assignments
    teachers = preprocess_teacher_assignments(teachers, students)

    # Create schedule
    schedule = create_schedule(students, teachers)

    if schedule:
        print("Schedule created successfully.")
        export_schedule(schedule)
        export_student_schedule(students, schedule)
    else:
        print("Failed to create a complete schedule.")
