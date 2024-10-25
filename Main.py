from ortools.sat.python import cp_model
import random
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment

import config
from config import (grade_8_required, grade_9_required, grade_10_required, grade_11_required, grade_12_required, language_courses, adst_courses, fine_arts_courses, science_11_12, grade_12_electives
)
from input import students, teachers

off_timetable_courses = ["MUSIC 9: CONCERT CHOIR",
                                "CHORAL MUSIC 10: CONCERT CHOIR",
                                "CHORAL MUSIC 11: CONCERT CHOIR",
                                "CHORAL MUSIC 12: CONCERT CHOIR"]

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
                                
# determines roughly if sorting is feasible
def preprocess_teacher_assignments(teachers, students):
    course_demand = {}

    # Calculate course demand from the students' selections, ignoring off-time table music courses
    for student, data in students.items():
        for course, (value, category) in data['courses'].items():
            if course not in off_timetable_courses:  # Ignore off-time table music courses
                if course not in course_demand:
                    course_demand[course] = 0
                course_demand[course] += 1

    # Check teacher availability and calculate how many teachers are missing for each course
    for course, demand in course_demand.items():
        teachers_available = sum(1 for teacher_data in teachers.values() if course in teacher_data['courses'])
        teachers_needed = (demand + config.MAX_CLASS_SIZE - 1) // config.MAX_CLASS_SIZE  # Calculate number of teachers needed

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

def create_schedule(students, teachers):
    model = cp_model.CpModel()
    periods = ['S1P1', 'S1P2', 'S1P3', 'S1P4', 'S2P1', 'S2P2', 'S2P3', 'S2P4']
    min_class_size = config.MIN_CLASS_SIZE
    ideal_class_size = config.IDEAL_CLASS_SIZE
    max_class_size = config.MAX_CLASS_SIZE
    max_teacher_courses = 7

    # Initialize assigned courses counter for each student
    for student in students:
        students[student]['assigned_courses'] = 0

    # Create variables
    student_assignments = {}
    teacher_assignments = {}
    class_sizes = {}
    off_timetable_assignments = {}

    # Decision Variables: Student Assignments
    for student, data in students.items():
        for course, (value, category) in data['courses'].items():
            # in off-time table period is not a unique identifier
            if course in off_timetable_courses:
                off_timetable_assignments[(student, course)] = model.NewBoolVar(f'{student}_{course}_off_timetable')
            # for "normal" classes, student, course, and every period is needed to be the unique identifier
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
    # every combination of courses and period
    for course in all_courses:
        for period in periods:
            class_sizes[(course, period)] = model.NewIntVar(0, len(students), f'size_{course}_{period}')

    # Constraints passed into solver
    # 1. No overlapping classes
    for student in students:
        for period in periods:
            model.Add(sum(student_assignments[(student, course, period)]
                          for course in students[student]['courses'] 
                          if course not in off_timetable_courses) <= 1
                     )

    # 2. 6-8 courses
    for student in students.keys():
        
        regular_courses = sum(
            student_assignments[(student, course, period)]
            for course in students[student]['courses'] if course not in off_timetable_courses
            for period in periods
        )

        off_timetable_courses_count = sum(
            off_timetable_assignments[(student, course)]
            for course in students[student]['courses'] if course in off_timetable_courses
        )

        model.Add(regular_courses + off_timetable_courses_count >= 6)
        model.Add(regular_courses + off_timetable_courses_count <= 8)

    # On regular time table teachers must have less than or equal to 7 courses
    for teacher, data in teachers.items():
        total_courses_assigned = sum(
        teacher_assignments[(teacher, course, period)]
        for course in data['courses'] if course not in off_timetable_courses
        for period in periods
        )

        model.Add(total_courses_assigned <= max_teacher_courses)

    # # 4. Merge small classes where possible -- absolutely not necessary
    # for course in all_courses:
    #     total_students = sum(student_assignments[(student, course, period)]
    #                         for student in students 
    #                         for period in periods 
    #                         if course in students[student]['courses']
    #                         )

    #     class_active = model.NewBoolVar(f'{course}_active')

    #     model.Add(total_students > 0).OnlyEnforceIf(class_active)
    #     model.Add(total_students == 0).OnlyEnforceIf(class_active.Not())

    #     model.Add(total_students >= min_class_size).OnlyEnforceIf(class_active)
    #     model.Add(total_students <= max_class_size).OnlyEnforceIf(class_active)

    #     num_periods_active = model.NewIntVar(1, len(periods), f'{course}_periods_active')
    #     model.AddDivisionEquality(num_periods_active, total_students, ideal_class_size)

    #     for period in periods:
    #         model.Add(class_sizes[(course, period)] <= ideal_class_size).OnlyEnforceIf(class_active)

    # Objective function: Prioritize student preferences and maximize course assignments
    objective_terms = []

    # Maximize student course preferences
    for student, data in students.items():
        for course, preference in data['preferences'].items():
            if course not in off_timetable_courses:
                for period in periods:
                    if (student, course, period) in student_assignments:
                        objective_terms.append(student_assignments[(student, course, period)] * preference)

    # Maximize number of courses assigned to each student (excluding off-timetable)
    for student, data in students.items():
        regular_courses_assigned = sum(student_assignments[(student, course, period)]
                                    for course in students[student]['courses'] if course not in off_timetable_courses
                                    for period in periods
                                    )
        objective_terms.append(regular_courses_assigned)

    # maximize the sum of preference values 
    model.Maximize(sum(objective_terms))

    for (teacher, course, period), teacher_var in teacher_assignments.items():
        student_vars_for_course_period = [
            student_assignments[(student, course, period)]
            for student in students
            if course in students[student]['courses']
        ]
        model.AddMaxEquality(teacher_var, student_vars_for_course_period)

    # Solve
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = config.MAX_MODEL_SOLVER_TIME
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        schedule = process_results(solver, student_assignments, teacher_assignments, class_sizes,
                                   students, teachers, periods)
        # export_schedule(schedule)
        return schedule
    else:
        print(f"No solution found. Solver status: {solver.StatusName()}")

def process_results(solver, student_assignments, teacher_assignments, class_sizes, students, teachers, periods):
    schedule = {period: {} for period in periods}
    consolidated_class_sizes = {}  # To store combined class sizes across periods

    # Process regular student assignments
    for (student, course, period), var in student_assignments.items():
        # solver.Value(var) = 0 or 1
        if solver.Value(var):
            if course not in schedule[period]:
                schedule[period][course] = {'students': [], 'teacher': None}
            schedule[period][course]['students'].append(student)

    # Process teacher assignments
    for (teacher, course, period), var in teacher_assignments.items():
        if solver.Value(var):
            if course in schedule[period]:
                schedule[period][course]['teacher'] = teacher

    # Combine class sizes across periods
    for (course, period), var in class_sizes.items():
        class_size = solver.Value(var)

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

    print(schedule)
    return schedule

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
      
        # Process regular courses for each period
        for period in periods:
            assigned_course = None
            for course, course_data in schedule[period].items():
                if student_name in course_data['students']:
                    assigned_course = course
                    assigned_courses_count += 1
                    break
            student_row[period] = assigned_course if assigned_course else 'Free'

        # Assign off-timetable courses if needed
        off_timetable_courses = [course for course, (value, category) in data['courses'].items() if category == 'Off-Timetable']

        if assigned_courses_count < 8 and off_timetable_courses:
            remaining_courses_needed = 8 - assigned_courses_count
            for i in range(remaining_courses_needed):
                if i < len(off_timetable_courses):
                    student_row[f'Off-Timetable Course {i+1}'] = off_timetable_courses[i]
                    assigned_courses_count += 1

        rows.append(student_row)

        # Track students with fewer than 8 courses
        if assigned_courses_count < 8:
            students_with_less_than_8_courses.append((student_name, student_number, assigned_courses_count))

    # Export student schedules to an Excel file
    df = pd.DataFrame(rows)
    df.to_excel('student_schedules.xlsx', index=False)
    print("Student schedules exported to student_schedules.xlsx")

    # Export the list of students with fewer than 8 courses
    if students_with_less_than_8_courses:
        df_missing_courses = pd.DataFrame(students_with_less_than_8_courses, columns=['Student Name', 'Student Number', 'Assigned Courses'])
        df_missing_courses.to_excel('students_with_less_than_8_courses.xlsx', index=False)
        print("List of students with fewer than 8 courses exported to students_with_less_than_8_courses.xlsx")

if __name__ == "__main__":
    schedule = create_schedule(students, teachers)

    if schedule:
        print("Schedule created successfully.")
        export_schedule(schedule)
        export_student_schedule(students, schedule)
    else:
        print("Failed to create a complete schedule.")
