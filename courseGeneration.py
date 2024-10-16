import pandas as pd
import os
import random
import openpyxl
from openpyxl.styles import Font, Alignment

from config import (grade_8_required, grade_9_required, grade_10_required, grade_11_required, grade_12_required, language_courses, adst_courses, fine_arts_courses, science_11_12, grade_12_electives
)
# test cases: max is 1000 students from student.xlsx. Random student-course is generated to studentCourses.xlsx
NUM_STUDENTS = 800

file_path_students_names = '/Users/ez/Documents/Course-Sorting-Algorithm/exampleInput/student.xlsx'
if not os.path.exists(file_path_students_names):
    print(f"Error: The file {file_path_students_names} does not exist.")
    exit()

df_student_names = pd.read_excel(file_path_students_names)
studentsNames = df_student_names.iloc[:, 0].tolist()
selected_students = studentsNames[:NUM_STUDENTS]

student_dict = {name: 218620 + index for index, name in enumerate(selected_students)}

print(f"Number of students processed: {len(student_dict)}")
print(f"Sample student data: {list(student_dict.items())[:5]}")

def generate_course_selection(grade):
    courses = []
    try:
        if grade == 8:
            courses = [course for course, _ in grade_8_required]
        elif grade == 9:
            courses = [course for course, _ in grade_9_required]
            courses.append(
                random.choice([course for course, course_grade in language_courses if course_grade == grade]))
            courses.append(random.choice([course for course, course_grade in adst_courses if course_grade == grade]))
            courses.append(
                random.choice([course for course, course_grade in fine_arts_courses if course_grade == grade]))
        elif grade == 10:
            courses = [course for course, _ in grade_10_required]
            courses.append(
                random.choice([course for course, course_grade in language_courses if course_grade == grade]))
            adst_fine_arts_choice = random.choice(
                [course for course, course_grade in adst_courses + fine_arts_courses if course_grade == grade])
            courses.append(adst_fine_arts_choice)
        elif grade == 11:
            courses = [course for course, _ in grade_11_required]
            courses.extend(
                random.sample([course for course, course_grade in science_11_12 if course_grade == grade], 2))
            courses.append(random.choice([course for course, course_grade in adst_courses if course_grade == grade]))
            courses.append(
                random.choice([course for course, course_grade in fine_arts_courses if course_grade == grade]))
            # Add one more elective to reach 8 courses
            additional_elective = random.choice(
                [course for course, course_grade in language_courses + adst_courses + fine_arts_courses if
                 course_grade == grade])
            courses.append(additional_elective)
        elif grade == 12:
            courses = [course for course, _ in grade_12_required]
            courses.extend(
                random.sample([course for course, course_grade in science_11_12 if course_grade == grade], 2))
            courses.append(random.choice([course for course, _ in grade_12_electives]))
            adst_fine_arts_choices = random.sample(
                [course for course, course_grade in adst_courses + fine_arts_courses if course_grade == grade], 2)
            courses.extend(adst_fine_arts_choices)
        else:
            raise ValueError(f"Invalid grade: {grade}")

        if len(courses) != 8:
            raise ValueError(f"Incorrect number of courses generated for grade {grade}: {len(courses)}")
        return courses
    except Exception as e:
        print(f"Error generating courses for grade {grade}: {str(e)}")
        return []

def generate_preferences(grade):
    preferences = []
    try:
        if grade in [9, 10, 11, 12]:
            # Filter ADST and Fine Arts courses by the student's grade
            adst_choices = [course for course, course_grade in adst_courses if course_grade == grade]
            fine_arts_choices = [course for course, course_grade in fine_arts_courses if course_grade == grade]

            # Ensure there are enough courses to sample from
            adst_sample_size = min(3, len(adst_choices))
            fine_arts_sample_size = min(2, len(fine_arts_choices))

            preferences += random.sample(adst_choices, adst_sample_size)
            preferences += random.sample(fine_arts_choices, fine_arts_sample_size)

        return preferences[:5]  # Return top 5 preferences
    except Exception as e:
        print(f"Error generating preferences for grade {grade}: {str(e)}")
        return []

# Create students dictionary with course selections and preferences
students = {}
for name, number in student_dict.items():
    grade = random.choice([8, 9, 10, 11, 12])
    students[name] = {
        "number": number,
        "grade": grade,
        "courses": generate_course_selection(grade),
        "preferences": generate_preferences(grade)
    }

def export_to_excel(students, filename='studentCourses.xlsx'):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Student Courses"

    headers = ["Student Name", "Student Number", "Grade", "Courses", "Preferences"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    for row, (name, data) in enumerate(students.items(), start=2):
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=data['number'])
        ws.cell(row=row, column=3, value=data['grade'])
        ws.cell(row=row, column=4, value=', '.join(data['courses']))
        ws.cell(row=row, column=5, value=', '.join(data['preferences']))  # Changed line

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save()
    print(f"Data exported to {filename}")

export_to_excel(students)
