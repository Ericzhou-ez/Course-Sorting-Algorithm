import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment

from config import (grade_8_required, grade_9_required, grade_10_required, grade_11_required, grade_12_required, language_courses, adst_courses, fine_arts_courses, science_11_12, grade_12_electives
)
offTimeTableMusicCourses = ["MUSIC 9: CONCERT CHOIR",
                            "CHORAL MUSIC 10: CONCERT CHOIR",
                            "CHORAL MUSIC 11: CONCERT CHOIR",
                            "CHORAL MUSIC 12: CONCERT CHOIR"]

students = {}
def populate_student_dict(student_input):
    try: 
        for index, row in student_input.iterrows():
        student_name = row['Student Name']
        student_number = row['Student Number']
        grade = row['Grade']

        courses = row['Courses'].split(', ')
        preferences = row['Preferences']

        if pd.isna(preferences):
            preferences = []
        else:
            preferences = preferences.split(', ')

        # preference values for future use
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
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")
        return
                 
file_path_students = 'exampleInput/studentCourses.xlsx'
df_students = pd.read_excel(file_path_students)
populate_student_dict(df_students)

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
        df_teachers = pd.read_excel(file_path)
        teachers = {}

        for _, row in df_teachers.iterrows():
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
        return

file_path_teachers = 'exampleInput/TeacherCourseMapping.xlsx'
teachers = read_teacher_data(file_path_teachers)

print(teachers, students)