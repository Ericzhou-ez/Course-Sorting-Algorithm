import pandas as pd

def analyze_student_enrollments(file_path):
    df = pd.read_excel(file_path)

    # Initialize counters
    course_counts = {}
    grade_counts = {8: 0, 9: 0, 10: 0, 11: 0, 12: 0}

    for index, row in df.iterrows():
        grade = row['Grade']
        courses = row['Courses'].split(', ')

        if grade in grade_counts:
            grade_counts[grade] += 1

        for course in courses:
            if course in course_counts:
                course_counts[course] += 1
            else:
                course_counts[course] = 1

    sorted_courses = sorted(course_counts.items(), key=lambda x: x[1], reverse=True)
    print("Number of students in each grade:")
    for grade, count in grade_counts.items():
        print(f"Grade {grade}: {count} students")

    print("\nNumber of students in each course (sorted):")
    for course, count in sorted_courses:
        print(f"{course}: {count} students")


# File path to sutdent's course preferences
file_path = '...path/to/your/Student Course File'
analyze_student_enrollments(file_path)
