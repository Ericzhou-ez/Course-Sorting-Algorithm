"""
Export schedule: school schedule (teacher/room/students) and student schedules.
"""

from typing import Dict, Any, List, Optional
import pandas as pd

from scheduler.config import get_config


def export_school_schedule(
    schedule: Dict[str, Dict[str, Dict[str, Any]]],
    teachers: Dict[str, Dict[str, Any]],
    output_path: str = "school_schedule.xlsx",
    *,
    periods: Optional[List[str]] = None,
) -> None:
    """
    Write school schedule to Excel: Period, Course, Teacher, Room, Students.
    Room is derived from teacher (using teacher name as room identifier).
    """
    cfg = get_config()
    periods = periods or cfg.periods

    rows = []
    for p in periods:
        for course, data in (schedule.get(p) or {}).items():
            teacher_names = data.get("teachers") or []
            teacher_ids = data.get("teacher_ids") or []
            students = data.get("students") or []
            
            # Use first teacher's name as room identifier (or "TBD" if no teacher)
            room = teacher_names[0] if teacher_names else "TBD"
            teacher = ", ".join(teacher_names) if teacher_names else "TBD"
            
            rows.append({
                "Period": p,
                "Course": course,
                "Teacher": teacher,
                "Room": room,
                "Students": ", ".join(students),
                "Class Size": len(students),
            })
    
    if rows:
        pd.DataFrame(rows).to_excel(output_path, index=False)
        print(f"Wrote {output_path} ({len(rows)} sections)")
    else:
        print(f"No sections; {output_path} not written.")


def export_student_schedules(
    schedule: Dict[str, Dict[str, Dict[str, Any]]],
    students: Dict[int, Dict[str, Any]],
    output_path: str = "student_schedules.xlsx",
    *,
    periods: Optional[List[str]] = None,
) -> None:
    """
    Write all student schedules to Excel: Student Name, Student Number, Grade, then one column per period with their course.
    """
    cfg = get_config()
    periods = periods or cfg.periods

    # Build student -> period -> course mapping
    student_courses: Dict[int, Dict[str, str]] = {}
    for sid in students:
        student_courses[sid] = {p: "" for p in periods}

    for period, courses in schedule.items():
        for course, data in courses.items():
            for student_name in (data.get("students") or []):
                # Find student ID by name
                sid = next((k for k, v in students.items() if v["name"] == student_name), None)
                if sid is not None:
                    student_courses[sid][period] = course

    # Build rows
    rows = []
    for sid in sorted(students.keys()):
        sdata = students[sid]
        row = {
            "Student Name": sdata["name"],
            "Student Number": sid,
            "Grade": sdata.get("grade", ""),
        }
        for p in periods:
            row[p] = student_courses[sid].get(p, "")
        rows.append(row)

    if rows:
        pd.DataFrame(rows).to_excel(output_path, index=False)
        print(f"Wrote {output_path} ({len(rows)} students)")
    else:
        print(f"No students; {output_path} not written.")
