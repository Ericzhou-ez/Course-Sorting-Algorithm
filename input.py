import pandas as pd
from typing import Dict, Any, List

PERIODS = ["S1P1","S1P2","S1P3","S1P4","S2P1","S2P2","S2P3","S2P4"]

# ---------- helpers ----------
def _yes(val) -> bool:
    return isinstance(val, str) and val.strip().lower().startswith("y")

def _split(cell: str) -> List[str]:
    if not isinstance(cell, str):
        return []
    return [c.strip() for c in cell.split(",") if c.strip()]

# ---------- teachers ----------
def load_teachers(path: str, default_sections: int = 7) -> Dict[str, Dict[str, Any]]:
    df = pd.read_excel(path)
    out: Dict[str, Dict[str, Any]] = {}
    for _, r in df.iterrows():
        last, first = str(r["Last Name"]).strip(), str(r["First Name"]).strip()
        key         = f"{last}_{first[0]}"          # Smith_J
        out[key] = {
            "name":   first + " " + last,
            "can_teach":    _split(r.get("Courses", "")),
            "max_sections": int(r["Classes"]) if not pd.isna(r["Classes"]) else default_sections,
            "room_capacity": (None if pd.isna(r["Room Capcity"])
                              else int(r["Room Capcity"])),
            "rotations": {
                "ADST":     _yes(r.get("ADST Rotation", "")),
                "FineArts": _yes(r.get("Fine Arts Rotation", ""))
            },
            "availability": {p: True for p in PERIODS}
        }
    return out

# ---------- students ----------
def load_students(path: str) -> Dict[int, Dict[str, Any]]:
    df = pd.read_excel(path)
    out: Dict[int, Dict[str, Any]] = {}
    for _, r in df.iterrows():
        number = int(r["Student Number"])
        out[number] = {
            "name":   str(r["Student Name"]).strip(),
            "grade":  int(r["Grade"]),
            "requests": _split(r.get("Courses", "")),
            "preferences": _split(r.get("Preferences", "")),
        }
    return out

teachers = load_teachers("exampleInput/TeacherCourseMapping.xlsx")
students = load_students("exampleInput/studentCourses.xlsx")

if __name__ == "__main__":
    # print any one teacher / student to verify
    sample_teacher_key  = next(iter(teachers))
    sample_student_key  = next(iter(students))

    print("Sample teacher ➜", teachers[sample_teacher_key])
    print("Sample student ➜", students[sample_student_key])