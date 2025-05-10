# -----------------------------------------------------------------------------
#  main.py  – v1 timetable generator (no Grade‑8 rotation yet)
#  (2025‑05‑09)
# -----------------------------------------------------------------------------
#  • Fix: replace unsupported abs() on linear expression with AddAbsEquality.
#  • Everything else unchanged from previous patch.
# -----------------------------------------------------------------------------

from ortools.sat.python import cp_model
from typing import Dict, Tuple, List, Any
import pandas as pd

from input import students, teachers

# -----------------------------------------------------------------------------
# CONSTANTS
# -----------------------------------------------------------------------------
PERIODS = ["S1P1", "S1P2", "S1P3", "S1P4", "S2P1", "S2P2", "S2P3", "S2P4"]
OFF_TIMETABLE_BLOCK = "OT"  # synthetic period for choir/out‑of‑timetable

GLOBAL_MAX_SIZE   = 30
GLOBAL_IDEAL_SIZE = 20
GLOBAL_MIN_SIZE   = 15

OFF_TIMETABLE_COURSES = [
    "MUSIC 9: CONCERT CHOIR",
    "CHORAL MUSIC 10: CONCERT CHOIR",
    "CHORAL MUSIC 11: CONCERT CHOIR",
    "CHORAL MUSIC 12: CONCERT CHOIR",
]

# -----------------------------------------------------------------------------
# Helper: class‑size caps
# -----------------------------------------------------------------------------

def course_cap(course: str, fallback_max: int = GLOBAL_MAX_SIZE) -> Tuple[int, int, int]:
    caps = [t["room_capacity"] for t in teachers.values() if course in t["can_teach"] and t["room_capacity"]]
    max_cap = min(caps) if caps else fallback_max
    ideal   = max_cap - 10
    min_cap = max(max_cap - 20, GLOBAL_MIN_SIZE)
    return min_cap, ideal, max_cap

# Pre‑compute qualified‑teacher map
QUALIFIED: Dict[str, List[str]] = {}
for tid, tdata in teachers.items():
    for c in tdata["can_teach"]:
        QUALIFIED.setdefault(c, []).append(tid)

# -----------------------------------------------------------------------------
# Build CP‑SAT model
# -----------------------------------------------------------------------------

def build_model():
    model = cp_model.CpModel()

    SA: Dict[Tuple[int, str, str], cp_model.IntVar] = {}
    TA: Dict[Tuple[str, str, str], cp_model.IntVar] = {}
    CA: Dict[Tuple[str, str], cp_model.IntVar] = {}

    # -------- course universe -------------------------------------------------
    all_courses: set[str] = set()
    for sdata in students.values():
        all_courses.update(sdata["requests"])
    all_courses.update(QUALIFIED.keys())

    # Student assignment vars
    for sid, sdata in students.items():
        for course in sdata["requests"]:
            periods = [OFF_TIMETABLE_BLOCK] if course in OFF_TIMETABLE_COURSES else PERIODS
            for p in periods:
                SA[(sid, course, p)] = model.NewBoolVar(f"SA_{sid}_{course}_{p}")

    # Teacher assignment vars
    for tid, tdata in teachers.items():
        for course in tdata["can_teach"]:
            periods = [OFF_TIMETABLE_BLOCK] if course in OFF_TIMETABLE_COURSES else PERIODS
            for p in periods:
                TA[(tid, course, p)] = model.NewBoolVar(f"TA_{tid}_{course}_{p}")

    # Class active vars
    for course in all_courses:
        periods = [OFF_TIMETABLE_BLOCK] if course in OFF_TIMETABLE_COURSES else PERIODS
        for p in periods:
            CA[(course, p)] = model.NewBoolVar(f"CA_{course}_{p}")

    # ----- Constraints --------------------------------------------------------
    # 1. Student‑course uniqueness
    for sid, sdata in students.items():
        for course in sdata["requests"]:
            per = [OFF_TIMETABLE_BLOCK] if course in OFF_TIMETABLE_COURSES else PERIODS
            model.Add(sum(SA[(sid, course, p)] for p in per) <= 1)

    # 2. No overlap per period
    for sid in students:
        for p in PERIODS:
            model.Add(sum(SA[(sid, c, p)] for c in students[sid]["requests"] if (sid, c, p) in SA) <= 1)

    # 3. Student load 6‑8
    for sid, sdata in students.items():
        total = sum(SA[(sid, c, p)] for c in sdata["requests"] for p in ([OFF_TIMETABLE_BLOCK] if c in OFF_TIMETABLE_COURSES else PERIODS))
        model.Add(total >= 6)
        model.Add(total <= 8)

    # 4. Link CA to participants
    for (course, p), active in CA.items():
        studs = [SA[(sid, course, p)] for sid in students if (sid, course, p) in SA]
        if studs:
            model.Add(sum(studs) >= 1).OnlyEnforceIf(active)
            model.Add(sum(studs) == 0).OnlyEnforceIf(active.Not())
        teas = [TA[(tid, course, p)] for tid in QUALIFIED.get(course, []) if (tid, course, p) in TA]
        if teas:
            model.Add(sum(teas) >= 1).OnlyEnforceIf(active)

    # 5. Size bounds
    size_vars: Dict[Tuple[str, str], cp_model.IntVar] = {}
    for (course, p), active in CA.items():
        mi, ideal, ma = course_cap(course)
        sz = model.NewIntVar(0, len(students), f"SZ_{course}_{p}")
        size_vars[(course, p)] = sz
        model.Add(sz == sum(SA[(sid, course, p)] for sid in students if (sid, course, p) in SA))
        model.Add(sz >= mi).OnlyEnforceIf(active)
        model.Add(sz <= ma).OnlyEnforceIf(active)
        model.Add(sz == 0).OnlyEnforceIf(active.Not())

    # 6. Teacher load and per‑period limit
    for tid, tdata in teachers.items():
        load = sum(TA[(tid, c, p)] for c in tdata["can_teach"] for p in ([OFF_TIMETABLE_BLOCK] if c in OFF_TIMETABLE_COURSES else PERIODS))
        model.Add(load <= tdata["max_sections"])
        for p in PERIODS:
            model.Add(sum(TA[(tid, c, p)] for c in tdata["can_teach"] if (tid, c, p) in TA) <= 1)

    # 7. Implications
    for (tid, course, p), tvar in TA.items():
        model.AddImplication(tvar, CA[(course, p)])
    for (sid, course, p), svar in SA.items():
        model.AddImplication(svar, CA[(course, p)])

    # ----- Objective ----------------------------------------------------------
    total_assigned = sum(SA.values())
    dev_vars = []
    for (course, p), sz in size_vars.items():
        _, ideal, _ = course_cap(course)
        dev = model.NewIntVar(0, GLOBAL_MAX_SIZE, f"DEV_{course}_{p}")
        model.AddAbsEquality(dev, sz - ideal)
        dev_vars.append(dev)
    model.Maximize(total_assigned * 1000 - sum(dev_vars))

    return model, SA, TA, CA

# -----------------------------------------------------------------------------
# Solve & export
# -----------------------------------------------------------------------------

def solve_and_export():
    model, SA, TA, CA = build_model()
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print("No schedule. Status:", solver.StatusName())
        return

    schedule: Dict[str, Dict[str, Dict[str, Any]]] = {p: {} for p in PERIODS + [OFF_TIMETABLE_BLOCK]}

    # Fill students
    for (sid, c, p), var in SA.items():
        if solver.Value(var):
            info = schedule[p].setdefault(c, {"students": [], "teachers": []})
            info["students"].append(students[sid]["name"])

    # Attach teachers
    for (tid, c, p), var in TA.items():
        if solver.Value(var):
            info = schedule[p].setdefault(c, {"students": [], "teachers": []})
            info["teachers"].append(teachers[tid]["name"])

    export_schedule(schedule)
    flag_students(schedule)
    print("Solver", solver.StatusName(), "| sections:", sum(len(v) for v in schedule.values()))

# -----------------------------------------------------------------------------
# Excel helpers
# -----------------------------------------------------------------------------

def export_schedule(schedule: Dict[str, Dict[str, Dict[str, Any]]]):
    rows = []
    for period, courses in schedule.items():
        for course, data in courses.items():
            rows.append({
                "Period": period,
                "Course": course,
                "Teachers": ", ".join(data["teachers"]),
                "Class Size": len(data["students"]),
                "Students": ", ".join(data["students"])
            })
    if rows:
        pd.DataFrame(rows).to_excel("timetable_output.xlsx", index=False)
        print("timetable_output.xlsx written (", len(rows), "sections )")
    else:
        print("No sections generated – timetable_output.xlsx not written")


def flag_students(schedule: Dict[str, Dict[str, Dict[str, Any]]]):
    counts = {sid: 0 for sid in students}
    for period in schedule.values():
        for sec in period.values():
            for name in sec["students"]:
                sid = next(k for k, v in students.items() if v["name"] == name)
                counts[sid] += 1
    flagged = [(students[sid]["name"], cnt) for sid, cnt in counts.items() if cnt < 8]
    if flagged:
        pd.DataFrame(flagged, columns=["Student", "Courses_Assigned"]).to_excel(
            "students_under_8_courses.xlsx", index=False)
        print("students_under_8_courses.xlsx written (", len(flagged), "students )")
    else:
        print("All students have 8 courses – no under‑loaded students")

# -----------------------------------------------------------------------------
if __name__ == "__main__":
    solve_and_export()
