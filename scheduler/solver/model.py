"""
Section-based CP-SAT model for timetable scheduling.
- No 6-8 hard rule; maximize assigned course-periods.
- One section = (course, period) with at most one teacher; size variable (0 = section closed).
- Redundant constraints and symmetry breaking to improve propagation.
"""

from typing import Dict, List, Set, Tuple, Any, Optional

from ortools.sat.python import cp_model

from scheduler.config import get_config


def _course_cap(
    course: str,
    teachers: Dict[str, Dict[str, Any]],
    cfg: Any,
) -> Tuple[int, int, int]:
    """(min, ideal, max) for a course. Max = min(room capacities) + capacity_slack."""
    caps = [
        t["room_capacity"]
        for t in teachers.values()
        if course in (t.get("can_teach") or []) and t.get("room_capacity")
    ]
    room_min = min(caps) if caps else cfg.global_max_class_size
    max_cap = min(room_min + cfg.capacity_slack, cfg.global_max_class_size)
    ideal = cfg.ideal_class_size
    min_cap = cfg.min_class_size
    return min_cap, ideal, max_cap


def build_model(
    students: Dict[int, Dict[str, Any]],
    teachers: Dict[str, Dict[str, Any]],
    *,
    off_timetable_courses: Optional[List[str]] = None,
) -> Tuple[cp_model.CpModel, Dict, Dict, Dict]:
    """
    Build CP-SAT model. Returns (model, SA, TA, size_vars).
    SA[(sid, course, period)] = 1 if student sid takes course in period.
    TA[(tid, course, period)] = 1 if teacher tid teaches course in period.
    size_vars[(course, period)] = enrollment in that section (0 if section not run).
    """
    cfg = get_config()
    off = set(off_timetable_courses or cfg.off_timetable_courses)
    periods = cfg.periods
    n_students = len(students)

    # Courses that appear in demand and have at least one teacher (on-timetable only)
    demand_courses: Set[str] = set()
    for s in students.values():
        for c in (s.get("requests") or []):
            if c not in off:
                demand_courses.add(c)
    qualified: Dict[str, List[str]] = {}
    for tid, t in teachers.items():
        for c in (t.get("can_teach") or []):
            if c in off:
                continue
            qualified.setdefault(c, []).append(tid)
    # Only model courses that are requested and have supply
    courses = [c for c in demand_courses if c in qualified]

    model = cp_model.CpModel()

    # --- Decision variables ---
    # Student assignment: sid takes course c in period p
    SA: Dict[Tuple[int, str, str], cp_model.IntVar] = {}
    for sid, sdata in students.items():
        for c in (sdata.get("requests") or []):
            if c not in courses:
                continue
            for p in periods:
                SA[(sid, c, p)] = model.NewBoolVar(f"SA_{sid}_{c}_{p}")

    # Teacher assignment: tid teaches course c in period p
    TA: Dict[Tuple[str, str, str], cp_model.IntVar] = {}
    for tid, t in teachers.items():
        for c in (t.get("can_teach") or []):
            if c not in courses:
                continue
            for p in periods:
                TA[(tid, c, p)] = model.NewBoolVar(f"TA_{tid}_{c}_{p}")

    # Section size: enrollment in (course, period). 0 means section not run.
    size_vars: Dict[Tuple[str, str], cp_model.IntVar] = {}
    section_active: Dict[Tuple[str, str], cp_model.IntVar] = {}
    for c in courses:
        for p in periods:
            _, _, max_cap = _course_cap(c, teachers, cfg)
            size_vars[(c, p)] = model.NewIntVar(0, max_cap, f"SZ_{c}_{p}")
            section_active[(c, p)] = model.NewBoolVar(f"active_{c}_{p}")

    # --- Hard constraints ---

    # 1. Student: at most one period per course
    for sid, sdata in students.items():
        for c in (sdata.get("requests") or []):
            if c not in courses:
                continue
            model.Add(sum(SA.get((sid, c, p), 0) for p in periods) <= 1)

    # 2. Student: at most one course per period (no clash)
    for sid in students:
        for p in periods:
            model.Add(
                sum(SA[(sid, c, p)] for c in (students[sid].get("requests") or []) if (sid, c, p) in SA) <= 1
            )

    # 3. Section size = number of students in (course, period)
    for (c, p), sz_var in size_vars.items():
        model.Add(
            sz_var == sum(SA.get((sid, c, p), 0) for sid in students if (sid, c, p) in SA)
        )

    # 4. Each (course, period) has at most one teacher
    for c in courses:
        for p in periods:
            model.Add(sum(TA.get((tid, c, p), 0) for tid in qualified.get(c, [])) <= 1)

    # 5. Link section_active to teacher_sum; enforce size bounds when active
    for c in courses:
        for p in periods:
            sz = size_vars[(c, p)]
            act = section_active[(c, p)]
            teacher_sum = sum(TA.get((tid, c, p), 0) for tid in qualified.get(c, []))
            # section_active == 1 iff teacher_sum >= 1
            model.Add(teacher_sum >= 1).OnlyEnforceIf(act)
            model.Add(teacher_sum <= 0).OnlyEnforceIf(act.Not())
            min_cap, _, max_cap = _course_cap(c, teachers, cfg)
            model.Add(sz >= min_cap).OnlyEnforceIf(act)
            model.Add(sz <= max_cap).OnlyEnforceIf(act)
            model.Add(sz == 0).OnlyEnforceIf(act.Not())
            model.Add(teacher_sum <= 1)

    # 6. Teacher load: total sections per teacher <= max_sections
    for tid, t in teachers.items():
        load = sum(
            TA.get((tid, c, p), 0)
            for c in (t.get("can_teach") or [])
            for p in periods
            if (tid, c, p) in TA
        )
        model.Add(load <= t.get("max_sections", cfg.max_teacher_sections))

    # 7. Teacher: at most one class per period
    for tid in teachers:
        for p in periods:
            model.Add(
                sum(TA.get((tid, c, p), 0) for c in (teachers[tid].get("can_teach") or []) if (tid, c, p) in TA)
                <= 1
            )

    # 8. Student can only be in (course, period) if that section is open (size > 0)
    for (sid, c, p), var in SA.items():
        model.Add(size_vars[(c, p)] >= 1).OnlyEnforceIf(var)

    # 9. Hard: total assignments = n_students * courses_per_student (if set)
    target = getattr(cfg, "courses_per_student_target", None)
    if target is not None:
        model.Add(sum(SA.values()) == n_students * target)

    # --- Optional symmetry breaking: for each course, prefer "earlier" periods first ---
    if get_config().symmetry_break_per_course:
        for c in courses:
            for i in range(len(periods) - 1):
                model.Add(size_vars[(c, periods[i])] >= size_vars[(c, periods[i + 1])])

    # --- Objective: maximize assignments, then minimize deviation from ideal size ---
    total_assigned = sum(SA.values())
    dev_vars = []
    for (c, p), sz in size_vars.items():
        _, ideal, _ = _course_cap(c, teachers, cfg)
        dev = model.NewIntVar(0, cfg.global_max_class_size, f"DEV_{c}_{p}")
        model.AddAbsEquality(dev, sz - ideal)
        dev_vars.append(dev)
    # Prioritize assignments; secondary minimize size deviation
    model.Maximize(total_assigned * 10000 - sum(dev_vars))

    return model, SA, TA, size_vars
