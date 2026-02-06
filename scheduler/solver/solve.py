"""
Run the CP-SAT solver and return a schedule structure.
"""

from typing import Dict, List, Any, Optional

from ortools.sat.python import cp_model

from scheduler.config import get_config
from scheduler.solver.model import build_model


def solve(
    students: Dict[int, Dict[str, Any]],
    teachers: Dict[str, Dict[str, Any]],
    *,
    off_timetable_courses: Optional[List[str]] = None,
    time_limit_seconds: Optional[float] = None,
) -> Optional[Dict[str, Dict[str, Dict[str, Any]]]]:
    """
    Build model, solve, and return schedule.
    Schedule: period -> course -> {"students": [names], "teachers": [names]}.
    Returns None if status is not OPTIMAL or FEASIBLE.
    """
    cfg = get_config()
    off = off_timetable_courses or cfg.off_timetable_courses
    time_limit = time_limit_seconds if time_limit_seconds is not None else cfg.solver_time_seconds
    
    model, SA, TA, size_vars = build_model(students, teachers, off_timetable_courses=off)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    if getattr(cfg, "solver_num_workers", 0) > 0:
        solver.parameters.num_search_workers = cfg.solver_num_workers

    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None

    periods = cfg.periods
    schedule: Dict[str, Dict[str, Dict[str, Any]]] = {p: {} for p in periods}

    for (sid, c, p), var in SA.items():
        if solver.Value(var):
            info = schedule[p].setdefault(c, {"students": [], "teachers": [], "teacher_ids": []})
            info["students"].append(students[sid]["name"])

    for (tid, c, p), var in TA.items():
        if solver.Value(var):
            info = schedule[p].setdefault(c, {"students": [], "teachers": [], "teacher_ids": []})
            info["teachers"].append(teachers[tid]["name"])
            info["teacher_ids"].append(tid)

    return schedule
