"""
Grade 8 (or other) rotations: assign students in a rotation section to N of M options.
Dynamic: configurable per school (some have 2-of-3, some 4 rotations, etc.).
"""

import random
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

from scheduler.config import get_config, RotationDef


@dataclass
class RotationAssignment:
    """Per-student assignment within a rotation section."""
    rotation_id: str
    period: str
    option_names: List[str]  # e.g. ["ADST A", "ADST B"]


def assign_rotation_options(
    section_students: List[str],
    rotation: RotationDef,
    period: str,
    *,
    rng: Optional[random.Random] = None,
) -> Dict[str, RotationAssignment]:
    """
    Given a list of student names in one section of a rotation, assign each student
    to num_slots_per_student of the num_options (e.g. 2 of 3). Returns
    student_name -> RotationAssignment.
    """
    rng = rng or random.Random()
    n_pick = rotation.num_slots_per_student
    n_options = rotation.num_options
    names = rotation.option_display_names or [f"{rotation.display_name} {i+1}" for i in range(n_options)]

    out: Dict[str, RotationAssignment] = {}
    for student in section_students:
        chosen = rng.sample(names, min(n_pick, n_options))
        out[student] = RotationAssignment(
            rotation_id=rotation.id,
            period=period,
            option_names=chosen,
        )
    return out


def apply_rotations_to_schedule(
    schedule: Dict[str, Dict[str, Dict[str, Any]]],
    students: Dict[int, Dict[str, Any]],
    *,
    rotations: Optional[List[RotationDef]] = None,
    rng: Optional[random.Random] = None,
) -> Dict[str, Dict[str, RotationAssignment]]:
    """
    For each rotation section in the schedule, assign each student to N-of-M options.
    Returns student_name -> { rotation_id: RotationAssignment }.
    """
    cfg = get_config()
    rotations = rotations or cfg.rotations
    rng = rng or random.Random()

    # Map rotation display_name -> RotationDef
    by_display: Dict[str, RotationDef] = {r.display_name: r for r in rotations}

    result: Dict[str, Dict[str, RotationAssignment]] = {}

    for period, courses in schedule.items():
        for course, data in courses.items():
            if course not in by_display:
                continue
            rotation = by_display[course]
            student_names = data.get("students") or []
            for name in student_names:
                if name not in result:
                    result[name] = {}
                result[name][rotation.id] = assign_rotation_options(
                    [name], rotation, period, rng=rng
                )[name]

    return result
