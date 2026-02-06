"""
Central, highly customizable configuration for the scheduler.
Change to adapt to different schools.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any

# -----------------------------------------------------------------------------
# Time structure
# -----------------------------------------------------------------------------
DEFAULT_PERIODS: List[str] = [
    "S1P1", "S1P2", "S1P3", "S1P4",
    "S2P1", "S2P2", "S2P3", "S2P4",
]
OFF_TIMETABLE_BLOCK: str = "OT"  # for courses that don't need a period (e.g. choir)

# -----------------------------------------------------------------------------
# Class size (all customizable)
# -----------------------------------------------------------------------------
DEFAULT_MIN_CLASS_SIZE: int = 5  # allow small sections so solver can place everyone
DEFAULT_IDEAL_CLASS_SIZE: int = 25
# Max = room_capacity + CAPACITY_SLACK (so +5 means allow 5 over room capacity if needed)
DEFAULT_CAPACITY_SLACK: int = 5
# Global fallback when room capacity is missing
DEFAULT_GLOBAL_MAX_CLASS_SIZE: int = 35
# (1) Cap: section size is never allowed above this, even if room + slack would be higher. (2) Fallback: when a course has no teacher with a room capacity in Excel, max section size becomes 35. So: “absolute max size” and “default when Excel has no room.”


# -----------------------------------------------------------------------------
# Teacher
# -----------------------------------------------------------------------------
DEFAULT_MAX_TEACHER_SECTIONS: int = 7  # 7 teaching blocks, 1 prep

# -----------------------------------------------------------------------------
# Courses that are not placed on the timetable (e.g. choir outside blocks)
# -----------------------------------------------------------------------------
DEFAULT_OFF_TIMETABLE_COURSES: List[str] = [
    "MUSIC 9: CONCERT CHOIR",
    "CHORAL MUSIC 10: CONCERT CHOIR",
    "CHORAL MUSIC 11: CONCERT CHOIR",
    "CHORAL MUSIC 12: CONCERT CHOIR",
]

# -----------------------------------------------------------------------------
# Solver
# -----------------------------------------------------------------------------
DEFAULT_SOLVER_TIME_SECONDS: float = 120.0
# Primary requests: soft goal (not hard). We maximize assignments; no strict "exactly 8" or "at least 6".
# Desired: as many students as possible get all 8 requested courses; up to N students may get fewer (report only).
DESIRED_MAX_STUDENTS_UNDER_8: int = 10  # reporting target, not a constraint
# Hard constraint: total assignments = num_students * this (set None to disable).
COURSES_PER_STUDENT_TARGET: Optional[int] = None
# Symmetry breaking (pack sections into earlier periods) can hurt solution quality; set True to enable.
SOLVER_SYMMETRY_BREAK_PER_COURSE: bool = False
# Parallel search workers (0 = auto).
SOLVER_NUM_WORKERS: int = 8
# All Excel outputs go into this directory (created if missing).
DEFAULT_OUTPUT_DIR: str = "output"

# -----------------------------------------------------------------------------
# Input column names (so different schools can use different Excel headers)
# -----------------------------------------------------------------------------
TEACHER_COLUMNS: Dict[str, str] = {
    "last_name": "Last Name",
    "first_name": "First Name",
    "courses": "Courses",
    "adst_rotation": "ADST Rotation",
    "fine_arts_rotation": "Fine Arts Rotation",
    "classes": "Classes",
    "room_capacity": "Room Capcity",  # note typo in original
}
STUDENT_COLUMNS: Dict[str, str] = {
    "name": "Student Name",
    "number": "Student Number",
    "grade": "Grade",
    "courses": "Courses",
    "preferences": "Preferences",
}

# -----------------------------------------------------------------------------
# Grade 8 rotations (dynamic: not all schools have these; N-of-M is configurable)
# -----------------------------------------------------------------------------
@dataclass
class RotationDef:
    """One rotation (e.g. ADST or Fine Arts for grade 8)."""
    id: str                          # e.g. "ADST", "FineArts"
    display_name: str                # e.g. "ADST Rotation", "Fine Arts Rotation"
    grade: int = 8
    # How many "slots" each student gets in this rotation (e.g. 2 = half semester each)
    num_slots_per_student: int = 2
    # How many options we choose from (e.g. 3 options, assign 2)
    num_options: int = 3
    # Generic course names for sections (no real course names needed; just for display)
    option_display_names: Optional[List[str]] = None  # e.g. ["ADST A", "ADST B", "ADST C"]
    # Teacher IDs who participate in this rotation (must have rotation flag in data)
    # If None, we infer from teacher data rotation flags.
    teacher_keys: Optional[List[str]] = None

    def __post_init__(self):
        if self.option_display_names is None:
            self.option_display_names = [f"{self.display_name} {i+1}" for i in range(self.num_options)]


DEFAULT_ROTATIONS: List[RotationDef] = [
    RotationDef(
        id="ADST",
        display_name="ADST Rotation",
        grade=8,
        num_slots_per_student=2,
        num_options=3,
        option_display_names=["ADST A", "ADST B", "ADST C"],
    ),
    RotationDef(
        id="FineArts",
        display_name="Fine Arts Rotation",
        grade=8,
        num_slots_per_student=2,
        num_options=3,
        option_display_names=["Fine Arts A", "Fine Arts B", "Fine Arts C"],
    ),
]


@dataclass
class SchedulerConfig:
    """Single config object; override any field to customize."""
    periods: List[str] = field(default_factory=lambda: list(DEFAULT_PERIODS))
    off_timetable_block: str = OFF_TIMETABLE_BLOCK
    min_class_size: int = DEFAULT_MIN_CLASS_SIZE
    ideal_class_size: int = DEFAULT_IDEAL_CLASS_SIZE
    capacity_slack: int = DEFAULT_CAPACITY_SLACK
    global_max_class_size: int = DEFAULT_GLOBAL_MAX_CLASS_SIZE
    max_teacher_sections: int = DEFAULT_MAX_TEACHER_SECTIONS
    off_timetable_courses: List[str] = field(default_factory=lambda: list(DEFAULT_OFF_TIMETABLE_COURSES))
    solver_time_seconds: float = DEFAULT_SOLVER_TIME_SECONDS
    desired_max_students_under_8: int = DESIRED_MAX_STUDENTS_UNDER_8
    teacher_columns: Dict[str, str] = field(default_factory=lambda: dict(TEACHER_COLUMNS))
    student_columns: Dict[str, str] = field(default_factory=lambda: dict(STUDENT_COLUMNS))
    rotations: List[RotationDef] = field(default_factory=lambda: list(DEFAULT_ROTATIONS))
    symmetry_break_per_course: bool = SOLVER_SYMMETRY_BREAK_PER_COURSE
    solver_num_workers: int = SOLVER_NUM_WORKERS
    output_dir: str = DEFAULT_OUTPUT_DIR
    courses_per_student_target: Optional[int] = COURSES_PER_STUDENT_TARGET

    def max_capacity_for_room(self, room_capacity: Optional[int]) -> int:
        if room_capacity is not None:
            return room_capacity + self.capacity_slack
        return self.global_max_class_size


_config: Optional[SchedulerConfig] = None


def get_config() -> SchedulerConfig:
    global _config
    if _config is None:
        _config = SchedulerConfig()
    return _config


def set_config(cfg: SchedulerConfig) -> None:
    global _config
    _config = cfg
