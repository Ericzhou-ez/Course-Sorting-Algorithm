# High-school course scheduling: modular, configurable SAT-based solver.
# Use scheduler.config to customize; scheduler.data to load/validate; scheduler.solver to solve.

from scheduler.config import get_config
from scheduler.data import load_and_validate  # noqa: F401
from scheduler.solver import solve
from scheduler.export import export_school_schedule, export_student_schedules

__all__ = [
    "get_config",
    "load_and_validate",
    "solve",
    "export_school_schedule",
    "export_student_schedules",
]
