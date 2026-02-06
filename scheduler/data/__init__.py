from scheduler.data.load import load_teachers, load_students
from scheduler.data.validate import validate_demand_supply, AlignmentResult


def load_and_validate(
    teachers_path: str,
    students_path: str,
    *,
    require_alignment: bool = True,
):
    """
    Load teachers and students, then validate demand vs supply.
    Returns (students, teachers, alignment_result).
    If require_alignment is True and alignment fails, raises ValueError.
    """
    teachers = load_teachers(teachers_path)
    students = load_students(students_path)
    alignment = validate_demand_supply(students, teachers)
    if require_alignment and not alignment.ok:
        error_msg = "Data alignment failed. Fix input data before solving.\n"
        error_msg += alignment.summary() + "\n"
        error_msg += alignment.detailed_report()
        raise ValueError(error_msg)
    return students, teachers, alignment


__all__ = [
    "load_teachers",
    "load_students",
    "load_and_validate",
    "validate_demand_supply",
    "AlignmentResult",
]
