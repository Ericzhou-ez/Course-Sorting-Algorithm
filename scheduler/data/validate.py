"""
Validate that demand (student requests) aligns with supply (teacher capacity).
Run this before building the solver; optionally refuse to solve if alignment fails.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple, Any

from scheduler.config import get_config


@dataclass
class AlignmentResult:
    """Result of demand-supply validation."""
    ok: bool
    messages: List[str] = field(default_factory=list)
    # Courses that students requested but no teacher can teach
    no_teacher: List[str] = field(default_factory=list)
    # For each course: (demand_count, min_sections_needed, total_teacher_slots_available)
    course_stats: Dict[str, Tuple[int, int, int]] = field(default_factory=dict)
    # Courses that are under-supplied (demand > we can possibly serve with capacity)
    under_supplied: List[str] = field(default_factory=list)
    # Courses approaching capacity limit (supply_slots < min_sections * threshold)
    approaching_limit: List[str] = field(default_factory=list)

    def summary(self) -> str:
        lines = self.messages.copy()
        if self.no_teacher:
            lines.append(f"Courses with no teacher: {', '.join(self.no_teacher)}")
        if self.under_supplied:
            lines.append(f"Courses under-supplied: {', '.join(self.under_supplied)}")
        if self.approaching_limit:
            lines.append(f"Courses approaching capacity limit: {', '.join(self.approaching_limit)}")
        return "\n".join(lines) if lines else "OK"

    def detailed_report(self, global_max_size: Optional[int] = None) -> str:
        """Detailed per-course breakdown of demand vs supply."""
        from scheduler.config import get_config
        cfg = get_config()
        max_size = global_max_size or cfg.global_max_class_size
        
        lines = ["\n=== COURSE STAFFING ANALYSIS ==="]
        if not self.course_stats:
            lines.append("No courses to analyze.")
            return "\n".join(lines)

        # Sort by severity: no_teacher > under_supplied > approaching_limit > ok
        def severity_key(course: str) -> Tuple[int, float]:
            if course in self.no_teacher:
                return (0, 0.0)  # Most severe
            if course in self.under_supplied:
                return (1, 0.0)
            if course in self.approaching_limit:
                return (2, 0.0)
            demand, min_sec, supply = self.course_stats[course]
            ratio = supply / min_sec if min_sec > 0 else float('inf')
            return (3, -ratio)  # Sort ok courses by supply/demand ratio (descending)

        sorted_courses = sorted(self.course_stats.keys(), key=severity_key)
        
        # Filter to only show courses with issues
        courses_with_issues = [
            c for c in sorted_courses
            if c in self.no_teacher or c in self.under_supplied or c in self.approaching_limit
        ]
        
        if not courses_with_issues:
            lines.append("All courses have adequate staffing.")
            lines.append("\n" + "=" * 40)
            return "\n".join(lines)

        for course in courses_with_issues:
            demand, min_sections, supply_slots = self.course_stats[course]
            max_capacity = supply_slots * max_size if supply_slots > 0 else 0
            ratio = supply_slots / min_sections if min_sections > 0 else 0.0

            status = []
            if course in self.no_teacher:
                status.append("NO TEACHER")
            elif course in self.under_supplied:
                status.append("UNDERSTAFFED")
            elif course in self.approaching_limit:
                status.append("APPROACHING LIMIT")
            else:
                status.append("✓ OK")

            lines.append(
                f"\n{', '.join(status)}: {course}"
            )
            lines.append(f"  Demand: {demand} students")
            lines.append(f"  Min sections needed (at max class size {max_size}): {min_sections}")
            lines.append(f"  Teacher section-slots available: {supply_slots}")
            lines.append(f"  Max capacity (if all slots used at {max_size}/section): {max_capacity}")
            lines.append(f"  Supply/Demand ratio: {ratio:.2f}x")
            if ratio < 1.0:
                lines.append(f"Cannot meet demand even at max capacity!")
            elif ratio < 1.2:
                lines.append(f"Very tight - may struggle to assign all students")

        lines.append("\n" + "=" * 40)
        return "\n".join(lines)


def _qualified_teachers(teachers: Dict[str, Dict[str, Any]], course: str) -> List[str]:
    return [tid for tid, t in teachers.items() if course in (t.get("can_teach") or [])]


def _supply_for_course(teachers: Dict[str, Dict[str, Any]], course: str) -> int:
    """Total section-slots that can be assigned to this course (sum of max_sections of qualified teachers)."""
    total = 0
    for tid, t in teachers.items():
        if course in (t.get("can_teach") or []):
            total += t.get("max_sections", 7)
    return total


def validate_demand_supply(
    students: Dict[int, Dict[str, Any]],
    teachers: Dict[str, Dict[str, Any]],
    *,
    off_timetable: Optional[List[str]] = None,
    min_class_size: Optional[int] = None,
    global_max_size: Optional[int] = None,
    approaching_limit_threshold: float = 1.2,
) -> AlignmentResult:
    """
    Check that every requested course has at least one teacher, and that total
    section capacity can theoretically meet demand (at max class size).
    
    Args:
        approaching_limit_threshold: Courses with supply_slots < min_sections * this are flagged
            as "approaching limit" (default 1.2 = 20% buffer).
    """
    cfg = get_config()
    off_timetable = off_timetable or cfg.off_timetable_courses
    min_class_size = min_class_size or cfg.min_class_size
    global_max_size = global_max_size or cfg.global_max_class_size

    # Demand per course (only for courses that go on the timetable)
    demand: Dict[str, int] = {}
    for s in students.values():
        for c in (s.get("requests") or []):
            if c in off_timetable:
                continue
            demand[c] = demand.get(c, 0) + 1

    messages: List[str] = []
    no_teacher: List[str] = []
    under_supplied: List[str] = []
    approaching_limit: List[str] = []
    course_stats: Dict[str, Tuple[int, int, int]] = {}

    for course, count in demand.items():
        qualified = _qualified_teachers(teachers, course)
        supply_slots = _supply_for_course(teachers, course) if qualified else 0

        # Min sections needed if we fill each section to global_max_size
        min_sections = (count + global_max_size - 1) // global_max_size if global_max_size else count
        if not qualified:
            no_teacher.append(course)
            course_stats[course] = (count, min_sections, 0)
            continue

        course_stats[course] = (count, min_sections, supply_slots)
        if supply_slots < min_sections:
            under_supplied.append(course)
        elif supply_slots < min_sections * approaching_limit_threshold:
            approaching_limit.append(course)

    if no_teacher:
        messages.append(f"ALIGNMENT FAIL: {len(no_teacher)} course(s) have no qualified teacher.")
    if under_supplied:
        messages.append(f"ALIGNMENT FAIL: {len(under_supplied)} course(s) are UNDERSTAFFED (demand exceeds max capacity).")
    if approaching_limit:
        messages.append(f"ALIGNMENT WARN: {len(approaching_limit)} course(s) are approaching capacity limit.")
    if not no_teacher and not under_supplied and not approaching_limit:
        messages.append("Demand-supply alignment OK (every requested course has supply and enough slots).")

    ok = len(no_teacher) == 0 and len(under_supplied) == 0
    return AlignmentResult(
        ok=ok,
        messages=messages,
        no_teacher=no_teacher,
        course_stats=course_stats,
        under_supplied=under_supplied,
        approaching_limit=approaching_limit,
    )
