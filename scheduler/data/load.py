"""
Load and normalize teacher and student data from Excel.
Data alignment: normalize course names so demand and supply use the same keys.
"""

import re
import pandas as pd
from typing import Dict, Any, List, Optional, Set

from scheduler.config import get_config


def _normalize_course_token(s: str) -> str:
    s = s.strip()
    # Collapse multiple spaces
    s = re.sub(r"\s+", " ", s)
    return s


def split_courses_cell(cell: Any) -> List[str]:
    """
    Split a cell that may contain course names separated by comma, period, or newline.
    Drops empty and known non-course tokens (e.g. 'Fine_Arts_rotation').
    """
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return []
    s = str(cell).strip()
    if not s:
        return []
    # Split on comma, period, or newline (some Excel cells use period by mistake)
    parts = re.split(r"[,.\n]+", s)
    # Known tokens that are not course names (rotation column names, typos)
    skip_tokens = {"fine_arts_rotation", "adst rotation", "fine arts rotation", ""}
    out = []
    for p in parts:
        t = _normalize_course_token(p)
        if not t or t.lower() in skip_tokens:
            continue
        out.append(t)
    return out


def _split_simple(cell: Any, sep: str = ",") -> List[str]:
    """Simple comma split for columns that don't mix periods (e.g. Preferences)."""
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return []
    s = str(cell).strip()
    if not s:
        return []
    return [_normalize_course_token(c) for c in s.split(sep) if _normalize_course_token(c)]


def _yes(val: Any) -> bool:
    return isinstance(val, str) and val.strip().lower().startswith("y")


def load_teachers(
    path: str,
    *,
    default_sections: Optional[int] = None,
    columns: Optional[Dict[str, str]] = None,
) -> Dict[str, Dict[str, Any]]:
    """
    Load teacher mapping from Excel.
    Uses split_courses_cell for Courses so that "CHORAL MUSIC 12. Fine_Arts_rotation" is parsed correctly.
    """
    cfg = get_config()
    col = columns or cfg.teacher_columns
    default_sections = default_sections if default_sections is not None else cfg.max_teacher_sections

    df = pd.read_excel(path)
    out: Dict[str, Dict[str, Any]] = {}
    periods = cfg.periods

    for _, r in df.iterrows():
        last = str(r.get(col["last_name"], "")).strip()
        first = str(r.get(col["first_name"], "")).strip()
        if not last and not first:
            continue
        key = f"{last}_{first[0]}" if first else last
        courses_raw = r.get(col["courses"], "")
        can_teach = split_courses_cell(courses_raw)
        classes_val = r.get(col["classes"], default_sections)
        max_sections = int(classes_val) if classes_val is not None and not pd.isna(classes_val) else default_sections
        rc = r.get(col["room_capacity"])
        room_capacity = None if rc is None or pd.isna(rc) else int(rc)

        rotations = {
            "ADST": _yes(r.get(col["adst_rotation"], "")),
            "FineArts": _yes(r.get(col["fine_arts_rotation"], "")),
        }
        cfg = get_config()
        extra = [rot.display_name for rot in (cfg.rotations or []) if rotations.get(rot.id, False)]
        can_teach = list(can_teach) + [x for x in extra if x not in can_teach]
        out[key] = {
            "name": f"{first} {last}".strip(),
            "can_teach": can_teach,
            "max_sections": max_sections,
            "room_capacity": room_capacity,
            "rotations": rotations,
            "availability": {p: True for p in periods},
        }
    return out


def load_students(
    path: str,
    *,
    columns: Optional[Dict[str, str]] = None,
) -> Dict[int, Dict[str, Any]]:
    """Load student course requests from Excel. Uses same course normalization as teachers."""
    cfg = get_config()
    col = columns or cfg.student_columns

    df = pd.read_excel(path)
    out: Dict[int, Dict[str, Any]] = {}

    for _, r in df.iterrows():
        num = r.get(col["number"])
        if num is None or pd.isna(num):
            continue
        number = int(num)
        name = str(r.get(col["name"], "")).strip()
        grade = int(r.get(col["grade"], 9))
        courses_raw = r.get(col["courses"], "")
        requests = split_courses_cell(courses_raw)
        prefs_raw = r.get(col["preferences"], "")
        preferences = _split_simple(prefs_raw)

        out[number] = {
            "name": name,
            "grade": grade,
            "requests": requests,
            "preferences": preferences,
        }
    return out


def course_universe(students: Dict[int, Dict[str, Any]], teachers: Dict[str, Dict[str, Any]]) -> Set[str]:
    """All course names that appear in either demand or supply."""
    out: Set[str] = set()
    for s in students.values():
        out.update(s.get("requests") or [])
    for t in teachers.values():
        out.update(t.get("can_teach") or [])
    return out
