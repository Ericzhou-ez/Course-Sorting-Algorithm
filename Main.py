#!/usr/bin/env python3
"""
Entry point: load data, align demand/supply, solve, export.
Run from repo root. Uses scheduler package (modular, configurable).
"""

import argparse
import os
import sys

from scheduler.config import get_config
from scheduler.data import load_and_validate
from scheduler.solver import solve
from scheduler.export import export_school_schedule, export_student_schedules
from scheduler.rotation import apply_rotations_to_schedule


def main():
    parser = argparse.ArgumentParser(description="High-school course scheduler (SAT-based)")
    parser.add_argument("--teachers", default="exampleInput/TeacherCourseMapping.xlsx", help="Teacher/course Excel path")
    parser.add_argument("--students", default="exampleInput/studentCourses.xlsx", help="Student courses Excel path")
    parser.add_argument("--out-dir", default=None, help="Directory for all output Excel files (default: output)")
    parser.add_argument("--no-require-alignment", action="store_true", help="Run even if demand/supply alignment fails")
    parser.add_argument("--time", type=float, default=None, help="Solver time limit (seconds)")
    args = parser.parse_args()

    cfg = get_config()
    time_limit = args.time if args.time is not None else cfg.solver_time_seconds

    print("Loading and validating data...")
    try:
        students, teachers, alignment = load_and_validate(
            args.teachers,
            args.students,
            require_alignment=not args.no_require_alignment,
        )
    except FileNotFoundError as e:
        print("ERROR: File not found.", e, file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print("ERROR:", e, file=sys.stderr)
        sys.exit(1)

    print(alignment.summary())
    print(alignment.detailed_report())
    print(f"\nStudents: {len(students)}, Teachers: {len(teachers)}")
    
    # Fail fast if understaffed courses exist
    if alignment.under_supplied or alignment.no_teacher:
        print("\nERROR: Cannot proceed with understaffed courses or courses with no teacher.")
        print("Fix the data before solving:")
        if alignment.no_teacher:
            print(f"  - Add teachers for: {', '.join(alignment.no_teacher)}")
        if alignment.under_supplied:
            print(f"  - Increase teacher capacity for: {', '.join(alignment.under_supplied)}")
        sys.exit(1)

    print("Solving...")
    schedule = solve(
        students,
        teachers,
        time_limit_seconds=time_limit,
    )

    if schedule is None:
        print("No feasible schedule found. Try relaxing constraints or check data.")
        sys.exit(1)

    out_dir = args.out_dir or cfg.output_dir
    os.makedirs(out_dir, exist_ok=True)
    print(f"Writing outputs to {out_dir}/...")
    export_school_schedule(
        schedule,
        teachers,
        output_path=os.path.join(out_dir, "school_schedule.xlsx")
    )
    export_student_schedules(
        schedule,
        students,
        output_path=os.path.join(out_dir, "student_schedules.xlsx")
    )

    # Optional: rotation option assignment for G8 (2-of-3 etc.)
    try:
        rotation_assignments = apply_rotations_to_schedule(schedule, students)
        if rotation_assignments:
            print(f"Rotation options assigned for {len(rotation_assignments)} students in rotation sections.")
    except Exception as e:
        print("Rotation assignment (optional):", e)

    print("Done.")


if __name__ == "__main__":
    main()
