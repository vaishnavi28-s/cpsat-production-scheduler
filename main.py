"""
Print Job Optimizer — entry point.

Usage:
    # Demo mode (uses sample CSV data — no credentials needed)
    python main.py

    # Snowflake mode (requires .env with SF_* credentials)
    python main.py --snowflake

    # Custom CSV path
    python main.py --csv path/to/jobs.csv

    # Custom output path
    python main.py --output results.xlsx
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from src.data_loader import load_from_csv, load_from_snowflake
from src.optimizer import run_optimizer
from src.export import export_to_excel

DEFAULT_CSV = os.path.join(os.path.dirname(__file__), "data", "sample_jobs.csv")
DEFAULT_OUTPUT = "combo_output.xlsx"


def main():
    parser = argparse.ArgumentParser(description="Print Job Combo Optimizer")
    parser.add_argument("--snowflake", action="store_true", help="Load from Snowflake instead of CSV")
    parser.add_argument("--csv", type=str, default=DEFAULT_CSV, help="Path to input CSV file")
    parser.add_argument("--output", type=str, default=DEFAULT_OUTPUT, help="Output Excel file path")
    parser.add_argument("--limit", type=int, default=5000, help="Max rows to fetch from Snowflake")
    args = parser.parse_args()

    print("=" * 60)
    print("   Print Job Combo Optimizer")
    print("=" * 60)

    if args.snowflake:
        print("Mode: Snowflake")
        jobs = load_from_snowflake(limit=args.limit)
    else:
        print(f"Mode: CSV ({args.csv})")
        jobs = load_from_csv(args.csv)

    print(f"\nLoaded {len(jobs)} print jobs")
    print("Running CP-SAT optimizer...")

    total_runs = run_optimizer(jobs)

    total_combos = sum(
        sum(1 for ids in runs.values() if len(ids) > 1)
        for runs in total_runs.values()
    )
    total_singles = sum(
        sum(1 for ids in runs.values() if len(ids) == 1)
        for runs in total_runs.values()
    )

    print(f"\nResults: {total_combos} combos | {total_singles} singles")
    export_to_excel(total_runs, jobs, path=args.output)
    print(f"\nDone! Output saved to: {args.output}")


if __name__ == "__main__":
    main()
