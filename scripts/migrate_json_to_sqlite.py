"""Import existing SpendSight JSON data into a SQLite database."""

from __future__ import annotations

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from spendsight_store import migrate_legacy_json  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description="Migrate SpendSight JSON files into SQLite.")
    parser.add_argument("--workspace", default=ROOT, help="SpendSight workspace directory.")
    parser.add_argument("--db", default=os.path.join(ROOT, "spendsight.db"), help="SQLite database path.")
    args = parser.parse_args()

    migrate_legacy_json(args.workspace, args.db)
    print(f"Migrated JSON data from {args.workspace} to {args.db}")


if __name__ == "__main__":
    main()
