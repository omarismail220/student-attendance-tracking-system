#!/usr/bin/env python3
"""
One-off migration: legacy ``data/groups.json`` + ``data/attendance.json`` → SQLite.

Usage (from project root, with virtualenv activated)::

    python migrate_json_to_sqlite.py

If ``data/attendance.db`` already contains groups, this script **does not** overwrite
them (same rules as server bootstrap). To force a clean import, move or delete
``data/attendance.db`` first, then run again.
"""

from __future__ import annotations

from pathlib import Path

from attendance_service import migrate_from_json_files, seed_default_groups_if_empty
from db import configure_engine, init_db, session_scope


def main() -> None:
    root = Path(__file__).resolve().parent
    configure_engine(root)
    init_db()
    with session_scope() as session:
        migrate_from_json_files(root, session)
        session.flush()
        seed_default_groups_if_empty(session)
    print("Done. SQLite database is under data/attendance.db (unless ATTENDANCE_DATABASE_URL is set).")


if __name__ == "__main__":
    main()
