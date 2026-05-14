"""
Business logic: status normalization, roster CRUD, attendance upserts, nested JSON for API,
and one-shot migration from legacy ``groups.json`` / ``attendance.json``.
"""

from __future__ import annotations

import json
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from sqlalchemy import and_, delete, func, select
from sqlalchemy.orm import Session, selectinload

from db import session_scope
from models import AttendanceRecord, Group, Student
from storage import DEFAULT_GROUPS, _data_dir, _read_json


def _parse_iso_date(s: str) -> Optional[date]:
    try:
        return date.fromisoformat(s.strip())
    except Exception:
        return None


def normalize_ui_status_to_db(raw: Any) -> Optional[str]:
    """
    Map frontend / legacy Arabic strings to canonical DB values.
    Returns None when the cell should be cleared (no DB row).
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    if s in ("present", "absent", "late"):
        return s
    if s in ("حاضر",):
        return "present"
    if s in ("غائب",):
        return "absent"
    if s in ("متاخر", "متأخر"):
        return "late"
    return None


def db_status_to_ui(db_val: str) -> str:
    return {"present": "حاضر", "absent": "غائب", "late": "متاخر"}.get(db_val, db_val)


def ensure_group_exists(session: Session, group_id: str) -> bool:
    return session.get(Group, group_id) is not None


def list_groups(session: Session) -> List[Dict[str, Any]]:
    rows = session.scalars(select(Group).order_by(Group.id)).all()
    out = []
    for g in rows:
        cnt = session.scalar(
            select(func.count()).select_from(Student).where(Student.group_id == g.id)
        )
        out.append({"id": g.id, "label": g.label, "time": g.time_slot, "count": int(cnt or 0)})
    return out


def get_group_meta_and_students(session: Session, group_id: str) -> Optional[Tuple[Group, List[str]]]:
    g = session.get(Group, group_id)
    if not g:
        return None
    names = [
        s.full_name
        for s in session.scalars(
            select(Student).where(Student.group_id == group_id).order_by(Student.sort_order)
        ).all()
    ]
    return g, names


def replace_group_students(session: Session, group_id: str, names: List[str]) -> None:
    """Full roster replace (Excel import). Removes old students and CASCADE attendance."""
    session.execute(delete(Student).where(Student.group_id == group_id))
    for i, nm in enumerate(names):
        session.add(Student(group_id=group_id, full_name=nm, sort_order=i))


def seed_default_groups_if_empty(session: Session) -> None:
    """Insert ``DEFAULT_GROUPS`` only when the ``groups`` table has zero rows (post-flush)."""
    session.flush()
    if session.scalar(select(Group.id).limit(1)):
        return
    for gid, meta in DEFAULT_GROUPS.items():
        session.add(
            Group(id=gid, label=meta["label"], time_slot=meta["time"]),
        )
        for i, nm in enumerate(meta.get("students", [])):
            session.add(Student(group_id=gid, full_name=nm, sort_order=i))


def migrate_from_json_files(app_root: Path, session: Session) -> bool:
    """
    If ``groups.json`` exists and DB has no groups, import groups + students.
    If ``attendance.json`` exists and the attendance table is empty, import rows.
    Returns True if any JSON file was consumed.
    """
    data = _data_dir(app_root)
    gpath = data / "groups.json"
    apath = data / "attendance.json"
    did = False

    if gpath.is_file() and not session.scalar(select(Group.id).limit(1)):
        raw = _read_json(gpath, None)
        if isinstance(raw, dict) and raw:
            for gid, meta in raw.items():
                if not isinstance(meta, dict):
                    continue
                session.add(
                    Group(
                        id=str(gid),
                        label=str(meta.get("label", gid)),
                        time_slot=str(meta.get("time", "")),
                    )
                )
                for i, nm in enumerate(meta.get("students") or []):
                    if isinstance(nm, str) and nm.strip():
                        session.add(
                            Student(group_id=str(gid), full_name=nm.strip(), sort_order=i)
                        )
            session.flush()
            did = True

    att_count = session.scalar(select(func.count()).select_from(AttendanceRecord)) or 0
    if apath.is_file() and att_count == 0:
        att = _read_json(apath, {})
        if isinstance(att, dict) and att:
            # Build name -> student_id per group
            for gid, dates_block in att.items():
                if not isinstance(dates_block, dict):
                    continue
                studs = session.scalars(
                    select(Student).where(Student.group_id == str(gid))
                ).all()
                by_name = {s.full_name: s for s in studs}
                for dstr, rec in dates_block.items():
                    if not isinstance(rec, dict):
                        continue
                    d = _parse_iso_date(str(dstr))
                    if not d:
                        continue
                    for name, raw in rec.items():
                        st = by_name.get(name)
                        if not st:
                            continue
                        db_s = normalize_ui_status_to_db(raw)
                        if db_s is None:
                            continue
                        existing = session.scalar(
                            select(AttendanceRecord).where(
                                and_(
                                    AttendanceRecord.group_id == str(gid),
                                    AttendanceRecord.student_id == st.id,
                                    AttendanceRecord.attendance_date == d,
                                )
                            )
                        )
                        if existing:
                            existing.status = db_s
                            existing.updated_at = datetime.utcnow()
                        else:
                            session.add(
                                AttendanceRecord(
                                    group_id=str(gid),
                                    student_id=st.id,
                                    attendance_date=d,
                                    status=db_s,
                                )
                            )
            did = True
    return did


def upsert_attendance_for_date(
    session: Session, group_id: str, attendance_date: date, records: Dict[str, Any]
) -> None:
    """
    For each student in ``records`` keyed by display name: upsert or delete row for that day.
    Unknown names are ignored. Empty / unmapped values delete the row for that day.
    """
    studs = session.scalars(
        select(Student).where(Student.group_id == group_id).order_by(Student.sort_order)
    ).all()
    by_name = {s.full_name: s for s in studs}

    for name, raw in records.items():
        st = by_name.get(name)
        if not st:
            continue
        db_s = normalize_ui_status_to_db(raw)
        row = session.scalar(
            select(AttendanceRecord).where(
                and_(
                    AttendanceRecord.group_id == group_id,
                    AttendanceRecord.student_id == st.id,
                    AttendanceRecord.attendance_date == attendance_date,
                )
            )
        )
        if db_s is None:
            if row:
                session.delete(row)
        elif row:
            row.status = db_s
            row.updated_at = datetime.utcnow()
        else:
            session.add(
                AttendanceRecord(
                    group_id=group_id,
                    student_id=st.id,
                    attendance_date=attendance_date,
                    status=db_s,
                )
            )


def fetch_attendance_nested(
    session: Session,
    group_id: str,
    *,
    start: Optional[date] = None,
    end: Optional[date] = None,
    student_name: Optional[str] = None,
) -> Dict[str, Dict[str, str]]:
    """
    Build ``{ date_str: { student_name: ui_status } }`` compatible with legacy frontend.

    Filters are optional; all three narrow the SQL query for performance.
    """
    q = select(AttendanceRecord).where(AttendanceRecord.group_id == group_id)
    if start:
        q = q.where(AttendanceRecord.attendance_date >= start)
    if end:
        q = q.where(AttendanceRecord.attendance_date <= end)
    if student_name and student_name.strip():
        st = session.scalar(
            select(Student).where(
                and_(Student.group_id == group_id, Student.full_name == student_name.strip())
            )
        )
        if not st:
            return {}
        q = q.where(AttendanceRecord.student_id == st.id)

    q = q.options(selectinload(AttendanceRecord.student))
    out: Dict[str, Dict[str, str]] = {}
    for row in session.scalars(q).all():
        dkey = row.attendance_date.isoformat()
        out.setdefault(dkey, {})[row.student.full_name] = db_status_to_ui(row.status)
    return out


def iter_attendance_matrix(
    session: Session,
    group_id: str,
    dates: List[date],
    student_names: List[str],
) -> Tuple[List[List[str]], Dict[str, Tuple[int, int, int, float]]]:
    """
    Returns (matrix rows aligned with student_names, each row list of cell tokens),
    and per-student summary: (present, absent, late, pct) over ``dates`` length.
    """
    if not dates:
        return [], {}

    start_d, end_d = min(dates), max(dates)
    studs = session.scalars(
        select(Student).where(Student.group_id == group_id).order_by(Student.sort_order)
    ).all()
    name_to_id = {s.full_name: s.id for s in studs}

    q = select(AttendanceRecord).where(
        and_(
            AttendanceRecord.group_id == group_id,
            AttendanceRecord.attendance_date >= start_d,
            AttendanceRecord.attendance_date <= end_d,
        )
    )
    cell_map: Dict[Tuple[int, date], str] = {}
    for row in session.scalars(q).all():
        cell_map[(row.student_id, row.attendance_date)] = row.status

    n_days = len(dates)
    matrix: List[List[str]] = []
    summaries: Dict[str, Tuple[int, int, int, float]] = {}

    for nm in student_names:
        sid = name_to_id.get(nm)
        row_cells: List[str] = []
        p = a = l = 0
        if sid is None:
            matrix.append(["—"] * n_days)
            summaries[nm] = (0, 0, 0, 0.0)
            continue
        for d in dates:
            st = cell_map.get((sid, d))
            if st == "present":
                row_cells.append("ح")
                p += 1
            elif st == "absent":
                row_cells.append("غ")
                a += 1
            elif st == "late":
                row_cells.append("م")
                l += 1
            else:
                row_cells.append("—")
        # Percentage: (present + late) / number of calendar slots in export
        denom = n_days if n_days else 1
        pct = round(100.0 * (p + l) / denom, 1)
        matrix.append(row_cells)
        summaries[nm] = (p, a, l, pct)

    return matrix, summaries


def day_records_for_export(
    session: Session, group_id: str, attendance_date: date
) -> Dict[str, str]:
    """Map student_name -> raw UI status string for single-day Excel (حاضر/غائب/متاخر)."""
    q = select(AttendanceRecord).where(
        and_(
            AttendanceRecord.group_id == group_id,
            AttendanceRecord.attendance_date == attendance_date,
        )
    )
    q = q.options(selectinload(AttendanceRecord.student))
    m: Dict[str, str] = {}
    for row in session.scalars(q).all():
        m[row.student.full_name] = db_status_to_ui(row.status)
    return m
