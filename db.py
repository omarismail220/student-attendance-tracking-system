"""
Database engine, session factory, and schema bootstrap.

Environment:
  ATTENDANCE_DATABASE_URL — optional SQLAlchemy URL (default: SQLite under ``data/``).
"""

from __future__ import annotations

import os
from contextlib import contextmanager
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import Session, sessionmaker

from models import Base

_engine = None
_SessionLocal = None


def get_database_url(app_root: Path) -> str:
    env = os.environ.get("ATTENDANCE_DATABASE_URL", "").strip()
    if env:
        return env
    data_dir = app_root / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    # check_same_thread=False allows Flask dev + WSGI threads on SQLite
    return f"sqlite:///{(data_dir / 'attendance.db').as_posix()}"


def configure_engine(app_root: Path) -> None:
    global _engine, _SessionLocal
    url = get_database_url(app_root)
    connect_args = {}
    if url.startswith("sqlite:"):
        connect_args["check_same_thread"] = False
    _engine = create_engine(url, echo=False, future=True, connect_args=connect_args)
    _SessionLocal = sessionmaker(bind=_engine, autoflush=False, autocommit=False, future=True)


def init_db() -> None:
    """Create tables if they do not exist."""
    Base.metadata.create_all(bind=_engine)


def get_engine():
    return _engine


@contextmanager
def session_scope():
    """Transactional session context manager."""
    assert _SessionLocal is not None
    session: Session = _SessionLocal()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()
