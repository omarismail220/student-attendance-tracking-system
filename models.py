"""
SQLAlchemy ORM models: groups, students, attendance with indexes and relationships.

``AttendanceRecord`` uses a composite unique constraint so the same student cannot have
duplicate rows for one calendar day; saves become upserts.
"""

from __future__ import annotations

from datetime import date, datetime

from sqlalchemy import (
    Date,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    String,
    UniqueConstraint,
    func,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship


class Base(DeclarativeBase):
    pass


class Group(Base):
    __tablename__ = "groups"

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    label: Mapped[str] = mapped_column(String(255), nullable=False)
    time_slot: Mapped[str] = mapped_column(String(64), nullable=False)

    students: Mapped[list["Student"]] = relationship(
        back_populates="group",
        cascade="all, delete-orphan",
        order_by="Student.sort_order",
    )


class Student(Base):
    __tablename__ = "students"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    group_id: Mapped[str] = mapped_column(
        String(32), ForeignKey("groups.id", ondelete="CASCADE"), nullable=False, index=True
    )
    full_name: Mapped[str] = mapped_column(String(512), nullable=False)
    sort_order: Mapped[int] = mapped_column(Integer, nullable=False, default=0)

    group: Mapped["Group"] = relationship(back_populates="students")
    attendance_rows: Mapped[list["AttendanceRecord"]] = relationship(
        back_populates="student",
        cascade="all, delete-orphan",
    )

    __table_args__ = (
        UniqueConstraint("group_id", "full_name", name="uq_students_group_full_name"),
        Index("ix_students_group_sort", "group_id", "sort_order"),
    )


class AttendanceRecord(Base):
    __tablename__ = "attendance"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    group_id: Mapped[str] = mapped_column(
        String(32), ForeignKey("groups.id", ondelete="CASCADE"), nullable=False, index=True
    )
    student_id: Mapped[int] = mapped_column(
        Integer, ForeignKey("students.id", ondelete="CASCADE"), nullable=False, index=True
    )
    attendance_date: Mapped[date] = mapped_column(Date, nullable=False, index=True)
    # Canonical values: present | absent | late (English for stable storage / APIs)
    status: Mapped[str] = mapped_column(String(16), nullable=False)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime, server_default=func.now(), onupdate=func.now(), nullable=False
    )

    group: Mapped["Group"] = relationship()
    student: Mapped["Student"] = relationship(back_populates="attendance_rows")

    __table_args__ = (
        UniqueConstraint(
            "group_id", "student_id", "attendance_date", name="uq_attendance_group_student_day"
        ),
        Index("ix_attendance_group_date", "group_id", "attendance_date"),
        Index("ix_attendance_student_date", "student_id", "attendance_date"),
    )
