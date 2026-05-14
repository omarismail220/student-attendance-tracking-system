from __future__ import annotations

import io
import os
import re
from datetime import date, datetime
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

from attendance_service import (
    day_records_for_export,
    fetch_attendance_nested,
    get_group_meta_and_students,
    iter_attendance_matrix,
    list_groups,
    migrate_from_json_files,
    replace_group_students,
    seed_default_groups_if_empty,
    upsert_attendance_for_date,
)
from db import configure_engine, init_db, session_scope
from excel_reports import (
    build_matrix_workbook,
    build_single_day_workbook,
    daterange_inclusive,
    month_dates,
)
from excel_students import build_student_template_roster_workbook, parse_student_workbook

# ── Flask app ─────────────────────────────────────────────────────────────────
_app_root = Path(__file__).resolve().parent
app = Flask(__name__)

MAX_UPLOAD_BYTES = 2 * 1024 * 1024
MAX_STUDENT_IMPORT_ROWS = 2000
MAX_EXPORT_RANGE_DAYS = 400

app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES

configure_engine(_app_root)
init_db()
with session_scope() as _boot_session:
    migrate_from_json_files(_app_root, _boot_session)
    _boot_session.flush()
    seed_default_groups_if_empty(_boot_session)

CORS(
    app,
    origins=[
        "https://student-attendance-tracking-system-3un4pgeat.vercel.app",
        "https://omarismail220.pythonanywhere.com",
        "http://127.0.0.1:5050",
        "http://localhost:5050",
        "http://localhost:3000",
        "http://127.0.0.1:3000",
        "null",
        re.compile(r"^https://[a-zA-Z0-9.-]+\.vercel\.app$"),
    ],
)

PA_SAMPLE_XLSX_DEFAULT = "/home/omarismail220/sample_attendance_A1.xlsx"


def resolve_sample_xlsx_path():
    env = os.environ.get("ATTENDANCE_SAMPLE_XLSX", "").strip()
    if env and os.path.isfile(env):
        return env
    local = os.path.join(_app_root, "sample_attendance_A1.xlsx")
    if os.path.isfile(local):
        return local
    if os.path.isfile(PA_SAMPLE_XLSX_DEFAULT):
        return PA_SAMPLE_XLSX_DEFAULT
    return None


def _parse_date_arg(name: str) -> date | None:
    raw = (request.args.get(name) or "").strip()
    if not raw:
        return None
    try:
        return date.fromisoformat(raw)
    except ValueError:
        return None


@app.route("/")
def index():
    return send_from_directory(str(_app_root), "index.html")


@app.route("/api/info", methods=["GET"])
def api_info():
    """Health + metadata; ``groups_count`` reflects SQLite ``groups`` table."""
    sample = resolve_sample_xlsx_path()
    with session_scope() as session:
        groups_list = list_groups(session)
    return jsonify(
        {
            "status": "ok",
            "service": "student-attendance-api",
            "version": "3.0.0",
            "storage": "sqlite",
            "groups_count": len(groups_list),
            "server_time": datetime.now().isoformat(timespec="seconds"),
            "sample_xlsx_available": bool(sample),
        }
    )


@app.route("/api/sample-attendance-xlsx", methods=["GET"])
def download_sample_attendance_xlsx():
    path = resolve_sample_xlsx_path()
    if not path:
        return (
            jsonify(
                {"error": "not_found", "message": "ملف النموذج غير موجود على الخادم"}
            ),
            404,
        )
    return send_file(
        path,
        as_attachment=True,
        download_name="sample_attendance_A1.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/groups", methods=["GET"])
def get_groups():
    """List groups from SQLite with live student counts."""
    with session_scope() as session:
        return jsonify(list_groups(session))


@app.route("/api/students/<group_id>", methods=["GET"])
def get_students(group_id):
    """
    Roster + group meta + full ``attendance`` nested map (legacy shape) for compatibility.

    Large histories can be heavy; prefer ``GET /api/attendance/<group_id>`` with filters
    for integrations that only need a window.
    """
    with session_scope() as session:
        meta = get_group_meta_and_students(session, group_id)
        if not meta:
            return jsonify({"error": "not_found", "message": "المجموعة غير موجودة"}), 404
        g, names = meta
        nested = fetch_attendance_nested(session, group_id)
        return jsonify(
            {
                "label": g.label,
                "time": g.time_slot,
                "students": names,
                "attendance": nested,
            }
        )


@app.route("/api/template/<group_id>", methods=["GET"])
def download_student_template(group_id):
    """RTL Excel roster template (prefilled names) for offline editing."""
    with session_scope() as session:
        meta = get_group_meta_and_students(session, group_id)
        if not meta:
            return jsonify({"error": "not_found", "message": "المجموعة غير موجودة"}), 404
        g, students = meta
        buf, fname = build_student_template_roster_workbook(
            group_id, g.label, g.time_slot, students
        )
    return send_file(
        io.BytesIO(buf),
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/import-students/<group_id>", methods=["POST"])
def import_students(group_id):
    """
    Replace roster in SQLite (``students`` rows, CASCADE removes orphan ``attendance``).
    """
    with session_scope() as session:
        if not get_group_meta_and_students(session, group_id):
            return jsonify({"error": "not_found", "message": "المجموعة غير موجودة"}), 404

    uf = request.files.get("file")
    if uf is None or uf.filename == "":
        return (
            jsonify({"success": False, "error": "missing_file", "message": "لم يُرفع أي ملف"}),
            400,
        )
    if not secure_filename(uf.filename).lower().endswith(".xlsx"):
        return (
            jsonify(
                {
                    "success": False,
                    "error": "bad_extension",
                    "message": "يُسمح فقط بملفات امتداد .xlsx",
                }
            ),
            415,
        )
    raw = uf.read()
    if len(raw) == 0:
        return (
            jsonify({"success": False, "error": "empty_file", "message": "الملف فارغ"}),
            400,
        )
    if len(raw) > MAX_UPLOAD_BYTES:
        return (
            jsonify({"success": False, "error": "file_too_large", "message": "حجم الملف كبير جداً"}),
            413,
        )
    names, parse_err = parse_student_workbook(
        raw, max_bytes=MAX_UPLOAD_BYTES, max_student_rows=MAX_STUDENT_IMPORT_ROWS
    )
    if parse_err:
        return (
            jsonify({"success": False, "error": "invalid_template", "message": parse_err}),
            422,
        )

    with session_scope() as session:
        replace_group_students(session, group_id, names)

    return jsonify({"success": True, "group_id": group_id, "imported_count": len(names)})


@app.route("/api/attendance/<group_id>", methods=["GET"])
def get_attendance_history(group_id):
    """
    Nested ``{ date: { student: status } }`` in Arabic UI strings.

    Optional filters (narrow SQL): ``start``, ``end`` (ISO dates), ``student`` (full name).
    """
    with session_scope() as session:
        if not get_group_meta_and_students(session, group_id):
            return jsonify({"error": "not_found"}), 404
        start = _parse_date_arg("start")
        end = _parse_date_arg("end")
        student = (request.args.get("student") or "").strip() or None
        if start and end and end < start:
            return (
                jsonify(
                    {
                        "error": "bad_range",
                        "message": "تاريخ النهاية يجب أن يكون بعد تاريخ البداية",
                    }
                ),
                400,
            )
        data = fetch_attendance_nested(session, group_id, start=start, end=end, student_name=student)
    return jsonify(data)


@app.route("/api/attendance", methods=["POST"])
def save_attendance_route():
    """
    Upsert per-student rows for one ``date``. Empty / unknown UI values delete the row
    for that day (no duplicate rows — unique on group + student + date).
    """
    data = request.json or {}
    gid = data.get("group_id")
    records = data.get("records", {})
    date_str = data.get("date", datetime.now().strftime("%Y-%m-%d"))
    if not isinstance(records, dict):
        records = {}
    try:
        ad = date.fromisoformat(str(date_str).strip())
    except ValueError:
        return jsonify({"error": "bad_date", "message": "صيغة التاريخ غير صالحة"}), 400

    with session_scope() as session:
        if not get_group_meta_and_students(session, gid):
            return jsonify({"error": "not_found"}), 404
        upsert_attendance_for_date(session, gid, ad, records)
    return jsonify({"status": "saved"})


@app.route("/api/export/<group_id>", methods=["GET"])
def export_attendance(group_id):
    """Single-day RTL Excel (unchanged URL); ``?date=YYYY-MM-DD`` optional."""
    ad = _parse_date_arg("date") or date.today()
    with session_scope() as session:
        meta = get_group_meta_and_students(session, group_id)
        if not meta:
            return jsonify({"error": "not found"}), 404
        g, names = meta
        rec = day_records_for_export(session, group_id, ad)
        blob = build_single_day_workbook(
            group_id, g.label, g.time_slot, ad, names, rec
        )
    return send_file(
        io.BytesIO(blob),
        as_attachment=True,
        download_name=f"attendance_{group_id}_{ad.isoformat()}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/export-range/<group_id>", methods=["GET"])
def export_attendance_range(group_id):
    """
    Matrix export: each column is a calendar day between ``start`` and ``end`` (inclusive).

    Query: ``start=YYYY-MM-DD`` & ``end=YYYY-MM-DD`` (required). Max span ``MAX_EXPORT_RANGE_DAYS``.
    """
    start = _parse_date_arg("start")
    end = _parse_date_arg("end")
    if not start or not end:
        return (
            jsonify(
                {
                    "error": "missing_params",
                    "message": "يُرجى تمرير start و end بصيغة YYYY-MM-DD",
                }
            ),
            400,
        )
    if end < start:
        start, end = end, start
    dates = daterange_inclusive(start, end)
    if len(dates) > MAX_EXPORT_RANGE_DAYS:
        return (
            jsonify(
                {
                    "error": "range_too_large",
                    "message": f"الحد الأقصى للفترة {MAX_EXPORT_RANGE_DAYS} يوماً",
                }
            ),
            400,
        )

    with session_scope() as session:
        meta = get_group_meta_and_students(session, group_id)
        if not meta:
            return jsonify({"error": "not_found", "message": "المجموعة غير موجودة"}), 404
        g, names = meta
        matrix, summaries = iter_attendance_matrix(session, group_id, dates, names)
        title = f"كشف حضور — {g.label} ({group_id})"
        subtitle = f"من {start.isoformat()} إلى {end.isoformat()} | الموعد: {g.time_slot}"
        legend = f"عدد الأيام: {len(dates)}"
        blob = build_matrix_workbook(title, subtitle, dates, names, matrix, summaries, legend)

    fn = f"attendance_range_{group_id}_{start.isoformat()}_{end.isoformat()}.xlsx"
    return send_file(
        io.BytesIO(blob),
        as_attachment=True,
        download_name=fn,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/export-month/<group_id>", methods=["GET"])
def export_attendance_month(group_id):
    """
    Full calendar month matrix. Query: ``year`` (default current), ``month`` 1–12 (default current).
    """
    now = datetime.now()
    try:
        year = int(request.args.get("year", now.year))
        month = int(request.args.get("month", now.month))
    except ValueError:
        return jsonify({"error": "bad_params", "message": "year و month يجب أن يكونا أرقاماً"}), 400
    if month < 1 or month > 12 or year < 2000 or year > 2100:
        return jsonify({"error": "bad_params", "message": "شهر أو سنة غير صالحين"}), 400

    dates = month_dates(year, month)
    start, end = dates[0], dates[-1]

    with session_scope() as session:
        meta = get_group_meta_and_students(session, group_id)
        if not meta:
            return jsonify({"error": "not_found", "message": "المجموعة غير موجودة"}), 404
        g, names = meta
        matrix, summaries = iter_attendance_matrix(session, group_id, dates, names)
        title = f"كشف حضور شهري — {g.label} ({group_id})"
        subtitle = f"{year}-{month:02d} | الموعد: {g.time_slot}"
        blob = build_matrix_workbook(title, subtitle, dates, names, matrix, summaries)

    fn = f"attendance_month_{group_id}_{year}_{month:02d}.xlsx"
    return send_file(
        io.BytesIO(blob),
        as_attachment=True,
        download_name=fn,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


API_ENDPOINTS = [
    "GET /api",
    "GET /api/info",
    "GET /api/groups",
    "GET /api/students/<group_id>",
    "POST /api/attendance",
    "GET /api/attendance/<group_id>",
    "GET /api/export/<group_id>",
    "GET /api/export-range/<group_id>?start=&end=",
    "GET /api/export-month/<group_id>?year=&month=",
    "GET /api/sample-attendance-xlsx",
    "GET /api/template/<group_id>",
    "POST /api/import-students/<group_id>",
    "GET /",
]


@app.route("/api", methods=["GET"])
def api_index():
    return jsonify({"service": "student-attendance-api", "endpoints": API_ENDPOINTS}), 200


@app.errorhandler(RequestEntityTooLarge)
def handle_too_large(_):
    return jsonify({"success": False, "error": "payload_too_large", "message": "طلب أكبر من الحد الأقصى"}), 413


@app.errorhandler(404)
def handle_404(exc):
    if request.path.startswith("/api"):
        return (
            jsonify(
                {"error": "not_found", "path": request.path, "endpoints": API_ENDPOINTS},
            ),
            404,
        )
    body = (
        '<!DOCTYPE html><html><body dir="rtl" style="font-family:sans-serif;padding:2rem;">'
        f"<h1>404</h1><p>المسار غير موجود: <code>{request.path}</code></p>"
        '<p><a href="/">الصفحة الرئيسية</a> — <a href="/api/info">/api/info</a></p></body></html>'
    )
    return body, 404


if __name__ == "__main__":
    app.run(debug=False, port=5050)
