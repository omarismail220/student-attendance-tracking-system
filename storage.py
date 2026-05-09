"""
JSON persistence layer for student groups and attendance records.

Data layout (writes under ``data_dir``):

* ``groups.json`` — mapping ``group_id`` → ``{label, time, students: [...]}``.
* ``attendance.json`` — mapping ``group_id`` → ``{date -> {student_name → status}}``.

All saves are atomic (write temp + replace) where possible so partial writes cannot
leave corrupt files on SIGKILL during rename (best-effort).
"""

from __future__ import annotations

import json
import os
import threading
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict

# Used only when ``groups.json`` is missing — same content as legacy in-app STUDENTS.
DEFAULT_GROUPS: Dict[str, Dict[str, Any]] = {
    "A1": {
        "label": "المجموعة A1",
        "time": "10:00 - 12:00",
        "students": [
            "احمد سعيد شرقاوي نصر الرقيق",
            "اساف عادل كامل فرج الله",
            "اسلام اشرف عبد المجيد عبد الفتاح",
            "حبيبه محمد كمال محمد عامر",
            "حال حاتم محمد الحسيني عبد العظيم",
            "خلود محمد مصطفى بشرى",
            "رانيا يوسف راشد عبد المجيد",
            "رزق محمد رزق محمد",
            "روان شريف عطيه عبدالفتاح",
            "روان مصطفى مصطفى السيد على",
            "زياد احمد مرعى عبد اللطيف قاسم",
            "سما حسام ممدوح شوقي",
            "عبد الرحمن ابراهيم سيد خليل",
            "عبد الرحمن احمد سعيد كامل سالم",
            "عبد الرحمن طارق عبد الحميد والى",
            "علاء أحمد عبدالله محمد",
            "عمار ياسر فتحي احمد",
        ],
    },
    "B1": {
        "label": "المجموعة B1",
        "time": "8:00 - 10:00",
        "students": [
            "فتحي محمد فتحي محمد الشاذلي",
            "لوجين ابراهيم مصطفى محمد",
            "ليلى شعبان احمد امام",
            "محمد أشرف يحيى عبدالله",
            "محمد محمد محمد محمد العراقي",
            "محمد هيثم محمد فضل على مسعود",
            "محمد ياسر على الجميل بدر",
            "محمود احمد محمود عبدالجواد كريم",
            "مصطفى محمد عبد الفتاح محمد",
            "ملك محمد امام ابراهيم",
            "منه جمال محمد دمرداش",
            "ميار حاتم عبده سعد عثمان",
            "ندى محمد ابراهيم محمد التميمي",
            "هادى شريف محمد الهادي محمد على",
            "هشام سعيد صالح عثمان",
            "والء على سيد محمد عبد الجواد",
            "يمنى إسماعيل إسماعيل إمام حسن",
            "يوسف احمد جلال محمد احمد عامر",
            "يوسف اشرف ابراهيم ذكى",
            "يوسف طارق محمد امام محمود",
            "يوسف محمد جمال احمد شربيني فروح",
        ],
    },
    "A2": {
        "label": "المجموعة A2",
        "time": "14:00 - 16:00",
        "students": [
            "احمد تميم فيصل تميم",
            "احمد محمد محمد احمد امين",
            "جلال هاني جلال فتحي",
            "حبيبه اسامه عبد المنعم عبد الرحيم",
            "رؤى هشام فهمى ريحان",
            "رحمه فؤاد سيد احمد مهدى",
            "زياد محمد فتحي عباس السيد",
            "ساره صالح محمد محمدين",
            "شهد ابراهيم محمد ابراهيم",
        ],
    },
    "B2": {
        "label": "المجموعة B2",
        "time": "12:00 - 14:00",
        "students": [
            "عبد الله نجاح حامد احمد عزازي",
            "عمرو علاء الدين سيد إبراهيم خليل",
            "مؤمن على فتحي عبد العاطي مبروك",
            "ماريا فيكتور عياد يوسف ابراهيم",
            "مازن محمد رائف حافظ",
            "مريم محمد حافظ عبد العال",
            "مصطفى عبد المحسن مصطفى محمد",
            "منه محمد ابو الحجاج محمد",
            "نور الدين ادهم نور الدين مصطفى",
            "رسمية مسعود سعد محمد ياسين",
        ],
    },
    "A3": {
        "label": "المجموعة A3",
        "time": "16:00 - 18:00",
        "students": [
            "ابراهيم مسعود سيد ابراهيم",
            "احمد سعيد فتحي السيد محمد",
            "اسراء سيد عبد العزيز جلال",
            "بافلي حشمت بولص حنا",
            "جنه هاني صالح عبد المنعم احمد",
            "جوزيف عادل عزيز عياد",
            "جومانا هشام سيد بدوي محمد",
            "حازم حسنى فهيم محمد",
            "شهد محمد رمضان إبراهيم",
            "عبد الله حسن السيد احمد موسى",
            "عبد الله طارق مصباح عبد الحميد",
            "عمر اسماعيل حلمى أبو ضيف",
        ],
    },
    "B3": {
        "label": "المجموعة B3",
        "time": "18:00 - 20:00",
        "students": [
            "عمر ايمن حنفى محمود",
            "عمر كمال عبد النبي طلبة زايد",
            "كنزى حمدى سالم أمين سالم",
            "محمد إبراهيم طويل محمد سالم",
            "محمد سيد محمد حسن",
            "مصطفى محمد محمود مصطفى سويدان",
            "ملك طارق عواد عيد",
            "منة الله عبد العال على عبد العال",
        ],
    },
    "A4": {
        "label": "المجموعة A4",
        "time": "16:00 - 18:00",
        "students": [
            "احمد ايهاب بديع محمد الماظ",
            "احمد سعيد صابر كامل احمد",
            "احمد عرفه عمر احمد مكاوي",
            "الاء كرم خليفه عبد ربه عبد الرحمن",
            "اميرة صبحي عبد الشافي حسن الصغير",
            "بسملة محمد عبد العزيز عبد الحافظ",
            "بسمه هاني حفني محمد",
            "تونى ايمن كمال يونان",
            "جنى ياسر على حامد",
            "حازم محمود عبد الحميد محمود محمد",
            "روان رضا رمضان حسن وهدان",
            "روان محمد سالم حسن",
            "روان وليد احمد محمد",
            "زياد عادل امين عبد الصمد",
            "زينب حسام الدين محمد حسين",
            "سارة حسين عبد الحليم الجمال",
            "سلفانا عاطف نصرى عزيز",
            "شمس الدين شوقي عبد الله عبد القادر",
            "شهد حسنى حامد احمد حسن",
            "صبحى عصام صبحى عبد القوى",
            "صفية يحيى احمد على",
            "عمار عماد باهى منصور",
        ],
    },
    "B4": {
        "label": "المجموعة B4",
        "time": "18:00 - 20:00",
        "students": [
            "عمر أحمد شلقامى ذكي",
            "عمر خالد عبد الله احمد حسن",
            "عهد عبد العليم محمود محمود قديرة",
            "فريدة حاتم عبد القادر عبد العزيز محمد",
            "محمد حسين عبد التواب علي",
            "محمد رمضان عبد التواب السيد",
            "محمد محمود عبد الحميد محمود",
            "مريم بخيت فكرى نصير",
            "مريم ياسر حسن قطب حسن",
            "ملك محمد محمد صالح النبالوى",
            "منه الله نبيل سعيد محمود",
            "منه محمد السيد حسن خميس",
            "ناصر حسن ثابت حسين",
            "نرمين حنا سعيد حنا",
            "نور احمد سعيد محمد احمد ابو العز",
            "نور الهدى عمار على محمد عبدالمجيد",
            "نور محمد حمدى حسن",
            "نورهان رفاعي محجوب خليفة محمد",
            "هاجر محمد احمد عبد الهادي",
            "هاله السيد احمد سالم",
            "هبة الله مدحت حسن صقر مبارك",
            "ياسمين عبد الفتاح محمد عبد الفتاح",
            "يمنى يوسف احمد حامد رضوان",
            "يوسف ابراهيم سيد ابراهيم",
            "يوسف عادل بديع نخلة حنين",
            "يوسف محمد سيد احمد احمد المدني",
        ],
    },
}


_lock = threading.Lock()


def _data_dir(app_root: Path | None = None) -> Path:
    root = app_root or Path(__file__).resolve().parent
    d = root / "data"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _atomic_write_json(path: Path, obj: Any) -> None:
    """Serialize ``obj`` to ``path`` via a temp file + replace."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    data = json.dumps(obj, ensure_ascii=False, indent=2)
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(data)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, path)


def _read_json(path: Path, default: Any) -> Any:
    if not path.is_file():
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_groups_disk(app_root: Path | None = None) -> Dict[str, Dict[str, Any]]:
    """
    Load groups from disk. If missing, seed from ``DEFAULT_GROUPS`` and persist
    so deployments start with a deterministic snapshot once.
    """
    path = _data_dir(app_root) / "groups.json"
    with _lock:
        raw = _read_json(path, None)
        if raw is None or not isinstance(raw, dict):
            data = deepcopy(DEFAULT_GROUPS)
            _atomic_write_json(path, data)
            return data
        return raw


def save_groups(groups: Dict[str, Dict[str, Any]], app_root: Path | None = None) -> None:
    """Overwrite ``groups.json`` with full mapping (validated caller-side)."""
    path = _data_dir(app_root) / "groups.json"
    with _lock:
        _atomic_write_json(path, groups)


def load_attendance_disk(app_root: Path | None = None) -> Dict[str, Dict[str, Dict[str, str]]]:
    """Load persisted attendance bundle; absent file → empty dict."""
    path = _data_dir(app_root) / "attendance.json"
    with _lock:
        raw = _read_json(path, {})
        return raw if isinstance(raw, dict) else {}


def save_attendance(
    attendance: Dict[str, Dict[str, Dict[str, str]]], app_root: Path | None = None
) -> None:
    """Persist the full in-memory attendance store."""
    path = _data_dir(app_root) / "attendance.json"
    with _lock:
        _atomic_write_json(path, attendance)


def prune_attendance_for_group(group_id: str, valid_names: set[str], attendance: Dict) -> None:
    """
    In-place: drop attendance keys whose student name no longer exists in ``valid_names``.
    Called after Excel import replaces the roster while keeping historical sheets consistent.
    """
    block = attendance.get(group_id)
    if not isinstance(block, dict):
        return
    for dat, rec in list(block.items()):
        if not isinstance(rec, dict):
            block.pop(dat, None)
            continue
        cleaned = {k: v for k, v in rec.items() if k in valid_names}
        block[dat] = cleaned
