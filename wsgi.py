# -*- coding: utf-8 -*-
"""
استخدم هذا الملف على PythonAnywhere:
1. ارفع app.py و index.html و wsgi.py (و sample_attendance_A1.xlsx إن رغبت) في نفس المجلد، مثلاً:
   /home/omarismail220/mysite/
2. من تبويب Web > WSGI configuration file — انسخ محتوى هذا الملف إلى هناك،
   أو اجعل المسار يشير إلى هذا الملف إن كان الخيار متاحاً.
3. غيّر PROJECT_HOME أدناه ليطابق المجلد الفعلي الذي يوجد فيه app.py ثم Reload.
"""
import os
import sys

# المجلد الذي فيه app.py: إما نفس مجلد wsgi.py أو عيّن ATTENDANCE_APP_HOME في لوحة Web
PROJECT_HOME = os.environ.get(
    "ATTENDANCE_APP_HOME",
    os.path.dirname(os.path.abspath(__file__)),
)

if PROJECT_HOME not in sys.path:
    sys.path.insert(0, PROJECT_HOME)

# كائن WSGI الذي يتوقعه PythonAnywhere
from app import app as application
