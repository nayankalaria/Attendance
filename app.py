"""
Attendance -> WhatsApp absence alerts. ONE file, run from terminal:

    pip install flask openpyxl playwright
    playwright install chromium
    python app.py

- Student list comes from students.xlsx (columns: Name | Class | Parent Phone).
  A sample file is created on first run. Edit it in Excel any time.
- Opens http://127.0.0.1:5000 in your browser: tick absent students, press send.
- Auto-send uses a Chrome window controlled by Playwright. The WhatsApp login is
  saved in ./wa_profile, so you normally scan the QR only ONCE (not daily).
- No Playwright? The app still works: it gives a one-click "Open chat" link per
  parent (WhatsApp opens with the message ready, you just press send).
- Every attendance + send result is saved in attendance_log.csv.
"""
import csv, os, queue, random, re, threading, time, urllib.parse, webbrowser
from datetime import date as _date, datetime

from flask import Flask, jsonify, request, render_template_string, redirect, url_for, session, send_file, make_response
from openpyxl import Workbook, load_workbook
from werkzeug.middleware.proxy_fix import ProxyFix
import hmac

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-to-a-random-secret")

try:
    from playwright.sync_api import sync_playwright
    HAS_PW = True
except ImportError:
    HAS_PW = False

SCHOOL = os.environ.get("SCHOOL_NAME", "School")
DATA_DIR = os.environ.get("DATA_DIR", "/data")
os.makedirs(DATA_DIR, exist_ok=True)
XLSX = os.path.join(DATA_DIR, "students.xlsx")
LOG = os.path.join(DATA_DIR, "attendance_log.csv")
PROFILE = os.path.abspath(os.path.join(DATA_DIR, "wa_profile"))
QR_FILE = os.path.join(DATA_DIR, "whatsapp_qr.png")
BUNDLED_XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), "students.xlsx")
if not os.path.exists(XLSX) and os.path.exists(BUNDLED_XLSX):
    import shutil
    shutil.copy2(BUNDLED_XLSX, XLSX)
FIELDS = ["date", "sid", "name", "class", "phone", "status", "sent", "note"]
LOCK = threading.Lock()


# ---------------- Excel students ----------------
def ensure_sample():
    if os.path.exists(XLSX):
        return
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Class", "Parent Phone"])
    ws.append(["Amit Patel", "5A", "9876543210"])
    ws.append(["Riya Shah", "5A", "9876500000"])
    wb.save(XLSX)


def norm_phone(v):
    d = re.sub(r"\D", "", str(v or ""))
    d = d.lstrip("0") if len(d) == 11 and d.startswith("0") else d
    return "91" + d if len(d) == 10 else d  # default India


def load_students():
    ensure_sample()
    ws = load_workbook(XLSX, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    head = [str(c or "").strip().lower() for c in rows[0]]

    def col(*names):
        for i, h in enumerate(head):
            if any(n in h for n in names):
                return i
        return None

    ci_name, ci_cls, ci_ph = col("name"), col("class", "std", "grade"), col("phone", "mobile", "contact")
    if ci_name is None or ci_ph is None:
        raise ValueError("students.xlsx needs columns: Name, Class, Parent Phone")
    out = []
    for r in rows[1:]:
        name = str(r[ci_name] or "").strip()
        phone = norm_phone(r[ci_ph])
        if not name or not phone:
            continue
        cls = str(r[ci_cls] or "").strip() if ci_cls is not None else ""
        out.append({"sid": f"{name}|{phone}", "name": name, "class": cls, "phone": phone})
    return out


# ---------------- CSV log ----------------
def read_log():
    if not os.path.exists(LOG):
        return {}

    with open(LOG, newline="", encoding="utf-8") as f:
        return {
            (r["date"], r["sid"]): r
            for r in csv.DictReader(f)
        }
