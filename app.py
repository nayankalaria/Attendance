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
        return {(r["date"], r["sid"]): r for r in csv.DictReader(f)}


def upsert(day, s, **kw):
    with LOCK:
        d = read_log()
        key = (day, s["sid"])
        row = d.get(key) or {"date": day, "sid": s["sid"], "status": "present", "sent": "0", "note": ""}
        row.update({"name": s["name"], "class": s["class"], "phone": s["phone"]})
        row.update({k: str(v) for k, v in kw.items()})
        d[key] = row
        with open(LOG, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDS)
            w.writeheader()
            w.writerows(d.values())


def message(s, day):
    c = f" ({s['class']})" if s["class"] else ""
    return f"Dear Parent,\nYour child {s['name']}{c} was absent from class on {day}.\n- {SCHOOL}"


# ---------------- WhatsApp (Playwright worker thread) ----------------
class WA:
    def __init__(self):
        self.state = "off"  # off | login | connected | error
        self.q = queue.Queue()
        self.job = {"running": False, "total": 0, "done": 0}
        self.thread = None
        self.error = ""
        self.qr_updated = 0
        self.qr_error = ""
        self.browser_started = False
        self.page_url = ""
        self.stop_event = threading.Event()

    def capture_qr(self, page):
        """Capture a visible WhatsApp pairing code and log useful diagnostics."""
        try:
            # WhatsApp has changed its QR markup over time. Try likely QR
            # elements first, then visible canvases (the QR is commonly a canvas).
            selectors = [
                'canvas[aria-label*="QR" i]',
                'img[alt*="QR" i]',
                'svg[aria-label*="QR" i]',
                'canvas',
                'div[data-ref] canvas',
            ]
            for selector in selectors:
                loc = page.locator(selector)
                count = min(loc.count(), 8)
                for index in range(count):
                    item = loc.nth(index)
                    try:
                        if not item.is_visible():
                            continue
                        box = item.bounding_box()
                        if not box or box.get("width", 0) < 100 or box.get("height", 0) < 100:
                            continue
                        item.screenshot(path=QR_FILE, timeout=10000)
                        if os.path.exists(QR_FILE) and os.path.getsize(QR_FILE) > 500:
                            self.qr_updated = time.time()
                            self.qr_error = ""
                            print(f"[wa] QR captured using {selector} ({int(box['width'])}x{int(box['height'])})", flush=True)
                            return True
                    except Exception as exc:
                        self.qr_error = str(exc)[:240]
            # A full-page screenshot is useful for diagnosing changed markup,
            # but is not served as a pairing QR.
            self.qr_error = self.qr_error or "WhatsApp login page loaded, but no visible QR canvas/image was found."
            print(f"[wa] QR not found; url={page.url!r}; title={page.title()!r}; detail={self.qr_error}", flush=True)
            try:
                print(f"[wa] visible canvases={page.locator('canvas:visible').count()}, images={page.locator('img:visible').count()}", flush=True)
                page.screenshot(path=os.path.join(DATA_DIR, "whatsapp_debug.png"), full_page=True, timeout=10000)
            except Exception as exc:
                print(f"[wa] diagnostic screenshot failed: {exc}", flush=True)
            return False
        except Exception as exc:
            self.qr_error = str(exc)[:240]
            print(f"[wa] QR capture exception: {self.qr_error}", flush=True)
            return False

    def start(self):
        if not HAS_PW:
            return
        if self.thread and self.thread.is_alive():
            return  # already running
        # (re)start cleanly: old thread is dead (window closed / crashed)
        self.error = ""
        self.qr_error = ""
        self.browser_started = False
        self.page_url = ""
        self.qr_updated = 0
        try:
            if os.path.exists(QR_FILE):
                os.remove(QR_FILE)
        except OSError:
            pass
        self.state = "login"
        self.job = {"running": False, "total": 0, "done": 0}
        self.q = queue.Queue()
        self.stop_event = threading.Event()
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def restart(self):
        """Stop the current browser loop cleanly, then start a fresh session."""
        if self.thread and self.thread.is_alive():
            self.stop_event.set()
            self.thread.join(timeout=12)
            if self.thread.is_alive():
                self.error = "WhatsApp browser is still shutting down. Please try again in a few seconds."
                return False
        self.start()
        return True

    def _run(self):
        try:
            print("[wa] Starting Playwright Chromium (headless)", flush=True)
            with sync_playwright() as p:
                ctx = p.chromium.launch_persistent_context(
                    PROFILE, headless=True,
                    args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
                    viewport={"width": 1280, "height": 900},
                )
                self.browser_started = True
                try:
                    page = ctx.pages[0] if ctx.pages else ctx.new_page()
                    page.on("dialog", lambda d: d.accept())
                    print("[wa] Navigating to WhatsApp Web", flush=True)
                    page.goto("https://web.whatsapp.com", wait_until="domcontentloaded", timeout=90000)
                    print(f"[wa] WhatsApp page loaded: {page.url}", flush=True)
                    while not self.stop_event.is_set():
                        if page.is_closed():
                            raise RuntimeError("WhatsApp page was closed. Click Reconnect.")
                        try:
                            self.page_url = page.url
                            if page.locator("#pane-side").count():
                                self.state = "connected"
                            else:
                                self.state = "login"
                                self.capture_qr(page)
                        except Exception as exc:
                            raise RuntimeError(f"WhatsApp page check failed: {str(exc)[:160]}")
                        try:
                            day, rows = self.q.get(timeout=2)
                        except queue.Empty:
                            continue
                        self._send_all(page, day, rows)
                finally:
                    try:
                        ctx.close()
                    except Exception:
                        pass
        except Exception as e:
            self.state, self.error = "error", str(e)[:200]
            self.job["running"] = False
            print(f"[wa] stopped: {self.error}", flush=True)

    def _send_all(self, page, day, rows):
        self.job = {"running": True, "total": len(rows), "done": 0}
        print(f"[send] {len(rows)} message(s) for {day}", flush=True)
        for s in rows:
            try:
                print(f"[send] -> {s['name']} ({s['phone']})", flush=True)
                self._send_one(page, s, day)
                upsert(day, s, sent=1, note="sent " + datetime.now().strftime("%H:%M:%S"))
                print("[send]    sent", flush=True)
            except Exception as e:
                print(f"[send]    FAILED: {e}", flush=True)
                try:
                    page.screenshot(path="last_error.png")
                except Exception:
                    pass
                upsert(day, s, sent=0, note="failed: " + str(e)[:80])
            self.job["done"] += 1
            time.sleep(random.uniform(5, 10))  # gap between messages = lower ban risk
        self.job["running"] = False

    def _send_one(self, page, s, day):
        url = f"https://web.whatsapp.com/send?phone={s['phone']}&text={urllib.parse.quote(message(s, day))}"
        page.goto(url)
        box = page.locator('div[contenteditable="true"][data-tab="10"], footer div[contenteditable="true"]')
        invalid = page.get_by_text(re.compile("phone number shared via url is invalid", re.I))
        for _ in range(60):  # wait up to ~60s for the chat to open
            if invalid.count():
                raise RuntimeError("number not on WhatsApp")
            if box.count() and box.first.is_visible():
                break
            time.sleep(1)
        else:
            raise RuntimeError("chat did not open (see last_error.png)")
        time.sleep(2)  # let the prefilled text appear
        if not box.first.inner_text().strip():
            raise RuntimeError("message text was not prefilled")
        send_btn = page.locator('button[aria-label="Send"], span[data-icon="send"]')
        if send_btn.count():
            send_btn.first.click()
        else:
            box.first.click()  # focus first, otherwise Enter does nothing
            page.keyboard.press("Enter")
        for _ in range(10):  # message box becomes empty once the message is sent
            time.sleep(1)
            if not box.first.inner_text().strip():
                time.sleep(1)
                return
        raise RuntimeError("could not press send (see last_error.png)")


wa = WA()

# ---------------- Web app ----------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "")
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SECURE=os.environ.get("COOKIE_SECURE", "true").lower() not in ("0", "false", "no"),
    SESSION_COOKIE_SAMESITE="Lax",
    PERMANENT_SESSION_LIFETIME=__import__("datetime").timedelta(hours=12),
)
APP_USERNAME = os.environ.get("APP_USERNAME", "")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

@app.before_request
def require_login():
    if request.endpoint in ("login", "healthz") or request.path.startswith("/static/"):
        return None
    if not APP_USERNAME or not APP_PASSWORD or not app.secret_key:
        return ("Application is not configured. Set APP_USERNAME, APP_PASSWORD and SECRET_KEY.", 503)
    if session.get("authenticated") and session.get("username") == APP_USERNAME:
        return None
    if request.path.startswith("/api/"):
        return jsonify(error="Authentication required"), 401
    return redirect(url_for("login"))

@app.get("/healthz")
def healthz():
    return jsonify(ok=True)

@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if (APP_USERNAME and APP_PASSWORD and
            hmac.compare_digest(username, APP_USERNAME) and
            hmac.compare_digest(password, APP_PASSWORD)):
            session.clear()
            session["authenticated"] = True
            session["username"] = APP_USERNAME
            session.permanent = True
            return redirect(url_for("home"))
        error = "Incorrect username or password."
    return render_template_string(LOGIN_PAGE, error=error)

@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.get("/api/qr.png")
def qr_image():
    if not os.path.exists(QR_FILE):
        return ("QR is not available yet. Keep this page open; check /api/status and Render logs for QR diagnostics.", 404)
    response = make_response(send_file(QR_FILE, mimetype="image/png"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return response

LOGIN_PAGE = """<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Attendance Alerts Login</title><style>
body{font-family:system-ui,sans-serif;background:#f5f6f8;display:grid;place-items:center;min-height:100vh;margin:0;color:#222}
main{background:white;padding:28px;border-radius:14px;width:min(360px,calc(100% - 56px));box-shadow:0 2px 12px #0001}
input,button{box-sizing:border-box;width:100%;padding:12px;margin:8px 0;border:1px solid #ccc;border-radius:7px;font:inherit}
button{background:#15803d;color:white;border:0;font-weight:600}.error{color:#b91c1c}
</style></head><body><main><h2>Attendance Alerts</h2><p>Sign in to continue.</p>
<form method="post"><input name="username" autocomplete="username" placeholder="Username" required>
<input name="password" type="password" autocomplete="current-password" placeholder="Password" required>
<button type="submit">Sign in</button></form><p class="error">{{ error }}</p></main></body></html>"""


@app.get("/api/status")
def status():
    return jsonify(
        pw=HAS_PW, wa=wa.state, error=wa.error, job=wa.job, school=SCHOOL,
        browser_started=wa.browser_started, page_url=wa.page_url,
        qr_ready=os.path.exists(QR_FILE), qr_updated=wa.qr_updated,
        qr_error=wa.qr_error,
    )


@app.post("/api/connect")
def connect():
    if not HAS_PW:
        return jsonify(error="Playwright not installed. Use the Open chat links instead."), 400
    if wa.thread and wa.thread.is_alive():
        if not wa.restart():
            return jsonify(error=wa.error), 409
    else:
        wa.start()
    return jsonify(ok=True)


@app.get("/api/students")
def students():
    day = request.args.get("date") or str(_date.today())
    try:
        studs = load_students()
    except Exception as e:
        return jsonify(error=str(e)), 400
    log = read_log()
    out = []
    for s in studs:
        r = log.get((day, s["sid"]), {})
        out.append({**s, "status": r.get("status", "present"), "sent": r.get("sent") in ("1", "manual"),
                    "note": r.get("note", ""),
                    "link": f"https://wa.me/{s['phone']}?text={urllib.parse.quote(message(s, day))}"})
    return jsonify(out)


@app.post("/api/send")
def send():
    b = request.get_json() or {}
    day, absent, force = b.get("date"), set(b.get("absent", [])), bool(b.get("force"))
    if wa.state != "connected":
        return jsonify(error="WhatsApp not connected. Click Connect and scan the QR."), 400
    if wa.job["running"]:
        return jsonify(error="A send is already running"), 409
    log = read_log()
    todo = []
    for s in load_students():
        isabs = s["sid"] in absent
        upsert(day, s, status="absent" if isabs else "present")
        if isabs and (force or log.get((day, s["sid"]), {}).get("sent") != "1"):
            todo.append(s)
    if not todo:
        return jsonify(error="Nothing to send: tick an absent student. Already-sent ones are skipped unless you tick \"Send again\"."), 400
    wa.job = {"running": True, "total": len(todo), "done": 0}
    wa.q.put((day, todo))
    return jsonify(queued=len(todo))


@app.post("/api/mark")
def mark():
    b = request.get_json() or {}
    for s in load_students():
        if s["sid"] == b.get("sid"):
            upsert(b["date"], s, status="absent", sent="manual", note="opened manually")
    return jsonify(ok=True)


PAGE = """<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Attendance Alerts</title><style>
body{font-family:system-ui,sans-serif;max-width:820px;margin:0 auto;padding:16px;background:#f5f6f8;color:#222}
.card{background:#fff;border-radius:10px;padding:16px;margin-bottom:16px;box-shadow:0 1px 3px #0001}
h2{margin:0 0 10px;font-size:18px}table{width:100%;border-collapse:collapse}
td,th{padding:8px 6px;border-bottom:1px solid #eee;text-align:left;font-size:14px}
button,input{font:inherit;padding:8px 12px;border-radius:6px;border:1px solid #ccc}
button{background:#16a34a;color:#fff;border:0;cursor:pointer}button.g{background:#6b7280}button:disabled{opacity:.5}
a.chat{color:#16a34a;font-weight:600;text-decoration:none}.ok{color:#16a34a;font-weight:600}.bad{color:#dc2626}small{color:#666}
</style></head><body>
<div class="card"><h2>1. WhatsApp</h2><div id="wa">...</div><div id="qrwrap" style="display:none;margin-top:12px"><img id="qr" alt="WhatsApp pairing QR" style="width:240px;max-width:100%;background:white"><p><small id="qrhint">Starting WhatsApp Web and waiting for the QR code…</small></p></div></div>
<div style="text-align:right;margin:8px 0"><form method="post" action="/logout"><button class="g" type="submit">Log out</button></form></div><div class="card"><h2>2. Attendance <input type="date" id="date" style="float:right"></h2>
<small>Students come from <b>students.xlsx</b>. Tick the box for ABSENT students.</small>
<table><thead><tr><th>Absent</th><th>Name</th><th>Class</th><th>Result</th><th>Manual</th></tr></thead><tbody id="tb"></tbody></table>
<p><button id="send">Send absence alerts</button> <button class="g" onclick="load()">Reload from Excel</button> <label><input type=checkbox id=force> Send again to already-sent</label> <span id="msg"></span></p></div>
<script>
const $=id=>document.getElementById(id);$('date').value=new Date().toLocaleDateString('en-CA');
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const j=async(u,o)=>(await fetch(u,o&&{headers:{'Content-Type':'application/json'},...o})).json();
async function st(){const s=await j('/api/status');let h='';
 if(s.wa==='connected'){h='<span class="ok">Connected</span>';$('qrwrap').style.display='none'}
 else if(s.wa==='login'){h='<span class="bad">Waiting for WhatsApp login</span> <button class="g" onclick="conn()">Reconnect WhatsApp</button> - scan the QR below.';$('qrwrap').style.display='block';if(s.qr_ready){$('qr').style.display='inline-block';$('qr').src='/api/qr.png?t='+Date.now();$('qrhint').textContent='Scan from WhatsApp → Linked devices → Link a device.'}else{$('qr').style.display='none';$('qrhint').textContent=s.qr_error||'Starting WhatsApp Web and waiting for the QR code…'}}
 else if(s.wa==='error'){h='<span class="bad">'+esc(s.error)+'</span> <button onclick="conn()">Reconnect</button>';$('qrwrap').style.display='none'}
 else {h=s.pw?'<button onclick="conn()">Connect WhatsApp</button>':'Auto-send unavailable (Playwright not installed). Use the <b>Open chat</b> links below.';$('qrwrap').style.display='none'}
 $('wa').innerHTML=h;$('send').disabled=s.wa!=='connected'||s.job.running;
 if(s.job.running){$('msg').textContent='Sending '+s.job.done+'/'+s.job.total+'...';load(true)}}
async function conn(){const r=await j('/api/connect',{method:'POST'});if(r.error)alert(r.error);st()}
async function load(keep){const c=new Set([...document.querySelectorAll('.ab:checked')].map(x=>x.value));
 const d=await j('/api/students?date='+$('date').value);if(d.error){$('tb').innerHTML='<tr><td colspan=5 class=bad>'+esc(d.error)+'</td></tr>';return}
 $('tb').innerHTML=d.map(s=>{const a=keep?c.has(s.sid):s.status==='absent';
  const r=s.sent?'<span class=ok>'+esc(s.note||'sent')+'</span>':(s.note?'<span class=bad>'+esc(s.note)+'</span>':'');
  return '<tr><td><input type=checkbox class=ab value="'+esc(s.sid)+'" '+(a?'checked':'')+'></td><td>'+esc(s.name)+'</td><td>'+esc(s.class)+'</td><td>'+r+
  '</td><td><a class=chat target=_blank href="'+esc(s.link)+'" onclick="mark(\\''+esc(s.sid).replace(/'/g,"\\\\'")+'\\')">Open chat</a></td></tr>'}).join('')}
async function mark(sid){await j('/api/mark',{method:'POST',body:JSON.stringify({date:$('date').value,sid})});setTimeout(()=>load(true),500)}
$('send').onclick=async()=>{const ids=[...document.querySelectorAll('.ab:checked')].map(x=>x.value);
 if(!confirm('Send alerts to parents of '+ids.length+' absent student(s)?'))return;
 const r=await j('/api/send',{method:'POST',body:JSON.stringify({date:$('date').value,absent:ids,force:$('force').checked})});
 $('msg').textContent=r.error||('Queued '+r.queued+' message(s)...')};
$('date').onchange=()=>load();load();st();setInterval(st,3000);
</script></body></html>"""


@app.get("/")
def home():
    return render_template_string(PAGE)


if __name__ == "__main__":
    ensure_sample()
    print("Starting Attendance Alerts")
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8080")), debug=False, threaded=True)
