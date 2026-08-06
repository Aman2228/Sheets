"""
app.py — OCS Master, web version.

Open the same URL from your phone or your PC — both hit the same Google
Sheet, so there's nothing to keep "in sync" manually.

Security model:
 - APP_PASSWORD env var gates the whole site (simple shared password, so a
   random person who finds your URL can't send email as you or touch your
   data). Set this in your host's environment variables, not in code.
 - Your IITD mailbox password is entered fresh in the browser each session,
   kept ONLY in server memory (never written to the Sheet/disk/logs), and
   auto-forgotten after 60 minutes — same behaviour as ocs_master.py's cache.
"""
import os, time, uuid, threading
from html import escape
from functools import wraps
from flask import Flask, request, redirect, url_for, session, render_template_string

import ocs_logic as L
from sheets_backend import load_workbook

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-key-change-me")

SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY")  # the long ID in the Sheet's URL
APP_PASSWORD = os.environ.get("APP_PASSWORD")
IITD_WEBMAIL_PASSWORD = os.environ.get("IITD_WEBMAIL_PASSWORD")
BROCHURE_PATH = os.environ.get("BROCHURE_FILE_PATH")  # optional, uploaded alongside app.py

# ---- in-memory password cache: {session_id: (password, expires_at)} ----
_pw_cache = {}
PW_TTL = 60 * 60  # 60 min, matches original

def _sid():
    if "sid" not in session:
        session["sid"] = uuid.uuid4().hex
    return session["sid"]

def get_cached_pw():
    if IITD_WEBMAIL_PASSWORD:
        return IITD_WEBMAIL_PASSWORD

    entry = _pw_cache.get(_sid())
    if not entry:
        return None

    pw, exp = entry
    if time.time() > exp:
        _pw_cache.pop(_sid(), None)
        return None

    return pw

def set_cached_pw(pw):
    _pw_cache[_sid()] = (pw, time.time() + PW_TTL)

def clear_cached_pw():
    _pw_cache.pop(_sid(), None)

# ---- in-memory job store for the continuous sender (runs in a background thread) ----
_jobs = {}  # job_id -> {"done": bool, "results": [...], "total": int}

def login_required(f):
    @wraps(f)
    def wrapper(*a, **kw):
        if APP_PASSWORD and not session.get("authed"):
            return redirect(url_for("login", next=request.path))
        return f(*a, **kw)
    return wrapper

def get_wb():
    if not SPREADSHEET_KEY:
        raise RuntimeError("SPREADSHEET_KEY env var is not set.")
    return load_workbook(SPREADSHEET_KEY)

# =====================================================================
# LAYOUT
# =====================================================================
BASE = """
<!doctype html><html><head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>OCS Master</title>
<style>
  :root{--fg:#1a1a1a;--muted:#6b6b6b;--line:#e3e3e3;--accent:#2a5bd7;--bad:#c0392b;--good:#1e8449;--bg:#fafafa;}
  *{box-sizing:border-box}
  body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;margin:0;background:var(--bg);color:var(--fg)}
  header{background:#111;color:#fff;padding:14px 16px;font-weight:600;display:flex;justify-content:space-between;align-items:center}
  header a{color:#fff;text-decoration:none;font-size:14px;opacity:.85}
  main{max-width:640px;margin:0 auto;padding:16px}
  .card{background:#fff;border:1px solid var(--line);border-radius:10px;padding:16px;margin-bottom:14px}
  a.btn,button{display:inline-block;background:var(--accent);color:#fff;border:none;border-radius:8px;
    padding:12px 16px;font-size:16px;text-decoration:none;cursor:pointer;width:100%;margin-top:8px}
  button.secondary,a.btn.secondary{background:#fff;color:var(--accent);border:1px solid var(--accent)}
  button.danger{background:var(--bad)}
  input[type=text],input[type=password],input[type=number]{width:100%;padding:11px;border:1px solid var(--line);
    border-radius:8px;font-size:16px;margin-top:4px}
  label{font-size:13px;color:var(--muted);display:block;margin-top:10px}
  .row{display:flex;gap:8px}
  .row>*{flex:1}
  .tag{display:inline-block;font-size:12px;padding:2px 8px;border-radius:12px;background:#eee;margin-left:6px}
  .tag.sent{background:#e8f6ee;color:var(--good)}
  .tag.fail{background:#fdecea;color:var(--bad)}
  ul.plain{list-style:none;padding:0;margin:0}
  ul.plain li{padding:10px 0;border-bottom:1px solid var(--line)}
  ul.plain li:last-child{border-bottom:none}
  .muted{color:var(--muted);font-size:13px}
  h2{font-size:17px;margin:0 0 10px}
  .grid2{display:grid;grid-template-columns:1fr 1fr;gap:10px}
  .flash{background:#fff8e1;border:1px solid #ffe08a;padding:10px 12px;border-radius:8px;margin-bottom:12px;font-size:14px}
</style></head><body>
<header><span>OCS Master</span>
<a href="{{ url_for('dashboard') }}">Menu</a>
</header>
<main>{% if flash %}<div class="flash">{{ flash }}</div>{% endif %}{{ body|safe }}</main>
</body></html>
"""

def page(body_html, flash=None):
    return render_template_string(BASE, body=body_html, flash=flash,
                                   app_password_set=bool(APP_PASSWORD))

# =====================================================================
# LOGIN
# =====================================================================
@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        return redirect(url_for("dashboard"))
    err = None
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["authed"] = True
            return redirect(request.args.get("next") or url_for("dashboard"))
        err = "Wrong password."
    return page(f"""
    <div class="card">
      <h2>Sign in</h2>
      <form method="post">
        <label>App password</label>
        <input type="password" name="password" autofocus>
        <button type="submit">Enter</button>
      </form>
    </div>""", flash=err)

# =====================================================================
# DASHBOARD
# =====================================================================
@app.route("/")
def dashboard():
    return page(f"""
    <div class="card">
      <h2>Choose dashboard</h2>
      <a class="btn" href="{url_for('hr_view')}">HR call log / lookup</a>
      <a class="btn secondary" href="{url_for('mail_dashboard')}">Mail dashboard</a>
    </div>
    """)

@app.route("/mail")
@login_required
def mail_dashboard():
    if IITD_WEBMAIL_PASSWORD:
        pw_ready = "stored in Render env ✓"
    else:
        pw_ready = "cached ✓" if get_cached_pw() else "not entered"

    return page(f"""
    <div class="card">
      <h2>Mail dashboard</h2>
      <a class="btn" href="{url_for('continuous_view')}">Continuous pitch sender</a>
      <a class="btn" href="{url_for('single_view')}">Single company sender</a>
      <a class="btn secondary" href="{url_for('bounces_view')}">Check bounces</a>
      <a class="btn secondary" href="{url_for('reconcile_view')}">Reconcile Sent folder</a>
      <a class="btn secondary" href="{url_for('debug_webmail')}">Test IITD webmail login</a>
      <a class="btn secondary" href="{url_for('debug_webmail_send')}">Send one Roundcube test email</a>
    </div>
    <div class="card">
      <p class="muted">IITD mailbox password: {pw_ready}</p>
      <form method="post" action="{url_for('forget_password')}">
        <button type="submit" class="danger">Forget cached password</button>
      </form>
      <a class="btn secondary" href="{url_for('dashboard')}">Back to home</a>
    </div>
    """)

@app.route("/forget-password", methods=["POST"])
@login_required
def forget_password():
    clear_cached_pw()
    return redirect(url_for("dashboard"))

def password_field(label="IITD mailbox password"):
    if IITD_WEBMAIL_PASSWORD:
        return '<p class="muted">Using IITD mailbox password from Render environment.</p>'

    cached = get_cached_pw()
    if cached:
        return '<p class="muted">Using cached password (expires in up to 60 min).</p>'

    return f"""<label>{label}</label><input type="password" name="password" required>"""

def resolve_pw(form):
    pw = get_cached_pw()
    if pw:
        return pw
    pw = form.get("password", "")
    if pw:
        set_cached_pw(pw)
    return pw

# =====================================================================
# CONTINUOUS SENDER
# =====================================================================
@app.route("/continuous", methods=["GET"])
@login_required
def continuous_view():
    wb = get_wb()
    queue, _ = L.build_continuous_queue(wb)
    rows = "".join(
        f"""<li><input type="checkbox" name="idx" value="{i}" checked>
            &nbsp;<b>{it['company']}</b> <span class="tag">{it['sheet']}</span><br>
            <span class="muted">{', '.join(it['emails'])}</span></li>"""
        for i, it in enumerate(queue)
    )
    if not queue:
        body = '<div class="card"><h2>Continuous sender</h2><p>Nothing eligible to send right now — everyone in scope is already logged as sent.</p></div>'
        return page(body)
    body = f"""
    <div class="card">
      <h2>Continuous sender — {len(queue)} eligible ({L.PER_SHEET} per sheet per round)</h2>
      <form method="post" action="{url_for('continuous_send')}">
        <ul class="plain">{rows}</ul>
        <p class="muted">Brochure attachment is temporarily disabled for Roundcube sending. The email body includes the downloads link.</p>
        {password_field()}
        <button type="submit">Send selected</button>
      </form>
    </div>"""
    return page(body)

def _run_continuous_job(job_id, items, pw, brochure):
    def cb(result):
        _jobs[job_id]["results"].append(result)
    try:
        wb = get_wb()
        L.send_continuous_batch(wb, items, pw, brochure, progress_cb=cb)
    except Exception as e:
        _jobs[job_id]["results"].append({"company": "—", "status": "ERROR", "detail": str(e)})
    _jobs[job_id]["done"] = True

@app.route("/continuous/send", methods=["POST"])
@login_required
def continuous_send():
    wb = get_wb()
    queue, _ = L.build_continuous_queue(wb)
    picked_idx = {int(i) for i in request.form.getlist("idx")}
    items = [it for i, it in enumerate(queue) if i in picked_idx]
    pw = resolve_pw(request.form)
    if not pw or not items:
        return redirect(url_for("continuous_view"))
    brochure = BROCHURE_PATH if request.form.get("brochure") else None
    job_id = uuid.uuid4().hex
    _jobs[job_id] = {"done": False, "results": [], "total": len(items)}
    threading.Thread(target=_run_continuous_job, args=(job_id, items, pw, brochure), daemon=True).start()
    return redirect(url_for("continuous_status", job_id=job_id))

@app.route("/continuous/status/<job_id>")
@login_required
def continuous_status(job_id):
    job = _jobs.get(job_id)
    if not job:
        return redirect(url_for("continuous_view"))
    lines = "".join(
        f"""<li>{escape(r['company'])} — <span class="tag {'sent' if r['status']=='SENT' else 'fail'}">{r['status']}</span>
            {f"<br><span class='muted'>Error: {escape(r.get('detail', ''))}</span>" if r.get('detail') else ""}
            </li>"""
        for r in job["results"]
    )
    refresh = "" if job["done"] else '<meta http-equiv="refresh" content="3">'
    status_line = "Done." if job["done"] else f"Sending… ({len(job['results'])}/{job['total']}, ~{L.DELAY_SEC}s between each)"
    body = f"""{refresh}
    <div class="card">
      <h2>Continuous sender — progress</h2>
      <p class="muted">{status_line}</p>
      <ul class="plain">{lines or '<li class="muted">Starting…</li>'}</ul>
      {'<a class="btn secondary" href="' + url_for('mail_dashboard') + '">Back to mail dashboard</a>' if job['done'] else ''}
    </div>"""
    return page(body)

# =====================================================================
# SINGLE COMPANY SENDER
# =====================================================================
@app.route("/single", methods=["GET", "POST"])
@login_required
def single_view():
    results_html = ""
    query = request.values.get("q", "")
    if query:
        wb = get_wb()
        matches = L.search_company(wb, query)
        if not matches:
            results_html = '<p class="muted">No match.</p>'
        else:
            items = ""
            for m in matches:
                tag = f'<span class="tag {"fail" if m["already_sent"] else ""}">already {m["delivery"] or "sent"}</span>' if m["already_sent"] else ""
                items += f"""
                <li>
                  <b>{m['display']}</b> {tag}<br>
                  <span class="muted">{', '.join(m['sheets'])}</span><br>
                  <a class="btn secondary" href="{url_for('single_send_view', key=m['key'])}">Choose &amp; send</a>
                </li>"""
            results_html = f'<ul class="plain">{items}</ul>'
    body = f"""
    <div class="card">
      <h2>Single company sender</h2>
      <form method="get">
        <input type="text" name="q" placeholder="Company name (partial ok)" value="{query}" autofocus>
        <button type="submit">Search</button>
      </form>
      {results_html}
    </div>"""
    return page(body)

@app.route("/single/send/<key>", methods=["GET", "POST"])
@login_required
def single_send_view(key):
    wb = get_wb()
    matches = {m["key"]: m for m in L.search_company(wb, key)}
    m = matches.get(key)
    if not m:
        # key search may not resolve directly; fall back to a direct lookup by re-searching companies
        companies, _, _ = L.build_company_data(wb)
        if key in companies:
            c = companies[key]
            m = {"key": key, "display": c["display"], "emails": c["emails"], "sheets": sorted(c["sheets"])}
        else:
            return redirect(url_for("single_view"))

    if request.method == "POST":
        chosen = request.form.getlist("email")
        custom = [e.strip() for e in request.form.get("custom_emails", "").split(",") if e.strip()]
        recipients = L.unique_emails(chosen + custom)
        if not recipients:
            return page('<div class="card">No recipient selected. <a href="javascript:history.back()">Back</a></div>')
        pw = resolve_pw(request.form)
        if not pw:
            return page('<div class="card">Password required. <a href="javascript:history.back()">Back</a></div>')
        brochure = BROCHURE_PATH if request.form.get("brochure") else None
        result = L.send_single(wb, key, m["display"], recipients, brochure, pw)
        status = result["status"]
        detail = result.get("detail", "")
        
        body = f"""<div class="card"><h2>{m['display']}</h2>
          <p><span class="tag {'sent' if status=='SENT' else 'fail'}">{status}</span></p>
          <p class="muted">To: {', '.join(recipients)}</p>
          {f'<p class="muted">Error: {escape(detail)}</p>' if detail else ''}
          <a class="btn" href="{url_for('single_view')}">Back to search</a></div>"""
        return page(body)

    checks = "".join(
        f'<label><input type="checkbox" name="email" value="{e}" checked> {e}</label>'
        for e in m["emails"]
    ) or '<p class="muted">No emails on file — add one below.</p>'
    body = f"""
    <div class="card">
      <h2>{m['display']}</h2>
      <form method="post">
        {checks}
        <label>Extra email(s), comma-separated</label>
        <input type="text" name="custom_emails" placeholder="name@company.com">
        <p class="muted">Brochure attachment is temporarily disabled for Roundcube sending. The email body includes the downloads link.</p>
        {password_field()}
        <button type="submit">Send</button>
      </form>
    </div>"""
    return page(body)

# =====================================================================
# CHECK BOUNCES
# =====================================================================
@app.route("/bounces", methods=["GET", "POST"])
@login_required
def bounces_view():
    if request.method == "POST":
        pw = resolve_pw(request.form)
        days = int(request.form.get("days") or 14)
        wb = get_wb()
        res = L.check_bounces(wb, pw, days=days)
        rows = "".join(f"<li>{r['company']} — <b>{r['status']}</b><br><span class='muted'>{', '.join(r['bounced'])}</span></li>"
                        for r in res["updated_rows"])
        body = f"""<div class="card"><h2>Bounce check result</h2>
          <p class="muted">Scanned {res['scanned']} inbox messages, last {days} days.</p>
          <p>{len(res['bounced'])} bounced address(es) found, {len(res['updated_rows'])} row(s) updated.</p>
          <ul class="plain">{rows}</ul>
          <a class="btn" href="{url_for('mail_dashboard')}">Back to mail dashboard</a></div>"""
        return page(body)
    body = f"""<div class="card"><h2>Check bounces</h2>
      <form method="post">
        <label>Days of Inbox to scan</label>
        <input type="number" name="days" value="14">
        {password_field()}
        <button type="submit">Run</button>
      </form></div>"""
    return page(body)

# =====================================================================
# RECONCILE
# =====================================================================
@app.route("/reconcile", methods=["GET", "POST"])
@login_required
def reconcile_view():
    if request.method == "POST":
        pw = resolve_pw(request.form)
        days = int(request.form.get("days") or 30)
        wb = get_wb()
        res = L.reconcile_sent(wb, pw, days=days)
        body = f"""<div class="card"><h2>Reconcile result</h2>
          <p class="muted">Scanned {res['scanned']} Sent messages, last {days} days.</p>
          <p>{res['matched']} pitch mail(s) matched · {res['added']} new row(s) added · {res['updated']} row(s) marked SENT.</p>
          <p class="muted">{len(res['unattributed'])} could not be attributed automatically.</p>
          <a class="btn" href="{url_for('mail_dashboard')}">Back to mail dashboard</a></div>"""
        return page(body)
    body = f"""<div class="card"><h2>Reconcile Sent folder</h2>
      <form method="post">
        <label>Days of Sent folder to scan</label>
        <input type="number" name="days" value="30">
        {password_field()}
        <button type="submit">Run</button>
      </form></div>"""
    return page(body)

# =====================================================================
# HR CONTACT / PHONE LOOKUP  +  CALL LOGS
# =====================================================================
@app.route("/hr", methods=["GET"])
def hr_view():
    query = request.args.get("q", "")
    results_html = ""
    if query:
        wb = get_wb()
        matches = L.hr_lookup(wb, query)
        if not matches:
            results_html = '<p class="muted">No match.</p>'
        else:
            items = ""
            for m in matches:
                items += f"""<li><b>{m['display']}</b>
                  <span class="muted">[{', '.join(m['sheets'])}] · {len(m['contacts'])} contact(s)</span><br>
                  <a class="btn secondary" href="{url_for('hr_company_view', key=m['key'], q=query)}">View / log calls</a></li>"""
            results_html = f'<ul class="plain">{items}</ul>'
    body = f"""
    <div class="card">
      <h2>HR contact / phone lookup</h2>
      <form method="get">
        <input type="text" name="q" placeholder="Company name (partial ok)" value="{query}" autofocus>
        <button type="submit">Search</button>
      </form>
      <a class="btn secondary" href="{url_for('manual_call_log_view', company=query)}">Log call for new / unlisted company</a>
      {results_html}
    </div>"""
    return page(body)

@app.route("/hr/manual", methods=["GET", "POST"])
def manual_call_log_view():
    company_prefill = request.values.get("company", "").strip()

    if request.method == "POST":
        company = request.form.get("company", "").strip()
        phone_raw = request.form.get("phone", "").strip()
        incident = request.form.get("incident", "").strip()
        caller_name = request.form.get("caller_name", "").strip()
        hr_name = request.form.get("hr_name", "").strip()
        hr_email = request.form.get("hr_email", "").strip()

        if not company:
            return page(f"""
            <div class="card">
              <h2>Nothing logged</h2>
              <p>Company name is required.</p>
              <a class="btn" href="{url_for('manual_call_log_view', company=company_prefill)}">Back</a>
            </div>
            """)

        if not phone_raw and not hr_email:
            return page(f"""
            <div class="card">
              <h2>Nothing logged</h2>
              <p>Please enter at least a phone number or HR email.</p>
              <a class="btn" href="{url_for('manual_call_log_view', company=company)}">Back</a>
            </div>
            """)

        wb = get_wb()

        entry = {
            "company": company,
            "phone": phone_raw,
            "incident": incident,
            "date": time.strftime("%Y-%m-%d"),
            "caller_name": caller_name,
            "hr_name": hr_name,
            "hr_email": hr_email,
        }

        L.append_call_logs(wb, [entry])

        return page(f"""
        <div class="card">
          <h2>Call logged</h2>
          <p>Logged call/contact info for {company}.</p>
          <a class="btn" href="{url_for('hr_view', q=company)}">View company in HR lookup</a>
          <a class="btn secondary" href="{url_for('manual_call_log_view')}">Log another new company</a>
        </div>
        """)

    body = f"""
    <div class="card">
      <h2>Log call for new / unlisted company</h2>
      <form method="post">
        <label>Company</label>
        <input type="text" name="company" value="{escape(company_prefill)}" placeholder="e.g. Hyundai" required>

        <label>Phone Number</label>
        <input type="text" name="phone" placeholder="e.g. +91 9876543210">

        <label>Incident</label>
        <input type="text" name="incident" placeholder="e.g. spoke to HR, asked to email brochure">

        <label>Caller Name</label>
        <input type="text" name="caller_name" placeholder="e.g. Aman">

        <label>HR Name</label>
        <input type="text" name="hr_name" placeholder="e.g. Priya Sharma">

        <label>HR Email</label>
        <input type="text" name="hr_email" placeholder="e.g. priya@company.com">

        <button type="submit">Save call log</button>
      </form>
      <a class="btn secondary" href="{url_for('hr_view')}">Back to HR lookup</a>
    </div>
    """

    return page(body)

@app.route("/hr/<key>", methods=["GET", "POST"])
def hr_company_view(key):
    wb = get_wb()
    q = request.values.get("q", key)
    matches = {m["key"]: m for m in L.hr_lookup(wb, q)}
    m = matches.get(key)
    if not m:
        # direct key fallback (e.g. bookmarked link with a different query)
        matches_all = {mm["key"]: mm for mm in L.hr_lookup(wb, key)}
        m = matches_all.get(key)
    if not m:
        return redirect(url_for("hr_view"))

    if request.method == "POST":
        phone_raw = request.form.get("phone", "").strip()
        incident = request.form.get("incident", "").strip()
        caller_name = request.form.get("caller_name", "").strip()
        hr_name = request.form.get("hr_name", "").strip()
        hr_email = request.form.get("hr_email", "").strip()

        if not phone_raw:
            return page(f"""<div class="card"><h2>Nothing logged</h2><p>No phone number was entered.</p><a class="btn" href="{url_for('hr_company_view', key=key, q=q)}">Back</a></div>""")
    
        entry = {
            "company": m["display"],
            "phone": phone_raw,
            "incident": incident,
            "date": time.strftime("%Y-%m-%d"),
            "caller_name": caller_name,
            "hr_name": hr_name,
            "hr_email": hr_email,
        }

        L.append_call_logs(wb, [entry])

        return page(f"""<div class="card"><h2>Call logged</h2><p>Logged call to {phone_raw} for {m['display']}.</p><a class="btn" href="{url_for('hr_company_view', key=key, q=q)}">Back</a></div>""")

    rows = "".join(
        f"""<tr><td>{c['sheet']}</td><td>{c['name'] or '-'}</td>
            <td>{c['phone'] or '-'}</td><td>{c['email'] or '-'}</td></tr>"""
        for c in m["contacts"]
    ) or '<tr><td colspan="4" class="muted">No HR name/email/phone on file.</td></tr>'

    phone_options = "".join(f'<option value="{escape(c["phone"])}" data-name="{escape(c["name"])}" data-email="{escape(c["email"])}">{escape(c["name"] or c["phone"])} ({escape(c["phone"])})</option>'for c in m["contacts"] if c["phone"])

    body = f"""
    <div class="card">
      <h2>{m['display']}</h2>
      <p class="muted">In: {', '.join(m['sheets'])}</p>
      <table style="width:100%;border-collapse:collapse;font-size:14px">
        <tr class="muted"><th align="left">Sheet</th><th align="left">Name</th><th align="left">Phone</th><th align="left">Email</th></tr>
        {rows}
      </table>
    </div>
    <div class="card">
      <h2>Log a call</h2>
      <form method="post">
        <label>Pick an existing contact's number</label>
        <select name="phone_pick" onchange="
          document.getElementsByName('phone')[0].value=this.value;
          document.getElementsByName('hr_name')[0].value=this.options[this.selectedIndex].dataset.name || '';
          document.getElementsByName('hr_email')[0].value=this.options[this.selectedIndex].dataset.email || '';
        ">
           <option value="">— or type a number below —</option>
           {phone_options}
         </select>

        <label>Phone Number</label>
        <input type="text" name="phone" placeholder="Type or pick above">

        <label>Incident</label>
        <input type="text" name="incident" placeholder="e.g. spoke to HR, follow up next week">

        <label>Caller Name</label>
        <input type="text" name="caller_name" placeholder="e.g. Aman">

        <label>HR Name</label>
        <input type="text" name="hr_name" placeholder="e.g. Priya Sharma">

        <label>HR Email</label>
        <input type="text" name="hr_email" placeholder="e.g. priya@company.com">

        <button type="submit">Save call log</button>
      </form>
      <a class="btn secondary" href="{url_for('hr_view', q=q)}">Back to search</a>
    </div>"""
    return page(body)

@app.route("/debug-webmail-send", methods=["GET", "POST"])
@login_required
def debug_webmail_send():
    if request.method == "POST":
        password = resolve_pw(request.form)
        recipients = L.emails_in(request.form.get("recipient", ""))

        if not password or not recipients:
            return page("""
            <div class="card">
              Enter your IITD mailbox password and one valid recipient email.
              <a href="javascript:history.back()">Back</a>
            </div>
            """)

        try:
            L.send_one_via_roundcube(
                password=password,
                recipients=recipients[:1],
                company="Roundcube delivery test",
                subject="IITD Roundcube delivery test",
                body=(
                    "This is a one-message delivery test from the IITD "
                    "outreach application through IITD Roundcube webmail."
                ),
            )
            status = '<span class="tag sent">TEST EMAIL SENT</span>'
            message = f"Sent to {escape(recipients[0])}."
        except L.WebmailSendError as e:
            status = '<span class="tag fail">SEND FAILED</span>'
            message = escape(str(e))

        return page(f"""
        <div class="card">
          <h2>Roundcube test send</h2>
          <p>{status}</p>
          <p class="muted">{message}</p>
          <a class="btn secondary" href="{url_for('mail_dashboard')}">
            Back to mail dashboard
          </a>
        </div>
        """)

    return page(f"""
    <div class="card">
      <h2>Send one Roundcube test email</h2>
      <p class="muted">
        Use only your own IITD email address for this test.
      </p>
      <form method="post">
        <label>Recipient</label>
        <input type="text" name="recipient" placeholder="yourname@iitd.ac.in" required>
        {password_field()}
        <button type="submit">Send one test email</button>
      </form>
    </div>
    """)

@app.route("/debug-webmail", methods=["GET", "POST"])
@login_required
def debug_webmail():
    if request.method == "POST":
        password = resolve_pw(request.form)

        if not password:
            return page(
                '<div class="card">Password required. '
                '<a href="javascript:history.back()">Back</a></div>'
            )

        try:
            roundcube_session, _ = L.roundcube_compose_session(password)
            roundcube_session.close()
            status = '<span class="tag sent">LOGIN WORKED</span>'
            message = (
                "Render successfully logged in to IITD Roundcube and opened a compose session over HTTPS. No email was sent."
            )
        except L.WebmailLoginError as e:
            status = '<span class="tag fail">LOGIN FAILED</span>'
            message = escape(str(e))

        return page(f"""
        <div class="card">
          <h2>IITD webmail connection</h2>
          <p>{status}</p>
          <p class="muted">{message}</p>
          <a class="btn secondary" href="{url_for('mail_dashboard')}">
            Back to mail dashboard
          </a>
        </div>
        """)

    return page(f"""
    <div class="card">
      <h2>Test IITD webmail login</h2>
      <p class="muted">
        This only signs in through Roundcube. It does not send an email.
      </p>
      <form method="post">
        {password_field()}
        <button type="submit">Test login</button>
      </form>
    </div>
    """)
    
@app.route("/debug-smtp")
@login_required
def debug_smtp():
    import socket
    import time

    tests = [
        ("smtp.iitd.ac.in", 465),
        ("smtp.iitd.ac.in", 587),
        
        ("mailstore.iitd.ac.in", 465),
        ("mailstore.iitd.ac.in", 587),
        ("mailstore.iitd.ac.in", 25),
        ("mailstore.iitd.ac.in", 993),
    
        ("webmail.iitd.ac.in", 465),
        ("webmail.iitd.ac.in", 587),
        ("webmail.iitd.ac.in", 25),
        ("webmail.iitd.ac.in", 443),
    
        ("mail.iitd.ac.in", 465),
        ("mail.iitd.ac.in", 587),
        ("mail.iitd.ac.in", 25),

    ]

    rows = ""

    for host, port in tests:
        start = time.time()
        try:
            s = socket.create_connection((host, port), timeout=10)
            s.close()
            rows += f"<li><b>{host}:{port}</b> — CONNECTED in {time.time() - start:.2f}s</li>"
        except Exception as e:
            rows += f"<li><b>{host}:{port}</b> — FAILED: {str(e)}</li>"

    return page(f"""
    <div class="card">
      <h2>SMTP / IMAP connectivity debug</h2>
      <ul class="plain">{rows}</ul>
      <a class="btn secondary" href="{url_for('mail_dashboard') if 'mail_dashboard' in globals() else url_for('dashboard')}">Back</a>
    </div>
    """)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
