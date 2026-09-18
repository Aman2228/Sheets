"""
app.py — OCS Master, web version.

Open the same URL from your phone or PC — both use the same Google Sheet.

Security model:
- Main menu and non-mail dashboards are public.
- APP_PASSWORD protects the Mail Dashboard and its private mail tools.
- SINGLE_SENDER_PASSWORD separately gates the public single-company sender.
- IITD mailbox password can be supplied through IITD_WEBMAIL_PASSWORD or
  entered in the browser and cached only in server memory for 60 minutes.
- IITD mailbox passwords are never written to Google Sheets, disk, or logs.

Required environment variables:
    SPREADSHEET_KEY
    APP_PASSWORD
    FLASK_SECRET_KEY

Optional:
    IITD_WEBMAIL_PASSWORD
    SINGLE_SENDER_PASSWORD
    BROCHURE_FILE_PATH
    PORT
"""

import os
import time
import uuid
import threading
from html import escape
from functools import wraps

from flask import (
    Flask,
    request,
    redirect,
    url_for,
    session,
    render_template_string,
)

import ocs_logic as L
from sheets_backend import load_workbook


# =====================================================================
# APP CONFIG
# =====================================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "FLASK_SECRET_KEY",
    "dev-key-change-me",
)

SPREADSHEET_KEY = os.environ.get("SPREADSHEET_KEY")
APP_PASSWORD = os.environ.get("APP_PASSWORD")
IITD_WEBMAIL_PASSWORD = os.environ.get("IITD_WEBMAIL_PASSWORD")
SINGLE_SENDER_PASSWORD = os.environ.get("SINGLE_SENDER_PASSWORD")
BROCHURE_PATH = os.environ.get("BROCHURE_FILE_PATH")

PW_TTL = 60 * 60  # 60 minutes


# =====================================================================
# IN-MEMORY PASSWORD CACHE
# =====================================================================

# {
#     session_id: (password, expires_at)
# }
_pw_cache = {}

_pw_cache_lock = threading.Lock()


def _sid():
    """Return/create a random identifier for the current browser session."""
    if "sid" not in session:
        session["sid"] = uuid.uuid4().hex

    return session["sid"]


def get_cached_pw():
    """
    Return the mailbox password currently available to this session.

    Priority:
    1. IITD_WEBMAIL_PASSWORD environment variable.
    2. Temporary in-memory browser-session cache.
    """

    if IITD_WEBMAIL_PASSWORD:
        return IITD_WEBMAIL_PASSWORD

    sid = _sid()

    with _pw_cache_lock:
        entry = _pw_cache.get(sid)

        if not entry:
            return None

        password, expires_at = entry

        if time.time() >= expires_at:
            _pw_cache.pop(sid, None)
            return None

        return password


def set_cached_pw(password):
    """Cache mailbox password in server RAM for 60 minutes."""

    if not password:
        return

    sid = _sid()

    with _pw_cache_lock:
        _pw_cache[sid] = (
            password,
            time.time() + PW_TTL,
        )


def clear_cached_pw():
    """Forget the current browser session's cached mailbox password."""

    sid = _sid()

    with _pw_cache_lock:
        _pw_cache.pop(sid, None)


def resolve_pw(form):
    """
    Resolve mailbox password.

    Uses an existing cached/environment password first.
    Otherwise reads the submitted password and caches it temporarily.
    """

    password = get_cached_pw()

    if password:
        return password

    password = (form.get("password") or "").strip()

    if password:
        set_cached_pw(password)

    return password


# =====================================================================
# BACKGROUND JOB STORE
# =====================================================================

# {
#     job_id: {
#         "done": bool,
#         "results": list,
#         "total": int
#     }
# }

_jobs = {}
_jobs_lock = threading.Lock()


# =====================================================================
# AUTH DECORATORS
# =====================================================================

def login_required(func):
    """Legacy decorator retained for compatibility.

    The main menu and non-mail dashboards are intentionally not protected
    by APP_PASSWORD. Mail routes use mail_login_required instead.
    """

    @wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)

    return wrapper

def mail_login_required(func):
    """Require APP_PASSWORD only for the Mail Dashboard."""

    @wraps(func)
    def wrapper(*args, **kwargs):

        if APP_PASSWORD and not session.get("mail_authed"):
            next_url = request.full_path

            return redirect(
                url_for(
                    "mail_login",
                    next=next_url,
                )
            )

        return func(*args, **kwargs)

    return wrapper
def single_sender_required(func):
    """
    Require SINGLE_SENDER_PASSWORD when configured.

    This is intentionally separate from APP_PASSWORD so the
    single-company sender can be exposed independently.
    """

    @wraps(func)
    def wrapper(*args, **kwargs):

        if (
            SINGLE_SENDER_PASSWORD
            and not session.get("single_sender_authed")
        ):
            next_url = request.full_path

            return redirect(
                url_for(
                    "single_sender_login",
                    next=next_url,
                )
            )

        return func(*args, **kwargs)

    return wrapper


# =====================================================================
# GOOGLE SHEET HELPERS
# =====================================================================

def get_wb():
    """Load the configured Google Sheet workbook."""

    if not SPREADSHEET_KEY:
        raise RuntimeError(
            "SPREADSHEET_KEY environment variable is not set."
        )

    return load_workbook(SPREADSHEET_KEY)


def get_sheet_case_insensitive(wb, wanted_name):
    """Find a worksheet regardless of capitalization/extra spaces."""

    wanted = str(wanted_name).strip().lower()

    for name in wb.sheetnames:
        if str(name).strip().lower() == wanted:
            return wb[name]

    return None


def _sheet_col(ws, possible_headers):
    """
    Find a worksheet column by trying multiple possible header names.

    Matching:
    - case-insensitive
    - ignores surrounding/excess spaces through L.clean()
    """

    wanted = {
        L.clean(header).strip().lower()
        for header in possible_headers
    }

    for column in range(1, ws.max_column + 1):
        value = L.clean(
            ws.cell(row=1, column=column).value
        ).strip().lower()

        if value in wanted:
            return column

    return None


def _cell(ws, row, col):
    """Safely return a cleaned worksheet cell."""

    if not col:
        return ""

    return L.clean(
        ws.cell(row=row, column=col).value
    )


def _render_simple_table(headers, rows):
    """Render a mobile-friendly HTML table."""

    head_html = "".join(
        f"<th align='left'>{escape(str(header))}</th>"
        for header in headers
    )

    if not rows:
        return "<p class='muted'>No rows found.</p>"

    body_html = ""

    for row in rows:
        body_html += "<tr>"

        for value in row:
            text = "" if value is None else str(value)

            body_html += (
                f"<td>{escape(text)}</td>"
            )

        body_html += "</tr>"

    return f"""
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:14px">
        <thead>
          <tr class="muted">
            {head_html}
          </tr>
        </thead>
        <tbody>
          {body_html}
        </tbody>
      </table>
    </div>
    """


def _safe_int(value, default):
    """Convert a form value to int without crashing on bad input."""

    try:
        number = int(str(value).strip())

        if number < 1:
            return default

        return number

    except (TypeError, ValueError):
        return default


# =====================================================================
# HTML LAYOUT
# =====================================================================

BASE = """
<!doctype html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>OCS Master</title>

<style>
  :root{
    --fg:#1a1a1a;
    --muted:#6b6b6b;
    --line:#e3e3e3;
    --accent:#2a5bd7;
    --bad:#c0392b;
    --good:#1e8449;
    --bg:#fafafa;
  }

  *{
    box-sizing:border-box;
  }

  body{
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
    margin:0;
    background:var(--bg);
    color:var(--fg);
  }

  header{
    background:#111;
    color:#fff;
    padding:14px 16px;
    font-weight:600;
    display:flex;
    justify-content:space-between;
    align-items:center;
  }

  header a{
    color:#fff;
    text-decoration:none;
    font-size:14px;
    opacity:.85;
  }

  main{
    max-width:700px;
    margin:0 auto;
    padding:16px;
  }

  .card{
    background:#fff;
    border:1px solid var(--line);
    border-radius:10px;
    padding:16px;
    margin-bottom:14px;
  }

  a.btn,
  button{
    display:inline-block;
    background:var(--accent);
    color:#fff;
    border:none;
    border-radius:8px;
    padding:12px 16px;
    font-size:16px;
    text-decoration:none;
    cursor:pointer;
    width:100%;
    margin-top:8px;
  }

  button.secondary,
  a.btn.secondary{
    background:#fff;
    color:var(--accent);
    border:1px solid var(--accent);
  }

  button.danger,
  a.btn.danger{
    background:var(--bad);
    color:#fff;
    border:none;
  }

  input[type=text],
  input[type=password],
  input[type=number],
  select{
    width:100%;
    padding:11px;
    border:1px solid var(--line);
    border-radius:8px;
    font-size:16px;
    margin-top:4px;
    background:#fff;
  }

  input[type=checkbox]{
    transform:scale(1.15);
    margin-right:6px;
  }

  label{
    font-size:13px;
    color:var(--muted);
    display:block;
    margin-top:10px;
  }

  .row{
    display:flex;
    gap:8px;
  }

  .row>*{
    flex:1;
  }

  .tag{
    display:inline-block;
    font-size:12px;
    padding:2px 8px;
    border-radius:12px;
    background:#eee;
    margin-left:6px;
  }

  .tag.sent{
    background:#e8f6ee;
    color:var(--good);
  }

  .tag.fail{
    background:#fdecea;
    color:var(--bad);
  }

  ul.plain{
    list-style:none;
    padding:0;
    margin:0;
  }

  ul.plain li{
    padding:10px 0;
    border-bottom:1px solid var(--line);
  }

  ul.plain li:last-child{
    border-bottom:none;
  }

  .muted{
    color:var(--muted);
    font-size:13px;
  }

  h2{
    font-size:17px;
    margin:0 0 10px;
  }

  .flash{
    background:#fff8e1;
    border:1px solid #ffe08a;
    padding:10px 12px;
    border-radius:8px;
    margin-bottom:12px;
    font-size:14px;
  }

  table th,
  table td{
    border-bottom:1px solid var(--line);
    padding:8px 6px;
    vertical-align:top;
  }

  table th{
    font-weight:600;
  }

  details.card{
    padding:0;
  }

  details.card summary{
    list-style:none;
    cursor:pointer;
    padding:16px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:10px;
  }

  details.card summary::-webkit-details-marker{
    display:none;
  }

  details.card summary::after{
    content:"Open";
    font-size:12px;
    color:var(--accent);
    border:1px solid var(--accent);
    border-radius:999px;
    padding:3px 8px;
  }

  details.card[open] summary::after{
    content:"Close";
  }

  details.card > div{
    padding:0 16px 16px;
  }

  .error-box{
    background:#fdecea;
    border:1px solid #f5c6c2;
    color:#7b241c;
    padding:12px;
    border-radius:8px;
    margin-bottom:12px;
  }

  .success-box{
    background:#e8f6ee;
    border:1px solid #b7dfc6;
    color:#145a32;
    padding:12px;
    border-radius:8px;
    margin-bottom:12px;
  }
</style>

</head>

<body>

<header>
  <span>OCS Master</span>

  <a href="{{ url_for('dashboard') }}">
    Menu
  </a>
</header>

<main>

{% if flash %}
  <div class="flash">
    {{ flash }}
  </div>
{% endif %}

{{ body|safe }}

</main>

</body>
</html>
"""


def page(body_html, flash=None):
    """Render a page using the common layout."""

    return render_template_string(
        BASE,
        body=body_html,
        flash=flash,
    )


# =====================================================================
# LOGIN
# =====================================================================

@app.route("/login", methods=["GET", "POST"])
def login():

    if not APP_PASSWORD:
        return redirect(
            url_for("dashboard")
        )

    error = None

    if request.method == "POST":

        password = request.form.get("password", "")

        if password == APP_PASSWORD:
            session["authed"] = True

            next_url = request.args.get("next")

            if next_url and next_url.startswith("/"):
                return redirect(next_url)

            return redirect(
                url_for("dashboard")
            )

        error = "Wrong password."

    return page(
        """
        <div class="card">
          <h2>Sign in</h2>

          <form method="post">

            <label>App password</label>

            <input
              type="password"
              name="password"
              autofocus
              required
            >

            <button type="submit">
              Enter
            </button>

          </form>
        </div>
        """,
        flash=error,
    )


# =====================================================================
# SINGLE SENDER LOGIN
# =====================================================================

@app.route("/single-login", methods=["GET", "POST"])
def single_sender_login():

    if not SINGLE_SENDER_PASSWORD:
        return redirect(
            url_for("single_public_view")
        )

    error = None

    if request.method == "POST":

        password = request.form.get("password", "")

        if password == SINGLE_SENDER_PASSWORD:

            session["single_sender_authed"] = True

            next_url = request.args.get("next")

            if next_url and next_url.startswith("/"):
                return redirect(next_url)

            return redirect(
                url_for("single_public_view")
            )

        error = "Wrong single sender password."

    return page(
        """
        <div class="card">

          <h2>Single company sender access</h2>

          <form method="post">

            <label>Single sender password</label>

            <input
              type="password"
              name="password"
              autofocus
              required
            >

            <button type="submit">
              Enter
            </button>

          </form>

        </div>
        """,
        flash=error,
    )

@app.route("/mail-login", methods=["GET", "POST"])
def mail_login():

    if not APP_PASSWORD:
        return redirect(
            url_for("mail_dashboard")
        )

    error = None

    if request.method == "POST":

        password = request.form.get(
            "password",
            ""
        )

        if password == APP_PASSWORD:

            session["mail_authed"] = True

            next_url = request.args.get("next")

            if next_url and next_url.startswith("/"):
                return redirect(next_url)

            return redirect(
                url_for("mail_dashboard")
            )

        error = "Wrong password."

    return page(
        """
        <div class="card">

          <h2>Mail Dashboard Access</h2>

          <p class="muted">
            Enter the app password to access
            the mail dashboard.
          </p>

          <form method="post">

            <label>Mail dashboard password</label>

            <input
              type="password"
              name="password"
              autofocus
              required
            >

            <button type="submit">
              Enter
            </button>

          </form>

        </div>
        """,
        flash=error,
    )
# =====================================================================
# DASHBOARD
# =====================================================================

@app.route("/")
def dashboard():

    return page(
        f"""
        <div class="card">

          <h2>Choose dashboard</h2>

          <a class="btn"
             href="{url_for('hr_view')}">
            HR call log / lookup
          </a>

          <a class="btn secondary"
             href="{url_for('sheet_data_view')}">
            View sheet data
          </a>

          <a class="btn secondary"
             href="{url_for('single_public_view')}">
            Single company sender
          </a>

          <a class="btn secondary"
             href="{url_for('mail_dashboard')}">
            Mail dashboard
          </a>

        </div>
        """
    )


# =====================================================================
# MAIL DASHBOARD
# =====================================================================

@app.route("/mail")
@mail_login_required
def mail_dashboard():

    if IITD_WEBMAIL_PASSWORD:
        password_status = "stored in Render environment ✓"

    elif get_cached_pw():
        password_status = "cached ✓"

    else:
        password_status = "not entered"

    return page(
        f"""
        <div class="card">

          <h2>Mail dashboard</h2>

          <a class="btn"
             href="{url_for('continuous_view')}">
            Continuous pitch sender
          </a>

          <a class="btn"
             href="{url_for('single_view')}">
            Single company sender
          </a>

          <a class="btn secondary"
             href="{url_for('bounces_view')}">
            Check bounces
          </a>

          <a class="btn secondary"
             href="{url_for('reconcile_view')}">
            Reconcile Sent folder
          </a>

          <a class="btn secondary"
             href="{url_for('debug_webmail')}">
            Test IITD webmail login
          </a>

          <a class="btn secondary"
             href="{url_for('debug_webmail_send')}">
            Send one Roundcube test email
          </a>

          <a class="btn secondary"
             href="{url_for('debug_smtp')}">
            SMTP / IMAP connectivity test
          </a>

        </div>

        <div class="card">

          <p class="muted">
            IITD mailbox password:
            {escape(password_status)}
          </p>

          <form
            method="post"
            action="{url_for('forget_password')}"
          >

            <button
              type="submit"
              class="danger"
            >
              Forget cached password
            </button>

          </form>

          <a
            class="btn secondary"
            href="{url_for('dashboard')}"
          >
            Back to home
          </a>

        </div>
        """
    )


@app.route("/forget-password", methods=["POST"])
@mail_login_required
def forget_password():

    clear_cached_pw()

    return redirect(
        url_for("dashboard")
    )


# =====================================================================
# MAIL PASSWORD FIELD
# =====================================================================

def password_field(label="IITD mailbox password"):

    if IITD_WEBMAIL_PASSWORD:
        return """
        <p class="muted">
          Using IITD mailbox password from Render environment.
        </p>
        """

    if get_cached_pw():
        return """
        <p class="muted">
          Using cached password.
          It expires after 60 minutes.
        </p>
        """

    return f"""
    <label>
      {escape(label)}
    </label>

    <input
      type="password"
      name="password"
      required
    >
    """


# =====================================================================
# SHEET DATA VIEWER
# =====================================================================

@app.route("/data")
def sheet_data_view():

    try:
        wb = get_wb()
    except Exception as exc:
        return page(
            f"""
            <div class="card">
              <h2>Sheet error</h2>
              <div class="error-box">
                {escape(str(exc))}
              </div>
              <a class="btn secondary"
                 href="{url_for('dashboard')}">
                Back to home
              </a>
            </div>
            """
        )

    tabs = [
        "Design",
        "Thermal",
        "Industrial",
        "Production",
    ]

    cards = ""

    for sheet_name in tabs:

        ws = get_sheet_case_insensitive(
            wb,
            sheet_name,
        )

        if ws is None:

            cards += f"""
            <details class="card">

              <summary>
                <b>{escape(sheet_name)}</b>
                <span class="tag fail">
                  not found
                </span>
              </summary>

              <div>
                <p class="muted">
                  Sheet not found.
                </p>
              </div>

            </details>
            """

            continue

        company_col = _sheet_col(
            ws,
            [
                "Company Name",
                "Company",
                "Organisation",
                "Organization",
            ],
        )

        hr_name_col = _sheet_col(
            ws,
            [
                "HR Name",
                "Name",
                "Contact Person",
                "Contact Name",
            ],
        )

        hr_phone_col = _sheet_col(
            ws,
            [
                "HR Phone",
                "HR Phone Number",
                "Phone Number",
                "Phone",
                "Contact Number",
                "Contact No.",
            ],
        )

        hr_email_col = _sheet_col(
            ws,
            [
                "HR Email",
                "Email",
                "Email ID",
                "Email Address",
                "Contact Email",
            ],
        )

        rows = []
        current_company = ""

        for row_number in range(
            2,
            ws.max_row + 1,
        ):

            company = _cell(
                ws,
                row_number,
                company_col,
            )

            hr_name = _cell(
                ws,
                row_number,
                hr_name_col,
            )

            hr_phone = _cell(
                ws,
                row_number,
                hr_phone_col,
            )

            hr_email = _cell(
                ws,
                row_number,
                hr_email_col,
            )

            if company:
                current_company = company

            else:
                company = current_company

            if not (
                company
                or hr_name
                or hr_phone
                or hr_email
            ):
                continue

            rows.append(
                [
                    company,
                    hr_name,
                    hr_phone,
                    hr_email,
                ]
            )

        cards += f"""
        <details class="card">

          <summary>

            <b>{escape(sheet_name)}</b>

            <span class="tag">
              {len(rows)} row(s)
            </span>

          </summary>

          <div>

            {_render_simple_table(
                [
                    "Company",
                    "HR Name",
                    "HR Phone Number",
                    "HR Email",
                ],
                rows,
            )}

          </div>

        </details>
        """

    # ---------------------------------------------------------------
    # CALL LOGS
    # ---------------------------------------------------------------

    ws = get_sheet_case_insensitive(
        wb,
        "Call Logs",
    )

    if ws is not None:

        columns = {}

        wanted_headers = [
            (
                "Company",
                [
                    "Company",
                    "Company Name",
                ],
            ),
            (
                "Phone Number",
                [
                    "Phone Number",
                    "Phone",
                    "HR Phone",
                    "Contact Number",
                ],
            ),
            (
                "Incident",
                [
                    "Incident",
                    "Status",
                    "Notes",
                    "Remarks",
                ],
            ),
            (
                "Date",
                [
                    "Date",
                    "Call Date",
                    "Timestamp",
                    "Date/Time",
                ],
            ),
            (
                "Caller Name",
                [
                    "Caller Name",
                    "Caller",
                    "Called By",
                ],
            ),
            (
                "HR Name",
                [
                    "HR Name",
                    "HRName",
                    "Contact Person",
                ],
            ),
            (
                "HR Email",
                [
                    "HR Email",
                    "HREmail",
                    "Email",
                    "Email ID",
                ],
            ),
            (
                "Success Flag",
                [
                    "Success Flag",
                    "Success",
                    "Successful",
                    "Call Success",
                ],
            ),
        ]

        for display_name, aliases in wanted_headers:
            columns[display_name] = _sheet_col(
                ws,
                aliases,
            )

        call_rows = []

        for row_number in range(
            2,
            ws.max_row + 1,
        ):

            row = [
                _cell(
                    ws,
                    row_number,
                    columns["Company"],
                ),
                _cell(
                    ws,
                    row_number,
                    columns["Phone Number"],
                ),
                _cell(
                    ws,
                    row_number,
                    columns["Incident"],
                ),
                _cell(
                    ws,
                    row_number,
                    columns["Date"],
                ),
                _cell(
                    ws,
                    row_number,
                    columns["Caller Name"],
                ),
                _cell(
                    ws,
                    row_number,
                    columns["HR Name"],
                ),
                _cell(
                    ws,
                    row_number,
                    columns["HR Email"],
                ),
                _cell(
                    ws,
                    row_number,
                    columns["Success Flag"],
                ),
            ]

            if not any(row):
                continue

            call_rows.append(row)

        cards += f"""
        <details class="card">

          <summary>

            <b>Call Logs</b>

            <span class="tag">
              {len(call_rows)} row(s)
            </span>

          </summary>

          <div>

            {_render_simple_table(
                [
                    "Company",
                    "Phone Number",
                    "Incident",
                    "Date",
                    "Caller Name",
                    "HR Name",
                    "HR Email",
                    "Success Flag",
                ],
                call_rows,
            )}

          </div>

        </details>
        """

    else:

        cards += """
        <details class="card">

          <summary>

            <b>Call Logs</b>

            <span class="tag fail">
              not found
            </span>

          </summary>

          <div>

            <p class="muted">
              Call Logs sheet not found.
            </p>

          </div>

        </details>
        """

    return page(
        f"""
        <div class="card">

          <h2>Sheet data viewer</h2>

          <p class="muted">
            Showing selected columns from Design,
            Thermal, Industrial, Production,
            and Call Logs.
          </p>

          <a
            class="btn secondary"
            href="{url_for('dashboard')}"
          >
            Back to home
          </a>

        </div>

        {cards}
        """
    )


# =====================================================================
# CONTINUOUS SENDER
# =====================================================================

@app.route("/continuous", methods=["GET"])
@mail_login_required
def continuous_view():

    try:
        wb = get_wb()

        queue, _ = L.build_continuous_queue(wb)

    except Exception as exc:

        return page(
            f"""
            <div class="card">

              <h2>Continuous sender</h2>

              <div class="error-box">
                {escape(str(exc))}
              </div>

              <a class="btn secondary"
                 href="{url_for('mail_dashboard')}">
                Back to mail dashboard
              </a>

            </div>
            """
        )

    if not queue:

        return page(
            """
            <div class="card">

              <h2>Continuous sender</h2>

              <p>
                Nothing eligible to send right now —
                everyone in scope is already logged as sent.
              </p>

              <a class="btn secondary"
                 href="{{ url_for('mail_dashboard') }}">
                Back to mail dashboard
              </a>

            </div>
            """
        )

    rows = ""

    for index, item in enumerate(queue):

        company = escape(
            str(item.get("company", ""))
        )

        sheet = escape(
            str(item.get("sheet", ""))
        )

        emails = escape(
            ", ".join(
                str(email)
                for email in item.get("emails", [])
            )
        )

        rows += f"""
        <li>

          <label style="margin-top:0">

            <input
              type="checkbox"
              name="idx"
              value="{index}"
              checked
            >

            <b>{company}</b>

            <span class="tag">
              {sheet}
            </span>

          </label>

          <span class="muted">
            {emails}
          </span>

        </li>
        """

    brochure_checked = (
        "checked"
        if BROCHURE_PATH
        else ""
    )

    body = f"""
    <div class="card">

      <h2>
        Continuous sender —
        {len(queue)} eligible
      </h2>

      <p class="muted">
        {L.PER_SHEET} per sheet per round.
      </p>

      <form
        method="post"
        action="{url_for('continuous_send')}"
      >

        <ul class="plain">
          {rows}
        </ul>

        <label>
          <input
            type="checkbox"
            name="brochure"
            {brochure_checked}
          >
          Attach brochure
        </label>

        {password_field()}

        <button type="submit">
          Send selected
        </button>

      </form>

    </div>

    <a
      class="btn secondary"
      href="{url_for('mail_dashboard')}"
    >
      Back to mail dashboard
    </a>
    """

    return page(body)


def _run_continuous_job(
    job_id,
    items,
    password,
    brochure,
):
    """Run continuous sending in a background thread."""

    def progress_callback(result):

        with _jobs_lock:

            job = _jobs.get(job_id)

            if job is not None:
                job["results"].append(result)

    try:

        wb = get_wb()

        L.send_continuous_batch(
            wb,
            items,
            password,
            brochure,
            progress_cb=progress_callback,
        )

    except Exception as exc:

        with _jobs_lock:

            job = _jobs.get(job_id)

            if job is not None:
                job["results"].append(
                    {
                        "company": "—",
                        "status": "ERROR",
                        "detail": str(exc),
                    }
                )

    finally:

        with _jobs_lock:

            job = _jobs.get(job_id)

            if job is not None:
                job["done"] = True


@app.route("/continuous/send", methods=["POST"])
@mail_login_required
def continuous_send():

    try:

        wb = get_wb()

        queue, _ = L.build_continuous_queue(wb)

        picked_idx = set()

        for value in request.form.getlist("idx"):

            try:
                picked_idx.add(int(value))

            except (TypeError, ValueError):
                continue

        items = [
            item
            for index, item in enumerate(queue)
            if index in picked_idx
        ]

        password = resolve_pw(request.form)

        if not items:

            return page(
                """
                <div class="card">

                  <h2>No companies selected</h2>

                  <p>
                    Please select at least one company.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        if not password:

            return page(
                """
                <div class="card">

                  <h2>Password required</h2>

                  <p>
                    IITD mailbox password is required.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        brochure = (
            BROCHURE_PATH
            if request.form.get("brochure")
            else None
        )

        job_id = uuid.uuid4().hex

        with _jobs_lock:

            _jobs[job_id] = {
                "done": False,
                "results": [],
                "total": len(items),
            }

        thread = threading.Thread(
            target=_run_continuous_job,
            args=(
                job_id,
                items,
                password,
                brochure,
            ),
            daemon=True,
        )

        thread.start()

        return redirect(
            url_for(
                "continuous_status",
                job_id=job_id,
            )
        )

    except Exception as exc:

        return page(
            f"""
            <div class="card">

              <h2>Could not start sender</h2>

              <div class="error-box">
                {escape(str(exc))}
              </div>

              <a class="btn secondary"
                 href="{url_for('continuous_view')}">
                Back
              </a>

            </div>
            """
        )


@app.route("/continuous/status/<job_id>")
@mail_login_required
def continuous_status(job_id):

    with _jobs_lock:
        job = _jobs.get(job_id)

        if job:
            job = {
                "done": job["done"],
                "results": list(job["results"]),
                "total": job["total"],
            }

    if not job:

        return redirect(
            url_for("continuous_view")
        )

    lines = ""

    for result in job["results"]:

        company = escape(
            str(result.get("company", "—"))
        )

        status = escape(
            str(result.get("status", "UNKNOWN"))
        )

        css_class = (
            "sent"
            if result.get("status") == "SENT"
            else "fail"
        )

        detail = result.get("detail", "")

        detail_html = ""

        if detail:

            detail_html = f"""
            <br>
            <span class="muted">
              Error: {escape(str(detail))}
            </span>
            """

        lines += f"""
        <li>

          {company}

          <span class="tag {css_class}">
            {status}
          </span>

          {detail_html}

        </li>
        """

    if job["done"]:

        status_line = "Done."

        refresh = ""

        back_button = f"""
        <a
          class="btn secondary"
          href="{url_for('mail_dashboard')}"
        >
          Back to mail dashboard
        </a>
        """

    else:

        completed = len(job["results"])

        status_line = (
            f"Sending… "
            f"({completed}/{job['total']}, "
            f"~{L.DELAY_SEC}s between each)"
        )

        refresh = """
        <meta http-equiv="refresh" content="3">
        """

        back_button = ""

    body = f"""
    {refresh}

    <div class="card">

      <h2>Continuous sender — progress</h2>

      <p class="muted">
        {escape(status_line)}
      </p>

      <ul class="plain">

        {lines or '''
        <li class="muted">
          Starting…
        </li>
        '''}

      </ul>

      {back_button}

    </div>
    """

    return page(body)


# =====================================================================
# SINGLE COMPANY SENDER — PRIVATE
# =====================================================================

@app.route("/single", methods=["GET"])
@mail_login_required
def single_view():

    results_html = ""

    query = (
        request.args.get("q", "")
        .strip()
    )

    if query:

        try:

            wb = get_wb()

            matches = L.search_company(
                wb,
                query,
            )

        except Exception as exc:

            results_html = f"""
            <div class="error-box">
              {escape(str(exc))}
            </div>
            """

        else:

            if not matches:

                results_html = f"""
                <p class="muted">
                  No match found for
                  "{escape(query)}".
                </p>

                <a
                  class="btn secondary"
                  href="{url_for(
                      'single_new_company_view',
                      company=query
                  )}"
                >
                  Send to new company
                </a>
                """

            else:

                items_html = ""

                for match in matches:

                    if match.get("already_sent"):

                        tag = (
                            '<span class="tag fail">'
                            f'already '
                            f'{escape(match.get("delivery") or "sent")}'
                            '</span>'
                        )

                    else:

                        tag = ""

                    display = escape(
                        str(match.get("display", ""))
                    )

                    sheets = escape(
                        ", ".join(
                            str(sheet)
                            for sheet in match.get(
                                "sheets",
                                [],
                            )
                        )
                    )

                    key = escape(
                        str(match.get("key", ""))
                    )

                    items_html += f"""
                    <li>

                      <b>{display}</b>
                      {tag}

                      <br>

                      <span class="muted">
                        {sheets}
                      </span>

                      <br>

                      <a
                        class="btn secondary"
                        href="{url_for(
                            'single_send_view',
                            key=key
                        )}"
                      >
                        Choose &amp; send
                      </a>

                    </li>
                    """

                results_html = f"""
                <ul class="plain">
                  {items_html}
                </ul>

                <div class="card"
                     style="margin-top:12px">

                  <p class="muted">
                    Not seeing the right company
                    in the results?
                  </p>

                  <a
                    class="btn secondary"
                    href="{url_for(
                        'single_new_company_view',
                        company=query
                    )}"
                  >
                    Send to new company
                  </a>

                </div>
                """

    body = f"""
    <div class="card">

      <h2>Single company sender</h2>

      <form method="get">

        <input
          type="text"
          name="q"
          placeholder="Company name (partial ok)"
          value="{escape(query)}"
          autofocus
        >

        <button type="submit">
          Search
        </button>

      </form>

      {results_html}

    </div>
    """

    return page(body)


@app.route("/single/send/<key>", methods=["GET", "POST"])
@mail_login_required
def single_send_view(key):

    try:

        wb = get_wb()

        matches = {
            match["key"]: match
            for match in L.search_company(
                wb,
                key,
            )
        }

        match = matches.get(key)

        if not match:

            companies, _, _ = L.build_company_data(wb)

            if key in companies:

                company_data = companies[key]

                match = {
                    "key": key,
                    "display": company_data["display"],
                    "emails": company_data["emails"],
                    "sheets": sorted(
                        company_data["sheets"]
                    ),
                }

    except Exception as exc:

        return page(
            f"""
            <div class="card">

              <h2>Company lookup error</h2>

              <div class="error-box">
                {escape(str(exc))}
              </div>

              <a class="btn secondary"
                 href="{url_for('single_view')}">
                Back
              </a>

            </div>
            """
        )

    if not match:

        return redirect(
            url_for("single_view")
        )

    if request.method == "POST":

        chosen = request.form.getlist("email")

        custom = [
            email.strip()
            for email in request.form.get(
                "custom_emails",
                "",
            ).split(",")
            if email.strip()
        ]

        recipients = L.unique_emails(
            chosen + custom
        )

        if not recipients:

            return page(
                """
                <div class="card">

                  <h2>No recipient selected</h2>

                  <p>
                    Please select or enter at least
                    one recipient.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        password = resolve_pw(request.form)

        if not password:

            return page(
                """
                <div class="card">

                  <h2>Password required</h2>

                  <p>
                    IITD mailbox password is required.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        brochure = (
            BROCHURE_PATH
            if request.form.get("brochure")
            else None
        )

        try:

            result = L.send_single(
                wb,
                key,
                match["display"],
                recipients,
                brochure,
                password,
            )

        except Exception as exc:

            result = {
                "status": "ERROR",
                "detail": str(exc),
            }

        status = result.get(
            "status",
            "ERROR",
        )

        detail = result.get(
            "detail",
            "",
        )

        status_class = (
            "sent"
            if status == "SENT"
            else "fail"
        )

        body = f"""
        <div class="card">

          <h2>
            {escape(str(match["display"]))}
          </h2>

          <p>
            <span class="tag {status_class}">
              {escape(str(status))}
            </span>
          </p>

          <p class="muted">
            To:
            {escape(", ".join(recipients))}
          </p>

          {
              f'<p class="muted">Error: '
              f'{escape(str(detail))}</p>'
              if detail
              else ''
          }

          <a
            class="btn"
            href="{url_for('single_view')}"
          >
            Back to search
          </a>

        </div>
        """

        return page(body)

    checks = ""

    for email in match.get("emails", []):

        checks += f"""
        <label>
          <input
            type="checkbox"
            name="email"
            value="{escape(str(email))}"
            checked
          >
          {escape(str(email))}
        </label>
        """

    if not checks:

        checks = """
        <p class="muted">
          No emails on file — add one below.
        </p>
        """

    brochure_checked = (
        "checked"
        if BROCHURE_PATH
        else ""
    )

    body = f"""
    <div class="card">

      <h2>
        {escape(str(match["display"]))}
      </h2>

      <form method="post">

        {checks}

        <label>
          Extra email(s), comma-separated
        </label>

        <input
          type="text"
          name="custom_emails"
          placeholder="name@company.com"
        >

        <label>
          <input
            type="checkbox"
            name="brochure"
            {brochure_checked}
          >
          Attach brochure
        </label>

        {password_field()}

        <button type="submit">
          Send
        </button>

      </form>

    </div>

    <a
      class="btn secondary"
      href="{url_for('single_view')}"
    >
      Back to search
    </a>
    """

    return page(body)


# =====================================================================
# SINGLE COMPANY — NEW COMPANY PRIVATE
# =====================================================================

@app.route("/single/new", methods=["GET", "POST"])
@mail_login_required
def single_new_company_view():

    company_prefill = (
        request.values.get(
            "company",
            "",
        ).strip()
    )

    if request.method == "POST":

        company = (
            request.form.get(
                "company",
                "",
            ).strip()
        )

        custom = (
            request.form.get(
                "emails",
                "",
            ).strip()
        )

        recipients = L.unique_emails(
            L.emails_in(custom)
        )

        if not company:

            return page(
                """
                <div class="card">

                  <h2>Missing company</h2>

                  <p>
                    Company name is required.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        if not recipients:

            return page(
                """
                <div class="card">

                  <h2>No valid email</h2>

                  <p>
                    Please enter at least one valid
                    recipient email.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        password = resolve_pw(request.form)

        if not password:

            return page(
                """
                <div class="card">

                  <h2>Password required</h2>

                  <p>
                    IITD mailbox password is required.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        try:

            wb = get_wb()

            brochure = (
                BROCHURE_PATH
                if request.form.get("brochure")
                else None
            )

            key = L.norm(company)

            result = L.send_single(
                wb=wb,
                key=key,
                company_display=company,
                recipients=recipients,
                brochure_path=brochure,
                pw=password,
            )

        except Exception as exc:

            result = {
                "status": "ERROR",
                "detail": str(exc),
            }

        status = result.get(
            "status",
            "ERROR",
        )

        detail = result.get(
            "detail",
            "",
        )

        status_class = (
            "sent"
            if status == "SENT"
            else "fail"
        )

        body = f"""
        <div class="card">

          <h2>
            {escape(company)}
          </h2>

          <p>
            <span class="tag {status_class}">
              {escape(str(status))}
            </span>
          </p>

          <p class="muted">
            To:
            {escape(", ".join(recipients))}
          </p>

          {
              f'<p class="muted">Error: '
              f'{escape(str(detail))}</p>'
              if detail
              else ''
          }

          <a
            class="btn"
            href="{url_for('single_view')}"
          >
            Back to single sender
          </a>

          <a
            class="btn secondary"
            href="{url_for('mail_dashboard')}"
          >
            Back to mail dashboard
          </a>

        </div>
        """

        return page(body)

    brochure_checked = (
        "checked"
        if BROCHURE_PATH
        else ""
    )

    body = f"""
    <div class="card">

      <h2>Send to new company</h2>

      <p class="muted">
        Use this when the company is not present
        in Design, Thermal, Industrial, Production,
        or Call Logs.

        The send will still be recorded in Sent Log.
      </p>

      <form method="post">

        <label>
          Full company name
        </label>

        <input
          type="text"
          name="company"
          value="{escape(company_prefill)}"
          placeholder="e.g. Nabhdrishti Aerospace"
          required
        >

        <label>
          Recipient email(s)
        </label>

        <input
          type="text"
          name="emails"
          placeholder="hr@company.com, careers@company.com"
          required
        >

        <label>
          <input
            type="checkbox"
            name="brochure"
            {brochure_checked}
          >
          Attach brochure
        </label>

        {password_field()}

        <button type="submit">
          Send pitch
        </button>

      </form>

      <a
        class="btn secondary"
        href="{url_for('single_view')}"
      >
        Back to search
      </a>

    </div>
    """

    return page(body)


# =====================================================================
# PUBLIC SINGLE COMPANY SENDER
# =====================================================================

@app.route(
    "/single-public",
    methods=["GET", "POST"],
)
@single_sender_required
def single_public_view():

    results_html = ""

    query = (
        request.values.get(
            "q",
            "",
        ).strip()
    )

    if query:

        try:

            wb = get_wb()

            matches = L.search_company(
                wb,
                query,
            )

        except Exception as exc:

            results_html = f"""
            <div class="error-box">
              {escape(str(exc))}
            </div>
            """

        else:

            if not matches:

                results_html = f"""
                <p class="muted">
                  No match found for
                  "{escape(query)}".
                </p>

                <a
                  class="btn secondary"
                  href="{url_for(
                      'single_public_new_company_view',
                      company=query
                  )}"
                >
                  Send to new company
                </a>
                """

            else:

                items = ""

                for match in matches:

                    if match.get("already_sent"):

                        tag = (
                            '<span class="tag fail">'
                            'already '
                            f'{escape(match.get("delivery") or "sent")}'
                            '</span>'
                        )

                    else:

                        tag = ""

                    display = escape(
                        str(match.get("display", ""))
                    )

                    sheets = escape(
                        ", ".join(
                            str(sheet)
                            for sheet in match.get(
                                "sheets",
                                [],
                            )
                        )
                    )

                    key = escape(
                        str(match.get("key", ""))
                    )

                    items += f"""
                    <li>

                      <b>{display}</b>
                      {tag}

                      <br>

                      <span class="muted">
                        {sheets}
                      </span>

                      <br>

                      <a
                        class="btn secondary"
                        href="{url_for(
                            'single_public_send_view',
                            key=key
                        )}"
                      >
                        Choose &amp; send
                      </a>

                    </li>
                    """

                results_html = f"""
                <ul class="plain">
                  {items}
                </ul>

                <div class="card"
                     style="margin-top:12px">

                  <p class="muted">
                    Not seeing the right company
                    in the results?
                  </p>

                  <a
                    class="btn secondary"
                    href="{url_for(
                        'single_public_new_company_view',
                        company=query
                    )}"
                  >
                    Send to new company
                  </a>

                </div>
                """

    body = f"""
    <div class="card">

      <h2>Single company sender</h2>

      <form method="get">

        <input
          type="text"
          name="q"
          placeholder="Company name (partial ok)"
          value="{escape(query)}"
          autofocus
        >

        <button type="submit">
          Search
        </button>

      </form>

      {results_html}

      <a
        class="btn secondary"
        href="{url_for('dashboard')}"
      >
        Back to main menu
      </a>

    </div>
    """

    return page(body)


@app.route(
    "/single-public/send/<key>",
    methods=["GET", "POST"],
)
@single_sender_required
def single_public_send_view(key):

    try:

        wb = get_wb()

        matches = {
            match["key"]: match
            for match in L.search_company(
                wb,
                key,
            )
        }

        match = matches.get(key)

        if not match:

            companies, _, _ = L.build_company_data(wb)

            if key in companies:

                company_data = companies[key]

                match = {
                    "key": key,
                    "display": company_data["display"],
                    "emails": company_data["emails"],
                    "sheets": sorted(
                        company_data["sheets"]
                    ),
                }

    except Exception as exc:

        return page(
            f"""
            <div class="card">

              <h2>Company lookup error</h2>

              <div class="error-box">
                {escape(str(exc))}
              </div>

              <a class="btn secondary"
                 href="{url_for('single_public_view')}">
                Back
              </a>

            </div>
            """
        )

    if not match:

        return redirect(
            url_for("single_public_view")
        )

    if request.method == "POST":

        chosen = request.form.getlist("email")

        custom = [
            email.strip()
            for email in request.form.get(
                "custom_emails",
                "",
            ).split(",")
            if email.strip()
        ]

        recipients = L.unique_emails(
            chosen + custom
        )

        if not recipients:

            return page(
                """
                <div class="card">

                  <h2>No recipient selected</h2>

                  <p>
                    Please select or enter
                    at least one recipient.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        password = resolve_pw(request.form)

        if not password:

            return page(
                """
                <div class="card">

                  <h2>Password required</h2>

                  <p>
                    IITD mailbox password is required.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        brochure = (
            BROCHURE_PATH
            if request.form.get("brochure")
            else None
        )

        try:

            result = L.send_single(
                wb,
                key,
                match["display"],
                recipients,
                brochure,
                password,
            )

        except Exception as exc:

            result = {
                "status": "ERROR",
                "detail": str(exc),
            }

        status = result.get(
            "status",
            "ERROR",
        )

        detail = result.get(
            "detail",
            "",
        )

        status_class = (
            "sent"
            if status == "SENT"
            else "fail"
        )

        body = f"""
        <div class="card">

          <h2>
            {escape(str(match["display"]))}
          </h2>

          <p>
            <span class="tag {status_class}">
              {escape(str(status))}
            </span>
          </p>

          <p class="muted">
            To:
            {escape(", ".join(recipients))}
          </p>

          {
              f'<p class="muted">Error: '
              f'{escape(str(detail))}</p>'
              if detail
              else ''
          }

          <a
            class="btn"
            href="{url_for('single_public_view')}"
          >
            Back to single sender
          </a>

          <a
            class="btn secondary"
            href="{url_for('dashboard')}"
          >
            Back to main menu
          </a>

        </div>
        """

        return page(body)

    checks = ""

    for email in match.get("emails", []):

        checks += f"""
        <label>

          <input
            type="checkbox"
            name="email"
            value="{escape(str(email))}"
            checked
          >

          {escape(str(email))}

        </label>
        """

    if not checks:

        checks = """
        <p class="muted">
          No emails on file — add one below.
        </p>
        """

    brochure_checked = (
        "checked"
        if BROCHURE_PATH
        else ""
    )

    body = f"""
    <div class="card">

      <h2>
        {escape(str(match["display"]))}
      </h2>

      <p class="muted">
        In:
        {escape(
            ", ".join(
                str(sheet)
                for sheet in match.get(
                    "sheets",
                    [],
                )
            )
        )}
      </p>

      <form method="post">

        {checks}

        <label>
          Extra email(s), comma-separated
        </label>

        <input
          type="text"
          name="custom_emails"
          placeholder="name@company.com"
        >

        <label>
          <input
            type="checkbox"
            name="brochure"
            {brochure_checked}
          >
          Attach brochure
        </label>

        {password_field()}

        <button type="submit">
          Send
        </button>

      </form>

      <a
        class="btn secondary"
        href="{url_for('single_public_view')}"
      >
        Back to search
      </a>

    </div>
    """

    return page(body)


# =====================================================================
# PUBLIC SINGLE — NEW COMPANY
# =====================================================================

@app.route(
    "/single-public/new",
    methods=["GET", "POST"],
)
@single_sender_required
def single_public_new_company_view():

    company_prefill = (
        request.values.get(
            "company",
            "",
        ).strip()
    )

    if request.method == "POST":

        company = (
            request.form.get(
                "company",
                "",
            ).strip()
        )

        custom = (
            request.form.get(
                "emails",
                "",
            ).strip()
        )

        recipients = L.unique_emails(
            L.emails_in(custom)
        )

        if not company:

            return page(
                """
                <div class="card">

                  <h2>Missing company</h2>

                  <p>
                    Please enter the company name.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        if not recipients:

            return page(
                """
                <div class="card">

                  <h2>No valid email</h2>

                  <p>
                    Please enter at least one
                    valid recipient email.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        password = resolve_pw(request.form)

        if not password:

            return page(
                """
                <div class="card">

                  <h2>Password required</h2>

                  <p>
                    IITD mailbox password is required.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        try:

            wb = get_wb()

            brochure = (
                BROCHURE_PATH
                if request.form.get("brochure")
                else None
            )

            key = L.norm(company)

            result = L.send_single(
                wb=wb,
                key=key,
                company_display=company,
                recipients=recipients,
                brochure_path=brochure,
                pw=password,
            )

        except Exception as exc:

            result = {
                "status": "ERROR",
                "detail": str(exc),
            }

        status = result.get(
            "status",
            "ERROR",
        )

        detail = result.get(
            "detail",
            "",
        )

        status_class = (
            "sent"
            if status == "SENT"
            else "fail"
        )

        body = f"""
        <div class="card">

          <h2>
            {escape(company)}
          </h2>

          <p>
            <span class="tag {status_class}">
              {escape(str(status))}
            </span>
          </p>

          <p class="muted">
            To:
            {escape(", ".join(recipients))}
          </p>

          {
              f'<p class="muted">Error: '
              f'{escape(str(detail))}</p>'
              if detail
              else ''
          }

          <a
            class="btn"
            href="{url_for('single_public_view')}"
          >
            Back to single sender
          </a>

          <a
            class="btn secondary"
            href="{url_for('dashboard')}"
          >
            Back to main menu
          </a>

        </div>
        """

        return page(body)

    brochure_checked = (
        "checked"
        if BROCHURE_PATH
        else ""
    )

    body = f"""
    <div class="card">

      <h2>Send to new company</h2>

      <p class="muted">
        Use this when the company is not present
        in the sheet results.

        The send will still be recorded
        in Sent Log.
      </p>

      <form method="post">

        <label>
          Full company name
        </label>

        <input
          type="text"
          name="company"
          value="{escape(company_prefill)}"
          placeholder="e.g. Nabhdrishti Aerospace"
          required
        >

        <label>
          Recipient email(s)
        </label>

        <input
          type="text"
          name="emails"
          placeholder="hr@company.com, careers@company.com"
          required
        >

        <label>
          <input
            type="checkbox"
            name="brochure"
            {brochure_checked}
          >
          Attach brochure
        </label>

        {password_field()}

        <button type="submit">
          Send pitch
        </button>

      </form>

      <a
        class="btn secondary"
        href="{url_for('single_public_view')}"
      >
        Back to search
      </a>

    </div>
    """

    return page(body)


# =====================================================================
# CHECK BOUNCES
# =====================================================================

@app.route(
    "/bounces",
    methods=["GET", "POST"],
)
@login_required
def bounces_view():

    if request.method == "POST":

        password = resolve_pw(request.form)

        if not password:

            return page(
                """
                <div class="card">

                  <h2>Password required</h2>

                  <p>
                    IITD mailbox password is required.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        days = _safe_int(
            request.form.get("days"),
            14,
        )

        try:

            wb = get_wb()

            result = L.check_bounces(
                wb,
                password,
                days=days,
            )

        except Exception as exc:

            return page(
                f"""
                <div class="card">

                  <h2>Bounce check failed</h2>

                  <div class="error-box">
                    {escape(str(exc))}
                  </div>

                  <a
                    class="btn secondary"
                    href="{url_for('mail_dashboard')}"
                  >
                    Back to mail dashboard
                  </a>

                </div>
                """
            )

        rows = ""

        for row in result.get(
            "updated_rows",
            [],
        ):

            company = escape(
                str(row.get("company", ""))
            )

            status = escape(
                str(row.get("status", ""))
            )

            bounced = escape(
                ", ".join(
                    str(email)
                    for email in row.get(
                        "bounced",
                        [],
                    )
                )
            )

            rows += f"""
            <li>

              {company}
              —
              <b>{status}</b>

              <br>

              <span class="muted">
                {bounced}
              </span>

            </li>
            """

        bounced_count = len(
            result.get(
                "bounced",
                [],
            )
        )

        updated_count = len(
            result.get(
                "updated_rows",
                [],
            )
        )

        scanned = result.get(
            "scanned",
            0,
        )

        body = f"""
        <div class="card">

          <h2>Bounce check result</h2>

          <p class="muted">
            Scanned {escape(str(scanned))}
            inbox messages,
            last {days} days.
          </p>

          <p>
            {bounced_count}
            bounced address(es) found,
            {updated_count}
            row(s) updated.
          </p>

          <ul class="plain">
            {rows or '''
            <li class="muted">
              No matching bounce updates.
            </li>
            '''}
          </ul>

          <a
            class="btn"
            href="{url_for('mail_dashboard')}"
          >
            Back to mail dashboard
          </a>

        </div>
        """

        return page(body)

    body = f"""
    <div class="card">

      <h2>Check bounces</h2>

      <form method="post">

        <label>
          Days of Inbox to scan
        </label>

        <input
          type="number"
          name="days"
          value="14"
          min="1"
          required
        >

        {password_field()}

        <button type="submit">
          Run
        </button>

      </form>

    </div>
    """

    return page(body)


# =====================================================================
# RECONCILE SENT
# =====================================================================

@app.route(
    "/reconcile",
    methods=["GET", "POST"],
)
@login_required
def reconcile_view():

    if request.method == "POST":

        password = resolve_pw(request.form)

        if not password:

            return page(
                """
                <div class="card">

                  <h2>Password required</h2>

                  <p>
                    IITD mailbox password is required.
                  </p>

                  <a class="btn"
                     href="javascript:history.back()">
                    Back
                  </a>

                </div>
                """
            )

        days = _safe_int(
            request.form.get("days"),
            30,
        )

        try:

            wb = get_wb()

            result = L.reconcile_sent(
                wb,
                password,
                days=days,
            )

        except Exception as exc:

            return page(
                f"""
                <div class="card">

                  <h2>Reconcile failed</h2>

                  <div class="error-box">
                    {escape(str(exc))}
                  </div>

                  <a
                    class="btn secondary"
                    href="{url_for('mail_dashboard')}"
                  >
                    Back to mail dashboard
                  </a>

                </div>
                """
            )

        scanned = result.get(
            "scanned",
            0,
        )

        matched = result.get(
            "matched",
            0,
        )

        added = result.get(
            "added",
            0,
        )

        updated = result.get(
            "updated",
            0,
        )

        unattributed = len(
            result.get(
                "unattributed",
                [],
            )
        )

        body = f"""
        <div class="card">

          <h2>Reconcile result</h2>

          <p class="muted">
            Scanned {escape(str(scanned))}
            Sent messages,
            last {days} days.
          </p>

          <p>
            {matched}
            pitch mail(s) matched
            ·
            {added}
            new row(s) added
            ·
            {updated}
            row(s) marked SENT.
          </p>

          <p class="muted">
            {unattributed}
            could not be attributed automatically.
          </p>

          <a
            class="btn"
            href="{url_for('mail_dashboard')}"
          >
            Back to mail dashboard
          </a>

        </div>
        """

        return page(body)

    body = f"""
    <div class="card">

      <h2>Reconcile Sent folder</h2>

      <form method="post">

        <label>
          Days of Sent folder to scan
        </label>

        <input
          type="number"
          name="days"
          value="30"
          min="1"
          required
        >

        {password_field()}

        <button type="submit">
          Run
        </button>

      </form>

    </div>
    """

    return page(body)


# =====================================================================
# HR CONTACT / PHONE LOOKUP
# =====================================================================

@app.route(
    "/hr",
    methods=["GET"],
)
@login_required
def hr_view():

    query = (
        request.args.get(
            "q",
            "",
        ).strip()
    )

    results_html = ""

    if query:

        try:

            wb = get_wb()

            matches = L.hr_lookup(
                wb,
                query,
            )

        except Exception as exc:

            results_html = f"""
            <div class="error-box">
              {escape(str(exc))}
            </div>
            """

        else:

            if not matches:

                results_html = f"""
                <p class="muted">
                  No match.
                </p>

                <a
                  class="btn secondary"
                  href="{url_for(
                      'manual_call_log_view',
                      company=query
                  )}"
                >
                  Log call for new / unlisted company
                </a>
                """

            else:

                items = ""

                for match in matches:

                    display = escape(
                        str(match.get("display", ""))
                    )

                    sheets = escape(
                        ", ".join(
                            str(sheet)
                            for sheet in match.get(
                                "sheets",
                                [],
                            )
                        )
                    )

                    contact_count = len(
                        match.get(
                            "contacts",
                            [],
                        )
                    )

                    key = escape(
                        str(match.get("key", ""))
                    )

                    items += f"""
                    <li>

                      <b>{display}</b>

                      <span class="muted">
                        [{sheets}]
                        ·
                        {contact_count}
                        contact(s)
                      </span>

                      <br>

                      <a
                        class="btn secondary"
                        href="{url_for(
                            'hr_company_view',
                            key=key,
                            q=query
                        )}"
                      >
                        View / log calls
                      </a>

                    </li>
                    """

                results_html = f"""
                <ul class="plain">
                  {items}
                </ul>
                """

    body = f"""
    <div class="card">

      <h2>HR contact / phone lookup</h2>

      <form method="get">

        <input
          type="text"
          name="q"
          placeholder="Company name (partial ok)"
          value="{escape(query)}"
          autofocus
        >

        <button type="submit">
          Search
        </button>

      </form>

      <a
        class="btn secondary"
        href="{url_for(
            'manual_call_log_view',
            company=query
        )}"
      >
        Log call for new / unlisted company
      </a>

      {results_html}

    </div>
    """

    return page(body)


# =====================================================================
# MANUAL CALL LOG
# =====================================================================

@app.route(
    "/hr/manual",
    methods=["GET", "POST"],
)
@login_required
def manual_call_log_view():

    company_prefill = (
        request.values.get(
            "company",
            "",
        ).strip()
    )

    if request.method == "POST":

        company = (
            request.form.get(
                "company",
                "",
            ).strip()
        )

        phone_raw = (
            request.form.get(
                "phone",
                "",
            ).strip()
        )

        incident = (
            request.form.get(
                "incident",
                "",
            ).strip()
        )

        caller_name = (
            request.form.get(
                "caller_name",
                "",
            ).strip()
        )

        success_flag = (
            "1"
            if request.form.get(
                "success_flag"
            ) == "1"
            else "0"
        )

        hr_name = (
            request.form.get(
                "hr_name",
                "",
            ).strip()
        )

        hr_email = (
            request.form.get(
                "hr_email",
                "",
            ).strip()
        )

        if not company:

            return page(
                f"""
                <div class="card">

                  <h2>Nothing logged</h2>

                  <p>
                    Company name is required.
                  </p>

                  <a
                    class="btn"
                    href="{url_for(
                        'manual_call_log_view',
                        company=company_prefill
                    )}"
                  >
                    Back
                  </a>

                </div>
                """
            )

        if not phone_raw and not hr_email:

            return page(
                f"""
                <div class="card">

                  <h2>Nothing logged</h2>

                  <p>
                    Please enter at least a phone
                    number or HR email.
                  </p>

                  <a
                    class="btn"
                    href="{url_for(
                        'manual_call_log_view',
                        company=company
                    )}"
                  >
                    Back
                  </a>

                </div>
                """
            )

        try:

            wb = get_wb()

            entry = {
                "company": company,
                "phone": phone_raw,
                "incident": incident,
                "date": time.strftime(
                    "%Y-%m-%d"
                ),
                "caller_name": caller_name,
                "success_flag": success_flag,
                "hr_name": hr_name,
                "hr_email": hr_email,
            }

            L.append_call_logs(
                wb,
                [entry],
            )

        except Exception as exc:

            return page(
                f"""
                <div class="card">

                  <h2>Could not save call</h2>

                  <div class="error-box">
                    {escape(str(exc))}
                  </div>

                  <a
                    class="btn"
                    href="javascript:history.back()"
                  >
                    Back
                  </a>

                </div>
                """
            )

        return page(
            f"""
            <div class="card">

              <h2>Call logged</h2>

              <p>
                Logged call/contact information
                for {escape(company)}.
              </p>

              <a
                class="btn"
                href="{url_for(
                    'hr_view',
                    q=company
                )}"
              >
                View company in HR lookup
              </a>

              <a
                class="btn secondary"
                href="{url_for(
                    'manual_call_log_view'
                )}"
              >
                Log another new company
              </a>

            </div>
            """
        )

    body = f"""
    <div class="card">

      <h2>
        Log call for new / unlisted company
      </h2>

      <form method="post">

        <label>
          Company
        </label>

        <input
          type="text"
          name="company"
          value="{escape(company_prefill)}"
          placeholder="e.g. Hyundai"
          required
        >

        <label>
          Phone Number
        </label>

        <input
          type="text"
          name="phone"
          placeholder="e.g. +91 9876543210"
        >

        <label>
          Incident
        </label>

        <input
          type="text"
          name="incident"
          placeholder="e.g. spoke to HR, asked to email brochure"
        >

        <label>
          Caller Name
        </label>

        <input
          type="text"
          name="caller_name"
          placeholder="e.g. Aman"
        >

        <label>
          Success Flag
        </label>

        <select name="success_flag">

          <option value="0" selected>
            0 — Not successful / wrong number / no pickup
          </option>

          <option value="1">
            1 — Successful / HR asked to mail
          </option>

        </select>

        <label>
          HR Name
        </label>

        <input
          type="text"
          name="hr_name"
          placeholder="e.g. Priya Sharma"
        >

        <label>
          HR Email
        </label>

        <input
          type="text"
          name="hr_email"
          placeholder="e.g. priya@company.com"
        >

        <button type="submit">
          Save call log
        </button>

      </form>

      <a
        class="btn secondary"
        href="{url_for('hr_view')}"
      >
        Back to HR lookup
      </a>

    </div>
    """

    return page(body)


# =====================================================================
# HR COMPANY VIEW
# =====================================================================

@app.route(
    "/hr/<key>",
    methods=["GET", "POST"],
)
@login_required
def hr_company_view(key):

    try:

        wb = get_wb()

        q = (
            request.values.get(
                "q",
                key,
            ).strip()
        )

        matches = {
            match["key"]: match
            for match in L.hr_lookup(
                wb,
                q,
            )
        }

        match = matches.get(key)

        if not match:

            matches_all = {
                item["key"]: item
                for item in L.hr_lookup(
                    wb,
                    key,
                )
            }

            match = matches_all.get(key)

    except Exception as exc:

        return page(
            f"""
            <div class="card">

              <h2>HR lookup error</h2>

              <div class="error-box">
                {escape(str(exc))}
              </div>

              <a
                class="btn secondary"
                href="{url_for('hr_view')}"
              >
                Back
              </a>

            </div>
            """
        )

    if not match:

        return redirect(
            url_for("hr_view")
        )

    if request.method == "POST":

        phone_raw = (
            request.form.get(
                "phone",
                "",
            ).strip()
        )

        incident = (
            request.form.get(
                "incident",
                "",
            ).strip()
        )

        caller_name = (
            request.form.get(
                "caller_name",
                "",
            ).strip()
        )

        success_flag = (
            "1"
            if request.form.get(
                "success_flag"
            ) == "1"
            else "0"
        )

        hr_name = (
            request.form.get(
                "hr_name",
                "",
            ).strip()
        )

        hr_email = (
            request.form.get(
                "hr_email",
                "",
            ).strip()
        )

        if not phone_raw:

            return page(
                f"""
                <div class="card">

                  <h2>Nothing logged</h2>

                  <p>
                    No phone number was entered.
                  </p>

                  <a
                    class="btn"
                    href="{url_for(
                        'hr_company_view',
                        key=key,
                        q=q
                    )}"
                  >
                    Back
                  </a>

                </div>
                """
            )

        entry = {
            "company": match["display"],
            "phone": phone_raw,
            "incident": incident,
            "date": time.strftime(
                "%Y-%m-%d"
            ),
            "caller_name": caller_name,
            "success_flag": success_flag,
            "hr_name": hr_name,
            "hr_email": hr_email,
        }

        try:

            L.append_call_logs(
                wb,
                [entry],
            )

        except Exception as exc:

            return page(
                f"""
                <div class="card">

                  <h2>Could not save call</h2>

                  <div class="error-box">
                    {escape(str(exc))}
                  </div>

                  <a
                    class="btn"
                    href="javascript:history.back()"
                  >
                    Back
                  </a>

                </div>
                """
            )

        return page(
            f"""
            <div class="card">

              <h2>Call logged</h2>

              <p>
                Logged call to
                {escape(phone_raw)}
                for
                {escape(str(match["display"]))}.
              </p>

              <a
                class="btn"
                href="{url_for(
                    'hr_company_view',
                    key=key,
                    q=q
                )}"
              >
                Back
              </a>

            </div>
            """
        )

    rows = ""

    for contact in match.get(
        "contacts",
        [],
    ):

        rows += f"""
        <tr>

          <td>
            {escape(str(
                contact.get("sheet", "")
            ))}
          </td>

          <td>
            {escape(str(
                contact.get("name") or "-"
            ))}
          </td>

          <td>
            {escape(str(
                contact.get("phone") or "-"
            ))}
          </td>

          <td>
            {escape(str(
                contact.get("email") or "-"
            ))}
          </td>

        </tr>
        """

    if not rows:

        rows = """
        <tr>
          <td
            colspan="4"
            class="muted"
          >
            No HR name/email/phone on file.
          </td>
        </tr>
        """

    phone_options = ""

    for contact in match.get(
        "contacts",
        [],
    ):

        phone = contact.get(
            "phone"
        )

        if not phone:
            continue

        name = str(
            contact.get("name") or ""
        )

        email = str(
            contact.get("email") or ""
        )

        phone_options += f"""
        <option
          value="{escape(str(phone))}"
          data-name="{escape(name)}"
          data-email="{escape(email)}"
        >
          {escape(name or str(phone))}
          ({escape(str(phone))})
        </option>
        """

    display = escape(
        str(match.get("display", ""))
    )

    sheets = escape(
        ", ".join(
            str(sheet)
            for sheet in match.get(
                "sheets",
                [],
            )
        )
    )

    body = f"""
    <div class="card">

      <h2>{display}</h2>

      <p class="muted">
        In: {sheets}
      </p>

      <div style="overflow-x:auto">

        <table
          style="
            width:100%;
            border-collapse:collapse;
            font-size:14px
          "
        >

          <tr class="muted">

            <th align="left">
              Sheet
            </th>

            <th align="left">
              Name
            </th>

            <th align="left">
              Phone
            </th>

            <th align="left">
              Email
            </th>

          </tr>

          {rows}

        </table>

      </div>

    </div>

    <div class="card">

      <h2>Log a call</h2>

      <form method="post">

        <label>
          Pick an existing contact's number
        </label>

        <select
          name="phone_pick"
          onchange="
            const selected =
              this.options[this.selectedIndex];

            document.getElementsByName('phone')[0].value =
              this.value;

            document.getElementsByName('hr_name')[0].value =
              selected.dataset.name || '';

            document.getElementsByName('hr_email')[0].value =
              selected.dataset.email || '';
          "
        >

          <option value="">
            — or type a number below —
          </option>

          {phone_options}

        </select>

        <label>
          Phone Number
        </label>

        <input
          type="text"
          name="phone"
          placeholder="Type or pick above"
        >

        <label>
          Incident
        </label>

        <input
          type="text"
          name="incident"
          placeholder="e.g. spoke to HR, follow up next week"
        >

        <label>
          Caller Name
        </label>

        <input
          type="text"
          name="caller_name"
          placeholder="e.g. Aman"
        >

        <label>
          Success Flag
        </label>

        <select name="success_flag">

          <option value="0" selected>
            0 — Not successful / wrong number / no pickup
          </option>

          <option value="1">
            1 — Successful / HR asked to mail
          </option>

        </select>

        <label>
          HR Name
        </label>

        <input
          type="text"
          name="hr_name"
          placeholder="e.g. Priya Sharma"
        >

        <label>
          HR Email
        </label>

        <input
          type="text"
          name="hr_email"
          placeholder="e.g. priya@company.com"
        >

        <button type="submit">
          Save call log
        </button>

      </form>

      <a
        class="btn secondary"
        href="{url_for(
            'hr_view',
            q=q
        )}"
      >
        Back to search
      </a>

    </div>
    """

    return page(body)


# =====================================================================
# ROUNDcube TEST SEND
# =====================================================================

@app.route(
    "/debug-webmail-send",
    methods=["GET", "POST"],
)
@login_required
def debug_webmail_send():

    if request.method == "POST":

        password = resolve_pw(request.form)

        recipients = L.emails_in(
            request.form.get(
                "recipient",
                "",
            )
        )

        if not password:

            return page(
                """
                <div class="card">

                  <h2>Password required</h2>

                  <p>
                    Enter your IITD mailbox password.
                  </p>

                  <a
                    class="btn"
                    href="javascript:history.back()"
                  >
                    Back
                  </a>

                </div>
                """
            )

        if not recipients:

            return page(
                """
                <div class="card">

                  <h2>Invalid recipient</h2>

                  <p>
                    Enter one valid recipient email.
                  </p>

                  <a
                    class="btn"
                    href="javascript:history.back()"
                  >
                    Back
                  </a>

                </div>
                """
            )

        try:

            L.send_one_via_roundcube(
                password=password,
                recipients=recipients[:1],
                company="Roundcube delivery test",
                subject="IITD Roundcube delivery test",
                body=(
                    "This is a one-message delivery test "
                    "from the IITD outreach application "
                    "through IITD Roundcube webmail."
                ),
            )

            status = """
            <span class="tag sent">
              TEST EMAIL SENT
            </span>
            """

            message = (
                f"Sent to "
                f"{escape(str(recipients[0]))}."
            )

        except L.WebmailSendError as exc:

            status = """
            <span class="tag fail">
              SEND FAILED
            </span>
            """

            message = escape(str(exc))

        except Exception as exc:

            status = """
            <span class="tag fail">
              SEND FAILED
            </span>
            """

            message = escape(str(exc))

        return page(
            f"""
            <div class="card">

              <h2>Roundcube test send</h2>

              <p>
                {status}
              </p>

              <p class="muted">
                {message}
              </p>

              <a
                class="btn secondary"
                href="{url_for('mail_dashboard')}"
              >
                Back to mail dashboard
              </a>

            </div>
            """
        )

    body = f"""
    <div class="card">

      <h2>
        Send one Roundcube test email
      </h2>

      <p class="muted">
        Use only your own IITD email address
        for this test.
      </p>

      <form method="post">

        <label>
          Recipient
        </label>

        <input
          type="text"
          name="recipient"
          placeholder="yourname@iitd.ac.in"
          required
        >

        {password_field()}

        <button type="submit">
          Send one test email
        </button>

      </form>

    </div>
    """

    return page(body)


# =====================================================================
# ROUNDcube LOGIN TEST
# =====================================================================

@app.route(
    "/debug-webmail",
    methods=["GET", "POST"],
)
@login_required
def debug_webmail():

    if request.method == "POST":

        password = resolve_pw(request.form)

        if not password:

            return page(
                """
                <div class="card">

                  <h2>Password required</h2>

                  <p>
                    Enter your IITD mailbox password.
                  </p>

                  <a
                    class="btn"
                    href="javascript:history.back()"
                  >
                    Back
                  </a>

                </div>
                """
            )

        try:

            roundcube_session, _ = (
                L.roundcube_compose_session(
                    password
                )
            )

            roundcube_session.close()

            status = """
            <span class="tag sent">
              LOGIN WORKED
            </span>
            """

            message = (
                "Render successfully logged in "
                "to IITD Roundcube and opened a "
                "compose session over HTTPS. "
                "No email was sent."
            )

        except L.WebmailLoginError as exc:

            status = """
            <span class="tag fail">
              LOGIN FAILED
            </span>
            """

            message = escape(str(exc))

        except Exception as exc:

            status = """
            <span class="tag fail">
              LOGIN FAILED
            </span>
            """

            message = escape(str(exc))

        return page(
            f"""
            <div class="card">

              <h2>
                IITD webmail connection
              </h2>

              <p>
                {status}
              </p>

              <p class="muted">
                {message}
              </p>

              <a
                class="btn secondary"
                href="{url_for('mail_dashboard')}"
              >
                Back to mail dashboard
              </a>

            </div>
            """
        )

    body = f"""
    <div class="card">

      <h2>
        Test IITD webmail login
      </h2>

      <p class="muted">
        This only signs in through Roundcube.
        It does not send an email.
      </p>

      <form method="post">

        {password_field()}

        <button type="submit">
          Test login
        </button>

      </form>

    </div>
    """

    return page(body)


# =====================================================================
# SMTP / IMAP CONNECTIVITY DEBUG
# =====================================================================

@app.route(
    "/debug-smtp",
    methods=["GET"],
)
@login_required
def debug_smtp():

    import socket

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

        started = time.time()

        try:

            sock = socket.create_connection(
                (host, port),
                timeout=10,
            )

            sock.close()

            elapsed = (
                time.time() - started
            )

            rows += f"""
            <li>

              <b>
                {escape(host)}:{port}
              </b>

              —
              <span class="tag sent">
                CONNECTED
              </span>

              <br>

              <span class="muted">
                {elapsed:.2f}s
              </span>

            </li>
            """

        except Exception as exc:

            rows += f"""
            <li>

              <b>
                {escape(host)}:{port}
              </b>

              —
              <span class="tag fail">
                FAILED
              </span>

              <br>

              <span class="muted">
                {escape(str(exc))}
              </span>

            </li>
            """

    return page(
        f"""
        <div class="card">

          <h2>
            SMTP / IMAP connectivity debug
          </h2>

          <p class="muted">
            This checks whether the Render server
            can establish a TCP connection to the
            listed IITD hosts and ports.
          </p>

          <ul class="plain">
            {rows}
          </ul>

          <a
            class="btn secondary"
            href="{url_for('mail_dashboard')}"
          >
            Back to mail dashboard
          </a>

        </div>
        """
    )


# =====================================================================
# ERROR HANDLERS
# =====================================================================

@app.errorhandler(404)
def page_not_found(error):

    return page(
        """
        <div class="card">

          <h2>Page not found</h2>

          <p>
            The requested page does not exist.
          </p>

          <a
            class="btn"
            href="/"
          >
            Back to home
          </a>

        </div>
        """
    ), 404


@app.errorhandler(500)
def internal_server_error(error):

    return page(
        """
        <div class="card">

          <h2>Server error</h2>

          <p>
            Something went wrong while processing
            the request.
          </p>

          <a
            class="btn"
            href="/"
          >
            Back to home
          </a>

        </div>
        """
    ), 500


# =====================================================================
# START SERVER
# =====================================================================

if __name__ == "__main__":

    port = _safe_int(
        os.environ.get("PORT"),
        5000,
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False,
    )

