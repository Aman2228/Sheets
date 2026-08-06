"""
ocs_logic.py — the actual OCS business logic, ported from ocs_master.py.

Same rules, same column layout, same resolver/aliasing, same rate-limit and
bounce-parsing behaviour as the original CLI tool. The only things removed
are input()/print() — every function here returns data, and app.py turns
that into web pages.
"""
import os, re, ssl, time, smtplib, imaplib, email, difflib
import requests
from urllib.parse import urljoin
from datetime import datetime, timedelta
from email.header import decode_header
from email.utils import parsedate_to_datetime, formataddr, formatdate, make_msgid
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# =====================================================================
# CONFIG  (edit these for your setup — same values as ocs_master.py)
# =====================================================================
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.iitd.ac.in")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))

IMAP_HOST = "mailstore.iitd.ac.in"; IMAP_PORT = 993

WEBMAIL_URL = os.environ.get(
    "IITD_WEBMAIL_URL",
    "https://webmail.iitd.ac.in/roundcube/"
)

class WebmailLoginError(Exception):
    pass


def roundcube_login(username, password):
    session = requests.Session()

    try:
        login_page = session.get(WEBMAIL_URL, timeout=TIMEOUT)
        login_page.raise_for_status()
    except requests.RequestException as e:
        raise WebmailLoginError(f"Could not open IITD webmail: {e}") from e

    token_match = re.search(
        r'name="_token"\s+value="([^"]+)"',
        login_page.text,
    )
    if not token_match:
        raise WebmailLoginError("IITD webmail login token was not found.")

    payload = {
        "_token": token_match.group(1),
        "_task": "login",
        "_action": "login",
        "_timezone": "Asia/Kolkata",
        "_url": "",
        "_user": username,
        "_pass": password,
    }

    try:
        response = session.post(
            urljoin(WEBMAIL_URL, "?_task=login"),
            data=payload,
            timeout=TIMEOUT,
        )
        response.raise_for_status()
    except requests.RequestException as e:
        raise WebmailLoginError(f"IITD webmail login request failed: {e}") from e

    if "_task=mail" not in response.url:
        raise WebmailLoginError(
            "IITD webmail login was not accepted. Check the username, password, or MFA."
        )

    return session

def check_roundcube_login(password):
    session = roundcube_login(WEBMAIL_USERNAME, password)
    try:
        response = session.get(
            urljoin(WEBMAIL_URL, "?_task=mail"),
            timeout=TIMEOUT,
        )
        response.raise_for_status()

        if "_task=mail" not in response.url:
            raise WebmailLoginError(
                "Roundcube session was not available after login."
            )

        return True
    except requests.RequestException as e:
        raise WebmailLoginError(
            f"Roundcube session check failed: {e}"
        ) from e
    finally:
        session.close()

def roundcube_compose_session(password):
    session = roundcube_login(WEBMAIL_USERNAME, password)

    try:
        compose_page = session.get(
            urljoin(WEBMAIL_URL, "?_task=mail&_action=compose"),
            timeout=TIMEOUT,
        )
        compose_page.raise_for_status()
    except requests.RequestException as e:
        session.close()
        raise WebmailLoginError(
            f"Could not open the Roundcube compose page: {e}"
        ) from e

    token_match = re.search(
        r'name="_token"\s+value="([^"]+)"|'
        r'"request_token":"([^"]+)"',
        compose_page.text,
    )
    if not token_match:
        session.close()
        raise WebmailLoginError(
            "Roundcube compose token was not found."
        )

    token = token_match.group(1) or token_match.group(2)

    compose_id_match = re.search(
        r'"compose_id":"([^"]+)"',
        compose_page.text,
    )
    if not compose_id_match:
        session.close()
        raise WebmailLoginError(
            "Roundcube compose ID was not found."
        )

    session.roundcube_compose_id = compose_id_match.group(1)
    return session, token

class WebmailSendError(Exception):
    pass


def send_one_via_roundcube(
    password,
    recipients,
    company,
    subject=None,
    body=None,
):

    try:
        payload = {
            "_token": token,
            "_id": session.roundcube_compose_id,
            "_from": "",
            "_to": ", ".join(recipients),
            "_cc": ", ".join(CC),
            "_bcc": ", ".join(BCC),
            "_replyto": "",
            "_subject": subject or SUBJECT,
            "_message": body if body is not None else body_for(company),
            "_is_html": "0",
            "_draft": "",
        }

        response = session.post(
            urljoin(WEBMAIL_URL, "?_task=mail&_action=send"),
            params={"_id": session.roundcube_compose_id},
            data=payload,
            timeout=TIMEOUT,
        )
        response.raise_for_status()

        if "sent_successfully" not in response.text:
            raise WebmailSendError(
                "Roundcube did not confirm that the message was sent."
            )

        return True

    except requests.RequestException as e:
        raise WebmailSendError(
            f"Roundcube send request failed: {e}"
        ) from e
    finally:
        session.close()

SENT_FOLDER = "Sent"

FROM_ADDR = "met252767@mech.iitd.ac.in"

WEBMAIL_USERNAME = os.environ.get("IITD_WEBMAIL_USERNAME", FROM_ADDR)

FROM_NAME = "Aman Vijaypratap Prajapati"
CC  = ["placement@admin.iitd.ac.in"]
BCC = ["met252947@mech.iitd.ac.in", "met252592@mech.iitd.ac.in",
       "met252832@mech.iitd.ac.in", "amanvprajapati8@gmail.com"]
SUBJECT     = "IIT Delhi Hiring Invitation for Internship and Placement Season 2027"
SUBJECT_KEY = "iit delhi hiring invitation"

PER_SHEET   = 3
DELAY_SEC   = 60
TIMEOUT     = 10
MAX_RETRIES = 1
LOG_SHEET   = "Sent Log"
SHEETS = ["Design", "Thermal", "Production", "Industrial"]

PORTAL = ("https://protect.checkpoint.com/v2/r05/___https:/ocs.iitd.ac.in/portal/recruiter/auth___."
          "YXBzMTphZGl0eWFiaXJsYW1hbmFnZW1lbnQ6YzpvOmNlZmVhYjc0NTgyZTJkNmRmZjA5ZTlkYjk3NmMwYzgwOjc6"
          "OGNmNDo3MGRlZDJlNDBkMDYzMzYxNzgzMmFlOTAwOWJmOThhOTRjMmIzMzlmMzVhZDMwMjJhNGYyMzM0NDQ3YWVhMjY4Omg6VDpO")
TUTORIAL = ("https://protect.checkpoint.com/v2/r05/___https:/owncloud.iitd.ac.in/nextcloud/index.php/s/"
            "8WAXMCBZ63zqiaP___.YXBzMTphZGl0eWFiaXJsYW1hbmFnZW1lbnQ6YzpvOmNlZmVhYjc0NTgyZTJkNmRmZjA5ZTlkYjk3"
            "NmMwYzgwOjc6ZDZmMjoyZWNhYWYyNjhlMGEyMDUwODA5MDkxMWU2Y2E2MjUzMzlmM2FkZjM1NTEwM2NhNWQwMmFlYjM4ODcxOTYwZmIzOmg6VDpO")
DOWNLOADS = "https://ocs.iitd.ac.in/downloads"

# =====================================================================
# SMALL HELPERS  (unchanged from ocs_master.py)
# =====================================================================
def clean(v): return "" if v is None else str(v).replace("\xa0", " ").strip()
def set_cell(ws, row, column, value):
    ws.cell(row=row, column=column).value = value

def norm(value):
    value = clean(value).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value).strip()
    aliases = {
        "airbus india":"airbus","airbus":"airbus","boeing india":"boeing","boeing":"boeing",
        "boeing india defence":"boeing","ge aerospace":"ge aerospace","rolls royce":"rolls royce",
        "pratt whitney":"pratt whitney","collins aerospace":"collins aerospace","cummins india":"cummins",
        "cummins":"cummins","mahindra mahindra":"mahindra","mahindra":"mahindra","tata motors":"tata motors",
        "tata steel":"tata steel","hyundai motor india":"hyundai","hyundai":"hyundai","alstom india":"alstom",
        "alstom":"alstom","bajaj auto":"bajaj auto","dassault systemes":"dassault","dassault syst mes":"dassault",
        "dassault":"dassault","siemens energy":"siemens energy","ge vernova":"ge vernova","jsw energy":"jsw energy",
        "bellatrix aerospace":"bellatrix","reliance new energy":"reliance new energy","valeo":"valeo",
        "bosch india":"bosch","bosch":"bosch","mtar technologies":"mtar",
        "mercedes benz r d india mbrdi":"mercedes","mercedes benz r d":"mercedes","mercedes benz r d india":"mercedes",
        "honeywell aerospace":"honeywell","honeywell":"honeywell",
        "tata advanced systems":"tasl","tasl":"tasl","larsen toubro":"l t","l t":"l t",
    }
    return aliases.get(value, value)

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
def emails_in(v): return EMAIL_RE.findall(clean(v))

def unique_emails(emails):
    out, seen = [], set()
    for e in emails:
        e = clean(e)
        if not e or e.lower() == "have to find": continue
        if e.lower() not in seen:
            seen.add(e.lower()); out.append(e)
    return out

def detect_col(ws, header):
    for c in range(1, ws.max_column + 1):
        if clean(ws.cell(row=1, column=c).value).lower() == header.lower():
            return c
    return None

FREE_DOMAINS = {"gmail","yahoo","hotmail","outlook","rediffmail","live",
                "icloud","protonmail","163","ymail"}

def _domain_root(email_addr):
    try:
        root = email_addr.split("@", 1)[1].lower().split(".")[0]
        return None if root in FREE_DOMAINS else root
    except Exception:
        return None

class CompanyResolver:
    def __init__(self):
        self.domain_to_key = {}
        self.name_to_key = {}
    def key(self, name, emails):
        nkey = norm(name)
        roots = set(filter(None, (_domain_root(e) for e in (emails or []))))
        for rt in roots:
            if rt in self.domain_to_key:
                k = self.domain_to_key[rt]
                self.name_to_key[nkey] = k
                for r2 in roots: self.domain_to_key.setdefault(r2, k)
                return k
        k = self.name_to_key.get(nkey, nkey)
        self.name_to_key[nkey] = k
        for rt in roots: self.domain_to_key.setdefault(rt, k)
        return k
    def key_for_name_only(self, name):
        return self.name_to_key.get(norm(name), norm(name))
    def key_for_emails(self, emails):
        for e in (emails or []):
            rt = _domain_root(e)
            if rt and rt in self.domain_to_key:
                return self.domain_to_key[rt]
        return None

# =====================================================================
# SENT LOG ENGINE  (unchanged logic)
# =====================================================================
LOG_HEADERS = ["Company Name","Mail Sent Flag","Mail Sent Date/Time","Emails Sent To",
               "Mail Reply","Phonic Conversation","Delivery Status","Bounced Emails"]

def get_log_sheet(wb):
    if LOG_SHEET in wb.sheetnames:
        ws = wb[LOG_SHEET]
    else:
        ws = wb.create_sheet(LOG_SHEET)
    cols = {}
    for h in LOG_HEADERS:
        found_col = None
        for c in range(1, ws.max_column + 2):
            val = ws.cell(row=1, column=c).value
            if val and clean(val).lower() == h.lower():
                found_col = c
                break
        if not found_col:
            found_col = ws.max_column + 1
            set_cell(ws, 1, found_col, h)
        cols[h] = found_col
    return ws, cols

def read_log_state(ws, cols, resolver):
    state = {}
    for r in range(2, ws.max_row + 1):
        company = clean(ws.cell(row=r, column=cols["Company Name"]).value)
        if not company: continue
        tried = emails_in(clean(ws.cell(row=r, column=cols["Emails Sent To"]).value))
        sent_flag = clean(ws.cell(row=r, column=cols["Mail Sent Flag"]).value).upper()
        sent = sent_flag.startswith("SENT")
        delivery = clean(ws.cell(row=r, column=cols["Delivery Status"]).value).lower()
        entry = {"row": r, "sent": sent, "delivery": delivery,
                  "emails_tried": {e.lower() for e in tried}, "display": company}
        k_name = resolver.key_for_name_only(company)
        state[k_name] = entry
        k_email = resolver.key_for_emails(tried)
        if k_email:
            state[k_email] = entry
    return state

def log_sent(ws, cols, state, resolver, company_display, emails):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    k = resolver.key_for_emails(emails) or resolver.key_for_name_only(company_display)
    r_target = None
    if k in state:
        r_target = state[k]["row"]
    else:
        for val in state.values():
            if norm(val["display"]) == norm(company_display):
                r_target = val["row"]; break
    if r_target:
        r = r_target
    else:
        r = ws.max_row + 1
        set_cell(ws, r, cols["Company Name"], company_display)
        state[k] = {"row": r, "sent": True, "delivery": "pending", "emails_tried": set(), "display": company_display}
    set_cell(ws, r, cols["Mail Sent Flag"], "SENT")
    set_cell(ws, r, cols["Mail Sent Date/Time"], ts)
    old = unique_emails(emails_in(clean(ws.cell(row=r, column=cols["Emails Sent To"]).value)))
    merged = unique_emails(old + emails)
    set_cell(ws, r, cols["Emails Sent To"], ", ".join(merged))
    if not clean(ws.cell(row=r, column=cols["Delivery Status"]).value):
        set_cell(ws, r, cols["Delivery Status"], "Pending")
    if k in state:
        state[k]["row"] = r; state[k]["sent"] = True
        state[k]["emails_tried"] |= {e.lower() for e in emails}
    return ts

# =====================================================================
# COMPANY READER
# =====================================================================
def build_company_data(wb):
    resolver = CompanyResolver()
    raw = {s: [] for s in SHEETS}

    for sheet_name in SHEETS:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        ncol = detect_col(ws, "Company Name")
        ecol = detect_col(ws, "HR Email")

        if ncol is None or ecol is None:
            continue

        bucket = None

        for r in range(2, ws.max_row + 1):
            cn = clean(ws.cell(row=r, column=ncol).value)

            if cn:
                bucket = {"name": cn, "emails": []}
                raw[sheet_name].append(bucket)

            if bucket is None:
                continue

            bucket["emails"].extend(emails_in(ws.cell(row=r, column=ecol).value))

    companies = {}
    order = {s: [] for s in SHEETS}

    for sheet_name in SHEETS:
        seen = set()

        for b in raw[sheet_name]:
            k = resolver.key(b["name"], b["emails"])

            if k not in companies:
                companies[k] = {
                    "display": b["name"],
                    "emails": [],
                    "sheets": set(),
                }

            companies[k]["sheets"].add(sheet_name)
            companies[k]["emails"].extend(b["emails"])

            if k not in seen:
                order[sheet_name].append(b["name"])
                seen.add(k)

    # Merge companies/emails discovered from Call Logs.
    call_log_companies = read_call_log_companies(wb, resolver)

    if call_log_companies:
        order.setdefault("Call Logs", [])

    for k, d in call_log_companies.items():
        if k not in companies:
            companies[k] = {
                "display": d["display"],
                "emails": [],
                "sheets": set(),
            }

        companies[k]["sheets"].add("Call Logs")
        companies[k]["emails"].extend(d["emails"])

        if d["display"] not in order["Call Logs"]:
            order["Call Logs"].append(d["display"])

    for k in companies:
        companies[k]["emails"] = unique_emails(companies[k]["emails"])

    return companies, order, resolver

# =====================================================================
# EMAIL BUILD + SEND + SAVE-TO-SENT  (unchanged)
# =====================================================================
def body_for(company):
    return f"""Dear {company} Recruitment Team,

Greetings from the Office of Career Services, IIT Delhi.

We are pleased to invite your esteemed organisation to participate in the Placement Season 2026-2027 at IIT Delhi. With a rich legacy of academic excellence and a diverse pool of highly skilled students across Undergraduate, Postgraduate, and PhD programs, IIT Delhi remains a preferred campus for top recruiters worldwide.

The Department of Mechanical Engineering at IIT Delhi offers highly skilled students with strong foundations in thermal sciences, fluid mechanics, computational modelling, and advanced manufacturing. Their technical expertise and problem-solving abilities make them well-suited for engineering roles at {company}.

We are also pleased to announce the commencement of the Internship Hiring process for Master's students for Summer 2027 and would be delighted to have your organization participate.

Registration & Resources:

Register for internship and/or full-time hiring (registrations Open Now - Please ensure you register for Internship 2027 and/or Placement 2027):
Link to Recruiter Portal
{PORTAL}

For guidance on the registration process and timelines, please refer to the resources below:

- Registration & JNF Creation Tutorial:  Watch Here
{TUTORIAL}
- Placement Brochure & Flyer: {DOWNLOADS}
- Important Timelines: {DOWNLOADS}

We look forward to a meaningful collaboration and welcoming your organisation this season.

For any queries or assistance, feel free to contact undersigned:

Mansvi Gour
Overall Coordinator (PG)
mee252753@iitd.ac.in
+91 9303762683

Kritik Rathore
Overall Coordinator (PG)
che252177@iitd.ac.in
+91 7023398924

Aman Prajapati
PG Coordinator
met252767@mech.iitd.ac.in
+91 7715079808

Warm regards,
Aman Vijaypratap Prajapati
Nucleus Team Member(PG)
Office of Career Services
Indian Institute of Technology Delhi
"""

def build_message(company, recipients, brochure_path):
    body = body_for(company)
    if brochure_path and os.path.isfile(brochure_path):
        msg = MIMEMultipart()
        msg.attach(MIMEText(body, "plain", "utf-8"))
        with open(brochure_path, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=os.path.basename(brochure_path))
        msg.attach(part)
    else:
        msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = SUBJECT
    msg["From"] = formataddr((FROM_NAME, FROM_ADDR))
    msg["To"] = ", ".join(recipients)
    msg["Cc"] = ", ".join(CC)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="mech.iitd.ac.in")
    return msg.as_string()

class RateLimitHit(Exception): pass
def is_rate_limit(t):
    t = str(t).lower()
    return ("sending rate too high" in t or "policy rejection" in t or "450" in t or "4.7.1" in t)

def send_one(pw, recipients, raw):
    ctx = ssl.create_default_context()
    last_error = ""

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            if SMTP_PORT == 465:
                with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx, timeout=TIMEOUT) as server:
                    server.login(FROM_ADDR, pw)
                    server.sendmail(FROM_ADDR, recipients + CC + BCC, raw)
            else:
                with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=TIMEOUT) as server:
                    server.ehlo()
                    server.starttls(context=ctx)
                    server.ehlo()
                    server.login(FROM_ADDR, pw)
                    server.sendmail(FROM_ADDR, recipients + CC + BCC, raw)

            return True

        except Exception as e:
            last_error = str(e)
            print(f"SMTP send failed on attempt {attempt}: {last_error}", flush=True)

            if is_rate_limit(e):
                raise RateLimitHit(str(e))

            if attempt < MAX_RETRIES:
                time.sleep(15)

    print(f"SMTP send ultimately failed: {last_error}", flush=True)
    return False

def save_to_sent(pw, raw):
    try:
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=TIMEOUT)
        imap.login(FROM_ADDR, pw)
        imap.append(SENT_FOLDER, "\\Seen", imaplib.Time2Internaldate(time.time()), raw.encode("utf-8"))
        imap.logout(); return True
    except Exception:
        return False

# =====================================================================
# full-message text extraction (bounces / reconcile)
# =====================================================================
def decode_hdr(val):
    if not val: return ""
    out = ""
    for t, e in decode_header(val):
        out += t.decode(e or "utf-8", "replace") if isinstance(t, bytes) else t
    return out

def whole_message_text(msg):
    chunks = []
    try:
        for h in ("To", "Cc", "Subject", "X-Failed-Recipients", "Original-Recipient"):
            v = decode_hdr(msg.get(h))
            if v: chunks.append(f"{h}: {v}")
    except Exception:
        pass
    def dec(p):
        try:
            pl = p.get_payload(decode=True)
            if pl: return pl.decode(p.get_content_charset() or "utf-8", "replace")
        except Exception:
            pass
        try:
            return str(p.get_payload())
        except Exception:
            return ""
    if msg.is_multipart():
        for part in msg.walk():
            try:
                for h in ("To","Original-Recipient","Final-Recipient","X-Failed-Recipients"):
                    hv = decode_hdr(part.get(h))
                    if hv: chunks.append(f"{h}: {hv}")
            except Exception:
                pass
            chunks.append(dec(part))
    else:
        chunks.append(dec(msg))
    text = "\n".join(c for c in chunks if c)
    return text + "\n" + re.sub(r"<[^>]+>", " ", text)

# =====================================================================
# MODE 1 (logic): CONTINUOUS QUEUE + SEND
# =====================================================================
def build_continuous_queue(wb):
    companies, order, resolver = build_company_data(wb)
    log_ws, cols = get_log_sheet(wb)
    state = read_log_state(log_ws, cols, resolver)

    def skip_and_emails(k, sheet_emails, disp):
        st = state.get(k)
        if not st:
            k_email = resolver.key_for_emails(sheet_emails)
            st = state.get(k_email) if k_email else None
        if not st:
            for val in state.values():
                if norm(val["display"]) == norm(disp):
                    st = val; break
        if not st or not st["sent"]:
            return (False, sheet_emails)
        if st["delivery"] == "all failed":
            new = [e for e in sheet_emails if e.lower() not in st["emails_tried"]]
            return (False, new) if new else (True, [])
        return (True, [])

    pointers = {s: 0 for s in SHEETS}
    used = set(); queue = []
    while any(pointers[s] < len(order.get(s, [])) for s in SHEETS):
        added = False
        for s in SHEETS:
            picked = 0
            lst = order.get(s, [])
            while pointers[s] < len(lst) and picked < PER_SHEET:
                disp = lst[pointers[s]]; pointers[s] += 1
                k = resolver.key_for_name_only(disp)
                data = companies.get(k)
                if data is None:
                    for kk, dd in companies.items():
                        if norm(dd["display"]) == norm(disp):
                            k, data = kk, dd; break
                if not k or k in used or data is None or not data["emails"]:
                    continue
                skip, emails_to_use = skip_and_emails(k, data["emails"], disp)
                if skip or not emails_to_use:
                    used.add(k); continue
                used.add(k)
                queue.append({"sheet": s, "company": data["display"], "key": k, "emails": emails_to_use})
                picked += 1; added = True
        if not added: break
    return queue, resolver

def send_continuous_batch(wb, items, pw, brochure_path, progress_cb=None, delay=DELAY_SEC):
    """items: list of {sheet, company, key, emails} selected by the user.
    Returns list of result dicts. Saves after every send (matches original)."""
    log_ws, cols = get_log_sheet(wb)
    _, _, resolver = build_company_data(wb)
    state = read_log_state(log_ws, cols, resolver)
    results = []
    for i, it in enumerate(items):
        company, recipients = it["company"], it["emails"]
        raw = build_message(company, recipients, brochure_path)
        try:
            ok = send_one(pw, recipients, raw)
        except RateLimitHit as e:
            wb.save()
            results.append({"company": company, "status": "RATE_LIMIT", "detail": str(e)})
            if progress_cb: progress_cb(results[-1])
            break
        if not ok:
            wb.save()
            results.append({"company": company, "status": "FAILED"})
            if progress_cb: progress_cb(results[-1])
            break
        saved = save_to_sent(pw, raw)
        ts = log_sent(log_ws, cols, state, resolver, company, recipients)
        wb.save()
        results.append({"company": company, "status": "SENT", "time": ts, "saved_to_sent": saved})
        if progress_cb: progress_cb(results[-1])
        if i < len(items) - 1:
            time.sleep(delay)
    return results

# =====================================================================
# MODE 2 (logic): SINGLE COMPANY SEARCH + SEND
# =====================================================================
def find_matches(query, companies):
    q = norm(query)
    if q in companies: return [q]
    m = []
    for k, d in companies.items():
        dn = norm(d["display"])
        if q in k or q in dn: m.append(k); continue
        qw, cw = set(q.split()), set(dn.split())
        if qw and qw.issubset(cw): m.append(k)
    if m: return m
    return difflib.get_close_matches(q, list(companies.keys()), n=8, cutoff=0.35)

def search_company(wb, query):
    companies, _, resolver = build_company_data(wb)
    log_ws, cols = get_log_sheet(wb)
    state = read_log_state(log_ws, cols, resolver)
    matches = find_matches(query, companies)
    out = []
    for k in matches:
        st = state.get(k)
        if not st:
            for val in state.values():
                if norm(val["display"]) == norm(companies[k]["display"]):
                    st = val; break
        out.append({
            "key": k,
            "display": companies[k]["display"],
            "emails": companies[k]["emails"],
            "sheets": sorted(companies[k]["sheets"]),
            "already_sent": bool(st and st["sent"]),
            "delivery": (st["delivery"] if st else "") or "",
            "tried": sorted(st["emails_tried"]) if st else [],
        })
    return out

def send_single(wb, key, company_display, recipients, brochure_path, pw):
    companies, _, resolver = build_company_data(wb)
    log_ws, cols = get_log_sheet(wb)
    state = read_log_state(log_ws, cols, resolver)
    raw = build_message(company_display, recipients, brochure_path)
    try:
        ok = send_one(pw, recipients, raw)
    except RateLimitHit as e:
        wb.save()
        return {"status": "RATE_LIMIT", "detail": str(e)}
    if not ok:
        wb.save()
        return {"status": "FAILED"}
    saved = save_to_sent(pw, raw)
    ts = log_sent(log_ws, cols, state, resolver, company_display, recipients)
    wb.save()
    return {"status": "SENT", "time": ts, "saved_to_sent": saved}

# =====================================================================
# MODE 3 (logic): CHECK BOUNCES
# =====================================================================
def is_bounce(msg):
    frm = decode_hdr(msg.get("From")).lower()
    subj = decode_hdr(msg.get("Subject")).lower()
    ctype = (msg.get_content_type() or "").lower()
    if "postmaster@" in frm or "mailer-daemon" in frm or "mail delivery" in frm:
        return True
    if ("undeliverable" in subj or "delivery has failed" in subj or
        "delivery status notification" in subj or "returned mail" in subj or
        "mail delivery failed" in subj or "failure notice" in subj):
        return True
    if "report-type=delivery-status" in ctype or "multipart/report" in ctype:
        return True
    return False

def failed_recips_full(msg):
    text = whole_message_text(msg)
    failed = []
    for m in re.finditer(r"Final-Recipient:\s*rfc822;\s*([^\s]+)", text, re.I):
        e = m.group(1).strip().strip("<>").lower()
        if EMAIL_RE.fullmatch(e) and e not in failed: failed.append(e)
    for m in re.finditer(r"(?:X-Failed-Recipients|Original-Recipient)[^\n:]*:\s*(?:rfc822;)?\s*([^\s,;]+)", text, re.I):
        e = m.group(1).strip().strip("<>").lower()
        if EMAIL_RE.fullmatch(e) and e not in failed: failed.append(e)
    blk = re.search(r"failed to these recipients[^\n]*\n(.*?)(?:\n\s*\n|your message)", text, re.I | re.S)
    if blk:
        for e in EMAIL_RE.findall(blk.group(1)):
            if e.lower() not in failed: failed.append(e.lower())
    our = {a.lower() for a in [FROM_ADDR] + CC + BCC}
    return [e for e in failed if e not in our]

def check_bounces(wb, pw, days=14):
    bounced = set()
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=TIMEOUT)
    imap.login(FROM_ADDR, pw); imap.select("INBOX")
    since = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
    typ, data = imap.search(None, f'(SINCE {since})')
    ids = data[0].split() if data and data[0] else []
    for num in ids:
        typ, md = imap.fetch(num, "(RFC822)")
        if typ != "OK" or not md or not md[0]: continue
        msg = email.message_from_bytes(md[0][1])
        if not is_bounce(msg): continue
        for e in failed_recips_full(msg): bounced.add(e)
    imap.logout()
    if not bounced:
        return {"scanned": len(ids), "bounced": [], "updated_rows": []}

    if LOG_SHEET not in wb.sheetnames:
        return {"scanned": len(ids), "bounced": sorted(bounced), "updated_rows": [], "no_log": True}
    ws, cols = get_log_sheet(wb)
    updated_rows = []
    for r in range(2, ws.max_row + 1):
        company = clean(ws.cell(row=r, column=cols["Company Name"]).value)
        if not company: continue
        sent_list = [e.lower() for e in emails_in(clean(ws.cell(row=r, column=cols["Emails Sent To"]).value))]
        if not sent_list: continue
        failed = [e for e in sent_list if e in bounced]
        if not failed:      status = "No bounce seen"
        elif len(failed) == len(sent_list): status = "All Failed"
        else:               status = "Partial"
        set_cell(ws, r, cols["Delivery Status"], status)
        set_cell(ws, r, cols["Bounced Emails"], ", ".join(failed))
        if failed:
            updated_rows.append({"company": company, "status": status, "bounced": failed})
    wb.save()
    return {"scanned": len(ids), "bounced": sorted(bounced), "updated_rows": updated_rows}

# =====================================================================
# MODE 4 (logic): RECONCILE SENT FOLDER
# =====================================================================
GREET_RE = re.compile(r"Dear\s+(.+?)\s+(?:Recruitment\s+)?Team\b", re.I | re.S)

def build_hr_contact_directory(wb, companies, resolver):
    """Reads all four sheets, returns canonical_company_key -> list of
    {sheet, company, name, email, phone}. Ported unchanged from ocs_master.py."""
    contacts = {}
    for sheet_name in SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        company_col = detect_col(ws, "Company Name")
        hr_name_col = detect_col(ws, "HR Name")
        hr_email_col = detect_col(ws, "HR Email")
        hr_phone_col = detect_col(ws, "HR Phone")
        if company_col is None:
            continue
        current_company_name = None
        current_company_key = None
        for row in range(2, ws.max_row + 1):
            company_name = clean(ws.cell(row=row, column=company_col).value)
            if company_name:
                current_company_name = company_name
                current_company_key = resolver.key_for_name_only(company_name)
            if current_company_key is None:
                continue
            hr_name = clean(ws.cell(row=row, column=hr_name_col).value) if hr_name_col else ""
            hr_email = clean(ws.cell(row=row, column=hr_email_col).value) if hr_email_col else ""
            hr_phone = clean(ws.cell(row=row, column=hr_phone_col).value) if hr_phone_col else ""
            if not hr_name and not hr_email and not hr_phone:
                continue
            placeholder_values = {"have to find", "production", "thermal", "industrial", "design", "na"}
            if (hr_name.lower() in placeholder_values and hr_email.lower() in placeholder_values
                    and hr_phone.lower() in placeholder_values):
                continue
            contacts.setdefault(current_company_key, []).append({
                "sheet": sheet_name, "company": current_company_name,
                "name": hr_name, "email": hr_email, "phone": hr_phone,
            })
    cleaned = {}
    for company_key, rows in contacts.items():
        seen = set(); unique_rows = []
        for row in rows:
            row_key = (row["sheet"].lower(), row["name"].lower(), row["email"].lower(), row["phone"].lower())
            if row_key not in seen:
                seen.add(row_key); unique_rows.append(row)
        cleaned[company_key] = unique_rows
        # Also include contacts discovered from Call Logs.
    call_log_companies = read_call_log_companies(wb, resolver)

    for company_key, data in call_log_companies.items():
        cleaned.setdefault(company_key, [])

        existing = {
            (
                row["sheet"].lower(),
                row["name"].lower(),
                row["email"].lower(),
                row["phone"].lower(),
            )
            for row in cleaned[company_key]
        }

        for row in data.get("contacts", []):
            row_key = (
                row["sheet"].lower(),
                row["name"].lower(),
                row["email"].lower(),
                row["phone"].lower(),
            )

            if row_key not in existing:
                existing.add(row_key)
                cleaned[company_key].append(row)
    return cleaned

def hr_lookup(wb, query):
    """Search companies + return contact rows for matches. Read-only."""
    companies, _, resolver = build_company_data(wb)
    contacts = build_hr_contact_directory(wb, companies, resolver)
    matches = find_matches(query, companies)
    out = []
    for k in matches:
        out.append({
            "key": k,
            "display": companies[k]["display"],
            "sheets": sorted(companies[k]["sheets"]),
            "contacts": contacts.get(k, []),
        })
    return out

# =====================================================================
# CALL LOGS  (append-only into the "Call Logs" sheet)
# =====================================================================
CALL_LOG_SHEET = "Call Logs"
CALL_LOG_FIELD_ALIASES = {
    "Company":      ["company", "company name"],
    "Phone Number": ["phone number", "phone", "contact no.", "contact number"],
    "Incident":     ["incident", "status", "notes", "remark", "remarks"],
    "Date":         ["Date"],
    "Caller Name":  ["Caller Name"], 
    "HR Name":      ["HR Name"],
    "HR Email":     ["HR Email"]
}

def _header_key(v):
    """
    Normalize sheet headers so that Date, date, Date/Time, Date Time,
    date-time etc. can be matched more reliably.
    """
    return re.sub(r"[^a-z0-9]+", " ", clean(v).lower()).strip()


def get_call_log_sheet(wb):
    target = CALL_LOG_SHEET.strip().lower()
    existing_name = None

    for name in wb.sheetnames:
        if name.strip().lower() == target:
            existing_name = name
            break

    if existing_name is not None:
        ws = wb[existing_name]
    else:
        ws = wb.create_sheet(CALL_LOG_SHEET)

    header_map = {}

    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=c).value
        hk = _header_key(val)
        if hk:
            header_map[hk] = c

    cols = {}

    for field, aliases in CALL_LOG_FIELD_ALIASES.items():
        found_col = None

        for a in aliases:
            ak = _header_key(a)
            if ak in header_map:
                found_col = header_map[ak]
                break

        if not found_col:
            found_col = ws.max_column + 1
            set_cell(ws, 1, found_col, field)
            header_map[_header_key(field)] = found_col

        cols[field] = found_col

    return ws, cols

def append_call_logs(wb, entries):
    """entries: list of call log records. Saves the workbook."""
    ws, cols = get_call_log_sheet(wb)
    added = 0

    for e in entries:
        r = ws.max_row + 1

        set_cell(ws, r, cols["Company"], e.get("company", ""))
        set_cell(ws, r, cols["Phone Number"], e.get("phone", ""))
        set_cell(ws, r, cols["Incident"], e.get("incident", ""))
        set_cell(ws, r, cols["Date"], e.get("date", ""))
        set_cell(ws, r, cols["Caller Name"], e.get("caller_name", ""))
        set_cell(ws, r, cols["HR Name"], e.get("hr_name", ""))
        set_cell(ws, r, cols["HR Email"], e.get("hr_email", ""))

        added += 1

    wb.save()
    return added

def read_call_log_companies(wb, resolver=None):
    """
    Reads Call Logs and returns company/email/contact data found there.

    Output format:
    {
      company_key: {
        "display": company name,
        "emails": [...],
        "contacts": [
          {
            "sheet": "Call Logs",
            "company": company,
            "name": hr_name,
            "email": hr_email,
            "phone": phone,
          }
        ]
      }
    }
    """
    if CALL_LOG_SHEET not in wb.sheetnames:
        return {}

    if resolver is None:
        resolver = CompanyResolver()

    ws, cols = get_call_log_sheet(wb)
    out = {}

    for r in range(2, ws.max_row + 1):
        company = clean(ws.cell(row=r, column=cols["Company"]).value)
        if not company:
            continue

        phone = clean(ws.cell(row=r, column=cols["Phone Number"]).value)
        hr_name = clean(ws.cell(row=r, column=cols["HR Name"]).value)
        hr_email_raw = clean(ws.cell(row=r, column=cols["HR Email"]).value)

        found_emails = unique_emails(emails_in(hr_email_raw))

        key = resolver.key(company, found_emails)

        if key not in out:
            out[key] = {
                "display": company,
                "emails": [],
                "contacts": [],
            }

        out[key]["emails"].extend(found_emails)

        if phone or hr_name or found_emails:
            out[key]["contacts"].append({
                "sheet": "Call Logs",
                "company": company,
                "name": hr_name,
                "email": ", ".join(found_emails) if found_emails else hr_email_raw,
                "phone": phone,
            })

    for key in out:
        out[key]["emails"] = unique_emails(out[key]["emails"])

        seen = set()
        unique_contacts = []

        for c in out[key]["contacts"]:
            ck = (
                c["sheet"].lower(),
                c["company"].lower(),
                c["name"].lower(),
                c["email"].lower(),
                c["phone"].lower(),
            )
            if ck not in seen:
                seen.add(ck)
                unique_contacts.append(c)

        out[key]["contacts"] = unique_contacts

    return out

def reconcile_sent(wb, pw, days=30):
    companies, _, resolver = build_company_data(wb)
    n2d = {norm(d["display"]): d["display"] for d in companies.values()}

    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=TIMEOUT)
    imap.login(FROM_ADDR, pw); imap.select(SENT_FOLDER)
    since = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
    typ, data = imap.search(None, f'(SINCE {since})')
    ids = data[0].split() if data and data[0] else []

    our = {a.lower() for a in [FROM_ADDR] + CC + BCC}
    hits, unattr, matched = {}, [], 0
    for num in ids:
        typ, md = imap.fetch(num, "(RFC822)")
        if typ != "OK" or not md or not md[0]: continue
        msg = email.message_from_bytes(md[0][1])
        subj = decode_hdr(msg.get("Subject")).lower()
        text = whole_message_text(msg)
        if SUBJECT_KEY not in subj and SUBJECT_KEY not in text.lower():
            continue
        matched += 1
        raw_recips = decode_hdr(msg.get("To")) + " " + decode_hdr(msg.get("Cc"))
        recips = [e.lower() for e in emails_in(raw_recips) if e.lower() not in our]
        for m in re.finditer(r"^(?:To|Cc):\s*(.+)$", text, re.I | re.M):
            for e in emails_in(m.group(1)):
                if e.lower() not in our and e.lower() not in recips:
                    recips.append(e.lower())
        try: dstr = parsedate_to_datetime(msg.get("Date")).strftime("%Y-%m-%d %H:%M")
        except Exception: dstr = ""
        comp_display = None
        k = resolver.key_for_emails(recips)
        if k and k in companies:
            comp_display = companies[k]["display"]
        if comp_display is None:
            g = GREET_RE.search(text)
            if g:
                gc = clean(g.group(1)); comp_display = n2d.get(norm(gc), gc)
        if comp_display is None:
            for comp_name in sorted(n2d.values(), key=len, reverse=True):
                if comp_name.lower() in text.lower():
                    comp_display = comp_name; break
        if comp_display is None:
            unattr.append({"to": decode_hdr(msg.get("To")), "date": dstr}); continue
        rec = hits.setdefault(norm(comp_display), {"display": comp_display, "emails": set(), "date": dstr})
        for e in recips: rec["emails"].add(e)
        if dstr and dstr > rec["date"]: rec["date"] = dstr
    imap.logout()

    ws, cols = get_log_sheet(wb)
    log_by_name = {}
    for r in range(2, ws.max_row + 1):
        cn = clean(ws.cell(row=r, column=cols["Company Name"]).value)
        if cn: log_by_name[norm(cn)] = r
    added = updated = 0
    for nk, rec in hits.items():
        emails_sorted = sorted(rec["emails"])
        if nk in log_by_name:
            r = log_by_name[nk]
            if not clean(ws.cell(row=r, column=cols["Mail Sent Flag"]).value).upper().startswith("SENT"):
                updated += 1
        else:
            r = ws.max_row + 1
            set_cell(ws, r, cols["Company Name"], rec["display"])
            log_by_name[nk] = r; added += 1
        set_cell(ws, r, cols["Mail Sent Flag"], "SENT")
        if not clean(ws.cell(row=r, column=cols["Mail Sent Date/Time"]).value) and rec["date"]:
            set_cell(ws, r, cols["Mail Sent Date/Time"], rec["date"])
        existing = [e for e in emails_in(clean(ws.cell(row=r, column=cols["Emails Sent To"]).value))]
        merged = existing[:]
        for e in emails_sorted:
            if e not in [x.lower() for x in merged]: merged.append(e)
        set_cell(ws, r, cols["Emails Sent To"], ", ".join(merged))
        if not clean(ws.cell(row=r, column=cols["Delivery Status"]).value):
            set_cell(ws, r, cols["Delivery Status"], "Pending")
    wb.save()
    return {"scanned": len(ids), "matched": matched, "added": added, "updated": updated,
            "unattributed": unattr}
