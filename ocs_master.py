"""
ocs_master.py  —  IIT Delhi OCS master tool (menu-driven)

Menu:
  1  Continuous pitch sender (3 per sheet; consults Sent Log + delivery status)
  2  Single-company sender (fuzzy search, custom emails, brochure; shows sent status)
  3  Check bounces  -> full-message scan; updates Delivery Status / Bounced Emails
  4  Reconcile Sent folder -> full-message scan; back-fills Sent Log
  5  Clear cached password
  0  Exit

All data read from Priority_Company_List_Restructured.xlsx (no hardcoded data).
"""

import os, re, ssl, time, base64, hashlib, getpass, imaplib, smtplib, tempfile, difflib, email
from datetime import datetime, timedelta
from email.header import decode_header
from email.utils import parsedate_to_datetime, formataddr, formatdate, make_msgid
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from openpyxl import load_workbook

# =====================================================================
# CONFIG
# =====================================================================
FILE          = "Priority_Company_List_Restructured.xlsx"
BROCHURE_FILE = "Mechanical Department Brochure.pdf"

SMTP_HOST = "smtp.iitd.ac.in";  SMTP_PORT = 465
IMAP_HOST = "mailstore.iitd.ac.in"; IMAP_PORT = 993
SENT_FOLDER = "Sent"

FROM_ADDR = "met252767@mech.iitd.ac.in"
FROM_NAME = "Aman Vijaypratap Prajapati"
CC  = ["placement@admin.iitd.ac.in"]
BCC = ["met252947@mech.iitd.ac.in", "met252592@mech.iitd.ac.in",
       "met252832@mech.iitd.ac.in", "amanvprajapati8@gmail.com"]
SUBJECT     = "IIT Delhi Hiring Invitation for Internship and Placement Season 2027"
SUBJECT_KEY = "iit delhi hiring invitation"

PER_SHEET   = 3
DELAY_SEC   = 60
TIMEOUT     = 60
MAX_RETRIES = 2
LOG_SHEET   = "Sent Log"
SHEETS = ["Design", "Thermal", "Production", "Industrial"]
PW_CACHE_MINUTES = 60

PORTAL = ("https://protect.checkpoint.com/v2/r05/___https:/ocs.iitd.ac.in/portal/recruiter/auth___."
          "YXBzMTphZGl0eWFiaXJsYW1hbmFnZW1lbnQ6YzpvOmNlZmVhYjc0NTgyZTJkNmRmZjA5ZTlkYjk3NmMwYzgwOjc6"
          "OGNmNDo3MGRlZDJlNDBkMDYzMzYxNzgzMmFlOTAwOWJmOThhOTRjMmIzMzlmMzVhZDMwMjJhNGYyMzM0NDQ3YWVhMjY4Omg6VDpO")
TUTORIAL = ("https://protect.checkpoint.com/v2/r05/___https:/owncloud.iitd.ac.in/nextcloud/index.php/s/"
            "8WAXMCBZ63zqiaP___.YXBzMTphZGl0eWFiaXJsYW1hbmFnZW1lbnQ6YzpvOmNlZmVhYjc0NTgyZTJkNmRmZjA5ZTlkYjk3"
            "NmMwYzgwOjc6ZDZmMjoyZWNhYWYyNjhlMGEyMDUwODA5MDkxMWU2Y2E2MjUzMzlmM2FkZjM1NTEwM2NhNWQwMmFlYjM4ODcxOTYwZmIzOmg6VDpO")
DOWNLOADS = "https://ocs.iitd.ac.in/downloads"

# =====================================================================
# SMALL HELPERS
# =====================================================================
def clean(v): return "" if v is None else str(v).replace("\xa0", " ").strip()

def norm(value):
    value = clean(value).lower()
    value = re.sub(
    r"$$$dup$$$|$$$pushkar$$$|$$$vivek$$$|$$$thermax$$$",
    "",
    value,
    )
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

# =====================================================================
# DOMAIN-BASED COMPANY RESOLVER
# =====================================================================
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
# PASSWORD CACHE
# =====================================================================
def _cache_path(): return os.path.join(tempfile.gettempdir(), ".ocs_pw_cache.bin")
def _machine_key():
    seed = (os.getlogin() if hasattr(os, "getlogin") else "user") + "|" + tempfile.gettempdir()
    return hashlib.sha256(seed.encode()).digest()
def _xor(data, key): return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))
def save_pw_cache(pw):
    try:
        blob = _xor(pw.encode("utf-8"), _machine_key())
        with open(_cache_path(), "wb") as f:
            f.write(str(int(time.time())).encode() + b"|" + base64.b64encode(blob))
    except Exception: pass
def load_pw_cache():
    try:
        p = _cache_path()
        if not os.path.isfile(p): return None
        with open(p, "rb") as f: raw = f.read()
        ts_str, b64 = raw.split(b"|", 1)
        if time.time() - int(ts_str) > PW_CACHE_MINUTES * 60:
            os.remove(p); return None
        return _xor(base64.b64decode(b64), _machine_key()).decode("utf-8")
    except Exception:
        try: os.remove(_cache_path())
        except Exception: pass
        return None
def get_password():
    pw = load_pw_cache()
    if pw:
        print("(using cached IITD password — expires within 10 min)"); return pw
    pw = getpass.getpass(f"IITD password for {FROM_ADDR}: ")
    save_pw_cache(pw); return pw
def clear_pw_cache():
    try: os.remove(_cache_path())
    except Exception: pass

# =====================================================================
# SENT LOG ENGINE
# =====================================================================
LOG_HEADERS = ["Company Name","Mail Sent Flag","Mail Sent Date/Time","Emails Sent To",
               "Mail Reply","Phonic Conversation","Delivery Status","Bounced Emails"]

def get_log_sheet(wb):
    """Enforces standard column headers to fix structuring issues"""
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
            ws.cell(row=1, column=found_col, value=h)
        cols[h] = found_col

    widths = {"Company Name":30,"Mail Sent Flag":14,"Mail Sent Date/Time":20,"Emails Sent To":55,
              "Mail Reply":30,"Phonic Conversation":30,"Delivery Status":18,"Bounced Emails":55}
    for name, col in cols.items():
        ws.column_dimensions[chr(64+col)].width = widths[name]
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
        
        # Store multiple references to ensure strong hit rates for duplicates
        k_name = resolver.key_for_name_only(company)
        state[k_name] = entry
        
        k_email = resolver.key_for_emails(tried)
        if k_email:
            state[k_email] = entry
            
    return state

def log_sent(ws, cols, state, resolver, company_display, emails):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    k = resolver.key_for_emails(emails) or resolver.key_for_name_only(company_display)
    
    # Fallback exact name search
    r_target = None
    if k in state:
        r_target = state[k]["row"]
    else:
        for val in state.values():
            if norm(val["display"]) == norm(company_display):
                r_target = val["row"]
                break

    if r_target:
        r = r_target
    else:
        r = ws.max_row + 1
        ws.cell(row=r, column=cols["Company Name"], value=company_display)
        state[k] = {"row": r, "sent": True, "delivery": "pending", "emails_tried": set(), "display": company_display}
    
    ws.cell(row=r, column=cols["Mail Sent Flag"], value="SENT")
    ws.cell(row=r, column=cols["Mail Sent Date/Time"], value=ts)
    old = unique_emails(emails_in(clean(ws.cell(row=r, column=cols["Emails Sent To"]).value)))
    merged = unique_emails(old + emails)
    ws.cell(row=r, column=cols["Emails Sent To"], value=", ".join(merged))
    if not clean(ws.cell(row=r, column=cols["Delivery Status"]).value):
        ws.cell(row=r, column=cols["Delivery Status"], value="Pending")
    
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
        ws = wb[sheet_name]
        ncol = detect_col(ws, "Company Name"); ecol = detect_col(ws, "HR Email")
        if ncol is None or ecol is None:
            print(f"Warning: '{sheet_name}' missing Company Name / HR Email. Skipping."); continue
        bucket = None
        for r in range(2, ws.max_row + 1):
            cn = clean(ws.cell(row=r, column=ncol).value)
            if cn:
                bucket = {"name": cn, "emails": []}
                raw[sheet_name].append(bucket)
            if bucket is None: continue
            bucket["emails"].extend(emails_in(ws.cell(row=r, column=ecol).value))
    companies = {}; order = {s: [] for s in SHEETS}
    for sheet_name in SHEETS:
        seen = set()
        for b in raw[sheet_name]:
            k = resolver.key(b["name"], b["emails"])
            if k not in companies:
                companies[k] = {"display": b["name"], "emails": [], "sheets": set()}
            companies[k]["sheets"].add(sheet_name)
            companies[k]["emails"].extend(b["emails"])
            if k not in seen:
                order[sheet_name].append(b["name"]); seen.add(k)
    for k in companies:
        companies[k]["emails"] = unique_emails(companies[k]["emails"])
    return companies, order, resolver

# =====================================================================
# EMAIL BUILD + SEND + SAVE-TO-SENT
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
    if brochure_path:
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
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx, timeout=TIMEOUT) as server:
                server.login(FROM_ADDR, pw)
                server.sendmail(FROM_ADDR, recipients + CC + BCC, raw)
            return True
        except Exception as e:
            if is_rate_limit(e): raise RateLimitHit(str(e))
            print(f"   attempt {attempt}/{MAX_RETRIES} failed: {e}")
            if attempt < MAX_RETRIES: time.sleep(15)
    return False

def save_to_sent(pw, raw):
    try:
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=TIMEOUT)
        imap.login(FROM_ADDR, pw)
        imap.append(SENT_FOLDER, "\\Seen", imaplib.Time2Internaldate(time.time()), raw.encode("utf-8"))
        imap.logout(); return True
    except Exception as e:
        print(f"   (warning: could not save to Sent: {e})"); return False

def ask_attach_brochure():
    while True:
        ans = input(f"Attach the brochure ('{BROCHURE_FILE}')? (y/n): ").strip().lower()
        if ans == "n": return None
        if ans == "y":
            if os.path.isfile(BROCHURE_FILE): return BROCHURE_FILE
            print(f"  WARNING: '{BROCHURE_FILE}' not found in this folder.")
            if input("  Send WITHOUT brochure? (y/n): ").strip().lower() == "y": return None
            continue
        print("Please type y or n.")

# =====================================================================
# SHARED: full-message text extraction (all parts, headers, DSN, attached mail)
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
# MODE 1: CONTINUOUS SENDER
# =====================================================================
def mode_continuous(dry=False):
    wb = load_workbook(FILE)
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
            return (False, sheet_emails)                 # never sent -> all
        if st["delivery"] == "all failed":
            new = [e for e in sheet_emails if e.lower() not in st["emails_tried"]]
            return (False, new) if new else (True, [])   # requeue only new
        return (True, [])                                # sent+other -> skip

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

    print("\n=== CONTINUOUS SENDER ===")
    print(f"Eligible unsent companies: {len(queue)}  ({PER_SHEET} per sheet per round)")
    for it in queue[:12]:
        print(f"  [{it['sheet']}] {it['company']} -> {it['emails']}")
    if dry:
        wb.save(FILE); print("\nDRY RUN — nothing sent."); return
    if not queue:
        wb.save(FILE); print("\nNothing to send."); return

    brochure = ask_attach_brochure()
    pw = get_password()
    sent_count = 0
    for pos, it in enumerate(queue, start=1):
        company, k, recipients = it["company"], it["key"], it["emails"]
        print("\n" + "-"*46)
        print(f"NEXT {pos}/{len(queue)}  [{it['sheet']}] {company}")
        print("To :", ", ".join(recipients)); print("Cc :", ", ".join(CC)); print("Bcc:", ", ".join(BCC))
        print("Attachment:", os.path.basename(brochure) if brochure else "None")
        print("-"*46)
        ans = input("Send? (y=send / n=skip / q=quit): ").strip().lower()
        if ans == "q": break
        if ans != "y": print("skipped."); continue
        raw = build_message(company, recipients, brochure)
        try:
            ok = send_one(pw, recipients, raw)
        except RateLimitHit:
            wb.save(FILE)
            print(f"\n*** RATE LIMIT at {company}. Not logged. Wait 30-60 min, re-run to resume. ***")
            print(f"Sent this session: {sent_count}"); return
        if not ok:
            wb.save(FILE); print(f"FAILED (timeout): {company} — will retry later."); continue
        save_to_sent(pw, raw)
        ts = log_sent(log_ws, cols, state, resolver, company, recipients)
        wb.save(FILE); sent_count += 1
        print(f"SENT ({ts}). Delivery Status: Pending  [session total {sent_count}]")
        print(f"   waiting {DELAY_SEC}s..."); time.sleep(DELAY_SEC)
    wb.save(FILE)
    print(f"\nDone. Sent this session: {sent_count}")

# =====================================================================
# MODE 2: SINGLE COMPANY
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

def ask_recipients(found):
    print("\nHR emails found (all sheets):")
    if found:
        for i, e in enumerate(found, 1): print(f"  {i}. {e}")
    else:
        print("  (none found)")
        return ask_custom()
    while True:
        a = input("\nSend to ALL these? (y/n): ").strip().lower()
        if a == "y": return found
        if a == "n": break
        print("Type y or n.")
    while True:
        w = input("Type a different email address? (y/n): ").strip().lower()
        if w in {"y","n"}: break
        print("Type y or n.")
    if w == "n":
        return ask_indexes(found)
    custom = ask_custom()
    if not custom:
        return ask_indexes(found)
    while True:
        s = input("\n1 = only the new one(s)\n2 = new + found\nChoose 1/2: ").strip()
        if s == "1": return unique_emails(custom)
        if s == "2": return unique_emails(found + custom)
        print("Type 1 or 2.")

def ask_indexes(found):
    while True:
        raw = input("Index number(s), comma-separated (e.g. 1,3): ").strip()
        try:
            idxs = [int(x) for x in raw.split(",") if x.strip()]
            picked = [found[i-1] for i in idxs if 1 <= i <= len(found)]
            if picked and len(picked) == len(idxs): return unique_emails(picked)
        except ValueError: pass
        print("Invalid. Try again.")

def ask_custom():
    raw = input("\nType email address(es), comma/space separated: ").strip()
    got = emails_in(raw)
    if not got: print("No valid email found."); return []
    return unique_emails(got)

def mode_single():
    wb = load_workbook(FILE)
    companies, _, resolver = build_company_data(wb)
    log_ws, cols = get_log_sheet(wb)
    state = read_log_state(log_ws, cols, resolver)

    print("\n=== SINGLE COMPANY SENDER ===  (q to return to menu)")
    while True:
        query = input("\nType a company name (partial ok; q=back): ").strip()
        if query.lower() in {"q","quit","exit"}:
            wb.save(FILE); return
        matches = find_matches(query, companies)
        if not matches: print("No match."); continue
        print("\nMatches:")
        for i, k in enumerate(matches, 1):
            st = state.get(k)
            if not st: 
                # Check mapping fallbacks specifically for visual logs
                for val in state.values():
                    if norm(val["display"]) == norm(companies[k]["display"]):
                        st = val; break

            tag = ""
            if st and st["sent"]:
                tag = f"  [ALREADY {st['delivery'].upper() or 'SENT'}]"
            print(f"  {i}. {companies[k]['display']} [in: {', '.join(sorted(companies[k]['sheets']))}]{tag}")
        ch = input("Pick number (Enter=first): ").strip()
        if not ch: key = matches[0]
        else:
            try:
                sel = int(ch)
                if not (1 <= sel <= len(matches)): print("Invalid."); continue
                key = matches[sel-1]
            except ValueError: print("Invalid."); continue

        company = companies[key]["display"]
        found = companies[key]["emails"]

        # WARN if already sent, using robust Sent Log matching
        st = state.get(key)
        if not st:
            k_email = resolver.key_for_emails(found)
            st = state.get(k_email) if k_email else None
        if not st:
            for val in state.values():
                if norm(val["display"]) == norm(company):
                    st = val; break

        if st and st["sent"]:
            dstat = st["delivery"] or "sent"
            print(f"\n!!! NOTE: '{company}' is already marked {dstat.upper()} in the Sent Log.")
            print(f"    Previously emailed: {', '.join(sorted(st['emails_tried'])) or '(none recorded)'}")
            if dstat == "all failed":
                new_avail = [e for e in found if e.lower() not in st["emails_tried"]]
                if new_avail:
                    print(f"    New untried email(s) available: {', '.join(new_avail)}")
                else:
                    print("    No new email available — sending again would repeat a failed address.")
            proceed = input("    Proceed anyway? (y/n): ").strip().lower()
            if proceed != "y":
                print("    Skipped."); continue

        recipients = ask_recipients(found)
        if not recipients: print("No recipient."); continue

        brochure = ask_attach_brochure()
        print("\nTo :", ", ".join(recipients)); print("Cc :", ", ".join(CC)); print("Bcc:", ", ".join(BCC))
        print("Attachment:", os.path.basename(brochure) if brochure else "None")
        if input("Type SEND to send, anything else to cancel: ").strip() != "SEND":
            print("Cancelled."); continue

        pw = get_password()
        raw = build_message(company, recipients, brochure)
        try:
            ok = send_one(pw, recipients, raw)
        except RateLimitHit:
            wb.save(FILE); print("RATE LIMIT. Not logged. Wait 30-60 min."); return
        if not ok:
            wb.save(FILE); print("Failed to send. Not logged."); continue
        saved = save_to_sent(pw, raw)
        ts = log_sent(log_ws, cols, state, resolver, company, recipients)
        wb.save(FILE)
        print(f"SENT ({ts}). Saved to Sent: {'Yes' if saved else 'No'}")
        if input("Send another? (y/n): ").strip().lower() != "y":
            wb.save(FILE); return

# =====================================================================
# MODE 3: CHECK BOUNCES  (full-message scan)
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

def mode_bounces(days=14):
    pw = get_password()
    bounced = set()
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=TIMEOUT)
    imap.login(FROM_ADDR, pw); imap.select("INBOX")
    since = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
    typ, data = imap.search(None, f'(SINCE {since})')
    ids = data[0].split() if data and data[0] else []
    print(f"Scanning {len(ids)} Inbox messages since {since} (full-message search)...")
    for num in ids:
        typ, md = imap.fetch(num, "(RFC822)")
        if typ != "OK" or not md or not md[0]: continue
        msg = email.message_from_bytes(md[0][1])
        if not is_bounce(msg): continue
        for e in failed_recips_full(msg): bounced.add(e)
    imap.logout()
    print(f"Bounced addresses found: {len(bounced)}")
    for e in sorted(bounced): print("   -", e)
    if not bounced: print("No bounces. Nothing to update."); return

    wb = load_workbook(FILE)
    if LOG_SHEET not in wb.sheetnames: print("No Sent Log."); return
    ws, cols = get_log_sheet(wb)
    updated = 0
    for r in range(2, ws.max_row + 1):
        company = clean(ws.cell(row=r, column=cols["Company Name"]).value)
        if not company: continue
        sent_list = [e.lower() for e in emails_in(clean(ws.cell(row=r, column=cols["Emails Sent To"]).value))]
        if not sent_list: continue
        failed = [e for e in sent_list if e in bounced]
        if not failed:      status = "No bounce seen"
        elif len(failed) == len(sent_list): status = "All Failed"
        else:               status = "Partial"
        ws.cell(row=r, column=cols["Delivery Status"], value=status)
        ws.cell(row=r, column=cols["Bounced Emails"], value=", ".join(failed))
        if failed:
            updated += 1
            print(f"  {company}: {status}  (bounced: {', '.join(failed)})")
    wb.save(FILE)
    print(f"\nDone. {updated} company row(s) had bounces.")

# =====================================================================
# MODE 4: RECONCILE SENT FOLDER  (full-message scan)
# =====================================================================
GREET_RE = re.compile(r"Dear\s+(.+?)\s+(?:Recruitment\s+)?Team\b", re.I | re.S)

def mode_reconcile(days=30):
    pw = get_password()
    wb = load_workbook(FILE)
    companies, _, resolver = build_company_data(wb)
    n2d = {norm(d["display"]): d["display"] for d in companies.values()}

    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=TIMEOUT)
    imap.login(FROM_ADDR, pw); imap.select(SENT_FOLDER)
    since = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
    typ, data = imap.search(None, f'(SINCE {since})')
    ids = data[0].split() if data and data[0] else []
    print(f"Scanning {len(ids)} Sent messages since {since} (full-message search)...")

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
        
        # Recipients mapping: Check To and Cc fields 
        raw_recips = decode_hdr(msg.get("To")) + " " + decode_hdr(msg.get("Cc"))
        recips = [e.lower() for e in emails_in(raw_recips) if e.lower() not in our]
        for m in re.finditer(r"^(?:To|Cc):\s*(.+)$", text, re.I | re.M):
            for e in emails_in(m.group(1)):
                if e.lower() not in our and e.lower() not in recips:
                    recips.append(e.lower())
                    
        try: dstr = parsedate_to_datetime(msg.get("Date")).strftime("%Y-%m-%d %H:%M")
        except Exception: dstr = ""
        
        # Company determination mapping
        comp_display = None
        k = resolver.key_for_emails(recips)
        if k and k in companies:
            comp_display = companies[k]["display"]
            
        if comp_display is None:
            g = GREET_RE.search(text)
            if g:
                gc = clean(g.group(1)); comp_display = n2d.get(norm(gc), gc)
                
        if comp_display is None:
            # Fallback body text scan for company names
            for comp_name in sorted(n2d.values(), key=len, reverse=True):
                if comp_name.lower() in text.lower():
                    comp_display = comp_name
                    break
                    
        if comp_display is None:
            unattr.append((decode_hdr(msg.get("To")), dstr)); continue
            
        rec = hits.setdefault(norm(comp_display), {"display": comp_display, "emails": set(), "date": dstr})
        for e in recips: rec["emails"].add(e)
        if dstr and dstr > rec["date"]: rec["date"] = dstr
    imap.logout()
    print(f"Pitch mails found: {matched} | companies identified: {len(hits)}")

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
            ws.cell(row=r, column=cols["Company Name"], value=rec["display"])
            log_by_name[nk] = r; added += 1
        ws.cell(row=r, column=cols["Mail Sent Flag"], value="SENT")
        if not clean(ws.cell(row=r, column=cols["Mail Sent Date/Time"]).value) and rec["date"]:
            ws.cell(row=r, column=cols["Mail Sent Date/Time"], value=rec["date"])
        existing = [e for e in emails_in(clean(ws.cell(row=r, column=cols["Emails Sent To"]).value))]
        merged = existing[:]
        for e in emails_sorted:
            if e not in [x.lower() for x in merged]: merged.append(e)
        ws.cell(row=r, column=cols["Emails Sent To"], value=", ".join(merged))
        if not clean(ws.cell(row=r, column=cols["Delivery Status"]).value):
            ws.cell(row=r, column=cols["Delivery Status"], value="Pending")
    wb.save(FILE)
    print(f"\nDone. {added} new row(s) added, {updated} previously-unlogged marked SENT.")
    if unattr:
        print(f"\n--- {len(unattr)} could NOT be attributed (log manually) ---")
        for to, d in unattr: print(f"   To: {to}   [{d}]")



# =====================================================================
# MODE 5: HR CONTACT / PHONE LOOKUP
# =====================================================================

def build_hr_contact_directory(wb, companies, resolver):
    """
    Reads all four data sheets and creates:

    canonical_company_key -> list of contact records

    Each contact record contains:
    - Sheet
    - HR Name
    - HR Email
    - HR Phone

    It reads blank continuation rows too, so extra HR contacts added
    below a company are included.
    """
    contacts = {}

    for sheet_name in SHEETS:
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
            company_name = clean(
                ws.cell(row=row, column=company_col).value
            )

            # A non-empty company cell starts a company block
            if company_name:
                current_company_name = company_name
                current_company_key = resolver.key_for_name_only(
                    company_name
                )

            # Continuation row: keep company above
            if current_company_key is None:
                continue

            hr_name = ""
            hr_email = ""
            hr_phone = ""

            if hr_name_col is not None:
                hr_name = clean(
                    ws.cell(row=row, column=hr_name_col).value
                )

            if hr_email_col is not None:
                hr_email = clean(
                    ws.cell(row=row, column=hr_email_col).value
                )

            if hr_phone_col is not None:
                hr_phone = clean(
                    ws.cell(row=row, column=hr_phone_col).value
                )

            # Ignore completely empty contact rows
            if not hr_name and not hr_email and not hr_phone:
                continue

            # Ignore placeholder-only records
            placeholder_values = {
                "have to find",
                "production",
                "thermal",
                "industrial",
                "design",
                "na",
            }

            if (
                hr_name.lower() in placeholder_values
                and hr_email.lower() in placeholder_values
                and hr_phone.lower() in placeholder_values
            ):
                continue

            contacts.setdefault(
                current_company_key,
                []
            ).append(
                {
                    "sheet": sheet_name,
                    "company": current_company_name,
                    "name": hr_name,
                    "email": hr_email,
                    "phone": hr_phone,
                }
            )

    # Remove exact duplicate contact rows but preserve order
    cleaned_contacts = {}

    for company_key, rows in contacts.items():
        seen = set()
        unique_rows = []

        for row in rows:
            row_key = (
                row["sheet"].lower(),
                row["name"].lower(),
                row["email"].lower(),
                row["phone"].lower(),
            )

            if row_key not in seen:
                seen.add(row_key)
                unique_rows.append(row)

        cleaned_contacts[company_key] = unique_rows

    return cleaned_contacts


def mode_hr_contact_lookup():
    """
    Menu mode:
    Search a company using a partial name and show all HR names,
    emails and contact numbers from all sheets.
    No mail is sent in this mode.
    """
    wb = load_workbook(FILE)

    companies, _, resolver = build_company_data(wb)

    contacts = build_hr_contact_directory(
        wb,
        companies,
        resolver,
    )

    print("\n==============================================")
    print(" HR CONTACT / PHONE LOOKUP")
    print("==============================================")
    print("Search by partial company name.")
    print("Examples: rolls, cum, tasl, john, siemens")
    print("Type q to return to the main menu.")

    while True:
        query = input(
            "\nCompany name (partial is fine; q = back): "
        ).strip()

        if query.lower() in {"q", "quit", "exit"}:
            return

        matches = find_matches(query, companies)

        if not matches:
            print("No company match found. Try another spelling.")
            continue

        print("\nMatching companies:")

        for index, company_key in enumerate(matches, start=1):
            company_data = companies[company_key]

            contact_count = len(
                contacts.get(company_key, [])
            )

            print(
                f"  {index}. {company_data['display']} "
                f"[{contact_count} contact row(s)]"
            )

        selection = input(
            "\nChoose a number, or press Enter for first result: "
        ).strip()

        if not selection:
            selected_key = matches[0]

        else:
            try:
                selected_index = int(selection)

                if (
                    selected_index < 1
                    or selected_index > len(matches)
                ):
                    print("Invalid selection.")
                    continue

                selected_key = matches[selected_index - 1]

            except ValueError:
                print("Invalid selection.")
                continue

        company_data = companies[selected_key]
        contact_rows = contacts.get(selected_key, [])

        print("\n----------------------------------------------")
        print(f"COMPANY: {company_data['display']}")
        print(
            "Found in sheets: "
            + ", ".join(
                sorted(company_data["sheets"])
            )
        )
        print("----------------------------------------------")

        if not contact_rows:
            print("No HR name / email / phone contact found.")
            continue

        print(
            f"{'No.':<5}"
            f"{'Sheet':<14}"
            f"{'HR Name':<32}"
            f"{'HR Phone':<26}"
            f"HR Email"
        )

        print("-" * 120)

        for index, row in enumerate(contact_rows, start=1):
            hr_name = row["name"] or "-"
            hr_phone = row["phone"] or "-"
            hr_email = row["email"] or "-"

            print(
                f"{index:<5}"
                f"{row['sheet']:<14}"
                f"{hr_name:<32}"
                f"{hr_phone:<26}"
                f"{hr_email}"
            )

        # ---- NEW: offer to record calls into the Call Logs sheet ----
        _maybe_log_calls(wb, company_data["display"], contact_rows)

def _maybe_log_calls(wb, company_display, contact_rows):
    """
    After showing the phone lookup for a company, ask whether any calls
    were made and, if so, append entries to the 'Call Logs' sheet.

    Phone selection:
      - type the INDEX shown in the table above, OR
      - type a brand-new phone number manually (not in the list).

    Date is auto-stamped with today's date. Existing data is preserved;
    only new rows are appended.
    """
    ans = input(
        "\nDid you make any call(s) to update in 'Call Logs'? (y/n): "
    ).strip().lower()
    if ans != "y":
        print("\nLookup only. Nothing written to Call Logs.")
        return

    log_ws, cols = get_call_log_sheet(wb)
    entries = []

    while True:
        raw = input(
            "\nPhone: type an INDEX from the list above, "
            "OR type a NEW number manually (q = done): "
        ).strip()

        if raw.lower() in {"q", "quit", "exit", ""}:
            break

        phone_value = None

        # If it is a pure index into the shown table, use that contact's phone.
        if raw.isdigit() and 1 <= int(raw) <= len(contact_rows):
            picked = contact_rows[int(raw) - 1]
            phone_value = picked["phone"] or ""
            if not phone_value:
                typed = input(
                    "  That contact has no phone on file. "
                    "Type the number to log: "
                ).strip()
                phone_value = typed
        else:
            # Treat the input as a manually typed phone number.
            phone_value = raw

        if not clean(phone_value):
            print("  No phone captured. Skipping this one.")
            continue

        incident = input("  Incident (what happened in the call): ").strip()

        entries.append({
            "company":  company_display,
            "phone":    phone_value,
            "incident": incident,
            "date":     datetime.now().strftime("%Y-%m-%d"),
        })

        more = input("  Log another call for this company? (y/n): ").strip().lower()
        if more != "y":
            break

    if not entries:
        print("\nNo call entries captured. Nothing written.")
        return

    added = _append_call_logs(log_ws, cols, entries)
    wb.save(FILE)
    print(
        f"\nAdded {added} new row(s) to the '{CALL_LOG_SHEET}' sheet "
        f"(dated {datetime.now().strftime('%Y-%m-%d')}). Existing data preserved."
    )

# =====================================================================
# CALL LOG UPDATER  (append-only into the existing "Call Logs" sheet)
# Columns expected in the sheet: Company | Phone Number | Incident | Date
# =====================================================================

CALL_LOG_SHEET = "Call Logs"
# The sheet's real headers (matched case-insensitively). We also accept a
# few common aliases so minor header wording differences still map cleanly.
CALL_LOG_FIELD_ALIASES = {
    "Company":      ["company", "company name"],
    "Phone Number": ["phone number", "phone", "contact no.", "contact number"],
    "Incident":     ["incident", "status", "notes", "remark", "remarks"],
    "Date":         ["date"],
}


def get_call_log_sheet(wb):
    """
    Locate the existing Call Logs sheet (case-INSENSITIVE match, so
    'Call logs', 'Call Logs', 'CALL LOGS' all resolve to the same sheet)
    and map its real header columns to our four logical fields
    (Company, Phone Number, Incident, Date).

    - Existing headers/columns/rows are NEVER modified or reordered.
    - The sheet is only created if NO case-variant of it already exists.
    """
    # --- Case-insensitive sheet resolution (this fixes the duplicate-sheet bug) ---
    target = CALL_LOG_SHEET.strip().lower()
    existing_name = None
    for name in wb.sheetnames:
        if name.strip().lower() == target:
            existing_name = name
            break

    if existing_name is not None:
        ws = wb[existing_name]          # append into the REAL existing sheet
        fresh = False
    else:
        ws = wb.create_sheet(CALL_LOG_SHEET)
        fresh = True

    # Read the current header row once.
    header_map = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=c).value
        if val is not None and clean(val) != "":
            header_map[clean(val).lower()] = c

    cols = {}
    for field, aliases in CALL_LOG_FIELD_ALIASES.items():
        found_col = None
        for a in aliases:
            if a in header_map:
                found_col = header_map[a]
                break
        if not found_col:
            if fresh and ws.max_row == 1 and ws.max_column == 1 and \
               ws.cell(row=1, column=1).value in (None, ""):
                found_col = 1
            else:
                found_col = ws.max_column + 1
            ws.cell(row=1, column=found_col, value=field)
            header_map[field.lower()] = found_col
        cols[field] = found_col

    return ws, cols

def _append_call_logs(ws, cols, entries):
    """
    Append one new row per entry. Existing rows are untouched.
    Each entry is a dict with keys: company, phone, incident, date.
    Returns the number of rows added.
    """
    added = 0
    for e in entries:
        r = ws.max_row + 1
        ws.cell(row=r, column=cols["Company"],      value=e.get("company", ""))
        ws.cell(row=r, column=cols["Phone Number"], value=e.get("phone", ""))
        ws.cell(row=r, column=cols["Incident"],     value=e.get("incident", ""))
        ws.cell(row=r, column=cols["Date"],         value=e.get("date", ""))
        added += 1
    return added

# =====================================================================
# MENU
# =====================================================================
def menu():
    if not os.path.isfile(FILE):
        print(f"Could not find '{FILE}' in this folder."); return
    while True:
        print("\n==============================================")
        print(" IIT DELHI OCS — MASTER TOOL")
        print("==============================================")
        print(" 1  Continuous pitch sender (3 per sheet)")
        print(" 2  Single-company sender")
        print(" 3  Check bounces (update delivery status)")
        print(" 4  Reconcile Sent folder into Sent Log")
        print(" 5  HR contact / phone lookup")
        print(" 6  Clear cached password now")
        print(" 0  Exit")
        choice = input("\nChoose (0-5): ").strip()
        try:
            if choice == "1":
                d = input("Dry run first? (y/n): ").strip().lower() == "y"
                mode_continuous(dry=d)
            elif choice == "2":
                mode_single()
            elif choice == "3":
                days = input("How many days of Inbox to scan? (default 14): ").strip()
                mode_bounces(int(days) if days.isdigit() else 14)
            elif choice == "4":
                days = input("How many days of Sent to scan? (default 30): ").strip()
                mode_reconcile(int(days) if days.isdigit() else 30)
            elif choice == "5":
                mode_hr_contact_lookup()
            elif choice == "6":
                clear_pw_cache()
                print("Cached password cleared.")
            elif choice == "0":
                print("Bye."); return
            else:
                print("Invalid choice.")
        except Exception as e:
            print(f"\nERROR: {e}\n(Returning to menu — anything already sent was saved.)")

if __name__ == "__main__":
    menu()