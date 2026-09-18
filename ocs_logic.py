"""
ocs_logic.py
------------
OCS business logic for the IIT Delhi recruiter-mail application.

This module contains the business logic originally implemented in
ocs_master.py, adapted for use by a web application.

Important:
- No input() calls.
- No CLI interaction.
- Functions return data/results.
- The web layer (app.py) is responsible for rendering pages.
- Roundcube webmail is used for sending mail.
- IMAP is used for bounce checking and sent-folder reconciliation.
"""

import os
import re
import ssl
import time
import smtplib
import imaplib
import email
import difflib
import logging
from datetime import datetime, timedelta
from email.header import decode_header
from email.utils import (
    parsedate_to_datetime,
    formataddr,
    formatdate,
    make_msgid,
)
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from urllib.parse import urljoin

import requests


# =====================================================================
# LOGGING
# =====================================================================

logger = logging.getLogger(__name__)


# =====================================================================
# CONFIG
# =====================================================================

SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.iitd.ac.in")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))

IMAP_HOST = os.environ.get("IMAP_HOST", "mailstore.iitd.ac.in")
IMAP_PORT = int(os.environ.get("IMAP_PORT", "993"))

WEBMAIL_URL = os.environ.get(
    "IITD_WEBMAIL_URL",
    "https://webmail.iitd.ac.in/roundcube/",
)

SENT_FOLDER = os.environ.get("SENT_FOLDER", "Sent")

FROM_ADDR = os.environ.get(
    "FROM_ADDR",
    "met252767@mech.iitd.ac.in",
)

WEBMAIL_USERNAME = os.environ.get(
    "IITD_WEBMAIL_USERNAME",
    FROM_ADDR,
)

FROM_NAME = os.environ.get(
    "FROM_NAME",
    "Aman Vijaypratap Prajapati",
)

CC = [
    "placement@admin.iitd.ac.in",
    "esn252272@iitd.ac.in",
]

BCC = [
    "mep252764@mech.iitd.ac.in",
    "amanvprajapati8@gmail.com",
]

SUBJECT = (
    "IIT Delhi Hiring Invitation for Internship and Placement Season 2027"
)

SUBJECT_KEY = "iit delhi hiring invitation"

PER_SHEET = 3
DELAY_SEC = 10
TIMEOUT = 20
ATTACHMENT_UPLOAD_TIMEOUT = 180
MAX_RETRIES = 1

SHEETS = [
    "Design",
    "Thermal",
    "Production",
    "Industrial",
]

LOG_SHEET = "Sent Log"

CALL_LOG_SHEET = "Call Logs"


# =====================================================================
# LINKS
# =====================================================================

PORTAL = (
    "https://protect.checkpoint.com/v2/r05/"
    "___https:/ocs.iitd.ac.in/portal/recruiter/auth___."
    "YXBzMTphZGl0eWFiaXJsYW1hbmFnZW1lbnQ6YzpvOmNlZmVhYjc0NTgyZTJkNmRmZjA5ZTlkYjk3NmMwYzgwOjc6"
    "OGNmNDo3MGRlZDJlNDBkMDYzMzYxNzgzMmFlOTAwOWJmOThhOTRjMmIzMzlmMzVhZDMwMjJhNGYyMzM0NDQ3YWVhMjY4Omg6VDpO"
)

TUTORIAL = (
    "https://protect.checkpoint.com/v2/r05/"
    "___https:/owncloud.iitd.ac.in/nextcloud/index.php/s/"
    "8WAXMCBZ63zqiaP___."
    "YXBzMTphZGl0eWFiaXJsYW1hbmFnZW1lbnQ6YzpvOmNlZmVhYjc0NTgyZTJkNmRmZjA5ZTlkYjk3"
    "NmMwYzgwOjc6ZDZmMjoyZWNhYWYyNjhlMGEyMDUwODA5MDkxMWU2Y2E2MjUzMzlmM2FkZjM1NTEw"
    "M2NhNWQwMmFlYjM4ODcxOTYwZmIzOmg6VDpO"
)

DOWNLOADS = "https://ocs.iitd.ac.in/downloads"


# =====================================================================
# EXCEPTIONS
# =====================================================================

class WebmailLoginError(Exception):
    """Raised when Roundcube login fails."""


class WebmailSendError(Exception):
    """Raised when Roundcube cannot send a message."""


class RateLimitHit(Exception):
    """Raised when SMTP reports a rate limit."""


# =====================================================================
# GENERAL HELPERS
# =====================================================================

def clean(value):
    """Convert a value to clean string form."""
    if value is None:
        return ""

    return str(value).replace("\xa0", " ").strip()


def set_cell(ws, row, column, value):
    """Safely write a workbook cell."""
    ws.cell(row=row, column=column).value = value


def norm(value):
    """
    Normalize company names for comparison.

    Aliases intentionally mirror the original resolver.
    """
    value = clean(value).lower()

    value = re.sub(
        r"[^a-z0-9]+",
        " ",
        value,
    ).strip()

    aliases = {
        "airbus india": "airbus",
        "airbus": "airbus",

        "boeing india": "boeing",
        "boeing": "boeing",
        "boeing india defence": "boeing",

        "ge aerospace": "ge aerospace",

        "rolls royce": "rolls royce",

        "pratt whitney": "pratt whitney",

        "collins aerospace": "collins aerospace",

        "cummins india": "cummins",
        "cummins": "cummins",

        "mahindra mahindra": "mahindra",
        "mahindra": "mahindra",

        "tata motors": "tata motors",
        "tata steel": "tata steel",

        "hyundai motor india": "hyundai",
        "hyundai": "hyundai",

        "alstom india": "alstom",
        "alstom": "alstom",

        "bajaj auto": "bajaj auto",

        "dassault systemes": "dassault",
        "dassault syst mes": "dassault",
        "dassault": "dassault",

        "siemens energy": "siemens energy",

        "ge vernova": "ge vernova",

        "jsw energy": "jsw energy",

        "bellatrix aerospace": "bellatrix",

        "reliance new energy": "reliance new energy",

        "valeo": "valeo",

        "bosch india": "bosch",
        "bosch": "bosch",

        "mtar technologies": "mtar",

        "mercedes benz r d india mbrdi": "mercedes",
        "mercedes benz r d": "mercedes",
        "mercedes benz r d india": "mercedes",
        "mercedes": "mercedes",

        "honeywell aerospace": "honeywell",
        "honeywell": "honeywell",

        "tata advanced systems": "tasl",
        "tasl": "tasl",

        "larsen toubro": "l t",
        "l t": "l t",
    }

    return aliases.get(value, value)


EMAIL_RE = re.compile(
    r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}",
)


def emails_in(value):
    """Return email addresses found inside arbitrary text."""
    return EMAIL_RE.findall(clean(value))


def unique_emails(emails):
    """Deduplicate email addresses while preserving order."""
    result = []
    seen = set()

    for email_addr in emails or []:
        email_addr = clean(email_addr)

        if not email_addr:
            continue

        if email_addr.lower() == "have to find":
            continue

        key = email_addr.lower()

        if key not in seen:
            seen.add(key)
            result.append(email_addr)

    return result


def detect_col(ws, header):
    """Find a column by header name."""
    target = clean(header).lower()

    for column in range(1, ws.max_column + 1):
        value = clean(
            ws.cell(row=1, column=column).value
        ).lower()

        if value == target:
            return column

    return None


# =====================================================================
# COMPANY RESOLVER
# =====================================================================

FREE_DOMAINS = {
    "gmail",
    "yahoo",
    "hotmail",
    "outlook",
    "rediffmail",
    "live",
    "icloud",
    "protonmail",
    "163",
    "ymail",
}


def _domain_root(email_addr):
    """
    Extract the first domain component.

    Example:
        hr@airbus.com -> airbus
        hr@airbus.co.in -> airbus
    """
    try:
        domain = email_addr.split("@", 1)[1].lower()
        root = domain.split(".", 1)[0]

        if root in FREE_DOMAINS:
            return None

        return root

    except (IndexError, AttributeError):
        return None


class CompanyResolver:
    """
    Resolves companies using both normalized names and email domains.
    """

    def __init__(self):
        self.domain_to_key = {}
        self.name_to_key = {}

    def key(self, name, emails):
        nkey = norm(name)

        roots = {
            root
            for root in (
                _domain_root(e)
                for e in (emails or [])
            )
            if root
        }

        # Existing domain match wins.
        for root in roots:
            if root in self.domain_to_key:
                key = self.domain_to_key[root]

                self.name_to_key[nkey] = key

                for root2 in roots:
                    self.domain_to_key.setdefault(root2, key)

                return key

        # Existing name match.
        key = self.name_to_key.get(nkey, nkey)

        self.name_to_key[nkey] = key

        for root in roots:
            self.domain_to_key.setdefault(root, key)

        return key

    def key_for_name_only(self, name):
        normalized = norm(name)
        return self.name_to_key.get(
            normalized,
            normalized,
        )

    def key_for_emails(self, emails):
        for email_addr in emails or []:
            root = _domain_root(email_addr)

            if root and root in self.domain_to_key:
                return self.domain_to_key[root]

        return None


# =====================================================================
# ROUNDCUBE LOGIN
# =====================================================================

def _extract_roundcube_token(html):
    """
    Extract CSRF/request token from several Roundcube variants.
    """

    patterns = [
        r'name=["\']_token["\']\s+value=["\']([^"\']+)["\']',
        r'name=["\']_token["\'][^>]+value=["\']([^"\']+)["\']',
        r'"request_token"\s*:\s*"([^"]+)"',
        r'rcmail\.set_env\(\s*["\']request_token["\']\s*,\s*["\']([^"\']+)["\']',
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            html or "",
            re.I | re.S,
        )

        if match:
            return match.group(1)

    return None


def roundcube_login(username, password):
    """
    Log into IIT Delhi Roundcube.

    Returns:
        requests.Session
    """

    if not username:
        raise WebmailLoginError(
            "Roundcube username is empty."
        )

    if not password:
        raise WebmailLoginError(
            "Roundcube password is empty."
        )

    session = requests.Session()

    try:
        login_page = session.get(
            WEBMAIL_URL,
            timeout=TIMEOUT,
        )

        login_page.raise_for_status()

    except requests.RequestException as exc:
        session.close()

        raise WebmailLoginError(
            f"Could not open IITD webmail: {exc}"
        ) from exc

    token = _extract_roundcube_token(
        login_page.text
    )

    if not token:
        session.close()

        raise WebmailLoginError(
            "IITD webmail login token was not found."
        )

    payload = {
        "_token": token,
        "_task": "login",
        "_action": "login",
        "_timezone": "Asia/Kolkata",
        "_url": "",
        "_user": username,
        "_pass": password,
    }

    try:
        response = session.post(
            urljoin(
                WEBMAIL_URL,
                "?_task=login",
            ),
            data=payload,
            timeout=TIMEOUT,
            allow_redirects=True,
        )

        response.raise_for_status()

    except requests.RequestException as exc:
        session.close()

        raise WebmailLoginError(
            f"IITD webmail login request failed: {exc}"
        ) from exc

    final_url = response.url or ""

    if "_task=mail" not in final_url:
        session.close()

        raise WebmailLoginError(
            "IITD webmail login was not accepted. "
            "Check the username, password, or MFA."
        )

    return session


def check_roundcube_login(password):
    """
    Validate Roundcube credentials.
    """

    session = roundcube_login(
        WEBMAIL_USERNAME,
        password,
    )

    try:
        response = session.get(
            urljoin(
                WEBMAIL_URL,
                "?_task=mail",
            ),
            timeout=TIMEOUT,
        )

        response.raise_for_status()

        if "_task=mail" not in response.url:
            raise WebmailLoginError(
                "Roundcube session was not available after login."
            )

        return True

    except requests.RequestException as exc:
        raise WebmailLoginError(
            f"Roundcube session check failed: {exc}"
        ) from exc

    finally:
        session.close()


# =====================================================================
# ROUNDCUBE COMPOSE
# =====================================================================

def _extract_compose_id(html):
    patterns = [
        r'name=["\']_id["\']\s+value=["\']([^"\']*)["\']',
        r'name=["\']_id["\'][^>]+value=["\']([^"\']*)["\']',
        r'["_\']id["\']\s*[:=]\s*["\']([^"\']+)["\']',
        r'_id=([A-Za-z0-9_-]+)',
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            html or "",
            re.I | re.S,
        )

        if match:
            return match.group(1)

    return ""


def _extract_roundcube_identity(html):
    """
    Extract sender identity from the compose page.
    """

    # Preferred form: select[name="_from"]
    select_match = re.search(
        r'<select[^>]+name=["\']_from["\'][^>]*>'
        r'(.*?)'
        r'</select>',
        html or "",
        re.I | re.S,
    )

    if select_match:
        select_html = select_match.group(1)

        # Selected option.
        selected_match = re.search(
            r'<option[^>]+value=["\']([^"\']+)["\'][^>]*'
            r'\bselected\b',
            select_html,
            re.I | re.S,
        )

        if selected_match:
            return selected_match.group(1)

        # Sometimes selected appears before value.
        selected_match = re.search(
            r'<option[^>]*\bselected\b[^>]+'
            r'value=["\']([^"\']+)["\']',
            select_html,
            re.I | re.S,
        )

        if selected_match:
            return selected_match.group(1)

        # Fall back to first option.
        first_match = re.search(
            r'<option[^>]+value=["\']([^"\']+)["\']',
            select_html,
            re.I | re.S,
        )

        if first_match:
            return first_match.group(1)

    # Input variant.
    input_patterns = [
        r'name=["\']_from["\']\s+value=["\']([^"\']+)["\']',
        r'name=["\']_from["\'][^>]+value=["\']([^"\']+)["\']',
        r'value=["\']([^"\']+)["\'][^>]+name=["\']_from["\']',
    ]

    for pattern in input_patterns:
        match = re.search(
            pattern,
            html or "",
            re.I | re.S,
        )

        if match:
            return match.group(1)

    return ""


def roundcube_compose_session(password):
    """
    Login and open a Roundcube compose session.

    Returns:
        (session, token)
    """

    session = roundcube_login(
        WEBMAIL_USERNAME,
        password,
    )

    try:
        compose_url = urljoin(
            WEBMAIL_URL,
            "?_task=mail&_action=compose",
        )

        compose_page = session.get(
            compose_url,
            timeout=TIMEOUT,
        )

        compose_page.raise_for_status()

        html = compose_page.text or ""

        token = _extract_roundcube_token(html)

        if not token:
            raise WebmailSendError(
                "Could not find Roundcube CSRF token on compose page."
            )

        compose_id = _extract_compose_id(html)

        identity = _extract_roundcube_identity(html)

        if not identity:
            raise WebmailSendError(
                "Could not find Roundcube sender identity "
                "on compose page."
            )

        session.roundcube_compose_id = compose_id
        session.roundcube_identity = identity

        logger.info(
            "Roundcube compose session created: id=%s",
            compose_id,
        )

        return session, token

    except Exception:
        session.close()
        raise


# =====================================================================
# ROUNDCUBE ATTACHMENT
# =====================================================================

def _extract_attachment_token(text, filename):
    """
    Try to extract an attachment identifier from a Roundcube upload
    response.
    """

    text = text or ""

    patterns = [
        r'"_attachments"\s*:\s*"([^"]+)"',
        r'"attachment"\s*:\s*"([^"]+)"',
        r'"id"\s*:\s*"([^"]+)"',
        r'"name"\s*:\s*"([^"]+)"',
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            re.I | re.S,
        )

        if match:
            value = match.group(1)

            if value:
                return value

    # Roundcube JavaScript callback.
    callback_match = re.search(
        r'add2attachment_list\((.*?)\)',
        text,
        re.I | re.S,
    )

    if callback_match:
        inside = callback_match.group(1)

        quoted = re.findall(
            r'["\']([^"\']+)["\']',
            inside,
        )

        # Prefer anything containing the filename.
        for candidate in quoted:
            if filename.lower() in candidate.lower():
                return candidate

        # Otherwise prefer longer token-like values.
        for candidate in quoted:
            if re.search(
                r"[A-Za-z0-9_-]{8,}",
                candidate,
            ):
                return candidate

    return ""


def roundcube_upload_attachment(
    session,
    token,
    compose_id,
    brochure_path,
):
    """
    Upload an attachment to the current compose session.

    Returns:
        attachment token or empty string.
    """

    if not brochure_path:
        return ""

    if not os.path.isfile(brochure_path):
        raise WebmailSendError(
            f"Brochure file not found: {brochure_path}"
        )

    filename = os.path.basename(
        brochure_path
    )

    upload_url = urljoin(
        WEBMAIL_URL,
        (
            "?_task=mail"
            "&_action=upload"
            f"&_id={compose_id}"
            "&_unlock=loading"
        ),
    )

    headers = {
        "X-Roundcube-Request": token,
        "Referer": urljoin(
            WEBMAIL_URL,
            (
                "?_task=mail"
                "&_action=compose"
                f"&_id={compose_id}"
            ),
        ),
    }

    try:
        with open(
            brochure_path,
            "rb",
        ) as file_obj:

            files = {
                "_attachments[]": (
                    filename,
                    file_obj,
                    "application/pdf",
                ),
            }

            data = {
                "_token": token,
                "_id": compose_id,
            }

            response = session.post(
                upload_url,
                data=data,
                files=files,
                headers=headers,
                timeout=ATTACHMENT_UPLOAD_TIMEOUT,
            )

        response.raise_for_status()

    except (OSError, requests.RequestException) as exc:
        raise WebmailSendError(
            f"Roundcube attachment upload failed: {exc}"
        ) from exc

    response_text = response.text or ""

    token_value = _extract_attachment_token(
        response_text,
        filename,
    )

    if token_value:
        return token_value

    # Some Roundcube installations attach the uploaded file to the
    # compose session server-side.
    lower_text = response_text.lower()

    if (
        "error" not in lower_text
        and "failed" not in lower_text
    ):
        logger.warning(
            "Attachment uploaded but no explicit "
            "attachment token was found."
        )
        return ""

    raise WebmailSendError(
        "Roundcube attachment upload returned an "
        f"unexpected response: {response_text[:1000]}"
    )


# =====================================================================
# EMAIL BODY
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

- Registration & JNF Creation Tutorial: Watch Here
{TUTORIAL}

- Placement Brochure & Flyer:
{DOWNLOADS}

- Important Timelines:
{DOWNLOADS}

We look forward to a meaningful collaboration and welcoming your organisation this season.

For any queries or assistance, feel free to contact the undersigned:

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
Nucleus Team Member (PG)
Office of Career Services
Indian Institute of Technology Delhi
"""


# =====================================================================
# EMAIL MESSAGE BUILDER
# =====================================================================

def build_message(
    company,
    recipients,
    brochure_path=None,
):
    """
    Build a raw RFC-compliant email message.
    """

    body = body_for(company)

    if brochure_path and os.path.isfile(
        brochure_path
    ):
        msg = MIMEMultipart()

        msg.attach(
            MIMEText(
                body,
                "plain",
                "utf-8",
            )
        )

        with open(
            brochure_path,
            "rb",
        ) as file_obj:

            part = MIMEApplication(
                file_obj.read(),
                _subtype="pdf",
            )

        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=os.path.basename(
                brochure_path
            ),
        )

        msg.attach(part)

    else:
        msg = MIMEText(
            body,
            "plain",
            "utf-8",
        )

    msg["Subject"] = SUBJECT

    msg["From"] = formataddr(
        (
            FROM_NAME,
            FROM_ADDR,
        )
    )

    msg["To"] = ", ".join(
        unique_emails(recipients)
    )

    if CC:
        msg["Cc"] = ", ".join(CC)

    msg["Date"] = formatdate(
        localtime=True
    )

    msg["Message-ID"] = make_msgid(
        domain="mech.iitd.ac.in"
    )

    return msg.as_string()


# =====================================================================
# ROUNDCUBE SEND
# =====================================================================

def send_one_via_roundcube(
    password,
    recipients,
    company,
    subject=None,
    body=None,
    brochure_path=None,
):
    """
    Send one message through Roundcube.

    Roundcube itself stores the sent message in Sent.
    """

    recipients = unique_emails(
        recipients
    )

    if not recipients:
        raise WebmailSendError(
            "No recipient email addresses were supplied."
        )

    session = None

    try:
        session, token = roundcube_compose_session(
            password
        )

        subject = subject or SUBJECT
        body = body or body_for(company)

        to_field = ", ".join(
            recipients
        )

        cc_field = ", ".join(CC)
        bcc_field = ", ".join(BCC)

        compose_id = getattr(
            session,
            "roundcube_compose_id",
            "",
        )

        identity = getattr(
            session,
            "roundcube_identity",
            "",
        )

        if not identity:
            raise WebmailSendError(
                "Roundcube sender identity is empty."
            )

        attachment_token = ""

        if brochure_path:
            attachment_token = (
                roundcube_upload_attachment(
                    session=session,
                    token=token,
                    compose_id=compose_id,
                    brochure_path=brochure_path,
                )
            )

        send_url = urljoin(
            WEBMAIL_URL,
            (
                "?_task=mail"
                "&_unlock=loading"
                "&_framed=1"
                "&_action=send"
            ),
        )

        payload = {
            "_token": token,
            "_task": "mail",
            "_action": "send",

            "_id": compose_id,
            "_from": identity,

            "_to": to_field,
            "_cc": cc_field,
            "_bcc": bcc_field,

            "_subject": subject,
            "_message": body,

            "_is_html": "0",
            "_priority": "0",
            "_store_target": "Sent",

            "_draft_saveid": "",

            "_attachments": attachment_token,

            "_references": "",
            "_in_reply_to": "",
            "_reply_uid": "",
            "_forward_uid": "",
            "_draft_uid": "",
        }

        headers = {
            "Referer": urljoin(
                WEBMAIL_URL,
                "?_task=mail&_action=compose",
            ),
            "X-Roundcube-Request": token,
        }

        response = session.post(
            send_url,
            data=payload,
            headers=headers,
            timeout=TIMEOUT,
            allow_redirects=True,
        )

        response.raise_for_status()

        response_text = response.text or ""

        lower_text = response_text.lower()

        success_markers = [
            "sent_successfully",
            "message sent successfully",
            "message_sent",
            "message sent",
        ]

        if any(
            marker in lower_text
            for marker in success_markers
        ):
            return True

        # Some Roundcube versions redirect to the mail page after
        # successful sending and do not return the exact marker.
        if (
            response.url
            and "_task=mail" in response.url
            and response.status_code in (200, 302)
            and "error" not in lower_text
        ):
            logger.info(
                "Roundcube returned a successful-looking "
                "mail response."
            )
            return True

        message_matches = re.findall(
            r"(?:show_message|display_message)"
            r"\((.*?)\)",
            response_text,
            re.I | re.S,
        )

        if message_matches:
            detail = message_matches[-1][:1000]
        else:
            detail = (
                f"HTTP {response.status_code}; "
                "Roundcube did not confirm send. "
                f"Response: {response_text[:1000]}"
            )

        raise WebmailSendError(detail)

    except requests.RequestException as exc:
        raise WebmailSendError(
            f"Roundcube network error: {exc}"
        ) from exc

    except WebmailSendError:
        raise

    except Exception as exc:
        raise WebmailSendError(
            f"Roundcube send failed: {exc}"
        ) from exc

    finally:
        if session is not None:
            try:
                session.close()
            except Exception:
                pass


# =====================================================================
# SMTP FALLBACK
# =====================================================================

def is_rate_limit(error):
    text = str(error).lower()

    return any(
        phrase in text
        for phrase in (
            "sending rate too high",
            "policy rejection",
            "450",
            "4.7.1",
        )
    )


def send_one(
    password,
    recipients,
    raw_message,
):
    """
    SMTP sender retained for compatibility with the original program.
    """

    recipients = unique_emails(
        recipients
    )

    if not recipients:
        return False

    context = ssl.create_default_context()

    last_error = ""

    for attempt in range(
        1,
        MAX_RETRIES + 1,
    ):
        try:
            if SMTP_PORT == 465:
                with smtplib.SMTP_SSL(
                    SMTP_HOST,
                    SMTP_PORT,
                    context=context,
                    timeout=TIMEOUT,
                ) as server:

                    server.login(
                        FROM_ADDR,
                        password,
                    )

                    server.sendmail(
                        FROM_ADDR,
                        recipients + CC + BCC,
                        raw_message,
                    )

            else:
                with smtplib.SMTP(
                    SMTP_HOST,
                    SMTP_PORT,
                    timeout=TIMEOUT,
                ) as server:

                    server.ehlo()

                    server.starttls(
                        context=context
                    )

                    server.ehlo()

                    server.login(
                        FROM_ADDR,
                        password,
                    )

                    server.sendmail(
                        FROM_ADDR,
                        recipients + CC + BCC,
                        raw_message,
                    )

            return True

        except Exception as exc:
            last_error = str(exc)

            logger.error(
                "SMTP send failed on attempt %s: %s",
                attempt,
                last_error,
            )

            if is_rate_limit(exc):
                raise RateLimitHit(
                    last_error
                )

            if attempt < MAX_RETRIES:
                time.sleep(15)

    logger.error(
        "SMTP send ultimately failed: %s",
        last_error,
    )

    return False


# =====================================================================
# IMAP SENT SAVE
# =====================================================================

def save_to_sent(
    password,
    raw_message,
):
    """
    Save a raw message manually into Sent.

    Normally Roundcube handles this itself.
    """

    imap = None

    try:
        imap = imaplib.IMAP4_SSL(
            IMAP_HOST,
            IMAP_PORT,
            timeout=TIMEOUT,
        )

        imap.login(
            FROM_ADDR,
            password,
        )

        status, _ = imap.append(
            SENT_FOLDER,
            "\\Seen",
            imaplib.Time2Internaldate(
                time.time()
            ),
            raw_message.encode(
                "utf-8"
            ),
        )

        return status == "OK"

    except Exception as exc:
        logger.error(
            "Could not save message to Sent: %s",
            exc,
        )

        return False

    finally:
        if imap is not None:
            try:
                imap.logout()
            except Exception:
                pass


# =====================================================================
# SENT LOG
# =====================================================================

LOG_HEADERS = [
    "Company Name",
    "Mail Sent Flag",
    "Mail Sent Date/Time",
    "Emails Sent To",
    "Mail Reply",
    "Phonic Conversation",
    "Delivery Status",
    "Bounced Emails",
]


def get_log_sheet(wb):
    """
    Get/create Sent Log and guarantee required headers.
    """

    if LOG_SHEET in wb.sheetnames:
        ws = wb[LOG_SHEET]
    else:
        ws = wb.create_sheet(LOG_SHEET)

    cols = {}

    for header in LOG_HEADERS:
        found_col = None

        for column in range(
            1,
            ws.max_column + 2,
        ):
            value = ws.cell(
                row=1,
                column=column,
            ).value

            if (
                value
                and clean(value).lower()
                == header.lower()
            ):
                found_col = column
                break

        if found_col is None:
            found_col = (
                ws.max_column + 1
            )

            set_cell(
                ws,
                1,
                found_col,
                header,
            )

        cols[header] = found_col

    return ws, cols


def read_log_state(
    ws,
    cols,
    resolver,
):
    """
    Read Sent Log into an in-memory state dictionary.
    """

    state = {}

    for row in range(
        2,
        ws.max_row + 1,
    ):
        company = clean(
            ws.cell(
                row=row,
                column=cols["Company Name"],
            ).value
        )

        if not company:
            continue

        tried = emails_in(
            clean(
                ws.cell(
                    row=row,
                    column=cols["Emails Sent To"],
                ).value
            )
        )

        sent_flag = clean(
            ws.cell(
                row=row,
                column=cols["Mail Sent Flag"],
            ).value
        ).upper()

        sent = sent_flag.startswith(
            "SENT"
        )

        delivery = clean(
            ws.cell(
                row=row,
                column=cols["Delivery Status"],
            ).value
        ).lower()

        entry = {
            "row": row,
            "sent": sent,
            "delivery": delivery,
            "emails_tried": {
                email_addr.lower()
                for email_addr in tried
            },
            "display": company,
        }

        name_key = resolver.key_for_name_only(
            company
        )

        state[name_key] = entry

        email_key = resolver.key_for_emails(
            tried
        )

        if email_key:
            state[email_key] = entry

    return state


def log_sent(
    ws,
    cols,
    state,
    resolver,
    company_display,
    emails,
):
    """
    Mark a company as SENT and merge recipient emails.
    """

    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M"
    )

    emails = unique_emails(
        emails
    )

    key = (
        resolver.key_for_emails(emails)
        or resolver.key_for_name_only(
            company_display
        )
    )

    target_row = None

    if key in state:
        target_row = state[key]["row"]

    else:
        for value in state.values():
            if norm(
                value["display"]
            ) == norm(company_display):

                target_row = value["row"]
                break

    if target_row:
        row = target_row

    else:
        row = ws.max_row + 1

        set_cell(
            ws,
            row,
            cols["Company Name"],
            company_display,
        )

        state[key] = {
            "row": row,
            "sent": True,
            "delivery": "pending",
            "emails_tried": set(),
            "display": company_display,
        }

    set_cell(
        ws,
        row,
        cols["Mail Sent Flag"],
        "SENT",
    )

    set_cell(
        ws,
        row,
        cols["Mail Sent Date/Time"],
        timestamp,
    )

    old_emails = unique_emails(
        emails_in(
            clean(
                ws.cell(
                    row=row,
                    column=cols["Emails Sent To"],
                ).value
            )
        )
    )

    merged = unique_emails(
        old_emails + emails
    )

    set_cell(
        ws,
        row,
        cols["Emails Sent To"],
        ", ".join(merged),
    )

    if not clean(
        ws.cell(
            row=row,
            column=cols["Delivery Status"],
        ).value
    ):
        set_cell(
            ws,
            row,
            cols["Delivery Status"],
            "Pending",
        )

    state[key]["row"] = row
    state[key]["sent"] = True
    state[key]["emails_tried"].update(
        e.lower()
        for e in emails
    )

    return timestamp


# =====================================================================
# CALL LOG HELPERS
# =====================================================================

CALL_LOG_FIELD_ALIASES = {
    "Company": [
        "company",
        "company name",
    ],

    "Phone Number": [
        "phone number",
        "phone",
        "contact no.",
        "contact number",
    ],

    "Incident": [
        "incident",
        "status",
        "notes",
        "remark",
        "remarks",
    ],

    "Date": [
        "date",
        "date/time",
        "date time",
    ],

    "Caller Name": [
        "caller name",
    ],

    "Success Flag": [
        "success flag",
    ],

    "HR Name": [
        "hr name",
    ],

    "HR Email": [
        "hr email",
    ],
}


def _header_key(value):
    """
    Normalize workbook headers.
    """

    return re.sub(
        r"[^a-z0-9]+",
        " ",
        clean(value).lower(),
    ).strip()


def get_call_log_sheet(wb):
    """
    Get/create Call Logs and ensure all required columns exist.
    """

    target = CALL_LOG_SHEET.strip().lower()

    existing_name = None

    for name in wb.sheetnames:
        if name.strip().lower() == target:
            existing_name = name
            break

    if existing_name:
        ws = wb[existing_name]
    else:
        ws = wb.create_sheet(
            CALL_LOG_SHEET
        )

    header_map = {}

    for column in range(
        1,
        ws.max_column + 1,
    ):
        value = ws.cell(
            row=1,
            column=column,
        ).value

        key = _header_key(value)

        if key:
            header_map[key] = column

    cols = {}

    for field, aliases in CALL_LOG_FIELD_ALIASES.items():

        found_col = None

        for alias in aliases:
            alias_key = _header_key(
                alias
            )

            if alias_key in header_map:
                found_col = header_map[
                    alias_key
                ]
                break

        if found_col is None:
            found_col = (
                ws.max_column + 1
            )

            set_cell(
                ws,
                1,
                found_col,
                field,
            )

            header_map[
                _header_key(field)
            ] = found_col

        cols[field] = found_col

    return ws, cols


def append_call_logs(
    wb,
    entries,
):
    """
    Append call-log records and save workbook.

    Returns:
        number of rows added
    """

    if not entries:
        return 0

    ws, cols = get_call_log_sheet(
        wb
    )

    added = 0

    for entry in entries:

        row = ws.max_row + 1

        success_flag = clean(
            entry.get(
                "success_flag",
                "0",
            )
        )

        success_flag = (
            "1"
            if success_flag == "1"
            else "0"
        )

        set_cell(
            ws,
            row,
            cols["Company"],
            clean(
                entry.get(
                    "company",
                    "",
                )
            ),
        )

        set_cell(
            ws,
            row,
            cols["Phone Number"],
            clean(
                entry.get(
                    "phone",
                    "",
                )
            ),
        )

        set_cell(
            ws,
            row,
            cols["Incident"],
            clean(
                entry.get(
                    "incident",
                    "",
                )
            ),
        )

        set_cell(
            ws,
            row,
            cols["Date"],
            clean(
                entry.get(
                    "date",
                    "",
                )
            ),
        )

        set_cell(
            ws,
            row,
            cols["Caller Name"],
            clean(
                entry.get(
                    "caller_name",
                    "",
                )
            ),
        )

        set_cell(
            ws,
            row,
            cols["Success Flag"],
            success_flag,
        )

        set_cell(
            ws,
            row,
            cols["HR Name"],
            clean(
                entry.get(
                    "hr_name",
                    "",
                )
            ),
        )

        set_cell(
            ws,
            row,
            cols["HR Email"],
            clean(
                entry.get(
                    "hr_email",
                    "",
                )
            ),
        )

        added += 1

    wb.save()

    return added


# =====================================================================
# COMPANY DATA
# =====================================================================

def build_company_data(wb):
    """
    Read all company data from the four primary sheets plus Call Logs.

    Returns:
        companies
        order
        resolver
    """

    resolver = CompanyResolver()

    raw = {
        sheet: []
        for sheet in SHEETS
    }

    for sheet_name in SHEETS:

        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        name_col = detect_col(
            ws,
            "Company Name",
        )

        email_col = detect_col(
            ws,
            "HR Email",
        )

        if (
            name_col is None
            or email_col is None
        ):
            continue

        bucket = None

        for row in range(
            2,
            ws.max_row + 1,
        ):

            company_name = clean(
                ws.cell(
                    row=row,
                    column=name_col,
                ).value
            )

            if company_name:
                bucket = {
                    "name": company_name,
                    "emails": [],
                }

                raw[
                    sheet_name
                ].append(bucket)

            if bucket is None:
                continue

            bucket["emails"].extend(
                emails_in(
                    ws.cell(
                        row=row,
                        column=email_col,
                    ).value
                )
            )

    companies = {}

    order = {
        sheet: []
        for sheet in SHEETS
    }

    for sheet_name in SHEETS:

        seen = set()

        for bucket in raw[sheet_name]:

            key = resolver.key(
                bucket["name"],
                bucket["emails"],
            )

            if key not in companies:
                companies[key] = {
                    "display": bucket["name"],
                    "emails": [],
                    "sheets": set(),
                }

            companies[key]["sheets"].add(
                sheet_name
            )

            companies[key]["emails"].extend(
                bucket["emails"]
            )

            if key not in seen:
                order[
                    sheet_name
                ].append(
                    bucket["name"]
                )

                seen.add(key)

    # ---------------------------------------------------------------
    # Merge Call Logs
    # ---------------------------------------------------------------

    call_log_companies = (
        read_call_log_companies(
            wb,
            resolver,
        )
    )

    if call_log_companies:
        order["Call Logs"] = []

    for key, data in (
        call_log_companies.items()
    ):

        if key not in companies:
            companies[key] = {
                "display": data["display"],
                "emails": [],
                "sheets": set(),
            }

        companies[key]["sheets"].add(
            "Call Logs"
        )

        companies[key]["emails"].extend(
            data["emails"]
        )

        if data["display"] not in order[
            "Call Logs"
        ]:
            order[
                "Call Logs"
            ].append(
                data["display"]
            )

    # ---------------------------------------------------------------
    # Final cleanup
    # ---------------------------------------------------------------

    for key in companies:

        companies[key]["emails"] = (
            unique_emails(
                companies[key]["emails"]
            )
        )

    return (
        companies,
        order,
        resolver,
    )


# =====================================================================
# CONTINUOUS QUEUE
# =====================================================================

def build_continuous_queue(wb):
    """
    Build the next batch of companies according to PER_SHEET.

    Returns:
        queue, resolver
    """

    companies, order, resolver = (
        build_company_data(wb)
    )

    log_ws, cols = get_log_sheet(
        wb
    )

    state = read_log_state(
        log_ws,
        cols,
        resolver,
    )

    def skip_and_emails(
        key,
        sheet_emails,
        display_name,
    ):
        state_entry = state.get(key)

        if not state_entry:
            email_key = (
                resolver.key_for_emails(
                    sheet_emails
                )
            )

            if email_key:
                state_entry = state.get(
                    email_key
                )

        if not state_entry:
            for value in state.values():
                if norm(
                    value["display"]
                ) == norm(display_name):
                    state_entry = value
                    break

        if (
            not state_entry
            or not state_entry["sent"]
        ):
            return False, sheet_emails

        # If every previous address failed,
        # allow new addresses to be tried.
        if state_entry[
            "delivery"
        ] == "all failed":

            new_emails = [
                email_addr
                for email_addr in sheet_emails
                if email_addr.lower()
                not in state_entry[
                    "emails_tried"
                ]
            ]

            if new_emails:
                return False, new_emails

            return True, []

        return True, []

    pointers = {
        sheet: 0
        for sheet in SHEETS
    }

    used = set()

    queue = []

    while any(
        pointers[sheet]
        < len(order.get(sheet, []))
        for sheet in SHEETS
    ):

        added = False

        for sheet in SHEETS:

            picked = 0

            company_list = order.get(
                sheet,
                [],
            )

            while (
                pointers[sheet]
                < len(company_list)
                and picked < PER_SHEET
            ):

                display_name = company_list[
                    pointers[sheet]
                ]

                pointers[sheet] += 1

                key = resolver.key_for_name_only(
                    display_name
                )

                data = companies.get(
                    key
                )

                if data is None:
                    for (
                        possible_key,
                        possible_data,
                    ) in companies.items():

                        if norm(
                            possible_data[
                                "display"
                            ]
                        ) == norm(
                            display_name
                        ):
                            key = possible_key
                            data = possible_data
                            break

                if (
                    not key
                    or key in used
                    or data is None
                    or not data["emails"]
                ):
                    continue

                skip, emails_to_use = (
                    skip_and_emails(
                        key,
                        data["emails"],
                        display_name,
                    )
                )

                if (
                    skip
                    or not emails_to_use
                ):
                    used.add(key)
                    continue

                used.add(key)

                queue.append(
                    {
                        "sheet": sheet,
                        "company": data[
                            "display"
                        ],
                        "key": key,
                        "emails": emails_to_use,
                    }
                )

                picked += 1
                added = True

        if not added:
            break

    return queue, resolver


# =====================================================================
# CONTINUOUS SEND
# =====================================================================

def send_continuous_batch(
    wb,
    items,
    pw,
    brochure_path,
    progress_cb=None,
    delay=DELAY_SEC,
):
    """
    Send selected queue items one by one.

    Workbook is saved after every successful/failed attempt.
    """

    if not items:
        return []

    log_ws, cols = get_log_sheet(
        wb
    )

    _, _, resolver = (
        build_company_data(wb)
    )

    state = read_log_state(
        log_ws,
        cols,
        resolver,
    )

    results = []

    for index, item in enumerate(
        items
    ):

        company = clean(
            item.get(
                "company",
                "",
            )
        )

        recipients = unique_emails(
            item.get(
                "emails",
                [],
            )
        )

        if not company:
            result = {
                "company": "",
                "status": "FAILED",
                "detail": "Company name is empty.",
            }

            results.append(result)

            if progress_cb:
                progress_cb(result)

            continue

        if not recipients:
            result = {
                "company": company,
                "status": "FAILED",
                "detail": "No recipient email address.",
            }

            results.append(result)

            if progress_cb:
                progress_cb(result)

            continue

        try:
            send_one_via_roundcube(
                password=pw,
                recipients=recipients,
                company=company,
                subject=SUBJECT,
                body=body_for(company),
                brochure_path=brochure_path,
            )

        except Exception as exc:

            try:
                wb.save()
            except Exception:
                pass

            result = {
                "company": company,
                "status": "FAILED",
                "detail": str(exc),
            }

            results.append(result)

            if progress_cb:
                progress_cb(result)

            # Stop batch on send failure, matching original behavior.
            break

        timestamp = log_sent(
            log_ws,
            cols,
            state,
            resolver,
            company,
            recipients,
        )

        wb.save()

        result = {
            "company": company,
            "status": "SENT",
            "time": timestamp,
            "saved_to_sent": True,
        }

        results.append(result)

        if progress_cb:
            progress_cb(result)

        if index < len(items) - 1:
            time.sleep(
                max(0, delay)
            )

    return results


# =====================================================================
# COMPANY SEARCH
# =====================================================================

def find_matches(
    query,
    companies,
):
    """
    Search companies by:
    1. exact key
    2. substring
    3. word containment
    4. fuzzy match
    """

    query_normalized = norm(
        query
    )

    if not query_normalized:
        return []

    if query_normalized in companies:
        return [
            query_normalized
        ]

    matches = []

    for key, data in companies.items():

        display_normalized = norm(
            data["display"]
        )

        if (
            query_normalized in key
            or query_normalized
            in display_normalized
        ):
            matches.append(key)
            continue

        query_words = set(
            query_normalized.split()
        )

        company_words = set(
            display_normalized.split()
        )

        if (
            query_words
            and query_words.issubset(
                company_words
            )
        ):
            matches.append(key)

    if matches:
        return matches

    return difflib.get_close_matches(
        query_normalized,
        list(companies.keys()),
        n=8,
        cutoff=0.35,
    )


def search_company(
    wb,
    query,
):
    """
    Search companies and return send/log state.
    """

    companies, _, resolver = (
        build_company_data(wb)
    )

    log_ws, cols = get_log_sheet(
        wb
    )

    state = read_log_state(
        log_ws,
        cols,
        resolver,
    )

    matches = find_matches(
        query,
        companies,
    )

    result = []

    for key in matches:

        data = companies[key]

        state_entry = state.get(key)

        if not state_entry:
            for value in state.values():
                if norm(
                    value["display"]
                ) == norm(
                    data["display"]
                ):
                    state_entry = value
                    break

        result.append(
            {
                "key": key,
                "display": data[
                    "display"
                ],
                "emails": data[
                    "emails"
                ],
                "sheets": sorted(
                    data["sheets"]
                ),
                "already_sent": bool(
                    state_entry
                    and state_entry["sent"]
                ),
                "delivery": (
                    state_entry[
                        "delivery"
                    ]
                    if state_entry
                    else ""
                )
                or "",
                "tried": sorted(
                    state_entry[
                        "emails_tried"
                    ]
                )
                if state_entry
                else [],
            }
        )

    return result


# =====================================================================
# SINGLE SEND
# =====================================================================

def send_single(
    wb,
    key,
    company_display,
    recipients,
    brochure_path,
    pw,
):
    """
    Send one company email.
    """

    del key  # Kept for API compatibility.

    company_display = clean(
        company_display
    )

    recipients = unique_emails(
        recipients
    )

    if not company_display:
        return {
            "status": "FAILED",
            "detail": "Company name is empty.",
        }

    if not recipients:
        return {
            "status": "FAILED",
            "detail": "No recipient email addresses.",
        }

    companies, _, resolver = (
        build_company_data(wb)
    )

    log_ws, cols = get_log_sheet(
        wb
    )

    state = read_log_state(
        log_ws,
        cols,
        resolver,
    )

    try:
        send_one_via_roundcube(
            password=pw,
            recipients=recipients,
            company=company_display,
            subject=SUBJECT,
            body=body_for(
                company_display
            ),
            brochure_path=brochure_path,
        )

    except Exception as exc:

        try:
            wb.save()
        except Exception:
            pass

        return {
            "status": "FAILED",
            "detail": str(exc),
        }

    timestamp = log_sent(
        log_ws,
        cols,
        state,
        resolver,
        company_display,
        recipients,
    )

    wb.save()

    return {
        "status": "SENT",
        "time": timestamp,
        "saved_to_sent": True,
    }


# =====================================================================
# MESSAGE DECODING
# =====================================================================

def decode_hdr(value):
    """
    Decode MIME encoded email headers safely.
    """

    if not value:
        return ""

    output = []

    try:
        parts = decode_header(
            value
        )

        for text, encoding in parts:

            if isinstance(
                text,
                bytes,
            ):
                output.append(
                    text.decode(
                        encoding
                        or "utf-8",
                        errors="replace",
                    )
                )
            else:
                output.append(
                    str(text)
                )

    except Exception:
        return clean(value)

    return "".join(
        output
    )


def _decode_payload(part):
    """
    Decode an email MIME part safely.
    """

    try:
        payload = part.get_payload(
            decode=True
        )

        if payload:
            charset = (
                part.get_content_charset()
                or "utf-8"
            )

            return payload.decode(
                charset,
                errors="replace",
            )

    except Exception:
        pass

    try:
        payload = part.get_payload()

        if isinstance(
            payload,
            str,
        ):
            return payload

        return str(payload)

    except Exception:
        return ""


def whole_message_text(msg):
    """
    Extract useful searchable text from a complete email.
    """

    chunks = []

    headers = [
        "To",
        "Cc",
        "Subject",
        "X-Failed-Recipients",
        "Original-Recipient",
        "Final-Recipient",
    ]

    for header in headers:
        try:
            value = decode_hdr(
                msg.get(header)
            )

            if value:
                chunks.append(
                    f"{header}: {value}"
                )

        except Exception:
            pass

    if msg.is_multipart():

        for part in msg.walk():

            try:
                for header in (
                    "To",
                    "Original-Recipient",
                    "Final-Recipient",
                    "X-Failed-Recipients",
                ):

                    value = decode_hdr(
                        part.get(header)
                    )

                    if value:
                        chunks.append(
                            f"{header}: {value}"
                        )

            except Exception:
                pass

            content = _decode_payload(
                part
            )

            if content:
                chunks.append(
                    content
                )

    else:

        content = _decode_payload(
            msg
        )

        if content:
            chunks.append(
                content
            )

    text = "\n".join(
        chunks
    )

    # Add a basic HTML-stripped copy as well.
    stripped_html = re.sub(
        r"<[^>]+>",
        " ",
        text,
    )

    return (
        text
        + "\n"
        + stripped_html
    )


# =====================================================================
# BOUNCE DETECTION
# =====================================================================

def is_bounce(msg):
    """
    Determine whether an email looks like a bounce.
    """

    sender = decode_hdr(
        msg.get("From")
    ).lower()

    subject = decode_hdr(
        msg.get("Subject")
    ).lower()

    content_type = (
        msg.get_content_type()
        or ""
    ).lower()

    if (
        "postmaster@" in sender
        or "mailer-daemon" in sender
        or "mail delivery" in sender
    ):
        return True

    if any(
        phrase in subject
        for phrase in (
            "undeliverable",
            "delivery has failed",
            "delivery status notification",
            "returned mail",
            "mail delivery failed",
            "failure notice",
        )
    ):
        return True

    if (
        "report-type=delivery-status"
        in content_type
        or "multipart/report"
        in content_type
    ):
        return True

    return False


def failed_recips_full(msg):
    """
    Extract failed recipient addresses from a bounce.
    """

    text = whole_message_text(
        msg
    )

    failed = []

    def add_email(value):
        value = (
            clean(value)
            .strip("<>")
            .strip()
            .lower()
        )

        if (
            EMAIL_RE.fullmatch(value)
            and value not in failed
        ):
            failed.append(value)

    # RFC 3464.
    for match in re.finditer(
        r"Final-Recipient:\s*rfc822;\s*([^\s]+)",
        text,
        re.I,
    ):
        add_email(
            match.group(1)
        )

    # X-Failed-Recipients and Original-Recipient.
    for match in re.finditer(
        r"(?:X-Failed-Recipients|Original-Recipient)"
        r"[^\n:]*:\s*"
        r"(?:rfc822;)?\s*"
        r"([^\s,;]+)",
        text,
        re.I,
    ):
        add_email(
            match.group(1)
        )

    # Common human-readable bounce block.
    block = re.search(
        r"failed\s+to\s+these\s+recipients"
        r"[^\n]*\n"
        r"(.*?)(?:\n\s*\n|your\s+message)",
        text,
        re.I | re.S,
    )

    if block:
        for email_addr in EMAIL_RE.findall(
            block.group(1)
        ):
            add_email(
                email_addr
            )

    own_addresses = {
        address.lower()
        for address in (
            [FROM_ADDR]
            + CC
            + BCC
        )
    }

    return [
        email_addr
        for email_addr in failed
        if email_addr not in own_addresses
    ]


# =====================================================================
# CHECK BOUNCES
# =====================================================================

def check_bounces(
    wb,
    pw,
    days=14,
):
    """
    Scan recent INBOX messages for bounces and update Sent Log.
    """

    bounced = set()

    imap = None

    try:
        imap = imaplib.IMAP4_SSL(
            IMAP_HOST,
            IMAP_PORT,
            timeout=TIMEOUT,
        )

        imap.login(
            FROM_ADDR,
            pw,
        )

        status, _ = imap.select(
            "INBOX"
        )

        if status != "OK":
            raise RuntimeError(
                "Could not select INBOX."
            )

        since_date = (
            datetime.now()
            - timedelta(days=days)
        ).strftime(
            "%d-%b-%Y"
        )

        status, data = imap.search(
            None,
            f"(SINCE {since_date})",
        )

        if status != "OK":
            ids = []
        else:
            ids = (
                data[0].split()
                if data
                and data[0]
                else []
            )

        for message_id in ids:

            status, message_data = (
                imap.fetch(
                    message_id,
                    "(RFC822)",
                )
            )

            if (
                status != "OK"
                or not message_data
            ):
                continue

            raw = None

            for item in message_data:
                if (
                    isinstance(item, tuple)
                    and len(item) >= 2
                ):
                    raw = item[1]
                    break

            if not raw:
                continue

            try:
                msg = (
                    email.message_from_bytes(
                        raw
                    )
                )
            except Exception:
                continue

            if not is_bounce(msg):
                continue

            bounced.update(
                failed_recips_full(
                    msg
                )
            )

    finally:
        if imap is not None:
            try:
                imap.logout()
            except Exception:
                pass

    if not bounced:
        return {
            "scanned": len(ids)
            if "ids" in locals()
            else 0,
            "bounced": [],
            "updated_rows": [],
        }

    if LOG_SHEET not in wb.sheetnames:
        return {
            "scanned": len(ids),
            "bounced": sorted(
                bounced
            ),
            "updated_rows": [],
            "no_log": True,
        }

    ws, cols = get_log_sheet(
        wb
    )

    updated_rows = []

    for row in range(
        2,
        ws.max_row + 1,
    ):

        company = clean(
            ws.cell(
                row=row,
                column=cols[
                    "Company Name"
                ],
            ).value
        )

        if not company:
            continue

        sent_list = [
            email_addr.lower()
            for email_addr in emails_in(
                clean(
                    ws.cell(
                        row=row,
                        column=cols[
                            "Emails Sent To"
                        ],
                    ).value
                )
            )
        ]

        if not sent_list:
            continue

        failed = [
            email_addr
            for email_addr in sent_list
            if email_addr in bounced
        ]

        if not failed:
            status = "No bounce seen"

        elif len(failed) == len(
            sent_list
        ):
            status = "All Failed"

        else:
            status = "Partial"

        set_cell(
            ws,
            row,
            cols["Delivery Status"],
            status,
        )

        set_cell(
            ws,
            row,
            cols["Bounced Emails"],
            ", ".join(failed),
        )

        if failed:
            updated_rows.append(
                {
                    "company": company,
                    "status": status,
                    "bounced": failed,
                }
            )

    wb.save()

    return {
        "scanned": len(ids),
        "bounced": sorted(
            bounced
        ),
        "updated_rows": updated_rows,
    }


# =====================================================================
# HR CONTACT DIRECTORY
# =====================================================================

def build_hr_contact_directory(
    wb,
    companies,
    resolver,
):
    """
    Build:

        company_key -> [
            {
                sheet,
                company,
                name,
                email,
                phone,
            }
        ]
    """

    del companies  # Retained for API compatibility.

    contacts = {}

    for sheet_name in SHEETS:

        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        company_col = detect_col(
            ws,
            "Company Name",
        )

        hr_name_col = detect_col(
            ws,
            "HR Name",
        )

        hr_email_col = detect_col(
            ws,
            "HR Email",
        )

        hr_phone_col = detect_col(
            ws,
            "HR Phone",
        )

        if company_col is None:
            continue

        current_company_name = None
        current_company_key = None

        for row in range(
            2,
            ws.max_row + 1,
        ):

            company_name = clean(
                ws.cell(
                    row=row,
                    column=company_col,
                ).value
            )

            if company_name:

                current_company_name = (
                    company_name
                )

                current_company_key = (
                    resolver.key_for_name_only(
                        company_name
                    )
                )

            if current_company_key is None:
                continue

            hr_name = (
                clean(
                    ws.cell(
                        row=row,
                        column=hr_name_col,
                    ).value
                )
                if hr_name_col
                else ""
            )

            hr_email = (
                clean(
                    ws.cell(
                        row=row,
                        column=hr_email_col,
                    ).value
                )
                if hr_email_col
                else ""
            )

            hr_phone = (
                clean(
                    ws.cell(
                        row=row,
                        column=hr_phone_col,
                    ).value
                )
                if hr_phone_col
                else ""
            )

            if (
                not hr_name
                and not hr_email
                and not hr_phone
            ):
                continue

            placeholder_values = {
                "have to find",
                "production",
                "thermal",
                "industrial",
                "design",
                "na",
            }

            if (
                hr_name.lower()
                in placeholder_values
                and hr_email.lower()
                in placeholder_values
                and hr_phone.lower()
                in placeholder_values
            ):
                continue

            contacts.setdefault(
                current_company_key,
                [],
            ).append(
                {
                    "sheet": sheet_name,
                    "company": current_company_name,
                    "name": hr_name,
                    "email": hr_email,
                    "phone": hr_phone,
                }
            )

    # ---------------------------------------------------------------
    # Merge contacts from Call Logs.
    # ---------------------------------------------------------------

    call_log_companies = (
        read_call_log_companies(
            wb,
            resolver,
        )
    )

    for company_key, data in (
        call_log_companies.items()
    ):

        contacts.setdefault(
            company_key,
            [],
        )

        existing = {
            (
                row["sheet"].lower(),
                row["name"].lower(),
                row["email"].lower(),
                row["phone"].lower(),
            )
            for row in contacts[
                company_key
            ]
        }

        for row in data.get(
            "contacts",
            [],
        ):

            row_key = (
                row["sheet"].lower(),
                row["name"].lower(),
                row["email"].lower(),
                row["phone"].lower(),
            )

            if row_key not in existing:
                existing.add(row_key)

                contacts[
                    company_key
                ].append(row)

    return contacts


def hr_lookup(
    wb,
    query,
):
    """
    Search companies and return HR contacts.
    """

    companies, _, resolver = (
        build_company_data(wb)
    )

    contacts = (
        build_hr_contact_directory(
            wb,
            companies,
            resolver,
        )
    )

    matches = find_matches(
        query,
        companies,
    )

    result = []

    for key in matches:

        data = companies[key]

        result.append(
            {
                "key": key,
                "display": data[
                    "display"
                ],
                "sheets": sorted(
                    data["sheets"]
                ),
                "contacts": contacts.get(
                    key,
                    [],
                ),
            }
        )

    return result


# =====================================================================
# CALL LOG COMPANY READER
# =====================================================================

def read_call_log_companies(
    wb,
    resolver=None,
):
    """
    Read Call Logs.

    Returns:

        {
            company_key: {
                "display": "...",
                "emails": [...],
                "contacts": [...]
            }
        }
    """

    if CALL_LOG_SHEET not in wb.sheetnames:
        return {}

    if resolver is None:
        resolver = CompanyResolver()

    ws, cols = get_call_log_sheet(
        wb
    )

    result = {}

    for row in range(
        2,
        ws.max_row + 1,
    ):

        company = clean(
            ws.cell(
                row=row,
                column=cols[
                    "Company"
                ],
            ).value
        )

        if not company:
            continue

        phone = clean(
            ws.cell(
                row=row,
                column=cols[
                    "Phone Number"
                ],
            ).value
        )

        hr_name = clean(
            ws.cell(
                row=row,
                column=cols[
                    "HR Name"
                ],
            ).value
        )

        hr_email_raw = clean(
            ws.cell(
                row=row,
                column=cols[
                    "HR Email"
                ],
            ).value
        )

        found_emails = unique_emails(
            emails_in(
                hr_email_raw
            )
        )

        key = resolver.key(
            company,
            found_emails,
        )

        if key not in result:
            result[key] = {
                "display": company,
                "emails": [],
                "contacts": [],
            }

        result[key]["emails"].extend(
            found_emails
        )

        if (
            phone
            or hr_name
            or found_emails
        ):
            result[key][
                "contacts"
            ].append(
                {
                    "sheet": "Call Logs",
                    "company": company,
                    "name": hr_name,
                    "email": (
                        ", ".join(
                            found_emails
                        )
                        if found_emails
                        else hr_email_raw
                    ),
                    "phone": phone,
                }
            )

    # ---------------------------------------------------------------
    # Deduplicate
    # ---------------------------------------------------------------

    for key in result:

        result[key]["emails"] = (
            unique_emails(
                result[key]["emails"]
            )
        )

        seen = set()
        unique_contacts = []

        for contact in result[key][
            "contacts"
        ]:

            contact_key = (
                contact["sheet"].lower(),
                contact["company"].lower(),
                contact["name"].lower(),
                contact["email"].lower(),
                contact["phone"].lower(),
            )

            if contact_key not in seen:
                seen.add(contact_key)

                unique_contacts.append(
                    contact
                )

        result[key][
            "contacts"
        ] = unique_contacts

    return result


# =====================================================================
# RECONCILE SENT FOLDER
# =====================================================================

GREET_RE = re.compile(
    r"Dear\s+(.+?)\s+(?:Recruitment\s+)?Team\b",
    re.I | re.S,
)


def reconcile_sent(
    wb,
    pw,
    days=30,
):
    """
    Reconcile messages in Sent against workbook companies.
    """

    companies, _, resolver = (
        build_company_data(wb)
    )

    normalized_to_display = {
        norm(
            data["display"]
        ): data["display"]
        for data in companies.values()
    }

    imap = None

    hits = {}
    unattributed = []
    matched = 0
    ids = []

    try:
        imap = imaplib.IMAP4_SSL(
            IMAP_HOST,
            IMAP_PORT,
            timeout=TIMEOUT,
        )

        imap.login(
            FROM_ADDR,
            pw,
        )

        status, _ = imap.select(
            SENT_FOLDER
        )

        if status != "OK":
            raise RuntimeError(
                f"Could not select Sent folder: {SENT_FOLDER}"
            )

        since_date = (
            datetime.now()
            - timedelta(days=days)
        ).strftime(
            "%d-%b-%Y"
        )

        status, data = imap.search(
            None,
            f"(SINCE {since_date})",
        )

        if status == "OK" and data:
            ids = (
                data[0].split()
                if data[0]
                else []
            )

        our_addresses = {
            address.lower()
            for address in (
                [FROM_ADDR]
                + CC
                + BCC
            )
        }

        for message_id in ids:

            status, message_data = (
                imap.fetch(
                    message_id,
                    "(RFC822)",
                )
            )

            if (
                status != "OK"
                or not message_data
            ):
                continue

            raw = None

            for item in message_data:
                if (
                    isinstance(item, tuple)
                    and len(item) >= 2
                ):
                    raw = item[1]
                    break

            if not raw:
                continue

            try:
                msg = (
                    email.message_from_bytes(
                        raw
                    )
                )
            except Exception:
                continue

            subject = decode_hdr(
                msg.get("Subject")
            ).lower()

            text = whole_message_text(
                msg
            )

            combined_lower = (
                subject
                + "\n"
                + text.lower()
            )

            if (
                SUBJECT_KEY
                not in combined_lower
            ):
                continue

            matched += 1

            raw_recipient_text = (
                decode_hdr(
                    msg.get("To")
                )
                + " "
                + decode_hdr(
                    msg.get("Cc")
                )
            )

            recipients = []

            for address in emails_in(
                raw_recipient_text
            ):

                address_lower = (
                    address.lower()
                )

                if (
                    address_lower
                    not in our_addresses
                    and address_lower
                    not in recipients
                ):
                    recipients.append(
                        address_lower
                    )

            # Also inspect extracted To/Cc lines.
            for match in re.finditer(
                r"^(?:To|Cc):\s*(.+)$",
                text,
                re.I | re.M,
            ):

                for address in emails_in(
                    match.group(1)
                ):

                    address_lower = (
                        address.lower()
                    )

                    if (
                        address_lower
                        not in our_addresses
                        and address_lower
                        not in recipients
                    ):
                        recipients.append(
                            address_lower
                        )

            try:
                date_header = msg.get(
                    "Date"
                )

                if date_header:
                    date_string = (
                        parsedate_to_datetime(
                            date_header
                        ).strftime(
                            "%Y-%m-%d %H:%M"
                        )
                    )
                else:
                    date_string = ""

            except Exception:
                date_string = ""

            company_display = None

            # -------------------------------------------------------
            # First: email-domain/company resolver.
            # -------------------------------------------------------

            key = resolver.key_for_emails(
                recipients
            )

            if (
                key
                and key in companies
            ):
                company_display = (
                    companies[key][
                        "display"
                    ]
                )

            # -------------------------------------------------------
            # Second: greeting.
            # -------------------------------------------------------

            if company_display is None:

                greeting_match = (
                    GREET_RE.search(
                        text
                    )
                )

                if greeting_match:

                    greeting_company = clean(
                        greeting_match.group(
                            1
                        )
                    )

                    company_display = (
                        normalized_to_display.get(
                            norm(
                                greeting_company
                            ),
                            greeting_company,
                        )
                    )

            # -------------------------------------------------------
            # Third: search company name in body.
            # -------------------------------------------------------

            if company_display is None:

                for company_name in sorted(
                    normalized_to_display.values(),
                    key=len,
                    reverse=True,
                ):

                    if (
                        company_name.lower()
                        in text.lower()
                    ):
                        company_display = (
                            company_name
                        )
                        break

            # -------------------------------------------------------
            # No match.
            # -------------------------------------------------------

            if company_display is None:

                unattributed.append(
                    {
                        "to": decode_hdr(
                            msg.get("To")
                        ),
                        "date": date_string,
                    }
                )

                continue

            # -------------------------------------------------------
            # Aggregate.
            # -------------------------------------------------------

            normalized_company = norm(
                company_display
            )

            record = hits.setdefault(
                normalized_company,
                {
                    "display": company_display,
                    "emails": set(),
                    "date": date_string,
                },
            )

            record["emails"].update(
                recipients
            )

            if (
                date_string
                and (
                    not record["date"]
                    or date_string
                    > record["date"]
                )
            ):
                record["date"] = (
                    date_string
                )

    finally:
        if imap is not None:
            try:
                imap.logout()
            except Exception:
                pass

    # ---------------------------------------------------------------
    # Update Sent Log.
    # ---------------------------------------------------------------

    ws, cols = get_log_sheet(
        wb
    )

    log_by_name = {}

    for row in range(
        2,
        ws.max_row + 1,
    ):

        company_name = clean(
            ws.cell(
                row=row,
                column=cols[
                    "Company Name"
                ],
            ).value
        )

        if company_name:
            log_by_name[
                norm(company_name)
            ] = row

    added = 0
    updated = 0

    for normalized_company, record in (
        hits.items()
    ):

        emails_sorted = sorted(
            record["emails"]
        )

        if normalized_company in log_by_name:

            row = log_by_name[
                normalized_company
            ]

            existing_flag = clean(
                ws.cell(
                    row=row,
                    column=cols[
                        "Mail Sent Flag"
                    ],
                ).value
            ).upper()

            if not existing_flag.startswith(
                "SENT"
            ):
                updated += 1

        else:

            row = ws.max_row + 1

            set_cell(
                ws,
                row,
                cols["Company Name"],
                record["display"],
            )

            log_by_name[
                normalized_company
            ] = row

            added += 1

        set_cell(
            ws,
            row,
            cols["Mail Sent Flag"],
            "SENT",
        )

        existing_date = clean(
            ws.cell(
                row=row,
                column=cols[
                    "Mail Sent Date/Time"
                ],
            ).value
        )

        if (
            not existing_date
            and record["date"]
        ):
            set_cell(
                ws,
                row,
                cols[
                    "Mail Sent Date/Time"
                ],
                record["date"],
            )

        existing_emails = emails_in(
            clean(
                ws.cell(
                    row=row,
                    column=cols[
                        "Emails Sent To"
                    ],
                ).value
            )
        )

        existing_lower = {
            address.lower()
            for address in existing_emails
        }

        merged = list(
            existing_emails
        )

        for address in emails_sorted:
            if (
                address.lower()
                not in existing_lower
            ):
                merged.append(
                    address
                )
                existing_lower.add(
                    address.lower()
                )

        set_cell(
            ws,
            row,
            cols[
                "Emails Sent To"
            ],
            ", ".join(merged),
        )

        if not clean(
            ws.cell(
                row=row,
                column=cols[
                    "Delivery Status"
                ],
            ).value
        ):
            set_cell(
                ws,
                row,
                cols[
                    "Delivery Status"
                ],
                "Pending",
            )

    wb.save()

    return {
        "scanned": len(ids),
        "matched": matched,
        "added": added,
        "updated": updated,
        "unattributed": unattributed,
    }
