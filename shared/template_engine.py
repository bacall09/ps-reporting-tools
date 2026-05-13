"""
shared/template_engine.py
─────────────────────────────────────────────────────────────────────────────
PSPT Customer Engagement — Template Engine

Responsibilities:
  1. Load and validate YAML template libraries (welcome, post-session, lifecycle)
  2. Auto-fill placeholders from DRS / NS / SFDC session data
  3. Render final subject + body for consultant review before send
  4. Generate a unique send identifier (UUID) per email
  5. Record sends to the session log (st.session_state)
  6. Expose the Smartsheet writeback payload for milestone date updates

Gmail API integration (pending IT approval):
  - This module produces the final rendered email + metadata
  - Actual send is delegated to shared/gmail_api.py (stub included below)
  - On send confirmation, this module triggers the SS writeback and session log

Usage:
    from shared.template_engine import (
        get_welcome_template,
        get_post_session_templates,
        get_lifecycle_template,
        render_template,
        record_send,
        get_session_send_log,
    )
"""

import re
import uuid
import yaml
import datetime
import streamlit as st
from pathlib import Path
from typing import Optional

# ── Path constants ────────────────────────────────────────────────────────────

_BASE = Path(__file__).parent / "templates"
_WELCOME_PATH     = _BASE / "welcome"      / "welcome_templates.yaml"
_POST_SESSION_PATH = _BASE / "post_session" / "post_session_templates.yaml"
_LIFECYCLE_PATH   = _BASE / "lifecycle"    / "lifecycle_templates.yaml"

# ── Session log key ───────────────────────────────────────────────────────────

_SESSION_LOG_KEY = "_email_send_log"

# ── Smartsheet field names ────────────────────────────────────────────────────
# These are the column names on the intermediate sheet as defined by VP PMO.
# The intermediate sheet does not use _OR suffixes — those are internal to the
# Blueprint layer. PSPT reads and writes the intermediate sheet only.
# Source of truth: SS_Data_Sheet_Template_with_notes____API_field_mapping_PMO_aligned.xlsx

SS_INTRO_EMAIL_SENT     = "Intro. Email Sent"
SS_CONFIG_ENABLEMENT    = "Enablement Session"
SS_WORKING_SESSION_1    = "Session #1"
SS_WORKING_SESSION_2    = "Session #2"
SS_UAT_SIGNOFF          = "UAT Signoff"
SS_PROD_CUTOVER         = "Prod Cutover"
SS_HYPERCARE_START      = "Hypercare Start"
SS_TRANSITION_SUPPORT   = "Transition to Support"
SS_GO_LIVE_DATE         = "Go-Live Date"       # Blueprint direct write — PMO approved


# ─────────────────────────────────────────────────────────────────────────────
# 1. Loaders
# ─────────────────────────────────────────────────────────────────────────────

def _load_yaml(path: Path) -> dict:
    """Load a YAML file, raise a clear error if missing."""
    if not path.exists():
        raise FileNotFoundError(
            f"Template file not found: {path}. "
            "Check that shared/templates/ is present in the deployment."
        )
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _welcome_library() -> dict:
    return _load_yaml(_WELCOME_PATH)


def _post_session_library() -> dict:
    return _load_yaml(_POST_SESSION_PATH)


def _lifecycle_library() -> dict:
    return _load_yaml(_LIFECYCLE_PATH)


# ─────────────────────────────────────────────────────────────────────────────
# 2. Template selectors
# ─────────────────────────────────────────────────────────────────────────────

def get_welcome_template(sku_key: str) -> Optional[dict]:
    """
    Return the welcome template dict for a given sku_key.
    sku_key should match the Product field from DRS (e.g. 'ZoneCapture').

    Returns None if no matching template is found (caller should fall back to
    showing a 'no template found' message and offering a blank draft).
    """
    lib = _welcome_library()
    for tmpl in lib.get("templates", []):
        if tmpl["sku_key"].lower() == sku_key.strip().lower():
            return tmpl
    return None


def list_welcome_templates() -> list[dict]:
    """Return all welcome templates as a list of {sku_key, display_name} dicts."""
    lib = _welcome_library()
    return [
        {"sku_key": t["sku_key"], "display_name": t["display_name"]}
        for t in lib.get("templates", [])
    ]


def get_post_session_templates(product_key: str) -> list[dict]:
    """
    Return all post-session templates for a product.
    product_key: 'capture' | 'approvals' | 'reconcile'
    """
    lib = _post_session_library()
    product = lib.get("products", {}).get(product_key.lower())
    if not product:
        return []
    return product.get("sessions", [])


def get_lifecycle_template(template_id: str) -> Optional[dict]:
    """
    Return a lifecycle template by its id.
    e.g. 'uat_signoff', 'go_live_hypercare_kickoff', 'hypercare_closure'
    """
    lib = _lifecycle_library()
    for tmpl in lib.get("templates", []):
        if tmpl["id"] == template_id:
            return tmpl
    return None


def list_lifecycle_templates() -> list[dict]:
    """Return all lifecycle templates as {id, category, name, trigger} dicts."""
    lib = _lifecycle_library()
    return [
        {
            "id": t["id"],
            "category": t["category"],
            "name": t["name"],
            "trigger": t["trigger"],
        }
        for t in lib.get("templates", [])
    ]


# ─────────────────────────────────────────────────────────────────────────────
# 3. Placeholder auto-fill
# ─────────────────────────────────────────────────────────────────────────────

def build_auto_context(project_row: dict, consultant_name: str, sfdc_contact: Optional[dict] = None) -> dict:
    """
    Build the auto-fill context dict from available data sources.

    project_row   : a single row from df_drs (as dict)
    consultant_name : logged-in user's name
    sfdc_contact  : dict from SFDC contacts (optional — may be None pre-Gmail API)

    Returns a dict of {PLACEHOLDER_KEY: resolved_value} for all auto-fillable fields.
    """
    customer = project_row.get("customer", "")
    product  = project_row.get("product", project_row.get("project_type", ""))

    # Generate the ps-support alias from customer name
    customer_slug = re.sub(r"[^A-Z0-9]", "", customer.upper())
    ps_support_email = f"ps-support+{customer_slug}@zoneandco.com"

    # Go-live and hypercare dates from DRS
    go_live_raw = project_row.get("go_live_date") or project_row.get("prod_cutover")
    go_live_str = ""
    hypercare_end_str = ""
    if go_live_raw:
        try:
            go_live_dt = (
                go_live_raw
                if isinstance(go_live_raw, datetime.date)
                else datetime.date.fromisoformat(str(go_live_raw)[:10])
            )
            go_live_str = go_live_dt.strftime("%d %B %Y")
            hypercare_end_str = (go_live_dt + datetime.timedelta(days=14)).strftime("%d %B %Y")
        except (ValueError, TypeError):
            pass

    ctx = {
        "CUSTOMER":               customer,
        "CUSTOMER_CONTACT_NAME":  sfdc_contact.get("contact_name", "") if sfdc_contact else "",
        "CONSULTANT_NAME":        consultant_name,
        "SENDER":                 consultant_name,
        "PRODUCT_NAME":           product,
        "PS_SUPPORT_EMAIL":       ps_support_email,
        "GO_LIVE_DATE":           go_live_str,
        "HYPERCARE_END_DATE":     hypercare_end_str,
        "SMARTSHEET_TRACKER_URL": project_row.get("project_link", ""),
    }
    return ctx


# ─────────────────────────────────────────────────────────────────────────────
# 4. Renderer
# ─────────────────────────────────────────────────────────────────────────────

def render_template(
    body_template: str,
    subject_template: str,
    auto_context: dict,
    manual_context: Optional[dict] = None,
) -> tuple[str, str]:
    """
    Render subject and body by substituting all {PLACEHOLDER} tokens.

    auto_context   : dict from build_auto_context()
    manual_context : dict of consultant-entered field values from the UI

    Returns (rendered_subject, rendered_body).
    Unreplaced placeholders are left visible so the consultant can spot them.
    """
    ctx = {**auto_context, **(manual_context or {})}

    def _replace(template: str) -> str:
        result = template
        for key, value in ctx.items():
            result = result.replace(f"{{{key}}}", str(value) if value else f"{{{key}}}")
        return result

    return _replace(subject_template), _replace(body_template)


# ─────────────────────────────────────────────────────────────────────────────
# 5. Send identifier + session log
# ─────────────────────────────────────────────────────────────────────────────

def generate_send_id(project_id: str, template_id: str) -> str:
    """
    Generate a unique identifier for a specific email send.
    Format: {project_id}__{template_id}__{uuid4_short}

    This ID is:
      - Embedded in the email as a hidden header (X-PSPT-Send-ID) for Gmail API
      - Written to the session log
      - Used as the reference when writing back to Smartsheet
    """
    short_uuid = str(uuid.uuid4()).replace("-", "")[:12]
    return f"{project_id}__{template_id}__{short_uuid}"


def _init_session_log():
    """Ensure the session log list exists in session state."""
    if _SESSION_LOG_KEY not in st.session_state:
        st.session_state[_SESSION_LOG_KEY] = []


def record_send(
    send_id: str,
    project_id: str,
    template_id: str,
    template_name: str,
    recipient_email: str,
    subject: str,
    ss_milestone_field: Optional[str],
    sent_at: Optional[datetime.datetime] = None,
) -> dict:
    """
    Record a sent email to the session log.

    Returns the log entry dict (can be used for display or further processing).
    This is called after a successful Gmail API send (or simulated send in pre-Gmail phase).
    """
    _init_session_log()

    entry = {
        "send_id":            send_id,
        "project_id":         project_id,
        "template_id":        template_id,
        "template_name":      template_name,
        "recipient_email":    recipient_email,
        "subject":            subject,
        "ss_milestone_field": ss_milestone_field,
        "sent_at":            (sent_at or datetime.datetime.utcnow()).isoformat(),
        "ss_writeback_done":  False,   # updated after Smartsheet writeback
    }
    st.session_state[_SESSION_LOG_KEY].append(entry)
    return entry


def get_session_send_log() -> list[dict]:
    """Return all email send records for this session."""
    _init_session_log()
    return list(st.session_state[_SESSION_LOG_KEY])


def mark_ss_writeback_done(send_id: str):
    """Mark a send log entry as having its Smartsheet writeback completed."""
    _init_session_log()
    for entry in st.session_state[_SESSION_LOG_KEY]:
        if entry["send_id"] == send_id:
            entry["ss_writeback_done"] = True
            break


# ─────────────────────────────────────────────────────────────────────────────
# 6. Smartsheet writeback payload
# ─────────────────────────────────────────────────────────────────────────────

def build_ss_writeback(
    project_id: str,
    milestone_field: str,
    date_value: Optional[datetime.date] = None,
    current_drs_row: Optional[dict] = None,
) -> dict:
    """
    Build the Smartsheet writeback payload for a milestone date update.

    milestone_field  : Smartsheet column name (e.g. 'Intro. Email Sent')
    date_value       : date to write; defaults to today
    current_drs_row  : the project's current DRS row as a dict (from df_drs).
                       When provided, the field is checked for an existing value
                       before writing — if already populated, writeback is skipped
                       to protect manually-entered or system-set dates.

    Returns a dict with:
      "project_id"  : str
      "fields"      : dict of {field: date_string} — may be empty if skipped
      "skipped"     : list of field names that were not written (already had a value)
      "skip_reason" : human-readable explanation for any skipped fields

    The returned dict is passed to shared/smartsheet_api.py::update_row().
    Only fields in the existing writeback allowlist are permitted.

    No-override policy
    ──────────────────
    Certain milestone fields can be set early for legitimate reasons (e.g. a
    consultant moves to Production at the start for a specific purpose, or the
    customer requests a delayed Hypercare start). In those cases the field will
    already have a value in DRS. This function will NOT overwrite an existing
    value — the consultant should update Smartsheet manually if a correction is
    genuinely needed.

    Fields subject to no-override check:
      Prod Cutover, Hypercare Start — most likely to be pre-populated legitimately.
      All other fields also respect this policy for safety.
    """
    WRITEBACK_ALLOWLIST = {
        SS_INTRO_EMAIL_SENT,       # Welcome email           → "Intro. Email Sent"
        SS_CONFIG_ENABLEMENT,      # Post-Session 1          → "Enablement Session"
        SS_WORKING_SESSION_1,      # Post-Session 2          → "Session #1"
        SS_WORKING_SESSION_2,      # Post-Session 3/Q&A      → "Session #2"
        SS_UAT_SIGNOFF,            # UAT Sign-Off            → "UAT Signoff"
        SS_PROD_CUTOVER,           # Go-Live kickoff         → "Prod Cutover"
        SS_HYPERCARE_START,        # Go-Live kickoff         → "Hypercare Start"
        SS_TRANSITION_SUPPORT,     # Hypercare Closure       → "Transition to Support"
        SS_GO_LIVE_DATE,           # Go-Live date            → "Go-Live Date"
        # Phase & Status handled separately by My Projects page
    }

    if milestone_field not in WRITEBACK_ALLOWLIST:
        raise ValueError(
            f"Field '{milestone_field}' is not in the Smartsheet writeback allowlist. "
            "Add it to WRITEBACK_ALLOWLIST only after confirming with SS admin."
        )

    write_date = date_value or datetime.date.today()
    skipped = []
    skip_reason = ""

    # No-override check: if DRS row is provided and the field already has a value, skip.
    if current_drs_row is not None:
        # Try both exact key and lowercased key (DRS column names can vary in case)
        existing = (
            current_drs_row.get(milestone_field)
            or current_drs_row.get(milestone_field.lower())
            or current_drs_row.get(milestone_field.replace(" ", "_").lower())
        )
        if existing and str(existing).strip() not in ("", "None", "nan", "NaT"):
            skipped.append(milestone_field)
            skip_reason = (
                f"'{milestone_field}' already has a value in Smartsheet "
                f"({existing}). No write performed to avoid overriding an "
                f"existing date. Update Smartsheet directly if a correction is needed."
            )

    fields_to_write = {} if skipped else {milestone_field: write_date.strftime("%Y-%m-%d")}

    return {
        "project_id":  project_id,
        "fields":      fields_to_write,
        "skipped":     skipped,
        "skip_reason": skip_reason,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 7. Project flag: needs welcome email?
# ─────────────────────────────────────────────────────────────────────────────

def project_needs_welcome_email(project_row: dict) -> bool:
    """
    Returns True if this project should be flagged as needing a Welcome email.

    Conditions:
      - Project status is active (not Closed, On Hold, Cancelled)
      - 'Intro. Email Sent' milestone date is empty/null in DRS
      - Project has a recognisable product key (a template exists for it)

    Used by My Projects page to surface the flag.
    """
    status = str(project_row.get("status", "")).lower()
    if any(s in status for s in ("closed", "on hold", "cancelled", "complete")):
        return False

    intro_sent = project_row.get("intro_email_sent") or project_row.get("Intro. Email Sent")
    if intro_sent:
        return False

    product = project_row.get("product", project_row.get("project_type", ""))
    template = get_welcome_template(product)
    return template is not None


# ─────────────────────────────────────────────────────────────────────────────
# 8. Gmail API stub (pending IT approval)
# ─────────────────────────────────────────────────────────────────────────────

class GmailSendResult:
    """Result object returned by send_email()."""
    def __init__(self, success: bool, message_id: Optional[str] = None, error: Optional[str] = None):
        self.success    = success
        self.message_id = message_id
        self.error      = error


def send_email(
    to: str,
    cc: list[str],
    subject: str,
    body: str,
    send_id: str,
) -> GmailSendResult:
    """
    Send an email via the Gmail API (shared PS inbox, delegated OAuth).

    PENDING IT APPROVAL — currently returns a simulated success so the
    session log and Smartsheet writeback can be developed and tested
    independently of the Gmail integration.

    When Gmail API is approved:
      1. Replace the stub below with actual API call via shared/gmail_api.py
      2. Pass send_id as the X-PSPT-Send-ID custom header for traceability
      3. Return the real Gmail message_id

    Args:
        to         : recipient email address
        cc         : list of CC addresses (consultant + any others)
        subject    : rendered email subject
        body       : rendered email body (plain text; HTML wrapper added by gmail_api)
        send_id    : unique identifier from generate_send_id()
    """
    # ── STUB: simulate success ────────────────────────────────────────────────
    # Remove this block and replace with gmail_api.send() when approved.
    simulated_message_id = f"SIMULATED_{send_id}"
    return GmailSendResult(success=True, message_id=simulated_message_id)
    # ─────────────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────────────────
# 9. Convenience: full send flow
# ─────────────────────────────────────────────────────────────────────────────

def execute_send(
    project_id: str,
    template_id: str,
    template_name: str,
    subject: str,
    body: str,
    recipient_email: str,
    cc_emails: list[str],
    ss_milestone_field: Optional[str],
    df_drs_row: Optional[dict] = None,
) -> tuple[bool, str]:
    """
    Full send pipeline:
      1. Generate a unique send ID
      2. Call send_email() (stub / real)
      3. On success: record to session log + build SS writeback payload
      4. Return (success: bool, message: str)

    The Smartsheet writeback itself is intentionally NOT called here —
    it's executed by the page after confirming the send, using the payload
    returned from build_ss_writeback(). This keeps the write separate from
    the send so failures in either are independently recoverable.
    """
    send_id = generate_send_id(project_id, template_id)

    result = send_email(
        to=recipient_email,
        cc=cc_emails,
        subject=subject,
        body=body,
        send_id=send_id,
    )

    if result.success:
        record_send(
            send_id=send_id,
            project_id=project_id,
            template_id=template_id,
            template_name=template_name,
            recipient_email=recipient_email,
            subject=subject,
            ss_milestone_field=ss_milestone_field,
        )
        return True, send_id
    else:
        return False, result.error or "Unknown send error"
