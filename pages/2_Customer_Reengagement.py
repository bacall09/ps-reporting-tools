import streamlit as st
import pandas as pd
from datetime import date, datetime
import re
from rapidfuzz import fuzz
from shared.constants import (
    EMPLOYEE_ROLES, ACTIVE_EMPLOYEES, PRODUCT_KEYWORDS,
    MILESTONE_COLS_MAP, SS_COL_MAP, NS_COL_MAP, SFDC_COL_MAP,
    name_matches,
)
from shared.loaders import (
    load_drs, load_sfdc, load_ns_time,
    calc_days_inactive, calc_last_milestone,
    fuzzy_match_sfdc, normalise_product_name, suggest_tier_from_days,
)
from shared.template_utils import (
    TEMPLATES, suggest_tier, fill_template,
    highlight_placeholders, extract_placeholders,
)




# ── Product keywords for fuzzy matching ──────────────────────────────────────
PRODUCT_KEYWORDS = [
    "Capture", "Approvals", "Reconcile", "PSP", "Payments", "SFTP",
    "E-Invoicing", "eInvoicing", "CC", "Premium", "ZoneCapture",
    "ZoneApprovals", "ZoneReconcile", "ZonePayments", "ZCapture",
    "ZApprovals", "ZReconcile",
]

def _extract_product_hints(text):
    """Return set of product keywords found in a string."""
    t = str(text).lower()
    return {kw for kw in PRODUCT_KEYWORDS if kw.lower() in t}

def _clean_account(text):
    """Normalise account name for fuzzy comparison — strip legal suffixes, punctuation."""
    t = str(text).lower()
    for stop in ["ltd","limited","inc","llc","plc","gmbh","the ","abf ","- za -","& co","co."]:
        t = t.replace(stop, " ")
    return re.sub(r"[^a-z0-9 ]", " ", t).split()

def fuzzy_match_sfdc(df_sfdc, project_name, account_name):
    """
    Find best SFDC row(s) for a DRS project using:
    1. Exact opportunity name match
    2. Fuzzy account name + product keyword overlap
    Returns matching rows and a confidence label.
    """
    if df_sfdc is None or df_sfdc.empty:
        return pd.DataFrame(), None

    # ── Exact opp name match ───────────────────────────────────────────────
    if "opportunity" in df_sfdc.columns:
        exact = df_sfdc[df_sfdc["opportunity"].astype(str).str.strip().str.lower()
                        == str(project_name).strip().lower()]
        if not exact.empty:
            return exact, "Exact match"

    # ── Fuzzy account + product keyword match ─────────────────────────────
    drs_words    = set(_clean_account(account_name))
    drs_products = _extract_product_hints(project_name)

    best_score  = 0
    best_rows   = pd.DataFrame()
    best_label  = None
    best_opp_id = None
    best_opp_nm = None

    for _, row in df_sfdc.iterrows():
        sfdc_account  = str(row.get("account",  ""))
        sfdc_opp      = str(row.get("opportunity", ""))

        # Account fuzzy score — word overlap + token set ratio for partial name matches
        sfdc_words    = set(_clean_account(sfdc_account))
        common        = drs_words & sfdc_words
        word_score    = len(common) / max(len(drs_words), 1) * 100
        fuzz_score    = fuzz.token_set_ratio(" ".join(drs_words), " ".join(sfdc_words))
        acct_score    = max(word_score, fuzz_score * 0.7)  # blend both signals

        # Product keyword overlap
        sfdc_products = _extract_product_hints(sfdc_opp)
        prod_match    = bool(drs_products & sfdc_products)

        # Combined score — account similarity weighted, product match bonus
        score = acct_score + (30 if prod_match else 0)

        if score > best_score:
            best_score   = score
            best_opp_id  = row.get("opportunity_id", None)
            best_opp_nm  = row.get("opportunity",    None)
            best_label   = (
                f"Fuzzy match ({int(acct_score)}% account · {'✅ product match' if prod_match else '⚠️ no product match'})"
            )

    # Only return if confident enough
    if best_score >= 60:
        # Return ALL contacts for the matched opportunity (not just the best-scoring row)
        if best_opp_id and "opportunity_id" in df_sfdc.columns:
            best_rows = df_sfdc[df_sfdc["opportunity_id"] == best_opp_id]
        elif best_opp_nm and "opportunity" in df_sfdc.columns:
            best_rows = df_sfdc[df_sfdc["opportunity"] == best_opp_nm]
        if not best_rows.empty:
            return best_rows, best_label

    # ── Last resort: account name only, high threshold ────────────────────
    if "account" in df_sfdc.columns and account_name:
        df_sfdc["_acct_score"] = df_sfdc["account"].apply(
            lambda x: fuzz.token_set_ratio(str(account_name).lower(), str(x).lower())
        )
        top = df_sfdc[df_sfdc["_acct_score"] >= 75].sort_values("_acct_score", ascending=False)
        df_sfdc.drop(columns=["_acct_score"], inplace=True)
        if not top.empty:
            return top, f"Account name match ({top.iloc[0].get('_acct_score', '?')}% similarity)"

    return pd.DataFrame(), None


# ── Active employee list (sourced from EMPLOYEE_ROLES — leavers excluded) ────
ACTIVE_EMPLOYEES = [
    "Arestarkhov, Yaroslav", "Barrio, Nairobi", "Bell, Stuart", "Cadelina",
    "Carpen, Anamaria", "Centinaje, Rhodechild", "Church, Jason G", "Cooke, Ellen",
    "Cruz, Daniel", "DiMarco, Nicole R", "Dolha, Madalina", "Dunn, Steven",
    "Finalle-Newton, Jesse", "Gardner, Cheryll L", "Hopkins, Chris", "Hughes, Madalyn",
    "Ickler, Georganne", "Isberg, Eric", "Jordanova, Marija", "Lappin, Thomas",
    "Law, Brandon", "Longalong, Santiago", "Mohammad, Manaan", "Morris, Lisa",
    "Murphy, Conor", "NAQVI, SYED", "Olson, Austin D", "Pallone, Daniel",
    "Porangada, Suraj", "Quiambao, Generalyn", "Raykova, Silvia",
    "Selvakumar, Sajithan", "Snee, Stefanie J", "Swanson, Patti",
    "Tuazon, Carol", "Zoric, Ivan",
]

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        h1, h2, h3, .stMarkdown, .stDataFrame, label, button { font-family: 'Manrope', sans-serif !important; }
        .email-box {
            background: var(--background-color, white);
            border: 1px solid #d0dff5;
            border-radius: 8px;
            padding: 20px 24px;
            font-family: 'Manrope', sans-serif;
            font-size: 14px;
            line-height: 1.7;
            white-space: pre-wrap;
            color: var(--text-color, #1e2c63);
        }
        .placeholder-missing {
            background: #fff3cd;
            border-radius: 3px;
            padding: 1px 4px;
            font-weight: 600;
            color: #856404;
        }
        .placeholder-filled {
            background: #d4edda;
            border-radius: 3px;
            padding: 1px 4px;
            font-weight: 600;
            color: #155724;
        }
        .tier-badge-1 { background:#EAF9F1; color:#1E8449; border-radius:6px; padding:3px 10px; font-weight:700; font-size:13px; }
        .tier-badge-2 { background:#FEF9E7; color:#9C6500; border-radius:6px; padding:3px 10px; font-weight:700; font-size:13px; }
        .tier-badge-3 { background:#FDECED; color:#C0392B; border-radius:6px; padding:3px 10px; font-weight:700; font-size:13px; }
        .tier-badge-4 { background:#f0e6ff; color:#6c3483; border-radius:6px; padding:3px 10px; font-weight:700; font-size:13px; }
        .cc-box {
            background: #f0f4ff;
            border-left: 3px solid #4472C4;
            padding: 10px 14px;
            border-radius: 4px;
            font-size: 13px;
            font-family: 'Manrope', sans-serif;
            color: #1e2c63;
            margin-bottom: 8px;
        }
    </style>
""", unsafe_allow_html=True)

# ── Hero ──────────────────────────────────────────────────────────────────────
def _title_suffix_from_browse():
    b = st.session_state.get("home_browse", "— My own view —") or ""
    if b.startswith("── ") and b.endswith(" ──"):
        return f" — {b[3:-3].strip()} Team"
    if b and b not in ("— My own view —", "— Select —", "👥 All team"):
        parts = [p.strip() for p in b.split(",")]
        return f" — {parts[1] + ' ' + parts[0] if len(parts)==2 else b}"
    return ""

st.markdown(f"""
    <div style='background:#1B2B5E;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;font-family:Manrope,sans-serif;position:relative;overflow:hidden'>
        <div style='position:absolute;right:-40px;top:-40px;width:220px;height:220px;border-radius:50%;background:radial-gradient(circle,rgba(91,141,239,0.15) 0%,transparent 70%);pointer-events:none'></div>
        <div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#ff4b40;margin-bottom:10px;font-family:Manrope,sans-serif'>Professional Services · Tools</div>
        <h1 style='color:#fff;margin:0;font-size:28px;font-weight:800;font-family:Manrope,sans-serif;line-height:1.15'>Customer Engagement{_title_suffix_from_browse()}</h1>
        <p style='color:rgba(255,255,255,0.6);margin:8px 0 0;font-size:14px;font-family:Manrope,sans-serif;line-height:1.6;max-width:520px'>Tier-based re-engagement communications for on-hold or stalled projects — auto-suggests outreach level based on days inactive.</p>
    </div>
""", unsafe_allow_html=True)

# ── Phase Banner ─────────────────────────────────────────────────────────────
st.markdown("""
<div style='background:#F0F4FF;border-left:4px solid #1E2C63;border-radius:6px;
            padding:16px 20px;margin:12px 0 20px;font-family:Manrope,sans-serif'>
    <div style='font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;
                color:#1E2C63;margin-bottom:10px'>Roadmap</div>
    <div style='display:flex;gap:32px;flex-wrap:wrap'>
        <div style='flex:1;min-width:240px'>
            <span style='background:#1E2C63;color:#fff;font-size:10px;font-weight:700;
                         padding:2px 8px;border-radius:10px;letter-spacing:1px'>PHASE 1 · NOW</span>
            <p style='margin:8px 0 0;font-size:13px;color:#333;line-height:1.6'>
                Access project contact details for all DRS-assigned projects.
                Templates for <strong>Customer Re-Engagement</strong> communications are pre-loaded
                and auto-suggested based on days inactive.
            </p>
        </div>
        <div style='flex:1;min-width:240px'>
            <span style='background:#808080;color:#fff;font-size:10px;font-weight:700;
                         padding:2px 8px;border-radius:10px;letter-spacing:1px'>PHASE 2 · COMING SOON</span>
            <p style='margin:8px 0 0;font-size:13px;color:#555;line-height:1.6'>
                Templates for <strong>Project Lifecycle</strong> communications will be added
                once finalised — covering kick-off, go-live, and project close communications.
            </p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Templates ─────────────────────────────────────────────────────────────────
TEMPLATES = {
    "Tier 1 — ~30 Day Check-In": {
        "tier": 1,
        "days_min": 30,
        "days_max": 59,
        "subject": "{PRODUCT NAME} — Checking In on Your Project",
        "cc_guidance": "No additional CC required.",
        "body": """Hi {CUSTOMER CONTACT NAME},

I hope you're doing well. I wanted to reach out as it's been a little while since we last connected on your {PRODUCT NAME} implementation.

When we last spoke, we were in the {CURRENT PHASE} phase. I appreciate that priorities can shift and schedules get busy — I just wanted to check in and see if you're ready to pick things back up, or if anything has changed on your end that I should be aware of.

We're here whenever you're ready. If it's helpful, I'm happy to schedule a short call to realign on next steps and get things moving again.

Looking forward to hearing from you.

Kind regards,
{IMPLEMENTATION CONSULTANT}
Professional Services | Zone & Co"""
    },
    "Tier 2 — ~60 Day Follow-Up": {
        "tier": 2,
        "days_min": 60,
        "days_max": 89,
        "subject": "{PRODUCT NAME} Project — Let's Reconnect",
        "cc_guidance": "CC: PS Leadership",
        "body": """Hi {CUSTOMER CONTACT NAME},

I'm following up on my earlier message regarding your {PRODUCT NAME} implementation. Our last activity together was on {LAST ACTIVITY DATE}, during the {CURRENT PHASE} phase, and I want to make sure we can keep your project on track.

Where Things Stand

As outlined in the implementation scope shared at the start of the project, estimated timelines are based on mutual availability and timely engagement from both teams. Extended pauses can affect a few things I want to flag for your awareness:

• Consultant availability: Prolonged delays may mean that your currently assigned consultant is reassigned to other projects. We'll do our best to maintain continuity, but early re-engagement helps.
• Session utilization: Your implementation includes a defined number of working sessions and enablement activities. These are available through to project closure.
• Environment changes: If your Sandbox environment is refreshed during the pause, reconfiguration may be required, which could impact the remaining scope.

Suggested Next Step

I'd recommend we schedule a 30-minute call to review where we left off, confirm your current priorities, and agree on an updated timeline. Could you let me know your availability over the next week or two?

We're committed to helping you get the most out of {PRODUCT NAME} and want to ensure a smooth path to Go-Live.

Kind regards,
{IMPLEMENTATION CONSULTANT}
Professional Services | Zone & Co"""
    },
    "Tier 3 — ~90 Day Escalation": {
        "tier": 3,
        "days_min": 90,
        "days_max": 179,
        "subject": "{PRODUCT NAME} Implementation — Re-Engagement Needed",
        "cc_guidance": "CC: PS Leadership · Account Manager ({ACCOUNT MANAGER}) · CS Manager",
        "body": """Hi {CUSTOMER CONTACT NAME},

I'm reaching out once more regarding your {PRODUCT NAME} implementation. It's now been over 90 days since our last project activity on {LAST ACTIVITY DATE}, and I want to make sure we can find a way forward together.

Project Status

Your project is currently in the {CURRENT PHASE} phase. We've been holding this open on our side, but I want to be transparent about a few things that extended inactivity can affect:

• Consultant reassignment: After this length of time, your assigned consultant may no longer be available when you're ready to resume. We'll work to ensure a smooth handover if that's the case, but there may be some ramp-up time.
• Environment drift: If your NetSuite Sandbox has been refreshed or modified during this period, reconfiguration of {PRODUCT NAME} may be required. Depending on the extent, this could impact the remaining scope and deliverables.
• Knowledge continuity: The longer the gap, the more likely it is that key decisions or configurations from earlier phases need to be revisited, which can extend the overall timeline.

Recommended Next Steps

I've copied your Customer Success Manager, {ACCOUNT MANAGER}, on this email so we can coordinate a plan together. Specifically, it would be helpful to:

• Schedule a re-alignment call to review the current project status and outstanding deliverables.
• Confirm availability of your project team and NetSuite administrator.
• Agree on an updated timeline to get the project back on track.

We understand that circumstances change, and we're here to help find the right path forward. Please let us know how you'd like to proceed.

Kind regards,
{IMPLEMENTATION CONSULTANT}
Professional Services | Zone & Co"""
    },
    "Tier 4 — ~6 Month Notification": {
        "tier": 4,
        "days_min": 180,
        "days_max": 99999,
        "subject": "Important: {PRODUCT NAME} Implementation — Service Term Update",
        "cc_guidance": "CC: PS Leadership · Account Manager ({ACCOUNT MANAGER}) · CS Manager",
        "body": """Hi {CUSTOMER CONTACT NAME},

I'm writing to provide an important update regarding your {PRODUCT NAME} implementation. Despite our previous outreach, we've been unable to reconnect with your team, and I want to ensure you have full visibility into where things stand so we can plan the best path forward.

Project Status

Your project has been in the {CURRENT PHASE} phase since our last activity on {LAST ACTIVITY DATE}. You have {REMAINING SESSIONS} sessions remaining in scope.

Service Term & Scope Reminders

I'd like to bring a few items from our agreed implementation scope to your attention:

• Service term: Your Professional Services engagement is valid for 12 months from your contract signature date. Your current service term expires on {SERVICE TERM EXPIRY}.
• Session utilization: Any unused sessions or deliverables remaining at the time of transition to Support will be considered forfeited and no longer valid.
• Resource availability: Extended delays may impact the continued availability of your assigned consultant. It is the customer's responsibility to notify Zone when ready to re-engage.
• Environment changes: If a Sandbox refresh has occurred during this period, reconfiguration may be required and could impact remaining scope.

Recommended Next Steps

To make the most of your remaining sessions and service term, I'd strongly recommend we put together a re-engagement plan as soon as possible. I've copied your Account Manager, {ACCOUNT MANAGER}, on this email so we can coordinate together.

• Schedule a re-alignment call to review the current project status and outstanding deliverables.
• Confirm availability of your project team and NetSuite administrator.
• Agree on an updated timeline that allows us to complete the implementation within your service term.

We want to ensure you get full value from your investment in {PRODUCT NAME}. Please let us know how you'd like to proceed.

Kind regards,
{IMPLEMENTATION CONSULTANT}
Professional Services | Zone & Co"""
    },
}

TIER_COLORS = {1: "#EAF9F1", 2: "#FEF9E7", 3: "#FDECED", 4: "#f0e6ff"}
TIER_TEXT   = {1: "#1E8449", 2: "#9C6500", 3: "#C0392B", 4: "#6c3483"}

# ── SFDC column map ───────────────────────────────────────────────────────────
SFDC_COL_MAP = {
    # Exact headers from SFDC contacts export
    "18 digit opportunity id":  "opportunity_id",
    "first name":               "first_name",
    "last name":                "last_name",
    "primary title":            "title",
    "email":                    "email",
    "opportunity owner":        "account_manager",
    "account name":             "account",
    "opportunity name":         "opportunity",
    "close date":               "close_date",
    # Fallback aliases
    "opportunity":              "opportunity",
    "account":                  "account",
    "contact name":             "contact_name",
    "full name":                "contact_name",
    "name":                     "contact_name",
    "contact":                  "contact_name",
    "contact email":            "email",
    "email address":            "email",
    "role":                     "title",
    "product":                  "product",
    "products":                 "product",
    "project type":             "product",
    "product family":           "product",
    "closed date":              "close_date",
    "stage":                    "stage",
    "territory":                "territory",
    "account owner":            "account_manager",
    "owner":                    "account_manager",
    "owner name":               "account_manager",
    "opportunity owner name":   "account_manager",
    "rep":                      "account_manager",
    "ae":                       "account_manager",
    "csm":                      "account_manager",
    # New columns
    "implementation contact":   "impl_contact_flag",
    "implementation contact exists": "impl_contact_flag",
    "contact roles":            "contact_roles",
    "opportunity owner email":  "account_manager_email",
    "owner email":              "account_manager_email",
    "opp contact role count":   "role_count",
    "partner contact":          "partner_contact",
    "primary":                  "is_primary",
    "primary contact":          "is_primary",
    "is primary":               "is_primary",
    "title":                    "title",
}

def load_sfdc(file):
    if file.name.endswith(".csv"):
        for enc in ["utf-8", "utf-8-sig", "windows-1252", "latin-1", "cp1252"]:
            try:
                file.seek(0)
                df = pd.read_csv(file, encoding=enc)
                break
            except (UnicodeDecodeError, LookupError):
                continue
        else:
            file.seek(0)
            df = pd.read_csv(file, encoding="latin-1")  # last resort
    else:
        df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    rename = {c: SFDC_COL_MAP[c.lower()] for c in df.columns if c.lower() in SFDC_COL_MAP}
    df = df.rename(columns=rename)
    if "close_date" in df.columns:
        df["close_date"] = pd.to_datetime(df["close_date"], errors="coerce")
    # Always build contact_name from first + last (source report has separate columns)
    if "first_name" in df.columns and "last_name" in df.columns:
        df["contact_name"] = (df["first_name"].fillna("") + " " + df["last_name"].fillna("")).str.strip()
    # Use opportunity name as the project label
    if "opportunity" not in df.columns and "opportunity_id" in df.columns:
        df["opportunity"] = df["opportunity_id"]
    # Normalise impl_contact_flag — checkbox may come as TRUE/FALSE, Yes/No, 1/0
    if "impl_contact_flag" in df.columns:
        df["impl_contact_flag"] = df["impl_contact_flag"].astype(str).str.strip().str.lower().isin(
            ["true", "yes", "1", "checked", "x"]
        )
    # Normalise is_primary flag — store as readable string to avoid date rendering
    if "is_primary" in df.columns:
        _prim_bool = df["is_primary"].astype(str).str.strip().str.lower().isin(
            ["true", "yes", "1", "checked", "x"]
        )
        df["is_primary"] = _prim_bool.map({True: "✓", False: ""})
    # Normalise contact_roles to lowercase string for matching
    if "contact_roles" in df.columns:
        df["contact_roles"] = df["contact_roles"].fillna("").astype(str).str.strip()
    return df

def suggest_tier(days_inactive):
    if days_inactive is None:
        return None
    for name, tmpl in TEMPLATES.items():
        if tmpl["days_min"] <= days_inactive <= tmpl["days_max"]:
            return name
    return list(TEMPLATES.keys())[-1]

def fill_template(template_key, fields):
    tmpl  = TEMPLATES[template_key]
    body  = tmpl["body"]
    subj  = tmpl["subject"]
    cc    = tmpl["cc_guidance"]

    for k, v in fields.items():
        if v:
            body = body.replace(f"{{{k}}}", str(v))
            subj = subj.replace(f"{{{k}}}", str(v))
            cc   = cc.replace(f"{{{k}}}", str(v))
    return subj, body, cc

def highlight_placeholders(text):
    """Return text with remaining {PLACEHOLDERS} wrapped in HTML spans."""
    def replacer(m):
        return f"<span class='placeholder-missing'>{m.group(0)}</span>"
    return re.sub(r'\{[A-Z _]+\}', replacer, text)

def extract_placeholders(text):
    return re.findall(r'\{([A-Z _]+)\}', text)

# ── Main ──────────────────────────────────────────────────────────────────────
# ── SS DRS column map ────────────────────────────────────────────────────────
SS_COL_MAP_OUT = {
    "project name":           "project_name",
    "name":                   "project_name",   # fallback: exports where header is just "Name"
    "project id":             "project_id",
    "project phase":          "phase",
    "project type":           "project_type",
    "status":                 "status",
    "start date":             "start_date",
    "start date (subscription)": "subscription_start_date",  # kept separate — not project start
    "go live date":           "go_live_date",
    "territory":              "territory",
    "billing type":           "billing_type",
    "billing":                "billing_type",
    "project manager":        "project_manager",
    "consultant":             "project_manager",
    "overall rag":            "rag",
    "schedule health":        "schedule_health",
    "risk level":             "risk_level",
    "client responsiveness":  "client_responsiveness",
    "last updated":           "last_updated",
    "modified":               "last_updated",
    "modified date":          "last_updated",
    "last modified":          "last_updated",
    "date modified":          "last_updated",
    "account name":           "account",
    "customer":               "account",
    "risk detail":            "risk_detail",
    "risk owner":             "risk_owner",
    "responsible for delay":  "responsible_delay",
    "delay summary":          "delay_summary",
    "intro. email sent":           "ms_intro_email",
    "standard config start":       "ms_config_start",
    "enablement session":          "ms_enablement",
    "session #1":                  "ms_session1",
    "session #2":                  "ms_session2",
    "uat signoff":                 "ms_uat_signoff",
    "prod cutover":                "ms_prod_cutover",
    "hypercare start":             "ms_hypercare_start",
    "close out remaining tasks":   "ms_close_out",
    "transition to support":       "ms_transition",
}

INACTIVE_PHASES_OUT = {
    "10. complete/pending final billing",
    "11. on hold",
    "12. ps review",
}

def load_drs(file):
    """Load SS DRS export and identify stale/unresponsive projects."""
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    rename = {col: SS_COL_MAP_OUT[col.lower()] for col in df.columns if col.lower() in SS_COL_MAP_OUT}
    df = df.rename(columns=rename)

    if "start_date" in df.columns:
        df["start_date"] = pd.to_datetime(df["start_date"], errors="coerce")
    if "go_live_date" in df.columns:
        df["go_live_date"] = pd.to_datetime(df["go_live_date"], errors="coerce")
    if "last_updated" in df.columns:
        df["last_updated"] = pd.to_datetime(df["last_updated"], errors="coerce")
    # Parse all milestone date columns
    for _ms_col in MILESTONE_COLS_MAP.keys():
        if _ms_col in df.columns:
            df[_ms_col] = pd.to_datetime(df[_ms_col], errors="coerce")

    # Filter to active FF projects only
    today = pd.Timestamp.today().normalize()
    if "phase" in df.columns:
        df = df[~df["phase"].str.strip().str.lower().isin(INACTIVE_PHASES_OUT)]
    if "billing_type" in df.columns:
        df = df[~df["billing_type"].str.strip().str.lower().isin({"t&m","time & material","time and material"})]
    # Calculate remaining sessions from milestone columns
    _session_ms_cols = ["ms_enablement", "ms_session1", "ms_session2"]
    _present_ms = [col for col in _session_ms_cols if col in df.columns]
    if _present_ms:
        def _count_remaining(row):
            return sum(1 for col in _present_ms if pd.isna(row[col]) or str(row[col]).strip() in ("", "nan", "None", "NaT"))
        df["remaining_sessions"] = df.apply(_count_remaining, axis=1).astype(int)
    else:
        df["remaining_sessions"] = None

    # Calculate last milestone (name + date of most recently completed milestone)
    df["last_milestone"] = df.apply(calc_last_milestone, axis=1)

    # Tag On Hold but keep in df — excluded from dropdown, shown in table
    ON_HOLD_VALS = {"on-hold","on hold","onhold","on_hold"}
    if "status" in df.columns:
        df["_on_hold"] = df["status"].str.strip().str.lower().isin(ON_HOLD_VALS)
    else:
        df["_on_hold"] = False

    # ── Inactivity signal hierarchy (DRS-only baseline, NS overrides in calc_days_inactive) ──
    # 1. Last Milestone date — days since most recently completed milestone
    # 2. Phase vs Start Date benchmark — days open minus expected phase duration
    # 3. Unknown (-1) — NS Time Entry will override when uploaded
    _PHASE_BENCH = {
        "00. onboarding": 7, "01. requirements and design": 14,
        "02. configuration": 21, "03. enablement/training": 14,
        "04. uat": 28, "05. prep for go-live": 7,
        "06. go-live (hypercare)": 14, "08. ready for support transition": 5,
        "09. phase 2 scoping": 14,
    }

    def _calc_inactivity_drs(row):
        # 1. Last milestone
        for _ms_col in reversed(list(MILESTONE_COLS_MAP.keys())):
            if _ms_col in row.index and pd.notna(row[_ms_col]):
                try:
                    return int((today - pd.Timestamp(row[_ms_col])).days), "Last Milestone"
                except: pass
        # 2. Phase vs start date
        if "start_date" in row.index and pd.notna(row["start_date"]):
            try:
                days_open    = int((today - pd.Timestamp(row["start_date"])).days)
                bench        = _PHASE_BENCH.get(str(row.get("phase","")).strip().lower(), 14)
                days_stalled = max(0, days_open - bench)
                return days_stalled, "Phase vs Start Date"
            except: pass
        return -1, "Unknown"

    _results = df.apply(_calc_inactivity_drs, axis=1)
    df["days_inactive"]      = _results.apply(lambda x: x[0]).astype(int)
    df["_inactivity_source"] = _results.apply(lambda x: x[1])

    # Store unmapped columns for debug
    df.attrs["unmapped_cols"] = [c for c in df.columns if c not in SS_COL_MAP_OUT.values()]

    return df


def normalise_product_name(raw):
    """Convert internal DRS project type labels to customer-facing product names."""
    if not raw or str(raw).strip().lower() in ("nan", "none", ""):
        return ""
    s = str(raw).strip()

    # Explicit mappings — checked first (case-insensitive key match)
    EXPLICIT = {
        "zoneapp: cc":           "CC Import",
        "zoneapp: psp":          "PSP",
        "zoneapp: sftp":         "SFTP",
        "zoneapp: e-invoicing":  "e-Invoicing",
        "zoneapp: einvoicing":   "e-Invoicing",
        "zoneapp: approvals":    "ZoneApprovals",
        "zoneapp: capture":      "ZoneCapture",
        "zoneapp: reconcile":    "ZoneReconcile",
        "zoneapp: payments":     "ZonePayments",
        "zoneapp: reconcile 2.0":"ZoneReconcile",
        "zoneapp: premium":      "Premium",
        "zoneapp: capture and e-invoicing": "ZoneCapture + e-Invoicing",
    }
    if s.lower() in EXPLICIT:
        return EXPLICIT[s.lower()]

    # Generic fallback: strip "ZoneApp: " prefix and add Zone
    import re as _re
    s = _re.sub(r"(?i)^zoneapp:\s*", "Zone", s)
    return s


# Milestone columns in delivery order — mapped internal name : display name
MILESTONE_COLS_MAP = {
    "ms_intro_email":    "Intro. Email Sent",
    "ms_config_start":   "Standard Config Start",
    "ms_enablement":     "Enablement Session",
    "ms_session1":       "Session #1",
    "ms_session2":       "Session #2",
    "ms_uat_signoff":    "UAT Signoff",
    "ms_prod_cutover":   "Prod Cutover",
    "ms_hypercare_start":"Hypercare Start",
    "ms_close_out":      "Close Out Remaining Tasks",
    "ms_transition":     "Transition to Support",
}

def calc_last_milestone(row):
    """Return 'Milestone Name · YYYY-MM-DD' for the latest completed milestone."""
    best_date = None
    best_name = None
    for col, label in MILESTONE_COLS_MAP.items():
        if col in row.index:
            val = row[col]
            try:
                dt = pd.to_datetime(val, errors="coerce")
                if pd.notna(dt):
                    if best_date is None or dt > best_date:
                        best_date = dt
                        best_name = label
            except: pass
    if best_name and best_date:
        return f"{best_name} · {best_date.strftime('%Y-%m-%d')}"
    return "—"


def suggest_tier_from_days(days):
    try:
        d = int(days)
        if d < 0:   return None   # unknown inactivity
        if d < 30:  return None   # follow-up flag only, no template
        for name, tmpl in TEMPLATES.items():
            if tmpl["days_min"] <= d <= tmpl["days_max"]:
                return name
    except:
        pass
    return None


NS_COL_MAP_OUT = {
    "employee":             "employee",
    "name":                 "employee",
    "project":              "project",
    "project name":         "project",
    "project id":           "project_id",
    "project id ":          "project_id",
    " project id":          "project_id",
    "project_id":           "project_id",
    "projectid":            "project_id",
    "project internal id":  "project_id",
    "job":                  "project_id",
    "billing type":         "billing_type",
    "project manager":      "project_manager",
    "date":                 "date",
    "transaction date":     "date",
    "hours":                "hours",
    "quantity":             "hours",
    "hours/quantity":       "hours",
}

def load_ns_time(file):
    """Load NS Time Detail — derive last time entry per project for inactivity."""
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    rename = {col: NS_COL_MAP_OUT[col.strip().lower()] for col in df.columns if col.strip().lower() in NS_COL_MAP_OUT}
    df = df.rename(columns=rename)
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    if "project_id" in df.columns:
        df["project_id"] = df["project_id"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    # Store original cols for debug
    df.attrs["original_cols"] = list(df.columns)
    return df


def calc_days_inactive(df_drs, df_ns):
    """Merge NS last time entry onto DRS projects to get objective inactivity days."""
    today = pd.Timestamp.today().normalize()
    if df_ns is None or "date" not in df_ns.columns:
        return df_drs

    def _clean_id(series):
        """Safely convert project_id series to clean integer string."""
        def _to_id(v):
            try:
                s = str(v).strip()
                if s in ("", "nan", "None", "NaN"): return ""
                # Handle float like 367986.0 → 367986
                return str(int(float(s)))
            except: return str(v).strip()
        return series.apply(_to_id)

    # Try join on project_id first, fall back to project name
    if "project_id" in df_ns.columns and "project_id" in df_drs.columns:
        df_ns  = df_ns.copy()
        df_drs = df_drs.copy()
        df_ns["project_id"]  = _clean_id(df_ns["project_id"])
        df_drs["project_id"] = _clean_id(df_drs["project_id"])
        last_entry = (
            df_ns[df_ns["date"].notna()]
            .groupby("project_id")["date"].max()
            .reset_index()
            .rename(columns={"date": "last_ns_entry"})
        )
        df_drs = df_drs.merge(last_entry, on="project_id", how="left")
    elif "project" in df_ns.columns and "project_name" in df_drs.columns:
        last_entry = (
            df_ns[df_ns["date"].notna()]
            .groupby("project")["date"].max()
            .reset_index()
            .rename(columns={"project": "project_name", "date": "last_ns_entry"})
        )
        df_drs = df_drs.merge(last_entry, on="project_name", how="left")

    if "last_ns_entry" in df_drs.columns:
        ns_days = (today - df_drs["last_ns_entry"]).dt.days.clip(lower=0)
        fallback = df_drs["days_inactive"] if "days_inactive" in df_drs.columns else pd.Series(-1, index=df_drs.index)
        df_drs["days_inactive"] = ns_days.where(df_drs["last_ns_entry"].notna(), fallback).fillna(-1).astype(int)
        _existing_src = df_drs["_inactivity_source"].tolist() if "_inactivity_source" in df_drs.columns else ["Unknown"] * len(df_drs)
        df_drs["_inactivity_source"] = [
            "NS Time Entry" if pd.notna(x) else _existing_src[i]
            for i, x in enumerate(df_drs["last_ns_entry"])
        ]
    return df_drs


# ── Outreach Log — session_state + /tmp file for cross-session persistence ────
import os, json as _json
LOG_PATH = "/tmp/ps_outreach_log.json"
_LOG_KEY  = "_outreach_log"

def _load_log():
    if _LOG_KEY in st.session_state:
        return list(st.session_state[_LOG_KEY])
    try:
        if os.path.exists(LOG_PATH):
            with open(LOG_PATH) as f:
                entries = _json.load(f)
                st.session_state[_LOG_KEY] = entries
                return entries
    except: pass
    return []

def _save_log(entries):
    entries = entries[-500:]
    st.session_state[_LOG_KEY] = entries
    try:
        with open(LOG_PATH, "w") as f:
            _json.dump(entries, f)
    except: pass

def _log_outreach(consultant, customer, project, tier_label, days_inactive, template):
    entries = _load_log()
    today_str = datetime.today().strftime("%Y-%m-%d")
    # Dedup — don't log the same project+consultant+tier on the same day more than once
    _already = any(
        e.get("project") == str(project) and
        e.get("consultant") == str(consultant) and
        e.get("date") == today_str and
        e.get("tier") == str(tier_label)
        for e in entries
    )
    if _already:
        return
    entries.append({
        "date":          datetime.today().strftime("%Y-%m-%d"),
        "consultant":    str(consultant),
        "customer":      str(customer),
        "project":       str(project),
        "tier":          str(tier_label),
        "days_inactive": int(days_inactive) if days_inactive else 0,
        "template":      str(template),
    })
    _save_log(entries)

def _get_project_log(project):
    return [e for e in _load_log() if e.get("project") == str(project)]


def main():

    # Safe defaults
    _va_name = None; _va_region = None; _is_group = False

    # ── Identity — use Home session state if available ────────────────────
    _session_name = st.session_state.get("consultant_name")

    if _session_name and _session_name != "— Select —":
        from shared.constants import get_role as _get_role, resolve_view_as, get_region_consultants
        from shared.config import EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE
        _home_browse = st.session_state.get("home_browse", "— My own view —")
        _va_name, _va_region, _is_group = resolve_view_as(
            _session_name, _home_browse, EMPLOYEE_ROLES,
            EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE, ACTIVE_EMPLOYEES
        )
        _role = _get_role(_session_name)
        is_manager    = _role in ("manager", "manager_only")
        selected_user = _va_name if _va_name else _session_name
        _disp = (selected_user.split(",")[1].strip() + " " + selected_user.split(",")[0].strip()
                 if "," in selected_user else selected_user)
        if _va_region:
            _disp = f"{_va_region} Team"
        _role_label = "Manager" if is_manager else "Consultant"
    else:
        # Not identified on Home — show Step 1 as before
        st.subheader("Step 2 — Who are you?")
        _u_col, _r_col = st.columns([3, 2])
        with _u_col:
            selected_user = st.selectbox(
                "Select your name",
                ["— Select —"] + ACTIVE_EMPLOYEES,
                key="selected_user"
            )
        with _r_col:
            user_role = st.radio(
                "Role",
                ["IC — my projects only", "Manager / Admin — all projects"],
                key="user_role",
                horizontal=False
            )
        is_manager = user_role.startswith("Manager")
        if selected_user == "— Select —":
            st.info("Select your name to get started — or go to the Home page to set your identity once for all pages.")
            return

    st.markdown("---")
    # ── Pull from session state (uploaded on Home) — local upload as fallback ──
    _from_session = {
        "df_sfdc": st.session_state.get("df_sfdc"),
        "df_drs":  st.session_state.get("df_drs"),
        "df_ns":   st.session_state.get("df_ns"),
    }
    _session_loaded = [k for k, v in _from_session.items() if v is not None]

    # Only show upload section if data not already loaded from Home
    sfdc_file = None
    drs_file  = None
    ns_file   = None

    if not _session_loaded:
        st.subheader("Step 3 — Upload Reports")
        st.caption("Upload one or more reports — more reports = better inactivity detection.")
        up1, up2, up3 = st.columns([3, 3, 3])
        with up1:
            st.markdown("**SFDC Contacts Export** — for contact lookup")
            st.markdown(
                "📁 [Download latest from shared Drive ↗]"
                "(https://drive.google.com/drive/u/1/folders/1VdI_WjuVclF5xN9fG7dEIz1WDu4QRE0m)"
            )
            st.caption("Required: First Name, Last Name, Email, Account Name, Opportunity Name")
            sfdc_file = st.file_uploader("Drop SFDC Contacts file here", type=["xlsx","xls","csv"], key="sfdc_outreach")
            if sfdc_file:
                import re as _re
                _dm = _re.search(r'(\d{8})', sfdc_file.name)
                if _dm:
                    try:
                        _fd   = datetime.strptime(_dm.group(1), "%Y%m%d")
                        _age  = (datetime.today() - _fd).days
                        if _age == 0:       st.success(f"📅 File dated today — current.")
                        elif _age <= 7:     st.info(f"📅 {_fd.strftime('%b %d, %Y')} — {_age}d old.")
                        else:               st.warning(f"📅 {_fd.strftime('%b %d, %Y')} — {_age}d old. Consider refreshing.")
                    except: pass
        with up2:
            st.markdown("**SS DRS Export** — project list & phase")
            st.markdown("[Download current SS DRS report ↗](#)", help="Link coming soon")
            st.caption("Required: Project Name, Project Phase, Project Type, Billing Type, Status")
            drs_file = st.file_uploader("Drop SS DRS file here", type=["xlsx","xls","csv"], key="drs_outreach")
        with up3:
            st.markdown("**NS Time Detail** — objective inactivity signal")
            st.markdown("[Download latest NS Time Detail ↗](https://3838224.app.netsuite.com/app/common/search/searchresults.nl?searchid=70652&saverun=T&whence=)")
            ns_file = st.file_uploader("Drop NS Time Detail file here", type=["xlsx","xls","csv"], key="ns_outreach")
    else:
        # Data loaded from Home page — no banner or override needed
        sfdc_file = None
        drs_file  = None
        ns_file   = None

    # Require at least one source — session state counts
    if not sfdc_file and not drs_file and not ns_file and not _session_loaded:
        st.info("Upload at least one report — or go to the Home page to load data once for all pages.")
        return

    # ── Load: local upload takes priority over session state ──────────────────
    df_sfdc = None
    df_drs  = None
    df_ns   = None

    if sfdc_file:
        try:
            df_sfdc = load_sfdc(sfdc_file)
        except Exception as e:
            st.error(f"Could not load SFDC file: {e}")
    else:
        df_sfdc = _from_session.get("df_sfdc")

    if drs_file:
        try:
            df_drs = load_drs(drs_file)
            if _va_region:
                _rc = get_region_consultants(_va_region, EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE, ACTIVE_EMPLOYEES)
                _filtered = df_drs[df_drs["project_manager"].astype(str).str.strip().str.lower().isin(_rc)]
                df_drs = _filtered if not _filtered.empty else df_drs
            elif not is_manager or _va_name:
                if "project_manager" in df_drs.columns:
                    _filtered = df_drs[df_drs["project_manager"].apply(lambda v: name_matches(v, selected_user))]
                    df_drs = _filtered if not _filtered.empty else df_drs
        except Exception as e:
            st.error(f"Could not load DRS file: {e}")
    else:
        df_drs = _from_session.get("df_drs")
        if df_drs is not None:
            if _va_region:
                _rc = get_region_consultants(_va_region, EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE, ACTIVE_EMPLOYEES)
                _filtered = df_drs[df_drs["project_manager"].astype(str).str.strip().str.lower().isin(_rc)]
                df_drs = _filtered if not _filtered.empty else df_drs
            elif not is_manager or _va_name:
                if "project_manager" in df_drs.columns:
                    _filtered = df_drs[df_drs["project_manager"].apply(lambda v: name_matches(v, selected_user))]
                    df_drs = _filtered if not _filtered.empty else df_drs

    if ns_file:
        try:
            df_ns = load_ns_time(ns_file)
            if "employee" in df_ns.columns:
                df_ns = df_ns[df_ns["employee"].apply(lambda v: name_matches(v, selected_user))]
            _ns_id_col = "project_id" if "project_id" in df_ns.columns else "project" if "project" in df_ns.columns else None
        except Exception as e:
            st.error(f"Could not load NS file: {e}")
    else:
        df_ns = _from_session.get("df_ns")

    # ── Merge NS inactivity into DRS (outside try/except so errors are visible) ─
    if df_drs is not None and df_ns is not None:
        try:
            df_drs = calc_days_inactive(df_drs, df_ns)

        except Exception as e:
            st.warning(f"Could not calculate inactivity from NS data: {e}")

        # ── Summary metrics ───────────────────────────────────────────────────
    msg_parts = []
    if df_sfdc is not None:
        msg_parts.append(f"SFDC: {df_sfdc['account'].nunique() if 'account' in df_sfdc.columns else '?'} accounts · {df_sfdc['opportunity'].nunique() if 'opportunity' in df_sfdc.columns else '?'} opportunities")
    if df_drs is not None:
        msg_parts.append(f"DRS: {len(df_drs):,} active FF projects")
        # Tier summary from DRS
        tier_counts = {name: len(df_drs[(df_drs["days_inactive"] >= t["days_min"]) & (df_drs["days_inactive"] <= t["days_max"])]) for name, t in TEMPLATES.items()}
        tier_counts = {k: v for k, v in tier_counts.items() if v > 0}



    # ── Set mode based on what's loaded ───────────────────────────────────
    if df_drs is not None:
        df   = df_drs
        mode = "drs"
    else:
        df   = df_sfdc
        mode = "sfdc"

    # ── Ensure df is filtered to match the view selection ─────────────────
    # df_drs is already filtered above; apply same logic to df_sfdc if in sfdc mode
    if mode == "sfdc" and df is not None and "opportunity_owner" in df.columns:
        if _va_region:
            _rc = get_region_consultants(_va_region, EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE, ACTIVE_EMPLOYEES)
            _f = df[df["opportunity_owner"].astype(str).str.strip().str.lower().isin(_rc)]
            if not _f.empty: df = _f
        elif _va_name or not is_manager:
            _target = _va_name if _va_name else selected_user
            _f = df[df["opportunity_owner"].apply(lambda v: name_matches(v, _target))]
            if not _f.empty: df = _f

    # ── Project overview table ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Your Projects")
    if df_drs is not None and not df_drs.empty:
        today_ts = pd.Timestamp.today().normalize()

        # ── Metric cards ──────────────────────────────────────────────────
        _total    = len(df_drs)
        _onhold   = int(df_drs["_on_hold"].sum()) if "_on_hold" in df_drs.columns else 0
        _active   = _total - _onhold
        _on_hold_mask = df_drs.get("_on_hold", pd.Series(False, index=df_drs.index)).astype(bool)
        _inactive = int(((df_drs["days_inactive"] >= 14) & ~_on_hold_mask).sum()) if "days_inactive" in df_drs.columns else 0

        mc1, mc2, mc3, mc4 = st.columns(4)
        with mc1:
            st.metric("Total Projects", _total)
        with mc2:
            st.metric("Active Projects", _active)
        with mc3:
            st.metric("On Hold", _onhold)
        with mc4:
            st.metric("Requiring Follow-Up", _inactive, help="Active projects with 14+ days inactivity (on-hold excluded)")

        st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)
        with st.expander("How is 'Days Inactive' calculated?", expanded=False):
            st.markdown("""
**Priority 1 — NS Time Entry** *(most accurate)*
Days since the last time entry booked to this project in NetSuite. Requires NS Time Detail upload.

**Priority 2 — Last Milestone**
Days since the most recently completed milestone in the SS DRS (e.g. Enablement Session, UAT Signoff).
Used when no NS time entries are available.

**Priority 3 — Phase vs Start Date**
Days the project has been open, minus the expected benchmark duration for its current phase.
For example: a project in Onboarding (benchmark: 7 days) that's been open 204 days = 197 days stalled.
Used when no NS entries and no milestones are present.

**Unknown (-1)** — shown when none of the above signals are available.
            """)

        # ── This Week's Initial Engagement Actions ───────────────────────
        st.subheader("This Week's Initial Engagement Actions")
        st.caption("Non-legacy projects with no Intro. Email Sent date — first outreach needed.")

        _intro_df = pd.DataFrame()
        if "ms_intro_email" in df_drs.columns and "legacy" in df_drs.columns:
            _non_legacy  = ~df_drs["legacy"].astype(bool)
            _no_intro    = df_drs["ms_intro_email"].isna() | (df_drs["ms_intro_email"].astype(str).str.strip().isin(["", "nan", "None", "NaT"]))
            _active_mask = df_drs.get("status", pd.Series("Active", index=df_drs.index)).astype(str).str.lower().isin(["active", "in progress", "onboarding", "implementation", ""])
            _intro_df    = df_drs[_non_legacy & _no_intro & _active_mask].copy()
        elif "ms_intro_email" in df_drs.columns:
            _no_intro = df_drs["ms_intro_email"].isna() | (df_drs["ms_intro_email"].astype(str).str.strip().isin(["", "nan", "None", "NaT"]))
            _intro_df = df_drs[_no_intro].copy()

        if not _intro_df.empty:
            _intro_cols = [c for c in ["account","project_name","project_type","phase",
                                        "project_manager","start_date","days_inactive"]
                           if c in _intro_df.columns]
            _intro_display = _intro_df[_intro_cols].copy()
            # Format all columns — convert any datetime-like values to readable date strings
            for _dc in _intro_display.columns:
                try:
                    _parsed = pd.to_datetime(_intro_display[_dc], errors="coerce")
                    if _parsed.notna().any():
                        _intro_display[_dc] = _parsed.dt.strftime("%-d %b %Y").where(_parsed.notna(), "")
                except Exception:
                    pass
            _intro_display.columns = [c.replace("_"," ").title() for c in _intro_display.columns]
            st.dataframe(_intro_display, use_container_width=True, hide_index=True)
            st.caption(f"{len(_intro_df)} project(s) awaiting initial introduction email.")
        else:
            st.success("✓ All non-legacy projects have an intro email on record.")

        st.markdown("---")

        # ── Weekly Action List ────────────────────────────────────────────
        st.subheader("This Week's Re-Engagement Actions")

        # Projects needing action: 30+ days inactive, not on hold, not recently logged
        _logged_projects = {e.get("project") for e in _load_log()
                           if e.get("consultant") == selected_user
                           and (datetime.today() - datetime.strptime(e.get("date","2000-01-01"), "%Y-%m-%d")).days <= 29}

        _action_cols = [c for c in ["account","project_name","project_type","phase",
                                     "client_responsiveness","days_inactive","_on_hold",
                                     opp_col if "opp_col" in dir() else "project_name"]
                        if c in df_drs.columns]

        _action_df = df_drs[
            (df_drs["days_inactive"] >= 14) &
            (~df_drs.get("_on_hold", pd.Series(False, index=df_drs.index)))
        ].copy() if "days_inactive" in df_drs.columns else pd.DataFrame()

        if not _action_df.empty:
            # Add tier and last logged
            def _tier_short(d):
                try:
                    d = int(d)
                    if d < 60:  return "Tier 1"
                    if d < 90:  return "Tier 2"
                    if d < 180: return "Tier 3"
                    return "Tier 4"
                except: return "—"

            _action_df["Suggested Tier"] = _action_df["days_inactive"].apply(_tier_short)
            _all_log_entries = _load_log()  # still used for prior outreach indicator
            # Use project_id as stable key for recent-log matching; fall back to project_name
            _recent_key = "project_id" if "project_id" in _action_df.columns else                           "project_name" if "project_name" in _action_df.columns else None
            if _recent_key:
                _action_df["⚠️ Recent"] = _action_df[_recent_key].astype(str).isin(
                    {str(p) for p in _logged_projects}
                )
            else:
                _action_df["⚠️ Recent"] = False
            # Flag escalated projects
            if "risk_level" in _action_df.columns:
                _action_df["⚠️ Risk"] = _action_df["risk_level"].fillna("").astype(str).str.strip().str.lower().map(
                    lambda r: "⚠️ Escalated" if "escalat" in str(r) else ("🔴 High" if "high" in str(r) else "")
                )

            # Sort: Tier 4 → 3 → 2 → 1, then by days desc
            _tier_order = {"Tier 4": 0, "Tier 3": 1, "Tier 2": 2, "Tier 1": 3}
            _action_df["_tier_sort"] = _action_df["Suggested Tier"].map(_tier_order).fillna(4)
            _action_df = _action_df.sort_values(["_tier_sort", "days_inactive"], ascending=[True, False])
            # Format start_date for display
            if "start_date" in _action_df.columns:
                _action_df["start_date"] = pd.to_datetime(_action_df["start_date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("—")

            # Display columns
            _disp_cols = {
                "project_name":          "Project",
                "status":                "Status",
                "phase":                 "Phase",
                "start_date":            "Start Date",
                "days_inactive":         "Days Inactive",
                "last_milestone":        "Last Milestone",
                "_inactivity_source":    "Inactivity Source",
                "client_responsiveness": "Responsiveness",
                "risk_level":            "Risk Level",
                "risk_detail":           "Risk Detail",
                "Suggested Tier":        "Suggested Tier",
            }
            _avail = [k for k in _disp_cols if k in _action_df.columns or k in ["Suggested Tier","Last Follow Up","⚠️ Risk","_inactivity_source"]]
            _show  = _action_df[[col for col in _avail if col in _action_df.columns]].rename(columns=_disp_cols).reset_index(drop=True)
            _show.columns = [str(col) for col in _show.columns]
            for _sc in _show.columns:
                if _show[_sc].dtype == object:
                    _show[_sc] = _show[_sc].fillna("—").astype(str)

            # Colour by tier
            def _style_action(row):
                t = row.get("Suggested Tier","")
                if t == "Tier 4": bg, fg = "#e8d5ff","#4a235a"
                elif t == "Tier 3": bg, fg = "#FDECED","#9C0006"
                elif t == "Tier 2": bg, fg = "#FFEB9C","#9C6500"
                elif t == "Tier 1": bg, fg = "#C6EFCE","#276221"
                else: return [""] * len(row)
                return [f"background-color:{bg};color:{fg}"] * len(row)

            st.dataframe(
                _show.style.apply(_style_action, axis=1),
                hide_index=True, use_container_width=True
            )
            st.caption(f"{len(_action_df)} project(s) requiring re-engagement · sorted by urgency")

            # Warn about escalated projects
            if "⚠️ Risk" in _action_df.columns:
                _escalated = _action_df[_action_df["⚠️ Risk"].str.contains("Escalat", na=False)]
                if not _escalated.empty:
                    _esc_names = ", ".join(_escalated["project_name"].astype(str).tolist()[:3]) if "project_name" in _escalated.columns else "—"
                    st.warning(f"⚠️ **{len(_escalated)} project(s) marked as Escalated** — check with CS before sending: {_esc_names}")

        else:
            st.success("✅ No projects requiring re-engagement this week.")


    else:
        st.info("Upload SS DRS to see your project overview.")

    st.markdown("---")
    st.subheader("Step 1 — Select a Project")

    # ── Column resolution — works for both SFDC and DRS modes ─────────────
    if mode == "sfdc":
        opp_col  = "opportunity"      if "opportunity"      in df.columns else df.columns[0]
        acc_col  = "account"          if "account"          in df.columns else None
        prod_col = "product"          if "product"          in df.columns else None
        name_key = opp_col
    else:
        # Use project_id as the stable key; fall back to project_name only if project_id absent
        opp_col  = "project_id"       if "project_id"       in df.columns else \
                   "project_name"     if "project_name"     in df.columns else df.columns[0]
        acc_col  = "account"          if "account"          in df.columns else None
        prod_col = "project_type"     if "project_type"     in df.columns else None
        name_key = opp_col
        # Display label column — always prefer project_name for readability in dropdown
        _disp_col = "project_name" if "project_name" in df.columns else opp_col

    # Build project list — concat Account : Project Name for clarity
    if acc_col and acc_col in df.columns:
        df["_proj_label"] = (
            df[acc_col].fillna("").astype(str).str.strip()
            + "  —  "
            + df[_disp_col if mode == "drs" else opp_col].fillna("").astype(str).str.strip()
        )
    else:
        df["_proj_label"] = df[_disp_col if mode == "drs" else opp_col].fillna("").astype(str).str.strip()

    label_to_opp  = df[["_proj_label", opp_col]].drop_duplicates().set_index("_proj_label")[opp_col].to_dict()

    # For DRS: show tier badge next to each project in dropdown
    if mode == "drs" and "days_inactive" in df.columns:
        days_map = df.set_index(opp_col)["days_inactive"].to_dict()
        def _label_with_tier(label):
            opp = label_to_opp.get(label, "")
            days = days_map.get(opp, 0)
            suggested = suggest_tier_from_days(days)
            t_num = TEMPLATES[suggested]["tier"]
            tier_short = {1:"T1",2:"T2",3:"T3",4:"T4"}.get(t_num,"")
            return f"[{tier_short} · {int(days)}d]  {label}"
        # Only show projects ≥ 14 days inactive in dropdown (all shown in table above)
        proj_options_all = sorted(label_to_opp.keys(), key=lambda x: -days_map.get(label_to_opp.get(x,""), 0))
        # Build on_hold lookup from df
        _oh_map = {}
        if "_on_hold" in df.columns:
            _oh_map = dict(zip(df[opp_col].astype(str), df["_on_hold"]))

        proj_options_filtered = [
            p for p in proj_options_all
            if not _oh_map.get(label_to_opp.get(p,""), False)  # exclude On Hold only
        ]

        if not proj_options_filtered:
            st.info("No active projects found.")
            return

        def _label_with_tier_clean(label):
            opp       = label_to_opp.get(label, "")
            days      = days_map.get(opp, 0)
            suggested = suggest_tier_from_days(days)
            days_str  = f"{int(days)} days inactive"
            tier_str  = f"Tier {TEMPLATES[suggested]['tier']}" if suggested else "No Tier"
            return f"{tier_str}  ·  {days_str}  —  {label}"

        proj_options_display = [_label_with_tier_clean(p) for p in proj_options_filtered]
        display_to_label     = dict(zip(proj_options_display, proj_options_filtered))
        # Handle jump from action list button
        _jump = st.session_state.pop("_jump_to_proj", None)
        _jump_display = None
        if _jump:
            for _d, _l in display_to_label.items():
                if label_to_opp.get(_l,"") == _jump or _l == _jump:
                    _jump_display = _d
                    break

        selected_display     = st.selectbox(
            f"Select project ({len(proj_options_filtered)} active projects)",
            proj_options_display,
            index=proj_options_display.index(_jump_display) if _jump_display and _jump_display in proj_options_display else 0,
            key="proj_select"
        )
        selected_label = display_to_label[selected_display]
        selected_proj  = label_to_opp[selected_label]
    else:
        proj_options  = sorted(label_to_opp.keys())
        selected_label = st.selectbox("Project / Opportunity", proj_options, key="proj_select")
        selected_proj  = label_to_opp[selected_label]

    proj_rows = df[df[opp_col].astype(str) == selected_proj]

    # Clear days widget from session state when project changes so value resets
    if st.session_state.get("_last_proj") != selected_proj:
        st.session_state["_last_proj"] = selected_proj
        for _k in list(st.session_state.keys()):
            if _k.startswith("days_in_"):
                del st.session_state[_k]

    # Project metadata
    account  = proj_rows[acc_col].iloc[0]  if acc_col and acc_col in proj_rows.columns and not proj_rows.empty else ""
    product  = proj_rows[prod_col].iloc[0] if prod_col and prod_col in proj_rows.columns and not proj_rows.empty else ""
    # Account manager name (→ email body) and email (→ suggested CC)
    def _get_am_row():
        """Return first SFDC row for this project — best source for AM details."""
        if "account_manager" in proj_rows.columns and not proj_rows.empty:
            return proj_rows.iloc[0]
        if df_sfdc is not None and "account_manager" in df_sfdc.columns:
            if "opportunity" in df_sfdc.columns and "project_name" in proj_rows.columns and not proj_rows.empty:
                _nm = str(proj_rows["project_name"].iloc[0]).strip().lower()
                _match = df_sfdc[df_sfdc["opportunity"].astype(str).str.strip().str.lower() == _nm]
                if not _match.empty:
                    return _match.iloc[0]
            # Fall back to first row in full SFDC (all rows share same opp owner)
            return df_sfdc.iloc[0] if not df_sfdc.empty else None
        return None

    _am_row    = _get_am_row()
    am         = str(_am_row["account_manager"]).strip()       if _am_row is not None and "account_manager"       in _am_row.index and str(_am_row["account_manager"]).strip().lower()       not in ("nan","none","") else ""
    am_email   = str(_am_row["account_manager_email"]).strip() if _am_row is not None and "account_manager_email" in _am_row.index and str(_am_row["account_manager_email"]).strip().lower() not in ("nan","none","") else ""
    close_dt = proj_rows["close_date"].iloc[0] if "close_date" in proj_rows.columns and not proj_rows.empty else None

    # For DRS mode — pre-set days inactive and phase from data
    if mode == "drs" and "days_inactive" in proj_rows.columns:
        _drs_days  = int(proj_rows["days_inactive"].iloc[0]) if not proj_rows.empty else 30
        _drs_phase = str(proj_rows["phase"].iloc[0]).strip() if "phase" in proj_rows.columns and not proj_rows.empty else ""
    else:
        _drs_days  = 30
        _drs_phase = ""

    # Days inactive input
    st.markdown("---")
    st.subheader("Step 2 — Set Context")

    # ── Effective tier — considers both calculated days AND any template override ──
    # Template selector is in Step 4 (rendered later) but persists in session state
    _tmpl_ss_key = f"tmpl_select_{str(selected_proj)[:30].replace(chr(32),'_')}"
    _override_tmpl = st.session_state.get(_tmpl_ss_key)
    _override_tier = TEMPLATES[_override_tmpl]["tier"] if _override_tmpl and _override_tmpl in TEMPLATES else 0
    # Map tier → minimum days threshold for field visibility
    _tier_days_map = {1: 30, 2: 60, 3: 90, 4: 180}
    _effective_days = max(_drs_days if _drs_days and _drs_days >= 0 else 0,
                          _tier_days_map.get(_override_tier, 0))

    # Try to derive product from SFDC data, allow manual override
    _sfdc_product = normalise_product_name(product) if product else ""

    col0, = st.columns([1])
    with col0:
        # Key includes selected_proj so widget resets when project changes
        _prod_key = f"product_name_{hash(selected_proj) if 'selected_proj' in dir() else 0}"
        product_name = st.text_input(
            "Product(s) being implemented *",
            value=_sfdc_product,
            placeholder="e.g. ZoneCapture, ZoneApprovals",
            key=_prod_key
        )

    if not product_name or product_name.strip() == "":
        st.warning("⚠️ Please enter the product(s) being implemented before continuing.")
        return

    col1, col2, col3 = st.columns([2, 2, 2])
    with col1:
        _days_label = "Days since last customer contact"
        if _drs_days == -1:
            _days_label += " (could not be calculated — please enter manually)"
            _drs_days = 30
        days_inactive = st.number_input(_days_label, min_value=0, value=_drs_days, step=1, key=f"days_in_{str(selected_proj)[:40].replace(chr(32),'_').replace(chr(39),'')}")
        if _drs_days != days_inactive:
            st.caption(f"✏️ Overridden — using {int(days_inactive)} days instead of calculated value")
    with col2:
        current_phase = st.selectbox(
            "Current project phase",
            ['00. Onboarding', '01. Requirements and Design', '02. Configuration', '03. Enablement/Training', '04. UAT', '05. Prep for Go-Live', '06. Go-Live (Hypercare)', '08. Ready for Support Transition', '09. Phase 2 Scoping', '10. Complete/Pending Final Billing', '11. On Hold'],
            index=next((i for i, p in enumerate([
                "00. Onboarding","01. Requirements and Design","02. Configuration",
                "03. Enablement/Training","04. UAT","05. Prep for Go-Live",
                "06. Go-Live (Hypercare)","08. Ready for Support Transition",
                "09. Phase 2 Scoping","10. Complete/Pending Final Billing","11. On Hold",
            ]) if _drs_phase.lower().startswith(p[:6].lower())), 2),
            key=f"phase_in_{str(selected_proj)[:40].replace(chr(32),'_').replace(chr(39),'')}"
        )
    with col3:
        _last_act_key = f"last_act_{str(selected_proj)[:40].replace(' ','_').replace(chr(39),'')}"
        if _last_act_key not in st.session_state:
            st.session_state[_last_act_key] = None
        last_activity_val = st.date_input(
            "Date of last activity *",
            value=st.session_state.get(_last_act_key),
            key=_last_act_key,
            help="Date you last heard from or engaged with the customer"
        )
        if last_activity_val is None:
            st.warning("⚠️ Please enter the date of last customer activity before continuing.")
            return
        last_activity = last_activity_val

    col4, col5 = st.columns([3, 3])
    with col4:
        # Flip "Last, First" → "First Last" for email signature
        _parts = selected_user.split(",", 1)
        _display_name = f"{_parts[1].strip()} {_parts[0].strip()}" if len(_parts) == 2 else selected_user
        consultant_name = st.text_input("Your name (Implementation Consultant)", value=_display_name, key="ic_name")
    with col5:
        if _effective_days >= 180:
            # Pre-fill from DRS milestone calculation — use df_drs directly
            # (proj_rows may have been reassigned to SFDC rows losing DRS columns)
            _drs_remaining = ""
            if df_drs is not None and "remaining_sessions" in df_drs.columns:
                # Match by project_id or project_name
                _drs_proj = None
                if "project_id" in df_drs.columns and "project_id" in df.columns:
                    _pid = str(df[df[opp_col].astype(str) == str(selected_proj)]["project_id"].iloc[0]) if not df[df[opp_col].astype(str) == str(selected_proj)].empty else None
                    if _pid:
                        _drs_proj = df_drs[df_drs["project_id"].astype(str) == _pid]
                if (_drs_proj is None or _drs_proj.empty) and "project_name" in df_drs.columns:
                    _drs_proj = df_drs[df_drs["project_name"].astype(str) == str(selected_proj)]
                if _drs_proj is not None and not _drs_proj.empty:
                    _rs_val = _drs_proj["remaining_sessions"].iloc[0]
                    if _rs_val is not None and str(_rs_val) not in ("nan","None",""):
                        _drs_remaining = str(int(_rs_val))
            remaining_sessions = st.text_input(
                "Remaining sessions",
                value=_drs_remaining,
                placeholder="e.g. 3",
                help="Calculated from SS DRS milestone columns (Enablement Session, Session #1, Session #2)",
                key=f"rem_sess_{str(selected_proj)[:20]}"
            )
        else:
            remaining_sessions = ""

    # Service term expiry — only shown for Tier 4
    if _effective_days >= 180:
        # Calculate subscription_start_date + 12m if available
        _sub_start = None
        if df_drs is not None and "subscription_start_date" in df_drs.columns:
            _sub_row = df_drs[df_drs[opp_col].astype(str) == str(selected_proj)] if opp_col in df_drs.columns else pd.DataFrame()
            if not _sub_row.empty:
                _sv = pd.to_datetime(_sub_row["subscription_start_date"].iloc[0], errors="coerce")
                if pd.notna(_sv):
                    _sub_start = (_sv + pd.DateOffset(months=12)).date()

        _svc_key = f"svc_exp_{str(selected_proj)[:30].replace(' ','_')}"
        # Only set session state default once per project — don't overwrite user edits
        if _svc_key not in st.session_state:
            st.session_state[_svc_key] = _sub_start

        if _sub_start:
            # Subscription date found — pre-populate with calculated expiry
            expiry_val = st.date_input(
                "Service term expiry date",
                value=st.session_state.get(_svc_key) or _sub_start,
                key=_svc_key,
                help="Auto-calculated: Start Date (Subscription) + 12 months. Edit if needed."
            )
            service_expiry = expiry_val
        else:
            # No subscription date in DRS — required manual entry
            st.markdown("**Service term expiry date** :red[*]")
            st.caption("Not found in DRS — enter the contract expiry date to continue (contract start + 12 months).")
            _manual = st.date_input(
                "Service term expiry date",
                value=None,
                key=_svc_key,
                help="12 months from contract/subscription start date",
                label_visibility="collapsed",
            )
            if _manual is None:
                st.error("Service term expiry date is required for Tier 4 outreach. Please enter a date to continue.")
                st.stop()
            service_expiry = _manual
    else:
        service_expiry = date.today()

    # ── Contacts ──────────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Step 3 — Select Recipients")

    # Safe defaults — overridden below if contacts are found
    to_emails             = []
    cc_emails             = []
    primary_contact_first = ""

    if mode == "drs" and df_sfdc is not None:
        _proj_nm   = str(proj_rows["project_name"].iloc[0]).strip() if not proj_rows.empty and "project_name" in proj_rows.columns else ""
        # Use account if available, otherwise fall back to project_name for account matching
        _acct_hint = str(account).strip() if account and str(account).strip() not in ("", "nan") else _proj_nm
        _sfdc_match, _match_label = fuzzy_match_sfdc(df_sfdc, _proj_nm, _acct_hint)
        if not _sfdc_match.empty:
            proj_rows = _sfdc_match.copy()
            mode      = "sfdc"
            st.caption(f"✅ SFDC contacts matched — {len(_sfdc_match)} contact(s) · {_match_label}")

        else:
            st.info("No SFDC contacts matched for this project. Account name or product may not overlap — add contacts manually.")
            to_emails             = []
            cc_emails             = []
            primary_contact_first = ""
    elif mode == "drs":
        st.info("Upload your SFDC Contacts file alongside the DRS to enable contact lookup.")

    if mode == "sfdc":
        email_col   = "email"             if "email"             in proj_rows.columns else None
        name_col    = "contact_name"      if "contact_name"      in proj_rows.columns else None
        roles_col   = "contact_roles"     if "contact_roles"     in proj_rows.columns else None
        flag_col    = "impl_contact_flag" if "impl_contact_flag" in proj_rows.columns else None
        partner_col = "partner_contact"   if "partner_contact"   in proj_rows.columns else None
        count_col   = "role_count"        if "role_count"        in proj_rows.columns else None

        if count_col and not proj_rows.empty:
            expected = pd.to_numeric(proj_rows[count_col].iloc[0], errors="coerce")
            actual   = len(proj_rows)
            if pd.notna(expected) and actual < int(expected):
                st.warning(f"⚠️ SFDC shows {int(expected)} contacts for this opportunity but only {actual} row(s) loaded — export may be missing contacts.")

        if email_col and name_col:
            # Build display cols — deduplicate to avoid Arrow errors
            _dcols_raw = [name_col, email_col, roles_col, flag_col]
            if "title" in proj_rows.columns and "title" not in _dcols_raw:
                _dcols_raw.append("title")
            if "is_primary" in proj_rows.columns and "is_primary" not in _dcols_raw:
                _dcols_raw.append("is_primary")
            display_cols = list(dict.fromkeys([c for c in _dcols_raw if c and c in proj_rows.columns]))
            contacts_display = proj_rows[display_cols].drop_duplicates().reset_index(drop=True)
            col_rename = {
                "contact_name": "Name", "email": "Email", "contact_roles": "Contact Roles",
                "title": "Title", "impl_contact_flag": "Impl. Contact ✓",
                "is_primary": "Primary ✓"
            }
            contacts_display = contacts_display.rename(columns={k: v for k, v in col_rename.items() if k in contacts_display.columns})
            # Final dedup of column names after rename
            contacts_display = contacts_display.loc[:, ~contacts_display.columns.duplicated()]
            st.dataframe(contacts_display, hide_index=True, use_container_width=True)

            all_emails = proj_rows[email_col].dropna().astype(str).tolist()

            # Priority:
            # 1. Implementation Contact checkbox (impl_contact_flag)
            # 2. Contact Roles contains "Implementation Contact"
            # 3. Primary checkbox (is_primary)
            # 4. Contact Roles contains "Primary"
            # 5. First contact (fallback)
            to_source    = "First contact (fallback)"
            suggested_to = all_emails[:1]

            primary_col = "is_primary" if "is_primary" in proj_rows.columns else None

            if flag_col:
                flagged = proj_rows[proj_rows[flag_col] == True]
                if not flagged.empty:
                    suggested_to = flagged[email_col].dropna().astype(str).tolist()
                    to_source    = "Implementation Contact checkbox"
            if to_source == "First contact (fallback)" and roles_col:
                impl_r = proj_rows[proj_rows[roles_col].str.contains("Implementation Contact", case=False, na=False)]
                if not impl_r.empty:
                    suggested_to = impl_r[email_col].dropna().astype(str).tolist()
                    to_source    = "Contact Roles: Implementation Contact"
            if to_source == "First contact (fallback)" and primary_col:
                prim_flag = proj_rows[proj_rows[primary_col] == "✓"]
                if not prim_flag.empty:
                    suggested_to = prim_flag[email_col].dropna().astype(str).tolist()
                    to_source    = "Primary contact checkbox"
            if to_source == "First contact (fallback)" and roles_col:
                prim_r = proj_rows[proj_rows[roles_col].str.contains("Primary", case=False, na=False)]
                if not prim_r.empty:
                    suggested_to = prim_r[email_col].dropna().astype(str).tolist()
                    to_source    = "Contact Roles: Primary"

            partner_emails = []
            if partner_col:
                partner_vals   = proj_rows[partner_col].dropna().astype(str)
                partner_emails = [v.strip() for v in partner_vals if v.strip() and v.strip().lower() not in ("nan","none","")]

            to_emails = st.multiselect("To:", all_emails, default=[e for e in suggested_to if e in all_emails], key="to_emails")
            st.caption(f"To: auto-selected via — {to_source}")
            _extra_to = st.text_input("Add additional To: address", placeholder="name@example.com", key="extra_to")
            if _extra_to and "@" in _extra_to and _extra_to not in to_emails:
                to_emails = to_emails + [_extra_to.strip()]

            # Build CC suggestions: partner contact + AM email (Tier 2+)
            _suggested_t = suggest_tier_from_days(int(days_inactive)) if "days_inactive" in dir() else None
            tier_num     = TEMPLATES[_suggested_t]["tier"] if _suggested_t else 1
            _am_cc       = [am_email] if am_email and tier_num >= 2 else []
            suggested_cc = list(dict.fromkeys(
                [e for e in partner_emails if e not in to_emails] + _am_cc
            ))
            cc_pool  = list(dict.fromkeys(all_emails + partner_emails + ([am_email] if am_email else [])))
            _cc_hints = []
            if [e for e in partner_emails if e not in to_emails]: _cc_hints.append("partner contact")
            if am_email and tier_num >= 2: _cc_hints.append(f"AM ({am}) pre-suggested for Tier {tier_num}+")
            cc_label  = "CC:" + (f" — {', '.join(_cc_hints)}" if _cc_hints else "")
            cc_emails = st.multiselect(cc_label, cc_pool, default=suggested_cc, key="cc_emails")
            if _cc_hints:
                st.caption("CC pre-suggestions — review before sending")
            _extra_cc = st.text_input("Add additional CC: address", placeholder="name@example.com", key="extra_cc")
            if _extra_cc and "@" in _extra_cc and _extra_cc not in cc_emails:
                cc_emails = cc_emails + [_extra_cc.strip()]
        else:
            st.warning("Email and/or contact name columns not detected. Check your export headers.")

            to_emails = []
            cc_emails = []

        # Primary contact first name from To: recipient
        primary_contact_first = ""
        if name_col and not proj_rows.empty:
            if email_col and to_emails:
                to_row    = proj_rows[proj_rows[email_col].astype(str) == to_emails[0]]
                full_name = str(to_row[name_col].iloc[0]) if not to_row.empty else str(proj_rows[name_col].iloc[0])
            else:
                full_name = str(proj_rows[name_col].iloc[0])
            primary_contact_first = full_name.split()[0] if full_name and full_name.lower() not in ("nan","") else ""
    # ── Template selection ────────────────────────────────────────────────────
    st.markdown("---")

    # ── Prior outreach indicator ──────────────────────────────────────────
    _prior = _get_project_log(selected_proj)
    if _prior:
        _prior_df = pd.DataFrame(_prior).sort_values("date", ascending=False)
        with st.expander(f"📋 Prior outreach logged — {len(_prior)} email(s) sent", expanded=True):
            st.dataframe(_prior_df[["date","consultant","tier","days_inactive","template"]].rename(columns={
                "date": "Date", "consultant": "Sent By", "tier": "Tier",
                "days_inactive": "Days Inactive", "template": "Template Used"
            }), hide_index=True, use_container_width=True)
        # Tier sequencing guard
        _logged_tiers = {e.get("tier","") for e in _prior}
        _tier1_sent   = any("Tier 1" in t for t in _logged_tiers)
        _tier2_sent   = any("Tier 2" in t for t in _logged_tiers)
        _tier3_sent   = any("Tier 3" in t for t in _logged_tiers)
    else:
        _tier1_sent = _tier2_sent = _tier3_sent = False
        if int(days_inactive) >= 60:
            st.warning("⚠️ No prior outreach logged for this project. Per sequencing rules, start with Tier 1 before escalating — unless Tier 1 was sent outside this tool.")

    st.subheader("Step 4 — Template")

    suggested = suggest_tier(int(days_inactive)) or list(TEMPLATES.keys())[0]
    tier_names = list(TEMPLATES.keys())
    suggested_idx = tier_names.index(suggested) if suggested in tier_names else 0

    t1, t2 = st.columns([3, 3])
    with t1:
        tier_num = TEMPLATES[suggested]["tier"]
        st.markdown(f"<div style='margin-bottom:6px'>Auto-suggested based on <b>{int(days_inactive)} days</b> inactive:</div>", unsafe_allow_html=True)
        st.markdown(f"<span class='tier-badge-{tier_num}'>{suggested}</span>", unsafe_allow_html=True)
    with t2:
        selected_template = st.selectbox("Override template if needed:", tier_names,
                                          index=suggested_idx,
                                          key=f"tmpl_select_{str(selected_proj)[:30].replace(chr(32),'_')}",
                                          label_visibility="visible")

    tmpl_info = TEMPLATES[selected_template]
    tier_num  = tmpl_info["tier"]

    # CC guidance
    cc_raw = tmpl_info["cc_guidance"].replace("{ACCOUNT MANAGER}", am if am else "{ACCOUNT MANAGER}")
    st.markdown(f"<div class='cc-box'>📋 <b>CC guidance:</b> {cc_raw}</div>", unsafe_allow_html=True)

    if tier_num in (3, 4) and not am:
        if df_sfdc is None:
            st.info("Upload your SFDC Contacts file to auto-populate the Account Manager name.")
        else:
            st.warning("Account Manager not found in SFDC data — fill {ACCOUNT MANAGER} manually in the template.")

    # ── Build filled template ─────────────────────────────────────────────────
    # Strip numeric prefix from phase for clean template insertion (e.g. "02. Configuration" → "Configuration")
    _phase_clean = re.sub(r'^\d+\.\s*', '', current_phase).strip()

    fields = {
        "CUSTOMER CONTACT NAME":  primary_contact_first,
        "PRODUCT NAME":           product_name,
        "CURRENT PHASE":          _phase_clean,
        "LAST ACTIVITY DATE":     last_activity.strftime("%B %d, %Y"),
        "IMPLEMENTATION CONSULTANT": consultant_name,
        "ACCOUNT MANAGER":        str(am) if am else "",
        "REMAINING SESSIONS":     remaining_sessions,
        "SERVICE TERM EXPIRY":    service_expiry.strftime("%B %d, %Y") if service_expiry else "[SERVICE TERM EXPIRY — please enter manually]",
    }

    subject, body, _ = fill_template(selected_template, fields)

    # ── Preview ───────────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Step 5 — Review & Send")

    # Remaining placeholders
    remaining = extract_placeholders(body) + extract_placeholders(subject)
    if remaining:
        st.warning(f"**{len(set(remaining))} placeholder(s) still need manual update:** {', '.join(f'{{{p}}}' for p in sorted(set(remaining)))}")
    else:
        st.success("All placeholders filled — ready to send.")

    # Email header display
    if to_emails:
        st.markdown(f"**To:** {', '.join(to_emails)}")
    if cc_emails:
        st.markdown(f"**CC:** {', '.join(cc_emails)}")
    st.markdown(f"**Subject:** {subject}")
    st.markdown("---")

    # Body preview with highlighted placeholders
    highlighted = highlight_placeholders(body.replace('\n', '<br>'))
    st.markdown(f"<div class='email-box'>{highlighted}</div>", unsafe_allow_html=True)

    # Copy button — plain text version
    st.markdown("<div style='margin-top:16px'></div>", unsafe_allow_html=True)
    # ── mailto link ───────────────────────────────────────────────────────────
    import urllib.parse
    _to_str  = ",".join(to_emails)  if to_emails  else ""
    _cc_str  = ",".join(cc_emails)  if cc_emails  else ""
    _subject = urllib.parse.quote(subject)
    _body    = urllib.parse.quote(body)
    _mailto  = f"mailto:{_to_str}?subject={_subject}&body={_body}"
    if _cc_str:
        _mailto = f"mailto:{_to_str}?cc={urllib.parse.quote(_cc_str)}&subject={_subject}&body={_body}"

    st.markdown("<div style='margin-top:12px'></div>", unsafe_allow_html=True)
    st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)
    btn_col1, btn_col2, btn_col3 = st.columns([3, 3, 3])
    with btn_col1:
        st.markdown(
            f"<a href='{_mailto}' target='_blank'>"
            f"<button style='background:#1e2c63;color:white;border:none;padding:10px 0;border-radius:6px;font-family:Manrope,sans-serif;font-size:14px;font-weight:600;cursor:pointer;width:100%;'>✉️ Open in Email</button></a>",
            unsafe_allow_html=True
        )
    with btn_col3:
        if st.button("📋 Log this outreach", key="log_btn", type="primary",
                     use_container_width=True):
            _tier_label = TEMPLATES.get(selected_template, {}).get("tier", "")
            _tier_str   = f"Tier {_tier_label}" if _tier_label else selected_template
            _log_customer = str(account) if account and str(account).lower() not in ("nan","") else selected_proj
            _log_outreach(
                consultant    = selected_user,   # store as "Last, First" for consistent filtering
                customer      = _log_customer,
                project       = selected_proj,
                tier_label    = _tier_str,
                days_inactive = days_inactive,
                template      = selected_template,
            )
            st.session_state["_log_success_msg"] = f"✅ Logged — {_tier_str} for {_log_customer} on {datetime.today().strftime('%Y-%m-%d')}"

        if st.session_state.get("_log_success_msg"):
            st.success(st.session_state["_log_success_msg"])
    with btn_col2:
        import json as _json2
        import streamlit.components.v1 as _components
        _html_wrapped = (
            "<div style='font-family:Manrope,Arial,sans-serif;font-size:14px;"
            "line-height:1.7;color:#1e2c63;'>"
            + highlighted + "</div>"
        )
        _h = _json2.dumps(_html_wrapped)
        _p = _json2.dumps(body)
        _components.html(f"""
<!DOCTYPE html><html><body style="margin:0;padding:0;font-family:Manrope,sans-serif;">
<button id="cb" style="background:#1e2c63;color:white;border:none;padding:10px 0;border-radius:6px;font-family:Manrope,sans-serif;font-size:14px;font-weight:600;cursor:pointer;width:100%;">&#128196; Copy Email Content (Formatted)</button>
<span id="st" style="margin-left:8px;font-size:13px;"></span>
<script>
var h={_h};
var p={_p};
document.getElementById("cb").addEventListener("click",function(){{
  navigator.clipboard.write([new ClipboardItem({{
    "text/html":new Blob([h],{{type:"text/html"}}),
    "text/plain":new Blob([p],{{type:"text/plain"}})
  }})]).then(function(){{
    document.getElementById("st").innerText="\u2705 Copied!";
    setTimeout(function(){{document.getElementById("st").innerText=""}},3000);
  }}).catch(function(){{
    document.getElementById("st").innerText="\u26a0\ufe0f Not supported — use plain text";
  }});
}});
</script></body></html>""", height=60)


    # ── Merge field reference ─────────────────────────────────────────────────
    with st.expander("📋 Merge field reference — what was filled vs. what still needs updating", expanded=False):
        ref_rows = []
        for k, v in fields.items():
            status = "✅ Filled" if v else "⚠️ Needs manual update"
            ref_rows.append({"Placeholder": f"{{{k}}}", "Value Used": v or "—", "Status": status})
        st.dataframe(pd.DataFrame(ref_rows), hide_index=True, use_container_width=True)

    # ── Tier guidance ─────────────────────────────────────────────────────────
    # ── Full outreach log ────────────────────────────────────────────────────
    st.markdown("---")
    with st.expander("📊 Full Outreach Log", expanded=False):
        st.caption("⚠️ Log entries are session-only and do not persist between logins. Download and record in your SmartSheet project tracker after each session.")
        _all_log = list(st.session_state.get(_LOG_KEY, []))
        if not _all_log:
            try:
                if os.path.exists(LOG_PATH):
                    with open(LOG_PATH) as _f:
                        _all_log = _json.load(_f)
            except: pass
        if _all_log:
            _log_df = pd.DataFrame(_all_log).sort_values("date", ascending=False)
            # Filter to current user — match both "First Last" and "Last, First" formats
            if selected_user and selected_user != "— Select —":
                _parts = selected_user.split(",", 1)
                _display = f"{_parts[1].strip()} {_parts[0].strip()}" if len(_parts) == 2 else selected_user
                _log_df = _log_df[
                    (_log_df["consultant"] == selected_user) |
                    (_log_df["consultant"] == _display)
                ]
            st.dataframe(
                _log_df[["date","consultant","customer","project","tier","days_inactive","template"]].rename(columns={
                    "date": "Date", "consultant": "Consultant", "customer": "Customer",
                    "project": "Project", "tier": "Tier",
                    "days_inactive": "Days Inactive", "template": "Template"
                }),
                hide_index=True, use_container_width=True
            )
            _cl1, _cl2 = st.columns([2, 2])
            with _cl1:
                if st.button("🗑 Clear my entries", key="clear_log", use_container_width=True):
                    _kept = [e for e in st.session_state.get(_LOG_KEY, [])
                             if e.get("consultant") != selected_user]
                    st.session_state[_LOG_KEY] = _kept
                    try:
                        with open(LOG_PATH, "w") as _lf:
                            _json.dump(_kept, _lf)
                    except: pass
                    st.rerun()
            with _cl2:
                if is_manager and st.button("🗑 Clear ALL entries", key="clear_all_log",
                                             use_container_width=True, type="primary"):
                    st.session_state[_LOG_KEY] = []
                    try:
                        with open(LOG_PATH, "w") as _lf:
                            _json.dump([], _lf)
                    except: pass
                    st.rerun()
        else:
            st.info("No outreach logged yet. Click '📋 Log this outreach' after composing an email.")

    with st.expander("ℹ️ Tier guidance & usage notes", expanded=False):
        st.markdown("""
**Tier 1 (~30 days)** — Keep it short and conversational. No contractual language. Goal is simply to prompt a reply.

**Tier 2 (~60 days)** — Introduces scope awareness without being heavy-handed. CC PS Leadership.

**Tier 3 (~90 days)** — First email where Account Manager is CC'd. Escalates practical consequences but avoids contractual terms. Also CC CS Manager and PS Leadership.

**Tier 4 (~6 months)** — Only email referencing service term, expiry date, and session forfeiture. Always CC Account Manager and CS Manager.

---

⚠️ **Always send tiers in sequential order.** If 60 days have passed but no Tier 1 has been sent, start with Tier 1.

⚠️ **Check with CS Manager before sending Tier 3 or 4** — they may have context about the customer's situation.

📝 **Log all outreach in your SmartSheet project tracker** for audit trail purposes.
        """)


if __name__ == "__main__":
    main()
