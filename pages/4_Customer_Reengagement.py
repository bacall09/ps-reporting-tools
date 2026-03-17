import streamlit as st
import pandas as pd
from datetime import date, datetime
import re
from rapidfuzz import fuzz

st.set_page_config(page_title="Customer Re-Engagement", page_icon=None, layout="wide")

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
st.markdown("""
    <div style='background-color:#1e2c63;padding:24px 32px;border-radius:8px;margin-bottom:24px;font-family:Manrope,sans-serif'>
        <h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif'>Customer Re-Engagement</h1>
        <p style='color:#aac4d0;margin:6px 0 0 0;font-size:14px;font-family:Manrope,sans-serif'>Re-engagement communications for unresponsive customers · Tier-based templates · Auto-suggests based on inactivity</p>
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
    "contact":                  "contact_name",
    "contact email":            "email",
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
    "project id":             "project_id",
    "project phase":          "phase",
    "project type":           "project_type",
    "status":                 "status",
    "start date":             "start_date",
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
    "enablement session":     "ms_enablement",
    "session #1":             "ms_session1",
    "session #2":             "ms_session2",
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

    # Filter to active FF projects only
    today = pd.Timestamp.today().normalize()
    if "phase" in df.columns:
        df = df[~df["phase"].str.strip().str.lower().isin(INACTIVE_PHASES_OUT)]
    if "billing_type" in df.columns:
        df = df[~df["billing_type"].str.strip().str.lower().isin({"t&m","time & material","time and material"})]
    # Calculate remaining sessions from milestone columns
    SESSION_MILESTONE_COLS = ["ms_enablement", "ms_session1", "ms_session2"]
    _present = [col for col in SESSION_MILESTONE_COLS if col in df.columns]
    if _present:
        def _count_remaining(row):
            return sum(1 for col in _present if pd.isna(row[col]) or str(row[col]).strip() in ("", "nan", "None"))
        df["remaining_sessions"] = df.apply(_count_remaining, axis=1)
    else:
        df["remaining_sessions"] = None

    # Tag On Hold but keep in df — excluded from dropdown, shown in table
    ON_HOLD_VALS = {"on-hold","on hold","onhold","on_hold"}
    if "status" in df.columns:
        df["_on_hold"] = df["status"].str.strip().str.lower().isin(ON_HOLD_VALS)
    else:
        df["_on_hold"] = False

    # Calculate days inactive
    # Only use last_updated if it mapped — do NOT fall back to start_date
    # (start_date is not an inactivity signal)
    if "last_updated" in df.columns:
        df["days_inactive"]        = (today - df["last_updated"]).dt.days.clip(lower=0).fillna(-1).astype(int)
        df["_inactivity_source"]   = "SS DRS Modified"
    else:
        # No reliable inactivity signal from DRS alone — set to -1 (Unknown)
        # Will be overridden by NS Time Detail if uploaded
        df["days_inactive"]        = -1
        df["_inactivity_source"]   = "Unknown — upload NS Time Detail"

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
        _fallback_src = str(df_drs["_inactivity_source"].iloc[0]) if "_inactivity_source" in df_drs.columns else "SS DRS Modified"
        df_drs["_inactivity_source"] = ["NS Time Entry" if pd.notna(x) else _fallback_src
                                         for x in df_drs["last_ns_entry"]]
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



    # ── Who is using this tool? ───────────────────────────────────────────
    st.subheader("Step 1 — Who are you?")
    selected_user = st.selectbox(
        "Select your name",
        ["— Select —"] + ACTIVE_EMPLOYEES,
        key="selected_user"
    )
    if selected_user == "— Select —":
        st.info("Select your name to filter projects to your portfolio.")
        return

    st.markdown("---")
    st.subheader("Step 2 — Upload Reports")
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

    if not sfdc_file and not drs_file and not ns_file:
        st.info("Upload at least one report. All three together give the most accurate inactivity detection.")
        return

    # ── Load whichever files were uploaded ────────────────────────────────
    df_sfdc = None
    df_drs  = None
    df_ns   = None

    if sfdc_file:
        try:
            df_sfdc = load_sfdc(sfdc_file)
        except Exception as e:
            st.error(f"Could not load SFDC file: {e}")

    if drs_file:
        try:
            df_drs = load_drs(drs_file)
            # Filter DRS to selected user
            if "project_manager" in df_drs.columns:
                df_drs = df_drs[df_drs["project_manager"].astype(str).str.strip() == selected_user]
                if df_drs.empty:
                    st.warning(f"No active FF projects found for {selected_user} in the DRS export.")
        except Exception as e:
            st.error(f"Could not load DRS file: {e}")

    if ns_file:
        try:
            df_ns = load_ns_time(ns_file)
            # Filter NS to selected user
            if "employee" in df_ns.columns:
                df_ns = df_ns[df_ns["employee"].astype(str).str.strip() == selected_user]
            # Show what ID column was found for join debugging
            _ns_id_col = "project_id" if "project_id" in df_ns.columns else "project" if "project" in df_ns.columns else None
        except Exception as e:
            st.error(f"Could not load NS file: {e}")

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

    # ── Project overview table ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Your Projects")
    if df_drs is not None and not df_drs.empty:
        today_ts = pd.Timestamp.today().normalize()

        # ── Metric cards ──────────────────────────────────────────────────
        _total    = len(df_drs)
        _onhold   = int(df_drs["_on_hold"].sum()) if "_on_hold" in df_drs.columns else 0
        _active   = _total - _onhold
        _inactive = int((df_drs["days_inactive"] >= 30).sum()) if "days_inactive" in df_drs.columns else 0

        mc1, mc2, mc3, mc4 = st.columns(4)
        with mc1:
            st.metric("Total Projects", _total)
        with mc2:
            st.metric("Active Projects", _active)
        with mc3:
            st.metric("On Hold", _onhold)
        with mc4:
            st.metric("Requiring Follow-Up", _inactive, help="Projects with 30+ days inactivity")

        st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)

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
            (df_drs["days_inactive"] >= 30) &
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
            _action_df["Last Logged"]    = _action_df["project_name"].apply(
                lambda p: next((e["date"] for e in sorted(_load_log(), key=lambda x: x.get("date",""), reverse=True)
                               if e.get("project") == str(p)), "—")
            ) if "project_name" in _action_df.columns else "—"
            _action_df["⚠️ Recent"] = _action_df["project_name"].isin(_logged_projects)

            # Sort: Tier 4 → 3 → 2 → 1, then by days desc
            _tier_order = {"Tier 4": 0, "Tier 3": 1, "Tier 2": 2, "Tier 1": 3}
            _action_df["_tier_sort"] = _action_df["Suggested Tier"].map(_tier_order).fillna(4)
            _action_df = _action_df.sort_values(["_tier_sort", "days_inactive"], ascending=[True, False])

            # Display columns
            _disp_cols = {
                "account":               "Customer",
                "project_name":          "Project",
                "phase":                 "Phase",
                "days_inactive":         "Days Inactive",
                "client_responsiveness": "Responsiveness",
                "Suggested Tier":        "Suggested Tier",
                "Last Logged":           "Last Logged",
            }
            _avail = [k for k in _disp_cols if k in _action_df.columns or k in ["Suggested Tier","Last Logged"]]
            _show  = _action_df[[c for c in _avail if c in _action_df.columns]].rename(columns=_disp_cols)

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

            # Per-row compose buttons
            st.markdown("**Compose email for:**")
            _btn_cols = st.columns(min(len(_action_df), 4))
            for _i, (_idx, _row) in enumerate(_action_df.head(8).iterrows()):
                _proj_id   = str(_row.get("project_name", ""))
                _proj_label = f"{_row.get('account', _proj_id)[:20]} · {_row.get('Suggested Tier','')}"
                with _btn_cols[_i % 4]:
                    if st.button(_proj_label, key=f"action_btn_{_i}", use_container_width=True):
                        st.session_state["_jump_to_proj"] = _proj_id
                        st.rerun()
        else:
            st.success("✅ No projects requiring re-engagement this week.")

        st.markdown("---")
        st.subheader("All Projects")

        overview_cols = {
            "account":                "Customer",
            "project_name":           "Project",
            "project_type":           "Project Type",
            "start_date":             "Start Date",
            "status":                 "Status",
            "phase":                  "Current Phase",
            "client_responsiveness":  "Client Responsiveness",
            "days_inactive":          "Days Inactive",
        }
        # Last Time Entry — from NS if available, else DRS Modified date
        if "last_ns_entry" in df_drs.columns:
            df_drs["last_time_entry"] = df_drs["last_ns_entry"].dt.strftime("%Y-%m-%d").fillna("—")
            df_drs["entry_source"]    = df_drs["last_ns_entry"].apply(
                lambda x: "NS Time Entry" if pd.notna(x) else
                          ("SS DRS Modified" if "last_updated" in df_drs.columns else "—")
            )
            # For rows with no NS entry, show DRS modified date instead
            if "last_updated" in df_drs.columns:
                mask = df_drs["last_ns_entry"].isna()
                df_drs.loc[mask, "last_time_entry"] = pd.to_datetime(
                    df_drs.loc[mask, "last_updated"], errors="coerce"
                ).dt.strftime("%Y-%m-%d").fillna("—")
            overview_cols["last_time_entry"] = "Last Time Entry"
            overview_cols["entry_source"]    = "Source"
        elif "last_updated" in df_drs.columns:
            df_drs["last_time_entry"] = pd.to_datetime(
                df_drs["last_updated"], errors="coerce"
            ).dt.strftime("%Y-%m-%d").fillna("—")
            df_drs["entry_source"]    = "SS DRS Modified"
            overview_cols["last_time_entry"] = "Last Time Entry"
            overview_cols["entry_source"]    = "Source"

        avail_cols = [c for c in overview_cols if c in df_drs.columns]
        overview_df = df_drs[avail_cols].rename(columns=overview_cols).copy()

        # Format start date
        if "Start Date" in overview_df.columns:
            overview_df["Start Date"] = pd.to_datetime(overview_df["Start Date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("—")

        # Add Suggested Tier column
        if "Days Inactive" in overview_df.columns:
            def _tier_label(days):
                try:
                    d = int(days)
                    if d < 0:   return "Unknown"
                    if d < 14:  return "—"
                    if d < 30:  return "Eligible for informal follow up"
                    if d < 60:  return "Tier 1"
                    if d < 90:  return "Tier 2"
                    if d < 180: return "Tier 3"
                    return "Tier 4"
                except: return "—"
            overview_df["Suggested Tier"] = overview_df["Days Inactive"].apply(_tier_label)
            overview_df["Days Inactive"]  = overview_df["Days Inactive"].apply(
                lambda x: "Unknown" if int(x) < 0 else int(x)
            )
            # Mark On Hold projects in table
            if "_on_hold" in df_drs.columns:
                on_hold_mask = df_drs["_on_hold"].values[:len(overview_df)]
                overview_df.loc[on_hold_mask, "Suggested Tier"] = "On Hold"

        # Colour-code rows by tier
        def _style_overview_row(row):
            days = row.get("Days Inactive", 0)
            try:
                d = int(days)
                if d >= 180: bg, fg = "#e8d5ff", "#4a235a"   # purple
                elif d >= 90:  bg, fg = "#FDECED", "#9C0006"  # red
                elif d >= 60:  bg, fg = "#FFEB9C", "#9C6500"  # amber
                elif d >= 30:  bg, fg = "#C6EFCE", "#276221"  # green
                else:          return [""] * len(row)
                return [f"background-color:{bg};color:{fg}"] * len(row)
            except:
                return [""] * len(row)

        sorted_df    = overview_df.sort_values("Days Inactive", ascending=False).copy()
        styled_overview = sorted_df.style.apply(_style_overview_row, axis=1)
        st.dataframe(styled_overview, hide_index=True, use_container_width=True)
        st.caption("🔔 Eligible for informal follow up = 14–29 days · On Hold projects shown for reference only")


    else:
        st.info("Upload SS DRS to see your project overview.")

    st.markdown("---")
    st.subheader("Step 3 — Select a Project")

    # ── Column resolution — works for both SFDC and DRS modes ─────────────
    if mode == "sfdc":
        opp_col  = "opportunity"      if "opportunity"      in df.columns else df.columns[0]
        acc_col  = "account"          if "account"          in df.columns else None
        prod_col = "product"          if "product"          in df.columns else None
        name_key = opp_col
    else:
        opp_col  = "project_name"     if "project_name"     in df.columns else df.columns[0]
        acc_col  = "account"          if "account"          in df.columns else None
        prod_col = "project_type"     if "project_type"     in df.columns else None
        name_key = opp_col

    # Build project list — concat Account : Project for clarity
    if acc_col and acc_col in df.columns:
        df["_proj_label"] = (
            df[acc_col].fillna("").astype(str).str.strip()
            + "  —  "
            + df[opp_col].fillna("").astype(str).str.strip()
        )
    else:
        df["_proj_label"] = df[opp_col].fillna("").astype(str).str.strip()

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
    st.subheader("Step 4 — Set Context")

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
        last_activity = st.date_input("Date of last activity", value=date.today(), key=f"last_act_{str(selected_proj)[:40].replace(chr(32),'_').replace(chr(39),'')}")

    col4, col5 = st.columns([3, 3])
    with col4:
        # Flip "Last, First" → "First Last" for email signature
        _parts = selected_user.split(",", 1)
        _display_name = f"{_parts[1].strip()} {_parts[0].strip()}" if len(_parts) == 2 else selected_user
        consultant_name = st.text_input("Your name (Implementation Consultant)", value=_display_name, key="ic_name")
    with col5:
        if int(days_inactive) >= 180:
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
    if int(days_inactive) >= 180:
        if close_dt and pd.notna(close_dt):
            expiry_default = (pd.Timestamp(close_dt) + pd.DateOffset(months=12)).date()
        else:
            expiry_default = date.today()
        service_expiry = st.date_input("Service term expiry date", value=expiry_default, key="svc_exp")
    else:
        service_expiry = (pd.Timestamp(close_dt) + pd.DateOffset(months=12)).date() if close_dt and pd.notna(close_dt) else date.today()

    # ── Contacts ──────────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Step 5 — Select Recipients")

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

    st.subheader("Step 6 — Template")

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
                                          index=suggested_idx, key="tmpl_select",
                                          label_visibility="visible")

    tmpl_info = TEMPLATES[selected_template]
    tier_num  = tmpl_info["tier"]

    # CC guidance
    cc_raw = tmpl_info["cc_guidance"].replace("{ACCOUNT MANAGER}", am if am else "{ACCOUNT MANAGER}")
    st.markdown(f"<div class='cc-box'>📋 <b>CC guidance:</b> {cc_raw}</div>", unsafe_allow_html=True)

    if tier_num in (3, 4) and not am:
        st.warning("Account Manager name not found in SFDC data — you'll need to fill {ACCOUNT MANAGER} manually.")

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
        "SERVICE TERM EXPIRY":    service_expiry.strftime("%B %d, %Y"),
    }

    subject, body, _ = fill_template(selected_template, fields)

    # ── Preview ───────────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Step 7 — Review & Send")

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
    btn_col1, = st.columns([1])
    with btn_col1:
        st.markdown(
            f"<a href='{_mailto}' target='_blank'>"
            f"<button style='background:#1e2c63;color:white;border:none;padding:10px 24px;"
            f"border-radius:6px;font-family:Manrope,sans-serif;font-size:14px;font-weight:600;"
            f"cursor:pointer;width:100%;margin-top:4px;'>✉️ Open in Email Client</button></a>",
            unsafe_allow_html=True
        )
    st.markdown("<div style='margin-top:6px'></div>", unsafe_allow_html=True)
    btn_col2, btn_col3 = st.columns([2, 2])
    with btn_col3:
        if st.button("📋 Log this outreach", key="log_btn", type="secondary"):
            _tier_label = TEMPLATES.get(selected_template, {}).get("tier", "")
            _tier_str   = f"Tier {_tier_label}" if _tier_label else selected_template
            _log_customer = str(account) if account and str(account).lower() not in ("nan","") else selected_proj
            _log_outreach(
                consultant    = consultant_name,
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
<button id="cb" style="background:#2980B9;color:white;border:none;padding:10px 24px;
border-radius:6px;font-family:Manrope,sans-serif;font-size:14px;font-weight:600;
cursor:pointer;">&#128196; Copy Formatted</button>
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
        _all_log = _load_log()
        if _all_log:
            _log_df = pd.DataFrame(_all_log).sort_values("date", ascending=False)
            # Filter to current user if selected
            if selected_user and selected_user != "— Select —":
                _log_df = _log_df[_log_df["consultant"] == selected_user]
            st.dataframe(
                _log_df[["date","consultant","customer","project","tier","days_inactive","template"]].rename(columns={
                    "date": "Date", "consultant": "Consultant", "customer": "Customer",
                    "project": "Project", "tier": "Tier",
                    "days_inactive": "Days Inactive", "template": "Template"
                }),
                hide_index=True, use_container_width=True
            )
            if st.button("🗑 Clear my log entries", key="clear_log"):
                _all    = _load_log()
                _kept   = [e for e in _all if e.get("consultant") != selected_user]
                _save_log(_kept)
                st.session_state["_log_success_msg"] = "✅ Your log entries cleared."
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
