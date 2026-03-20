"""
PS Tools — Shared Loaders & Calculation Helpers
Data loading, fuzzy matching, inactivity calc, template utilities.
"""
import re
import pandas as pd
import streamlit as st
from datetime import date, datetime
from rapidfuzz import fuzz

from shared.constants import PRODUCT_KEYWORDS, MILESTONE_COLS_MAP, PHASE_BENCHMARKS, SS_COL_MAP
from shared.config import DEFAULT_SCOPE

# Alias so load_drs can reference it — constants uses SS_COL_MAP, loaders used _OUT suffix
SS_COL_MAP_OUT = SS_COL_MAP

# Phases considered closed/inactive — excluded from DRS active project list
INACTIVE_PHASES_OUT = {
    "closed", "complete", "completed", "cancelled", "canceled",
    "transitioned to support", "transition to support", "closed - won",
    "closed - lost", "10. closed", "11. cancelled",
}


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


