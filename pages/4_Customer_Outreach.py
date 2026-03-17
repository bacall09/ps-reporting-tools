import streamlit as st
import pandas as pd
from datetime import date, datetime
import re

st.set_page_config(page_title="Customer Outreach", page_icon=None, layout="wide")

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
            background: #f8faff;
            border: 1px solid #d0dff5;
            border-radius: 8px;
            padding: 20px 24px;
            font-family: Arial, sans-serif;
            font-size: 14px;
            line-height: 1.7;
            white-space: pre-wrap;
            color: #1a1a1a;
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
            margin-bottom: 8px;
        }
    </style>
""", unsafe_allow_html=True)

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
    <div style='background-color:#1e2c63;padding:24px 32px;border-radius:8px;margin-bottom:24px;font-family:Manrope,sans-serif'>
        <h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif'>Customer Outreach</h1>
        <p style='color:#aac4d0;margin:6px 0 0 0;font-size:14px;font-family:Manrope,sans-serif'>Re-engagement communications for unresponsive customers · Auto-suggests tier based on inactivity</p>
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
    "contact roles":            "contact_roles",
    "opp contact role count":   "role_count",
    "partner contact":          "partner_contact",
}

def load_sfdc(file):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
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
    if "status" in df.columns:
        df = df[~df["status"].str.strip().str.lower().isin({"on-hold","on hold","onhold","on_hold"})]

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


def suggest_tier_from_days(days):
    for name, tmpl in TEMPLATES.items():
        if tmpl["days_min"] <= int(days) <= tmpl["days_max"]:
            return name
    return list(TEMPLATES.keys())[-1]


NS_COL_MAP_OUT = {
    "employee":             "employee",
    "name":                 "employee",
    "project":              "project",
    "project name":         "project",
    "project id":           "project_id",
    "internal id":          "project_id",
    "id":                   "project_id",
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
    rename = {col: NS_COL_MAP_OUT[col.lower()] for col in df.columns if col.lower() in NS_COL_MAP_OUT}
    df = df.rename(columns=rename)
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    if "project_id" in df.columns:
        df["project_id"] = df["project_id"].apply(lambda x: str(int(float(str(x)))) if str(x) not in ("nan","None","") else "")
    # Store original cols for debug
    df.attrs["original_cols"] = list(df.columns)
    return df


def calc_days_inactive(df_drs, df_ns):
    """Merge NS last time entry onto DRS projects to get objective inactivity days."""
    today = pd.Timestamp.today().normalize()
    if df_ns is None or "date" not in df_ns.columns:
        return df_drs

    # Try join on project_id first, fall back to project name
    if "project_id" in df_ns.columns and "project_id" in df_drs.columns:
        df_ns["project_id"] = df_ns["project_id"].astype(str).str.strip().str.split(".").str[0]
        df_drs["project_id"] = df_drs["project_id"].astype(str).str.strip().str.split(".").str[0]
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
        # Override days_inactive with NS-based calculation where available
        ns_days = (today - df_drs["last_ns_entry"]).dt.days.clip(lower=0)
        df_drs["days_inactive"] = ns_days.where(
            df_drs["last_ns_entry"].notna(),
            df_drs["days_inactive"]  # keep DRS modified date value as fallback
        ).fillna(-1).astype(int)
        df_drs["_inactivity_source"] = df_drs["last_ns_entry"].apply(
            lambda x: "NS Time Entry" if pd.notna(x) else df_drs["_inactivity_source"].iloc[0] if "_inactivity_source" in df_drs.columns else "Unknown"
        )
    return df_drs


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
        st.caption("Required: Project Name, Project Phase, Project Type, Billing Type, Status")
        drs_file = st.file_uploader("Drop SS DRS file here", type=["xlsx","xls","csv"], key="drs_outreach")

    with up3:
        st.markdown("**NS Time Detail** — objective inactivity signal")
        st.caption("Same export used on the Utilization Report page")
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
            # Show what mapped — helps debug missing columns
            _mapped   = [col for col in ["last_updated","client_responsiveness","project_manager","project_id"] if col in df_drs.columns]
            _unmapped = [col for col in ["last_updated","client_responsiveness","project_manager","project_id"] if col not in df_drs.columns]
            if _unmapped:
                st.warning(f"DRS: These expected columns did not map: {_unmapped} — check column names in your export match exactly.")
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
        except Exception as e:
            st.error(f"Could not load NS file: {e}")

    # ── Merge NS inactivity into DRS ───────────────────────────────────────
    if df_drs is not None and df_ns is not None:
        df_drs = calc_days_inactive(df_drs, df_ns)

    # ── Summary metrics ───────────────────────────────────────────────────
    msg_parts = []
    if df_sfdc is not None:
        msg_parts.append(f"SFDC: {df_sfdc['account'].nunique() if 'account' in df_sfdc.columns else '?'} accounts · {df_sfdc['opportunity'].nunique() if 'opportunity' in df_sfdc.columns else '?'} opportunities")
    if df_drs is not None:
        msg_parts.append(f"DRS: {len(df_drs):,} active FF projects")
        # Tier summary from DRS
        tier_counts = {name: len(df_drs[(df_drs["days_inactive"] >= t["days_min"]) & (df_drs["days_inactive"] <= t["days_max"])]) for name, t in TEMPLATES.items()}
        tier_counts = {k: v for k, v in tier_counts.items() if v > 0}
        if tier_counts:
            cols = st.columns(len(tier_counts))
            for i, (tier, n) in enumerate(tier_counts.items()):
                t_num = TEMPLATES[tier]["tier"]
                with cols[i]:
                    st.markdown(f"<div style='background:{TIER_COLORS[t_num]};color:{TIER_TEXT[t_num]};padding:8px 12px;border-radius:6px;font-size:13px;font-weight:700'>{tier}<br><span style='font-size:20px'>{n}</span> project(s)</div>", unsafe_allow_html=True)
    if msg_parts:
        st.success(" · ".join(msg_parts))

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
        overview_cols = {
            "account":                "Customer",
            "project_name":           "Project",
            "project_type":           "Project Type",
            "start_date":             "Start Date",
            "status":                 "Status",
            "phase":                  "Current Phase",
            "client_responsiveness":  "Client Responsiveness",
            "days_inactive":          "Days Inactive",
            "rag":                    "RAG",
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
                    if d < 0: return "Unknown"
                    for name, t in TEMPLATES.items():
                        if t["days_min"] <= d <= t["days_max"]:
                            return f"Tier {t['tier']}"
                    return "< 30 days"
                except: return "—"
            overview_df["Suggested Tier"] = overview_df["Days Inactive"].apply(_tier_label)
            overview_df["Days Inactive"]  = overview_df["Days Inactive"].apply(
                lambda x: "Unknown" if int(x) < 0 else int(x)
            )

        # Colour-code rows by tier
        def _style_overview(row):
            days = row.get("Days Inactive", 0)
            try:
                d = int(days)
                if d >= 180: return ["background-color:#f0e6ff"] * len(row)
                if d >= 90:  return ["background-color:#FDECED"] * len(row)
                if d >= 60:  return ["background-color:#FFEB9C"] * len(row)
                if d >= 30:  return ["background-color:#EAF9F1"] * len(row)
            except: pass
            return [""] * len(row)

        styled_overview = overview_df.sort_values("Days Inactive", ascending=False).style.apply(_style_overview, axis=1)
        st.dataframe(styled_overview, hide_index=True, use_container_width=True)
        st.caption("🟢 30d+ · 🟡 60d+ · 🔴 90d+ · 🟣 180d+ · Dropdown below shows only projects ≥ 14 days inactive")

        # Debug expander — shows what DRS columns were detected
        with st.expander("🔍 DRS column detection — click if Days Inactive shows Unknown", expanded=False):
            detected = [col for col in df_drs.columns if not col.startswith("_")]
            st.caption("Columns detected in your DRS export:")
            st.write(detected)
            st.caption("For inactivity to work from DRS, the export needs a column named: `Modified`, `Last Updated`, or `Modified Date`. If yours is named differently let us know and we'll add it to the map.")
            if df_ns is not None:
                st.success("✅ NS Time Detail uploaded — inactivity will be calculated from last time entry.")
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
        proj_options_filtered = [p for p in proj_options_all if days_map.get(label_to_opp.get(p,""), 0) >= 14
                                  or days_map.get(label_to_opp.get(p,""), 0) == -1]

        if not proj_options_filtered:
            st.info("No projects with 14+ days inactivity found. All projects are up to date.")
            return

        def _label_with_tier_clean(label):
            opp   = label_to_opp.get(label, "")
            days  = days_map.get(opp, 0)
            suggested = suggest_tier_from_days(days)
            t_num = TEMPLATES[suggested]["tier"]
            days_str = "Unknown inactivity" if int(days) < 0 else f"{int(days)} days inactive"
            return f"Tier {t_num}  ·  {days_str}  —  {label}"

        proj_options_display = [_label_with_tier_clean(p) for p in proj_options_filtered]
        display_to_label     = dict(zip(proj_options_display, proj_options_filtered))
        selected_display     = st.selectbox(
            f"Select project ({len(proj_options_filtered)} with 14+ days inactive)",
            proj_options_display,
            key="proj_select"
        )
        selected_label = display_to_label[selected_display]
        selected_proj  = label_to_opp[selected_label]
    else:
        proj_options  = sorted(label_to_opp.keys())
        selected_label = st.selectbox("Project / Opportunity", proj_options, key="proj_select")
        selected_proj  = label_to_opp[selected_label]

    proj_rows = df[df[opp_col].astype(str) == selected_proj]

    # Project metadata
    account  = proj_rows[acc_col].iloc[0]  if acc_col and acc_col in proj_rows.columns and not proj_rows.empty else ""
    product  = proj_rows[prod_col].iloc[0] if prod_col and prod_col in proj_rows.columns and not proj_rows.empty else ""
    am_col   = "account_manager"           if "account_manager" in df.columns else None
    am       = proj_rows[am_col].iloc[0]   if am_col and not proj_rows.empty else ""
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
    _sfdc_product = str(product).strip() if product and str(product).lower() not in ("nan", "") else ""

    col0, = st.columns([1])
    with col0:
        product_name = st.text_input(
            "Product(s) being implemented",
            value=_sfdc_product,
            placeholder="e.g. ZoneCapture, ZoneApprovals",
            key="product_name"
        )

    col1, col2, col3 = st.columns([2, 2, 2])
    with col1:
        days_inactive = st.number_input("Days since last customer contact", min_value=0, value=_drs_days, step=1, key="days_in")
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
            key="phase_in"
        )
    with col3:
        last_activity = st.date_input("Date of last activity", value=date.today(), key="last_act")

    col4, col5 = st.columns([3, 3])
    with col4:
        # Flip "Last, First" → "First Last" for email signature
        _parts = selected_user.split(",", 1)
        _display_name = f"{_parts[1].strip()} {_parts[0].strip()}" if len(_parts) == 2 else selected_user
        consultant_name = st.text_input("Your name (Implementation Consultant)", value=_display_name, key="ic_name")
    with col5:
        if int(days_inactive) >= 180:
            remaining_sessions = st.text_input("Remaining sessions", placeholder="e.g. 3", key="rem_sess")
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

    if mode == "drs":
        st.info("SS DRS does not include customer contact details. To/CC fields are left blank — add contacts manually or use the mailto link after composing.")
        to_emails             = []
        cc_emails             = []
        primary_contact_first = ""

    elif mode == "sfdc":
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
            display_cols = [col for col in [name_col, email_col, roles_col, "title", flag_col] if col and col in proj_rows.columns]
            contacts_display = proj_rows[display_cols].drop_duplicates().reset_index(drop=True)
            col_rename = {
                "contact_name": "Name", "email": "Email", "contact_roles": "Contact Roles",
                "title": "Title", "impl_contact_flag": "Impl. Contact ✓"
            }
            contacts_display = contacts_display.rename(columns={k: v for k, v in col_rename.items() if k in contacts_display.columns})
            st.dataframe(contacts_display, hide_index=True, use_container_width=True)

            all_emails = proj_rows[email_col].dropna().astype(str).tolist()

            # Priority: 1. impl_contact_flag  2. Implementation Contact role  3. Primary role  4. First contact
            to_source    = "First contact (fallback)"
            suggested_to = all_emails[:1]
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
                else:
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

            suggested_cc = [e for e in partner_emails if e not in to_emails]
            cc_pool  = list(set(all_emails + partner_emails))
            cc_label = "CC:" + (" — Partner contact pre-suggested" if suggested_cc else "")
            cc_emails = st.multiselect(cc_label, cc_pool, default=suggested_cc, key="cc_emails")
            if suggested_cc:
                st.caption("CC: Partner contact pre-suggested — review before sending")
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
    st.subheader("Step 6 — Template")

    suggested = suggest_tier(int(days_inactive))
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
    st.text_area(
        "Plain text (select all → copy):",
        value=f"Subject: {subject}\n\n{body}",
        height=300,
        key="copy_area",
        label_visibility="visible"
    )

    # ── mailto link ───────────────────────────────────────────────────────────
    import urllib.parse
    _to_str  = ",".join(to_emails)  if to_emails  else ""
    _cc_str  = ",".join(cc_emails)  if cc_emails  else ""
    _subject = urllib.parse.quote(subject)
    _body    = urllib.parse.quote(body)
    _mailto  = f"mailto:{_to_str}?subject={_subject}&body={_body}"
    if _cc_str:
        _mailto = f"mailto:{_to_str}?cc={urllib.parse.quote(_cc_str)}&subject={_subject}&body={_body}"

    st.markdown(
        f"<a href='{_mailto}' target='_blank'>"
        f"<button style='background:#1e2c63;color:white;border:none;padding:10px 24px;"
        f"border-radius:6px;font-family:Manrope,sans-serif;font-size:14px;font-weight:600;"
        f"cursor:pointer;margin-top:8px;'>✉️ Open in Email Client</button></a>",
        unsafe_allow_html=True
    )
    st.caption("Opens your default email client (Gmail, Outlook, etc.) with subject and body pre-filled · Review placeholders before sending")

    # ── Merge field reference ─────────────────────────────────────────────────
    with st.expander("📋 Merge field reference — what was filled vs. what still needs updating", expanded=False):
        ref_rows = []
        for k, v in fields.items():
            status = "✅ Filled" if v else "⚠️ Needs manual update"
            ref_rows.append({"Placeholder": f"{{{k}}}", "Value Used": v or "—", "Status": status})
        st.dataframe(pd.DataFrame(ref_rows), hide_index=True, use_container_width=True)

    # ── Tier guidance ─────────────────────────────────────────────────────────
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
