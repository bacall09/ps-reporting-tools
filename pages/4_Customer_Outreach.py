import streamlit as st
import pandas as pd
from datetime import date, datetime
import re

st.set_page_config(page_title="Customer Outreach", page_icon=None, layout="wide")

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
    "project name":          "project_name",
    "project id":            "project_id",
    "project phase":         "phase",
    "project type":          "project_type",
    "status":                "status",
    "start date":            "start_date",
    "go live date":          "go_live_date",
    "territory":             "territory",
    "billing type":          "billing_type",
    "billing":               "billing_type",
    "project manager":       "project_manager",
    "consultant":            "project_manager",
    "overall rag":           "rag",
    "schedule health":       "schedule_health",
    "risk level":            "risk_level",
    "last updated":          "last_updated",
    "modified":              "last_updated",
    "account name":          "account",
    "customer":              "account",
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
    rename = {c: SS_COL_MAP_OUT[c.lower()] for c in df.columns if c.lower() in SS_COL_MAP_OUT}
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

    # Calculate days since start as proxy for inactivity (if no last_updated col)
    if "last_updated" in df.columns:
        df["days_inactive"] = (today - df["last_updated"]).dt.days.clip(lower=0)
    elif "start_date" in df.columns:
        df["days_inactive"] = (today - df["start_date"]).dt.days.clip(lower=0)
    else:
        df["days_inactive"] = 0

    return df


def suggest_tier_from_days(days):
    for name, tmpl in TEMPLATES.items():
        if tmpl["days_min"] <= int(days) <= tmpl["days_max"]:
            return name
    return list(TEMPLATES.keys())[-1]


def main():


    st.subheader("Step 1 — Upload Report")

    src_col1, src_col2 = st.columns([3, 3])
    with src_col1:
        data_source = st.radio(
            "Data source",
            ["SFDC Contacts Export", "SS DRS Export"],
            horizontal=True,
            key="data_source"
        )

    # ── SFDC path ──────────────────────────────────────────────────────────
    if data_source == "SFDC Contacts Export":
        st.caption("Required: First Name, Last Name, Email, Account Name, Opportunity Name · Optional: Close Date, Opportunity Owner, Implementation Contact, Contact Roles, Partner Contact")
        sfdc_file = st.file_uploader("Drop your SFDC contacts report here", type=["xlsx","xls","csv"], key="sfdc_outreach")
        if not sfdc_file:
            st.info("Upload your Salesforce contacts export — one row per contact per opportunity.")
            return
        try:
            df = load_sfdc(sfdc_file)
        except Exception as e:
            st.error(f"Could not load file: {e}")
            return
        mode = "sfdc"
        st.success(f"Loaded {len(df):,} rows · {df['account'].nunique() if 'account' in df.columns else '?'} accounts · {df['opportunity'].nunique() if 'opportunity' in df.columns else '?'} opportunities")

    # ── SS DRS path ────────────────────────────────────────────────────────
    else:
        st.caption("Required: Project Name, Project Phase, Project Type, Billing Type, Status · Optional: Project Manager, Account Name, Start Date, Last Updated")
        drs_file = st.file_uploader("Drop your SS DRS export here", type=["xlsx","xls","csv"], key="drs_outreach")
        if not drs_file:
            st.info("Upload your SmartSheets DRS export — the tool will surface active projects and calculate days open.")
            return
        try:
            df = load_drs(drs_file)
        except Exception as e:
            st.error(f"Could not load file: {e}")
            return
        mode = "drs"

        # ── Consultant filter for ICs / Managers ──────────────────────────
        if "project_manager" in df.columns:
            all_consultants = sorted(df["project_manager"].dropna().astype(str).unique())
            selected_consultant = st.selectbox(
                "Filter by Consultant (ICs: select your name · Managers: select any)",
                ["All"] + all_consultants,
                key="drs_consultant"
            )
            if selected_consultant != "All":
                df = df[df["project_manager"].astype(str) == selected_consultant]

        # Show stale project summary
        tier_counts = {}
        for name, tmpl in TEMPLATES.items():
            n = len(df[(df["days_inactive"] >= tmpl["days_min"]) & (df["days_inactive"] <= tmpl["days_max"])])
            if n > 0:
                tier_counts[name] = n

        if tier_counts:
            cols = st.columns(len(tier_counts))
            for i, (tier, n) in enumerate(tier_counts.items()):
                t_num = TEMPLATES[tier]["tier"]
                bg = [TIER_COLORS[t_num]]
                with cols[i]:
                    st.markdown(f"<div style='background:{TIER_COLORS[t_num]};color:{TIER_TEXT[t_num]};padding:8px 12px;border-radius:6px;font-size:13px;font-weight:700'>{tier}<br><span style='font-size:20px'>{n}</span> project(s)</div>", unsafe_allow_html=True)
        else:
            st.success("No projects meeting re-engagement thresholds (30+ days).")

        st.success(f"Loaded {len(df):,} active FF projects")

    st.markdown("---")
    st.subheader("Step 2 — Select a Project")

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
        proj_options = sorted(label_to_opp.keys(), key=lambda x: -days_map.get(label_to_opp.get(x,""), 0))
        proj_options_display = [_label_with_tier(p) for p in proj_options]
        display_to_label = dict(zip(proj_options_display, proj_options))
        selected_display = st.selectbox("Project — sorted by most days inactive", proj_options_display, key="proj_select")
        selected_label   = display_to_label[selected_display]
        selected_proj    = label_to_opp[selected_label]
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
    st.subheader("Step 3 — Set Inactivity & Context")

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
        consultant_name = st.text_input("Your name (Implementation Consultant)", placeholder="e.g. Jane Smith", key="ic_name")
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
    st.subheader("Step 4 — Select Recipients")

    if mode == "drs":
        st.info("SS DRS does not include customer contact details. To/CC fields are left blank — paste your contacts manually or switch to SFDC mode for contact lookup.")
        to_emails  = []
        cc_emails  = []
        primary_contact_first = ""
    
    if mode == "sfdc":
      email_col   = "email"             if "email"             in proj_rows.columns else None
    name_col    = "contact_name"      if "contact_name"      in proj_rows.columns else None
    roles_col   = "contact_roles"     if "contact_roles"     in proj_rows.columns else None
    flag_col    = "impl_contact_flag" if "impl_contact_flag" in proj_rows.columns else None
    partner_col = "partner_contact"   if "partner_contact"   in proj_rows.columns else None
    count_col   = "role_count"        if "role_count"        in proj_rows.columns else None

    # ── Contact role count warning ─────────────────────────────────────────
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

        # ── Priority logic for To: field ───────────────────────────────────
        # 1. impl_contact_flag = True
        # 2. contact_roles contains "Implementation Contact"
        # 3. contact_roles contains "Primary"
        # 4. First contact (fallback)
        to_source = "First contact (fallback)"
        suggested_to = all_emails[:1]

        if flag_col:
            flagged = proj_rows[proj_rows[flag_col] == True]
            if not flagged.empty:
                suggested_to = flagged[email_col].dropna().astype(str).tolist()
                to_source = "Implementation Contact checkbox"
        if to_source == "First contact (fallback)" and roles_col:
            impl_r = proj_rows[proj_rows[roles_col].str.contains("Implementation Contact", case=False, na=False)]
            if not impl_r.empty:
                suggested_to = impl_r[email_col].dropna().astype(str).tolist()
                to_source = "Contact Roles: Implementation Contact"
            else:
                prim_r = proj_rows[proj_rows[roles_col].str.contains("Primary", case=False, na=False)]
                if not prim_r.empty:
                    suggested_to = prim_r[email_col].dropna().astype(str).tolist()
                    to_source = "Contact Roles: Primary"

        # ── Partner contact → suggest as CC ────────────────────────────────
        partner_emails = []
        if partner_col:
            partner_vals = proj_rows[partner_col].dropna().astype(str)
            partner_emails = [v.strip() for v in partner_vals if v.strip() and v.strip().lower() not in ("nan","none","")]

        to_emails = st.multiselect("To:", all_emails, default=[e for e in suggested_to if e in all_emails], key="to_emails")
        st.caption(f"To: auto-selected via — {to_source}")

        suggested_cc = [e for e in partner_emails if e not in to_emails]
        cc_label = "CC:" + (" — Partner contact pre-suggested" if suggested_cc else "")
        cc_pool  = list(set(all_emails + partner_emails))
        cc_emails = st.multiselect(cc_label, cc_pool, default=suggested_cc, key="cc_emails")
        if suggested_cc:
            st.caption("CC: Partner contact pre-suggested — review before sending")

    else:
        st.warning("Email and/or contact name columns not detected. Check your export headers.")
        to_emails = []
        cc_emails = []
    # Primary contact first name — use first To: recipient if available (SFDC mode only)
    if mode == "sfdc":
     primary_contact_first = ""
     if name_col and not proj_rows.empty:
        if email_col and to_emails:
            to_row = proj_rows[proj_rows[email_col].astype(str) == to_emails[0]]
            full_name = str(to_row[name_col].iloc[0]) if not to_row.empty else str(proj_rows[name_col].iloc[0])
        else:
            full_name = str(proj_rows[name_col].iloc[0])
        primary_contact_first = full_name.split()[0] if full_name and full_name.lower() not in ("nan","") else ""

    # ── Template selection ────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Step 5 — Template")

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
    st.subheader("Step 6 — Review & Copy")

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
