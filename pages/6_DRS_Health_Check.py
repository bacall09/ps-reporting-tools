"""
PS Tools — DRS Health Check
Logical consistency validator for SS DRS data.
Flags fields and combinations that don't align with expected project state.
"""
import streamlit as st
import pandas as pd
from datetime import date

st.session_state["current_page"] = "DRS Health Check"

from shared.loaders import load_drs
from shared.constants import MILESTONE_COLS_MAP, name_matches




st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        h1,h2,h3,h4,p,div,label,button { font-family: 'Manrope', sans-serif !important; }
        .sev-error   { display:inline-block; font-size:11px; font-weight:700; padding:2px 8px;
                       border-radius:4px; background:rgba(192,57,43,0.15); color:#C0392B; margin-right:6px; }
        .sev-warning { display:inline-block; font-size:11px; font-weight:700; padding:2px 8px;
                       border-radius:4px; background:rgba(243,156,18,0.15); color:#D68910; margin-right:6px; }
        .sev-info    { display:inline-block; font-size:11px; font-weight:700; padding:2px 8px;
                       border-radius:4px; background:rgba(68,114,196,0.15); color:#4472C4; margin-right:6px; }
        .rule-row    { border:1px solid rgba(128,128,128,0.15); border-radius:6px;
                       padding:10px 14px; margin-bottom:6px; }
        .rule-title  { font-size:13px; font-weight:600; color:inherit; margin-bottom:2px; }
        .rule-desc   { font-size:12px; opacity:0.65; }
        .summary-card { border:1px solid rgba(128,128,128,0.2); border-radius:8px;
                        padding:14px 18px; text-align:center; }
        .summary-val  { font-size:28px; font-weight:700; }
        .summary-lbl  { font-size:12px; opacity:0.6; margin-top:2px; }
        .divider { border:none; border-top:1px solid rgba(128,128,128,0.15); margin:16px 0; }
    </style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
# Use passthrough key if set (navigated from My Projects), else fall back to home_browse
_b = (st.session_state.pop("_va_passthrough", None) or
      st.session_state.get("_browse_passthrough") or
      st.session_state.get("home_browse", "")) or ""
if _b.startswith("── ") and _b.endswith(" ──"):
    _drs_title_sfx = f" — {_b[3:-3].strip()} Team"
elif _b and _b not in ("— My own view —", "— Select —", "👥 All team"):
    _bp = [p.strip() for p in _b.split(",")]
    _drs_title_sfx = f" — {_bp[1] + ' ' + _bp[0] if len(_bp)==2 else _b}"
else:
    _drs_title_sfx = ""

_hero = st.empty()
_hero.markdown(f"<div style='background:#050D1F;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;font-family:Manrope,sans-serif;position:relative;overflow:hidden'> <div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3B9EFF;margin-bottom:10px;font-family:Manrope,sans-serif'>Professional Services · Reporting</div> <h1 style='color:#fff;margin:0;font-size:28px;font-weight:800;font-family:Manrope,sans-serif;line-height:1.15'>DRS Health Check{_drs_title_sfx}</h1> <p style='color:rgba(255,255,255,0.6);margin:8px 0 0;font-size:14px;font-family:Manrope,sans-serif;line-height:1.6;max-width:520px'>Logical consistency validator for Smartsheet DRS data — flags fields and combinations that don't align with expected project state.</p> </div>", unsafe_allow_html=True)

# ── Data source ───────────────────────────────────────────────────────────────
df_drs = st.session_state.get("df_drs")

if df_drs is None:
    st.info("Upload your SS DRS Export (or load it on the Home page) to run the health check.")
    st.stop()

# ── Apply view_as filter (respects home_browse for managers) ──────────────────
from shared.constants import get_role as _get_role, resolve_view_as, get_region_consultants
from shared.config import EMPLOYEE_LOCATION as _EL3, PS_REGION_MAP as _RM3, PS_REGION_OVERRIDE as _RO3
from shared.constants import ACTIVE_EMPLOYEES as _AE3, EMPLOYEE_ROLES as _ER3
_session_name = st.session_state.get("consultant_name", "")
if _session_name:
    _home_browse = (st.session_state.pop("_browse_passthrough", None) or
                    st.session_state.get("home_browse", "— My own view —"))
    _va_name, _va_region, _is_group = resolve_view_as(
        _session_name, _home_browse, _ER3, _EL3, _RM3, _RO3, _AE3
    )
    _role = _get_role(_session_name)
    _is_manager = _role in ("manager", "manager_only")
    if _va_region and "project_manager" in df_drs.columns:
        _rc = get_region_consultants(_va_region, _EL3, _RM3, _RO3, _AE3)
        _filtered = df_drs[df_drs["project_manager"].astype(str).str.strip().str.lower().isin(_rc)]
        if not _filtered.empty: df_drs = _filtered
    elif _va_name or not _is_manager:
        _target = _va_name if _va_name else _session_name
        if "project_manager" in df_drs.columns:
            _filtered = df_drs[df_drs["project_manager"].apply(lambda v: name_matches(v, _target))]
            if not _filtered.empty: df_drs = _filtered

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# RULE ENGINE
# ══════════════════════════════════════════════════════════════════════════════

today = pd.Timestamp.today().normalize()

# Helper — safe column getter that returns None if col missing
def _get(row, col, default=None):
    v = row.get(col, default)
    if v is None: return default
    if isinstance(v, float) and pd.isna(v): return default
    if str(v).strip().lower() in ("", "nan", "none", "nat"): return default
    return v

def _is_date(val):
    try:
        return pd.notna(pd.to_datetime(val))
    except Exception:
        return False

def _days_since(val):
    try:
        dt = pd.to_datetime(val)
        if pd.isna(dt): return None
        return int((today - dt).days)
    except Exception:
        return None

def _days_until(val):
    try:
        dt = pd.to_datetime(val)
        if pd.isna(dt): return None
        return int((dt - today).days)
    except Exception:
        return None

# Phase ordering — used to check if milestones match phase
PHASE_ORDER = [
    "00. onboarding",
    "01. requirements and design",
    "02. configuration",
    "03. enablement/training",
    "04. uat",
    "05. prep for go-live",
    "06. go-live (hypercare)",
    "07. hypercare",
    "08. ready for support transition",
    "09. phase 2 scoping",
]

# Milestones in delivery order — used to check sequence
MS_ORDER = [
    "ms_intro_email", "ms_config_start", "ms_enablement",
    "ms_session1", "ms_session2", "ms_uat_signoff",
    "ms_prod_cutover", "ms_hypercare_start", "ms_close_out", "ms_transition",
]

# Which phase each milestone should be completed by
MS_EXPECTED_BY_PHASE = {
    "ms_intro_email":     "00. onboarding",
    "ms_config_start":    "02. configuration",
    "ms_enablement":      "03. enablement/training",
    "ms_session1":        "03. enablement/training",
    "ms_session2":        "03. enablement/training",
    "ms_uat_signoff":     "04. uat",
    "ms_prod_cutover":    "06. go-live (hypercare)",
    "ms_hypercare_start": "06. go-live (hypercare)",
    "ms_close_out":       "08. ready for support transition",
    "ms_transition":      "08. ready for support transition",
}

def _phase_idx(phase_str):
    p = str(phase_str).strip().lower()
    for i, ph in enumerate(PHASE_ORDER):
        if p.startswith(ph[:6]) or ph in p or p in ph:
            return i
    return -1

# ── Run rules ─────────────────────────────────────────────────────────────────
findings = []  # list of dicts: {project, severity, category, rule, description, expected}

for _, row in df_drs.iterrows():
    proj   = _get(row, "project_name") or _get(row, "project_id") or "Unknown project"
    phase  = _get(row, "phase", "")
    status = str(_get(row, "status", "") or "").strip().lower()
    rag    = str(_get(row, "rag",    "") or "").strip().upper()[:1]
    resp   = str(_get(row, "client_responsiveness", "") or "").strip().lower()
    pm     = _get(row, "project_manager")
    go_live   = _get(row, "effective_go_live_date") or _get(row, "go_live_date")
    start_dt  = _get(row, "start_date")
    actual_h  = _get(row, "actual_hours")
    budget_h  = _get(row, "budgeted_hours")
    change_ord= _get(row, "change_order")
    days_inac = _get(row, "days_inactive")
    phase_idx = _phase_idx(phase) if phase else -1

    # Legacy projects had no milestone tracking — skip all milestone checks
    _legacy_raw = _get(row, "legacy", "")
    # Handle both string and boolean forms (loader normalises to bool)
    is_legacy = (
        _legacy_raw is True or
        str(_legacy_raw).strip().lower() in ("yes", "true", "1", "y")
    )

    def flag(severity, category, rule, description, expected=""):
        findings.append({
            "project":     proj,
            "severity":    severity,
            "category":    category,
            "rule":        rule,
            "description": description,
            "expected":    expected,
        })

    # ── Completeness ──────────────────────────────────────────────────────────
    if not pm:
        flag("Warning", "Completeness",
             "No Project Manager assigned",
             "Project Manager / Consultant field is blank.",
             "Every active project should have a named consultant.")

    if not phase:
        flag("Error", "Completeness",
             "Phase is blank",
             "Project phase is missing — cannot assess progress or flag milestones.",
             "Set the current phase in SS DRS.")

    if not go_live and phase_idx >= 0 and phase_idx < 7:
        flag("Warning", "Completeness",
             "No Go Live date set",
             "Project has no Go Live date and is not yet in hypercare or later.",
             "Set a target Go Live date so timelines and inactivity signals work correctly.")

    # ── Date logic ────────────────────────────────────────────────────────────
    if _is_date(start_dt) and _is_date(go_live):
        if pd.to_datetime(go_live) < pd.to_datetime(start_dt):
            flag("Error", "Date Logic",
                 "Go Live date is before Start date",
                 f"Start: {pd.to_datetime(start_dt).strftime('%d %b %Y')} · "
                 f"Go Live: {pd.to_datetime(go_live).strftime('%d %b %Y')}",
                 "Go Live date must be after Start date.")

    if _is_date(start_dt):
        days_open = _days_since(start_dt)
        if days_open is not None and days_open > 180 and not _is_date(go_live):
            flag("Warning", "Date Logic",
                 "Open > 180 days with no Go Live date",
                 f"Project has been open {days_open} days with no Go Live date set.",
                 "Set a Go Live date or close the project if complete.")

    if _is_date(go_live):
        days_to_gl = _days_until(go_live)
        if days_to_gl is not None and days_to_gl < 0:
            # Go live in the past
            late_phases = {"06. go-live (hypercare)", "07. hypercare",
                           "08. ready for support transition", "09. phase 2 scoping"}
            if phase_idx >= 0 and not any(
                str(phase).strip().lower().startswith(lp[:6]) for lp in late_phases
            ):
                flag("Error", "Date Logic",
                     "Go Live date passed but phase not updated",
                     f"Go Live was {abs(days_to_gl)}d ago but phase is still '{phase}'.",
                     "Advance phase to Go-Live / Hypercare or update the Go Live date.")

        if days_to_gl is not None and 0 <= days_to_gl <= 14:
            early_phases = [p for p in PHASE_ORDER if p <= "04. uat"]
            if phase_idx >= 0 and any(
                str(phase).strip().lower().startswith(ep[:6]) for ep in early_phases
            ):
                flag("Warning", "Date Logic",
                     "Go Live within 14 days but phase is UAT or earlier",
                     f"Go Live is in {days_to_gl}d but phase is '{phase}'.",
                     "Confirm go-live is still on track or update the date.")

    # ── Status vs RAG ─────────────────────────────────────────────────────────
    if status in ("on track", "green") and rag == "R":
        flag("Error", "Status Conflict",
             "Status is On Track but RAG is Red",
             f"Status shows '{_get(row,'status','')}' but Overall RAG is Red.",
             "Align status and RAG — one of them needs to be updated.")

    if status in ("on track", "green") and rag == "A":
        flag("Warning", "Status Conflict",
             "Status is On Track but RAG is Amber",
             f"Status shows '{_get(row,'status','')}' but Overall RAG is Amber.",
             "Consider whether status should reflect the amber RAG signal.")

    # ── Activity vs status ────────────────────────────────────────────────────
    if "hold" in status and days_inac is not None and days_inac >= 0 and days_inac < 14:
        flag("Warning", "Activity Conflict",
             "On Hold but recently active",
             f"Status is On Hold but project has NS time entries within the last {days_inac}d.",
             "If work has resumed, update the status from On Hold.")

    if resp in ("responsive", "highly responsive") and days_inac is not None and days_inac > 30:
        flag("Warning", "Activity Conflict",
             "Client marked Responsive but project is stale",
             f"Client Responsiveness is '{_get(row,'client_responsiveness','')}' "
             f"but project has been inactive {days_inac}d.",
             "Update Client Responsiveness to reflect actual recent engagement.")

    # ── On Hold data quality checks ───────────────────────────────────────────
    sentiment  = str(_get(row, "client_sentiment",       "") or "").strip().lower()
    oh_reason  = str(_get(row, "on_hold_reason",         "") or "").strip()
    oh_delay   = str(_get(row, "responsible_for_delay",  "") or "").strip()

    if "hold" in status:
        if not oh_reason or oh_reason in ("—", "nan", "None"):
            flag("Warning", "On Hold Data Quality",
                 "On Hold Reason not set",
                 "Project is On Hold but no On Hold Reason has been recorded.",
                 "Set On Hold Reason in the DRS — required for all on-hold projects.")
        if not oh_delay or oh_delay in ("—", "nan", "None"):
            flag("Warning", "On Hold Data Quality",
                 "Responsible for Delay not set",
                 "Project is On Hold but Responsible for Delay has not been recorded.",
                 "Set Responsible for Delay in the DRS — required for all on-hold projects.")

    if "hold" in status and days_inac is not None and days_inac >= 14:
        if resp in ("highly engaged", "highly responsive", "responsive"):
            flag("Warning", "On Hold Data Quality",
                 "Engagement rating inconsistent with On Hold status",
                 f"Client Responsiveness is '{_get(row,'client_responsiveness','')}' "
                 f"but project has been On Hold for {days_inac}d. "
                 f"This rating should reflect current engagement, not historical.",
                 "Review and update Client Responsiveness — consider 'Neutral' or 'Not Responsive'.")

        if sentiment in ("positive",):
            flag("Warning", "On Hold Data Quality",
                 "Sentiment rating inconsistent with On Hold status",
                 f"Client Sentiment is '{_get(row,'client_sentiment','')}' "
                 f"but project has been On Hold for {days_inac}d. "
                 f"Positive sentiment is unlikely for a stalled project.",
                 "Review and update Client Sentiment to reflect current client relationship.")

    # ── Hours vs scope ────────────────────────────────────────────────────────
    if actual_h and budget_h:
        try:
            a = float(actual_h); b = float(budget_h)
            if a > b:
                overage = round(a - b, 2)
                co = str(change_ord or "").strip().lower()
                if not co or co in ("no", "false", "0", "none", "nan", ""):
                    flag("Warning", "Hours vs Scope",
                         "Actual hours exceed budget — no Change Order flagged",
                         f"Actual: {a}h · Budget: {b}h · Overage: {overage}h. "
                         f"No Change Order recorded.",
                         "Log a Change Order or review budget allocation.")
        except Exception:
            pass

    if _is_date(start_dt):
        days_open = _days_since(start_dt)
        if days_open is not None and days_open > 30:
            try:
                a = float(actual_h or 0)
                if a == 0:
                    flag("Info", "Hours vs Scope",
                         "No hours logged after 30+ days",
                         f"Project started {days_open}d ago but has 0 actual hours recorded.",
                         "Confirm project is active and NS time is being logged.")
            except Exception:
                pass

    # ── Milestone sequence ────────────────────────────────────────────────────
    # Legacy projects had no milestone tracking — skip all milestone checks
    if is_legacy:
        pass  # milestone checks exempt for legacy projects
    else:
        ms_dates = {}
        for ms_col in MS_ORDER:
            if ms_col in row.index:
                v = row.get(ms_col)
                if _is_date(v):
                    ms_dates[ms_col] = pd.to_datetime(v)

        # Check date ordering — each milestone should be >= the previous one
        prev_col, prev_date = None, None
        for ms_col in MS_ORDER:
            if ms_col not in ms_dates:
                prev_col, prev_date = None, None
                continue
            dt = ms_dates[ms_col]
            if prev_date is not None and dt < prev_date:
                flag("Error", "Milestone Sequence",
                     f"Milestone out of sequence: {MILESTONE_COLS_MAP.get(ms_col, ms_col)}",
                     f"'{MILESTONE_COLS_MAP.get(ms_col, ms_col)}' ({dt.strftime('%d %b %Y')}) "
                     f"is dated before '{MILESTONE_COLS_MAP.get(prev_col, prev_col)}' "
                     f"({prev_date.strftime('%d %b %Y')}).",
                     "Milestone dates should follow delivery order.")
            prev_col, prev_date = ms_col, dt

        # ── Phase vs milestones ───────────────────────────────────────────────
        # Flag milestones completed that are ahead of current phase
        if phase_idx >= 0:
            for ms_col, expected_phase in MS_EXPECTED_BY_PHASE.items():
                exp_idx = _phase_idx(expected_phase)
                if ms_col in ms_dates and exp_idx > phase_idx + 1:
                    flag("Warning", "Phase vs Milestone",
                         f"Milestone ahead of current phase: {MILESTONE_COLS_MAP.get(ms_col, ms_col)}",
                         f"'{MILESTONE_COLS_MAP.get(ms_col, ms_col)}' is completed but "
                         f"current phase is '{phase}' — this milestone is expected in a later phase.",
                         "Check whether phase needs to be advanced.")

        # Flag milestones that should be done by now but are missing
        if phase_idx >= 0:
            for ms_col, expected_phase in MS_EXPECTED_BY_PHASE.items():
                exp_idx = _phase_idx(expected_phase)
                if exp_idx >= 0 and phase_idx > exp_idx and ms_col not in ms_dates:
                    # Only flag the most critical ones to avoid noise
                    critical = {"ms_intro_email", "ms_uat_signoff", "ms_prod_cutover"}
                    if ms_col in critical:
                        flag("Warning", "Phase vs Milestone",
                             f"Expected milestone missing: {MILESTONE_COLS_MAP.get(ms_col, ms_col)}",
                             f"Phase is '{phase}' but '{MILESTONE_COLS_MAP.get(ms_col, ms_col)}' "
                             f"has no completion date recorded.",
                             "Complete or back-date the milestone if it has been done.")

# ══════════════════════════════════════════════════════════════════════════════
# RESULTS
# ══════════════════════════════════════════════════════════════════════════════
df_findings = pd.DataFrame(findings)

if df_findings.empty:
    st.success(f"✓ No issues found across {len(df_drs):,} projects — DRS data looks clean.")
    st.stop()

n_error   = len(df_findings[df_findings["severity"] == "Error"])
n_warning = len(df_findings[df_findings["severity"] == "Warning"])
n_info    = len(df_findings[df_findings["severity"] == "Info"])
n_projects = df_findings["project"].nunique()

# ── Summary cards ─────────────────────────────────────────────────────────────
c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    st.markdown(f'<div class="summary-card"><div class="summary-val">{len(df_drs):,}</div><div class="summary-lbl">Projects checked</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="summary-card"><div class="summary-val">{n_projects}</div><div class="summary-lbl">Projects with issues</div></div>', unsafe_allow_html=True)
with c3:
    col = "#C0392B" if n_error > 0 else "inherit"
    st.markdown(f'<div class="summary-card"><div class="summary-val" style="color:{col}">{n_error}</div><div class="summary-lbl">Errors</div></div>', unsafe_allow_html=True)
with c4:
    col = "#D68910" if n_warning > 0 else "inherit"
    st.markdown(f'<div class="summary-card"><div class="summary-val" style="color:{col}">{n_warning}</div><div class="summary-lbl">Warnings</div></div>', unsafe_allow_html=True)
with c5:
    st.markdown(f'<div class="summary-card"><div class="summary-val">{n_info}</div><div class="summary-lbl">Info</div></div>', unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── Filters ───────────────────────────────────────────────────────────────────
fc1, fc2, fc3 = st.columns([2, 2, 3])
with fc1:
    sev_filter = st.multiselect(
        "Severity",
        options=["Error", "Warning", "Info"],
        default=["Error", "Warning"],
        key="hc_sev",
    )
with fc2:
    cat_options = sorted(df_findings["category"].unique())
    cat_filter = st.multiselect(
        "Category",
        options=cat_options,
        default=cat_options,
        key="hc_cat",
    )
with fc3:
    proj_options = ["All projects"] + sorted(df_findings["project"].unique())
    proj_filter = st.selectbox("Project", options=proj_options, key="hc_proj")

filtered = df_findings[
    df_findings["severity"].isin(sev_filter) &
    df_findings["category"].isin(cat_filter)
]
if proj_filter != "All projects":
    filtered = filtered[filtered["project"] == proj_filter]

st.markdown(f"**{len(filtered)} issue(s)** shown")
st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── Results grouped by project ────────────────────────────────────────────────
SEV_ORDER = {"Error": 0, "Warning": 1, "Info": 2}
filtered = filtered.copy()
filtered["_sev_ord"] = filtered["severity"].map(SEV_ORDER)
filtered = filtered.sort_values(["_sev_ord", "project"])

for proj_name, proj_group in filtered.groupby("project", sort=False):
    n_e = len(proj_group[proj_group["severity"] == "Error"])
    n_w = len(proj_group[proj_group["severity"] == "Warning"])

    # Expander label must be plain text — Streamlit doesn't render HTML in titles
    summary_text = ""
    if n_e: summary_text += f"  ·  {n_e} Error{'s' if n_e>1 else ''}"
    if n_w: summary_text += f"  ·  {n_w} Warning{'s' if n_w>1 else ''}"

    with st.expander(f"{proj_name}{summary_text}", expanded=(n_e > 0)):
        # Show badge summary at top of expander body
        badge_html = ""
        if n_e: badge_html += f'<span class="sev-error">{n_e} Error{"s" if n_e>1 else ""}</span> '
        if n_w: badge_html += f'<span class="sev-warning">{n_w} Warning{"s" if n_w>1 else ""}</span>'
        if badge_html:
            st.markdown(f'<div style="margin-bottom:10px">{badge_html}</div>', unsafe_allow_html=True)

        for _, finding in proj_group.iterrows():
            sev   = finding["severity"]
            sev_cls = {"Error": "sev-error", "Warning": "sev-warning", "Info": "sev-info"}.get(sev, "sev-info")
            st.markdown(f"""
            <div class="rule-row">
                <div class="rule-title">
                    <span class="{sev_cls}">{sev}</span>
                    {finding['category']} · {finding['rule']}
                </div>
                <div class="rule-desc">{finding['description']}</div>
                {"<div class='rule-desc' style='margin-top:4px;color:#4472C4'>→ " + finding['expected'] + "</div>" if finding['expected'] else ""}
            </div>
            """, unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── Export ────────────────────────────────────────────────────────────────────
_pm_cols = ["project_manager"] if "project_manager" in filtered.columns else []
_export_cols = ["project"] + _pm_cols + ["severity","category","rule","description","expected"]
export = filtered[[c for c in _export_cols if c in filtered.columns]].rename(columns={
    "project": "Project", "project_manager": "Consultant", "severity": "Severity", "category": "Category",
    "rule": "Rule", "description": "Description", "expected": "Expected State",
})
st.download_button(
    label="⬇ Download findings as CSV",
    data=export.to_csv(index=False),
    file_name=f"drs_health_check_{date.today().strftime('%Y%m%d')}.csv",
    mime="text/csv",
)
