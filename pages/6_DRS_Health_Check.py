"""
PS Tools — DRS Health Check
Logical consistency validator for SS DRS data.
Flags fields and combinations that don't align with expected project state.
"""
import streamlit as st

def _get_effective_browse(session_name, role, el, rm, ro, ae, cd, sidebar_key):
    """Return (browse_str, is_all_team) robust to session wipe and emoji encoding."""
    _b = (st.session_state.get("_browse_passthrough") or
          st.session_state.get("home_browse", "")) or ""
    _is_mgr = role in ("manager", "manager_only", "reporting_only")
    if _is_mgr and not _b:
        # No passthrough — show local selector defaulting to All team
        _active = sorted([n for n in cd if n])
        _bopts = ["👥 All team"]
        _by_r = {}
        for _cn in _active:
            _loc = el.get(_cn, "")
            _rg = ro.get(_cn, rm.get(_loc, "Other"))
            _by_r.setdefault(_rg, []).append(_cn)
        for _rg in sorted(_by_r.keys()):
            _bopts.append(f"── {_rg} ──")
            _bopts.extend(_by_r[_rg])
        with st.sidebar:
            st.markdown("**View as:**")
            _b = st.selectbox("View as", _bopts, key=sidebar_key,
                              label_visibility="collapsed")
    _clean = _b.replace("👥", "").strip().lower()
    _is_all = _clean in ("all team", "") or _b in ("👥 All team", "All team")
    return _b, _is_all

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
        .rule-card { border:0.5px solid var(--color-border-tertiary); border-radius:10px; overflow:hidden; margin-bottom:12px; }
        .rule-card-hdr { padding:14px 18px; display:flex; align-items:flex-start; gap:14px; }
        .rule-card-body { padding:0 18px 14px; }
        .rule-rail-red    { width:4px; flex-shrink:0; background:#E24B4A; border-radius:2px; min-height:40px; }
        .rule-rail-amber  { width:4px; flex-shrink:0; background:#EF9F27; border-radius:2px; min-height:40px; }
        .rule-rail-blue   { width:4px; flex-shrink:0; background:#4472C4; border-radius:2px; min-height:40px; }
        .rule-title-lg { font-size:15px; font-weight:600; color:var(--color-text-primary); margin-bottom:4px; }
        .rule-sub { font-size:12px; color:var(--color-text-secondary); line-height:1.5; }
        .stat-num { font-size:28px; font-weight:600; line-height:1; }
        .stat-sub { font-size:11px; color:var(--color-text-secondary); margin-top:2px; }
        .prog-bar { height:4px; background:rgba(128,128,128,.15); border-radius:2px; margin-top:6px; overflow:hidden; }
        .prog-fill-red   { height:100%; background:#E24B4A; border-radius:2px; }
        .prog-fill-amber { height:100%; background:#EF9F27; border-radius:2px; }
        .proj-row { display:flex; gap:10px; padding:8px 0; border-bottom:0.5px solid rgba(128,128,128,.1); align-items:flex-start; font-size:13px; }
        .proj-row:last-child { border-bottom:none; }
        .proj-name { font-weight:500; color:var(--color-text-primary); }
        .proj-type { font-size:11px; color:var(--color-text-secondary); }
        .proj-detail { font-size:12px; color:var(--color-text-secondary); }
        .fix-pill { display:inline-block; font-size:11px; padding:2px 9px; border-radius:20px; border:0.5px solid var(--color-border-secondary); color:var(--color-text-secondary); background:var(--color-background-secondary); }
        .status-fixed   { font-size:11px; font-weight:600; color:#27AE60; }
        .status-snoozed { font-size:11px; font-weight:600; color:#f59e0b; }
        .status-open    { font-size:11px; color:var(--color-text-secondary); }
        .strip-bar { background:rgba(59,130,246,.08); border:0.5px solid rgba(59,130,246,.2); border-radius:8px; padding:10px 16px; margin-bottom:16px; display:flex; align-items:center; justify-content:space-between; font-size:12px; }
        .cat-chip { display:inline-block; font-size:12px; padding:4px 12px; border-radius:20px; border:0.5px solid var(--color-border-tertiary); cursor:pointer; margin-right:6px; margin-bottom:6px; }
        .cat-chip-active { background:var(--color-background-info); color:var(--color-text-info); border-color:var(--color-border-info); }
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
_hero.markdown(f"<div style='background:linear-gradient(135deg,#1a56db 0%,#050D1F 55%,#050D1F 100%);padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;font-family:Manrope,sans-serif;position:relative;overflow:hidden'> <div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3B9EFF;margin-bottom:10px;font-family:Manrope,sans-serif'>Professional Services · Reporting</div> <h1 style='color:#fff;margin:0;font-size:28px;font-weight:800;font-family:Manrope,sans-serif;line-height:1.15'>DRS Health Check{_drs_title_sfx}</h1> <p style='color:rgba(255,255,255,0.6);margin:8px 0 0;font-size:14px;font-family:Manrope,sans-serif;line-height:1.6;max-width:520px'>Logical consistency validator for Smartsheet DRS data — flags fields and combinations that don't align with expected project state.</p> </div>", unsafe_allow_html=True)

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
    from shared.constants import CONSULTANT_DROPDOWN as _CD3
    _home_browse, _drs_all_team = _get_effective_browse(
        _session_name, _get_role(_session_name), _EL3, _RM3, _RO3, _AE3, _CD3, "drs_view_as"
    )
    _role = _get_role(_session_name)
    _is_manager = _role in ("manager", "manager_only")
    if _is_manager and _drs_all_team:
        _va_name, _va_region = None, None
    else:
        _va_name, _va_region, _is_group = resolve_view_as(
            _session_name, _home_browse, _ER3, _EL3, _RM3, _RO3, _AE3
        )
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
# RESULTS — v2: group by rule, session state, bulk fix surface
# ══════════════════════════════════════════════════════════════════════════════
df_findings = pd.DataFrame(findings)

if df_findings.empty:
    st.success(f"✓ No issues found across {len(df_drs):,} projects — DRS data looks clean.")
    st.stop()

n_error   = int((df_findings["severity"] == "Error").sum())
n_warning = int((df_findings["severity"] == "Warning").sum())
n_info    = int((df_findings["severity"] == "Info").sum())
n_projects = int(df_findings["project"].nunique())

# ── Session state helpers ─────────────────────────────────────────────────────
# State key: "drs_state_{project}_{rule_key}" → "open" | "fixed" | "snoozed"
# Snooze note: "drs_snooze_note_{project}_{rule_key}"
def _state_key(proj, rule):
    return f"drs_state_{proj}_{rule}".replace(" ", "_")[:120]

def _snooze_key(proj, rule):
    return f"drs_snooze_{proj}_{rule}".replace(" ", "_")[:120]

def _get_state(proj, rule):
    return st.session_state.get(_state_key(proj, rule), "open")

def _set_state(proj, rule, val):
    st.session_state[_state_key(proj, rule)] = val

# Add state to findings
df_findings["_state"] = df_findings.apply(
    lambda r: _get_state(r["project"], r["rule"]), axis=1
)

# ── Compute action-oriented header counts ─────────────────────────────────────
open_findings   = df_findings[df_findings["_state"] == "open"]
fixed_count     = int((df_findings["_state"] == "fixed").sum())
snoozed_count   = int((df_findings["_state"] == "snoozed").sum())

# "Action required errors ≥30d old" — date logic errors that are overdue
_date_err = open_findings[
    (open_findings["severity"] == "Error") &
    (open_findings["category"] == "Date Logic")
]
n_action_required = len(_date_err)

# "1-click fixable" — projects where the fix is a phase advance
_phase_fixable = open_findings[
    open_findings["rule"] == "Go Live date passed but phase not updated"
]
n_fixable = len(_phase_fixable)

# ── Strip header ──────────────────────────────────────────────────────────────
_strip_parts = []
if fixed_count:
    _strip_parts.append(f"<b style='color:#27AE60'>{fixed_count} fixed</b> staged for sync")
if snoozed_count:
    _strip_parts.append(f"<b style='color:#f59e0b'>{snoozed_count} snoozed</b>")
_strip_msg = " · ".join(_strip_parts) if _strip_parts else "No fixes staged yet this session"

st.markdown(
    f"<div class='strip-bar'>"
    f"<span style='color:var(--color-text-secondary)'>{_strip_msg}</span>"
    f"<span style='color:var(--color-text-secondary)'>Fixes staged locally — export CSV to apply in Smartsheet</span>"
    f"</div>",
    unsafe_allow_html=True
)

# ── Action metric strip (replaces 5-metric grid) ──────────────────────────────
_m1, _m2, _m3, _m4 = st.columns(4)
with _m1:
    st.markdown(
        f"<div style='border:0.5px solid var(--color-border-tertiary);border-radius:8px;padding:14px 16px'>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px'>Projects flagged</div>"
        f"<div style='font-size:28px;font-weight:600;color:var(--color-text-primary);line-height:1'>{n_projects}</div>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);margin-top:3px'>of {len(df_drs)} checked · "
        f"{n_error} errors · {n_warning} warnings</div></div>",
        unsafe_allow_html=True
    )
with _m2:
    _col = "#E24B4A" if n_action_required > 0 else "var(--color-text-primary)"
    st.markdown(
        f"<div style='border:0.5px solid var(--color-border-tertiary);border-radius:8px;padding:14px 16px'>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px'>Action required</div>"
        f"<div style='font-size:28px;font-weight:600;color:{_col};line-height:1'>{n_action_required}</div>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);margin-top:3px'>date logic errors · open</div></div>",
        unsafe_allow_html=True
    )
with _m3:
    _col3 = "#27AE60" if n_fixable > 0 else "var(--color-text-primary)"
    st.markdown(
        f"<div style='border:0.5px solid var(--color-border-tertiary);border-radius:8px;padding:14px 16px'>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px'>1-click fixable</div>"
        f"<div style='font-size:28px;font-weight:600;color:{_col3};line-height:1'>{n_fixable}</div>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);margin-top:3px'>phase advance · same fix</div></div>",
        unsafe_allow_html=True
    )
with _m4:
    st.markdown(
        f"<div style='border:0.5px solid var(--color-border-tertiary);border-radius:8px;padding:14px 16px'>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px'>Fixed this session</div>"
        f"<div style='font-size:28px;font-weight:600;color:#27AE60;line-height:1'>{fixed_count}</div>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);margin-top:3px'>{snoozed_count} snoozed</div></div>",
        unsafe_allow_html=True
    )

st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

# ── Category filter chips ─────────────────────────────────────────────────────
all_cats  = sorted(df_findings["category"].unique())
_cat_key  = "drs_v2_cat"
if _cat_key not in st.session_state:
    st.session_state[_cat_key] = "All"

_chip_html = ""
for _c in ["All"] + all_cats:
    _cnt = len(df_findings) if _c == "All" else int((df_findings["category"] == _c).sum())
    _active = "cat-chip-active" if st.session_state[_cat_key] == _c else ""
    _chip_html += f"<span class='cat-chip {_active}' onclick="">{_c} {_cnt}</span>"

_cat_cols = st.columns(len(all_cats) + 2)
for _ci, _c in enumerate(["All"] + all_cats):
    with _cat_cols[_ci]:
        _cnt = len(df_findings) if _c == "All" else int((df_findings["category"] == _c).sum())
        _is_active = st.session_state[_cat_key] == _c
        if st.button(
            f"{_c} · {_cnt}",
            key=f"drs_chip_{_c}",
            type="primary" if _is_active else "secondary",
            use_container_width=True,
        ):
            st.session_state[_cat_key] = _c
            st.rerun()

_active_cat = st.session_state[_cat_key]
_cat_findings = df_findings if _active_cat == "All" else df_findings[df_findings["category"] == _active_cat]

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# RULE CARDS — one card per unique (severity, category, rule) combination
# ══════════════════════════════════════════════════════════════════════════════
_SEV_RANK = {"Error": 0, "Warning": 1, "Info": 2}
_SEV_RAIL = {"Error": "rule-rail-red", "Warning": "rule-rail-amber", "Info": "rule-rail-blue"}
_SEV_COL  = {"Error": "#E24B4A", "Warning": "#EF9F27", "Info": "#4472C4"}
_SEV_BG   = {"Error": "rgba(226,75,74,.08)", "Warning": "rgba(245,158,11,.08)", "Info": "rgba(68,114,196,.08)"}

# Group by rule
_rule_groups = (
    _cat_findings
    .groupby(["severity", "category", "rule"], sort=False)
    .apply(lambda g: g)
    .reset_index(drop=True)
)
_rules = (
    _cat_findings[["severity", "category", "rule", "description", "expected"]]
    .drop_duplicates(subset=["severity", "category", "rule"])
    .assign(_rank=lambda d: d["severity"].map(_SEV_RANK))
    .sort_values("_rank")
    .reset_index(drop=True)
)

for _, rule_row in _rules.iterrows():
    sev      = rule_row["severity"]
    category = rule_row["category"]
    rule     = rule_row["rule"]
    desc     = rule_row["description"]
    expected = rule_row["expected"]

    _proj_group = _cat_findings[
        (_cat_findings["severity"] == sev) &
        (_cat_findings["rule"] == rule)
    ].copy()
    _proj_group["_state"] = _proj_group.apply(
        lambda r: _get_state(r["project"], r["rule"]), axis=1
    )

    n_total   = len(_proj_group)
    n_open    = int((_proj_group["_state"] == "open").sum())
    n_fixed   = int((_proj_group["_state"] == "fixed").sum())
    n_snoozed = int((_proj_group["_state"] == "snoozed").sum())
    pct_done  = int((n_fixed + n_snoozed) / n_total * 100) if n_total else 0

    _rail  = _SEV_RAIL.get(sev, "rule-rail-blue")
    _scol  = _SEV_COL.get(sev, "#4472C4")
    _sbg   = _SEV_BG.get(sev, "rgba(68,114,196,.08)")

    # Oldest project for this rule
    _oldest_days = None
    for _, _fr in _proj_group.iterrows():
        _gl = _fr.get("description", "")
        import re as _re_drs
        _m = _re_drs.search(r"(\d+)d ago", str(_gl))
        if _m:
            _d = int(_m.group(1))
            if _oldest_days is None or _d > _oldest_days:
                _oldest_days = _d

    _oldest_str = f" · oldest {_oldest_days}d" if _oldest_days else ""

    # Build bulk fix hint
    _BULK_HINTS = {
        "Go Live date passed but phase not updated":
            "Advance phase to 06. Go-Live / Hypercare, or push the go-live date out.",
        "On Hold Reason not set":
            "Set On Hold Reason + Responsible for Delay in My Projects → Project Detail.",
        "Responsible for Delay not set":
            "Set Responsible for Delay in My Projects → Project Detail.",
        "Engagement rating inconsistent with On Hold status":
            "Update Client Responsiveness to Not Responsive or Neutral.",
        "Sentiment rating inconsistent with On Hold status":
            "Update Client Sentiment to Neutral or Concerned.",
        "Actual hours exceed budget — no Change Order flagged":
            "Not bulk-fixable — each overrun needs a per-project judgment (log CO vs accept).",
    }
    _bulk_hint = _BULK_HINTS.get(rule, expected)

    with st.expander(
        f"{sev.upper()} · {category} — {rule}  ({n_total} project{'s' if n_total!=1 else ''}{_oldest_str})",
        expanded=(sev == "Error" and n_open > 0)
    ):
        # Rule summary row
        _hcol1, _hcol2 = st.columns([3, 1])
        with _hcol1:
            st.markdown(
                f"<div style='background:{_sbg};border-left:3px solid {_scol};"
                f"border-radius:0 6px 6px 0;padding:10px 14px;margin-bottom:12px'>"
                f"<div style='font-size:12px;color:{_scol};font-weight:600;margin-bottom:3px'>"
                f"{sev.upper()} · {category}</div>"
                f"<div style='font-size:13px;color:var(--color-text-secondary);line-height:1.5'>{_bulk_hint}</div>"
                f"</div>",
                unsafe_allow_html=True
            )
        with _hcol2:
            st.markdown(
                f"<div style='text-align:right;padding:4px 0'>"
                f"<div style='font-size:11px;color:var(--color-text-secondary)'>"
                f"Open {n_open} · Fixed {n_fixed} · Snoozed {n_snoozed}</div>"
                f"<div class='prog-bar' style='margin-top:6px'>"
                f"<div class='{'prog-fill-red' if sev=='Error' else 'prog-fill-amber'}' "
                f"style='width:{pct_done}%'></div></div>"
                f"<div style='font-size:11px;color:var(--color-text-secondary);margin-top:4px'>"
                f"{pct_done}% resolved this session</div>"
                f"</div>",
                unsafe_allow_html=True
            )

        # Bulk action bar — only for rules with a clear uniform fix
        _BULK_FIXABLE = {
            "Go Live date passed but phase not updated",
            "On Hold Reason not set",
            "Responsible for Delay not set",
            "Engagement rating inconsistent with On Hold status",
            "Sentiment rating inconsistent with On Hold status",
        }
        _open_projs = _proj_group[_proj_group["_state"] == "open"]["project"].tolist()

        if rule in _BULK_FIXABLE and _open_projs:
            _bulk_c1, _bulk_c2, _bulk_c3 = st.columns([2, 1, 1])
            with _bulk_c1:
                st.markdown(
                    f"<div style='font-size:12px;color:var(--color-text-secondary);padding-top:6px'>"
                    f"{len(_open_projs)} open project{'s' if len(_open_projs)!=1 else ''} — "
                    f"mark all as fixed or snoozed</div>",
                    unsafe_allow_html=True
                )
            with _bulk_c2:
                if st.button(
                    f"✓ Mark all {len(_open_projs)} fixed",
                    key=f"drs_bulk_fix_{sev}_{rule[:30]}",
                    use_container_width=True,
                    type="primary",
                ):
                    for _p in _open_projs:
                        _set_state(_p, rule, "fixed")
                    st.rerun()
            with _bulk_c3:
                if st.button(
                    f"⏸ Snooze all",
                    key=f"drs_bulk_snooze_{sev}_{rule[:30]}",
                    use_container_width=True,
                ):
                    for _p in _open_projs:
                        _set_state(_p, rule, "snoozed")
                    st.rerun()

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

        # Per-project rows — sorted by days overdue desc, then name
        _proj_sorted = _proj_group.copy()
        def _extract_days(desc_str):
            import re as _re2
            _m2 = _re2.search(r"(\d+)d ago", str(desc_str))
            return int(_m2.group(1)) if _m2 else 0
        _proj_sorted["_days"] = _proj_sorted["description"].apply(_extract_days)
        _proj_sorted = _proj_sorted.sort_values(["_days", "project"], ascending=[False, True])

        for _, finding in _proj_sorted.iterrows():
            _proj   = finding["project"]
            _desc   = finding["description"]
            _state  = _get_state(_proj, rule)
            _fkey   = f"drs_row_{_proj}_{rule}".replace(" ","_")[:100]

            # Parse project name and type from combined string
            _proj_parts = _proj.split(" — ") if " — " in _proj else [_proj]
            _pname = _proj_parts[0].strip()[:35]
            _ptype = " · ".join(_proj_parts[1:])[:40] if len(_proj_parts) > 1 else ""

            # State badge
            if _state == "fixed":
                _sbadge = "<span class='status-fixed'>✓ Fixed · staged</span>"
            elif _state == "snoozed":
                _snote = st.session_state.get(_snooze_key(_proj, rule), "")
                _sbadge = f"<span class='status-snoozed'>⏸ Snoozed{' · ' + _snote if _snote else ''}</span>"
            else:
                _sbadge = "<span class='status-open'>Open</span>"

            # Extract days overdue for colour
            import re as _re3
            _dm = _re3.search(r"(\d+)d ago", str(_desc))
            _days_str = _dm.group(0) if _dm else ""
            _days_int = int(_dm.group(1)) if _dm else 0
            _days_col = "#E24B4A" if _days_int > 90 else ("#EF9F27" if _days_int > 30 else "var(--color-text-secondary)")

            st.markdown(
                f"<div class='proj-row'>"
                f"<div style='flex:1;min-width:0'>"
                f"<div class='proj-name'>{_pname}</div>"
                f"<div class='proj-type'>{_ptype}</div>"
                f"<div class='proj-detail' style='margin-top:3px'>{_desc.replace(_days_str, f'<b style="color:{_days_col}">{_days_str}</b>', 1) if _days_str else _desc}</div>"
                f"</div>"
                f"<div style='flex-shrink:0;text-align:right;min-width:120px'>"
                f"{_sbadge}</div>"
                f"</div>",
                unsafe_allow_html=True
            )

            # Inline Fix / Snooze buttons (shown only when Open)
            if _state == "open":
                _bc1, _bc2, _bc3 = st.columns([3, 1, 1])
                with _bc2:
                    if st.button("✓ Fixed", key=f"{_fkey}_fix", use_container_width=True, type="primary"):
                        _set_state(_proj, rule, "fixed")
                        st.rerun()
                with _bc3:
                    if st.button("⏸ Snooze", key=f"{_fkey}_snooze", use_container_width=True):
                        _set_state(_proj, rule, "snoozed")
                        st.rerun()
            elif _state in ("fixed", "snoozed"):
                _bc1, _bc2 = st.columns([5, 1])
                with _bc2:
                    if st.button("↩ Reopen", key=f"{_fkey}_reopen", use_container_width=True):
                        _set_state(_proj, rule, "open")
                        st.rerun()

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── Export ────────────────────────────────────────────────────────────────────
_ec1, _ec2 = st.columns([3, 1])
with _ec1:
    # Only export open findings (fixed/snoozed excluded)
    _open_df = df_findings[df_findings.apply(
        lambda r: _get_state(r["project"], r["rule"]) == "open", axis=1
    )]
    _n_staged = fixed_count
    st.markdown(
        f"<div style='font-size:12px;color:var(--color-text-secondary);padding-top:8px'>"
        f"{len(_open_df)} open findings · "
        f"{_n_staged} fixed/staged excluded from export · "
        f"apply staged fixes directly in Smartsheet DRS or via My Projects</div>",
        unsafe_allow_html=True
    )
with _ec2:
    _pm_cols  = ["project_manager"] if "project_manager" in df_findings.columns else []
    _exp_cols = ["project"] + _pm_cols + ["severity","category","rule","description","expected"]
    _export   = _open_df[[c for c in _exp_cols if c in _open_df.columns]].rename(columns={
        "project": "Project", "project_manager": "Consultant",
        "severity": "Severity", "category": "Category",
        "rule": "Rule", "description": "Description", "expected": "Expected State",
    })
    st.download_button(
        label="⬇ Download findings as CSV",
        data=_export.to_csv(index=False),
        file_name=f"drs_health_check_{date.today().strftime('%Y%m%d')}.csv",
        mime="text/csv",
        use_container_width=True,
    )
