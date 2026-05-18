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
            "project":          proj,
            "project_id":       str(_get(row, "project_id", "") or ""),
            "consultant":       str(pm or ""),
            "phase":            str(phase or ""),
            "go_live_disp":     pd.to_datetime(go_live).strftime("%-d %b %Y") if _is_date(go_live) else "",
            "days_val":         abs(_days_until(go_live)) if (_is_date(go_live) and _days_until(go_live) is not None and _days_until(go_live) < 0) else (_days_since(start_dt) if _is_date(start_dt) else None),
            "severity":         severity,
            "category":         category,
            "rule":             rule,
            "description":      description,
            "expected":         expected,
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
# RESULTS — v3: sortable table + per-category tabs, read-only
# ══════════════════════════════════════════════════════════════════════════════
df_findings = pd.DataFrame(findings)

if df_findings.empty:
    st.success(f"✓ No issues found across {len(df_drs):,} projects — DRS data looks clean.")
    st.stop()

# ── Computed counts ───────────────────────────────────────────────────────────
n_error    = int((df_findings["severity"] == "Error").sum())
n_warning  = int((df_findings["severity"] == "Warning").sum())
n_info     = int((df_findings["severity"] == "Info").sum())
n_projects = int(df_findings["project"].nunique())
n_total    = len(df_findings)

# Oldest issue in days
_days_col = df_findings["days_val"].dropna()
oldest_d   = int(_days_col.max()) if len(_days_col) else 0

# ── 4 action metrics ──────────────────────────────────────────────────────────
_m1, _m2, _m3, _m4 = st.columns(4)
with _m1:
    st.markdown(
        f"<div style='border:0.5px solid var(--color-border-tertiary);border-radius:8px;padding:14px 16px'>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);text-transform:uppercase;"
        f"letter-spacing:.5px;margin-bottom:4px'>Projects flagged</div>"
        f"<div style='font-size:28px;font-weight:600;line-height:1'>{n_projects}</div>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);margin-top:3px'>"
        f"of {len(df_drs)} checked</div></div>", unsafe_allow_html=True)
with _m2:
    _ec = "#E24B4A" if n_error > 0 else "var(--color-text-primary)"
    st.markdown(
        f"<div style='border:0.5px solid var(--color-border-tertiary);border-radius:8px;padding:14px 16px'>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);text-transform:uppercase;"
        f"letter-spacing:.5px;margin-bottom:4px'>Errors</div>"
        f"<div style='font-size:28px;font-weight:600;color:{_ec};line-height:1'>{n_error}</div>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);margin-top:3px'>"
        f"{n_warning} warnings · {n_info} info</div></div>", unsafe_allow_html=True)
with _m3:
    _dc = "#E24B4A" if oldest_d > 90 else ("#EF9F27" if oldest_d > 30 else "var(--color-text-primary)")
    st.markdown(
        f"<div style='border:0.5px solid var(--color-border-tertiary);border-radius:8px;padding:14px 16px'>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);text-transform:uppercase;"
        f"letter-spacing:.5px;margin-bottom:4px'>Oldest issue</div>"
        f"<div style='font-size:28px;font-weight:600;color:{_dc};line-height:1'>{oldest_d}d</div>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);margin-top:3px'>"
        f"days since go-live / start</div></div>", unsafe_allow_html=True)
with _m4:
    st.markdown(
        f"<div style='border:0.5px solid var(--color-border-tertiary);border-radius:8px;padding:14px 16px'>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);text-transform:uppercase;"
        f"letter-spacing:.5px;margin-bottom:4px'>Total findings</div>"
        f"<div style='font-size:28px;font-weight:600;line-height:1'>{n_total}</div>"
        f"<div style='font-size:11px;color:var(--color-text-secondary);margin-top:3px'>"
        f"across {df_findings['category'].nunique()} rule categories</div></div>",
        unsafe_allow_html=True)

st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS — Overview + one per category
# ══════════════════════════════════════════════════════════════════════════════
_SEV_RANK = {"Error": 0, "Warning": 1, "Info": 2}
df_findings["_sev_rank"] = df_findings["severity"].map(_SEV_RANK).fillna(9)

# Sort by severity then days desc for default view
df_sorted = df_findings.sort_values(
    ["_sev_rank", "days_val"], ascending=[True, False]
).reset_index(drop=True)

# Build tab labels
_cats = sorted(df_findings["category"].unique(),
               key=lambda c: -int((df_findings["category"] == c).sum()))
_tab_labels = [f"Overview · {n_total}"] + [
    f"{c} · {int((df_findings['category']==c).sum())}" for c in _cats
]
_tabs = st.tabs(_tab_labels)

# ── Helper: render a sortable st.dataframe ────────────────────────────────────
def _sev_display(s):
    """Map severity to short display string for column."""
    return s

def _days_display(v):
    if v is None or (hasattr(v, '__class__') and v.__class__.__name__ in ('float',)) and str(v) == 'nan':
        return None
    try:
        return int(v)
    except Exception:
        return None

# ── OVERVIEW TAB ──────────────────────────────────────────────────────────────
with _tabs[0]:
    st.markdown(
        "<div style='font-size:12px;color:var(--color-text-secondary);"
        "margin-bottom:10px'>All findings · sorted by severity then days overdue. "
        "Click any column header to re-sort.</div>",
        unsafe_allow_html=True
    )

    _ov_df = df_sorted[[
        "project", "consultant", "category", "severity", "rule",
        "phase", "go_live_disp", "days_val", "description"
    ]].copy()
    _ov_df.columns = [
        "Project", "Consultant", "Category", "Severity", "Rule",
        "Phase", "Go live", "Days", "Description"
    ]
    _ov_df["Days"] = _ov_df["Days"].apply(_days_display)

    st.dataframe(
        _ov_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Project":     st.column_config.TextColumn("Project",     width="medium"),
            "Consultant":  st.column_config.TextColumn("Consultant",  width="small"),
            "Category":    st.column_config.TextColumn("Category",    width="small"),
            "Severity":    st.column_config.TextColumn("Severity",    width="small"),
            "Rule":        st.column_config.TextColumn("Rule",        width="medium"),
            "Phase":       st.column_config.TextColumn("Phase",       width="small"),
            "Go live":     st.column_config.TextColumn("Go live",     width="small"),
            "Days":        st.column_config.NumberColumn(
                               "Days",
                               help="Days since go-live passed (date logic) or since project start (hours/completeness)",
                               width="small",
                               format="%d",
                           ),
            "Description": st.column_config.TextColumn("Description", width="large"),
        },
    )

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # CSV export
    _ec1, _ec2 = st.columns([4, 1])
    with _ec1:
        st.markdown(
            f"<div style='font-size:12px;color:var(--color-text-secondary);padding-top:8px'>"
            f"{n_total} findings across {n_projects} projects · "
            f"fix issues in <b>My Projects → Project Detail</b> then re-sync DRS to clear flags</div>",
            unsafe_allow_html=True
        )
    with _ec2:
        _exp_cols = ["Project", "Consultant", "Category", "Severity", "Rule",
                     "Phase", "Go live", "Days", "Description"]
        st.download_button(
            label="⬇ Download CSV",
            data=_ov_df[_exp_cols].to_csv(index=False),
            file_name=f"drs_health_check_{date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

# ── CATEGORY TABS ─────────────────────────────────────────────────────────────
# Column config per category — show the most relevant columns for each rule type
_CAT_CONFIG = {
    "Date Logic": {
        "cols":  ["Project", "Consultant", "Rule", "Phase", "Go live", "Days", "Description"],
        "note":  "Fix: update phase or go-live date in My Projects → Project Detail, then sync to Smartsheet.",
        "sort":  "Days",
    },
    "On Hold Data Quality": {
        "cols":  ["Project", "Consultant", "Rule", "Phase", "Days", "Description"],
        "note":  "Fix: set On Hold Reason, Responsible for Delay, and client fields in My Projects → Project Detail.",
        "sort":  "Days",
    },
    "Hours vs Scope": {
        "cols":  ["Project", "Consultant", "Rule", "Phase", "Days", "Description"],
        "note":  "Fix: log a Change Order in Smartsheet or review budget allocation. Each project needs individual judgment.",
        "sort":  "Days",
    },
    "Completeness": {
        "cols":  ["Project", "Consultant", "Rule", "Phase", "Description"],
        "note":  "Fix: complete missing fields in Smartsheet DRS or My Projects → Project Detail.",
        "sort":  None,
    },
    "Status Conflict": {
        "cols":  ["Project", "Consultant", "Rule", "Phase", "Description"],
        "note":  "Fix: align Status and RAG in My Projects → Project Detail.",
        "sort":  None,
    },
    "Activity Conflict": {
        "cols":  ["Project", "Consultant", "Rule", "Phase", "Days", "Description"],
        "note":  "Fix: update client responsiveness or project status in My Projects → Project Detail.",
        "sort":  "Days",
    },
    "Milestone Sequence": {
        "cols":  ["Project", "Consultant", "Rule", "Phase", "Description"],
        "note":  "Fix: correct milestone dates in My Projects → Project Detail, then sync to Smartsheet.",
        "sort":  None,
    },
    "Phase vs Milestone": {
        "cols":  ["Project", "Consultant", "Rule", "Phase", "Description"],
        "note":  "Fix: advance phase or back-date milestone in My Projects → Project Detail.",
        "sort":  None,
    },
}

_SEV_COL = {"Error": "#E24B4A", "Warning": "#EF9F27", "Info": "#4472C4"}

for _ti, _cat in enumerate(_cats):
    with _tabs[_ti + 1]:
        _cat_df = df_sorted[df_sorted["category"] == _cat].copy()
        _n_cat  = len(_cat_df)
        _n_err  = int((_cat_df["severity"] == "Error").sum())
        _n_warn = int((_cat_df["severity"] == "Warning").sum())
        _cfg    = _CAT_CONFIG.get(_cat, {
            "cols": ["Project", "Consultant", "Rule", "Severity", "Phase", "Days", "Description"],
            "note": "Fix issues in My Projects → Project Detail, then sync to Smartsheet.",
            "sort": "Days",
        })

        # Severity breakdown line
        _sev_parts = []
        if _n_err:  _sev_parts.append(f"<span style='color:#E24B4A;font-weight:600'>{_n_err} error{'s' if _n_err!=1 else ''}</span>")
        if _n_warn: _sev_parts.append(f"<span style='color:#EF9F27;font-weight:600'>{_n_warn} warning{'s' if _n_warn!=1 else ''}</span>")
        _n_info_cat = int((_cat_df["severity"] == "Info").sum())
        if _n_info_cat: _sev_parts.append(f"{_n_info_cat} info")

        st.markdown(
            f"<div style='display:flex;align-items:baseline;gap:12px;margin-bottom:8px'>"
            f"<span style='font-size:13px;color:var(--color-text-secondary)'>"
            f"{_n_cat} finding{'s' if _n_cat!=1 else ''} · "
            f"{'  ·  '.join(_sev_parts)}</span></div>",
            unsafe_allow_html=True
        )

        # Fix note
        st.markdown(
            f"<div style='font-size:12px;color:var(--color-text-secondary);"
            f"background:var(--color-background-secondary);border-radius:6px;"
            f"padding:8px 12px;margin-bottom:10px;border-left:3px solid var(--color-border-secondary);"
            f"border-radius:0 6px 6px 0'>"
            f"<b style='color:var(--color-text-primary)'>How to fix:</b> {_cfg['note']}</div>",
            unsafe_allow_html=True
        )

        # Build display df — rename to match column config keys
        _display = _cat_df[[
            "project", "consultant", "severity", "rule",
            "phase", "go_live_disp", "days_val", "description"
        ]].copy()
        _display.columns = [
            "Project", "Consultant", "Severity", "Rule",
            "Phase", "Go live", "Days", "Description"
        ]
        _display["Days"] = _display["Days"].apply(_days_display)

        # Sort
        if _cfg["sort"] == "Days" and "Days" in _display.columns:
            _display = _display.sort_values("Days", ascending=False, na_position="last")

        # Only keep relevant columns
        _show_cols = [c for c in _cfg["cols"] if c in _display.columns]

        # Column config for this tab
        _col_cfg = {
            "Project":     st.column_config.TextColumn("Project",     width="medium"),
            "Consultant":  st.column_config.TextColumn("Consultant",  width="small"),
            "Severity":    st.column_config.TextColumn("Severity",    width="small"),
            "Rule":        st.column_config.TextColumn("Rule",        width="medium"),
            "Phase":       st.column_config.TextColumn("Phase",       width="small"),
            "Go live":     st.column_config.TextColumn("Go live",     width="small"),
            "Days":        st.column_config.NumberColumn(
                               "Days overdue",
                               help="Days since go-live passed or since project started",
                               width="small",
                               format="%d",
                           ),
            "Description": st.column_config.TextColumn("Description", width="large"),
        }

        st.dataframe(
            _display[_show_cols].reset_index(drop=True),
            use_container_width=True,
            hide_index=True,
            column_config={k: v for k, v in _col_cfg.items() if k in _show_cols},
        )

        # Per-category CSV
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        _cc1, _cc2 = st.columns([4, 1])
        with _cc1:
            _n_proj_cat = int(_cat_df["project"].nunique())
            st.markdown(
                f"<div style='font-size:12px;color:var(--color-text-secondary);padding-top:8px'>"
                f"{_n_cat} findings across {_n_proj_cat} project{'s' if _n_proj_cat!=1 else ''}</div>",
                unsafe_allow_html=True
            )
        with _cc2:
            st.download_button(
                label="⬇ Download CSV",
                data=_display[_show_cols].to_csv(index=False),
                file_name=f"drs_{_cat.lower().replace(' ','_')}_{date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
                key=f"dl_{_cat}",
            )
