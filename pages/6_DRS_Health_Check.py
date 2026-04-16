"""
PS Tools — DRS Health Check
Logical consistency validator for SS DRS data.
Flags fields and combinations that don't align with expected project state.
"""
import streamlit as st
import pandas as pd
from datetime import date

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

st.markdown(f"""
<div style='background:#050D1F;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;font-family:Manrope,sans-serif;position:relative;overflow:hidden'>
    <svg style='position:absolute;right:-40px;top:50%;transform:translateY(-50%);opacity:0.06;width:200px;height:200px;pointer-events:none' viewBox='0 0 1482 1286.25' xmlns='http://www.w3.org/2000/svg'><g fill='#3B9EFF' fill-rule='evenodd'><path d='M975.127,924.953c2.608-2.68,1.744-5.496-.42-7.829l-57.415-61.872c-2.463-2.655-5.025-2.878-8.443-.991-10.398,5.739-19.024,12.314-27.949,19.885-83.252,70.621-197.471,155.494-298.93,195.556-17.993,7.105-35.256,13.178-54.191,17.329-62.148,13.627-131.853,15.491-192.702-5.298-64.93-22.183-113.878-68.722-142.715-130.542-28.647-61.415-22.393-131.406,11.352-189.217,2.598-2.793,1.405-6.055-1.389-8.184-35.341-26.918-40.303-33.439-69.367-65.686-1.449-1.607-4.102-2.401-5.903-1.138-13.105,9.189-23.232,20.534-33.172,32.961-16.499,20.629-29.73,42.605-38.718,67.541-5.127,10.469-8.378,20.486-10.885,32.065-13.633,62.973-7.701,128.685,17.402,188.142,23.839,56.463,65.297,103.638,114.77,139.169,32.418,23.283,66.848,42.548,103.476,58.385,25.142,10.871,50.281,18.994,76.934,25.12,96.392,22.153,188.876,4.496,276.774-38.393,42.916-20.94,83.188-45.685,121.922-73.568,75.733-54.514,154.643-126.72,219.571-193.435ZM1445.252,792.261c-7.628-38.507-22.817-74.472-43.124-107.897-35.582-58.566-85.801-106.77-139.329-149.092-69.784-55.176-145.355-102.407-225.163-141.162-2.165-1.052-4.941.388-5.391,1.627-.426,1.171-.463,3.413.931,4.628,20.341,17.734,39.847,35.55,58.599,55.093,13.286,14.465,26.223,28.012,37.022,44.544,19.784,30.289,35.735,62.168,50.127,95.397,34.512,31.926,64.863,67.358,90.813,106.359,42.427,63.765,57.696,142.663,37.453,217.116-11.436,42.061-34.763,80.507-64.388,112.265-55.859,59.882-133.144,94.711-214.71,99.157-32.507,1.773-64.093-.538-96.013-6.503-28.16-5.262-70.299-23.997-96.538-36.626-2.312-1.112-4.605-.743-6.449.974-12.635,11.76-25.076,22.901-39.051,33.146l-43.32,31.757c-2.68,1.965-2.195,5.562.439,7.808,70.707,60.309,165.779,100.179,259.837,97.033,39.996-1.336,78.686-6.594,117.486-16.111,94.178-23.099,174.952-71.91,236.526-146.957,23.873-29.096,44.355-60.51,59.779-94.956,29.172-65.148,38.357-137.461,24.463-207.601ZM601.099,242.903c-12.268,10.522-48.215,44.405-47.219,60.482.993,16.01,10.781,31.195,25.227,38.155,14.47,6.972,41.303-10.055,53.886-18.311l65.495-42.972c26.305-17.259,52.496-32.716,80.08-47.834l57.464-31.494c20.451-11.209,41.123-19.851,63.235-27.448,35.852-12.318,72.313-18.084,110.322-17.747,29.787.263,58.398,3.408,86.939,11.449,44.037,12.405,82.745,35.987,114.027,69.974,20.347,22.106,37.598,45.332,51.026,71.732,6.962,13.688,13.008,27.156,16.103,42.311,6.48,31.729,12.267,85.992-.676,115.916-6.013,13.902-13.009,26.627-18.289,40.753-.847,2.264-.768,4.767,1.387,6.461l81.366,63.967c2.003,1.574,5.098.298,6.46-1.592,19.285-26.745,34.599-55.578,45.667-86.804,10.617-29.953,15.416-60.246,15.218-92.192-.482-77.938-29.055-152.791-79.976-211.891-67.16-77.946-169.264-137.487-272.877-146.244-33.524-2.834-66.192-1.328-99.421,3.091-82.214,10.934-149.21,45.218-216.385,92.267-48.269,33.807-94.373,69.644-139.062,107.973ZM72.687,567.553c20.03,44.974,54.35,86.652,88.718,121.568,19.447,19.756,38.882,38.258,60.393,55.711l73.052,59.268c30.921,25.086,74.954,56.331,111.096,72.278,11.713,5.168,23.385,8.99,35.917,11.295,12.922,2.375,24.878,1.136,37.309-3.088,18.441-6.266,35.538-14.698,52.671-24.006,1.792-.974,2.85-2.213,3.058-3.936.179-1.483-.47-3.163-1.914-4.548-14.129-13.542-27.174-27.284-42.195-40.056l-78.193-66.48-93.5-82.422c-23.176-20.43-44.471-41.737-65.536-64.239-15.19-16.227-28.591-32.64-40.05-51.639-20.601-34.157-31.396-72.282-30.182-112.398.614-20.279,2.364-39.861,7.45-59.369,8.872-34.031,50.72-76.652,77.451-99.125,3.767-7.04,2.459-14.401,2.885-21.735.884-15.227,3.244-29.908,5.647-44.959,4.285-26.824,22.718-58.984,38.899-80.638,1.348-1.805,1.936-3.535.891-4.937-.951-1.277-2.618-2.49-4.589-2.222-52.436,7.145-104.92,34.806-146.088,67.704-25.632,20.484-48.458,43.456-68.934,69.137-46.339,58.118-62.952,131.49-53.428,204.864,4.697,36.186,14.376,70.75,29.171,103.971ZM1196.886,310.029c-4.882-10.39-12.371-18.773-20.659-26.723-18.771-18.007-40.425-31.674-64.291-42.362-57.569-25.783-110.906-28.064-173.214-22.213-61.067,5.735-111.183,25.069-164.567,54.081-24.678,13.412-48.301,26.866-71.885,42.28l-105.247,68.787c-85.308,55.756-195.138,156.138-256.755,237.876-1.598,2.12-2.206,4.81-.222,6.912l76.342,80.886c1.468,1.556,2.9,1.672,4.715,1.249,1.397-.326,1.99-1.717,2.793-3.377,3.117-6.44,6.665-11.977,11.238-17.864,38.52-49.59,82.099-94.54,130.222-135.261,40.87-34.583,82.783-67.442,126.68-98.902,83.71-59.991,188.529-115.793,291.15-127.921,23.653-2.795,46.328-.575,69.656,3.405,27.197,4.641,52.661,12.543,78.69,21.347l38.004,12.855c13.849,4.685,27.221-3.226,30.503-17.755,2.725-12.064,2.293-25.708-3.154-37.301Z'/></g></svg>
    <div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3B9EFF;margin-bottom:10px;font-family:Manrope,sans-serif'>Professional Services · Reporting</div>
    <h1 style='color:#fff;margin:0;font-size:28px;font-weight:800;font-family:Manrope,sans-serif;line-height:1.15'>DRS Health Check{_drs_title_sfx}</h1>
    <p style='color:rgba(255,255,255,0.6);margin:8px 0 0;font-size:14px;font-family:Manrope,sans-serif;line-height:1.6;max-width:520px'>Logical consistency validator for Smartsheet DRS data — flags fields and combinations that don't align with expected project state.</p>
</div>
""", unsafe_allow_html=True)

# ── Data source ───────────────────────────────────────────────────────────────
df_drs = st.session_state.get("df_drs")

if df_drs is not None:
    st.success("✓ Using SS DRS loaded from the Home page.")
    with st.expander("Override uploaded data for this page"):
        uploaded = st.file_uploader(
            "Drop SS DRS Export here",
            type=["xlsx", "csv"],
            key="drs_health_upload",
        )
        if uploaded:
            try:
                df_drs = load_drs(uploaded)
                st.success(f"Loaded {len(df_drs):,} rows from `{uploaded.name}`")
            except Exception as e:
                st.error(f"Could not load file: {e}")
                df_drs = None
else:
    uploaded = st.file_uploader(
        "Drop SS DRS Export here (or load on Home page)",
        type=["xlsx", "csv"],
        key="drs_health_upload",
    )
    if uploaded:
        try:
            df_drs = load_drs(uploaded)
            st.success(f"Loaded {len(df_drs):,} rows from `{uploaded.name}`")
        except Exception as e:
            st.error(f"Could not load file: {e}")
            df_drs = None

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
    go_live   = _get(row, "go_live_date")
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
