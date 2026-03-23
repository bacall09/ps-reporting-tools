"""
PS Tools — My Projects
Per-consultant project working list: snapshot metrics, needs-action items,
active projects table, on-hold projects, and a change export flow.
"""
import streamlit as st
import pandas as pd
import io
from datetime import date, timedelta

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN, ACTIVE_EMPLOYEES,
    MILESTONE_COLS_MAP, get_role, is_manager, LEAVER_EXIT_DATES,
)
from shared.config import (
    AVAIL_HOURS, EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP, DEFAULT_SCOPE,
)
from shared.loaders import calc_days_inactive, calc_last_milestone
from shared.template_utils import suggest_tier, TEMPLATES

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
<style>
    html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
    .section-label { font-size:11px; font-weight:700; text-transform:uppercase;
                     letter-spacing:0.8px; color:#4472C4; margin-bottom:8px; }
    .metric-card   { background:transparent; border:1px solid rgba(128,128,128,0.2);
                     border-radius:8px; padding:16px 20px; margin-bottom:12px; }
    .metric-val    { font-size:26px; font-weight:700; color:inherit; }
    .metric-lbl    { font-size:12px; opacity:0.6; margin-top:2px; }
    .proj-flag     { display:inline-block; font-size:11px; font-weight:600;
                     padding:2px 7px; border-radius:4px; margin-right:4px; }
    .flag-error    { background:rgba(231,76,60,0.15);  color:#E74C3C; }
    .flag-warn     { background:rgba(243,156,18,0.15); color:#D68910; }
    .flag-info     { background:rgba(68,114,196,0.15); color:#4472C4; }
    .flag-ok       { background:rgba(39,174,96,0.15);  color:#27AE60; }
    .rag-R { display:inline-block;width:10px;height:10px;border-radius:50%;background:#E74C3C;margin-right:5px; }
    .rag-A { display:inline-block;width:10px;height:10px;border-radius:50%;background:#F39C12;margin-right:5px; }
    .rag-G { display:inline-block;width:10px;height:10px;border-radius:50%;background:#27AE60;margin-right:5px; }
    .rag-  { display:inline-block;width:10px;height:10px;border-radius:50%;background:rgba(128,128,128,0.4);margin-right:5px; }
    .divider { border:none; border-top:1px solid rgba(128,128,128,0.2); margin:20px 0; }
    .edit-form { background:rgba(68,114,196,0.06); border:1px solid rgba(68,114,196,0.2);
                 border-radius:8px; padding:14px 16px; margin-top:8px; }
</style>
""", unsafe_allow_html=True)

# ── Identity ──────────────────────────────────────────────────────────────────
selected = st.session_state.get("consultant_name", "")
role     = get_role(selected) if selected else "consultant"
today    = pd.Timestamp.today().normalize()

# ── Sidebar — View As (managers/team leads only) ─────────────────────────────
view_as = selected
with st.sidebar:
    if role == "manager":
        from shared.config import EMPLOYEE_LOCATION as _EL, PS_REGION_MAP as _RM, PS_REGION_OVERRIDE as _RO

        def _gr(name):
            if name in _RO: return _RO[name]
            return _RM.get(_EL.get(name, ""), "Other")

        _active_c = sorted([
            n for n in CONSULTANT_DROPDOWN
            if get_role(n) in ("consultant", "manager")
            and len(EMPLOYEE_ROLES.get(n, {}).get("products", [])) > 0
        ])
        _by_rgn = {}
        for _cn in _active_c:
            _by_rgn.setdefault(_gr(_cn), []).append(_cn)

        _bopts = ["— My own projects —"]
        for _rg in sorted(_by_rgn.keys()):
            _bopts.append(f"── {_rg} ──")
            _bopts.extend(_by_rgn[_rg])

        st.markdown("**View as:**")
        _browse = st.selectbox("Consultant", _bopts,
                               key="myproj_browse", label_visibility="collapsed")

        if _browse != "— My own projects —" and not (_browse.startswith("── ") and _browse.endswith(" ──")):
            view_as = _browse

# ── Data ─────────────────────────────────────────────────────────────────────
df_drs = st.session_state.get("df_drs")
df_ns  = st.session_state.get("df_ns")

if df_drs is None:
    st.markdown('<div class="section-label">My Projects</div>', unsafe_allow_html=True)
    st.info("Upload SS DRS Export in the sidebar to see your projects.")
    st.stop()

# Filter to this consultant's projects
_va_parts = [p.strip() for p in view_as.split(",")]
_va_variants = {view_as.lower(), _va_parts[0].lower()}
if len(_va_parts) == 2:
    _va_variants.add(f"{_va_parts[1].strip()} {_va_parts[0]}".lower())

def _match_pm(v):
    v = str(v).strip().lower()
    return (v in _va_variants or
            any(v == nv or v.startswith(nv + " ") or v.endswith(" " + nv)
                for nv in _va_variants))

pm_col = df_drs.get("project_manager", pd.Series(dtype=str))
my_drs = df_drs[pm_col.apply(lambda v: _match_pm(str(v)))].copy()

# If no match and not a manager, show all (fallback)
if my_drs.empty and not is_manager(view_as):
    my_drs = df_drs.copy()

# Enrich with NS days inactive if available
if df_ns is not None:
    try:
        my_drs = calc_days_inactive(my_drs, df_ns)
    except Exception:
        pass

# Separate on-hold
on_hold = my_drs[my_drs.get("_on_hold", pd.Series(False, index=my_drs.index)) == True].copy()
active  = my_drs[my_drs.get("_on_hold", pd.Series(False, index=my_drs.index)) != True].copy()

# ── Helper: DRS health flags per project ─────────────────────────────────────
PHASE_ORDER = [
    "00. onboarding", "01. requirements and design", "02. configuration",
    "03. enablement/training", "04. uat", "05. prep for go-live",
    "06. go-live", "07. data migration", "08. ready for support transition",
    "09. phase 2 scoping",
]

def _phase_idx(p):
    pl = str(p).strip().lower()
    for i, ph in enumerate(PHASE_ORDER):
        if pl.startswith(ph[:6]) or ph in pl or pl in ph:
            return i
    return -1

def _get_flags(row):
    """Return list of (severity, field_key, message, editable) for a project row.
    Only flags things a consultant can actually fix in SS DRS.
    Go Live Date and Overall RAG are calculated by SS — not editable directly.
    """
    flags = []
    phase     = str(row.get("phase", "") or "").strip()
    go_live   = row.get("go_live_date")
    start_dt  = row.get("start_date")
    is_legacy = bool(row.get("legacy", False))
    phase_idx = _phase_idx(phase)

    # Go Live passed but phase not updated — fix by updating Phase (editable)
    if pd.notna(go_live) and phase_idx >= 0:
        if pd.Timestamp(go_live) < today and phase_idx < _phase_idx("06. go-live"):
            flags.append(("error", "phase", "Go Live passed — update Phase to 06. Go-Live", True))

    # Go Live before Start date — not editable (Go Live is calculated in SS)
    # Surface as info so consultant knows to raise with admin
    if pd.notna(go_live) and pd.notna(start_dt):
        if pd.Timestamp(go_live) < pd.Timestamp(start_dt):
            flags.append(("error", None, "Go Live before Start date — raise with admin (not editable)", False))

    # Missing intro email — editable milestone date
    if not is_legacy and phase_idx > _phase_idx("00. onboarding"):
        if not pd.notna(row.get("ms_intro_email")):
            flags.append(("warn", "ms_intro_email", "Intro email date not recorded", True))

    # Schedule health blank or concerning — editable
    sched = str(row.get("schedule_health", "") or "").strip()
    if not sched:
        flags.append(("warn", "schedule_health", "Schedule Health not set", True))

    # Missing critical milestones (non-legacy)
    if not is_legacy and phase_idx >= 0:
        critical = {
            "ms_uat_signoff":  _phase_idx("04. uat"),
            "ms_prod_cutover": _phase_idx("05. prep for go-live"),
            "ms_hypercare_start": _phase_idx("06. go-live"),
        }
        for ms_col, exp_idx in critical.items():
            if phase_idx > exp_idx and not pd.notna(row.get(ms_col)):
                flags.append(("warn", ms_col, f"Missing milestone: {MILESTONE_COLS_MAP.get(ms_col, ms_col)}", True))

    return flags

active["_flags"] = active.apply(_get_flags, axis=1)
active["_n_errors"] = active["_flags"].apply(lambda f: sum(1 for s,_,_m,_e in f if s=="error"))
active["_n_warns"]  = active["_flags"].apply(lambda f: sum(1 for s,_,_m,_e in f if s=="warn"))
active["_needs_action"] = (
    (active["_n_errors"] > 0) |
    (active["_n_warns"] > 0) |
    (active.get("days_inactive", pd.Series(-1, index=active.index)) >= 14)
)

# ── Header ────────────────────────────────────────────────────────────────────
_disp = view_as.split(",")[1].strip() + " " + view_as.split(",")[0] if "," in view_as else view_as
st.markdown(f'<div style="font-size:24px;font-weight:700;margin-bottom:4px">My Projects — {_disp}</div>',
            unsafe_allow_html=True)
st.markdown(f'<div style="font-size:13px;opacity:0.6;margin-bottom:20px">'
            f'{today.strftime("%A, %B %-d %Y")} · {len(active)} active · {len(on_hold)} on hold'
            f'</div>', unsafe_allow_html=True)
st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Snapshot metrics
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Snapshot</div>', unsafe_allow_html=True)

# ── Projects by phase ────────────────────────────────────────────────────────
phase_counts = active["phase"].value_counts() if "phase" in active.columns else pd.Series()

# ── Going live this week ──────────────────────────────────────────────────────
_next7 = today + timedelta(days=7)
going_live_soon = active[
    active["go_live_date"].notna() &
    (active["go_live_date"] >= today) &
    (active["go_live_date"] <= _next7)
] if "go_live_date" in active.columns else pd.DataFrame()

# ── In hypercare week 1 ───────────────────────────────────────────────────────
_14d_ago = today - timedelta(days=14)
in_hypercare = active[
    active["phase"].fillna("").str.lower().str.contains("06|go-live|go live|hypercare", na=False) &
    (active["go_live_date"].notna() if "go_live_date" in active.columns else pd.Series(False, index=active.index)) &
    (active["go_live_date"] >= _14d_ago if "go_live_date" in active.columns else pd.Series(False, index=active.index))
] if "go_live_date" in active.columns else pd.DataFrame()

# ── Missing intro email (non-legacy) ─────────────────────────────────────────
_phase_post_onboarding = active["phase"].fillna("").apply(
    lambda p: _phase_idx(p) > _phase_idx("00. onboarding")
)
missing_intro = active[
    ~active.get("legacy", pd.Series(False, index=active.index)).astype(bool) &
    _phase_post_onboarding &
    (~active["ms_intro_email"].notna() if "ms_intro_email" in active.columns
     else pd.Series(True, index=active.index))
] if "ms_intro_email" in active.columns else pd.DataFrame()

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown('<div class="metric-card">'
                f'<div class="metric-val">{len(active)}</div>'
                f'<div class="metric-lbl">Active Projects</div>'
                '</div>', unsafe_allow_html=True)
    # Mini phase breakdown
    if not phase_counts.empty:
        for ph, cnt in phase_counts.head(5).items():
            ph_short = str(ph).split(".")[-1].strip()[:25] if ph else "Unknown"
            st.markdown(f'<div style="font-size:11px;opacity:0.7;padding:1px 0">'
                        f'{cnt} · {ph_short}</div>', unsafe_allow_html=True)

with c2:
    col = "#E74C3C" if len(going_live_soon) > 0 else "inherit"
    st.markdown('<div class="metric-card">'
                f'<div class="metric-val" style="color:{col}">{len(going_live_soon)}</div>'
                f'<div class="metric-lbl">Going live this week</div>'
                '</div>', unsafe_allow_html=True)
    for _, r in going_live_soon.iterrows():
        pn = str(r.get("project_name",""))[:30]
        gl = pd.Timestamp(r["go_live_date"]).strftime("%-d %b")
        st.markdown(f'<div style="font-size:11px;opacity:0.7;padding:1px 0">{pn} · {gl}</div>',
                    unsafe_allow_html=True)

with c3:
    col = "#F39C12" if len(in_hypercare) > 0 else "inherit"
    st.markdown('<div class="metric-card">'
                f'<div class="metric-val" style="color:{col}">{len(in_hypercare)}</div>'
                f'<div class="metric-lbl">In hypercare (week 1)</div>'
                '</div>', unsafe_allow_html=True)
    for _, r in in_hypercare.iterrows():
        pn = str(r.get("project_name",""))[:30]
        gl = pd.Timestamp(r["go_live_date"]).strftime("%-d %b")
        st.markdown(f'<div style="font-size:11px;opacity:0.7;padding:1px 0">{pn} · went live {gl}</div>',
                    unsafe_allow_html=True)

with c4:
    col = "#E74C3C" if len(missing_intro) > 0 else "inherit"
    st.markdown('<div class="metric-card">'
                f'<div class="metric-val" style="color:{col}">{len(missing_intro)}</div>'
                f'<div class="metric-lbl">Missing intro email</div>'
                '</div>', unsafe_allow_html=True)
    if len(missing_intro) > 0:
        st.markdown('<div style="font-size:11px;opacity:0.6;padding:2px 0">'
                    'Excl. legacy projects</div>', unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Needs Action
# ══════════════════════════════════════════════════════════════════════════════
needs_action = active[active["_needs_action"]].sort_values(
    ["_n_errors", "_n_warns", "days_inactive"], ascending=[False, False, False]
)

st.markdown('<div class="section-label">Needs Action</div>', unsafe_allow_html=True)

if needs_action.empty:
    st.markdown('<span style="font-size:13px;color:#27AE60">✓ No projects currently need attention.</span>',
                unsafe_allow_html=True)
else:
    st.markdown(f'**{len(needs_action)} project(s)** need attention', unsafe_allow_html=True)

    # Session state for queued changes
    if "queued_changes" not in st.session_state:
        st.session_state["queued_changes"] = {}

    for _ri, (idx, row) in enumerate(needs_action.iterrows()):
        proj_name  = str(row.get("project_name", "—"))
        phase      = str(row.get("phase", "—"))
        rag        = str(row.get("rag", "") or "").strip().upper()[:1]
        days_inac  = int(row.get("days_inactive", 0))
        flags      = row.get("_flags", [])
        tier       = suggest_tier(days_inac) if days_inac >= 14 else None

        # Build expander label
        _label_parts = [proj_name]
        if row["_n_errors"]: _label_parts.append(f"{row['_n_errors']} Error{'s' if row['_n_errors']>1 else ''}")
        if row["_n_warns"]:  _label_parts.append(f"{row['_n_warns']} Warning{'s' if row['_n_warns']>1 else ''}")
        if days_inac >= 14:  _label_parts.append(f"{days_inac}d inactive")

        with st.expander("  ·  ".join(_label_parts), expanded=(row["_n_errors"] > 0)):
            col_info, col_action = st.columns([3, 2])

            with col_info:
                # Flag badges — show editable vs info-only differently
                badge_html = ""
                for sev, field_key, msg, editable in flags:
                    cls = "flag-error" if sev == "error" else "flag-warn"
                    icon = "✏️ " if editable else "ℹ️ "
                    badge_html += f'<span class="proj-flag {cls}" title="{msg}">{icon}{msg}</span><br>'
                if badge_html:
                    st.markdown(badge_html, unsafe_allow_html=True)
                    st.markdown("")

                st.markdown(f"**Phase:** {phase}")
                rag_dot = f'<span class="rag-{rag}"></span>'
                st.markdown(f'**RAG:** {rag_dot}{rag or "—"}', unsafe_allow_html=True)
                if days_inac >= 0:
                    st.markdown(f"**Days inactive:** {days_inac}")
                go_live = row.get("go_live_date")
                if pd.notna(go_live):
                    st.markdown(f"**Go Live:** {pd.Timestamp(go_live).strftime('%-d %b %Y')}")

            with col_action:
                # Re-engagement draft button
                if tier and tier in TEMPLATES:
                    st.markdown(f"**Re-engagement:** {tier}")
                    if st.button("Draft outreach →", key=f"na_draft_{_ri}", type="primary"):
                        st.session_state["_jump_to_proj"] = proj_name
                        st.session_state["_jump_tier"]    = tier
                        st.switch_page("pages/2_Customer_Reengagement.py")

                # Edit form — only editable fields
                editable_flags = [(fk, msg) for _, fk, msg, ed in flags if ed and fk]
                if editable_flags:
                    st.markdown("**Update in DRS:**")
                    _edit_key = f"edit_{proj_name[:30]}"
                    _queued   = st.session_state["queued_changes"].get(proj_name, {})
                    _new_vals = dict(_queued)

                    PHASE_OPTIONS = [
                        "00. Onboarding", "01. Requirements and Design",
                        "02. Configuration", "03. Enablement/Training",
                        "04. UAT", "05. Prep for Go-Live",
                        "06. Go-Live (Hypercare)", "07. Data Migration",
                        "08. Ready for Support Transition", "09. Phase 2 Scoping",
                        "10. Complete/Pending Final Billing",
                    ]
                    SCHEDULE_OPTIONS = ["On Track", "At Risk", "Behind", "Significantly Behind"]
                    RESOURCE_OPTIONS = ["Available", "At Capacity", "Over Capacity"]
                    SCOPE_OPTIONS    = ["On Track", "At Risk", "Changed", "Unchanged"]

                    with st.container():
                        st.markdown('<div class="edit-form">', unsafe_allow_html=True)
                        for field_key, label in editable_flags:
                            current_raw = row.get(field_key)

                            if field_key == "phase":
                                _idx = next((i for i, p in enumerate(PHASE_OPTIONS)
                                             if phase.lower() in p.lower()), 0)
                                _val = st.selectbox("Phase", PHASE_OPTIONS,
                                                    index=_idx, key=f"{_edit_key}_phase")
                                if _val != phase: _new_vals["Project Phase"] = _val

                            elif field_key == "schedule_health":
                                _cur = str(current_raw or "")
                                _idx = SCHEDULE_OPTIONS.index(_cur) if _cur in SCHEDULE_OPTIONS else 0
                                _val = st.selectbox("Schedule Health", SCHEDULE_OPTIONS,
                                                    index=_idx, key=f"{_edit_key}_sched")
                                if _val != _cur: _new_vals["Schedule Health"] = _val

                            elif field_key in ("ms_intro_email","ms_uat_signoff",
                                               "ms_prod_cutover","ms_hypercare_start",
                                               "ms_config_start","ms_enablement",
                                               "ms_session1","ms_session2",
                                               "ms_close_out","ms_transition"):
                                _ms_label = MILESTONE_COLS_MAP.get(field_key, label)
                                _cur_date = pd.Timestamp(current_raw).date() if pd.notna(current_raw) else date.today()
                                _val = st.date_input(_ms_label, value=_cur_date,
                                                     key=f"{_edit_key}_{field_key}")
                                _ss_col = {
                                    "ms_intro_email":     "Intro. Email Sent",
                                    "ms_config_start":    "Standard Config Start",
                                    "ms_enablement":      "Enablement Session",
                                    "ms_session1":        "Session #1",
                                    "ms_session2":        "Session #2",
                                    "ms_uat_signoff":     "UAT Signoff",
                                    "ms_prod_cutover":    "Prod Cutover",
                                    "ms_hypercare_start": "Hypercare Start",
                                    "ms_close_out":       "Close Out Remaining Tasks",
                                    "ms_transition":      "Transition to Support",
                                }.get(field_key, field_key)
                                if str(_val) != str(_cur_date): _new_vals[_ss_col] = str(_val)

                        c_save, c_clear = st.columns(2)
                        with c_save:
                            if st.button("Queue change", key=f"save_{_edit_key}",
                                         use_container_width=True):
                                if _new_vals:
                                    _new_vals["Project Name"] = proj_name
                                    st.session_state["queued_changes"][proj_name] = _new_vals
                                    st.success("✓ Queued")
                        with c_clear:
                            if _queued and st.button("Clear", key=f"clear_{_edit_key}",
                                                      use_container_width=True):
                                st.session_state["queued_changes"].pop(proj_name, None)
                                st.rerun()
                        st.markdown('</div>', unsafe_allow_html=True)

                    if proj_name in st.session_state["queued_changes"]:
                        st.markdown('<span style="font-size:11px;color:#27AE60">✓ Change queued</span>',
                                    unsafe_allow_html=True)

# ── Export queued changes ─────────────────────────────────────────────────────
_queued_all = st.session_state.get("queued_changes", {})
if _queued_all:
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    st.markdown(f'**{len(_queued_all)} change(s) queued for export**')
    _export_rows = [{"project_name": pn, **fields} for pn, fields in _queued_all.items()]
    _export_df = pd.DataFrame(_export_rows)
    st.dataframe(_export_df, use_container_width=True, hide_index=True)

    _buf = io.BytesIO()
    _export_df.to_csv(_buf, index=False)
    st.download_button(
        "⬇ Export changes to CSV",
        data=_buf.getvalue(),
        file_name=f"drs_changes_{date.today().strftime('%Y%m%d')}.csv",
        mime="text/csv",
        type="primary",
    )
    if st.button("Clear all queued changes", key="clear_all_changes"):
        st.session_state["queued_changes"] = {}
        st.rerun()

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — All Active Projects
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Active Projects</div>', unsafe_allow_html=True)

if active.empty:
    st.info("No active projects found.")
else:
    def _build_table_row(row):
        flags = row.get("_flags", [])
        flag_str = " ".join(
            ("🔴" if s=="error" else "🟡") for s, _, _m, _e in flags
        ) if flags else "✓"
        days_inac = int(row.get("days_inactive", -1))
        return {
            "Project":        str(row.get("project_name", "")),
            "Phase":          str(row.get("phase", "—")),
            "RAG":            str(row.get("rag", "") or "—").strip(),
            "Days Inactive":  days_inac if days_inac >= 0 else "—",
            "Go Live":        pd.Timestamp(row["go_live_date"]).strftime("%-d %b %Y")
                              if pd.notna(row.get("go_live_date")) else "—",
            "Schedule":       str(row.get("schedule_health", "") or "—").strip(),
            "Last Milestone": str(row.get("last_milestone", "—")),
            "Flags":          flag_str,
        }

    _tbl = pd.DataFrame([_build_table_row(r) for _, r in active.iterrows()])
    st.dataframe(
        _tbl,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Project":       st.column_config.TextColumn("Project",       width="large"),
            "Phase":         st.column_config.TextColumn("Phase",         width="medium"),
            "RAG":           st.column_config.TextColumn("RAG",           width="small"),
            "Days Inactive": st.column_config.TextColumn("Days Inactive", width="small"),
            "Go Live":       st.column_config.TextColumn("Go Live",       width="small"),
            "Schedule":      st.column_config.TextColumn("Schedule",      width="small"),
            "Last Milestone":st.column_config.TextColumn("Last Milestone",width="medium"),
            "Flags":         st.column_config.TextColumn("Flags",         width="small"),
        }
    )

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — On Hold
# ══════════════════════════════════════════════════════════════════════════════
with st.expander(f"On Hold ({len(on_hold)} projects)", expanded=False):
    if on_hold.empty:
        st.markdown("No on-hold projects.")
    else:
        _oh_tbl = pd.DataFrame([{
            "Project":    str(r.get("project_name", "")),
            "Phase":      str(r.get("phase", "—")),
            "RAG":        str(r.get("rag", "") or "—").strip(),
            "Go Live":    pd.Timestamp(r["go_live_date"]).strftime("%-d %b %Y")
                          if pd.notna(r.get("go_live_date")) else "—",
            "Days Inactive": int(r.get("days_inactive", -1)) if r.get("days_inactive", -1) >= 0 else "—",
        } for _, r in on_hold.iterrows()])
        st.dataframe(_oh_tbl, use_container_width=True, hide_index=True)

st.markdown(
    '<div style="font-size:11px;opacity:0.4;text-align:center;margin-top:20px">'
    'PS Reporting Tools · Internal use only · Data loaded this session only</div>',
    unsafe_allow_html=True
)
