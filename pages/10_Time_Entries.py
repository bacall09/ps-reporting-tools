"""
Zone PS — Time Entries
Session-based activity log → draft NS time entries → CSV export.
"""
import streamlit as st
import pandas as pd
from datetime import date, datetime
import io

st.session_state["current_page"] = "Time Entries"

from shared.activity_log import (
    log_activity, get_log_df, clear_log, log_count,
    to_ns_export, ACTIVITY_TYPE_LIST, ACTIVITY_TYPES, NS_HOUR_INCREMENTS
)
from shared.constants import EMPLOYEE_ROLES

# ── Auth ──────────────────────────────────────────────────────────────────────
_session_name = st.session_state.get("consultant_name", "")
if not _session_name:
    st.warning("Sign in on the Home page to use Time Entries.")
    st.stop()

# ── Page styles ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .te-hero { background:#050D1F; padding:32px 40px 28px; border-radius:10px; margin-bottom:24px; position:relative; overflow:hidden; font-family:Manrope,sans-serif; }
    .section-label { font-size: 13px; font-weight: 700; text-transform: uppercase;
                     letter-spacing: 0.8px; color: #4472C4; margin-bottom: 8px; }
    .te-stat { font-size: 28px; font-weight: 700; color: inherit; }
    .te-stat-lbl { font-size: 12px; opacity: 0.6; margin-top: 2px; }
    .te-empty { text-align: center; padding: 48px; opacity: 0.4; font-size: 14px; }
</style>
""", unsafe_allow_html=True)

# ── Hero banner ───────────────────────────────────────────────────────────────
_hero = st.empty()
_hero.markdown("<div class='te-hero'><div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3B9EFF;margin-bottom:8px;font-family:Manrope,sans-serif'>Professional Services · My Work</div><h1 style='color:#fff;margin:0;font-size:28px;font-weight:800;font-family:Manrope,sans-serif'>Time Entries</h1><p style='color:rgba(255,255,255,0.6);margin:8px 0 0;font-size:14px;font-family:Manrope,sans-serif;max-width:520px'>Log project activity during your session. Review, adjust, and export draft time entries ready for NetSuite import.</p></div>", unsafe_allow_html=True)

# ── Summary metrics ───────────────────────────────────────────────────────────
log_df = get_log_df()
_total_entries = len(log_df)
_total_hours   = float(log_df["hours"].sum()) if not log_df.empty else 0.0
_projects      = log_df["project_id"].nunique() if not log_df.empty else 0

m1, m2, m3 = st.columns(3)
with m1:
    st.markdown(f'<div class="te-stat">{_total_entries}</div>'
                f'<div class="te-stat-lbl">Draft entries this session</div>',
                unsafe_allow_html=True)
with m2:
    st.markdown(f'<div class="te-stat">{_total_hours:.2f}h</div>'
                f'<div class="te-stat-lbl">Total hours logged</div>',
                unsafe_allow_html=True)
with m3:
    st.markdown(f'<div class="te-stat">{_projects}</div>'
                f'<div class="te-stat-lbl">Projects this session</div>',
                unsafe_allow_html=True)

st.markdown('<hr style="margin:20px 0;opacity:0.15">', unsafe_allow_html=True)

# ── Manual entry form ─────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Add Entry</div>', unsafe_allow_html=True)

# Build project dropdown from DRS session data filtered to this consultant
_df_drs_te  = st.session_state.get("df_drs")
_drs_loaded = _df_drs_te is not None and not _df_drs_te.empty
_proj_options = {}  # {display_name: project_id}

if _drs_loaded:
    from shared.constants import name_matches
    # Try filtering to this consultant's projects
    if "project_manager" in _df_drs_te.columns:
        _my_drs = _df_drs_te[_df_drs_te["project_manager"].apply(
            lambda v: name_matches(v, _session_name)
        )]
        # If name match returns nothing, show all as fallback with a note
        if _my_drs.empty:
            _my_drs = _df_drs_te
            _proj_match_warn = True
        else:
            _proj_match_warn = False
    else:
        _my_drs = _df_drs_te
        _proj_match_warn = False

    if "project_name" in _my_drs.columns:
        for _, _pr in _my_drs.iterrows():
            _pname = str(_pr.get("project_name","") or "").strip()
            _pid   = str(_pr.get("project_id","") or "").strip()
            if _pname and _pname not in ("nan","None"):
                _proj_options[_pname] = _pid
else:
    _proj_match_warn = False

_proj_names = sorted(_proj_options.keys())

if not _drs_loaded:
    st.info("Load SS DRS on the Home page to enable the project dropdown.", icon="ℹ️")

with st.form("te_add_form", clear_on_submit=True):
    fc1, fc2, fc3 = st.columns([2, 2, 1])
    with fc1:
        if _proj_names:
            _f_proj_name = st.selectbox("Project *", ["— Select project —"] + _proj_names)
            _f_proj_id   = _proj_options.get(_f_proj_name, "") if _f_proj_name != "— Select project —" else ""
            if _f_proj_id:
                st.caption(f"Project ID: {_f_proj_id}")
            if _proj_match_warn:
                st.caption("⚠️ Showing all projects — name match to your profile not found")
        else:
            _f_proj_name = st.text_input("Project Name *", placeholder="e.g. Acme Corp - ZCapture Implementation")
            _f_proj_id   = st.text_input("Project ID *", placeholder="e.g. 157425")
    with fc2:
        _f_activity = st.selectbox("Activity Type *", ACTIVITY_TYPE_LIST)
        _f_date     = st.date_input("Date", value=date.today())
    with fc3:
        _f_hrs = st.number_input("Hours *", min_value=0.25, max_value=24.0,
                                  value=ACTIVITY_TYPES.get(_f_activity, {}).get("default_hrs", 0.25),
                                  step=0.25, format="%.2f")
        _f_emp = st.text_input("Employee", value=_session_name, disabled=True)

    _f_notes = st.text_input("Notes (optional)", placeholder="Additional context for this entry")

    _submitted = st.form_submit_button("＋ Add Entry", type="primary", use_container_width=True)
    if _submitted:
        _proj_id_val   = _f_proj_id.strip() if isinstance(_f_proj_id, str) else str(_f_proj_id)
        # project_name: use the display name from dropdown, never the ID
        _proj_name_val = (_f_proj_name if _f_proj_name != "— Select project —" else "") if _proj_names else _f_proj_name
        # Guard: if name accidentally equals the id, clear it so it's re-resolved
        if _proj_name_val and _proj_name_val.strip() == _proj_id_val.strip():
            _proj_name_val = ""
        if not _proj_id_val and not _proj_name_val:
            st.error("Please select a project.")
        elif not _proj_id_val:
            st.error("Project ID could not be resolved — check DRS data is loaded on Home page.")
        else:
            log_activity(
                project_id    = _proj_id_val,
                project_name  = _proj_name_val,
                activity_type = _f_activity,
                employee      = _session_name,
                notes         = _f_notes.strip(),
                entry_date    = _f_date,
            )
            _log = st.session_state.get("activity_log", [])
            if _log:
                _log[-1]["hours"] = _f_hrs
                _log[-1]["memo"]  = f"{ACTIVITY_TYPES[_f_activity]['ns_memo']}" + (f" — {_f_notes.strip()}" if _f_notes.strip() else "")
                st.session_state["activity_log"] = _log
            st.rerun()

st.markdown('<hr style="margin:20px 0;opacity:0.15">', unsafe_allow_html=True)

# ── Draft entries table ───────────────────────────────────────────────────────
st.markdown('<div class="section-label">Draft Entries — Review & Edit</div>',
            unsafe_allow_html=True)

log_df = get_log_df()  # refresh after possible add

if log_df.empty:
    st.markdown('<div class="te-empty">No entries yet this session.<br>'
                'Add entries manually above, or they are captured automatically '
                'when you use Customer Engagement.</div>', unsafe_allow_html=True)
else:
    # Editable table
    _edit_cols = ["date","project_id","project_name","activity_type","hours","memo","employee"]
    _display   = log_df[_edit_cols].copy()

    # Ensure types are compatible with column config
    _display["date"]  = pd.to_datetime(_display["date"], errors="coerce").dt.date
    _display["hours"] = pd.to_numeric(_display["hours"], errors="coerce").fillna(0.25)

    _edited = st.data_editor(
        _display,
        column_config={
            "date":          st.column_config.DateColumn("Date",         width="small"),
            "project_id":    st.column_config.TextColumn("Project ID",   width="small"),
            "project_name":  st.column_config.TextColumn("Project Name", width="medium"),
            "activity_type": st.column_config.SelectboxColumn(
                                 "Activity Type", options=ACTIVITY_TYPE_LIST, width="medium"),
            "hours":         st.column_config.NumberColumn(
                                 "Hours", min_value=0.25, max_value=24.0,
                                 step=0.25, format="%.2f", width="small"),
            "memo":          st.column_config.TextColumn("NS Memo",      width="large"),
            "employee":      st.column_config.TextColumn("Employee",     disabled=True, width="medium"),
        },
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="te_edit_table",
    )

    # Sync edits back to session state
    if _edited is not None:
        _updated = log_df.copy()
        for col in _edit_cols:
            if col in _edited.columns and col != "employee":
                _updated[col] = _edited[col].values
        st.session_state["activity_log"] = _updated.to_dict("records")

    # ── Summary by project ────────────────────────────────────────────────
    st.markdown('<div class="section-label" style="margin-top:16px">Summary by Project</div>',
                unsafe_allow_html=True)
    _summary = (log_df.groupby(["project_id","project_name"])
                .agg(Entries=("id","count"), Hours=("hours","sum"))
                .reset_index()
                .rename(columns={"project_id":"Project ID","project_name":"Project Name"}))
    _summary["Hours"] = _summary["Hours"].apply(lambda h: f"{h:.2f}h")
    st.dataframe(_summary, use_container_width=True, hide_index=True)

    st.markdown('<hr style="margin:20px 0;opacity:0.15">', unsafe_allow_html=True)

    # ── Export ────────────────────────────────────────────────────────────
    st.markdown('<div class="section-label">Export to NetSuite</div>',
                unsafe_allow_html=True)
    st.caption("Export as CSV formatted for NetSuite Time Entry import. "
               "Review all entries before importing.")

    _export_df = to_ns_export(log_df)
    _buf = io.BytesIO()
    _export_df.to_csv(_buf, index=False)

    ex1, ex2, ex3 = st.columns([2, 1, 1])
    with ex1:
        st.download_button(
            label=f"⬇ Export {len(log_df)} entries to CSV",
            data=_buf.getvalue(),
            file_name=f"time_entries_{_session_name.replace(', ','_')}_{date.today().isoformat()}.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True,
        )
    with ex3:
        if st.button("🗑 Clear all entries", use_container_width=True):
            clear_log()
            st.rerun()

st.markdown('<hr style="margin:20px 0;opacity:0.15">', unsafe_allow_html=True)

# ── How it works ──────────────────────────────────────────────────────────────
with st.expander("ℹ️ How Time Entries works"):
    st.markdown("""
**Automatic capture** — When you draft a customer email on the Customer Engagement page,
an entry is automatically added here with the project ID and a default duration.

**Manual entry** — Add any project activity using the form above. Select the project ID,
activity type, and adjust the hours to the nearest 0.25h increment.

**Edit before export** — All entries are drafts. Edit the hours, memo, or date in the
table above before exporting.

**NetSuite import** — The exported CSV is formatted for NS Time Entry bulk import with
columns: Date, Project ID, Hours, Memo, Employee.

**Session only** — Entries are stored in your current browser session and cleared when
you close the tab. Export before closing.
    """)
