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
    .section-label { font-size: 11px; font-weight: 700; text-transform: uppercase;
                     letter-spacing: 0.8px; color: #4472C4; margin-bottom: 8px; }
    .te-stat { font-size: 28px; font-weight: 700; color: inherit; }
    .te-stat-lbl { font-size: 12px; opacity: 0.6; margin-top: 2px; }
    .te-empty { text-align: center; padding: 48px; opacity: 0.4; font-size: 14px; }
</style>
""", unsafe_allow_html=True)

# ── Hero banner ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="te-hero">
    <div style="font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;
                color:#3B9EFF;margin-bottom:8px;font-family:Manrope,sans-serif">
        PROFESSIONAL SERVICES · MY WORK
    </div>
    <svg style='position:absolute;right:-40px;top:50%;transform:translateY(-50%);opacity:0.06;width:200px;height:200px;pointer-events:none' viewBox='0 0 1482 1286.25' xmlns='http://www.w3.org/2000/svg'><g fill='#3B9EFF' fill-rule='evenodd'><path d='M975.127,924.953c2.608-2.68,1.744-5.496-.42-7.829l-57.415-61.872c-2.463-2.655-5.025-2.878-8.443-.991-10.398,5.739-19.024,12.314-27.949,19.885-83.252,70.621-197.471,155.494-298.93,195.556-17.993,7.105-35.256,13.178-54.191,17.329-62.148,13.627-131.853,15.491-192.702-5.298-64.93-22.183-113.878-68.722-142.715-130.542-28.647-61.415-22.393-131.406,11.352-189.217,2.598-2.793,1.405-6.055-1.389-8.184-35.341-26.918-40.303-33.439-69.367-65.686-1.449-1.607-4.102-2.401-5.903-1.138-13.105,9.189-23.232,20.534-33.172,32.961-16.499,20.629-29.73,42.605-38.718,67.541-5.127,10.469-8.378,20.486-10.885,32.065-13.633,62.973-7.701,128.685,17.402,188.142,23.839,56.463,65.297,103.638,114.77,139.169,32.418,23.283,66.848,42.548,103.476,58.385,25.142,10.871,50.281,18.994,76.934,25.12,96.392,22.153,188.876,4.496,276.774-38.393,42.916-20.94,83.188-45.685,121.922-73.568,75.733-54.514,154.643-126.72,219.571-193.435ZM1445.252,792.261c-7.628-38.507-22.817-74.472-43.124-107.897-35.582-58.566-85.801-106.77-139.329-149.092-69.784-55.176-145.355-102.407-225.163-141.162-2.165-1.052-4.941.388-5.391,1.627-.426,1.171-.463,3.413.931,4.628,20.341,17.734,39.847,35.55,58.599,55.093,13.286,14.465,26.223,28.012,37.022,44.544,19.784,30.289,35.735,62.168,50.127,95.397,34.512,31.926,64.863,67.358,90.813,106.359,42.427,63.765,57.696,142.663,37.453,217.116-11.436,42.061-34.763,80.507-64.388,112.265-55.859,59.882-133.144,94.711-214.71,99.157-32.507,1.773-64.093-.538-96.013-6.503-28.16-5.262-70.299-23.997-96.538-36.626-2.312-1.112-4.605-.743-6.449.974-12.635,11.76-25.076,22.901-39.051,33.146l-43.32,31.757c-2.68,1.965-2.195,5.562.439,7.808,70.707,60.309,165.779,100.179,259.837,97.033,39.996-1.336,78.686-6.594,117.486-16.111,94.178-23.099,174.952-71.91,236.526-146.957,23.873-29.096,44.355-60.51,59.779-94.956,29.172-65.148,38.357-137.461,24.463-207.601ZM601.099,242.903c-12.268,10.522-48.215,44.405-47.219,60.482.993,16.01,10.781,31.195,25.227,38.155,14.47,6.972,41.303-10.055,53.886-18.311l65.495-42.972c26.305-17.259,52.496-32.716,80.08-47.834l57.464-31.494c20.451-11.209,41.123-19.851,63.235-27.448,35.852-12.318,72.313-18.084,110.322-17.747,29.787.263,58.398,3.408,86.939,11.449,44.037,12.405,82.745,35.987,114.027,69.974,20.347,22.106,37.598,45.332,51.026,71.732,6.962,13.688,13.008,27.156,16.103,42.311,6.48,31.729,12.267,85.992-.676,115.916-6.013,13.902-13.009,26.627-18.289,40.753-.847,2.264-.768,4.767,1.387,6.461l81.366,63.967c2.003,1.574,5.098.298,6.46-1.592,19.285-26.745,34.599-55.578,45.667-86.804,10.617-29.953,15.416-60.246,15.218-92.192-.482-77.938-29.055-152.791-79.976-211.891-67.16-77.946-169.264-137.487-272.877-146.244-33.524-2.834-66.192-1.328-99.421,3.091-82.214,10.934-149.21,45.218-216.385,92.267-48.269,33.807-94.373,69.644-139.062,107.973ZM72.687,567.553c20.03,44.974,54.35,86.652,88.718,121.568,19.447,19.756,38.882,38.258,60.393,55.711l73.052,59.268c30.921,25.086,74.954,56.331,111.096,72.278,11.713,5.168,23.385,8.99,35.917,11.295,12.922,2.375,24.878,1.136,37.309-3.088,18.441-6.266,35.538-14.698,52.671-24.006,1.792-.974,2.85-2.213,3.058-3.936.179-1.483-.47-3.163-1.914-4.548-14.129-13.542-27.174-27.284-42.195-40.056l-78.193-66.48-93.5-82.422c-23.176-20.43-44.471-41.737-65.536-64.239-15.19-16.227-28.591-32.64-40.05-51.639-20.601-34.157-31.396-72.282-30.182-112.398.614-20.279,2.364-39.861,7.45-59.369,8.872-34.031,50.72-76.652,77.451-99.125,3.767-7.04,2.459-14.401,2.885-21.735.884-15.227,3.244-29.908,5.647-44.959,4.285-26.824,22.718-58.984,38.899-80.638,1.348-1.805,1.936-3.535.891-4.937-.951-1.277-2.618-2.49-4.589-2.222-52.436,7.145-104.92,34.806-146.088,67.704-25.632,20.484-48.458,43.456-68.934,69.137-46.339,58.118-62.952,131.49-53.428,204.864,4.697,36.186,14.376,70.75,29.171,103.971ZM1196.886,310.029c-4.882-10.39-12.371-18.773-20.659-26.723-18.771-18.007-40.425-31.674-64.291-42.362-57.569-25.783-110.906-28.064-173.214-22.213-61.067,5.735-111.183,25.069-164.567,54.081-24.678,13.412-48.301,26.866-71.885,42.28l-105.247,68.787c-85.308,55.756-195.138,156.138-256.755,237.876-1.598,2.12-2.206,4.81-.222,6.912l76.342,80.886c1.468,1.556,2.9,1.672,4.715,1.249,1.397-.326,1.99-1.717,2.793-3.377,3.117-6.44,6.665-11.977,11.238-17.864,38.52-49.59,82.099-94.54,130.222-135.261,40.87-34.583,82.783-67.442,126.68-98.902,83.71-59.991,188.529-115.793,291.15-127.921,23.653-2.795,46.328-.575,69.656,3.405,27.197,4.641,52.661,12.543,78.69,21.347l38.004,12.855c13.849,4.685,27.221-3.226,30.503-17.755,2.725-12.064,2.293-25.708-3.154-37.301Z'/></g></svg>
    <h1 style="color:#fff;margin:0;font-size:28px;font-weight:800;font-family:Manrope,sans-serif">
        Time Entries
    </h1>
    <p style="color:rgba(255,255,255,0.6);margin:8px 0 0;font-size:14px;
              font-family:Manrope,sans-serif;max-width:520px">
        Log project activity during your session. Review, adjust, and export
        draft time entries ready for NetSuite import.
    </p>
</div>
""", unsafe_allow_html=True)

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
