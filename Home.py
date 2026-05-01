"""
PS Platform & Tools — Entry point
Auth, sidebar identity + upload hub, multi-page navigation.
"""
import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="PS Platform — Zone & Co",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="expanded",
)

from auth import require_auth, sign_out
from components import inject_css, init_filters
from config import get_role, ACTIVE_EMPLOYEES

inject_css()

# ── Auth gate ─────────────────────────────────────────────────────────────────
if not require_auth():
    st.stop()

# ── Session setup ──────────────────────────────────────────────────────────────
_roster  = st.session_state.get("consultant_name", "")
_display = st.session_state.get("display_name", "")
_role    = get_role(_roster) if _roster else "consultant"
st.session_state["_role"] = _role

init_filters()

# Force cache bust when loader version changes
_LOADER_VERSION = "v20260422b"
_drs_vkey = f"df_drs_{_LOADER_VERSION}"
_ns_vkey  = f"df_ns_{_LOADER_VERSION}"
if _drs_vkey not in st.session_state:
    for k in [k for k in list(st.session_state.keys())
              if k in ("df_drs",) or (k.startswith("df_drs_v") and k != _drs_vkey)]:
        del st.session_state[k]
    st.session_state[_drs_vkey] = True
if _ns_vkey not in st.session_state:
    for k in [k for k in list(st.session_state.keys())
              if k in ("df_ns",) or (k.startswith("df_ns_v") and k != _ns_vkey)]:
        del st.session_state[k]
    st.session_state[_ns_vkey] = True

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    # Brand header
    st.markdown(
        '<div style="padding:8px 0 12px">'
        '<div style="font-weight:800;font-size:17px;color:#0E223D;letter-spacing:-0.01em">'
        'Zone &amp; Co</div>'
        '<div style="font-size:11px;color:#6B7A8F;letter-spacing:.08em;text-transform:uppercase">'
        'PS Platform</div>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.divider()

    # User identity
    st.caption(f"Signed in as **{_display}**")
    st.caption(_role.replace("_", " ").title())
    if st.button("Sign out", key="signout"):
        sign_out()

    st.divider()

    # ── Upload hub ────────────────────────────────────────────────────────────
    st.markdown("**Upload data**")
    st.caption("Upload once — available across all pages this session.")

    # SS DRS — always first
    from shared.smartsheet_api import ss_available, load_sheet_as_df as _ss_load
    _ss_ready = ss_available()
    st.markdown(
        '<span class="upload-label">SS DRS</span>',
        unsafe_allow_html=True,
    )
    if st.button(
        "⟳ Sync SS DRS data",
        key="hub_drs_api",
        use_container_width=True,
        disabled=not _ss_ready,
        help="Fetch live DRS data from Smartsheet API" if _ss_ready
             else "SMARTSHEET_TOKEN / SMARTSHEET_DRS_ID not configured",
    ):
        with st.spinner("Syncing from Smartsheet…"):
            try:
                st.session_state["df_drs"] = _ss_load()
                st.session_state["_drs_source"] = "api"
                st.toast("DRS synced from Smartsheet ✓")
            except Exception as e:
                st.error(f"Smartsheet error: {e}")

    # NS Time Detail
    st.markdown(
        '<a class="upload-link" href="https://3838224.app.netsuite.com/app/common/search/'
        'searchresults.nl?searchid=66732&saverun=T&whence=" target="_blank">'
        '↗ Open NS Time Detail Search</a>'
        '<span class="upload-label">NS Time Detail</span>',
        unsafe_allow_html=True,
    )
    ns_file = st.file_uploader(
        "NS Time Detail", type=["xlsx", "csv"],
        key="hub_ns", label_visibility="collapsed",
    )

    # SFDC Contacts
    st.markdown(
        '<a class="upload-link" href="https://drive.google.com/drive/u/1/folders/'
        '1VdI_WjuVclF5xN9fG7dEIz1WDu4QRE0m" target="_blank">'
        '↗ Open SFDC Contacts (Google Drive)</a>'
        '<span class="upload-label">SFDC Contacts</span>',
        unsafe_allow_html=True,
    )
    sfdc_file = st.file_uploader(
        "SFDC Contacts", type=["xlsx", "csv"],
        key="hub_sfdc", label_visibility="collapsed",
    )

    # Manager-only uploads
    ns_ua_file = rev_file = tm_sow_file = None
    if _role in ("manager", "manager_only", "reporting_only"):
        st.divider()
        st.caption("Manager · reporting only")

        st.markdown(
            '<a class="upload-link" href="https://3838224.app.netsuite.com/app/common/search/'
            'searchresults.nl?searchid=68439&whence=" target="_blank">'
            '↗ Open NS Unassigned Projects</a>'
            '<span class="upload-label">NS Unassigned Projects</span>',
            unsafe_allow_html=True,
        )
        ns_ua_file = st.file_uploader(
            "NS Unassigned Projects", type=["xlsx", "csv"],
            key="hub_ns_unassigned", label_visibility="collapsed",
            help="Required for Capacity Outlook",
        )

        st.markdown(
            '<a class="upload-link" href="https://3838224.app.netsuite.com/app/common/search/'
            'searchresults.nl?searchid=75183&whence=" target="_blank">'
            '↗ Open NS FF Revenue Charges</a>'
            '<span class="upload-label">NS FF Revenue Charges</span>',
            unsafe_allow_html=True,
        )
        rev_file = st.file_uploader(
            "NS FF Revenue Charges", type=["xlsx", "csv"],
            key="hub_revenue", label_visibility="collapsed",
            help="Required for Revenue Report",
        )

        st.markdown(
            '<a class="upload-link" href="https://zoneandco.lightning.force.com/lightning/page/'
            'analytics?wave__assetType=report&wave__assetId=00OUh00000PeTZZMA3" target="_blank">'
            '↗ Open SFDC T&amp;M SOW Report</a>'
            '<span class="upload-label">SFDC T&amp;M SOW</span>',
            unsafe_allow_html=True,
        )
        tm_sow_file = st.file_uploader(
            "SFDC T&M SOW", type=["xlsx", "csv"],
            key="hub_tm_sow", label_visibility="collapsed",
            help="Required for T&M Revenue Report",
        )

    # ── Process uploads ───────────────────────────────────────────────────────
    from shared.loaders import load_drs, load_ns_time, load_sfdc
    _drs_file = None  # DRS comes from API only in v2

    for label, session_key, loader_fn, file_obj in [
        ("NS Time",     "df_ns",            load_ns_time, ns_file),
        ("SFDC",        "df_sfdc",          load_sfdc,    sfdc_file),
        ("NS Unassigned","df_ns_unassigned", load_ns_time, ns_ua_file),
        ("FF Revenue",  "df_revenue",       load_ns_time, rev_file),
        ("T&M SOW",     "df_tm_sow",        load_sfdc,    tm_sow_file),
    ]:
        if file_obj is not None and session_key not in st.session_state:
            with st.spinner(f"Loading {label}…"):
                try:
                    st.session_state[session_key] = loader_fn(file_obj)
                    st.toast(f"{label} loaded ✓")
                except Exception as e:
                    st.error(f"{label} error: {e}")

    st.divider()

    # ── Loaded data status ────────────────────────────────────────────────────
    _status_items = [
        ("df_drs",           "SS DRS"),
        ("df_ns",            "NS Time"),
        ("df_sfdc",          "SFDC"),
        ("df_ns_unassigned", "NS Unassigned"),
        ("df_revenue",       "FF Revenue"),
        ("df_tm_sow",        "T&M SOW"),
    ]
    _any_loaded = any(k in st.session_state for k, _ in _status_items)
    if _any_loaded:
        for key, label in _status_items:
            if key in st.session_state:
                df = st.session_state[key]
                n = len(df) if hasattr(df, "__len__") else "?"
                st.markdown(
                    f'<span style="font-size:12px;color:#1F9D66">✓ {label}</span>'
                    f'<span style="font-size:11px;color:#6B7A8F"> · {n} rows</span>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<span style="font-size:12px;color:#B4B2A9">○ {label}</span>',
                    unsafe_allow_html=True,
                )
        if st.button("Clear loaded data", use_container_width=True, key="clear_data"):
            for key, _ in _status_items:
                st.session_state.pop(key, None)
            st.rerun()

# ── Navigation ────────────────────────────────────────────────────────────────
_is_manager = _role in ("manager", "manager_only", "reporting_only")

_consultant_pages = [
    st.Page("pages/1_briefing.py",    title="Daily Briefing",       icon="☀️",  default=True),
    st.Page("pages/2_projects.py",    title="My Projects",          icon="📁"),
    st.Page("pages/3_engagement.py",  title="Customer Engagement",  icon="✉️"),
    st.Page("pages/4_profile.py",     title="Customer Profile",     icon="🏢"),
    st.Page("pages/5_utilization.py", title="Utilization",          icon="📈"),
    st.Page("pages/6_drs.py",         title="DRS Health Check",     icon="🛡️"),
    st.Page("pages/8_time.py",        title="Time Entries",         icon="⏱️"),
]

_manager_pages = [
    st.Page("pages/7_reporting.py",  title="Project Health",       icon="📊"),
    st.Page("pages/9_portfolio.py",  title="Portfolio Analytics",  icon="🗂️"),
    st.Page("pages/10_whs.py",       title="Workload Health",      icon="⚖️"),
    st.Page("pages/11_capacity.py",  title="Capacity Outlook",     icon="🔭"),
    st.Page("pages/12_revenue.py",   title="Revenue Report",       icon="💰"),
]

_other_pages = [
    st.Page("pages/13_vibe.py", title="Vibe Check", icon="🎯"),
    st.Page("pages/14_help.py", title="Help",        icon="❓"),
]

pages = _consultant_pages
if _is_manager:
    pages = pages + _manager_pages
pages = pages + _other_pages

pg = st.navigation(pages)
pg.run()
