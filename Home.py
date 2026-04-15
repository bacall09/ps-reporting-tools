"""
PS Tools — Home (entrypoint)
Handles auth, navigation, and upload hub.
With st.navigation in Streamlit 1.32+, Home registers pages and calls pg.run().
The login gate lives in a shared helper called by every page.
"""
import streamlit as st
import pandas as pd
import streamlit_authenticator as stauth

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN, ACTIVE_EMPLOYEES,
    get_role, is_manager, LEAVER_EXIT_DATES,
)
from shared.loaders import load_drs, load_ns_time, load_sfdc

# Increment this when loaders change — forces session cache to invalidate
_LOADER_VERSION = "v20260409e"

st.set_page_config(page_title="PS Tools", page_icon=None, layout="wide")

# ── Build credentials ─────────────────────────────────────────────────────────
def _to_dict(obj):
    try:
        d = dict(obj)
        return {k: _to_dict(v) for k, v in d.items()}
    except (TypeError, ValueError):
        return obj

_secrets_creds = st.secrets.get("credentials", {})
_creds = {
    "usernames": {
        u: _to_dict(d)
        for u, d in _to_dict(_secrets_creds.get("usernames", {})).items()
    }
}
_cookie_raw = _to_dict(st.secrets.get("cookie", {}))
_cookie = _cookie_raw or {"name": "ps_tools_auth", "key": "fallback_key", "expiry_days": 30}

authenticator = stauth.Authenticate(
    credentials        = _creds,
    cookie_name        = _cookie.get("name", "ps_tools_auth"),
    cookie_key         = _cookie.get("key", "fallback_key"),
    cookie_expiry_days = int(_cookie.get("expiry_days", 30)),
)

# ── Auth state ────────────────────────────────────────────────────────────────
_auth_user   = st.session_state.get("username", "")
_auth_name   = st.session_state.get("name", "")
_roster      = _creds["usernames"].get(_auth_user, {}).get("full_roster_name", "")
_role        = get_role(_roster) if _roster else None

if _roster and st.session_state.get("consultant_name") != _roster:
    st.session_state["consultant_name"] = _roster

# ── Register navigation ───────────────────────────────────────────────────────
_consultant_pages = [
    st.Page("pages/1_Daily_Briefing.py",        title="Daily Briefing"),
    st.Page("pages/8_My_Projects.py",            title="My Projects"),
    st.Page("pages/2_Customer_Reengagement.py", title="Customer Engagement"),
    st.Page("pages/3_Utilization_Report.py",    title="Utilization Report"),
    st.Page("pages/4_Workload_Health_Score.py", title="Workload Health Score"),
    st.Page("pages/6_DRS_Health_Check.py",      title="DRS Health Check"),
    # st.Page("pages/7_Vibe_Check.py",          title="Vibe Check ✨"),  # hidden
]
_manager_pages = [
    st.Page("pages/5_Capacity_Outlook.py",  title="Capacity Outlook"),
    st.Page("pages/9_Revenue_Report.py",    title="Revenue Report"),
]
_help_pages = [
    st.Page("pages/9_Help.py", title="Help"),
]

if _role in ("manager", "manager_only"):
    pg = st.navigation({"My Tools": _consultant_pages, "Management": _manager_pages, "Info": _help_pages})
elif _role == "reporting_only":
    pg = st.navigation({"Management": _manager_pages})
else:
    pg = st.navigation({"My Tools": _consultant_pages, "Info": _help_pages})

# ── Login gate (shown instead of page content when not authenticated) ─────────
if not st.session_state.get("authentication_status"):
    _user_options  = {d.get("name", u): u for u, d in _creds["usernames"].items()}
    _display_names = sorted(_user_options.keys())

    st.markdown("<div style='margin-top:60px'></div>", unsafe_allow_html=True)
    _col = st.columns([1, 2, 1])[1]
    with _col:
        st.markdown("""
        <div style='background:#1e2c63;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px'>
            <div style='font-size:11px;color:#a0aec0;letter-spacing:2px;
                        text-transform:uppercase;margin-bottom:8px'>Professional Services</div>
            <h1 style='color:#fff;margin:0;font-size:28px;font-weight:700'>PS Reporting Tools</h1>
            <p style='color:#a0aec0;margin:10px 0 0;font-size:13px'>Sign in to continue.</p>
        </div>""", unsafe_allow_html=True)

        _sel = st.selectbox("Select your name", ["— Select —"] + _display_names, key="login_name")
        _pw  = st.text_input("Password", type="password",
                             placeholder="Enter your password", key="login_pw")
        if st.button("Sign in →", use_container_width=True, type="primary", key="login_btn"):
            if _sel == "— Select —":
                st.warning("Please select your name.")
            elif not _pw:
                st.warning("Please enter your password.")
            else:
                import bcrypt as _bc
                _uname = _user_options[_sel]
                _hash  = _creds["usernames"][_uname].get("password", "")
                try:    _ok = _bc.checkpw(_pw.encode(), _hash.encode())
                except: _ok = False
                if _ok:
                    _r = _creds["usernames"][_uname].get("full_roster_name", "")
                    st.session_state["authentication_status"] = True
                    st.session_state["username"]              = _uname
                    st.session_state["name"]                  = _sel
                    if _r: st.session_state["consultant_name"] = _r
                    st.rerun()
                else:
                    st.error("Incorrect username or password.")

    with st.expander("🔑 Need to reset your password?"):
        st.caption("Generate a new hash and send it to your admin.")
        _rc = st.columns([1, 2, 1])[1]
        with _rc:
            _rn  = st.selectbox("Your name", ["— Select —"] + _display_names, key="reset_name")
            _rp1 = st.text_input("New password", type="password", key="reset_pw1")
            _rp2 = st.text_input("Confirm password", type="password", key="reset_pw2")
            if st.button("Generate new hash", use_container_width=True, key="reset_btn"):
                if _rn == "— Select —": st.warning("Select your name.")
                elif len(_rp1) < 8:     st.warning("Password must be at least 8 characters.")
                elif _rp1 != _rp2:      st.error("Passwords don't match.")
                else:
                    import bcrypt as _bc
                    _nh = _bc.hashpw(_rp1.encode(), _bc.gensalt()).decode()
                    st.success("Hash generated! Send the below to your admin:")
                    st.code(f'[credentials.usernames.{_user_options[_rn]}]\npassword = "{_nh}"',
                            language="toml")
    st.stop()

# ── Authenticated: sidebar ────────────────────────────────────────────────────
_display_first = _roster.split(",")[1].strip() if "," in _roster else _roster
_upload_role   = get_role(_roster) if _roster else None

with st.sidebar:
    st.markdown(f"#### {_display_first}")
    st.caption(f"Signed in as **{_auth_name}**")
    if st.button("Sign out", key="home_signout"):
        for k in ["authentication_status","username","name","consultant_name",
                  "df_drs","df_ns","df_sfdc","df_ns_unassigned"]:
            st.session_state.pop(k, None)
        st.rerun()
    st.markdown("---")

    # ── View As + Filter (managers only) ──────────────────────────────────────
    if _upload_role == "manager":
        from shared.config import EMPLOYEE_LOCATION as _EL, PS_REGION_MAP as _RM, PS_REGION_OVERRIDE as _RO

        def _get_region(name):
            if name in _RO: return _RO[name]
            return _RM.get(_EL.get(name, ""), "Other")

        _active_c = sorted([
            n for n in CONSULTANT_DROPDOWN
            if get_role(n) in ("consultant", "manager")
            and len(EMPLOYEE_ROLES.get(n, {}).get("products", [])) > 0
        ])
        _by_rgn = {}
        for _cn in _active_c:
            _by_rgn.setdefault(_get_region(_cn), []).append(_cn)

        _mgr_only = sorted([n for n in ACTIVE_EMPLOYEES if get_role(n) == "manager_only"])
        _bopts = ["— My own view —", "👥 All team"]
        for _rg in sorted(_by_rgn.keys()):
            _bopts.append(f"── {_rg} ──")
            _bopts.extend(_by_rgn[_rg])
        if _mgr_only:
            _bopts.append("── Managers ──")
            _bopts.extend(_mgr_only)

        st.markdown("**View as:**")
        _browse = st.selectbox("Browse team", _bopts,
                               key="home_browse", label_visibility="collapsed")

        _all_prods = sorted({
            p for n in _active_c
            for p in EMPLOYEE_ROLES.get(n, {}).get("products", []) if p and p != "All"
        })
        st.markdown("**Filter by product:**")
        _product_filter_home = st.selectbox("Product", ["All products"] + _all_prods,
                                            key="home_product_filter", label_visibility="collapsed")

        # Store in session state so Daily Briefing can read them
        st.session_state["_view_browse"]   = _browse
        st.session_state["_product_filter"] = _product_filter_home
        st.markdown("---")

    st.markdown("**Upload data**")
    st.caption("Upload once — available across all pages this session.")

    drs_file  = st.file_uploader("SS DRS Export",  type=["xlsx","csv"], key="hub_drs")
    st.markdown('<a href="https://www.smartsheet.com" target="_blank" style="font-size:11px;opacity:0.6;">↗ Open SS DRS Report</a>', unsafe_allow_html=True)

    ns_file   = st.file_uploader("NS Time Detail", type=["xlsx","csv"], key="hub_ns")
    st.markdown('<a href="https://3838224.app.netsuite.com/app/common/search/searchresults.nl?searchid=66732&amp;saverun=T&amp;whence=" target="_blank" style="font-size:11px;opacity:0.6;">↗ Open NS Time Detail Search</a>', unsafe_allow_html=True)

    sfdc_file = st.file_uploader("SFDC Contacts",  type=["xlsx","csv"], key="hub_sfdc")
    st.markdown('<a href="https://drive.google.com/drive/u/1/folders/1VdI_WjuVclF5xN9fG7dEIz1WDu4QRE0m" target="_blank" style="font-size:11px;opacity:0.6;">↗ Open SFDC Contacts (Google Drive)</a>', unsafe_allow_html=True)

    ns_ua_file = (
        st.file_uploader("NS Unassigned Projects", type=["xlsx","csv"], key="hub_ns_unassigned",
                         help="Required for Capacity Outlook")
        if _upload_role in ("manager","manager_only","reporting_only") else None
    )
    if _upload_role in ("manager","manager_only","reporting_only"):
        st.markdown('<a href="https://3838224.app.netsuite.com/app/common/search/searchresults.nl?searchid=68439&whence=" target="_blank" style="font-size:11px;opacity:0.6;">↗ Open NS Unassigned Projects</a>', unsafe_allow_html=True)

    rev_file = (
        st.file_uploader("NS FF Revenue Charges", type=["xlsx","csv"], key="hub_revenue",
                         help="Required for Revenue Report")
        if _upload_role in ("manager","manager_only","reporting_only") else None
    )
    tm_sow_file = (
        st.file_uploader("SFDC T&M SOW", type=["xlsx","csv"], key="hub_tm_sow",
                         help="Required for T&M Revenue Report")
        if _upload_role in ("manager","manager_only","reporting_only") else None
    )
    # Clear stale versioned caches on new deploy
    for _vk in [f"df_ns_{_LOADER_VERSION}", f"df_drs_{_LOADER_VERSION}"]:
        if _vk not in st.session_state:
            for _ok in [k for k in list(st.session_state.keys())
                        if k.startswith(_vk.split("_v")[0]) and k != _vk.split("_v")[0]
                        and "_v20" in k]:
                del st.session_state[_ok]

    for _lbl, _key, _ldr, _f in [
        ("SS DRS","df_drs",load_drs,drs_file),
        ("NS Time","df_ns",load_ns_time,ns_file),
        ("SFDC","df_sfdc",load_sfdc,sfdc_file),
    ]:
        if _f and _key not in st.session_state:
            try:    st.session_state[_key] = _ldr(_f)
            except Exception as e: st.error(f"{_lbl}: {e}")
    if ns_ua_file and "df_ns_unassigned" not in st.session_state:
        try:
            import pandas as _pd
            st.session_state["df_ns_unassigned"] = (
                _pd.read_excel(ns_ua_file) if not ns_ua_file.name.endswith(".csv")
                else _pd.read_csv(ns_ua_file)
            )
        except Exception as e: st.error(f"NS Unassigned: {e}")
    _rev_key = f"df_revenue_{_LOADER_VERSION}"
    # Clear old version if present
    if _rev_key not in st.session_state:
        for _ok in [k for k in st.session_state if k.startswith("df_revenue")]:
            del st.session_state[_ok]
    if rev_file and _rev_key not in st.session_state:
        try:
            from shared.loaders import load_revenue as _lr
            st.session_state["df_revenue"] = _lr(rev_file)
            st.session_state[_rev_key] = True
        except Exception as e: st.error(f"Revenue: {e}")
    if tm_sow_file and "df_tm_sow" not in st.session_state:
        try:
            from shared.loaders import load_tm_sow as _ltm
            st.session_state["df_tm_sow"] = _ltm(tm_sow_file)
        except Exception as e: st.error(f"T&M SOW: {e}")
    st.markdown("---")
    st.markdown("**Session data**")
    _si = [("SS DRS","df_drs"),("NS Time","df_ns"),("SFDC","df_sfdc")]
    if _upload_role in ("manager","manager_only","reporting_only"):
        _si.append(("NS Unassigned","df_ns_unassigned"))
        _si.append(("FF Revenue","df_revenue"))
        _si.append(("T&M SOW","df_tm_sow"))
    for _lbl, _key in _si:
        _ok = _key in st.session_state
        st.markdown(f'<div style="font-size:12px;color:{"#27AE60" if _ok else "rgba(128,128,128,0.4)"};'
                    f'padding:3px 0">{"✓" if _ok else "○"}&nbsp; {_lbl}</div>',
                    unsafe_allow_html=True)
    if any(k in st.session_state for k in ["df_drs","df_ns","df_sfdc","df_ns_unassigned","df_revenue","df_tm_sow"]):
        if st.button("Clear loaded data", use_container_width=True, key="home_clear"):
            # Clear dataframes and uploader widget states
            keys_to_clear = [
                "df_drs","df_ns","df_sfdc","df_ns_unassigned",
                "hub_drs","hub_ns","hub_sfdc","hub_ns_unassigned",
            ]
            # Also clear any FormData/file widget state Streamlit holds internally
            for k in list(st.session_state.keys()):
                if k in ["df_drs","df_ns","df_sfdc","df_ns_unassigned","df_revenue","df_tm_sow"] or k.startswith("hub_"):
                    del st.session_state[k]
            st.rerun()

# ── Run the selected page ─────────────────────────────────────────────────────
pg.run()
