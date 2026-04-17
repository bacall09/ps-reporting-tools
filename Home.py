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
_LOADER_VERSION = "v20260415a"

st.set_page_config(page_title="PS Projects & Tools", page_icon=None, layout="wide")

st.logo("zone_ps_logo.svg", size="large", link=None)



# ── Register navigation (must be called before any other st.* that could fail) ──
_consultant_pages = [
    st.Page("views/1_Daily_Briefing.py",        title="Daily Briefing",        icon="📋"),
    st.Page("views/8_My_Projects.py",            title="My Projects",           icon="📁"),
    st.Page("views/11_Project_Health.py",        title="Project Health",        icon="🏥"),
    st.Page("views/0_Customer_Profile.py",     title="Customer Profile",      icon="👤"),
    st.Page("views/2_Customer_Reengagement.py",  title="Customer Engagement",   icon="💬"),
    st.Page("views/3_Utilization_Report.py",     title="Utilization Report",    icon="📊"),
    st.Page("views/4_Workload_Health_Score.py",  title="Workload Health Score", icon="⚖️"),
    st.Page("views/6_DRS_Health_Check.py",       title="DRS Health Check",      icon="🔍"),
    st.Page("views/10_Time_Entries.py",           title="Time Entries",          icon="⏱️"),
]
_manager_pages = [
    st.Page("views/13_Portfolio_Analytics.py",   title="Portfolio Analytics",   icon="📈"),
    st.Page("views/5_Capacity_Outlook.py",       title="Capacity Outlook",      icon="🗓️"),
    st.Page("views/9_Revenue_Report.py",         title="Revenue Report",        icon="💰"),
]
_help_pages = [
    st.Page("views/9_Help.py", title="Help",                                    icon="❓"),
]

# ── Build credentials ─────────────────────────────────────────────────────────
def _to_dict(obj):
    try:
        d = dict(obj)
        return {k: _to_dict(v) for k, v in d.items()}
    except (TypeError, ValueError):
        return obj

try:
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
except Exception as _auth_init_err:
    _creds = {"usernames": {}}
    _cookie = {"name": "ps_tools_auth", "key": "fallback_key", "expiry_days": 30}
    authenticator = None

# ── Auth state ────────────────────────────────────────────────────────────────
_auth_user   = st.session_state.get("username", "")
_auth_name   = st.session_state.get("name", "")
_roster      = _creds["usernames"].get(_auth_user, {}).get("full_roster_name", "")
_role        = get_role(_roster) if _roster else None

if _roster and st.session_state.get("consultant_name") != _roster:
    st.session_state["consultant_name"] = _roster

# ── Navigation (page lists defined above, called here with role context) ────────
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
        <div style='background:#050D1F;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;position:relative;overflow:hidden'>
            <svg style='position:absolute;right:-40px;top:50%;transform:translateY(-50%);opacity:0.06;width:200px;height:200px;pointer-events:none' viewBox='0 0 1482 1286.25' xmlns='http://www.w3.org/2000/svg'><g fill='#3B9EFF' fill-rule='evenodd'><path d='M975.127,924.953c2.608-2.68,1.744-5.496-.42-7.829l-57.415-61.872c-2.463-2.655-5.025-2.878-8.443-.991-10.398,5.739-19.024,12.314-27.949,19.885-83.252,70.621-197.471,155.494-298.93,195.556-17.993,7.105-35.256,13.178-54.191,17.329-62.148,13.627-131.853,15.491-192.702-5.298-64.93-22.183-113.878-68.722-142.715-130.542-28.647-61.415-22.393-131.406,11.352-189.217,2.598-2.793,1.405-6.055-1.389-8.184-35.341-26.918-40.303-33.439-69.367-65.686-1.449-1.607-4.102-2.401-5.903-1.138-13.105,9.189-23.232,20.534-33.172,32.961-16.499,20.629-29.73,42.605-38.718,67.541-5.127,10.469-8.378,20.486-10.885,32.065-13.633,62.973-7.701,128.685,17.402,188.142,23.839,56.463,65.297,103.638,114.77,139.169,32.418,23.283,66.848,42.548,103.476,58.385,25.142,10.871,50.281,18.994,76.934,25.12,96.392,22.153,188.876,4.496,276.774-38.393,42.916-20.94,83.188-45.685,121.922-73.568,75.733-54.514,154.643-126.72,219.571-193.435ZM1445.252,792.261c-7.628-38.507-22.817-74.472-43.124-107.897-35.582-58.566-85.801-106.77-139.329-149.092-69.784-55.176-145.355-102.407-225.163-141.162-2.165-1.052-4.941.388-5.391,1.627-.426,1.171-.463,3.413.931,4.628,20.341,17.734,39.847,35.55,58.599,55.093,13.286,14.465,26.223,28.012,37.022,44.544,19.784,30.289,35.735,62.168,50.127,95.397,34.512,31.926,64.863,67.358,90.813,106.359,42.427,63.765,57.696,142.663,37.453,217.116-11.436,42.061-34.763,80.507-64.388,112.265-55.859,59.882-133.144,94.711-214.71,99.157-32.507,1.773-64.093-.538-96.013-6.503-28.16-5.262-70.299-23.997-96.538-36.626-2.312-1.112-4.605-.743-6.449.974-12.635,11.76-25.076,22.901-39.051,33.146l-43.32,31.757c-2.68,1.965-2.195,5.562.439,7.808,70.707,60.309,165.779,100.179,259.837,97.033,39.996-1.336,78.686-6.594,117.486-16.111,94.178-23.099,174.952-71.91,236.526-146.957,23.873-29.096,44.355-60.51,59.779-94.956,29.172-65.148,38.357-137.461,24.463-207.601ZM601.099,242.903c-12.268,10.522-48.215,44.405-47.219,60.482.993,16.01,10.781,31.195,25.227,38.155,14.47,6.972,41.303-10.055,53.886-18.311l65.495-42.972c26.305-17.259,52.496-32.716,80.08-47.834l57.464-31.494c20.451-11.209,41.123-19.851,63.235-27.448,35.852-12.318,72.313-18.084,110.322-17.747,29.787.263,58.398,3.408,86.939,11.449,44.037,12.405,82.745,35.987,114.027,69.974,20.347,22.106,37.598,45.332,51.026,71.732,6.962,13.688,13.008,27.156,16.103,42.311,6.48,31.729,12.267,85.992-.676,115.916-6.013,13.902-13.009,26.627-18.289,40.753-.847,2.264-.768,4.767,1.387,6.461l81.366,63.967c2.003,1.574,5.098.298,6.46-1.592,19.285-26.745,34.599-55.578,45.667-86.804,10.617-29.953,15.416-60.246,15.218-92.192-.482-77.938-29.055-152.791-79.976-211.891-67.16-77.946-169.264-137.487-272.877-146.244-33.524-2.834-66.192-1.328-99.421,3.091-82.214,10.934-149.21,45.218-216.385,92.267-48.269,33.807-94.373,69.644-139.062,107.973ZM72.687,567.553c20.03,44.974,54.35,86.652,88.718,121.568,19.447,19.756,38.882,38.258,60.393,55.711l73.052,59.268c30.921,25.086,74.954,56.331,111.096,72.278,11.713,5.168,23.385,8.99,35.917,11.295,12.922,2.375,24.878,1.136,37.309-3.088,18.441-6.266,35.538-14.698,52.671-24.006,1.792-.974,2.85-2.213,3.058-3.936.179-1.483-.47-3.163-1.914-4.548-14.129-13.542-27.174-27.284-42.195-40.056l-78.193-66.48-93.5-82.422c-23.176-20.43-44.471-41.737-65.536-64.239-15.19-16.227-28.591-32.64-40.05-51.639-20.601-34.157-31.396-72.282-30.182-112.398.614-20.279,2.364-39.861,7.45-59.369,8.872-34.031,50.72-76.652,77.451-99.125,3.767-7.04,2.459-14.401,2.885-21.735.884-15.227,3.244-29.908,5.647-44.959,4.285-26.824,22.718-58.984,38.899-80.638,1.348-1.805,1.936-3.535.891-4.937-.951-1.277-2.618-2.49-4.589-2.222-52.436,7.145-104.92,34.806-146.088,67.704-25.632,20.484-48.458,43.456-68.934,69.137-46.339,58.118-62.952,131.49-53.428,204.864,4.697,36.186,14.376,70.75,29.171,103.971ZM1196.886,310.029c-4.882-10.39-12.371-18.773-20.659-26.723-18.771-18.007-40.425-31.674-64.291-42.362-57.569-25.783-110.906-28.064-173.214-22.213-61.067,5.735-111.183,25.069-164.567,54.081-24.678,13.412-48.301,26.866-71.885,42.28l-105.247,68.787c-85.308,55.756-195.138,156.138-256.755,237.876-1.598,2.12-2.206,4.81-.222,6.912l76.342,80.886c1.468,1.556,2.9,1.672,4.715,1.249,1.397-.326,1.99-1.717,2.793-3.377,3.117-6.44,6.665-11.977,11.238-17.864,38.52-49.59,82.099-94.54,130.222-135.261,40.87-34.583,82.783-67.442,126.68-98.902,83.71-59.991,188.529-115.793,291.15-127.921,23.653-2.795,46.328-.575,69.656,3.405,27.197,4.641,52.661,12.543,78.69,21.347l38.004,12.855c13.849,4.685,27.221-3.226,30.503-17.755,2.725-12.064,2.293-25.708-3.154-37.301Z'/></g></svg>
            <div style='position:relative;z-index:1'>
            <div style='font-size:11px;color:#3B9EFF;letter-spacing:2px;font-weight:700;text-transform:uppercase;margin-bottom:8px'>Professional Services</div>
            <h1 style='color:#fff;margin:0;font-size:28px;font-weight:700'>PS Projects &amp; Tools</h1>
            <p style='color:rgba(255,255,255,0.45);margin:10px 0 0;font-size:13px'>Sign in to continue.</p>
            </div>
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
        # Product filter — shown on Daily Briefing only
        # Use st.context if available (Streamlit >= 1.36), otherwise check page marker
        _show_prod_filter = True
        try:
            _page_path = str(getattr(st.context, "page_script_hash", "") or "")
        except Exception:
            _page_path = ""
        _cur_page = st.session_state.get("current_page", "")
        if _cur_page and _cur_page != "Daily Briefing":
            _show_prod_filter = False
        if _show_prod_filter:
            st.markdown("**Filter by product:**")
            _product_filter_home = st.selectbox("Product", ["All products"] + _all_prods,
                                                key="home_product_filter", label_visibility="collapsed")
        else:
            _product_filter_home = st.session_state.get("_product_filter", "All products")

        # Store in session state so Daily Briefing can read them
        st.session_state["_view_browse"]   = _browse
        st.session_state["_product_filter"] = _product_filter_home
        st.markdown("---")

    st.markdown("**Upload data**")
    st.caption("Upload once — available across all pages this session.")

    drs_file  = st.file_uploader("SS DRS Export",  type=["xlsx","csv"], key="hub_drs", label_visibility="collapsed")
    st.markdown('<a href="https://www.smartsheet.com" target="_blank" style="font-size:11px;opacity:0.6;">↗ Open SS DRS Report</a>', unsafe_allow_html=True)

    ns_file   = st.file_uploader("NS Time Detail", type=["xlsx","csv"], key="hub_ns", label_visibility="collapsed")
    st.markdown('<a href="https://3838224.app.netsuite.com/app/common/search/searchresults.nl?searchid=66732&amp;saverun=T&amp;whence=" target="_blank" style="font-size:11px;opacity:0.6;">↗ Open NS Time Detail Search</a>', unsafe_allow_html=True)

    sfdc_file = st.file_uploader("SFDC Contacts",  type=["xlsx","csv"], key="hub_sfdc", label_visibility="collapsed")
    st.markdown('<a href="https://drive.google.com/drive/u/1/folders/1VdI_WjuVclF5xN9fG7dEIz1WDu4QRE0m" target="_blank" style="font-size:11px;opacity:0.6;">↗ Open SFDC Contacts (Google Drive)</a>', unsafe_allow_html=True)

    ns_ua_file = (
        st.file_uploader("NS Unassigned Projects", type=["xlsx","csv"], key="hub_ns_unassigned", label_visibility="collapsed",
                         help="Required for Capacity Outlook")
        if _upload_role in ("manager","manager_only","reporting_only") else None
    )
    if _upload_role in ("manager","manager_only","reporting_only"):
        st.markdown('<a href="https://3838224.app.netsuite.com/app/common/search/searchresults.nl?searchid=68439&whence=" target="_blank" style="font-size:11px;opacity:0.6;">↗ Open NS Unassigned Projects</a>', unsafe_allow_html=True)

    rev_file = (
        st.file_uploader("NS FF Revenue Charges", type=["xlsx","csv"], key="hub_revenue", label_visibility="collapsed",
                         help="Required for Revenue Report")
        if _upload_role in ("manager","manager_only","reporting_only") else None
    )
    tm_sow_file = (
        st.file_uploader("SFDC T&M SOW", type=["xlsx","csv"], key="hub_tm_sow", label_visibility="collapsed",
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
