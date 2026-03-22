"""
PS Tools — Home
Auth + upload hub. Redirects to Daily Briefing after login.
"""
import streamlit as st
import pandas as pd
import streamlit_authenticator as stauth

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN, ACTIVE_EMPLOYEES,
    get_role, is_manager, LEAVER_EXIT_DATES,
)
from shared.loaders import load_drs, load_ns_time, load_sfdc

st.set_page_config(page_title="PS Tools", page_icon=None, layout="wide")

# ── Build credentials from secrets ───────────────────────────────────────────
def _to_dict(obj):
    try:
        d = dict(obj)
        return {k: _to_dict(v) for k, v in d.items()}
    except (TypeError, ValueError):
        return obj

_secrets_creds = st.secrets.get("credentials", {})
_creds = {
    "usernames": {
        uname: _to_dict(udata)
        for uname, udata in _to_dict(_secrets_creds.get("usernames", {})).items()
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

# ── Navigation must be registered BEFORE any st.switch_page calls ─────────────
_auth_name  = st.session_state.get("name", "")
_auth_user  = st.session_state.get("username", "")
_roster     = _creds["usernames"].get(_auth_user, {}).get("full_roster_name", "")
_role       = get_role(_roster) if _roster else None

_consultant_pages = [
    st.Page("pages/1_Daily_Briefing.py",        title="Daily Briefing"),
    st.Page("pages/2_Customer_Reengagement.py", title="Customer Re-Engagement"),
    st.Page("pages/3_Utilization_Report.py",    title="Utilization Report"),
    st.Page("pages/4_Workload_Health_Score.py", title="Workload Health Score"),
    st.Page("pages/6_DRS_Health_Check.py",      title="DRS Health Check"),
    st.Page("pages/7_Vibe_Check.py",            title="Vibe Check ✨"),
]
_manager_pages = [st.Page("pages/5_Capacity_Outlook.py", title="Capacity Outlook")]

if _role in ("manager", "manager_only"):
    st.navigation({"My Tools": _consultant_pages, "Management": _manager_pages})
else:
    st.navigation({"My Tools": _consultant_pages})

# ── Login gate ────────────────────────────────────────────────────────────────
if not st.session_state.get("authentication_status"):
    _user_options  = {d.get("name", u): u for u, d in _creds["usernames"].items()}
    _display_names = sorted(_user_options.keys())

    st.markdown("""
    <div style='background:#1e2c63;padding:32px 40px 28px;border-radius:10px;
                max-width:480px;margin:60px auto 24px'>
        <div style='font-size:11px;color:#a0aec0;letter-spacing:2px;
                    text-transform:uppercase;margin-bottom:8px'>Professional Services</div>
        <h1 style='color:#fff;margin:0;font-size:28px;font-weight:700'>PS Reporting Tools</h1>
        <p style='color:#a0aec0;margin:10px 0 0;font-size:13px'>Sign in to continue.</p>
    </div>""", unsafe_allow_html=True)

    with st.form("login_form", clear_on_submit=False):
        _col = st.columns([1, 2, 1])[1]
        with _col:
            _sel  = st.selectbox("Select your name", ["— Select —"] + _display_names)
            _pw   = st.text_input("Password", type="password", placeholder="Enter your password")
            _btn  = st.form_submit_button("Sign in →", use_container_width=True, type="primary")
            if _btn:
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
                        st.switch_page("pages/1_Daily_Briefing.py")
                    else:
                        st.error("Incorrect password. Default: Zone{LastName}! e.g. ZoneSwanson!")

    with st.expander("🔑 Need to reset your password?"):
        st.caption("Generate a new hash and send it to your admin to update in Streamlit secrets.")
        with st.form("reset_form"):
            _rc = st.columns([1, 2, 1])[1]
            with _rc:
                _rn  = st.selectbox("Your name", ["— Select —"] + _display_names, key="reset_name")
                _rp1 = st.text_input("New password", type="password", key="reset_pw1")
                _rp2 = st.text_input("Confirm password", type="password", key="reset_pw2")
                if st.form_submit_button("Generate new hash", use_container_width=True):
                    if _rn == "— Select —": st.warning("Select your name.")
                    elif len(_rp1) < 8:     st.warning("Password must be at least 8 characters.")
                    elif _rp1 != _rp2:      st.error("Passwords don't match.")
                    else:
                        import bcrypt as _bc
                        _nh = _bc.hashpw(_rp1.encode(), _bc.gensalt()).decode()
                        st.success("Hash generated! Send the below to your admin:")
                        st.code(f'[credentials.usernames.{_user_options[_rn]}]\npassword = "{_nh}"', language="toml")
    st.stop()

# ── Authenticated ─────────────────────────────────────────────────────────────
if _roster and st.session_state.get("consultant_name") != _roster:
    st.session_state["consultant_name"] = _roster

# ── Sidebar: identity + upload hub ───────────────────────────────────────────
_display_first = _roster.split(",")[1].strip() if "," in _roster else _roster
with st.sidebar:
    st.markdown(f"#### {_display_first}")
    st.caption(f"Signed in as **{_auth_name}**")
    authenticator.logout("Sign out", location="sidebar", key="sidebar_logout")
    st.markdown("---")
    st.markdown("**Upload data**")
    st.caption("Upload once — available across all pages this session.")

    _upload_role = get_role(_roster) if _roster else None
    drs_file  = st.file_uploader("SS DRS Export",  type=["xlsx","csv"], key="hub_drs")
    ns_file   = st.file_uploader("NS Time Detail", type=["xlsx","csv"], key="hub_ns")
    sfdc_file = st.file_uploader("SFDC Contacts",  type=["xlsx","csv"], key="hub_sfdc")
    ns_unassigned_file = (
        st.file_uploader("NS Unassigned Projects", type=["xlsx","csv"], key="hub_ns_unassigned",
                         help="Required for Capacity Outlook")
        if _upload_role in ("manager", "manager_only") else None
    )

    for _lbl, _key, _loader, _file in [
        ("SS DRS",        "df_drs",  load_drs,     drs_file),
        ("NS Time",       "df_ns",   load_ns_time, ns_file),
        ("SFDC Contacts", "df_sfdc", load_sfdc,    sfdc_file),
    ]:
        if _file and _key not in st.session_state:
            try:    st.session_state[_key] = _loader(_file)
            except Exception as e: st.error(f"{_lbl}: {e}")

    if ns_unassigned_file and "df_ns_unassigned" not in st.session_state:
        try:
            import pandas as _pd
            st.session_state["df_ns_unassigned"] = (
                _pd.read_excel(ns_unassigned_file)
                if not ns_unassigned_file.name.endswith(".csv")
                else _pd.read_csv(ns_unassigned_file)
            )
        except Exception as e: st.error(f"NS Unassigned: {e}")

    st.markdown("---")
    st.markdown("**Session data**")
    _si = [("SS DRS","df_drs"),("NS Time","df_ns"),("SFDC","df_sfdc")]
    if _upload_role in ("manager","manager_only"): _si.append(("NS Unassigned","df_ns_unassigned"))
    for _lbl, _key in _si:
        _ok = _key in st.session_state
        st.markdown(
            f'<div style="font-size:12px;color:{"#27AE60" if _ok else "rgba(128,128,128,0.4)"};padding:3px 0">'
            f'{"✓" if _ok else "○"}&nbsp; {_lbl}</div>', unsafe_allow_html=True)
    if any(k in st.session_state for k in ["df_drs","df_ns","df_sfdc","df_ns_unassigned"]):
        if st.button("Clear loaded data", use_container_width=True):
            for k in ["df_drs","df_ns","df_sfdc","df_ns_unassigned"]:
                st.session_state.pop(k, None)
            st.rerun()

# Home has no content of its own — Daily Briefing is the landing page
st.switch_page("pages/1_Daily_Briefing.py")
