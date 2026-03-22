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

    # Plain widgets (no st.form — st.switch_page cannot fire inside a form)
    _col = st.columns([1, 2, 1])[1]
    with _col:
        _sel = st.selectbox("Select your name", ["— Select —"] + _display_names,
                            key="login_name")
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
                    st.rerun()  # re-run Home — auth passes → switch_page fires below
                else:
                    st.error("Incorrect password. Default: Zone{LastName}! e.g. ZoneSwanson!")

    with st.expander("🔑 Need to reset your password?"):
        st.caption("Generate a new hash and send it to your admin to update in Streamlit secrets.")
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

# ── Authenticated ─────────────────────────────────────────────────────────────
if _roster and st.session_state.get("consultant_name") != _roster:
    st.session_state["consultant_name"] = _roster

# Redirect immediately — sidebar renders on the destination page
st.switch_page("pages/1_Daily_Briefing.py")
