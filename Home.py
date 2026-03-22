"""
PS Tools — Home / Daily Briefing
Upload once. Everything loads here and stays available for the whole session.
"""
import streamlit as st
import pandas as pd
import streamlit_authenticator as stauth
from datetime import date, datetime

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN, ACTIVE_EMPLOYEES,
    MILESTONE_COLS_MAP, get_role, is_manager, LEAVER_EXIT_DATES,
)
from shared.config import (
    AVAIL_HOURS, EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP, DEFAULT_SCOPE,
)
from shared.loaders import (
    load_drs, load_ns_time, load_sfdc,
    calc_days_inactive, calc_last_milestone,
    suggest_tier_from_days,
)
from shared.template_utils import TEMPLATES, suggest_tier

st.set_page_config(page_title="PS Tools", page_icon=None, layout="wide")

# ══════════════════════════════════════════════════════════════════════════════
# AUTHENTICATION
# ══════════════════════════════════════════════════════════════════════════════
# Credentials are stored in Streamlit secrets (never in the repo).
# st.secrets returns AttrDict objects — must be converted to plain dicts recursively.
def _to_dict(obj):
    """Recursively convert AttrDict / secrets objects to plain dicts."""
    try:
        d = dict(obj)
        return {k: _to_dict(v) for k, v in d.items()}
    except (TypeError, ValueError):
        return obj

# Build credentials dict
_secrets_creds = st.secrets.get("credentials", {})
_usernames_raw = _secrets_creds.get("usernames", {})
_creds = {
    "usernames": {
        uname: _to_dict(udata)
        for uname, udata in _to_dict(_usernames_raw).items()
    }
}

_cookie_raw = st.secrets.get("cookie", {})
_cookie = _to_dict(_cookie_raw) if _cookie_raw else {
    "name": "ps_tools_auth",
    "key": "fallback_key_change_me",
    "expiry_days": 30,
}

authenticator = stauth.Authenticate(
    credentials        = _creds,
    cookie_name        = _cookie.get("name", "ps_tools_auth"),
    cookie_key         = _cookie.get("key", "fallback_key"),
    cookie_expiry_days = int(_cookie.get("expiry_days", 30)),
)

# ── Custom login UI (replaces default stauth form) ────────────────────────────
# Check cookie first — may already be authenticated
_auth_status = st.session_state.get("authentication_status")

if not _auth_status:
    # Build display name → username lookup for the dropdown
    _user_options = {
        udata.get("name", uname): uname
        for uname, udata in _creds["usernames"].items()
    }
    _display_names = sorted(_user_options.keys())

    # Hero header
    st.markdown("""
    <div style='background:#1e2c63;padding:32px 40px 28px;border-radius:10px;
                max-width:480px;margin:60px auto 24px'>
        <div style='font-size:11px;color:#a0aec0;letter-spacing:2px;
                    text-transform:uppercase;margin-bottom:8px'>Professional Services</div>
        <h1 style='color:#fff;margin:0;font-size:28px;font-weight:700'>PS Reporting Tools</h1>
        <p style='color:#a0aec0;margin:10px 0 0;font-size:13px'>Sign in to continue.</p>
    </div>
    """, unsafe_allow_html=True)

    # Login form
    with st.form("login_form", clear_on_submit=False):
        _col = st.columns([1, 2, 1])[1]  # centre column
        with _col:
            _selected_display = st.selectbox(
                "Select your name",
                options=["— Select —"] + _display_names,
            )
            _password_input = st.text_input(
                "Password",
                type="password",
                placeholder="Enter your password",
            )
            _login_btn = st.form_submit_button("Sign in →", use_container_width=True, type="primary")

            if _login_btn:
                if _selected_display == "— Select —":
                    st.warning("Please select your name.")
                elif not _password_input:
                    st.warning("Please enter your password.")
                else:
                    _username = _user_options[_selected_display]
                    _stored_hash = _creds["usernames"][_username].get("password", "")
                    import bcrypt as _bcrypt
                    try:
                        _match = _bcrypt.checkpw(
                            _password_input.encode(),
                            _stored_hash.encode()
                        )
                    except Exception:
                        _match = False

                    if _match:
                        st.session_state["authentication_status"] = True
                        st.session_state["username"] = _username
                        st.session_state["name"]     = _selected_display
                        # Resolve roster name and set immediately
                        _login_creds = _creds["usernames"].get(_username, {})
                        _login_roster = _login_creds.get("full_roster_name", "")
                        if _login_roster:
                            st.session_state["consultant_name"] = _login_roster
                        st.switch_page("pages/1_Daily_Briefing.py")
                    else:
                        st.error("Incorrect password. Your default is Zone{LastName}! e.g. ZoneSwanson!")

    # Password reset expander
    with st.expander("🔑 Need to reset your password?"):
        st.caption("Enter your new password below. Copy the hash and send it to your admin to update in Streamlit secrets.")
        with st.form("reset_form"):
            _r_col = st.columns([1, 2, 1])[1]
            with _r_col:
                _r_name   = st.selectbox("Your name", ["— Select —"] + _display_names, key="reset_name")
                _r_pw1    = st.text_input("New password", type="password", key="reset_pw1")
                _r_pw2    = st.text_input("Confirm password", type="password", key="reset_pw2")
                _r_btn    = st.form_submit_button("Generate new hash", use_container_width=True)
                if _r_btn:
                    if _r_name == "— Select —":
                        st.warning("Select your name.")
                    elif not _r_pw1 or len(_r_pw1) < 8:
                        st.warning("Password must be at least 8 characters.")
                    elif _r_pw1 != _r_pw2:
                        st.error("Passwords don't match.")
                    else:
                        import bcrypt as _bcrypt
                        _new_hash = _bcrypt.hashpw(_r_pw1.encode(), _bcrypt.gensalt()).decode()
                        _r_user   = _user_options[_r_name]
                        st.success("Hash generated! Send the below to your admin:")
                        st.code(f'[credentials.usernames.{_r_user}]\npassword = "{_new_hash}"', language="toml")
    st.stop()

_auth_status = st.session_state.get("authentication_status")
_auth_name   = st.session_state.get("name", "")
_auth_user   = st.session_state.get("username", "")

# ── Authenticated — resolve full roster name from secrets ────────────────────
_user_creds  = _creds["usernames"].get(_auth_user, {})
_roster_name = _user_creds.get("full_roster_name", "")

# Auto-set consultant_name in session state from login
if _roster_name and st.session_state.get("consultant_name") != _roster_name:
    st.session_state["consultant_name"] = _roster_name

# ── Sidebar — sign out + upload hub ──────────────────────────────────────────
_display_first = _roster_name.split(",")[1].strip() if "," in _roster_name else _roster_name
with st.sidebar:
    st.markdown(f"#### {_display_first}")
    st.caption(f"Signed in as **{_auth_name}**")
    authenticator.logout("Sign out", location="sidebar", key="sidebar_logout")
    st.markdown("---")
    st.markdown("**Upload data**")
    st.caption("Upload once — available across all pages this session.")

    _role_for_upload = get_role(_roster_name) if _roster_name else None

    drs_file  = st.file_uploader("SS DRS Export",  type=["xlsx","csv"], key="hub_drs")
    ns_file   = st.file_uploader("NS Time Detail", type=["xlsx","csv"], key="hub_ns")
    sfdc_file = st.file_uploader("SFDC Contacts",  type=["xlsx","csv"], key="hub_sfdc")

    if _role_for_upload in ("manager", "manager_only"):
        ns_unassigned_file = st.file_uploader(
            "NS Unassigned Projects", type=["xlsx","csv"], key="hub_ns_unassigned",
            help="Required for Capacity Outlook"
        )
    else:
        ns_unassigned_file = None

    for _lbl, _key, _loader, _file in [
        ("SS DRS",        "df_drs",  load_drs,      drs_file),
        ("NS Time",       "df_ns",   load_ns_time,  ns_file),
        ("SFDC Contacts", "df_sfdc", load_sfdc,     sfdc_file),
    ]:
        if _file and _key not in st.session_state:
            try:
                st.session_state[_key] = _loader(_file)
            except Exception as e:
                st.error(f"{_lbl}: {e}")

    if ns_unassigned_file and "df_ns_unassigned" not in st.session_state:
        try:
            import pandas as _pd
            st.session_state["df_ns_unassigned"] = (
                _pd.read_excel(ns_unassigned_file)
                if not ns_unassigned_file.name.endswith(".csv")
                else _pd.read_csv(ns_unassigned_file)
            )
        except Exception as e:
            st.error(f"NS Unassigned: {e}")

    st.markdown("---")
    st.markdown("**Session data**")
    _status_items = [("SS DRS","df_drs"),("NS Time","df_ns"),("SFDC","df_sfdc")]
    if _role_for_upload in ("manager","manager_only"):
        _status_items.append(("NS Unassigned","df_ns_unassigned"))
    for _lbl, _key in _status_items:
        _loaded = _key in st.session_state
        st.markdown(
            f'<div style="font-size:12px;color:{"#27AE60" if _loaded else "rgba(128,128,128,0.4)"};padding:3px 0">'
            f'{"✓" if _loaded else "○"}&nbsp; {_lbl}</div>',
            unsafe_allow_html=True,
        )
    if any(k in st.session_state for k in ["df_drs","df_ns","df_sfdc","df_ns_unassigned"]):
        st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)
        if st.button("Clear loaded data", use_container_width=True):
            for k in ["df_drs","df_ns","df_sfdc","df_ns_unassigned"]:
                st.session_state.pop(k, None)
            st.rerun()

# ── Role-aware page navigation ────────────────────────────────────────────────
_name = st.session_state.get("consultant_name", "") or ""
_role = get_role(_name) if _name else None

_consultant_pages = [
    st.Page("pages/1_Daily_Briefing.py",       title="Daily Briefing"),
    st.Page("pages/2_Customer_Reengagement.py", title="Customer Re-Engagement"),
    st.Page("pages/3_Utilization_Report.py",    title="Utilization Report"),
    st.Page("pages/4_Workload_Health_Score.py", title="Workload Health Score"),
    st.Page("pages/6_DRS_Health_Check.py",      title="DRS Health Check"),
    st.Page("pages/7_Vibe_Check.py",            title="Vibe Check ✨"),
]
_manager_pages = [
    st.Page("pages/5_Capacity_Outlook.py", title="Capacity Outlook"),
]

if _role in ("manager", "manager_only"):
    st.navigation({
        "My Tools": _consultant_pages,
        "Management": _manager_pages,
    })
else:
    st.navigation({
        "My Tools": _consultant_pages,
    })

