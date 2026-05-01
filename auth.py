"""
PS Platform — Authentication
Simple secrets-based auth. No external library required.

secrets.toml structure:
  [users.pattiswanson]
  password_hash    = "<bcrypt hash>"
  full_roster_name = "Swanson, Patti"
  display_name     = "Patti"

Generate a hash:
  import bcrypt
  bcrypt.hashpw(b"yourpassword", bcrypt.gensalt()).decode()
"""
from __future__ import annotations
import streamlit as st
import bcrypt


def _get_users() -> dict:
    """Return users dict from secrets."""
    try:
        return dict(st.secrets.get("users", {}))
    except Exception:
        return {}


def check_password(username: str, password: str) -> bool:
    """Return True if credentials are valid."""
    users = _get_users()
    user = users.get(username)
    if not user:
        return False
    try:
        stored = user.get("password_hash", "")
        return bcrypt.checkpw(password.encode(), stored.encode())
    except Exception:
        return False


def get_user_info(username: str) -> dict:
    """Return display_name and full_roster_name for a user."""
    users = _get_users()
    user = users.get(username, {})
    return {
        "display_name":     user.get("display_name", username),
        "full_roster_name": user.get("full_roster_name", ""),
    }


def render_login() -> bool:
    """
    Render the login screen. Returns True if authenticated.
    Sets st.session_state keys: authenticated, username, display_name, consultant_name.
    """
    # Hero login card
    st.markdown(
        "<div style='"
        "background:linear-gradient(135deg,#0E223D 0%,#1A3257 100%);"
        "padding:40px 48px 36px;"
        "border-radius:14px;"
        "max-width:440px;"
        "margin:80px auto 0;"
        "'>"
        "<div style='font-size:11px;color:#73DAE3;letter-spacing:2px;"
        "font-weight:700;text-transform:uppercase;margin-bottom:10px'>"
        "Professional Services</div>"
        "<h1 style='color:#fff;margin:0;font-size:26px;font-weight:800'>"
        "PS Platform &amp; Tools</h1>"
        "<p style='color:rgba(255,255,255,0.5);margin:10px 0 24px;font-size:13px'>"
        "Sign in to continue.</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    with st.form("login_form", clear_on_submit=False):
        st.markdown("<div style='max-width:440px;margin:0 auto'>", unsafe_allow_html=True)
        username = st.text_input("Username", placeholder="username")
        password = st.text_input("Password", type="password", placeholder="password")
        submitted = st.form_submit_button("Sign in", use_container_width=True, type="primary")
        st.markdown("</div>", unsafe_allow_html=True)

    if submitted:
        if check_password(username, password):
            info = get_user_info(username)
            st.session_state["authenticated"]   = True
            st.session_state["username"]        = username
            st.session_state["display_name"]    = info["display_name"]
            st.session_state["consultant_name"] = info["full_roster_name"]
            st.rerun()
        else:
            st.error("Invalid username or password.", icon="🔒")
            return False

    return st.session_state.get("authenticated", False)


def require_auth() -> bool:
    """
    Call at the top of every page.
    Returns True if authenticated, False (and renders login) if not.
    """
    if st.session_state.get("authenticated"):
        return True
    render_login()
    return False


def sign_out() -> None:
    """Clear auth state and rerun."""
    for key in ["authenticated", "username", "display_name", "consultant_name",
                "df_drs", "df_ns", "df_sfdc", "df_ns_unassigned",
                "df_revenue", "df_tm_sow", "_product_filter", "view", "product"]:
        st.session_state.pop(key, None)
    st.rerun()
