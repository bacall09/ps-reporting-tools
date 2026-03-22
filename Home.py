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
    MILESTONE_COLS_MAP, get_role, is_manager,
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
                        st.rerun()
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

# ── Role-aware page navigation ────────────────────────────────────────────────
# NOTE: Home.py must NOT be included as a st.Page — nav.run() re-executes the
# entrypoint causing infinite recursion. Home is the implicit default because
# it IS the entrypoint. st.navigation() here just controls sidebar grouping.
_name = st.session_state.get("consultant_name", "") or ""
_role = get_role(_name) if _name and _name != "— Select —" else None

_consultant_pages = [
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

st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        h1,h2,h3,h4,p,div,label,button { font-family: 'Manrope', sans-serif !important; }
        /* Use inherit so Streamlit theme controls text colour in both light and dark */
        .brief-header  { font-size: 24px; font-weight: 700; color: inherit;
                         margin-bottom: 4px; }
        .brief-sub     { font-size: 13px; margin-bottom: 20px; opacity: 0.6; }
        .section-label { font-size: 11px; font-weight: 700; text-transform: uppercase;
                         letter-spacing: 0.8px; color: #4472C4; margin-bottom: 8px; }
        /* Transparent cards — border adapts to theme */
        .metric-card   { background: transparent;
                         border: 1px solid rgba(128,128,128,0.2);
                         border-radius: 8px; padding: 16px 20px; margin-bottom: 12px; }
        .metric-val    { font-size: 26px; font-weight: 700; color: inherit; }
        .metric-lbl    { font-size: 12px; opacity: 0.6; margin-top: 2px; }
        .proj-card     { background: transparent;
                         border: 1px solid rgba(128,128,128,0.15);
                         border-radius: 8px; padding: 14px 18px; margin-bottom: 8px; }
        .proj-name     { font-size: 14px; font-weight: 700; color: inherit; margin-bottom: 4px; }
        .proj-meta     { font-size: 12px; opacity: 0.6; }
        .rag-R{display:inline-block;width:10px;height:10px;border-radius:50%;background:#E74C3C;margin-right:6px;}
        .rag-A{display:inline-block;width:10px;height:10px;border-radius:50%;background:#F39C12;margin-right:6px;}
        .rag-G{display:inline-block;width:10px;height:10px;border-radius:50%;background:#27AE60;margin-right:6px;}
        .rag- {display:inline-block;width:10px;height:10px;border-radius:50%;background:rgba(128,128,128,0.4);margin-right:6px;}
        .action-badge{display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:4px;margin-right:6px;}
        .badge-red   {background:rgba(231,76,60,0.15);color:#E74C3C;}
        .badge-amber {background:rgba(243,156,18,0.15);color:#D68910;}
        .badge-blue  {background:rgba(68,114,196,0.15);color:#4472C4;}
        .badge-gray  {background:rgba(128,128,128,0.12);color:inherit;opacity:0.7;}
        .badge-green {background:rgba(39,174,96,0.15);color:#27AE60;}
        .divider{border:none;border-top:1px solid rgba(128,128,128,0.2);margin:20px 0;}
        .data-ok  {font-size:12px;color:#27AE60;padding:3px 0;}
        .data-miss{font-size:12px;opacity:0.35;padding:3px 0;}
    </style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — identity + upload hub
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    # ── Identity (from login) ─────────────────────────────────────────────────
    _display_first = _roster_name.split(",")[1].strip() if "," in _roster_name else _roster_name
    st.markdown(f"#### {_display_first}")
    st.caption(f"Signed in as **{_auth_name}**")
    authenticator.logout("Sign out", location="sidebar", key="sidebar_logout")
    st.markdown("---")

    selected = _roster_name  # identity comes from login, not a selectbox

    role = get_role(selected) if selected != "— Select —" else None
    view_as = selected
    _product_filter = "All products"  # default, overridden below for managers
    if role in ("manager","manager_only") and selected != "— Select —":
        from shared.config import EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE
        from shared.constants import EMPLOYEE_ROLES, get_role as _get_role

        def _get_region(name):
            if name in PS_REGION_OVERRIDE:
                return PS_REGION_OVERRIDE[name]
            loc = EMPLOYEE_LOCATION.get(name) or next(
                (EMPLOYEE_LOCATION[k] for k in EMPLOYEE_LOCATION if name.startswith(k.split(",")[0])), None
            )
            return PS_REGION_MAP.get(loc, "Other") if loc else "Other"

        _active_consultants = sorted([
            n for n in CONSULTANT_DROPDOWN if _get_role(n) in ("consultant", "manager")
        ])

        # ── Step 1: View As (who) ────────────────────────────────────────────
        _by_region_all = {}
        for name in _active_consultants:
            _by_region_all.setdefault(_get_region(name), []).append(name)

        # Include managers-only in the browse list
        _mgr_only_names = sorted([
            e for e in ACTIVE_EMPLOYEES
            if get_role(e) == "manager_only"
        ])

        _browse_options = ["— My own view —", "👥 All team"]
        for region in sorted(_by_region_all.keys()):
            _browse_options.append(f"── {region} ──")
            _browse_options.extend(_by_region_all[region])
        if _mgr_only_names:
            _browse_options.append("── Managers ──")
            _browse_options.extend(_mgr_only_names)

        st.markdown("---")
        st.markdown("**View as:**")
        browse = st.selectbox(
            "Browse team",
            options=_browse_options,
            key="home_browse",
            label_visibility="collapsed",
        )

        if browse == "— My own view —":
            view_as = selected
        elif browse.startswith("── ") and browse.endswith(" ──"):
            # Region header selected — show all consultants in that region
            _region_selected = browse.replace("── ", "").replace(" ──", "").strip()
            if _region_selected == "Managers":
                view_as = "ALL_MANAGERS"
            else:
                view_as = f"REGION:{_region_selected}"
        elif browse == "👥 All team":
            view_as = "ALL"
        else:
            view_as = browse

        # ── Step 2: Filter by product (refines the view above) ───────────────
        _all_products = sorted({
            p for n in _active_consultants
            for p in EMPLOYEE_ROLES.get(n, {}).get("products", [])
            if p and p != "All"
        })

        st.markdown("**Filter by product:**")
        _product_filter = st.selectbox(
            "Product",
            options=["All products"] + _all_products,
            key="home_product_filter",
            label_visibility="collapsed",
        )
    # ── Upload hub ────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**Upload data**")
    st.caption("Upload once — available across all pages this session.")

    drs_file  = st.file_uploader("SS DRS Export",  type=["xlsx","csv"], key="hub_drs")
    ns_file   = st.file_uploader("NS Time Detail", type=["xlsx","csv"], key="hub_ns")
    sfdc_file = st.file_uploader("SFDC Contacts",  type=["xlsx","csv"], key="hub_sfdc")

    # NS Unassigned — managers only
    ns_unassigned_file = None
    if role in ("manager", "manager_only") and selected != "— Select —":
        ns_unassigned_file = st.file_uploader(
            "NS Unassigned Projects",
            type=["xlsx","csv"],
            key="hub_ns_unassigned",
            help="Required for Capacity Outlook — different export from NS Time Detail"
        )

    for label, key, loader, file in [
        ("SS DRS",         "df_drs",          load_drs,      drs_file),
        ("NS Time",        "df_ns",           load_ns_time,  ns_file),
        ("SFDC Contacts",  "df_sfdc",         load_sfdc,     sfdc_file),
    ]:
        if file and key not in st.session_state:
            try:
                st.session_state[key] = loader(file)
            except Exception as e:
                st.error(f"{label}: {e}")

    # NS Unassigned — load raw (Capacity Outlook has its own loader)
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

    # ── Status indicator ──────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**Session data**")
    _status_items = [("SS DRS","df_drs"),("NS Time","df_ns"),("SFDC","df_sfdc")]
    if role in ("manager","manager_only") and selected != "— Select —":
        _status_items.append(("NS Unassigned","df_ns_unassigned"))
    for label, key in _status_items:
        loaded = key in st.session_state
        st.markdown(
            f'<div class="{"data-ok" if loaded else "data-miss"}">'
            f'{"✓" if loaded else "○"}&nbsp; {label}</div>',
            unsafe_allow_html=True,
        )

    _all_keys = ["df_drs","df_ns","df_sfdc","df_ns_unassigned"]
    if any(k in st.session_state for k in _all_keys):
        st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)
        if st.button("Clear loaded data", use_container_width=True):
            for k in _all_keys:
                st.session_state.pop(k, None)
            st.rerun()

# Guard handled by authentication above

# ══════════════════════════════════════════════════════════════════════════════
# PULL DATA FROM SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
df_drs  = st.session_state.get("df_drs")
df_ns   = st.session_state.get("df_ns")

today     = date.today()
month_key = today.strftime("%Y-%m")
view_name = view_as

# ── Product type crosswalk (EMPLOYEE_ROLES labels → DRS project_type values)
_PT_MAP = {
    "Billing":             ["billing", "zonebilling"],
    "Capture":             ["zonecapture", "capture", "zonecapture + e-invoicing"],
    "Approvals":           ["zoneapprovals", "approvals"],
    "Reconcile":           ["zonereconcile", "reconcile", "reconcile 2.0"],
    "Payments":            ["zonepayments", "payments"],
    "Payroll":             ["payroll", "zonepayroll"],
    "Reporting":           ["reporting", "zonereporting"],
    "e-Invoicing":         ["e-invoicing", "einvoicing", "zonecapture + e-invoicing"],
    "PSP":                 ["psp"],
    "CC Statement Import": ["cc import", "cc statement import"],
    "SFTP Connector":      ["sftp"],
}

# Resolve active product filter
_active_product = _product_filter if _product_filter != "All products" else None
_product_project_types = set(_PT_MAP.get(_active_product, [])) if _active_product else None

def _name_variants(full_name):
    if full_name in ("ALL",) or full_name.startswith("REGION:") or full_name == "ALL_MANAGERS":
        return []
    parts = [p.strip() for p in full_name.split(",")]
    variants = [full_name.lower()]
    if len(parts) == 2:
        variants.append(f"{parts[1]} {parts[0]}".lower())
        variants.append(parts[0].lower())
    return variants

view_variants = _name_variants(view_name)

# Build set of names for region/manager views
_view_name_set = None
if view_name.startswith("REGION:"):
    _region_sel = view_name.split(":", 1)[1]
    from shared.config import EMPLOYEE_LOCATION as _EL, PS_REGION_MAP as _RM, PS_REGION_OVERRIDE as _RO
    def _gr(n):
        if n in _RO: return _RO[n]
        _l = _EL.get(n, "")
        return _RM.get(_l, "Other")
    _view_name_set = {
        n.lower() for n in CONSULTANT_DROPDOWN if _gr(n) == _region_sel
    }
elif view_name == "ALL_MANAGERS":
    _view_name_set = {n.lower() for n in ACTIVE_EMPLOYEES if get_role(n) == "manager_only"}

def _match_name(val):
    """Match a name value against the current view selection."""
    if view_name in ("ALL",):
        return True
    if _view_name_set is not None:
        return any(ns in str(val).lower() for ns in _view_name_set)
    return any(v in str(val).lower() for v in view_variants)

def _match_project_type(val):
    """Return True if project_type matches the active product filter."""
    if _product_project_types is None:
        return True
    return str(val).strip().lower() in _product_project_types

# Filter DRS — by PM name first, then refine by project_type if product filter active
_is_group_view = view_name in ("ALL", "ALL_MANAGERS") or view_name.startswith("REGION:")
my_projects = pd.DataFrame()
if df_drs is not None and not df_drs.empty:
    if _is_group_view:
        if _view_name_set is not None:
            pm_col = df_drs.get("project_manager", pd.Series(dtype=str))
            _drs_by_who = df_drs[pm_col.apply(lambda v: _match_name(str(v)))]
        else:
            _drs_by_who = df_drs
    else:
        pm_mask = df_drs.get("project_manager", pd.Series(dtype=str)).apply(lambda v: _match_name(str(v)))
        _drs_by_who = df_drs[pm_mask]
        if _drs_by_who.empty and not is_manager(view_name):
            _drs_by_who = df_drs
            st.caption("ℹ️ No PM column matched — showing all projects.")

    if _product_project_types is not None and not _drs_by_who.empty:
        pt_col = _drs_by_who.get("project_type", pd.Series(dtype=str))
        my_projects = _drs_by_who[pt_col.apply(lambda v: _match_project_type(str(v)))].copy()
    else:
        my_projects = _drs_by_who.copy()

# Filter NS — by employee name
my_ns = pd.DataFrame()
if df_ns is not None and not df_ns.empty:
    if _is_group_view and _view_name_set is None:
        my_ns = df_ns.copy()
    else:
        emp_mask = df_ns.get("employee", pd.Series(dtype=str)).apply(lambda v: _match_name(str(v)))
        my_ns = df_ns[emp_mask].copy()

# Enrich DRS with NS inactivity
if not my_projects.empty and df_ns is not None:
    try:
        my_projects = calc_days_inactive(my_projects, df_ns)
    except Exception:
        pass

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
# Greeting always uses the logged-in user's first name
_my_display = selected.split(",")[1].strip() if "," in selected else selected

# Sub-line context: what view + product filter is active
if view_name == "ALL":
    _view_label  = "All products" if not _active_product else _active_product
    _view_sub    = f"Viewing: All team · {_view_label}"
elif view_name.startswith("REGION:"):
    _region_label = view_name.split(":", 1)[1]
    _prod_label   = f" · {_active_product}" if _active_product else ""
    _view_sub     = f"Viewing: {_region_label} team{_prod_label}"
elif view_name == "ALL_MANAGERS":
    _view_sub     = "Viewing: Managers"
elif view_name != selected:
    _view_display = view_name.split(",")[1].strip() if "," in view_name else view_name
    _prod_label   = f" · {_active_product}" if _active_product else ""
    _view_sub     = f"Viewing: {_view_display}{_prod_label}"
elif _active_product:
    _view_sub     = f"Viewing: My projects · {_active_product}"
else:
    _view_sub     = None

emp_info     = EMPLOYEE_ROLES.get(selected, {})
emp_role     = emp_info.get("role", "Consultant")
emp_products = emp_info.get("products", [])

ch, cm = st.columns([3, 1])
with ch:
    _hour = datetime.now().hour
    _greeting = (
        "Good morning" if _hour < 12
        else "Good afternoon" if _hour < 17
        else "Good evening"
    )
    st.markdown(f'<div class="brief-header">{_greeting}, {_my_display}</div>', unsafe_allow_html=True)
    _sub_parts = [emp_role, ", ".join(emp_products) if emp_products else "All Products", today.strftime("%A, %B %-d %Y")]
    if _view_sub:
        _sub_parts.append(_view_sub)
    st.markdown(f'<div class="brief-sub">{' · '.join(_sub_parts)}</div>', unsafe_allow_html=True)
with cm:
    loc = EMPLOYEE_LOCATION.get(selected, "")
    if isinstance(loc, tuple): loc = loc[0]
    region = PS_REGION_OVERRIDE.get(selected, PS_REGION_MAP.get(loc, ""))
    if region:
        st.markdown(f'<div class="badge-blue action-badge" style="margin-top:12px">{region}</div>', unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Utilization
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">This Month — Utilization</div>', unsafe_allow_html=True)

if view_name in ("ALL",) or view_name.startswith("REGION:") or view_name == "ALL_MANAGERS":
    # Sum available hours across all consultants in scope
    if view_name == "ALL":
        _scope = CONSULTANT_DROPDOWN
    elif view_name.startswith("REGION:"):
        _scope = [n for n in CONSULTANT_DROPDOWN if _gr(n) == view_name.split(":",1)[1]]
    else:
        _scope = [n for n in ACTIVE_EMPLOYEES if get_role(n) == "manager_only"]
    if _active_product:
        _scope = [n for n in _scope if _active_product in EMPLOYEE_ROLES.get(n, {}).get("products", [])]
    _avail_total = 0
    for _cn in _scope:
        _cl = EMPLOYEE_LOCATION.get(_cn, "")
        if isinstance(_cl, tuple): _cl = _cl[0]
        _cr = PS_REGION_OVERRIDE.get(_cn, PS_REGION_MAP.get(_cl, ""))
        _ch = AVAIL_HOURS.get(_cl, AVAIL_HOURS.get(_cr, {})).get(month_key)
        if _ch and isinstance(_ch, (int, float)):
            _avail_total += _ch
        elif _ch and isinstance(_ch, tuple):
            _avail_total += _ch[0]
    avail = _avail_total if _avail_total > 0 else None
    loc_key = None
else:
    loc_key = EMPLOYEE_LOCATION.get(view_name, "")
    if isinstance(loc_key, tuple): loc_key = loc_key[0]
    avail = AVAIL_HOURS.get(loc_key, {}).get(month_key) if loc_key else None

if not my_ns.empty and "date" in my_ns.columns and "hours" in my_ns.columns:
    my_ns["date"] = pd.to_datetime(my_ns["date"], errors="coerce")
    month_ns = my_ns[my_ns["date"].dt.strftime("%Y-%m") == month_key].copy()

    bt_col    = "billing_type" if "billing_type" in month_ns.columns else None
    if bt_col:
        _bt       = month_ns[bt_col].fillna("").str.strip().str.lower()
        admin_hrs = round(month_ns[_bt == "internal"]["hours"].sum(), 1)
        tm_hrs    = round(month_ns[_bt == "t&m"]["hours"].sum(), 1)
        ff_rows   = month_ns[_bt == "fixed fee"].copy()
    else:
        admin_hrs = 0.0
        tm_hrs    = round(month_ns["hours"].sum(), 1)
        ff_rows   = pd.DataFrame()

    ff_credit = 0.0; ff_overrun = 0.0
    if not ff_rows.empty and "project" in ff_rows.columns:
        ff_rows = ff_rows.sort_values("date")
        _con: dict = {}
        for _, _r in ff_rows.iterrows():
            _proj  = str(_r.get("project","")).strip()
            _ptype = str(_r.get("project_type","")).strip()
            _hrs   = float(_r.get("hours",0) or 0)
            if _hrs <= 0: continue
            _m = [(k,float(v)) for k,v in DEFAULT_SCOPE.items() if k.strip().lower() in _ptype.lower()]
            _sc = max(_m, key=lambda x: len(x[0]))[1] if _m else None
            if _sc is None: ff_credit += _hrs; continue
            _used = _con.get(_proj, 0.0); _rem = _sc - _used
            if _rem <= 0:
                ff_overrun += _hrs
            elif _hrs <= _rem:
                ff_credit += _hrs; _con[_proj] = _used + _hrs
            else:
                ff_credit += _rem; ff_overrun += _hrs - _rem; _con[_proj] = _sc

    credit_hrs  = round(tm_hrs + ff_credit, 2)
    overrun_hrs = round(ff_overrun, 2)
    admin_hrs   = round(admin_hrs, 2)
    credit_pct  = round(credit_hrs  / avail * 100, 2) if avail else None
    overrun_pct = round(overrun_hrs / avail * 100, 2) if avail else None
    admin_pct   = round(admin_hrs   / avail * 100, 2) if avail else None

    total_booked = round(month_ns[month_ns["hours"] > 0]["hours"].sum(), 1)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        v = f"{avail}h" if avail else "—"
        lbl = "Available this month" if avail else "Available hrs (location not mapped)"
        st.markdown(f'<div class="metric-card"><div class="metric-val">{v}</div><div class="metric-lbl">{lbl}</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><div class="metric-val">{total_booked}h</div><div class="metric-lbl">Hours booked this month</div></div>', unsafe_allow_html=True)
    with c3:
        if credit_pct is not None:
            col = "#27AE60" if credit_pct >= 70 else ("#F39C12" if credit_pct >= 60 else "#E74C3C")
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{col}">{credit_pct}%</div><div class="metric-lbl">Utilization credit &nbsp;·&nbsp; {credit_hrs}h credited</div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">Utilization credit %</div></div>', unsafe_allow_html=True)
    with c4:
        if overrun_pct is not None:
            col = "#E74C3C" if overrun_pct > 10 else ("#F39C12" if overrun_pct > 0 else "#718096")
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{col}">{overrun_pct}%</div><div class="metric-lbl">FF overrun &nbsp;·&nbsp; {overrun_hrs}h over budget</div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">FF project overrun %</div></div>', unsafe_allow_html=True)
    with c5:
        v = f"{admin_pct}%" if admin_pct is not None else "—"
        st.markdown(f'<div class="metric-card"><div class="metric-val">{v}</div><div class="metric-lbl">Internal / admin &nbsp;·&nbsp; {admin_hrs}h</div></div>', unsafe_allow_html=True)
else:
    if df_ns is None:
        st.info("Upload NS Time Detail in the sidebar to see your utilization snapshot.")
    else:
        st.warning(f"No time entries found for **{view_name}** in the NS file.")

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Re-engagement actions
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Re-Engagement Actions</div>', unsafe_allow_html=True)

if df_drs is None:
    st.info("Upload SS DRS Export in the sidebar to see re-engagement actions.")
elif my_projects.empty:
    st.warning(f"No projects found for **{view_name}** in the DRS file.")
elif "days_inactive" not in my_projects.columns:
    st.info("Upload NS Time Detail alongside DRS to calculate project inactivity.")
else:
    stale = my_projects[my_projects["days_inactive"] >= 14].sort_values("days_inactive", ascending=False)
    if stale.empty:
        st.markdown('<span class="action-badge badge-green">All clear</span> No projects flagged for re-engagement.', unsafe_allow_html=True)
    else:
        st.markdown(f"**{len(stale)} project(s)** need re-engagement outreach")

        for _ri, (_, row) in enumerate(stale.iterrows()):
            proj_name  = row.get("project_name", "—")
            days_inac  = int(row.get("days_inactive", 0))
            phase      = row.get("phase", "—")
            tier       = suggest_tier(days_inac)
            tier_label = tier if tier else "Monitor"

            with st.expander(f"{proj_name} — {days_inac}d inactive · {tier_label}"):
                ci, ca = st.columns([1, 2])
                with ci:
                    st.markdown(f"**Phase:** {phase}")
                    st.markdown(f"**Days inactive:** {days_inac}")
                    try:
                        lm = calc_last_milestone(row)
                        if lm and lm != "—": st.markdown(f"**Last milestone:** {lm}")
                    except Exception: pass
                with ca:
                    if tier and tier in TEMPLATES:
                        tmpl = TEMPLATES[tier]
                        st.markdown(f"**Suggested template:** {tier}")
                        st.markdown(f"*Subject:* {tmpl.get('subject','')}")
                        if st.button("Draft outreach →", key=f"draft_{_ri}_{proj_name[:20]}", type="primary"):
                            st.session_state["_jump_to_proj"] = proj_name
                            st.session_state["_jump_tier"]    = tier
                            st.switch_page("pages/2_Customer_Reengagement.py")
                    else:
                        st.markdown("No template matched for this inactivity window.")

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — Welcome email actions (placeholder)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Welcome Email Actions</div>', unsafe_allow_html=True)
st.markdown(
    '<div style="border:1px dashed rgba(128,128,128,0.3);border-radius:8px;padding:16px 20px;opacity:0.6">' +
    '<div style="font-size:13px;font-weight:600;margin-bottom:4px">Coming soon</div>' +
    '<div style="font-size:12px">Newly assigned projects that haven\'t had a welcome / intro email sent yet will surface here, with a pre-filled template ready to send.</div>' +
    '</div>',
    unsafe_allow_html=True,
)

st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.caption("PS Reporting Tools · Internal use only · Data loaded this session only")
