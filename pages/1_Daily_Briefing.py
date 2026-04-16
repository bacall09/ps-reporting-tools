"""
PS Tools — Daily Briefing
Month-to-date utilization snapshot, team breakdown, and re-engagement actions.
"""
import streamlit as st
import pandas as pd
from datetime import date, datetime

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN, ACTIVE_EMPLOYEES,
    MILESTONE_COLS_MAP, get_role, is_manager, LEAVER_EXIT_DATES,
)
from shared.config import (
    AVAIL_HOURS, EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP, DEFAULT_SCOPE,
)
from shared.loaders import (
    load_drs, load_ns_time,
    calc_days_inactive, calc_last_milestone,
)
from shared.template_utils import TEMPLATES, suggest_tier
from shared.whs import consultant_whs, workload_level, GREEN, AMBER, RED


# ── Pull identity from session state ─────────────────────────────────────────
selected = st.session_state.get("consultant_name", "")
role     = get_role(selected) if selected else "consultant"

# View As and Filter are set by Home.py sidebar — just read from session state
_browse          = st.session_state.get("_view_browse", "— My own view —")
_product_filter  = st.session_state.get("_product_filter", "All products")

view_as = selected
if role == "manager":
    if _browse == "— My own view —":
        view_as = selected
    elif _browse.startswith("── ") and _browse.endswith(" ──"):
        _rs = _browse.replace("── ","").replace(" ──","").strip()
        view_as = "ALL_MANAGERS" if _rs == "Managers" else f"REGION:{_rs}"
    elif _browse == "👥 All team":
        view_as = "ALL"
    else:
        view_as = _browse

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        h1,h2,h3,h4,p,div,label,button { font-family: 'Manrope', sans-serif !important; }
        .brief-header  { font-size: 24px; font-weight: 700; color: inherit; margin-bottom: 4px; }
        .brief-sub     { font-size: 13px; margin-bottom: 20px; opacity: 0.6; }
        .section-label { font-size: 11px; font-weight: 700; text-transform: uppercase;
                         letter-spacing: 0.8px; color: #4472C4; margin-bottom: 8px; }
        .metric-card   { background: transparent; border: 1px solid rgba(128,128,128,0.2);
                         border-radius: 8px; padding: 16px 20px; margin-bottom: 12px; }
        .metric-val    { font-size: 26px; font-weight: 700; color: inherit; }
        .metric-lbl    { font-size: 12px; opacity: 0.6; margin-top: 2px; }
        .metric-help   { display:inline-block; margin-left:5px; font-size:11px; opacity:0.55;
                         cursor:help; position:relative; }
        .metric-help:hover::after {
            content: attr(data-tip);
            position:absolute; left:50%; transform:translateX(-50%);
            top:calc(100% + 8px); background:#0E223D; color:#fff;
            font-size:12.5px; font-weight:500; padding:11px 15px; border-radius:8px;
            white-space:normal; width:280px; z-index:99999; line-height:1.65;
            box-shadow:0 4px 20px rgba(0,0,0,0.6); opacity:1 !important;
            letter-spacing:0.1px;
            pointer-events:none;
        }
        .action-badge{display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:4px;margin-right:6px;}
        .badge-red   {background:rgba(192,57,43,0.15);color:#C0392B;}
        .badge-amber {background:rgba(243,156,18,0.15);color:#D68910;}
        .badge-blue  {background:rgba(68,114,196,0.15);color:#4472C4;}
        .badge-gray  {background:rgba(128,128,128,0.12);color:inherit;opacity:0.7;}
        .badge-green {background:rgba(39,174,96,0.15);color:#27AE60;}
        .divider{border:none;border-top:1px solid rgba(128,128,128,0.2);margin:20px 0;}
    </style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PULL DATA FROM SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
df_drs  = st.session_state.get("df_drs")
df_ns   = st.session_state.get("df_ns")

today     = date.today()
month_key = today.strftime("%Y-%m")
view_name = view_as

# ── Product type crosswalk (EMPLOYEE_ROLES labels → DRS project_type values)
# DRS uses "ZoneApp: X", "ZoneBill: X", "ZonePay: X", "ZoneRpt: X" format
_PT_MAP = {
    "Billing":             ["zonebill: zb_standard", "zonebill: zb_premium", "zonebill: zab implementation",
                            "zonebill: zab partner impl", "zonebill: optimization", "zonebill: optimization audit",
                            "zonebill: subscription services", "billing", "zonebilling"],
    "Capture":             ["zoneapp: capture", "zoneapp: capture & e-invoicing", "zonecapture", "capture"],
    "Approvals":           ["zoneapp: approvals", "zoneapprovals", "approvals"],
    "Reconcile":           ["zoneapp: reconcile", "zoneapp: reconcile 2.0", "zoneapp: reconcile psp",
                            "zonereconcile", "reconcile", "reconcile 2.0"],
    "Payments":            ["zoneapp: payments", "zoneapp: ap payment", "zonepayments", "payments"],
    "Payroll":             ["zonepay: implementation", "zonepay: optimization", "payroll", "zonepayroll"],
    "Reporting":           ["zonerpt: install, dwh", "zonerpt: optimization", "reporting", "zonereporting"],
    "e-Invoicing":         ["zoneapp: e-invoicing", "zoneapp: capture & e-invoicing", "e-invoicing", "einvoicing"],
    "PSP":                 ["zoneapp: reconcile psp", "psp"],
    "CC Statement Import": ["zoneapp: cc statement import", "cc import", "cc statement import"],
    "SFTP Connector":      ["zoneapp: sftp connector", "sftp"],
    "All":                 ["zoneapps: consulting", "zep: implementation", "zep: optimization"],
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

# ── Region helper (module level so avail hours can use it) ───────────────────
def _gr(n):
    """Return PS region for a consultant name."""
    if n in PS_REGION_OVERRIDE: return PS_REGION_OVERRIDE[n]
    _l = EMPLOYEE_LOCATION.get(n, "")
    return PS_REGION_MAP.get(_l, "Other")

view_variants = _name_variants(view_name)

# Build set of names for region/manager views
_view_name_set = None
if view_name.startswith("REGION:"):
    _region_sel = view_name.split(":", 1)[1]
    # Only include people who actually do product delivery for name matching
    _region_names = [
        n for n in CONSULTANT_DROPDOWN
        if _gr(n) == _region_sel
        and len(EMPLOYEE_ROLES.get(n, {}).get("products", [])) > 0
    ]
    # Include name variants but NOT first-name-only (too ambiguous)
    _view_name_set = set()
    for n in _region_names:
        parts = [p.strip() for p in n.split(",")]
        _view_name_set.add(n.lower())                              # "longalong, santiago"
        _view_name_set.add(parts[0].lower())                       # "longalong" (last name)
        if len(parts) == 2:
            _view_name_set.add(f"{parts[1]} {parts[0]}".lower())  # "santiago longalong"
elif view_name == "ALL_MANAGERS":
    _view_name_set = {n.lower() for n in ACTIVE_EMPLOYEES if get_role(n) == "manager_only"}

def _match_name(val):
    """Match a name value against the current view selection."""
    if view_name == "ALL":
        return True
    if _view_name_set is not None:
        v = str(val).strip().lower()
        # Check full string match OR "First Last" / "Last, First" exact match
        return v in _view_name_set or any(
            v == ns or v.startswith(ns + " ") or v.endswith(" " + ns)
            for ns in _view_name_set
        )
    return any(v in str(val).lower() for v in view_variants)

def _match_project_type(val):
    """Return True if project_type matches the active product filter."""
    if _product_project_types is None:
        return True
    v = str(val).strip().lower()
    # Exact match first
    if v in _product_project_types:
        return True
    # Contains match — handles slight variations
    return any(pt in v or v in pt for pt in _product_project_types)

# Filter DRS — by PM name first, then refine by project_type if product filter active
_is_group_view = view_name in ("ALL", "ALL_MANAGERS") or view_name.startswith("REGION:")

# Build project scope lookup AND overrun map from DRS
# DRS has cumulative actual_hours + budgeted_hours — far more reliable than NS month slice
_proj_scope:   dict = {}  # project_name.lower() → budget hours
_proj_actual:  dict = {}  # project_name.lower() → actual hours to date
_proj_overrun: dict = {}  # project_name.lower() → overrun hours (actual - budget, if > 0)

if df_drs is not None and not df_drs.empty:
    for _, _dr in df_drs.iterrows():
        _pn  = str(_dr.get("project_name","")).strip()
        _pid = str(_dr.get("project_id","")).strip()
        _pt  = str(_dr.get("project_type","")).strip()
        _bgt = _dr.get("budgeted_hours", None)
        _act = _dr.get("actual_hours", None)
        _co  = _dr.get("change_order", 0) or 0
        try:
            _budget = float(_bgt) + float(_co) if _bgt else None
        except (TypeError, ValueError):
            _budget = None
        # Fallback scope from DEFAULT_SCOPE if no explicit budget
        if _budget is None:
            _m = [(k, float(v)) for k, v in DEFAULT_SCOPE.items()
                  if k.strip().lower() in _pt.lower()]
            _budget = max(_m, key=lambda x: len(x[0]))[1] if _m else None
        try:
            _actual = float(_act) if _act is not None else None
        except (TypeError, ValueError):
            _actual = None

        for _key in [_pn.lower(), _pid.lower()]:
            if not _key or _key in ("", "nan"): continue
            if _budget: _proj_scope[_key]  = _budget
            if _actual: _proj_actual[_key] = _actual
            if _budget and _actual and _actual > _budget:
                _proj_overrun[_key] = round(_actual - _budget, 2)

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

# Filter NS — by employee name, then also by project_type if product filter active
my_ns = pd.DataFrame()
if df_ns is not None and not df_ns.empty:
    if _is_group_view and _view_name_set is None:
        _ns_by_who = df_ns
    else:
        emp_mask = df_ns.get("employee", pd.Series(dtype=str)).apply(lambda v: _match_name(str(v)))
        _ns_by_who = df_ns[emp_mask]
    # Apply product filter to NS if project_type column exists
    if _product_project_types is not None and "project_type" in _ns_by_who.columns:
        _ns_pt = _ns_by_who["project_type"].fillna("").str.strip().str.lower()
        my_ns = _ns_by_who[_ns_pt.isin(_product_project_types)].copy()
    else:
        my_ns = _ns_by_who.copy()

# ── END DEBUG ────────────────────────────────────────────────────────────────
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

_hour = datetime.now().hour
_greeting = (
    "Good morning" if _hour < 12
    else "Good afternoon" if _hour < 17
    else "Good evening"
)
_sub_parts = [emp_role, ", ".join(emp_products) if emp_products else "All Products", today.strftime("%A, %B %-d %Y")]
if _view_sub:
    _sub_parts.append(_view_sub)
loc = EMPLOYEE_LOCATION.get(selected, "")
if isinstance(loc, tuple): loc = loc[0]
region = PS_REGION_OVERRIDE.get(selected, PS_REGION_MAP.get(loc, ""))
_region_pill = (
    f"<span style='display:inline-block;margin-top:12px;padding:4px 12px;border-radius:20px;"
    f"background:rgba(59,158,255,0.15);border:1px solid rgba(59,158,255,0.35);color:#3B9EFF;"
    f"font-size:11px;font-weight:700;letter-spacing:.5px'>{region}</span>"
) if region else ""
_sub_str = " · ".join(_sub_parts)

st.markdown(
    f"<div style='background:#050D1F;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;"
    f"font-family:Manrope,sans-serif;position:relative;overflow:hidden'>"
    f"<svg style='position:absolute;right:-40px;top:50%;transform:translateY(-50%);opacity:0.06;width:200px;height:200px;pointer-events:none' viewBox='0 0 1482 1286.25' xmlns='http://www.w3.org/2000/svg'><g fill='#3B9EFF' fill-rule='evenodd'><path d='M975.127,924.953c2.608-2.68,1.744-5.496-.42-7.829l-57.415-61.872c-2.463-2.655-5.025-2.878-8.443-.991-10.398,5.739-19.024,12.314-27.949,19.885-83.252,70.621-197.471,155.494-298.93,195.556-17.993,7.105-35.256,13.178-54.191,17.329-62.148,13.627-131.853,15.491-192.702-5.298-64.93-22.183-113.878-68.722-142.715-130.542-28.647-61.415-22.393-131.406,11.352-189.217,2.598-2.793,1.405-6.055-1.389-8.184-35.341-26.918-40.303-33.439-69.367-65.686-1.449-1.607-4.102-2.401-5.903-1.138-13.105,9.189-23.232,20.534-33.172,32.961-16.499,20.629-29.73,42.605-38.718,67.541-5.127,10.469-8.378,20.486-10.885,32.065-13.633,62.973-7.701,128.685,17.402,188.142,23.839,56.463,65.297,103.638,114.77,139.169,32.418,23.283,66.848,42.548,103.476,58.385,25.142,10.871,50.281,18.994,76.934,25.12,96.392,22.153,188.876,4.496,276.774-38.393,42.916-20.94,83.188-45.685,121.922-73.568,75.733-54.514,154.643-126.72,219.571-193.435ZM1445.252,792.261c-7.628-38.507-22.817-74.472-43.124-107.897-35.582-58.566-85.801-106.77-139.329-149.092-69.784-55.176-145.355-102.407-225.163-141.162-2.165-1.052-4.941.388-5.391,1.627-.426,1.171-.463,3.413.931,4.628,20.341,17.734,39.847,35.55,58.599,55.093,13.286,14.465,26.223,28.012,37.022,44.544,19.784,30.289,35.735,62.168,50.127,95.397,34.512,31.926,64.863,67.358,90.813,106.359,42.427,63.765,57.696,142.663,37.453,217.116-11.436,42.061-34.763,80.507-64.388,112.265-55.859,59.882-133.144,94.711-214.71,99.157-32.507,1.773-64.093-.538-96.013-6.503-28.16-5.262-70.299-23.997-96.538-36.626-2.312-1.112-4.605-.743-6.449.974-12.635,11.76-25.076,22.901-39.051,33.146l-43.32,31.757c-2.68,1.965-2.195,5.562.439,7.808,70.707,60.309,165.779,100.179,259.837,97.033,39.996-1.336,78.686-6.594,117.486-16.111,94.178-23.099,174.952-71.91,236.526-146.957,23.873-29.096,44.355-60.51,59.779-94.956,29.172-65.148,38.357-137.461,24.463-207.601ZM601.099,242.903c-12.268,10.522-48.215,44.405-47.219,60.482.993,16.01,10.781,31.195,25.227,38.155,14.47,6.972,41.303-10.055,53.886-18.311l65.495-42.972c26.305-17.259,52.496-32.716,80.08-47.834l57.464-31.494c20.451-11.209,41.123-19.851,63.235-27.448,35.852-12.318,72.313-18.084,110.322-17.747,29.787.263,58.398,3.408,86.939,11.449,44.037,12.405,82.745,35.987,114.027,69.974,20.347,22.106,37.598,45.332,51.026,71.732,6.962,13.688,13.008,27.156,16.103,42.311,6.48,31.729,12.267,85.992-.676,115.916-6.013,13.902-13.009,26.627-18.289,40.753-.847,2.264-.768,4.767,1.387,6.461l81.366,63.967c2.003,1.574,5.098.298,6.46-1.592,19.285-26.745,34.599-55.578,45.667-86.804,10.617-29.953,15.416-60.246,15.218-92.192-.482-77.938-29.055-152.791-79.976-211.891-67.16-77.946-169.264-137.487-272.877-146.244-33.524-2.834-66.192-1.328-99.421,3.091-82.214,10.934-149.21,45.218-216.385,92.267-48.269,33.807-94.373,69.644-139.062,107.973ZM72.687,567.553c20.03,44.974,54.35,86.652,88.718,121.568,19.447,19.756,38.882,38.258,60.393,55.711l73.052,59.268c30.921,25.086,74.954,56.331,111.096,72.278,11.713,5.168,23.385,8.99,35.917,11.295,12.922,2.375,24.878,1.136,37.309-3.088,18.441-6.266,35.538-14.698,52.671-24.006,1.792-.974,2.85-2.213,3.058-3.936.179-1.483-.47-3.163-1.914-4.548-14.129-13.542-27.174-27.284-42.195-40.056l-78.193-66.48-93.5-82.422c-23.176-20.43-44.471-41.737-65.536-64.239-15.19-16.227-28.591-32.64-40.05-51.639-20.601-34.157-31.396-72.282-30.182-112.398.614-20.279,2.364-39.861,7.45-59.369,8.872-34.031,50.72-76.652,77.451-99.125,3.767-7.04,2.459-14.401,2.885-21.735.884-15.227,3.244-29.908,5.647-44.959,4.285-26.824,22.718-58.984,38.899-80.638,1.348-1.805,1.936-3.535.891-4.937-.951-1.277-2.618-2.49-4.589-2.222-52.436,7.145-104.92,34.806-146.088,67.704-25.632,20.484-48.458,43.456-68.934,69.137-46.339,58.118-62.952,131.49-53.428,204.864,4.697,36.186,14.376,70.75,29.171,103.971ZM1196.886,310.029c-4.882-10.39-12.371-18.773-20.659-26.723-18.771-18.007-40.425-31.674-64.291-42.362-57.569-25.783-110.906-28.064-173.214-22.213-61.067,5.735-111.183,25.069-164.567,54.081-24.678,13.412-48.301,26.866-71.885,42.28l-105.247,68.787c-85.308,55.756-195.138,156.138-256.755,237.876-1.598,2.12-2.206,4.81-.222,6.912l76.342,80.886c1.468,1.556,2.9,1.672,4.715,1.249,1.397-.326,1.99-1.717,2.793-3.377,3.117-6.44,6.665-11.977,11.238-17.864,38.52-49.59,82.099-94.54,130.222-135.261,40.87-34.583,82.783-67.442,126.68-98.902,83.71-59.991,188.529-115.793,291.15-127.921,23.653-2.795,46.328-.575,69.656,3.405,27.197,4.641,52.661,12.543,78.69,21.347l38.004,12.855c13.849,4.685,27.221-3.226,30.503-17.755,2.725-12.064,2.293-25.708-3.154-37.301Z'/></g></svg>"
        f"<div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;"
    f"color:#3B9EFF;margin-bottom:10px;font-family:Manrope,sans-serif'>Professional Services · Daily Briefing</div>"
    f"<h1 style='color:#fff;margin:0;font-size:28px;font-weight:800;font-family:Manrope,sans-serif;line-height:1.15'>"
    f"{_greeting}, {_my_display}</h1>"
    f"<p style='color:rgba(255,255,255,0.6);margin:8px 0 0;font-size:14px;font-family:Manrope,sans-serif;line-height:1.6'>"
    f"{_sub_str}</p>"
    f"{_region_pill}"
    f"</div>",
    unsafe_allow_html=True
)

# ── Pre-compute snapshot data so briefing can appear before utilization ───────
_ioh     = my_projects.get("_on_hold", pd.Series(False, index=my_projects.index)).astype(bool) if not my_projects.empty else pd.Series(dtype=bool)
_active  = my_projects[~_ioh].copy() if not my_projects.empty else pd.DataFrame()
_snap_today = pd.Timestamp.today().normalize()
_snap_n7    = _snap_today + pd.Timedelta(days=7)
_snap_14    = _snap_today - pd.Timedelta(days=14)
_PHASE_ORDER_DB = ["00. onboarding","01. requirements and design","02. configuration",
                "03. enablement/training","04. uat","05. prep for go-live",
                "06. go-live","07. data migration","08. ready for support transition","09. phase 2 scoping"]
def _pidx_db(p):
    pl = str(p).strip().lower()
    for i, ph in enumerate(_PHASE_ORDER_DB):
        if pl.startswith(ph[:6]) or ph in pl or pl in ph: return i
    return -1
if not _active.empty:
    _gls = (_active[_active["go_live_date"].notna() & (_active["go_live_date"] >= _snap_today) & (_active["go_live_date"] <= _snap_n7)]
            .sort_values("go_live_date") if "go_live_date" in _active.columns else pd.DataFrame())
    _ihc = (_active[_active["go_live_date"].notna() & (_active["go_live_date"] >= _snap_14) & (_active["go_live_date"] < _snap_today)]
            .sort_values("go_live_date") if "go_live_date" in _active.columns else pd.DataFrame())
    _leg_s    = _active.get("legacy", pd.Series(False, index=_active.index)).astype(bool)
    _onb_plus = _active["phase"].fillna("").apply(lambda p: _pidx_db(p) >= 0)
    _no_intro = (~_active["ms_intro_email"].notna()) if "ms_intro_email" in _active.columns else pd.Series(True, index=_active.index)
    _mi       = _active[(~_leg_s) & _onb_plus & _no_intro] if "ms_intro_email" in _active.columns else pd.DataFrame()
    _stale    = _active[_active["days_inactive"].fillna(0) >= 14].sort_values("days_inactive", ascending=False) if "days_inactive" in _active.columns else pd.DataFrame()
    _rag_red  = pd.DataFrame(); _rag_yellow = pd.DataFrame()
    if "rag" in _active.columns:
        _rv = _active["rag"].fillna("").astype(str).str.strip().str.lower()
        _rag_red = _active[_rv == "red"]; _rag_yellow = _active[_rv == "yellow"]
else:
    _gls = _ihc = _mi = _stale = _rag_red = _rag_yellow = pd.DataFrame()

# ── This Week Briefing ────────────────────────────────────────────────────
_p1, _p2, _p3 = [], [], []

if len(_rag_red) > 0:
    _red_names = ", ".join(str(r.get("project_name","")).split(" - ")[0][:25] for _, r in _rag_red.head(3).iterrows())
    _p1.append(f"**{len(_rag_red)} project{'s' if len(_rag_red)>1 else ''} flagged Red RAG** ({_red_names}{'...' if len(_rag_red)>3 else ''}) — these need your attention first.")

if len(_gls) > 0:
    _gl_names = ", ".join(str(r.get("project_name","")).split(" - ")[0][:20] for _, r in _gls.iterrows())
    _p1.append(f"**{len(_gls)} project{'s are' if len(_gls)>1 else ' is'} going live this week** ({_gl_names}) — confirm readiness and have cutover support in place.")

if len(_ihc) > 0:
    _ihc_names = ", ".join(str(r.get("project_name","")).split(" - ")[0][:20] for _, r in _ihc.head(3).iterrows())
    _p1.append(f"**{len(_ihc)} project{'s are' if len(_ihc)>1 else ' is'} in hypercare** ({_ihc_names}) — check in proactively and log any post-go-live issues.")

if len(_rag_yellow) > 0:
    _yel_names = ", ".join(str(r.get("project_name","")).split(" - ")[0][:25] for _, r in _rag_yellow.head(2).iterrows())
    _p2.append(f"**{len(_rag_yellow)} Yellow RAG project{'s' if len(_rag_yellow)>1 else ''}** ({_yel_names}) — review blockers before they escalate to Red.")

if len(_stale) > 0:
    _stale_top  = _stale.iloc[0]
    _stale_name = str(_stale_top.get("project_name","")).split(" - ")[0][:25]
    _stale_days = int(_stale_top.get("days_inactive", 0))
    if len(_stale) == 1:
        _p2.append(f"**1 project needs re-engagement** ({_stale_name}, {_stale_days}d inactive) — send a check-in to re-establish momentum.")
    else:
        _p2.append(f"**{len(_stale)} projects need re-engagement**, led by {_stale_name} ({_stale_days}d inactive) — use Customer Engagement to draft outreach.")

if len(_mi) > 0:
    _p2.append(f"**{len(_mi)} project{'s are' if len(_mi)>1 else ' is'} missing an intro email** — a quick win to close before end of week.")

_oh_count = int(_ioh.sum()) if hasattr(_ioh, "sum") else 0
if _oh_count > 0:
    _p3.append(f"You have **{_oh_count} project{'s' if _oh_count>1 else ''} on hold** — ensure On Hold Reason and Responsible for Delay are recorded on each.")

if _p1 or _p2 or _p3:

    # Build paragraph prose — conversational, not bullet list
    def _proj_list(df, n=3):
        names = [str(r.get("project_name","")).split(" - ")[0].strip()[:28] for _, r in df.head(n).iterrows()]
        if len(df) > n: names.append(f"and {len(df)-n} more")
        return ", ".join(names)

    _para_attn = ""
    if len(_rag_red) > 0:
        _para_attn += f"{len(_rag_red)} project{'s are' if len(_rag_red)>1 else ' is'} on Red RAG ({_proj_list(_rag_red)}). "
    if len(_gls) > 0:
        _para_attn += f"{'With' if _para_attn else ''} {_proj_list(_gls)} going live this week — confirm cutover readiness before the end of the day. "
    if len(_ihc) > 0:
        _para_attn += f"{_proj_list(_ihc)} {'are' if len(_ihc)>1 else 'is'} in week-one hypercare — a proactive check-in today would be timely."

    _para_quick = ""
    if len(_rag_yellow) > 0:
        _para_quick += f"{_proj_list(_rag_yellow)} {'are' if len(_rag_yellow)>1 else 'is'} at Yellow RAG — a quick review of blockers now could prevent escalation. "
    if len(_stale) > 0:
        _stale_top  = _stale.iloc[0]
        _stale_name = str(_stale_top.get("project_name","")).split(" - ")[0].strip()[:28]
        _stale_days = int(_stale_top.get("days_inactive",0))
        if len(_stale) == 1:
            _para_quick += f"{_stale_name} hasn't had contact in {_stale_days} days — a short check-in would re-establish momentum. "
        else:
            _para_quick += f"{len(_stale)} projects are overdue for outreach, with {_stale_name} leading at {_stale_days} days inactive — use Customer Engagement to draft messages. "
    if len(_mi) > 0:
        _para_quick += f"{len(_mi)} project{'s are' if len(_mi)>1 else ' is'} missing an intro email date. If the emails were sent, logging the dates is a quick close."

    _oh_count = int(_ioh.sum()) if hasattr(_ioh, "sum") else 0
    _para_house = ""
    if _oh_count > 0:
        _para_house = f"{_oh_count} project{'s are' if _oh_count>1 else ' is'} on hold. Make sure each has an On Hold Reason and Responsible for Delay recorded — these are flagged in DRS Health Check if missing."

    _bhtml = """<div style='border-radius:8px;border:1px solid rgba(59,158,255,0.2);overflow:hidden;margin-bottom:16px;font-family:Manrope,sans-serif'>
  <div style='background:rgba(59,158,255,0.07);padding:10px 20px;border-bottom:1px solid rgba(59,158,255,0.15);display:flex;align-items:center;gap:10px'>
<span style='font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3B9EFF'>This Week&#39;s Focus</span>
<span style='font-size:11px;color:var(--text-color,inherit);opacity:0.35;margin-left:auto'>Rule-based · AI briefings coming soon</span>
  </div>
  <div style='padding:18px 22px;display:flex;flex-direction:column;gap:14px'>"""

    if _para_attn:
        _bhtml += f"""<div style='display:flex;gap:14px;align-items:flex-start'>
<div style='flex-shrink:0;width:3px;background:#C0392B;border-radius:2px;align-self:stretch;min-height:36px'></div>
<div>
  <div style='font-size:10px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#C0392B;margin-bottom:5px'>Needs attention</div>
  <div style='font-size:13px;color:inherit;line-height:1.7'>{_para_attn.strip()}</div>
</div>
  </div>"""

    if _para_quick:
        _sep = "<div style='height:1px;background:rgba(128,128,128,0.12)'></div>" if _para_attn else ""
        _bhtml += f"""{_sep}<div style='display:flex;gap:14px;align-items:flex-start'>
<div style='flex-shrink:0;width:3px;background:#3B9EFF;border-radius:2px;align-self:stretch;min-height:36px'></div>
<div>
  <div style='font-size:10px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#3B9EFF;margin-bottom:5px'>Quick wins this week</div>
  <div style='font-size:13px;color:inherit;line-height:1.7'>{_para_quick.strip()}</div>
</div>
  </div>"""

    if _para_house:
        _sep2 = "<div style='height:1px;background:rgba(128,128,128,0.12)'></div>" if (_para_attn or _para_quick) else ""
        _bhtml += f"""{_sep2}<div style='display:flex;gap:14px;align-items:flex-start'>
<div style='flex-shrink:0;width:3px;background:rgba(128,128,128,0.3);border-radius:2px;align-self:stretch;min-height:36px'></div>
<div>
  <div style='font-size:10px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:inherit;opacity:0.45;margin-bottom:5px'>Housekeeping</div>
  <div style='font-size:13px;color:inherit;opacity:0.6;line-height:1.7'>{_para_house.strip()}</div>
</div>
  </div>"""

    _bhtml += "</div></div>"
    st.markdown(_bhtml, unsafe_allow_html=True)


st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Utilization
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">This Month — Utilization</div>', unsafe_allow_html=True)

prior_htd: dict = {}  # seeded by FF engine; referenced by _build_row

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
    avail = round(_avail_total, 2) if _avail_total > 0 else None
    loc_key = None
else:
    loc_key = EMPLOYEE_LOCATION.get(view_name, "")
    if isinstance(loc_key, tuple): loc_key = loc_key[0]
    avail = AVAIL_HOURS.get(loc_key, {}).get(month_key) if loc_key else None

# Safe defaults in case NS not uploaded
tm_hrs = 0.0; ff_credit = 0.0; ff_overrun = 0.0; ff_unscoped = 0.0
admin_hrs = 0.0; credit_hrs = 0.0; overrun_hrs = 0.0; util_hrs = 0.0
util_pct = None; overrun_pct = None; admin_pct = None; total_booked = 0.0

if not my_ns.empty and "date" in my_ns.columns and "hours" in my_ns.columns:
    my_ns["date"] = pd.to_datetime(my_ns["date"], errors="coerce")
    month_ns = my_ns[my_ns["date"].dt.strftime("%Y-%m") == month_key].copy()

    bt_col = "billing_type" if "billing_type" in month_ns.columns else None
    if bt_col:
        _bt       = month_ns[bt_col].fillna("").str.strip().str.lower()
        admin_hrs = round(month_ns[_bt == "internal"]["hours"].sum(), 2)
        tm_hrs    = round(month_ns[_bt == "t&m"]["hours"].sum(), 2)
        # Match Util Report Rule 3: anything not internal or T&M = Fixed Fee
        ff_rows   = month_ns[(_bt != "internal") & (_bt != "t&m")].copy()
    else:
        admin_hrs = 0.0
        tm_hrs    = round(month_ns["hours"].sum(), 2)
        ff_rows   = pd.DataFrame()

    ff_credit = 0.0; ff_overrun = 0.0; ff_unscoped = 0.0
    if not ff_rows.empty and "project" in ff_rows.columns:
        ff_rows = ff_rows.sort_values(["project", "date"])

        # Build prior_htd — same formula as Util Report assign_credits
        prior_htd: dict = {}
        if "hours_to_date" in month_ns.columns:
            _htd_grp_cols = [c for c in ["project","project_type"] if c in month_ns.columns]
            for _key, _grp in month_ns.groupby(_htd_grp_cols):
                _pk = tuple(" ".join(str(k).strip().split()) for k in (_key if isinstance(_key, tuple) else (_key,)))
                try:
                    _max_htd  = float(_grp["hours_to_date"].dropna().astype(float).max() or 0)
                    _period_h = float(_grp["hours"].sum() or 0)
                    prior_htd[_pk] = max(0.0, _max_htd - _period_h)
                except Exception:
                    prior_htd[_pk] = 0.0

        _con: dict = {}
        for _, _r in ff_rows.iterrows():
            _proj  = " ".join(str(_r.get("project","")).split())
            _ptype = str(_r.get("project_type","")).strip()
            _hrs   = float(_r.get("hours", 0) or 0)
            _date  = str(_r.get("date", ""))[:10]
            if _hrs <= 0: continue

            _m  = [(k, float(v)) for k, v in DEFAULT_SCOPE.items()
                   if k.strip().lower() in _ptype.lower()]
            _sc = max(_m, key=lambda x: len(x[0]))[1] if _m else None

            if _sc is None:
                ff_unscoped += _hrs
                continue

            _ck = (_proj, _ptype)  # composite key: project + product type
            if _ck not in _con: _con[_ck] = prior_htd.get(_ck, prior_htd.get((_proj,), 0.0))
            _used = _con[_ck]; _rem = _sc - _used
            if _rem <= 0:
                ff_overrun += _hrs
            elif _hrs <= _rem:
                ff_credit += _hrs; _con[_ck] = _used + _hrs
            else:
                ff_credit += _rem; ff_overrun += _hrs - _rem; _con[_ck] = _sc

    credit_hrs  = round(tm_hrs + ff_credit, 2)
    overrun_hrs  = round(ff_overrun, 2)
    admin_hrs    = round(admin_hrs, 2)
    ff_total_hrs = round(ff_credit + ff_overrun, 2)  # all FF hours
    util_hrs     = round(tm_hrs + ff_credit, 2)       # T&M + FF within scope
    util_pct     = round(util_hrs    / avail * 100, 2) if avail else None
    overrun_pct  = round(overrun_hrs / avail * 100, 2) if avail else None
    admin_pct    = round(admin_hrs   / avail * 100, 2) if avail else None
    ff_pct       = round(ff_total_hrs/ avail * 100, 2) if avail else None

    total_booked = round(month_ns[month_ns["hours"] > 0]["hours"].sum(), 2)

    def _fmt_hrs(h):
        """Format hours: round to 1dp max, drop trailing zeros.
        e.g. 176.0→176, 167.2→167.2, 176.25→176.25, 6600.10000001→6600.1"""
        if h is None: return "—"
        # Round to 2dp first to kill float noise, then strip trailing zeros
        rounded = round(h, 2)
        s = f"{rounded:.2f}".rstrip('0').rstrip('.')
        return f"{s}h"

    # Pre-calculate WHS so it sits in the same row
    # WHS: use the viewed consultant name (view_as or logged-in)
    _whs_name = view_name if (view_name and not _is_group_view) else selected
    # Normalise "Last, First" to match project_manager format
    _whs_score, _whs_label, _whs_col = consultant_whs(_whs_name, df_drs) if (df_drs is not None and not _is_group_view) else (None, "—", "#718096")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        v   = _fmt_hrs(avail)
        lbl = "Available this month" if avail else "Available hrs (location not mapped)"
        st.markdown(f'<div class="metric-card"><div class="metric-val">{v}</div><div class="metric-lbl">{lbl}<span class="metric-help" data-tip="Total available hours based on consultant location less Bank or Government holidays.">ⓘ</span></div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><div class="metric-val">{_fmt_hrs(total_booked)}</div><div class="metric-lbl">Hours booked this month<span class="metric-help" data-tip="Total hours logged in NetSuite for this period across all project types (Fixed Fee, T&M, and Internal).">ⓘ</span></div></div>', unsafe_allow_html=True)
    with c3:
        if util_pct is not None:
            col = "#27AE60" if util_pct >= 70 else ("#F39C12" if util_pct >= 60 else "#C0392B")
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{col}">{util_pct}%</div><div class="metric-lbl">Util % &nbsp;·&nbsp; {_fmt_hrs(util_hrs)} credited<span class="metric-help" data-tip="Utilization credit hours as a % of Available hours. Credits = T&M hours + Fixed Fee hours within Scope. Overrun hours and Internal/Admin hours are excluded.">ⓘ</span></div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">Util %<span class="metric-help" data-tip="Utilization credit hours as a % of Available hours. Credits = T&M hours + Fixed Fee hours within Scope. Overrun hours and Internal/Admin hours are excluded.">ⓘ</span></div></div>', unsafe_allow_html=True)
    with c4:
        if overrun_pct is not None:
            col = "#C0392B" if overrun_pct > 10 else ("#F39C12" if overrun_pct > 0 else "#718096")
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{col}">{overrun_pct}%</div><div class="metric-lbl">FF overrun % &nbsp;·&nbsp; {_fmt_hrs(overrun_hrs)} over budget<span class="metric-help" data-tip="Fixed Fee hours logged beyond the scoped budget as a % of available hours. A non-zero value means one or more FF projects has exceeded its allocated hours and should be reviewed.">ⓘ</span></div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">FF overrun %<span class="metric-help" data-tip="Fixed Fee hours logged beyond the scoped budget as a % of available hours. A non-zero value means one or more FF projects has exceeded its allocated hours and should be reviewed.">ⓘ</span></div></div>', unsafe_allow_html=True)
    with c5:
        if admin_pct is not None:
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:#718096">{admin_pct}%</div><div class="metric-lbl">Internal % &nbsp;·&nbsp; {_fmt_hrs(admin_hrs)}<span class="metric-help" data-tip="Hours logged against Internal or Admin projects as a % of Available hours. Includes non-billable tasks, internal meetings, PTO and Admin time.">ⓘ</span></div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">Internal %<span class="metric-help" data-tip="Hours logged against Internal or Admin projects as a % of Available hours. Includes non-billable tasks, internal meetings, PTO and Admin time.">ⓘ</span></div></div>', unsafe_allow_html=True)
    with c6:
        if _whs_score is not None:
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_whs_col}">{_whs_score}</div><div class="metric-lbl">WHS &nbsp;·&nbsp; {_whs_label}<span class="metric-help" data-tip="Workload Health Score: a composite score based on number of active projects, phase distribution, overrun count, and stale projects. Higher scores indicate higher risk of consultant overload.">ⓘ</span></div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">WHS<span class="metric-help" data-tip="Workload Health Score: a composite score based on number of active projects, phase distribution, overrun count, and stale projects. Higher scores indicate higher risk of consultant overload.">ⓘ</span></div></div>', unsafe_allow_html=True)

    # UNCONFIGURED FF hours warning — matches Util Report Watch List behaviour
    if ff_unscoped > 0:
        st.warning(f"⚠️ {_fmt_hrs(ff_unscoped)} on FF projects with NO SCOPE DEFINED — see Utilization Report Watch List tab in the downloaded report.")
else:
    if df_ns is None:
        st.info("Upload NS Time Detail in the sidebar to see your utilization snapshot.")
    else:
        _view_label_for_warn = (
            view_name.split(":",1)[1] + " team" if view_name.startswith("REGION:") else
            "All team" if view_name == "ALL" else
            view_name.split(",")[1].strip() if "," in view_name else view_name
        )
        st.warning(f"No time entries found for **{_view_label_for_warn}** in the NS file.")

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── Consultant breakdown table (group views only) ─────────────────────────────
if _is_group_view and not my_ns.empty and "employee" in my_ns.columns:
    st.markdown('<div class="section-label">Team Breakdown — This Month</div>', unsafe_allow_html=True)

    my_ns["date"] = pd.to_datetime(my_ns["date"], errors="coerce")
    _month_ns_all = my_ns[my_ns["date"].dt.strftime("%Y-%m") == month_key].copy()

    _scope_names  = [n for n in (_region_names if _view_name_set else CONSULTANT_DROPDOWN)
                     if get_role(n) != "manager_only"
                     and len(EMPLOYEE_ROLES.get(n, {}).get("products", [])) > 0]
    _region_key   = view_name.split(":",1)[1] if view_name.startswith("REGION:") else None

    # Find leavers who were active in this month and belong to this region
    _month_start = pd.to_datetime(f"{month_key}-01")
    _month_end   = _month_start + pd.offsets.MonthEnd(0)
    _days_in_month = _month_end.day

    _leaver_scope = []
    for _ln, _exit_str in LEAVER_EXIT_DATES.items():
        if _exit_str is None: continue
        _exit_dt = pd.to_datetime(_exit_str)
        if _exit_dt < _month_start: continue          # left before this month
        if _region_key and _gr(_ln) != _region_key: continue  # wrong region
        _leaver_scope.append((_ln, _exit_dt))

    def _build_row(cn, is_leaver=False, exit_dt=None):
        _parts = [p.strip() for p in cn.split(",")]
        _variants = {cn.lower(), _parts[0].lower()}
        if len(_parts) == 2:
            _variants.add(f"{_parts[1]} {_parts[0]}".lower())

        _emp_mask = _month_ns_all["employee"].astype(str).str.strip().str.lower().apply(
            lambda v: v in _variants or any(
                v == nv or v.startswith(nv+" ") or v.endswith(" "+nv) for nv in _variants)
        )
        _emp_rows = _month_ns_all[_emp_mask]

        if _emp_rows.empty:
            _total = 0.0; _ff = 0.0; _tm = 0.0; _admin = 0.0
            _ff_util = 0.0; _ff_over = 0.0
        else:
            _bt    = _emp_rows.get("billing_type", pd.Series(dtype=str)).fillna("").str.strip().str.lower()
            _total = round(_emp_rows["hours"].sum(), 2)
            _ff    = round(_emp_rows[_bt == "fixed fee"]["hours"].sum(), 2)
            _tm    = round(_emp_rows[_bt == "t&m"]["hours"].sum(), 2)
            _admin = round(_emp_rows[_bt == "internal"]["hours"].sum(), 2)

            # Per-consultant FF credit/overrun — mirrors top-level engine exactly
            # Use same rule: anything not internal or T&M = Fixed Fee
            _bt_cn = _emp_rows.get("billing_type", pd.Series(dtype=str)).fillna("").str.strip().str.lower()
            _ff_rows_cn = _emp_rows[(_bt_cn != "internal") & (_bt_cn != "t&m")].sort_values("date") if not _emp_rows.empty else pd.DataFrame()
            _ff_util = 0.0; _ff_over = 0.0
            if not _ff_rows_cn.empty and "project" in _ff_rows_cn.columns:
                _con_cn: dict = {}
                for _, _fr in _ff_rows_cn.iterrows():
                    _fp    = " ".join(str(_fr.get("project","")).split())
                    _ftype = str(_fr.get("project_type","")).strip()
                    _fh    = float(_fr.get("hours", 0) or 0)
                    if _fh <= 0: continue
                    _fm = [(k, float(v)) for k, v in DEFAULT_SCOPE.items()
                           if k.strip().lower() in _ftype.lower()]
                    _fsc = max(_fm, key=lambda x: len(x[0]))[1] if _fm else None
                    if _fsc is None:
                        _ff_util += _fh; continue  # UNCONFIGURED: not counted as overrun
                    _fck = (_fp, _ftype)  # composite key: project + product type (matches top-level)
                    if _fck not in _con_cn:
                        _con_cn[_fck] = prior_htd.get(_fck, prior_htd.get((_fp,), 0.0))
                    _fused = _con_cn[_fck]; _frem = _fsc - _fused
                    if _frem <= 0:
                        _ff_over += _fh
                    elif _fh <= _frem:
                        _ff_util += _fh; _con_cn[_fck] = _fused + _fh
                    else:
                        _ff_util += _frem; _ff_over += _fh - _frem; _con_cn[_fck] = _fsc
            _ff_util = round(_ff_util, 2)
            _ff_over = round(_ff_over, 2)

        _cl = EMPLOYEE_LOCATION.get(cn, "")
        _cr = PS_REGION_OVERRIDE.get(cn, PS_REGION_MAP.get(_cl, ""))
        _avail_full = AVAIL_HOURS.get(_cl, AVAIL_HOURS.get(_cr, {})).get(month_key)
        if isinstance(_avail_full, tuple): _avail_full = _avail_full[0]
        _avail_full = float(_avail_full) if _avail_full else None

        # Prorate if leaver
        if is_leaver and exit_dt is not None and _avail_full:
            _days_worked = min(exit_dt.day, _days_in_month)
            _avail_cn = round(_avail_full * _days_worked / _days_in_month, 2)
        else:
            _avail_cn = _avail_full

        _util_h_cn   = round(_ff_util + _tm, 2)
        _util_pct_cn = round(_util_h_cn / _avail_cn * 100, 1) if _avail_cn and _util_h_cn > 0 else None
        _over_pct_cn = round(_ff_over   / _avail_cn * 100, 1) if _avail_cn and _ff_over > 0 else None
        _int_pct_cn  = round(_admin     / _avail_cn * 100, 1) if _avail_cn and _admin > 0 else None

        _display  = f"{_parts[1].strip()} {_parts[0]}" if len(_parts) == 2 else cn
        if is_leaver and exit_dt:
            _display += " *"

        return {
            "Consultant":    _display,
            "Avail h":       _avail_cn or "—",
            "FF Util h":     _ff_util or "—",
            "FF Overrun h":  _ff_over or "—",
            "T&M h":         _tm or "—",
            "Internal h":    _admin or "—",
            "Util %":        f"{_util_pct_cn}%" if _util_pct_cn is not None else "—",
            "FF Overrun %":  f"{_over_pct_cn}%" if _over_pct_cn is not None else "—",
            "Internal %":    f"{_int_pct_cn}%" if _int_pct_cn is not None else "—",
        }

    _rows = [_build_row(cn) for cn in _scope_names]
    _rows += [_build_row(ln, is_leaver=True, exit_dt=ex) for ln, ex in _leaver_scope]

    if _rows:
        _tbl = pd.DataFrame(_rows)
        st.dataframe(
            _tbl,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Consultant":   st.column_config.TextColumn("Consultant",   width="medium"),
                "Avail h":      st.column_config.TextColumn("Avail h",      width="small"),
                "FF Util h":    st.column_config.TextColumn("FF Util h",    width="small"),
                "FF Overrun h": st.column_config.TextColumn("FF Overrun h", width="small"),
                "T&M h":        st.column_config.TextColumn("T&M h",        width="small"),
                "Internal h":   st.column_config.TextColumn("Internal h",   width="small"),
                "Util %":       st.column_config.TextColumn("Util %",       width="small"),
                "FF Overrun %": st.column_config.TextColumn("FF Overrun %", width="small"),
                "Internal %":   st.column_config.TextColumn("Internal %",   width="small"),
            }
        )
        if _leaver_scope:
            st.caption("* Available hours prorated")

# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Project Snapshot
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">My Projects — Snapshot</div>', unsafe_allow_html=True)

if df_drs is None:
    st.info("Upload SS DRS Export in the sidebar to see your project snapshot.")
elif my_projects.empty:
    st.info("No projects found in the DRS file for your profile.")
else:
    # Phase breakdown (still needed for metric cards)
    _today   = _snap_today; _n7 = _snap_n7; _14 = _snap_14
    _pc = sorted(
        [(("Unassigned" if str(ph) in ("—", "", "nan") else ph), cnt)
         for ph, cnt in _active["phase"].fillna("Unassigned").value_counts().items()
         if cnt > 0],
        key=lambda x: (_pidx_db(x[0]) if x[0] != "Unassigned" else 999)
    )
    snap1, snap2, snap3, snap4, snap5, snap6, snap7 = st.columns(7)
    with snap1:
        st.markdown(f'<div class="metric-card"><div class="metric-val">{len(_active)}</div><div class="metric-lbl">Active Projects</div></div>', unsafe_allow_html=True)
        for ph, cnt in _pc:
            st.markdown(f'<div style="font-size:13px;opacity:.65;padding:1px 0">{cnt} · {str(ph).split(".")[-1].strip()[:22]}</div>', unsafe_allow_html=True)
    with snap2:
        _col = "#27AE60" if len(_gls) > 0 else "inherit"
        st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_col}">{len(_gls)}</div><div class="metric-lbl">Going live this week</div></div>', unsafe_allow_html=True)
        for _, r in _gls.iterrows():
            _cust = str(r.get("project_name",""))
            _cust = _cust.split(" - ")[0].strip() if " - " in _cust else _cust[:28]
            st.markdown(f'<div style="font-size:13px;opacity:.65;padding:1px 0">{_cust[:28]} · {pd.Timestamp(r["go_live_date"]).strftime("%-d %b")}</div>', unsafe_allow_html=True)
    with snap3:
        _col = "#F39C12" if len(_ihc) > 0 else "inherit"
        st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_col}">{len(_ihc)}</div><div class="metric-lbl">In hypercare (week 1)</div></div>', unsafe_allow_html=True)
        for _, r in _ihc.iterrows():
            _cust = str(r.get("project_name",""))
            _cust = _cust.split(" - ")[0].strip() if " - " in _cust else _cust[:28]
            st.markdown(f'<div style="font-size:13px;opacity:.65;padding:1px 0">{_cust[:28]} · {pd.Timestamp(r["go_live_date"]).strftime("%-d %b")}</div>', unsafe_allow_html=True)
    with snap4:
        _col = "#C0392B" if len(_mi) > 0 else "inherit"
        st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_col}">{len(_mi)}</div><div class="metric-lbl">Missing intro email</div></div>', unsafe_allow_html=True)
        if len(_mi) > 0:
            st.markdown('<div style="font-size:13px;opacity:.55">Excl. legacy projects</div>', unsafe_allow_html=True)
    with snap5:
        _col = "#C0392B" if len(_stale) > 0 else "inherit"
        st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_col}">{len(_stale)}</div><div class="metric-lbl">Need re-engagement</div></div>', unsafe_allow_html=True)
        if len(_stale) > 0:
            st.markdown('<div style="font-size:13px;opacity:.55">14+ days inactive</div>', unsafe_allow_html=True)
    def _rag_label(r):
        """Format as 'Customer : Product abbrev' for RAG card list items."""
        _pn  = str(r.get("project_name","") or "")
        _cust = _pn.split(" - ")[0].strip() if " - " in _pn else _pn
        _cust = _cust[:22]
        _pt   = str(r.get("project_type","") or "")
        # Abbreviate product type: strip ZoneApp:/ZoneBill: prefix, keep first word
        _prod = _pt.split(":")[-1].strip() if ":" in _pt else _pt
        _prod = _prod.split()[0] if _prod else ""
        return f"{_cust} : {_prod}" if _prod else _cust

    with snap6:
        _col = "#C0392B" if len(_rag_red) > 0 else "inherit"
        st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_col}">{len(_rag_red)}</div><div class="metric-lbl">Red RAG</div></div>', unsafe_allow_html=True)
        for _, _rr in _rag_red.head(3).iterrows():
            st.markdown(f'<div style="font-size:12px;opacity:.65;padding:1px 0">{_rag_label(_rr)}</div>', unsafe_allow_html=True)
    with snap7:
        _col = "#F39C12" if len(_rag_yellow) > 0 else "inherit"
        st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_col}">{len(_rag_yellow)}</div><div class="metric-lbl">Yellow RAG</div></div>', unsafe_allow_html=True)
        for _, _ry in _rag_yellow.head(3).iterrows():
            st.markdown(f'<div style="font-size:12px;opacity:.65;padding:1px 0">{_rag_label(_ry)}</div>', unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.caption("PS Reporting Tools · Internal use only · Data loaded this session only")
