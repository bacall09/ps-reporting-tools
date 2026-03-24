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
        .action-badge{display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:4px;margin-right:6px;}
        .badge-red   {background:rgba(231,76,60,0.15);color:#E74C3C;}
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
    f"background:rgba(62,207,178,0.15);border:1px solid rgba(62,207,178,0.3);color:#3ECFB2;"
    f"font-size:11px;font-weight:700;letter-spacing:.5px'>{region}</span>"
) if region else ""
_sub_str = " · ".join(_sub_parts)

st.markdown(
    f"<div style='background:#1B2B5E;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;"
    f"font-family:Manrope,sans-serif;position:relative;overflow:hidden'>"
    f"<div style='position:absolute;right:-40px;top:-40px;width:220px;height:220px;border-radius:50%;"
    f"background:radial-gradient(circle,rgba(91,141,239,0.15) 0%,transparent 70%);pointer-events:none'></div>"
    f"<div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;"
    f"color:#3ECFB2;margin-bottom:10px;font-family:Manrope,sans-serif'>Professional Services · Daily Briefing</div>"
    f"<h1 style='color:#fff;margin:0;font-size:28px;font-weight:800;font-family:Manrope,sans-serif;line-height:1.15'>"
    f"{_greeting}, {_my_display}</h1>"
    f"<p style='color:rgba(255,255,255,0.6);margin:8px 0 0;font-size:14px;font-family:Manrope,sans-serif;line-height:1.6'>"
    f"{_sub_str}</p>"
    f"{_region_pill}"
    f"</div>",
    unsafe_allow_html=True
)

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
        _debug_rows = []  # for reconciliation expander
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
                _debug_rows.append({"Date": _date, "Project": _proj, "Type": _ptype,
                                     "Hours": _hrs, "Scope": "—", "Prior HTD": "—",
                                     "Remaining": "—", "Credited": 0.0, "Overrun": 0.0, "Tag": "UNCONFIGURED"})
                continue

            _ck = (_proj, _ptype)  # composite key: project + product type
            if _ck not in _con: _con[_ck] = prior_htd.get(_ck, prior_htd.get((_proj,), 0.0))
            _used = _con[_ck]; _rem = _sc - _used
            if _rem <= 0:
                ff_overrun += _hrs
                _debug_rows.append({"Date": _date, "Project": _proj, "Type": _ptype,
                                     "Hours": _hrs, "Scope": _sc, "Prior HTD": round(_used, 2),
                                     "Remaining": 0.0, "Credited": 0.0, "Overrun": _hrs, "Tag": "OVERRUN"})
            elif _hrs <= _rem:
                ff_credit += _hrs; _con[_ck] = _used + _hrs
                _debug_rows.append({"Date": _date, "Project": _proj, "Type": _ptype,
                                     "Hours": _hrs, "Scope": _sc, "Prior HTD": round(_used, 2),
                                     "Remaining": round(_rem, 2), "Credited": _hrs, "Overrun": 0.0, "Tag": "CREDITED"})
            else:
                ff_credit += _rem; ff_overrun += _hrs - _rem; _con[_ck] = _sc
                _debug_rows.append({"Date": _date, "Project": _proj, "Type": _ptype,
                                     "Hours": _hrs, "Scope": _sc, "Prior HTD": round(_used, 2),
                                     "Remaining": round(_rem, 2), "Credited": round(_rem, 2),
                                     "Overrun": round(_hrs - _rem, 2), "Tag": "PARTIAL"})

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

    # ── Reconciliation debug expander ─────────────────────────────────────────
    if _debug_rows:
        with st.expander(f"🔍 Util reconciliation detail — {len(_debug_rows)} FF rows · compare with Util Report › Detail tab", expanded=False):
            import pandas as _dpd
            _debug_df = _dpd.DataFrame(_debug_rows)
            st.dataframe(_debug_df, hide_index=True, use_container_width=True)
            st.caption(f"T&M: {tm_hrs}h · FF credited: {ff_credit}h · FF overrun: {ff_overrun}h · Internal: {admin_hrs}h · Unscoped: {ff_unscoped}h")
            st.caption("Prior HTD = hours on this project before this period. Scope = configured cap. Compare Tag column with Util Report Detail tab to find discrepancies.")

    def _fmt_hrs(h):
        """Format hours: round to 1dp max, drop trailing zeros.
        e.g. 176.0→176, 167.2→167.2, 176.25→176.25, 6600.10000001→6600.1"""
        if h is None: return "—"
        # Round to 2dp first to kill float noise, then strip trailing zeros
        rounded = round(h, 2)
        s = f"{rounded:.2f}".rstrip('0').rstrip('.')
        return f"{s}h"

    # Pre-calculate WHS so it sits in the same row
    _whs_score, _whs_label, _whs_col = consultant_whs(selected, df_drs) if (df_drs is not None and not _is_group_view) else (None, "—", "#718096")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        v   = _fmt_hrs(avail)
        lbl = "Available this month" if avail else "Available hrs (location not mapped)"
        st.markdown(f'<div class="metric-card"><div class="metric-val">{v}</div><div class="metric-lbl">{lbl}</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><div class="metric-val">{_fmt_hrs(total_booked)}</div><div class="metric-lbl">Hours booked this month</div></div>', unsafe_allow_html=True)
    with c3:
        if util_pct is not None:
            col = "#27AE60" if util_pct >= 70 else ("#F39C12" if util_pct >= 60 else "#E74C3C")
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{col}">{util_pct}%</div><div class="metric-lbl">Util % &nbsp;·&nbsp; {_fmt_hrs(util_hrs)} credited</div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">Util %</div></div>', unsafe_allow_html=True)
    with c4:
        if overrun_pct is not None:
            col = "#E74C3C" if overrun_pct > 10 else ("#F39C12" if overrun_pct > 0 else "#718096")
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{col}">{overrun_pct}%</div><div class="metric-lbl">FF overrun % &nbsp;·&nbsp; {_fmt_hrs(overrun_hrs)} over budget</div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">FF overrun %</div></div>', unsafe_allow_html=True)
    with c5:
        if admin_pct is not None:
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:#718096">{admin_pct}%</div><div class="metric-lbl">Internal % &nbsp;·&nbsp; {_fmt_hrs(admin_hrs)}</div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">Internal %</div></div>', unsafe_allow_html=True)
    with c6:
        if _whs_score is not None:
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_whs_col}">{_whs_score}</div><div class="metric-lbl">WHS &nbsp;·&nbsp; {_whs_label}</div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">WHS</div></div>', unsafe_allow_html=True)

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

            # Per-consultant FF credit/overrun — same logic as top-level engine
            _ff_rows_cn = _emp_rows[_bt == "fixed fee"].sort_values("date") if not _emp_rows.empty else pd.DataFrame()
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
                    if _fp not in _con_cn:
                        _con_cn[_fp] = prior_htd.get(_fp, 0.0)
                    _fused = _con_cn[_fp]; _frem = _fsc - _fused
                    if _frem <= 0:
                        _ff_over += _fh
                    elif _fh <= _frem:
                        _ff_util += _fh; _con_cn[_fp] = _fused + _fh
                    else:
                        _ff_util += _frem; _ff_over += _fh - _frem; _con_cn[_fp] = _fsc
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
    _ioh     = my_projects.get("_on_hold", pd.Series(False, index=my_projects.index)).astype(bool)
    _active  = my_projects[~_ioh].copy()
    _today   = pd.Timestamp.today().normalize()
    _n7      = _today + pd.Timedelta(days=7)
    _14      = _today - pd.Timedelta(days=14)

    # Phase breakdown
    _PHASE_ORDER = ["00. onboarding","01. requirements and design","02. configuration",
                    "03. enablement/training","04. uat","05. prep for go-live",
                    "06. go-live","07. data migration","08. ready for support transition","09. phase 2 scoping"]
    def _pidx_db(p):
        pl = str(p).strip().lower()
        for i, ph in enumerate(_PHASE_ORDER):
            if pl.startswith(ph[:6]) or ph in pl or pl in ph: return i
        return -1

    _pc = sorted(
        [(("Unassigned" if str(ph) in ("—", "", "nan") else ph), cnt)
         for ph, cnt in _active["phase"].fillna("Unassigned").value_counts().items()
         if cnt > 0],
        key=lambda x: (_pidx_db(x[0]) if x[0] != "Unassigned" else 999)
    )

    # Going live this week
    _gls = (_active[_active["go_live_date"].notna() & (_active["go_live_date"] >= _today) & (_active["go_live_date"] <= _n7)]
            .sort_values("go_live_date") if "go_live_date" in _active.columns else pd.DataFrame())

    # In hypercare
    _ihc = (_active[_active["go_live_date"].notna() & (_active["go_live_date"] >= _14) & (_active["go_live_date"] < _today)]
            .sort_values("go_live_date") if "go_live_date" in _active.columns else pd.DataFrame())

    # Missing intro email
    _leg      = _active.get("legacy", pd.Series(False, index=_active.index)).astype(bool)
    _onb_plus = _active["phase"].fillna("").apply(lambda p: _pidx_db(p) >= 0)
    _no_intro = (~_active["ms_intro_email"].notna()) if "ms_intro_email" in _active.columns else pd.Series(True, index=_active.index)
    _mi       = _active[(~_leg) & _onb_plus & _no_intro] if "ms_intro_email" in _active.columns else pd.DataFrame()

    # Unscoped hours banner — matches Util Report behaviour
    if ff_unscoped > 0:
        st.warning(f"⚠️ {round(ff_unscoped, 2)}h on FF projects with NO SCOPE DEFINED — review in Utilization Report.")

    # Re-engagement
    _stale = pd.DataFrame()
    if "days_inactive" in my_projects.columns:
        _stale = _active[_active["days_inactive"].fillna(0) >= 14].sort_values("days_inactive", ascending=False)

    snap1, snap2, snap3, snap4, snap5 = st.columns(5)
    with snap1:
        st.markdown(f'<div class="metric-card"><div class="metric-val">{len(_active)}</div><div class="metric-lbl">Active Projects</div></div>', unsafe_allow_html=True)
        for ph, cnt in _pc:
            st.markdown(f'<div style="font-size:11px;opacity:.65;padding:1px 0">{cnt} · {str(ph).split(".")[-1].strip()[:22]}</div>', unsafe_allow_html=True)
    with snap2:
        _col = "#27AE60" if len(_gls) > 0 else "inherit"
        st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_col}">{len(_gls)}</div><div class="metric-lbl">Going live this week</div></div>', unsafe_allow_html=True)
        for _, r in _gls.iterrows():
            _cust = str(r.get("project_name",""))
            _cust = _cust.split(" - ")[0].strip() if " - " in _cust else _cust[:28]
            st.markdown(f'<div style="font-size:11px;opacity:.65;padding:1px 0">{_cust[:28]} · {pd.Timestamp(r["go_live_date"]).strftime("%-d %b")}</div>', unsafe_allow_html=True)
    with snap3:
        _col = "#F39C12" if len(_ihc) > 0 else "inherit"
        st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_col}">{len(_ihc)}</div><div class="metric-lbl">In hypercare (week 1)</div></div>', unsafe_allow_html=True)
        for _, r in _ihc.iterrows():
            _cust = str(r.get("project_name",""))
            _cust = _cust.split(" - ")[0].strip() if " - " in _cust else _cust[:28]
            st.markdown(f'<div style="font-size:11px;opacity:.65;padding:1px 0">{_cust[:28]} · {pd.Timestamp(r["go_live_date"]).strftime("%-d %b")}</div>', unsafe_allow_html=True)
    with snap4:
        _col = "#E74C3C" if len(_mi) > 0 else "inherit"
        st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_col}">{len(_mi)}</div><div class="metric-lbl">Missing intro email</div></div>', unsafe_allow_html=True)
        if len(_mi) > 0:
            st.markdown('<div style="font-size:11px;opacity:.55">Excl. legacy projects</div>', unsafe_allow_html=True)
    with snap5:
        _col = "#E74C3C" if len(_stale) > 0 else "inherit"
        st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_col}">{len(_stale)}</div><div class="metric-lbl">Need re-engagement</div></div>', unsafe_allow_html=True)
        if len(_stale) > 0:
            st.markdown('<div style="font-size:11px;opacity:.55">14+ days inactive</div>', unsafe_allow_html=True)





st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.caption("PS Reporting Tools · Internal use only · Data loaded this session only")
