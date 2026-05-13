"""
PS Tools — Daily Briefing
Month-to-date utilization snapshot, team breakdown, and re-engagement actions.
"""
import streamlit as st
import pandas as pd
from datetime import date, datetime

st.session_state["current_page"] = "Daily Briefing"

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
        .section-label { font-size: 13px; font-weight: 700; text-transform: uppercase;
                         letter-spacing: 0.8px; color: #4472C4; margin-bottom: 8px; }
        .metric-card   { background: transparent; border: 1px solid rgba(128,128,128,0.2);
                         border-radius: 8px; padding: 16px 20px; margin-bottom: 12px; }
        .metric-val { font-size: 32px; font-weight: 700; color: inherit; }
        .metric-lbl { font-size: 14px; opacity: 0.6; margin-top: 2px; }
        .snap-btn { display:block; width:100%; text-align:center; font-size:11px;
            color:inherit !important; opacity:0.5; padding:2px 0; margin-top:2px;
            border:1px solid rgba(128,128,128,0.25); border-radius:4px;
            text-decoration:none; cursor:pointer; background:transparent; }
        .snap-btn:hover { opacity:0.9; background:rgba(128,128,128,0.08); }
        .metric-help { display:inline-block; margin-left:5px; font-size:13px; opacity:0.55;
                         cursor:help; position:relative; }
        .metric-help:hover::after {
            content: attr(data-tip);
            position:absolute; left:50%; transform:translateX(-50%);
            top:calc(100% + 8px); background:#0E223D; color:#ffffff;
            font-size:13px; font-weight:600; padding:12px 16px; border-radius:8px;
            white-space:normal; width:290px; z-index:99999; line-height:1.7;
            box-shadow:0 4px 20px rgba(0,0,0,0.7); opacity:1 !important;
            letter-spacing:0.15px; border:1px solid rgba(59,158,255,0.3);
            pointer-events:none;
        }
        .action-badge{display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:4px;margin-right:6px;}
        .badge-red   {background:rgba(192,57,43,0.15);color:#C0392B;}
        .badge-amber {background:rgba(243,156,18,0.15);color:#D68910;}
        .badge-blue  {background:rgba(68,114,196,0.15);color:#4472C4;}
        .badge-gray  {background:rgba(128,128,128,0.12);color:inherit;opacity:0.7;}
        .badge-green {background:rgba(39,174,96,0.15);color:#27AE60;}
        .divider{border:none;border-top:1px solid rgba(128,128,128,0.2);margin:20px 0;}
        /* Util metric value colors via data attribute on parent column */
        [data-util-color="green"] [data-testid="stMetricValue"] { color:#27AE60 !important; }
        [data-util-color="amber"] [data-testid="stMetricValue"] { color:#F39C12 !important; }
        [data-util-color="red"]   [data-testid="stMetricValue"] { color:#C0392B !important; }
        [data-util-color="grey"]  [data-testid="stMetricValue"] { color:#718096 !important; }
        button[data-testid="baseButton-secondary"] {
            font-size: 11px !important;
            padding: 2px 6px !important;
            min-height: 24px !important;
            height: 24px !important;
            line-height: 20px !important;
        }
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
_region_names  = []   # always defined regardless of view_name branch
if view_name == "ALL":
    # ALL view: scope to active employees only (excludes leavers + unassigned)
    # Build name variants for every active employee
    _view_name_set = set()
    for n in ACTIVE_EMPLOYEES:
        parts = [p.strip() for p in n.split(",")]
        _view_name_set.add(n.lower())
        _view_name_set.add(parts[0].lower())
        if len(parts) == 2:
            _view_name_set.add(f"{parts[1]} {parts[0]}".lower())
            _view_name_set.add(parts[1].lower())
elif view_name.startswith("REGION:"):
    _region_sel = view_name.split(":", 1)[1]
    # Only include people who actually do product delivery for name matching
    _region_names = [
        n for n in CONSULTANT_DROPDOWN
        if _gr(n) == _region_sel
        and len(EMPLOYEE_ROLES.get(n, {}).get("products", [])) > 0
    ]
    # Include name variants — also add DRS display name aliases from PS_REGION_OVERRIDE
    _view_name_set = set()
    for n in _region_names:
        parts = [p.strip() for p in n.split(",")]
        _view_name_set.add(n.lower())                              # "longalong, santiago"
        _view_name_set.add(parts[0].lower())                       # "longalong" (last name)
        if len(parts) == 2:
            _view_name_set.add(f"{parts[1]} {parts[0]}".lower())  # "santiago longalong"
            _view_name_set.add(parts[1].strip().lower())           # "santiago" (first name)
    # Add DRS display name aliases (e.g. "Caroline Tuazon" for Tuazon, Carol)
    for _alias, _alias_region in PS_REGION_OVERRIDE.items():
        if _alias_region == _region_sel:
            _view_name_set.add(_alias.lower())
            _alias_parts = _alias.strip().split()
            if len(_alias_parts) >= 2:
                _view_name_set.add(_alias_parts[-1].lower())
elif view_name == "ALL_MANAGERS":
    _view_name_set = {n.lower() for n in ACTIVE_EMPLOYEES if get_role(n) == "manager_only"}

def _match_name(val):
    """Match a name value against the current view selection."""
    if _view_name_set is not None:
        v = str(val).strip().lower()
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
            pm_col = df_drs.get("project_manager", pd.Series(dtype="object"))
            _drs_by_who = df_drs[pm_col.apply(lambda v: _match_name(str(v)))]
        else:
            _drs_by_who = df_drs
    else:
        pm_mask = df_drs.get("project_manager", pd.Series(dtype="object")).apply(lambda v: _match_name(str(v)))
        _drs_by_who = df_drs[pm_mask]
        if _drs_by_who.empty and not is_manager(view_name):
            _drs_by_who = df_drs
            st.caption("ℹ️ No PM column matched — showing all projects.")

    if _product_project_types is not None and not _drs_by_who.empty:
        pt_col = _drs_by_who.get("project_type", pd.Series(dtype="object"))
        my_projects = _drs_by_who[pt_col.apply(lambda v: _match_project_type(str(v)))].copy()
    else:
        my_projects = _drs_by_who.copy()

# Filter NS — by employee name, then also by project_type if product filter active
my_ns = pd.DataFrame()
if df_ns is not None and not df_ns.empty:
    if _is_group_view and _view_name_set is None:
        _ns_by_who = df_ns
    else:
        emp_mask = df_ns.get("employee", pd.Series(dtype="object")).apply(lambda v: _match_name(str(v)))
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

_hero = st.empty()
_hero.markdown(
    f"<div style='background:#050D1F;padding:28px 32px 24px;border-radius:10px;margin-bottom:16px;font-family:Manrope,sans-serif;'>"
    f"<div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3B9EFF;margin-bottom:8px;'>Professional Services · Daily Briefing</div>"
    f"<h1 style='color:#fff;margin:0;font-size:26px;font-weight:700;font-family:Manrope,sans-serif;'>{_greeting}, {_my_display}.</h1>"
    f"<p style='color:rgba(255,255,255,0.4);margin:6px 0 0;font-size:13px;'>Loading...</p>"
    f"</div>",
    unsafe_allow_html=True,
)

# ── Pre-compute snapshot data so hero metrics can be populated ────────────────
_ioh     = my_projects.get("_on_hold", pd.Series(False, index=my_projects.index)).astype(bool) if not my_projects.empty else pd.Series(dtype=bool)
_active  = my_projects[~_ioh].copy() if not my_projects.empty else pd.DataFrame()
# ── Deduped project counts for metrics (avoids multi-row DRS inflation) ──
_id_col_dc    = "project_id" if "project_id" in _active.columns else "project_name"
_n_active_dc  = int(_active[_id_col_dc].nunique()) if not _active.empty else 0
_n_onhold_dc  = int(my_projects[_ioh]["project_id"].nunique()) if not my_projects.empty and "project_id" in my_projects.columns else int(_ioh.sum())

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
    _gld_col = "effective_go_live_date" if "effective_go_live_date" in _active.columns else "go_live_date"
    _gls = (_active[_active[_gld_col].notna() & (_active[_gld_col] >= _snap_today) & (_active[_gld_col] <= _snap_n7)]
            .sort_values(_gld_col) if _gld_col in _active.columns else pd.DataFrame())
    _ihc = (_active[_active[_gld_col].notna() & (_active[_gld_col] >= _snap_14) & (_active[_gld_col] < _snap_today)]
            .sort_values(_gld_col) if _gld_col in _active.columns else pd.DataFrame())
    _leg_s    = _active.get("legacy", pd.Series(False, index=_active.index)).astype(bool)
    _onb_plus = _active["phase"].fillna("").apply(lambda p: _pidx_db(p) >= 0)
    _no_intro = (~_active["ms_intro_email"].notna()) if "ms_intro_email" in _active.columns else pd.Series(True, index=_active.index)
    _mi       = _active[(~_leg_s) & _onb_plus & _no_intro] if "ms_intro_email" in _active.columns else pd.DataFrame()
    _stale    = _active[_active["days_inactive"].fillna(0) >= 14].sort_values("days_inactive", ascending=False) if "days_inactive" in _active.columns else pd.DataFrame()
    _rag_red  = pd.DataFrame(); _rag_yellow = pd.DataFrame()
    # Include on-hold projects in RAG — a red on-hold is still a red
    _all_proj = my_projects.copy() if not my_projects.empty else pd.DataFrame()
    if "rag" in _all_proj.columns:
        _rv = _all_proj["rag"].fillna("").astype(str).str.strip().str.lower()
        _rag_red = _all_proj[_rv == "red"]; _rag_yellow = _all_proj[_rv == "yellow"]
    # Projects 9+ and 12+ months from start, not yet at phase 08
    _pre_trans = _active[_active["phase"].fillna("").apply(
        lambda p: _pidx_db(p) < _pidx_db("08. ready for support transition")
    )] if "phase" in _active.columns else _active
    if "start_date" in _pre_trans.columns and not _pre_trans.empty:
        _sd = pd.to_datetime(_pre_trans["start_date"], errors="coerce")
        _months_active = (_snap_today - _sd).dt.days / 30.44
        _proj_12mo = _pre_trans[_months_active >= 12].copy()
        _proj_9mo  = _pre_trans[(_months_active >= 9) & (_months_active < 12)].copy()
    else:
        _proj_9mo = _proj_12mo = pd.DataFrame()
else:
    _gls = _ihc = _mi = _stale = _rag_red = _rag_yellow = _proj_9mo = _proj_12mo = pd.DataFrame()

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

_oh_count = _n_onhold_dc
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
        _para_attn += f"{'With' if _para_attn else ''} {_proj_list(_gls)} {'is' if len(_gls)==1 else 'are'} going live this week — confirm cutover readiness before the end of the day. "
    if len(_mi) > 0:
        _para_attn += f"{len(_mi)} project{'s are' if len(_mi)>1 else ' is'} missing an intro email date — if already sent, logging them is a quick close. "
    if len(_stale) > 0:
        _stale_top  = _stale.iloc[0]
        _stale_name = str(_stale_top.get("project_name","")).split(" - ")[0].strip()[:28]
        _stale_days = int(_stale_top.get("days_inactive",0))
        if len(_stale) == 1:
            _para_attn += f"{_stale_name} hasn't had contact in {_stale_days} days — a short check-in would re-establish momentum. "
        else:
            _para_attn += f"{len(_stale)} projects are overdue for outreach, with {_stale_name} leading at {_stale_days} days inactive — use Customer Engagement to draft messages. "
    if len(_proj_12mo) > 0:
        _12mo_names = ", ".join(str(r.get("project_name","")).split(" - ")[0].strip()[:22] for _, r in _proj_12mo.head(2).iterrows())
        _extra = f" and {len(_proj_12mo)-2} more" if len(_proj_12mo) > 2 else ""
        _para_attn += f"{len(_proj_12mo)} project{'s have' if len(_proj_12mo)>1 else ' has'} been active for 12+ months without reaching support transition ({_12mo_names}{_extra}) — these likely need an escalation review."

    _para_reminder = ""
    if len(_gls) > 0 or len(_ihc) > 0:
        _r_parts = []
        if len(_gls) > 0:
            _r_parts.append(f"{len(_gls)} customer{'s' if len(_gls)>1 else ''} going live this week")
        if len(_ihc) > 0:
            _r_parts.append(f"{len(_ihc)} in week-one hypercare")
        _para_reminder = " and ".join(_r_parts) + " — a proactive check-in today would be timely."

    _para_quick = ""
    if len(_proj_9mo) > 0:
        _para_quick += f"{len(_proj_9mo)} project{'s are' if len(_proj_9mo)>1 else ' is'} approaching the 12-month mark (9–12 months active) — worth a proactive check on timeline and transition plan. "
    if len(_rag_yellow) > 0:
        _para_quick += f"{_proj_list(_rag_yellow)} {'are' if len(_rag_yellow)>1 else 'is'} at Yellow RAG — a quick review of blockers now could prevent escalation."

    _oh_count = _n_onhold_dc
    _para_house = ""
    if _oh_count > 0:
        _para_house = f"{_oh_count} project{'s are' if _oh_count>1 else ' is'} on hold. Make sure each has an On Hold Reason and Responsible for Delay recorded — these are flagged in DRS Health Check if missing."

    _bhtml = """<div style='border-radius:8px;border:1px solid rgba(59,158,255,0.2);overflow:hidden;margin-bottom:16px;font-family:Manrope,sans-serif'>
  <div style='background:rgba(59,158,255,0.07);padding:10px 20px;border-bottom:1px solid rgba(59,158,255,0.15);display:flex;align-items:center;gap:10px'>
<span style='font-size:13px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3B9EFF'>This Week&#39;s Focus</span>
<span style='font-size:11px;color:var(--text-color,inherit);opacity:0.35;margin-left:auto'>Rule-based · AI briefings coming soon</span>
  </div>
  <div style='padding:18px 22px;display:flex;flex-direction:column;gap:14px'>"""

    if _para_reminder:
        _bhtml += f"""<div style='display:flex;gap:14px;align-items:flex-start'>
<div style='flex-shrink:0;width:3px;background:#27AE60;border-radius:2px;align-self:stretch;min-height:36px'></div>
<div>
  <div style='font-size:13px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#27AE60;margin-bottom:5px'>Friendly reminder</div>
  <div style='font-size:13px;color:inherit;line-height:1.7'>{_para_reminder.strip()}</div>
</div>
  </div>"""

    if _para_attn:
        _sep_r = "<div style='height:1px;background:rgba(128,128,128,0.12)'></div>" if _para_reminder else ""
        _bhtml += f"""{_sep_r}<div style='display:flex;gap:14px;align-items:flex-start'>
<div style='flex-shrink:0;width:3px;background:#C0392B;border-radius:2px;align-self:stretch;min-height:36px'></div>
<div>
  <div style='font-size:13px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#C0392B;margin-bottom:5px'>Needs attention</div>
  <div style='font-size:13px;color:inherit;line-height:1.7'>{_para_attn.strip()}</div>
</div>
  </div>"""

    if _para_quick:
        _sep = "<div style='height:1px;background:rgba(128,128,128,0.12)'></div>" if (_para_attn or _para_reminder) else ""
        _bhtml += f"""{_sep}<div style='display:flex;gap:14px;align-items:flex-start'>
<div style='flex-shrink:0;width:3px;background:#3B9EFF;border-radius:2px;align-self:stretch;min-height:36px'></div>
<div>
  <div style='font-size:13px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#3B9EFF;margin-bottom:5px'>Quick wins this week</div>
  <div style='font-size:13px;color:inherit;line-height:1.7'>{_para_quick.strip()}</div>
</div>
  </div>"""

    if _para_house:
        _sep2 = "<div style='height:1px;background:rgba(128,128,128,0.12)'></div>" if (_para_attn or _para_quick) else ""
        _bhtml += f"""{_sep2}<div style='display:flex;gap:14px;align-items:flex-start'>
<div style='flex-shrink:0;width:3px;background:rgba(128,128,128,0.3);border-radius:2px;align-self:stretch;min-height:36px'></div>
<div>
  <div style='font-size:13px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:inherit;opacity:0.45;margin-bottom:5px'>Housekeeping</div>
  <div style='font-size:13px;color:inherit;opacity:0.6;line-height:1.7'>{_para_house.strip()}</div>
</div>
  </div>"""

    _bhtml += "</div></div>"
    # Briefing text rendered inside _col_brief below


st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Utilization
# ══════════════════════════════════════════════════════════════════════════════
# This Month — Utilization section label hidden (replaced by hero banner + donut panel)
# st.markdown('<div class="section-label">This Month — Utilization</div>', unsafe_allow_html=True)

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
_whs_score = None; _whs_label = "—"; _whs_col = "#718096"
_wk_util_pct = None; _wk_total = None; _wk_billable_h = None
_gl14_col = None; _gl14 = pd.DataFrame()
_month_n = 0
_week_start = pd.Timestamp.today().normalize() - pd.Timedelta(days=pd.Timestamp.today().weekday())

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

        # Build prior_htd — keyed by project_id (preferred) then project name
        # prior = max(hours_to_date) - sum(period hours) per project
        prior_htd: dict = {}
        if "hours_to_date" in month_ns.columns:
            _has_pid = "project_id" in month_ns.columns
            _grp_col = "project_id" if _has_pid else "project"
            for _proj_key, _grp in month_ns.groupby(_grp_col):
                _key = str(_proj_key).strip()
                try:
                    _max_htd    = float(_grp["hours_to_date"].dropna().astype(float).max() or 0)
                    _period_hrs = float(_grp["hours"].dropna().astype(float).sum() or 0)
                    prior_htd[_key] = max(0.0, _max_htd - _period_hrs)
                except Exception:
                    prior_htd[_key] = 0.0
            # Also index by project name for fallback lookups
            if _has_pid and "project" in month_ns.columns:
                for _pid, _grp in month_ns.groupby("project_id"):
                    _pname = " ".join(str(_grp["project"].iloc[0]).split())
                    if _pname and _pname not in prior_htd:
                        prior_htd[_pname] = prior_htd.get(str(_pid).strip(), 0.0)

        import re as _re_db
        _con: dict = {}
        for _, _r in ff_rows.iterrows():
            _proj  = " ".join(str(_r.get("project","")).split())
            _ptype = str(_r.get("project_type","")).strip()
            _hrs   = float(_r.get("hours", 0) or 0)
            if _hrs <= 0: continue

            # Premium: check Time Item SKU for IMPL10/IMPL20 first (matches Util Report)
            _ptype_lower = _ptype.strip().lower()
            if "premium" in _ptype_lower:
                _sku = str(_r.get("time_item_sku", "") or "")
                _sku_nums = _re_db.findall(r"IMPL(\d+)", _sku.upper())
                if _sku_nums:
                    _sc = float(_sku_nums[0])
                else:
                    _proj_name = str(_r.get("project","") or "")
                    _name_nums = _re_db.findall(r"(?<![\d])(10|20)(?![\d])", _proj_name)
                    _sc = float(_name_nums[0]) if _name_nums else None
            else:
                _m  = [(k, float(v)) for k, v in DEFAULT_SCOPE.items()
                       if k.strip().lower() in _ptype_lower]
                _sc = max(_m, key=lambda x: len(x[0]))[1] if _m else None

            if _sc is None:
                ff_unscoped += _hrs
                continue

            # Use project_id as key where available, fall back to project name
            _pid_r = str(_r.get("project_id","")).strip()
            _pkey  = _pid_r if _pid_r and _pid_r != "nan" else _proj
            if _pkey not in _con:
                _con[_pkey] = prior_htd.get(_pkey, prior_htd.get(_proj, 0.0))
            _used = _con[_pkey]; _rem = _sc - _used
            if _rem <= 0:
                ff_overrun += _hrs
            elif _hrs <= _rem:
                ff_credit += _hrs; _con[_pkey] = _used + _hrs
            else:
                ff_credit += _rem; ff_overrun += _hrs - _rem; _con[_pkey] = _sc

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

    # ── Weekly utilization (Mon–Sun current week) ─────────────────────────────
    _week_start = today - pd.Timedelta(days=today.weekday())  # Monday
    _week_end   = _week_start + pd.Timedelta(days=6)           # Sunday
    _UTIL_TARGET_WK = 70.0
    _week_avail = avail / (pd.bdate_range(today.replace(day=1),
                           (today.replace(day=28)+pd.Timedelta(days=4)).replace(day=1)-pd.Timedelta(days=1)).size
                          ) * len(pd.bdate_range(_week_start, min(_week_end, today))) if avail else None
    if not _is_group_view and df_ns is not None:
        _wk_ns  = my_ns[
            (pd.to_datetime(my_ns["date"], errors="coerce") >= pd.Timestamp(_week_start)) &
            (pd.to_datetime(my_ns["date"], errors="coerce") <= pd.Timestamp(_week_end))
        ] if "date" in my_ns.columns else pd.DataFrame()
        _wk_ff_hrs  = float(_wk_ns[_wk_ns.get("billing_type", pd.Series(dtype=str)).str.lower().str.strip() == "fixed fee"]["hours"].sum()) if not _wk_ns.empty and "billing_type" in _wk_ns.columns else 0.0
        _wk_tm_hrs  = float(_wk_ns[_wk_ns.get("billing_type", pd.Series(dtype=str)).str.lower().str.strip() == "t&m"]["hours"].sum()) if not _wk_ns.empty and "billing_type" in _wk_ns.columns else 0.0
        _wk_total   = round(_wk_ff_hrs + _wk_tm_hrs, 1)
        _wk_bdays_month = len(pd.bdate_range(today.replace(day=1),
                             (today.replace(day=28)+pd.Timedelta(days=4)).replace(day=1)-pd.Timedelta(days=1)))
        _wk_avail_h    = round(float(avail) / _wk_bdays_month * 5, 1) if avail and _wk_bdays_month else 40.0
        _wk_full_h     = _wk_avail_h
        _wk_billable_h = round(_wk_full_h * 0.70, 1)
        _wk_util_pct   = round(_wk_total / _wk_billable_h * 100) if _wk_billable_h else None
    elif _is_group_view and df_ns is not None and not my_ns.empty:
        # Group view — aggregate team weekly util
        _wk_ns_g = my_ns[
            (pd.to_datetime(my_ns["date"], errors="coerce") >= pd.Timestamp(_week_start)) &
            (pd.to_datetime(my_ns["date"], errors="coerce") <= pd.Timestamp(_week_end))
        ] if "date" in my_ns.columns else pd.DataFrame()
        if not _wk_ns_g.empty and "billing_type" in _wk_ns_g.columns:
            _wk_total = round(float(
                _wk_ns_g[_wk_ns_g["billing_type"].str.lower().str.strip().isin(["fixed fee","t&m"])]["hours"].sum()
            ), 1)
        else:
            _wk_total = 0.0
        _wk_bdays_month = len(pd.bdate_range(today.replace(day=1),
                             (today.replace(day=28)+pd.Timedelta(days=4)).replace(day=1)-pd.Timedelta(days=1)))
        _wk_full_h     = round(float(avail) / _wk_bdays_month * 5, 1) if avail and _wk_bdays_month else 40.0
        _wk_billable_h = round(_wk_full_h * 0.70, 1)
        _wk_util_pct   = round(_wk_total / _wk_billable_h * 100) if _wk_billable_h else None
    else:
        _wk_total = None; _wk_util_pct = None; _wk_full_h = None; _wk_billable_h = None

    # ── Overrun MTD ───────────────────────────────────────────────────────────
    _overrun_mtd = overrun_hrs  # already computed above

    # ── Week number and quarter ───────────────────────────────────────────────
    _week_num = today.isocalendar()[1]
    _month_n  = today.month
    _fy_q     = f"Q{(((_month_n - 1) // 3) + 1)} FY{str(today.year)[2:]}"
    _view_sub_hero = ""
    if view_name == "ALL":
        _view_sub_hero = " · Viewing: All team"
    elif view_name.startswith("REGION:"):
        _view_sub_hero = f" · Viewing: {view_name.split(':',1)[1]} team"
    elif view_name == "ALL_MANAGERS":
        _view_sub_hero = " · Viewing: Managers"
    elif view_name != selected:
        _vd = view_name.split(",")[1].strip() if "," in view_name else view_name
        _view_sub_hero = f" · Viewing: {_vd}"
    _date_line = f"{today.strftime('%A')} · {today.strftime('%B %-d')} · {_fy_q} — Week {_week_num}{_view_sub_hero}"

    # ── Summary sentence ──────────────────────────────────────────────────────
    _ss_parts = []
    _n_due_wk = len([p for p in _p1 + _p2 + _p3 if p]) if not _is_group_view else 0
    _n_ms_14  = len(_mi) + (len(_gls) if not _gls.empty else 0)
    _n_intro  = len(_mi) if not _mi.empty else 0
    if _n_ms_14 > 0:
        _ss_parts.append(f"<b style='color:var(--color-text-primary);font-weight:600;'>{_n_ms_14} milestone{'s' if _n_ms_14 > 1 else ''} in the next 14 days</b>")
    if _n_intro > 0:
        _ss_parts.append(f"<b style='color:var(--color-text-primary);font-weight:600;'>{_n_intro} intro email{'s' if _n_intro > 1 else ''} pending</b>")
    if len(_rag_red) > 0:
        _ss_parts.append(f"<b style='color:#E24B4A;font-weight:600;'>{len(_rag_red)} red RAG project{'s' if len(_rag_red) > 1 else ''}</b>")
    _summary_str = (
        "You have " + ", ".join(_ss_parts[:-1]) + (" and " if len(_ss_parts) > 1 else "") + _ss_parts[-1] + "."
        if _ss_parts else "All clear — no urgent items this week."
    )
    # Bold the key numbers/phrases in summary
    import re as _re
    _summary_str = _re.sub(r'<b>(\d+[^<]*?)</b>', r"<b style='color:rgba(255,255,255,0.9);'></b>", _summary_str)

    # ── Go-lives in next 14 days ──────────────────────────────────────────────
    _gl14_col = "effective_go_live_date" if "effective_go_live_date" in _active.columns else "go_live_date"
    _gl14 = (_active[
        _active[_gl14_col].notna() &
        (_active[_gl14_col] >= pd.Timestamp(_snap_today)) &
        (_active[_gl14_col] <= pd.Timestamp(_snap_today + pd.Timedelta(days=14)))
    ].sort_values(_gl14_col) if _gl14_col in _active.columns and not _active.empty else pd.DataFrame())
    _n_gl14 = int(_gl14["project_id"].nunique()) if "project_id" in _gl14.columns else len(_gl14)
    _gl14_sub = ""
    if _n_gl14 > 0 and not _gl14.empty:
        _first_gl = _gl14.iloc[0]
        _gl_cust  = str(_first_gl.get("project_name","")).split(" - ")[0][:22]
        _gl_date  = pd.Timestamp(_first_gl.get(_gl14_col)).strftime("%-d %b")
        _gl14_sub = f"Next: {_gl_cust} · {_gl_date}"

    # ── Overrun week amount from weekly NS slice ──────────────────────────────
    _wk_overrun = round(overrun_hrs, 1)  # MTD total — weekly split not available at this point

    # ── Pacing badge ──────────────────────────────────────────────────────────
    if _wk_util_pct is not None:
        _wk_gap = _wk_util_pct - _UTIL_TARGET_WK
        if _wk_gap >= 0:
            _pace_badge = f"<span style='font-size:11px;padding:2px 7px;border-radius:4px;background:rgba(99,153,34,0.25);color:#97C459;font-weight:500;'>+{round(_wk_gap)}% ahead</span>"
        else:
            _pace_badge = f"<span style='font-size:11px;padding:2px 7px;border-radius:4px;background:rgba(226,75,74,0.25);color:#F09595;font-weight:500;'>{round(_wk_gap)}% behind</span>"
    else:
        _pace_badge = ""

    # ── Overrun display ───────────────────────────────────────────────────────
    _overrun_color          = "#F09595" if _wk_overrun > 0 else "rgba(255,255,255,0.85)"
    _overrun_color_adaptive = "#E24B4A" if _wk_overrun > 0 else "var(--color-text-primary)"
    _overrun_val   = f"{_wk_overrun}h" if _wk_overrun else "—"
    _overrun_sub   = "MTD overrun" if _wk_overrun else "no overrun this month"
    _n_overrun_proj = sum(1 for _ in [True] if _wk_overrun > 0)
    _overrun_badge = f"<span style='font-size:11px;padding:2px 7px;border-radius:4px;background:rgba(226,75,74,0.2);color:#F09595;font-weight:500;'>{_n_overrun_proj} project{'s' if _n_overrun_proj != 1 else ''}</span>" if _wk_overrun > 0 else ""
    _overrun_badge_adaptive = _overrun_badge  # same style

    # ── Render hero ───────────────────────────────────────────────────────────
    _hero.markdown(
        f"<div style='background:#050D1F;padding:28px 32px 24px;border-radius:10px;margin-bottom:16px;font-family:Manrope,sans-serif;'>"
        f"<div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3B9EFF;margin-bottom:8px;'>Professional Services · Daily Briefing</div>"
        f"<div style='font-size:13px;font-weight:500;color:#08A9B7;margin-bottom:6px;'>{_date_line}</div>"
        f"<h1 style='color:#fff;margin:0;font-size:26px;font-weight:700;font-family:Manrope,sans-serif;line-height:1.2'>{_greeting}, {_my_display}.</h1>"
        f"<p style='color:rgba(255,255,255,0.55);margin:6px 0 0;font-size:13px;font-family:Manrope,sans-serif;line-height:1.6'>{_summary_str}</p>"
        f"<div style='display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin-top:18px;padding-top:16px;border-top:0.5px solid rgba(255,255,255,0.1);'>"
        f"<div><div style='font-size:10px;text-transform:uppercase;letter-spacing:0.6px;color:rgba(255,255,255,0.4);margin-bottom:4px;'>Week utilization</div>"
        f"<div style='font-size:26px;font-weight:600;color:#fff;line-height:1.1;'>{_wk_util_pct}%</div>"
        f"<div style='font-size:12px;color:var(--color-text-secondary);margin-top:3px;'>"
        f"{f'{_wk_total}h / {_wk_billable_h}h billable · {_wk_full_h}h total' if _wk_total is not None else 'No time data this week'}</div>"
        f"<div style='margin-top:4px;'>{_pace_badge}</div></div>"
        f"<div><div style='font-size:11px;text-transform:uppercase;letter-spacing:0.6px;color:rgba(255,255,255,0.5);margin-bottom:5px;font-weight:500;'>Open projects</div>"
        f"<div style='font-size:26px;font-weight:600;color:#fff;line-height:1.1;'>{_n_active_dc}</div>"
        f"<div style='font-size:12px;color:rgba(255,255,255,0.45);margin-top:3px;'>{_n_onhold_dc} on hold · {int(_n_active_dc + _n_onhold_dc)} assigned total</div></div>"
        f"<div><div style='font-size:10px;text-transform:uppercase;letter-spacing:0.6px;color:rgba(255,255,255,0.4);margin-bottom:4px;'>Go-lives next 14d</div>"
        f"<div style='font-size:26px;font-weight:600;color:#fff;line-height:1.1;'>{_n_gl14}</div>"
        f"<div style='font-size:12px;color:rgba(255,255,255,0.45);margin-top:3px;'>{_gl14_sub if _gl14_sub else 'None scheduled'}</div></div>"
        f"<div><div style='font-size:10px;text-transform:uppercase;letter-spacing:0.6px;color:rgba(255,255,255,0.4);margin-bottom:4px;'>FF overrun</div>"
        f"<div style='font-size:26px;font-weight:600;color:{_overrun_color};line-height:1.1;'>{_overrun_val}</div>"
        f"<div style='font-size:12px;color:rgba(255,255,255,0.45);margin-top:3px;'>{_overrun_sub}</div>"
        f"{_overrun_badge}</div>"
        f"</div></div>",
        unsafe_allow_html=True,
    )

    # [hidden] c1, c2, c3, c4, c5, c6 = st.columns(6)
    # [hidden] with c1:
    # [hidden] _lbl = "Available this month" if avail else "Available hrs (location not mapped)"
    # [hidden] st.markdown(f'<div class="metric-card"><div class="metric-val">{_fmt_hrs(avail)}</div><div class="metric-lbl">{_lbl} <span class="metric-help" data-tip="Total available hours based on consultant location less Bank or Government holidays.">ⓘ</span></div></div>', unsafe_allow_html=True)
    # [hidden] with c2:
    # [hidden] st.markdown(f'<div class="metric-card"><div class="metric-val">{_fmt_hrs(total_booked)}</div><div class="metric-lbl">Hours booked this month <span class="metric-help" data-tip="Total hours logged in NetSuite for this period across all project types (Fixed Fee, T&M, and Internal).">ⓘ</span></div></div>', unsafe_allow_html=True)
    # [hidden] with c3:
    # [hidden] if util_pct is not None:
            # Pacing-aware colouring — compare MTD util against pro-rated target
    # [hidden] _UTIL_TARGET = 70.0
    # [hidden] _month_start = today.replace(day=1)
    # [hidden] _month_end   = (today.replace(day=28) + pd.Timedelta(days=4)).replace(day=1) - pd.Timedelta(days=1)
    # [hidden] _days_total  = len(pd.bdate_range(_month_start, _month_end))
    # [hidden] _days_elapsed = len(pd.bdate_range(_month_start, today))
    # [hidden] _pacing_pct  = round(_UTIL_TARGET * _days_elapsed / _days_total, 1) if _days_total else _UTIL_TARGET
    # [hidden] _gap         = util_pct - _pacing_pct
    # [hidden] if _gap >= 0:
    # [hidden] _ucol      = "#27AE60"
    # [hidden] _pace_tag  = f'<div style="margin-top:5px;display:inline-block;font-size:9px;font-weight:700;padding:1px 6px;border-radius:8px;letter-spacing:.8px;background:rgba(39,174,96,.15);color:#27AE60">On pace · target {_pacing_pct}%</div>'
    # [hidden] elif _gap >= -10:
    # [hidden] _ucol      = "#F39C12"
    # [hidden] _pace_tag  = f'<div style="margin-top:5px;display:inline-block;font-size:9px;font-weight:700;padding:1px 6px;border-radius:8px;letter-spacing:.8px;background:rgba(243,156,18,.15);color:#F39C12">Behind pace · target {_pacing_pct}%</div>'
    # [hidden] else:
    # [hidden] _ucol      = "#C0392B"
    # [hidden] _pace_tag  = f'<div style="margin-top:5px;display:inline-block;font-size:9px;font-weight:700;padding:1px 6px;border-radius:8px;letter-spacing:.8px;background:rgba(192,57,43,.15);color:#C0392B">Behind target · goal {_UTIL_TARGET}%</div>'
    # [hidden] st.markdown(
    # [hidden] f'<div class="metric-card"><div class="metric-val" style="color:{_ucol}">{util_pct}%</div>' +
    # [hidden] f'<div class="metric-lbl">Util % · {_fmt_hrs(util_hrs)} credited <span class="metric-help" data-tip="Utilization credit hours as a % of Available hours. Colour reflects pacing: compares MTD util against the pro-rated {_UTIL_TARGET}% target for how far through the month we are.">ⓘ</span></div>' +
    # [hidden] f'{_pace_tag}</div>',
    # [hidden] unsafe_allow_html=True
    # [hidden] )
    # [hidden] else:
    # [hidden] st.metric("Util %", "—", help="Utilization credit hours as a % of Available hours.")
    # [hidden] with c4:
    # [hidden] if overrun_pct is not None:
    # [hidden] _ocol = "#C0392B" if overrun_pct > 10 else ("#F39C12" if overrun_pct > 0 else "#718096")
    # [hidden] st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_ocol}">{overrun_pct}%</div><div class="metric-lbl">FF overrun % · {_fmt_hrs(overrun_hrs)} over <span class="metric-help" data-tip="Fixed Fee hours logged beyond the scoped budget as a % of available hours. A non-zero value means one or more FF projects has exceeded its allocated hours and should be reviewed.">ⓘ</span></div></div>', unsafe_allow_html=True)
    # [hidden] else:
    # [hidden] st.metric("FF overrun %", "—", help="Fixed Fee hours logged beyond the scoped budget as a % of available hours.")
    # [hidden] with c5:
    # [hidden] if admin_pct is not None:
    # [hidden] st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:#718096">{admin_pct}%</div><div class="metric-lbl">Internal % · {_fmt_hrs(admin_hrs)} <span class="metric-help" data-tip="Hours logged against Internal or Admin projects as a % of Available hours. Includes non-billable tasks, internal meetings, PTO and Admin time.">ⓘ</span></div></div>', unsafe_allow_html=True)
    # [hidden] else:
    # [hidden] st.metric("Internal %", "—", help="Hours logged against Internal or Admin projects as a % of Available hours.")
    # [hidden] with c6:
    # [hidden] if _whs_score is not None:
    # [hidden] st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{_whs_col}">{_whs_score}</div><div class="metric-lbl">WHS · {_whs_label} <span class="metric-help" data-tip="Workload Health Score: a composite score based on number of active projects, phase distribution, overrun count, and stale projects. Higher scores indicate higher risk of consultant overload.">ⓘ</span></div></div>', unsafe_allow_html=True)
    # [hidden] else:
    # [hidden] st.metric("WHS", "—", help="Workload Health Score: a composite score based on number of active projects, phase distribution, overrun count, and stale projects.")
    # [hidden]     # UNCONFIGURED FF hours warning — matches Util Report Watch List behaviour
    # [hidden] if ff_unscoped > 0:
    # [hidden] st.warning(f"⚠️ {_fmt_hrs(ff_unscoped)} on FF projects with NO SCOPE DEFINED — see Utilization Report Watch List tab in the downloaded report.")
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

# ── Consultant breakdown table (group views only) ─────────────────────────────
if _is_group_view and not my_ns.empty and "employee" in my_ns.columns:
    try:
        st.markdown('<div class="section-label">Team Breakdown — This Week</div>', unsafe_allow_html=True)

        my_ns["date"] = pd.to_datetime(my_ns["date"], errors="coerce")
        _month_ns_all = my_ns[my_ns["date"].dt.strftime("%Y-%m") == month_key].copy()

        # Weekly NS slice (Mon–Sun current week) — used for hours columns
        _wk_start_bd = today - pd.Timedelta(days=today.weekday())
        _wk_end_bd   = _wk_start_bd + pd.Timedelta(days=6)
        _wk_ns_all   = my_ns[
            (my_ns["date"] >= pd.Timestamp(_wk_start_bd)) &
            (my_ns["date"] <= pd.Timestamp(_wk_end_bd))
        ].copy()

        # Business days in month and this week — for weekly avail proration
        _bdays_month = len(pd.bdate_range(
            today.replace(day=1),
            (today.replace(day=28) + pd.Timedelta(days=4)).replace(day=1) - pd.Timedelta(days=1)
        ))

        _region_key  = view_name.split(":",1)[1] if view_name.startswith("REGION:") else None
        _month_start = pd.to_datetime(f"{month_key}-01")
        _month_end   = _month_start + pd.offsets.MonthEnd(0)
        _days_in_month = _month_end.day

        # ── Build set of DRS project_manager names for manager-inclusion check ──
        # Managers are included in the table only if they:
        #   (a) appear as project_manager on any DRS row, OR
        #   (b) have NS time entries this month against non-internal billing types
        _drs_pm_names: set = set()
        if df_drs is not None and not df_drs.empty and "project_manager" in df_drs.columns:
            for _pm_val in df_drs["project_manager"].dropna().astype(str):
                _drs_pm_names.add(_pm_val.strip().lower())
                _pm_parts = [p.strip() for p in _pm_val.split(",")]
                if len(_pm_parts) == 2:
                    _drs_pm_names.add(f"{_pm_parts[1]} {_pm_parts[0]}".lower())
                    _drs_pm_names.add(_pm_parts[0].lower())

        def _is_manager_included(cn: str) -> bool:
            """Return True if a manager-role consultant should appear in the table."""
            _p = [p.strip() for p in cn.split(",")]
            _v = {cn.lower(), _p[0].lower()}
            if len(_p) == 2:
                _v.add(f"{_p[1]} {_p[0]}".lower())
            # (a) PM on a DRS project
            if any(v in _drs_pm_names for v in _v):
                return True
            # (b) has non-internal time this month
            if not _month_ns_all.empty and "billing_type" in _month_ns_all.columns:
                _emp_m = _month_ns_all["employee"].astype(str).str.strip().str.lower().apply(
                    lambda x: x in _v or any(
                        x == nv or x.startswith(nv+" ") or x.endswith(" "+nv) for nv in _v)
                )
                _bt_m = _month_ns_all[_emp_m]["billing_type"].fillna("").str.strip().str.lower()
                if (_bt_m != "internal").any():
                    return True
            return False

        # Base scope: exclude manager_only always; include managers only if they pass the check above
        _base_pool = _region_names if _view_name_set else CONSULTANT_DROPDOWN
        _scope_names = []
        for _cn in _base_pool:
            _cn_role = get_role(_cn)
            if _cn_role == "manager_only":
                continue
            if _cn_role == "manager" and not _is_manager_included(_cn):
                continue
            if len(EMPLOYEE_ROLES.get(_cn, {}).get("products", [])) == 0 and _cn_role != "manager":
                continue
            _scope_names.append(_cn)

        # Leavers active this month in this region
        _leaver_scope = []
        for _ln, _exit_str in LEAVER_EXIT_DATES.items():
            if _exit_str is None: continue
            _exit_dt = pd.to_datetime(_exit_str)
            if _exit_dt < _month_start: continue
            if _region_key and _gr(_ln) != _region_key: continue
            _leaver_scope.append((_ln, _exit_dt))

        # ── RAG lookup: per-consultant red/yellow count from DRS ──────────────
        # Keyed on project_manager name variants → {red: int, yellow: int}
        _rag_by_pm: dict = {}
        if df_drs is not None and not df_drs.empty and "rag" in df_drs.columns and "project_manager" in df_drs.columns:
            for _, _dr in df_drs.iterrows():
                _pm_raw = str(_dr.get("project_manager", "") or "").strip()
                _rv     = str(_dr.get("rag", "") or "").strip().lower()
                if not _pm_raw or _rv not in ("red", "yellow"):
                    continue
                # Normalise PM name to "First Last" key for lookup
                _pm_p = [p.strip() for p in _pm_raw.split(",")]
                _pm_key = f"{_pm_p[1]} {_pm_p[0]}".lower() if len(_pm_p) == 2 else _pm_raw.lower()
                if _pm_key not in _rag_by_pm:
                    _rag_by_pm[_pm_key] = {"red": 0, "yellow": 0}
                _rag_by_pm[_pm_key][_rv] += 1
                # Also index by "Last, First" and last-name-only for robust matching
                _rag_by_pm[_pm_raw.lower()] = _rag_by_pm[_pm_key]
                if _pm_p:
                    _rag_by_pm[_pm_p[0].lower()] = _rag_by_pm[_pm_key]

        def _get_rag_counts(cn: str) -> tuple[int, int]:
            """Return (red_count, yellow_count) for projects managed by cn."""
            _p = [p.strip() for p in cn.split(",")]
            _keys = [cn.lower(), _p[0].lower()]
            if len(_p) == 2:
                _keys.append(f"{_p[1]} {_p[0]}".lower())
            for _k in _keys:
                if _k in _rag_by_pm:
                    return _rag_by_pm[_k]["red"], _rag_by_pm[_k]["yellow"]
            return 0, 0

        def _build_row(cn, is_leaver=False, exit_dt=None):
            _parts = [p.strip() for p in cn.split(",")]
            _variants = {cn.lower(), _parts[0].lower()}
            if len(_parts) == 2:
                _variants.add(f"{_parts[1]} {_parts[0]}".lower())

            # Month NS rows for this consultant (for overrun calculation)
            _emp_mask_mo = _month_ns_all["employee"].astype(str).str.strip().str.lower().apply(
                lambda v: v in _variants or any(
                    v == nv or v.startswith(nv+" ") or v.endswith(" "+nv) for nv in _variants)
            )
            _emp_rows_mo = _month_ns_all[_emp_mask_mo]

            # Week NS rows for this consultant (for hours-this-week columns)
            _emp_mask_wk = _wk_ns_all["employee"].astype(str).str.strip().str.lower().apply(
                lambda v: v in _variants or any(
                    v == nv or v.startswith(nv+" ") or v.endswith(" "+nv) for nv in _variants)
            ) if not _wk_ns_all.empty else pd.Series(dtype=bool)
            _emp_rows_wk = _wk_ns_all[_emp_mask_wk] if not _wk_ns_all.empty else pd.DataFrame()

            # FF credit/overrun uses MONTH data (matches top-level engine)
            if _emp_rows_mo.empty:
                _ff_util = 0.0; _ff_over = 0.0
            else:
                _bt_cn = _emp_rows_mo.get("billing_type", pd.Series(dtype="object")).fillna("").str.strip().str.lower()
                _ff_rows_cn = _emp_rows_mo[(_bt_cn != "internal") & (_bt_cn != "t&m")].sort_values("date")
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
                            _ff_util += _fh; continue
                        _pid_fr = str(_fr.get("project_id","")).strip()
                        _fpkey  = _pid_fr if _pid_fr and _pid_fr != "nan" else _fp
                        if _fpkey not in _con_cn:
                            _con_cn[_fpkey] = prior_htd.get(_fpkey, prior_htd.get(_fp, 0.0))
                        _fused = _con_cn[_fpkey]; _frem = _fsc - _fused
                        if _frem <= 0:
                            _ff_over += _fh
                        elif _fh <= _frem:
                            _ff_util += _fh; _con_cn[_fpkey] = _fused + _fh
                        else:
                            _ff_util += _frem; _ff_over += _fh - _frem; _con_cn[_fpkey] = _fsc
            _ff_util = round(_ff_util, 2)
            _ff_over = round(_ff_over, 2)

            # Internal % uses MONTH data (same as before)
            _admin = 0.0
            if not _emp_rows_mo.empty and "billing_type" in _emp_rows_mo.columns:
                _bt_mo = _emp_rows_mo["billing_type"].fillna("").str.strip().str.lower()
                _admin = round(_emp_rows_mo[_bt_mo == "internal"]["hours"].sum(), 2)

            # Weekly util hours (FF + T&M this week only — for the Util % column)
            _wk_ff = 0.0; _wk_tm = 0.0
            if not _emp_rows_wk.empty and "billing_type" in _emp_rows_wk.columns:
                _bt_wk = _emp_rows_wk["billing_type"].fillna("").str.strip().str.lower()
                _wk_ff = round(_emp_rows_wk[(_bt_wk != "internal") & (_bt_wk != "t&m")]["hours"].sum(), 2)
                _wk_tm = round(_emp_rows_wk[_bt_wk == "t&m"]["hours"].sum(), 2)

            # Available hours — monthly total prorated to weekly
            _cl = EMPLOYEE_LOCATION.get(cn, "")
            _cr = PS_REGION_OVERRIDE.get(cn, PS_REGION_MAP.get(_cl, ""))
            _avail_full = AVAIL_HOURS.get(_cl, AVAIL_HOURS.get(_cr, {})).get(month_key)
            if isinstance(_avail_full, tuple): _avail_full = _avail_full[0]
            _avail_full = float(_avail_full) if _avail_full else None

            if is_leaver and exit_dt is not None and _avail_full:
                _days_worked = min(exit_dt.day, _days_in_month)
                _avail_mo = round(_avail_full * _days_worked / _days_in_month, 2)
            else:
                _avail_mo = _avail_full

            # Weekly avail: monthly ÷ business days in month × 5
            _avail_wk = round(_avail_mo / _bdays_month * 5, 1) if _avail_mo and _bdays_month else None

            # Util %: weekly billable ÷ weekly avail (apples-to-apples)
            _wk_bill    = round(_wk_ff + _wk_tm, 2)
            _util_pct   = round(_wk_bill / _avail_wk * 100, 1) if _avail_wk and _wk_bill > 0 else None
            # FF Overrun %: MTD overrun ÷ monthly avail (shows severity relative to capacity)
            _over_pct   = round(_ff_over / _avail_mo * 100, 1) if _avail_mo and _ff_over > 0 else None
            # Internal %: MTD internal ÷ monthly avail
            _int_pct    = round(_admin   / _avail_mo * 100, 1) if _avail_mo and _admin > 0 else None

            _display = f"{_parts[1].strip()} {_parts[0]}" if len(_parts) == 2 else cn
            if is_leaver and exit_dt:
                _display += " *"

            _whs_s, _whs_l, _ = consultant_whs(cn, df_drs) if df_drs is not None else (None, "—", None)
            _rag_r, _rag_y    = _get_rag_counts(cn)

            # Status — utilization signal only (RAG is its own column)
            # Priority: Overrun > High WHS > OK > No data
            if _ff_over > 0:
                _status = "Overrun"
            elif _whs_s is not None and _whs_s >= 75:
                _status = "High WHS"
            elif _wk_bill > 0 or not _emp_rows_mo.empty:
                _status = "OK"
            else:
                _status = "No data"

            return {
                "_cn":          cn,           # canonical name for sorting
                "Consultant":   _display,
                "_whs_score":   _whs_s,
                "_whs_label":   _whs_l or "—",
                "_avail_wk":    _avail_wk,
                "_util_pct":    _util_pct,
                "_ff_over":     _ff_over,
                "_over_pct":    _over_pct,
                "_int_pct":     _int_pct,
                "_rag_r":       _rag_r,
                "_rag_y":       _rag_y,
                "_status":      _status,
            }

        _rows = [_build_row(cn) for cn in _scope_names]
        _rows += [_build_row(ln, is_leaver=True, exit_dt=ex) for ln, ex in _leaver_scope]

        if _rows:
            import numpy as _np

            _raw = pd.DataFrame(_rows)

            # ── Sort: Overrun first → High WHS → OK → No data ─────────────────
            _STATUS_ORDER = {"Overrun": 0, "High WHS": 1, "OK": 2, "No data": 3}
            _raw["_sort_status"] = _raw["_status"].map(_STATUS_ORDER).fillna(9)
            # Within same status, sort by overrun hours desc, then WHS desc
            _raw = _raw.sort_values(
                ["_sort_status", "_ff_over", "_whs_score"],
                ascending=[True, False, False],
            ).reset_index(drop=True)

            # ── Build display columns ─────────────────────────────────────────
            def _whs_str(score, label):
                if score is None or (isinstance(score, float) and _np.isnan(float(score or _np.nan))):
                    return "No data"
                return f"{score:.1f} · {label}"

            def _fmt(v, decimals=1, suffix=""):
                if v is None: return ""
                try:
                    f = float(v)
                    if _np.isnan(f) or f == 0: return ""
                    return f"{f:.{decimals}f}{suffix}"
                except (TypeError, ValueError):
                    return ""

            def _rag_str(r, y):
                parts = []
                if r > 0: parts.append(f"{r}R")
                if y > 0: parts.append(f"{y}Y")
                return " · ".join(parts) if parts else ""

            # Totals
            def _safe_sum(col):
                s = _raw[col].dropna()
                return round(float(s.sum()), 2) if len(s) else None

            _t_avail_wk  = _safe_sum("_avail_wk")
            _t_ff_over   = _safe_sum("_ff_over")
            _t_rag_r     = int(_raw["_rag_r"].sum())
            _t_rag_y     = int(_raw["_rag_y"].sum())
            _whs_scores  = _raw["_whs_score"].dropna()
            _avg_whs     = round(float(_whs_scores.mean()), 1) if len(_whs_scores) else None
            _n_consult   = len(_raw)

            # Weighted team util %: sum of (util_h) / sum of (avail_wk)
            # Recalculate from raw data is not available here — use average of non-null util_pct
            _t_util_p = round(float(_raw["_util_pct"].dropna().mean()), 1) if _raw["_util_pct"].notna().any() else None
            _t_over_p = round(_t_ff_over / (_safe_sum("_avail_wk") or 1) * 100, 1) if _t_ff_over else None

            # Build final display DataFrame
            _disp_rows = []
            for _, _r in _raw.iterrows():
                _disp_rows.append({
                    "Consultant":   _r["Consultant"],
                    "Workload":     _whs_str(_r["_whs_score"], _r["_whs_label"]),
                    "Avail (wk)":   _fmt(_r["_avail_wk"], 1),
                    "Util %":       float(_r["_util_pct"]) if _r["_util_pct"] is not None else 0.0,
                    "FF Overrun":   _fmt(_r["_ff_over"], 2, "h") + (f"  {_fmt(_r['_over_pct'], 1, '%')}" if _r["_over_pct"] else ""),
                    "Internal %":   _fmt(_r["_int_pct"], 1, "%"),
                    "RAG":          _rag_str(int(_r["_rag_r"]), int(_r["_rag_y"])),
                    "Status":       _r["_status"],
                    # sentinels for styling
                    "_ff_over_raw": _r["_ff_over"],
                    "_status_raw":  _r["_status"],
                    "_whs_raw":     _r["_whs_score"],
                    "_rag_r_raw":   _r["_rag_r"],
                })

            _tbl = pd.DataFrame(_disp_rows)

            _totals_disp = pd.DataFrame([{
                "Consultant":   f"Team total · {_n_consult}",
                "Workload":     f"AVG {_avg_whs}" if _avg_whs is not None else "—",
                "Avail (wk)":   _fmt(_t_avail_wk, 1),
                "Util %":       float(_t_util_p) if _t_util_p is not None else 0.0,
                "FF Overrun":   _fmt(_t_ff_over, 2, "h"),
                "Internal %":   "",
                "RAG":          _rag_str(_t_rag_r, _t_rag_y),
                "Status":       "",
            }])

            # ── Styler: overrun row highlight + WHS high row tint ─────────────
            _OVER_ROW  = "background-color: rgba(239,68,68,0.06);"
            _OVER_CELL = "background-color: rgba(239,68,68,0.13); color: #dc2626; font-weight: 500;"
            _WHS_ROW   = "background-color: rgba(239,159,39,0.06);"
            _RAG_CELL  = "color: #a32d2d; font-weight: 500;"

            def _row_style(row):
                n = len(row)
                base = [""] * n
                s = row.get("_status_raw", "")
                o = float(row.get("_ff_over_raw") or 0)
                r = float(row.get("_rag_r_raw") or 0)
                if s == "Overrun":
                    base = [_OVER_ROW] * n
                elif s == "High WHS":
                    base = [_WHS_ROW] * n
                idx = list(row.index)
                # Overrun cell
                if "FF Overrun" in idx and o > 0:
                    base[idx.index("FF Overrun")] = _OVER_CELL
                # RAG cell red tint when red projects present
                if "RAG" in idx and r > 0:
                    base[idx.index("RAG")] = _RAG_CELL
                return base

            _display_cols = ["Consultant","Workload","Avail (wk)","Util %","FF Overrun","Internal %","RAG","Status"]
            _styler = _tbl.style.apply(_row_style, axis=1)

            # ── Render ────────────────────────────────────────────────────────
            st.dataframe(
                _styler,
                use_container_width=True,
                hide_index=True,
                column_order=_display_cols,
                column_config={
                    "Consultant":   st.column_config.TextColumn("Consultant",   width="medium"),
                    "Workload":     st.column_config.TextColumn("Workload",     width="small"),
                    "Avail (wk)":   st.column_config.TextColumn("Avail (wk)",  width="small"),
                    "Util %":       st.column_config.ProgressColumn(
                                        "Util %",
                                        help="Weekly billable hours (FF + T&M this week) ÷ weekly available hours",
                                        format="%.1f%%",
                                        min_value=0,
                                        max_value=100,
                                        width="medium",
                                    ),
                    "FF Overrun":   st.column_config.TextColumn(
                                        "FF Overrun",
                                        help="MTD Fixed Fee hours beyond scope · hours and % of monthly available",
                                        width="small",
                                    ),
                    "Internal %":   st.column_config.TextColumn(
                                        "Internal %",
                                        help="MTD internal hours as % of monthly available",
                                        width="small",
                                    ),
                    "RAG":          st.column_config.TextColumn(
                                        "RAG",
                                        help="Red and Yellow RAG projects managed by this consultant (from DRS)",
                                        width="small",
                                    ),
                    "Status":       st.column_config.TextColumn(
                                        "Status",
                                        help="Overrun: FF hours exceeded scope  ·  High WHS: workload score ≥ 75  ·  OK: active, no flags",
                                        width="small",
                                    ),
                    # hide sentinels
                    "_ff_over_raw": None,
                    "_status_raw":  None,
                    "_whs_raw":     None,
                    "_rag_r_raw":   None,
                },
            )

            # Totals row
            st.dataframe(
                _totals_disp.style.set_properties(**{
                    "font-weight": "600",
                    "border-top":  "2px solid rgba(128,128,128,0.3)",
                }),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Consultant": st.column_config.TextColumn("Consultant", width="medium"),
                    "Workload":   st.column_config.TextColumn("Workload",   width="small"),
                    "Avail (wk)": st.column_config.TextColumn("Avail (wk)", width="small"),
                    "Util %":     st.column_config.ProgressColumn(
                                      "Util %", format="%.1f%%",
                                      min_value=0, max_value=100, width="medium",
                                  ),
                    "FF Overrun": st.column_config.TextColumn("FF Overrun", width="small"),
                    "Internal %": st.column_config.TextColumn("Internal %", width="small"),
                    "RAG":        st.column_config.TextColumn("RAG",        width="small"),
                    "Status":     st.column_config.TextColumn("Status",     width="small"),
                },
            )

            if _leaver_scope:
                st.caption("* Available hours prorated to exit date")
            st.caption(
                "Avail (wk): monthly available ÷ business days × 5  ·  "
                "Util %: billable hours this week ÷ weekly avail  ·  "
                "FF Overrun & Internal %: month-to-date  ·  "
                "Status: Overrun > High WHS (≥75) > OK"
            )

    except Exception as _db_err:
        st.warning(
            "\u26a0\ufe0f Team breakdown could not load. This usually means the "
            "uploaded time data is from an older export format. "
            "Please upload a current NS Time Detail export and try again.",
            icon="\u26a0\ufe0f"
        )
        with st.expander("Technical details (for support)", expanded=False):
            st.code(str(_db_err))

# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Project Snapshot
# ══════════════════════════════════════════════════════════════════════════════

if df_drs is None:
    st.info("Upload SS DRS Export in the sidebar to see your project snapshot.")
elif my_projects.empty:
    st.info("No projects found in the DRS file for your profile.")
else:
    _today = _snap_today; _n7 = _snap_n7; _14 = _snap_14

    def _rag_label(r):
        _pn   = str(r.get("project_name","") or "")
        _cust = _pn.split(" - ")[0].strip()[:22] if " - " in _pn else _pn[:22]
        _pt   = str(r.get("project_type","") or "")
        # Extract product family: "ZonePay: Implementation" → "ZonePayroll"
        # "ZEP: Implementation" → "ZEP"
        # Map common short codes to display names
        _PROD_DISPLAY = {
            "zonepay":    "ZonePayroll",
            "zonebill":   "ZoneBilling",
            "zoneapp":    "ZoneApps",
            "zonerpt":    "ZoneReporting",
            "zep":        "ZEP",
            "zab":        "ZoneBilling",
        }
        if _pt:
            _prefix = _pt.split(":")[0].strip().lower()
            _prod = _PROD_DISPLAY.get(_prefix, _pt.split(":")[0].strip())
        else:
            _prod = ""
        return f"{_cust} : {_prod}" if _prod else _cust

    # Phase breakdown — exclude Unassigned, sort by phase order
    _PHASE_ABBREV = {
        "00. onboarding":                   "Onboarding",
        "01. requirements and design":       "Requirements and Design",
        "02. configuration":                 "Configuration",
        "03. enablement/training":           "Enablement/Training",
        "04. uat":                           "UAT",
        "05. prep for go-live":              "Prep for Go-Live",
        "06. go-live":                       "Go-Live",
        "07. data migration":                "Data Migration",
        "08. ready for support transition":  "Ready for Support Transition",
        "09. phase 2 scoping":               "Phase 2 Scoping",
    }
    _pc_all = {}
    if "phase" in _active.columns:
        for ph, cnt in _active["phase"].fillna("").value_counts().items():
            _ph_clean = str(ph).strip()
            if _ph_clean and _ph_clean not in ("", "nan", "None"):
                _pc_all[_ph_clean] = cnt
    # Sort by phase index, keep Unassigned at end
    _pc_sorted = sorted(
        [(ph, cnt) for ph, cnt in _pc_all.items()],
        key=lambda x: (_pidx_db(x[0]) if _pidx_db(x[0]) >= 0 else 999)
    )
    # Separate unassigned
    _pc_assigned   = [(ph, cnt) for ph, cnt in _pc_sorted if _pidx_db(ph) >= 0]
    _unassigned_ct = sum(cnt for ph, cnt in _pc_sorted if _pidx_db(ph) < 0)

    # ── New donut panel: RAG, Phase breakdown, My projects snapshot, Project age ─

    def _donut_svg(segments, cx=32, cy=32, r=24, sw=7):
        """Build stacked donut SVG circles from list of (pct, color) tuples."""
        total = sum(p for p, _ in segments)
        if total == 0:
            return f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="rgba(128,128,128,0.2)" stroke-width="{sw}"/>'
        circumference = 2 * 3.14159 * r
        out = f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="rgba(128,128,128,0.2)" stroke-width="{sw}"/>'
        offset = 38  # start at top
        for pct, color in segments:
            dash = round((pct / total) * circumference, 1)
            gap  = round(circumference - dash, 1)
            out += (f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{color}" '
                    f'stroke-width="{sw}" stroke-dasharray="{dash} {gap}" '
                    f'stroke-dashoffset="{offset}" stroke-linecap="round"/>')
            offset -= dash
        return out

    def _donut_card(title, center_val, center_sub, segments, legend_rows, note=None):
        """Render a donut card via st.markdown."""
        svg_inner = _donut_svg(segments)
        legend_html = ""
        for dot_color, label, val in legend_rows:
            legend_html += (
                f"<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;'>"
                f"<div style='display:flex;align-items:center;gap:6px;font-size:13px;color:var(--color-text-secondary);'>"
                f"<div style='width:8px;height:8px;border-radius:50%;background:{dot_color};flex-shrink:0;'></div>{label}</div>"
                f"<span style='font-size:14px;font-weight:600;color:var(--color-text-primary);'>{val}</span></div>"
            )
        note_html = (
            f"<div style='margin-top:8px;padding-top:8px;border-top:1px solid rgba(128,128,128,0.18);"
            f"font-size:11px;color:var(--color-text-secondary);'>{note}</div>"
        ) if note else ""
        st.markdown(
            f"<div style='background:var(--color-background-primary);"
            f"border:1px solid rgba(128,128,128,0.25);"
            f"border-radius:10px;padding:14px 16px;margin-bottom:10px;'>"
            f"<div style='font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:0.6px;"
            f"color:var(--color-text-secondary);margin-bottom:12px;'>{title}</div>"
            f"<div style='display:flex;align-items:center;gap:16px;'>"
            f"<div style='position:relative;width:92px;height:92px;flex-shrink:0;'>"
            f"<svg viewBox='0 0 64 64' width='92' height='92'>{svg_inner}</svg>"
            f"<div style='position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;'>"
            f"<span style='font-size:16px;font-weight:600;color:var(--color-text-primary);line-height:1;'>{center_val}</span>"
            f"<span style='font-size:11px;color:var(--color-text-secondary);margin-top:2px;'>{center_sub}</span>"
            f"</div></div>"
            f"<div style='flex:1;'>{legend_html}</div></div>"
            f"{note_html}</div>",
            unsafe_allow_html=True
        )

    # ── Column layout: briefing text left, donuts right ───────────────────────
    _col_brief, _col_donuts = st.columns([1.4, 1.0], gap="large")

    with _col_donuts:
        # ── My utilization donut ──────────────────────────────────────────────
        if util_pct is not None:
            _UTIL_TARGET_DISP = 70.0
            _util_segs = [(util_pct, "#08A9B7"), (max(0, 100 - util_pct), "var(--color-border-tertiary)")]
            _pu_badge  = ""
            _month_start2 = today.replace(day=1)
            _month_end2   = (today.replace(day=28)+pd.Timedelta(days=4)).replace(day=1)-pd.Timedelta(days=1)
            _days_total2  = len(pd.bdate_range(_month_start2, _month_end2))
            _days_elapsed2= len(pd.bdate_range(_month_start2, today))
            _gap2 = (_wk_util_pct or 0) - _UTIL_TARGET_DISP
            if _gap2 >= 0:
                _pace_c = "#97C459"; _pace_str = f"+{round(_gap2)}% ahead"
            else:
                _pace_c = "#F09595"; _pace_str = f"{round(_gap2)}% behind"
            st.markdown(
                f"<div style='background:var(--color-background-primary);border:1px solid rgba(128,128,128,0.25);border-radius:10px;padding:14px 16px;margin-bottom:10px;'>"
                f"<div style='font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:0.6px;color:var(--color-text-secondary);margin-bottom:10px;'>My utilization · week of {_week_start.strftime('%b %-d')}</div>"
                f"<div style='display:flex;align-items:center;gap:14px;'>"
                f"<div style='position:relative;width:92px;height:92px;flex-shrink:0;'>"
                f"<svg viewBox='0 0 64 64' width='92' height='92'>"
                f"<circle cx='32' cy='32' r='24' fill='none' stroke='rgba(128,128,128,0.2)' stroke-width='7'/>"
                f"<circle cx='32' cy='32' r='24' fill='none' stroke='#27AE60' stroke-width='7' "
                f"stroke-dasharray='{round((_wk_util_pct or 0)/100*151,1)} {round((1-(_wk_util_pct or 0)/100)*151,1)}' stroke-dashoffset='38' stroke-linecap='round'/>"
                f"</svg>"
                f"<div style='position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;'>"
                f"<span style='font-size:13px;font-weight:500;color:var(--color-text-primary);line-height:1;'>{_wk_util_pct or '—'}%</span>"
                f"<span style='font-size:9px;color:var(--color-text-secondary);margin-top:1px;'>util</span></div></div>"
                f"<div style='flex:1;display:flex;flex-direction:column;gap:5px;'>"
                f"<div style='display:flex;justify-content:space-between;'><span style='font-size:12px;color:var(--color-text-secondary);'>Billable hrs</span><span style='font-size:12px;font-weight:500;'>{_wk_total or '—'}h / {_wk_billable_h or '—'}h</span></div>"
                f"<div style='display:flex;justify-content:space-between;'><span style='font-size:12px;color:var(--color-text-secondary);'>Target</span><span style='font-size:12px;font-weight:500;'>{int(_UTIL_TARGET_DISP)}%</span></div>"
                f"<div style='display:flex;justify-content:space-between;align-items:center;'><span style='font-size:12px;color:var(--color-text-secondary);'>Pacing</span>"
                f"<span style='font-size:11px;padding:1px 6px;border-radius:3px;background:rgba(128,128,128,0.1);color:{_pace_c};font-weight:500;'>{_pace_str}</span></div>"
                f"<div style='display:flex;justify-content:space-between;'><span style='font-size:12px;color:var(--color-text-secondary);'>Overrun hrs</span>"
                f"<span style='font-size:12px;font-weight:500;color:{'#E24B4A' if overrun_hrs > 0 else 'var(--color-text-primary)'};'>{_fmt_hrs(overrun_hrs)} MTD</span></div>"
                f"</div></div></div>",
                unsafe_allow_html=True
            )

        # ── WHS above donuts ──────────────────────────────────────────────────
        if _whs_score is not None:
            st.markdown(
                f"<div style='background:var(--color-background-primary);border:1px solid rgba(128,128,128,0.25);"
                f"border-radius:10px;padding:12px 14px;margin-bottom:10px;"
                f"display:flex;justify-content:space-between;align-items:center;'>"
                f"<div><div style='font-size:10px;font-weight:500;text-transform:uppercase;letter-spacing:0.6px;"
                f"color:var(--color-text-secondary);margin-bottom:4px;'>Workload health score</div>"
                f"<div style='font-size:11px;color:var(--color-text-secondary);'>Composite · projects, phases, overrun, stale</div></div>"
                f"<div style='text-align:right;'>"
                f"<div style='font-size:28px;font-weight:500;color:{_whs_col};'>{_whs_score}</div>"
                f"<div style='font-size:11px;color:{_whs_col};'>{_whs_label}</div></div></div>",
                unsafe_allow_html=True
            )

        # 1 — RAG & risk
        _n_rag_total  = int(my_projects["project_id"].nunique()) if not my_projects.empty and "project_id" in my_projects.columns else 0
        _n_green      = max(0, _n_rag_total - len(_rag_red) - len(_rag_yellow))
        _rv_all       = my_projects["rag"].fillna("").astype(str).str.strip().str.lower() if "rag" in my_projects.columns and not my_projects.empty else pd.Series(dtype=str)
        _n_unrated    = int((_rv_all == "").sum()) if len(_rv_all) > 0 else 0
        _n_green_conf = max(0, _n_green - _n_unrated)
        _n_rag_rated  = len(_rag_yellow) + len(_rag_red)  # Red + Yellow only = "at risk"
        _rag_note     = f"* {_n_unrated} unrated project{'s' if _n_unrated != 1 else ''} — RAG not yet set" if _n_unrated > 0 else None
        _donut_card(
            "RAG &amp; risk",
            str(_n_rag_rated), "at risk",
            [(_n_green_conf, "#639922"), (len(_rag_yellow), "#EF9F27"), (len(_rag_red), "#E24B4A"), (_n_unrated, "#888780")],
            [
                ("#639922", "Green",   str(_n_green_conf)),
                ("#EF9F27", "Yellow",  str(len(_rag_yellow))),
                ("#E24B4A", "Red",     str(len(_rag_red))),
                ("#888780", "Unrated", str(_n_unrated)),
            ],
            note=_rag_note
        )

        # 2 — Phase breakdown
        _PHASE_GROUPS = [
            ("Onboarding / Req",   ["00.", "01."], "#534AB7"),
            ("Config / Training",  ["02.", "03."], "#378ADD"),
            ("UAT",                ["04."],        "#08A9B7"),
            ("Prep / Go-Live",     ["05.", "06."], "#EF9F27"),
            ("Data Migration",     ["07."],        "#E24B4A"),
            ("Ready for Support",  ["08.", "10."], "#639922"),
            ("Phase 2 Scoping",    ["09."],        "#888780"),
        ]
        _phase_counts      = {}
        _phase_blank       = 0
        _phase_unmatched   = 0
        _phase_pending_bil = 0
        if "phase" in _active.columns and not _active.empty:
            for ph in _active["phase"].fillna(""):
                pl = str(ph).strip().lower()
                if not pl:
                    _phase_blank += 1
                    continue
                matched = False
                for grp_name, prefixes, _ in _PHASE_GROUPS:
                    if any(pl.startswith(p) for p in prefixes):
                        # Special callout for 10. still in progress
                        if pl.startswith("10."):
                            _phase_pending_bil += 1
                        _phase_counts[grp_name] = _phase_counts.get(grp_name, 0) + 1
                        matched = True
                        break
                if not matched:
                    _phase_unmatched += 1
        _ph_note_parts = []
        if _phase_blank > 0:
            _ph_note_parts.append(f"{_phase_blank} project{'s' if _phase_blank != 1 else ''} with no phase set")
        if _phase_pending_bil > 0:
            _ph_note_parts.append(f"{_phase_pending_bil} marked Complete/Pending Billing — check status")
        if _phase_unmatched > 0:
            _ph_note_parts.append(f"{_phase_unmatched} in unrecognised phase")
        _ph_note = " · ".join(_ph_note_parts) if _ph_note_parts else None
        _ph_segs   = [(_phase_counts.get(g, 0), c) for g, _, c in _PHASE_GROUPS if _phase_counts.get(g, 0) > 0]
        _ph_legend = [(c, g, str(_phase_counts.get(g, 0))) for g, _, c in _PHASE_GROUPS if _phase_counts.get(g, 0) > 0]
        _donut_card(
            "Phase breakdown",
            str(_n_active_dc), "open",
            _ph_segs, _ph_legend,
            note=_ph_note
        )


        # 3 — My projects snapshot
        _n_assigned  = int(my_projects["project_id"].nunique()) if not my_projects.empty and "project_id" in my_projects.columns else 0
        _n_open      = _n_active_dc  # active = not on hold
        _n_onhold    = _n_onhold_dc
        _n_pend_cl   = int(my_projects[my_projects.get("_pending_close", pd.Series(False, index=my_projects.index)).astype(bool)]["project_id"].nunique()) if not my_projects.empty and "project_id" in my_projects.columns else 0
        # Active = projects with NS time this month
        _this_month_key = today.strftime("%Y-%m")
        _active_proj_ids: set = set()
        if df_ns is not None and not my_ns.empty and "project_id" in my_ns.columns and "date" in my_ns.columns:
            _mn2 = my_ns.copy()
            _mn2["_m"] = pd.to_datetime(_mn2["date"], errors="coerce").dt.strftime("%Y-%m")
            _active_proj_ids = set(_mn2[_mn2["_m"] == _this_month_key]["project_id"].dropna().unique())
        _n_time_active = len(_active_proj_ids)
        _snap_note = "* Active = open projects with recent time booked"
        _donut_card(
            "My projects snapshot",
            str(_n_assigned), "assigned",
            [(_n_open, "#08A9B7"), (_n_time_active, "#378ADD"), (_n_onhold, "#888780"), (_n_pend_cl, "#534AB7")],
            [
                ("#08A9B7", "Open",          str(_n_open)),
                ("#378ADD", "Active *",      str(_n_time_active)),
                ("#888780", "On hold",       str(_n_onhold)),
                ("#534AB7", "Pending close", str(_n_pend_cl)),
            ],
            note=_snap_note
        )

        # 4 — Project age
        _age_bands = [("<3 months", 0, 3, "#08A9B7"), ("3–6 months", 3, 6, "#378ADD"),
                      ("6–9 months", 6, 9, "#EF9F27"), ("9–12 months", 9, 12, "#E24B4A")]
        _age_counts = {b[0]: 0 for b in _age_bands}
        _age_12plus = 0
        _avg_months = None
        if "start_date" in _active.columns and not _active.empty:
            _sd2 = pd.to_datetime(_active["start_date"], errors="coerce")
            _mo  = (_snap_today - _sd2).dt.days / 30.44
            _avg_months = round(_mo.dropna().mean(), 1) if not _mo.dropna().empty else None
            for _, m in _mo.dropna().items():
                for band_name, lo, hi, _ in _age_bands:
                    if lo <= m < hi:
                        _age_counts[band_name] += 1
                        break
                else:
                    if m >= 12:
                        _age_12plus += 1
        _age_segs   = [(_age_counts[b[0]], b[3]) for b in _age_bands]
        _age_legend = [(b[3], b[0], str(_age_counts[b[0]])) for b in _age_bands]
        _age_note   = f"{_age_12plus} project{'s' if _age_12plus != 1 else ''} at 12+ months — escalation review recommended" if _age_12plus > 0 else None
        _donut_card(
            "Project age",
            str(_avg_months) if _avg_months else "—", "avg mo",
            _age_segs, _age_legend,
            note=_age_note
        )

        # WHS card moved above donuts

    with _col_brief:
        # ── This Week's Focus (briefing text) ────────────────────────────────
        if '_bhtml' in dir() or '_bhtml' in locals():
            pass  # rendered below
        try:
            st.markdown(_bhtml, unsafe_allow_html=True)
        except Exception:
            pass

        # ── Priority task list ────────────────────────────────────────────────
        if not _is_group_view and df_drs is not None and not my_projects.empty:
            _due_items = []

            # Intro emails pending — uncapped, concrete action with SS writeback
            for _, _mr in _mi.iterrows():
                _cust  = str(_mr.get("project_name","")).split(" - ")[0][:35]
                _pid   = str(_mr.get("project_id",""))
                _ptype = str(_mr.get("project_type","")).strip()
                _sub   = f"{_ptype} · {_pid}" if _ptype else _pid
                _due_items.append({"dot": "#EF9F27", "name": f"Intro email pending — {_cust}", "sub": _sub, "row_id": _mr.get("_ss_row_id"), "type": "intro"})

            # Go-lives this week — confirm readiness, uncapped
            for _, _gr2 in _gls.iterrows():
                _cust  = str(_gr2.get("project_name","")).split(" - ")[0][:35]
                _pid   = str(_gr2.get("project_id",""))
                _ptype = str(_gr2.get("project_type","")).strip()
                _gl_d  = pd.Timestamp(_gr2.get(_gld_col)).strftime("%-d %b") if _gld_col in _gr2.index else ""
                _sub   = f"{_ptype} · {_pid} · Scheduled {_gl_d}" if _ptype else f"{_pid} · Scheduled {_gl_d}"
                _due_items.append({"dot": "#639922", "name": f"Confirm go-live readiness — {_cust}", "sub": _sub, "row_id": None, "type": "golive"})

            # Stale projects — send outreach via Customer Engagement, uncapped
            for _, _sr in _stale.iterrows():
                _cust  = str(_sr.get("project_name","")).split(" - ")[0][:35]
                _pid   = str(_sr.get("project_id",""))
                _ptype = str(_sr.get("project_type","")).strip()
                _days  = int(_sr.get("days_inactive", 0))
                _sub   = f"{_ptype} · {_pid} · {_days} days inactive" if _ptype else f"{_pid} · {_days} days inactive"
                _due_items.append({"dot": "#888780", "name": f"Re-engage — {_cust}", "sub": _sub, "row_id": None, "type": "stale"})

            if _due_items:
                with st.container(border=True):
                    _hc1, _hc2 = st.columns([3, 1])
                    _hc1.markdown("<span style='font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.6px;color:var(--color-text-secondary);'>This week's priorities</span>", unsafe_allow_html=True)
                    _hc2.markdown(f"<span style='font-size:12px;color:var(--color-text-secondary);float:right;'>{len(_due_items)} items</span>", unsafe_allow_html=True)

                    for idx, item in enumerate(_due_items):
                        # Key on type + project ID (first token of sub) + DRS load ts
                        _pid_key  = str(item.get("sub","")).split("·")[- 1 if item["type"] == "stale" else 0].strip()[:20].replace(" ","")
                        _drs_ts   = str(st.session_state.get("drs_load_ts", ""))[:10]
                        _done_key = f"done_{item['type']}_{_pid_key}_{_drs_ts}"
                        _is_done  = st.session_state.get(_done_key, False)
                        _ic1, _ic2 = st.columns([6, 1])
                        with _ic1:
                            _strike = "text-decoration:line-through;opacity:0.4;" if _is_done else ""
                            st.markdown(
                                f"<div style='display:flex;align-items:flex-start;gap:8px;padding:4px 0;{_strike}'>"
                                f"<div style='width:7px;height:7px;border-radius:50%;background:{item['dot']};flex-shrink:0;margin-top:5px;'></div>"
                                f"<div><div style='font-size:13px;font-weight:500;'>{item['name']}</div>"
                                f"<div style='font-size:11px;color:var(--color-text-secondary);'>{item['sub']}</div></div>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
                        with _ic2:
                            if not _is_done:
                                if st.button("✓ Done", key=f"btn_{idx}", use_container_width=True):
                                    st.session_state[_done_key] = True
                                    if item["type"] == "intro":
                                        try:
                                            from shared.smartsheet_api import write_row_updates
                                            _row_id = item.get("row_id")
                                            if _row_id:
                                                _ok, _errs = write_row_updates([{
                                                    "_ss_row_id": _row_id,
                                                    "project_name": item["name"],
                                                    "changes": {"ms_intro_email": pd.Timestamp.today().normalize()}
                                                }])
                                                if _ok:
                                                    st.toast("✓ Intro email date updated in Smartsheet", icon="✅")
                                                else:
                                                    _err_msg = _errs[0] if _errs else "unknown error"
                                                    st.toast(f"⚠ SS write failed: {_err_msg}", icon="⚠️")
                                            else:
                                                # No row ID — DRS loaded from CSV upload, not API
                                                st.toast("Marked locally — sync DRS via API to write back", icon="ℹ️")
                                        except Exception as _e:
                                            st.toast(f"Saved locally — sync pending ({type(_e).__name__})", icon="ℹ️")
                                    st.rerun()
                            else:
                                st.markdown("<span style='font-size:11px;color:var(--color-text-secondary);'>Done ✓</span>", unsafe_allow_html=True)



        # Next 14 days milestone list
        _ms14_items = []
        if not _active.empty:
            _ms_cols = {
                "ms_intro_email":    ("Intro email", "#378ADD"),
                "ms_config_start":   ("Config start", "#534AB7"),
                "ms_uat_signoff":    ("UAT sign-off", "#EF9F27"),
                "ms_hypercare_start":("Hypercare start", "#534AB7"),
            }
            _gl14_items = []
            if _gl14_col in _active.columns:
                for _, _gr3 in _gl14.iterrows():
                    _cust = str(_gr3.get("project_name","")).split(" - ")[0][:28]
                    _pid2 = str(_gr3.get("project_id",""))
                    _dt   = pd.Timestamp(_gr3.get(_gl14_col))
                    _gl14_items.append({"dt": _dt, "name": "Go-live", "cust": _cust, "pid": _pid2, "dot": "#639922", "badge": "Go-live", "badge_bg": "#EAF3DE", "badge_col": "#3B6D11"})

            for ms_col, (ms_label, ms_color) in _ms_cols.items():
                if ms_col in _active.columns:
                    for _, _row in _active.iterrows():
                        _dt2 = pd.to_datetime(_row.get(ms_col), errors="coerce")
                        if pd.isna(_dt2): continue
                        if _snap_today <= _dt2 <= _snap_today + pd.Timedelta(days=14):
                            _cust = str(_row.get("project_name","")).split(" - ")[0][:28]
                            _pid2 = str(_row.get("project_id",""))
                            _ms14_items.append({"dt": _dt2, "name": ms_label, "cust": _cust, "pid": _pid2, "dot": ms_color, "badge": None})

            _ms14_items = sorted(_gl14_items + _ms14_items, key=lambda x: x["dt"])

        if _ms14_items:
            st.markdown(
                f"<div style='background:var(--color-background-primary);border:1px solid rgba(128,128,128,0.25);"
                f"border-radius:10px;padding:14px 16px;margin-bottom:12px;'>"
                f"<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;'>"
                f"<span style='font-size:10px;font-weight:500;text-transform:uppercase;letter-spacing:0.6px;color:var(--color-text-secondary);'>Next 14 days</span>"
                f"<span style='font-size:12px;color:var(--color-text-secondary);'>{len(_ms14_items)} scheduled</span></div>",
                unsafe_allow_html=True
            )
            for item in _ms14_items[:8]:
                _days_away = (item["dt"].normalize() - pd.Timestamp(_snap_today)).days
                _day_str   = item["dt"].strftime("%b %-d")
                _rel_str   = f"in {_days_away}d" if _days_away > 0 else "today"
                _badge_html = ""
                if item.get("badge"):
                    _bg  = item.get("badge_bg", "var(--color-background-info)")
                    _tc  = item.get("badge_col", "var(--color-text-info)")
                    _badge_html = f"<span style='font-size:10px;padding:2px 7px;border-radius:4px;background:{_bg};color:{_tc};font-weight:500;flex-shrink:0;'>{item['badge']}</span>"
                st.markdown(
                    f"<div style='display:flex;align-items:center;gap:10px;padding:7px 0;border-bottom:1px solid rgba(128,128,128,0.18);'>"
                    f"<div style='min-width:44px;'>"
                    f"<div style='font-size:12px;font-weight:500;color:var(--color-text-primary);'>{_day_str}</div>"
                    f"<div style='font-size:10px;color:var(--color-text-secondary);'>{_rel_str}</div></div>"
                    f"<div style='width:7px;height:7px;border-radius:50%;background:{item['dot']};flex-shrink:0;'></div>"
                    f"<div style='flex:1;min-width:0;'>"
                    f"<div style='font-size:13px;font-weight:500;color:var(--color-text-primary);'>{item['name']}</div>"
                    f"<div style='font-size:11px;color:var(--color-text-secondary);margin-top:1px;'>{item['cust']} · {item['pid']}</div></div>"
                    f"{_badge_html}</div>",
                    unsafe_allow_html=True
                )
            st.markdown("</div>", unsafe_allow_html=True)



st.caption("PS Projects & Tools · Internal use only · Data loaded this session only")
