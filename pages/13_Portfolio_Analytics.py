"""
PS Tools — Portfolio Analytics
Manager-level portfolio view: team utilization trends, project risk distribution,
workload by consultant, and product mix analysis.
"""
import streamlit as st
import math
import pandas as pd

st.session_state["current_page"] = "Portfolio Analytics"

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN, ACTIVE_EMPLOYEES,
    MILESTONE_COLS_MAP, get_role, is_manager, name_matches,
    resolve_name, get_ff_scope, DEFAULT_SCOPE,
)
from shared.constants import NO_ACCESS
try:
    from shared.constants import _LEAVERS as _LEAVER_SET
except ImportError:
    _LEAVER_SET = set()
_ALL_LEAVERS = set(NO_ACCESS) | _LEAVER_SET
from shared.utils import calc_consultant_util
from shared.config import (
    AVAIL_HOURS, EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP,
)

st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
<style>
    html,body,[class*="css"]{font-family:'Manrope',sans-serif!important}
    .section-label { font-size: 13px;font-weight:700;text-transform:uppercase;
                   letter-spacing:.8px;color:#4472C4;margin-bottom:8px}
    .metric-card{border:1px solid rgba(128,128,128,.2);border-radius:8px;
                 padding:16px 20px;margin-bottom:4px}
    .metric-val { font-size: 32px;font-weight:700;color:inherit}
    .metric-lbl { font-size: 14px;opacity:.6;margin-top:2px}
    .metric-sub{font-size:13px;opacity:.5;margin-top:3px}
    .divider{border:none;border-top:1px solid rgba(128,128,128,.2);margin:20px 0}
    .pill{display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;
          border-radius:10px;letter-spacing:.8px}
    .pill-green{background:rgba(39,174,96,.15);color:#27AE60}
    .pill-amber{background:rgba(243,156,18,.15);color:#F39C12}
    .pill-red{background:rgba(192,57,43,.15);color:#C0392B}
    .pill-grey{background:rgba(128,128,128,.12);color:inherit;opacity:.65}
    .pill-blue{background:rgba(59,158,255,.12);color:#3B9EFF}
    .bar-track{background:rgba(128,128,128,.12);border-radius:4px;height:8px;overflow:hidden}
    .bar-fill{height:8px;border-radius:4px}
</style>
""", unsafe_allow_html=True)

# ── Identity & auth ───────────────────────────────────────────────────────────
selected = st.session_state.get("consultant_name", "")
role     = get_role(selected) if selected else "consultant"
today    = pd.Timestamp.today().normalize()

if not selected:
    st.warning("Sign in on the Home page to view Portfolio Analytics.")
    st.stop()

# ── View As ───────────────────────────────────────────────────────────────────
view_as    = selected
_va_region = None
if role in ("manager", "manager_only", "reporting_only"):
    _pick = (st.session_state.get("_browse_passthrough") or
             st.session_state.get("home_browse", "")) or ""
    if not _pick:
        # No passthrough — show local selector
        from shared.constants import CONSULTANT_DROPDOWN as _CDP
        _bopts_pa = ["👥 All team"]
        _by_r_pa = {}
        for _cn_pa in sorted(_CDP):
            _loc_pa = EMPLOYEE_LOCATION.get(_cn_pa, "")
            _rg_pa = PS_REGION_OVERRIDE.get(_cn_pa, PS_REGION_MAP.get(_loc_pa, "Other"))
            _by_r_pa.setdefault(_rg_pa, []).append(_cn_pa)
        for _rg_pa in sorted(_by_r_pa.keys()):
            _bopts_pa.append(f"── {_rg_pa} ──")
            _bopts_pa.extend(_by_r_pa[_rg_pa])
        with st.sidebar:
            st.markdown("**View as:**")
            _pick = st.selectbox("View as", _bopts_pa, key="pa_view_as",
                                 label_visibility="collapsed")
    _pick_clean = _pick.replace("👥", "").strip().lower()
    if _pick and _pick.startswith("── ") and _pick.endswith(" ──"):
        _va_region = _pick[3:-3].strip()
    elif _pick_clean in ("all team", "") or _pick in ("👥 All team", "All team"):
        _va_region = None  # show all team — handled by manager_only path
        view_as    = selected
    elif _pick and _pick not in ("— My own view —", "— Select —", ""):
        view_as    = _pick
        _va_region = None

# ── Load data ─────────────────────────────────────────────────────────────────
df_drs = st.session_state.get("df_drs")
df_ns  = st.session_state.get("df_ns")

# ── Hero ──────────────────────────────────────────────────────────────────────
_zone_svg = """"""

if _va_region:
    _view_label = _va_region + " Team"
elif view_as != selected:
    _vp = [p.strip() for p in view_as.split(",")]
    _view_label = f"{_vp[1].split()[0]} {_vp[0]}" if len(_vp) == 2 else view_as
elif role in ("manager", "manager_only", "reporting_only"):
    _mgr_loc2    = EMPLOYEE_LOCATION.get(selected, "")
    _mgr_region2 = PS_REGION_OVERRIDE.get(selected, PS_REGION_MAP.get(_mgr_loc2, "Other"))
    _view_label  = _mgr_region2 + " Team" if _mgr_region2 else "My Team"
else:
    _vp = [p.strip() for p in selected.split(",")]
    _view_label = f"{_vp[1].split()[0]} {_vp[0]}" if len(_vp) == 2 else selected
_hero = st.empty()
_hero.markdown(f"<div style='background:#050D1F;padding:32px 40px 28px;border-radius:10px; margin-bottom:24px;font-family:Manrope,sans-serif; position:relative;overflow:hidden'> {_zone_svg} <div style='font-size:13px;font-weight:700;letter-spacing:2.5px; text-transform:uppercase;color:#3B9EFF;margin-bottom:10px'> Professional Services · Tools</div> <h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif'> Portfolio Analytics</h1> <p style='color:rgba(255,255,255,0.45);margin:6px 0 0;font-size:14px; font-family:Manrope,sans-serif'> {_view_label} · {today.strftime("%B %Y")}</p> </div>", unsafe_allow_html=True)

if df_drs is None:
    st.info("Load SS DRS on the Home page to view Portfolio Analytics.")
    st.stop()

# ── Build consultant list for this view ───────────────────────────────────────
def _get_region_consultants(region):
    _rc = set()
    for _n in CONSULTANT_DROPDOWN:
        _nl = EMPLOYEE_LOCATION.get(_n, "")
        _nr = PS_REGION_OVERRIDE.get(_n, PS_REGION_MAP.get(_nl, "Other"))
        if _nr == region:
            _rc.add(_n)
    return _rc

if _va_region:
    _team_consultants = _get_region_consultants(_va_region)
elif role == "manager_only":
    _team_consultants = set(ACTIVE_EMPLOYEES)
elif role in ("manager", "reporting_only"):
    _pick_curr = (st.session_state.get("_browse_passthrough") or
                  st.session_state.get("home_browse", "")) or ""
    _pick_curr_clean = _pick_curr.replace("👥", "").strip().lower()
    if _pick_curr_clean in ("all team", "") or _pick_curr in ("👥 All team", "All team"):
        _team_consultants = set(ACTIVE_EMPLOYEES)
    elif view_as != selected:
        _team_consultants = {view_as}
    else:
        # Default: manager's own region
        _mgr_loc    = EMPLOYEE_LOCATION.get(selected, "")
        _mgr_region = PS_REGION_OVERRIDE.get(selected, PS_REGION_MAP.get(_mgr_loc, "Other"))
        _team_consultants = _get_region_consultants(_mgr_region)
        if not _team_consultants:
            _team_consultants = set(CONSULTANT_DROPDOWN)
else:
    # Consultant — scope to themselves or View As target
    _team_consultants = {view_as}

# ── Filter DRS to team ────────────────────────────────────────────────────────
pm_col = df_drs.get("project_manager", pd.Series(dtype=str)).fillna("")

def _in_team(v):
    return any(name_matches(v, _n) for _n in _team_consultants)

team_drs = df_drs[pm_col.apply(_in_team)].copy() if _team_consultants else df_drs.copy()

# ── Leaver-assigned and unassigned projects (always from full df_drs) ─────────
def _is_leaver_pm(v):
    v = str(v).strip()
    if not v or v.lower() in ("", "nan", "none"): return False
    return any(name_matches(v, ln) for ln in _ALL_LEAVERS)

def _is_unassigned_pm(v):
    v = str(v).strip()
    return not v or v.lower() in ("", "nan", "none", "unassigned", "tbd")

_all_pm_col     = df_drs.get("project_manager", pd.Series(dtype=str)).fillna("")
_leaver_drs     = df_drs[_all_pm_col.apply(_is_leaver_pm)].copy()
_unassigned_drs = df_drs[_all_pm_col.apply(_is_unassigned_pm)].copy()
_n_leaver       = int(_leaver_drs["project_id"].nunique()) if "project_id" in _leaver_drs.columns and not _leaver_drs.empty else len(_leaver_drs)
_n_unassigned   = int(_unassigned_drs["project_id"].nunique()) if "project_id" in _unassigned_drs.columns and not _unassigned_drs.empty else len(_unassigned_drs)
# Derive pending_close from phase col — use contains for robustness against casing/spacing variants
if "_pending_close" in df_drs.columns:
    _pc_mask = df_drs["_pending_close"].astype(bool)
elif "phase" in df_drs.columns:
    _ph_norm = df_drs["phase"].fillna("").str.strip().str.lower()
    _pc_mask = _ph_norm.str.contains("complete", na=False) & _ph_norm.str.contains("pending", na=False)
else:
    _pc_mask = pd.Series(False, index=df_drs.index)
_pending_close_drs = df_drs[_pc_mask].copy()
_n_pending_close   = int(_pending_close_drs["project_id"].nunique()) if "project_id" in _pending_close_drs.columns and not _pending_close_drs.empty else len(_pending_close_drs)

# Active projects (not on hold)
_ioh     = team_drs.get("_on_hold", pd.Series(False, index=team_drs.index)).astype(bool)
_active  = team_drs[~_ioh].copy()
# ── Deduped project counts for metrics (avoids multi-row DRS inflation) ──
_id_col_dc    = "project_id" if "project_id" in _active.columns else "project_name"
_n_active_dc  = int(_active[_id_col_dc].nunique()) if not _active.empty else 0
_n_onhold_dc  = int(team_drs[_ioh]["project_id"].nunique()) if not team_drs.empty and "project_id" in team_drs.columns else int(_ioh.sum())


# Phase order
PHASE_ORDER = [
    "00. onboarding","01. requirements and design","02. configuration",
    "03. enablement/training","04. uat","05. prep for go-live",
    "06. go-live","07. data migration","08. ready for support transition",
    "09. phase 2 scoping",
]
def _pidx(p):
    pl = str(p).strip().lower()
    for i, ph in enumerate(PHASE_ORDER):
        if pl.startswith(ph[:6]) or ph in pl or pl in ph: return i
    return -1

# ── Portfolio Snapshot metrics ────────────────────────────────────────────────
_n_active   = _n_active_dc
_n_onhold   = _n_onhold_dc
_n_total    = _n_active_dc + _n_onhold_dc
# Count consultants who actually have projects in DRS (matches workload table)
if not team_drs.empty and "project_manager" in team_drs.columns:
    _active_pms = set(team_drs["project_manager"].dropna().str.strip().str.lower().unique())
    _n_team = sum(
        1 for _n in _team_consultants
        if any(name_matches(pm, _n) for pm in _active_pms)
    )
else:
    _n_team = len(_team_consultants)
_oh_denom   = _n_active + _n_onhold

_rag_col    = team_drs.get("rag", pd.Series(dtype=str)).fillna("").str.strip().str.lower()
_n_red      = int((_rag_col == "red").sum())
_n_yellow   = int((_rag_col == "yellow").sum())
_n_at_risk  = _n_red + _n_yellow
_risk_pct   = round(100 * _n_at_risk / (_n_active + _n_onhold)) if (_n_active + _n_onhold) else 0
_oh_pct     = round(100 * _n_onhold / _oh_denom) if _oh_denom else 0

_durations  = []
if "start_date" in _active.columns:
    _sd = pd.to_datetime(_active["start_date"], errors="coerce")
    _durations = [((today - s).days / 30.44) for s in _sd if pd.notna(s)]
_avg_dur = round(sum(_durations) / len(_durations), 1) if _durations else None

_pre_trans = _active[_active["phase"].fillna("").apply(
    lambda p: _pidx(p) < _pidx("08. ready for support transition")
)] if "phase" in _active.columns else _active
if "start_date" in _pre_trans.columns:
    _sd2 = pd.to_datetime(_pre_trans["start_date"], errors="coerce")
    _n_9mo = int((((today - _sd2).dt.days / 30.44) >= 9).sum())
else:
    _n_9mo = 0

st.markdown('<div class="section-label">Portfolio Snapshot</div>', unsafe_allow_html=True)

_lv_col = "#F39C12" if _n_leaver > 0 else "inherit"
_ua_col = "#C0392B" if _n_unassigned > 0 else "inherit"
_pc_col = "#F39C12" if _n_pending_close > 0 else "inherit"
_rc     = "#C0392B" if _risk_pct >= 20 else ("#F39C12" if _risk_pct >= 10 else "inherit")
_dur_txt2 = f"{_avg_dur}mo" if _avg_dur else "—"
_durc2  = "#F39C12" if (_avg_dur or 0) >= 9 else "inherit"
_9moc   = "#F39C12" if _n_9mo > 0 else "inherit"

st.markdown(f"""
<style>
.snap-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(110px, 1fr));
    gap: 8px;
    margin-bottom: 12px;
}}
.snap-card {{
    border: 1px solid rgba(128,128,128,.2);
    border-radius: 8px;
    padding: 14px 16px;
    min-width: 0;
}}
.snap-val  {{ font-size: 28px; font-weight: 700; line-height: 1.1; color: inherit; }}
.snap-lbl  {{ font-size: 13px; opacity: .6; margin-top: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.snap-sub  {{ font-size: 11px; opacity: .45; margin-top: 2px; }}
</style>
<div class="snap-grid">
  <div class="snap-card">
    <div class="snap-val">{_n_active}</div>
    <div class="snap-lbl">Active</div>
    <div class="snap-sub">{_n_team} consultants</div>
  </div>
  <div class="snap-card">
    <div class="snap-val" style="color:#F39C12">{_n_onhold}</div>
    <div class="snap-lbl">On hold</div>
    <div class="snap-sub">{_oh_pct}% of total</div>
  </div>
  <div class="snap-card">
    <div class="snap-val" style="color:{_lv_col}">{_n_leaver}</div>
    <div class="snap-lbl">Need reassignment</div>
  </div>
  <div class="snap-card">
    <div class="snap-val" style="color:{_ua_col}">{_n_unassigned}</div>
    <div class="snap-lbl">Unassigned</div>
  </div>
  <div class="snap-card">
    <div class="snap-val" style="color:{_pc_col}">{_n_pending_close}</div>
    <div class="snap-lbl">Pending close</div>
    <div class="snap-sub">Active status · closed phase</div>
  </div>
  <div class="snap-card">
    <div class="snap-val" style="color:{_rc}">{_risk_pct}%</div>
    <div class="snap-lbl">Projects at risk</div>
    <div class="snap-sub">{_n_red} Red · {_n_yellow} Yellow</div>
  </div>
  <div class="snap-card">
    <div class="snap-val" style="color:{_durc2}">{_dur_txt2}</div>
    <div class="snap-lbl">Avg duration</div>
    <div class="snap-sub">active projects</div>
  </div>
  <div class="snap-card">
    <div class="snap-val" style="color:{_9moc}">{_n_9mo}</div>
    <div class="snap-lbl">9+ months active</div>
    <div class="snap-sub">excl. phase 08+</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Expandable detail for leaver-assigned and unassigned ─────────────────────
if _n_leaver > 0 or _n_unassigned > 0 or _n_pending_close > 0:
    _exp1, _exp2, _exp3 = st.columns(3)
    if _n_leaver > 0:
        with _exp1:
            with st.expander(f"⚠️ {_n_leaver} project{'s' if _n_leaver != 1 else ''} need reassignment — view"):
                _lv_cols = ["project_name","project_manager","project_type","_on_hold","ps_region"]
                _lv_show = _leaver_drs[[c for c in _lv_cols if c in _leaver_drs.columns]].copy()
                _lv_show = _lv_show.rename(columns={
                    "project_name": "Project", "project_manager": "Former PM",
                    "project_type": "Product", "_on_hold": "On Hold", "ps_region": "Region"
                })
                if "On Hold" in _lv_show.columns:
                    _lv_show["Status"] = _lv_show["On Hold"].apply(lambda x: "On Hold" if x else "Active")
                    _lv_show = _lv_show.drop(columns=["On Hold"])
                st.dataframe(_lv_show, use_container_width=True, hide_index=True)
    if _n_unassigned > 0:
        with _exp2:
            with st.expander(f"⚠️ {_n_unassigned} unassigned project{'s' if _n_unassigned != 1 else ''} — view & assign"):
                _ua_cols = ["project_name","project_type","_on_hold","start_date"]
                _ua_show = _unassigned_drs[[c for c in _ua_cols if c in _unassigned_drs.columns]].copy()
                _ua_show = _ua_show.rename(columns={
                    "project_name": "Project", "project_type": "Product",
                    "_on_hold": "On Hold", "start_date": "Start Date"
                })
                if "On Hold" in _ua_show.columns:
                    _ua_show["Status"] = _ua_show["On Hold"].apply(lambda x: "On Hold" if x else "Active")
                    _ua_show = _ua_show.drop(columns=["On Hold"])
                if "Start Date" in _ua_show.columns:
                    _ua_show["Start Date"] = pd.to_datetime(_ua_show["Start Date"], errors="coerce").dt.strftime("%b %Y").fillna("—")
                st.dataframe(_ua_show, use_container_width=True, hide_index=True)
    if _n_pending_close > 0:
        with _exp3:
            with st.expander(f"⚠️ {_n_pending_close} project{'s' if _n_pending_close != 1 else ''} pending close — view & action"):
                _pc_cols = ["project_name","project_manager","project_type","status","ps_region"]
                _pc_show = _pending_close_drs[[c for c in _pc_cols if c in _pending_close_drs.columns]].copy()
                _pc_show = _pc_show.rename(columns={
                    "project_name": "Project", "project_manager": "PM",
                    "project_type": "Product", "status": "Status", "ps_region": "Region"
                })
                st.caption("Phase 10 (Complete/Pending Final Billing) with status not yet Closed. Update status in DRS.")
                st.dataframe(_pc_show, use_container_width=True, hide_index=True)

# ── Product Mix ───────────────────────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Product Mix</div>', unsafe_allow_html=True)

if "project_type" in _active.columns:
    # ── Family → type keyword mapping ─────────────────────────────────────────
    # Each tuple: (family_name, family_color, {type_keyword: display_label})
    # project_type is matched case-insensitively after stripping "ZoneApp:" prefix
    _FAMILIES = [
        ("ZoneApps", "#1D9E75", {
            "capture":             "Capture",
            "approvals":           "Approvals",
            "reconcile 2.0":       "Reconcile 2.0",
            "reconcile psp":       "Reconcile PSP",
            "reconcile":           "Reconcile",
            "e-invoicing":         "e-Invoicing",
            "einvoicing":          "e-Invoicing",
            "cc statement import": "CC Statement Import",
            "sftp connector":      "SFTP Connector",
            "sftp":                "SFTP Connector",
            "payments":            "Payments",
            "ap payment":          "AP Payment",
            "premium":             "Premium",
            "consulting":          "Consulting",
        }),
        ("ZoneBilling", "#3266ad", {
            "zab implementation":    "ZAB Implementation",
            "zab partner":           "ZAB Partner Impl",
            "zb_standard":           "ZB_Standard",
            "zb standard":           "ZB_Standard",
            "zb_premium":            "ZB_Premium",
            "zb premium":            "ZB_Premium",
            "optimization audit":    "Optimization Audit",
            "subscription services": "Subscription Services",
            "optimization":          "Optimization",
        }),
        ("ZonePayroll", "#D85A30", {
            "zep: implementation": "ZEP Implementation",
            "zep: optimization":   "ZEP Optimization",
            "zep implementation":  "ZEP Implementation",
            "zep optimization":    "ZEP Optimization",
            "zep":                 "ZEP",
            "implementation":      "Implementation",
            "optimization":        "Optimization",
            "support":             "Support",
        }),
        ("ZoneReporting", "#888780", {
            "install":      "Install / DWH",
            "dwh":          "Install / DWH",
            "optimization": "Optimization",
        }),
    ]

    def _classify_to_family(raw_pt):
        """Return (family_name, type_label) for a raw project_type string."""
        pt = str(raw_pt).strip()
        # Strip ZoneApp:, ZoneBill:, ZonePay:, ZoneRpt: prefixes
        if ":" in pt:
            pt_clean = pt.split(":", 1)[1].strip().lower()
            prefix   = pt.split(":", 1)[0].strip().lower()
        else:
            pt_clean = pt.lower()
            prefix   = ""
        # Determine family from prefix first
        prefix_map = {
            "zoneapp": 0, "zoneapps": 0,
            "zonebill": 1, "zonebilling": 1,
            "zonepay": 2, "zonepayroll": 2,
            "zonerpt": 3, "zonereporting": 3,
        }
        family_idx = prefix_map.get(prefix, None)
        # Search in the identified family (or all families if prefix unknown)
        search_families = [_FAMILIES[family_idx]] if family_idx is not None else _FAMILIES
        for fam_name, fam_col, kw_map in search_families:
            # Longest-match first to avoid "reconcile" matching before "reconcile 2.0"
            for kw in sorted(kw_map.keys(), key=len, reverse=True):
                if kw in pt_clean:
                    return fam_name, kw_map[kw]
        return "Other", pt.split(":")[-1].strip() or pt

    # Count by family and type
    _fam_type_counts = {}   # family_name → {type_label: count}
    _fam_totals      = {}   # family_name → total
    _other_counts    = {}   # for unmatched types
    _grand_total     = 0

    for raw_pt in _active["project_type"].fillna("Unknown"):
        fam, typ = _classify_to_family(raw_pt)
        if fam == "Other":
            _other_counts[typ] = _other_counts.get(typ, 0) + 1
        else:
            if fam not in _fam_type_counts:
                _fam_type_counts[fam] = {}
            _fam_type_counts[fam][typ] = _fam_type_counts[fam].get(typ, 0) + 1
        _fam_totals[fam] = _fam_totals.get(fam, 0) + 1
        _grand_total += 1

    _pm_left, _pm_right = st.columns(2)
    _fam_left  = [f for f in _FAMILIES if f[0] in ("ZoneApps", "ZoneReporting")]
    _fam_right = [f for f in _FAMILIES if f[0] in ("ZoneBilling", "ZonePayroll")]
    if _other_counts:
        _fam_right.append(("Other", "#B4B2A9", _other_counts))

    def _render_family(fam_name, fam_col, kw_map_or_counts):
        fam_total = _fam_totals.get(fam_name, 0)
        if fam_total == 0:
            return
        fam_pct = round(100 * fam_total / _grand_total) if _grand_total else 0
        # Family header + family-level bar (% of total portfolio)
        st.markdown(
            f"<div style='margin-bottom:6px'>"
            f"<div style='display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px'>"
            f"<span style='font-size:13px;font-weight:600;color:var(--color-text-primary)'>{fam_name}</span>"
            f"<span style='font-size:12px;color:var(--color-text-secondary)'>{fam_total} · {fam_pct}%</span>"
            f"</div>"
            f"<div style='background:rgba(128,128,128,.12);border-radius:3px;height:5px'>"
            f"<div style='height:5px;border-radius:3px;width:{fam_pct}%;background:{fam_col}'></div>"
            f"</div></div>",
            unsafe_allow_html=True
        )
        # Type rows indented
        type_counts = _fam_type_counts.get(fam_name, {}) if fam_name != "Other" else _other_counts
        if not type_counts:
            return
        type_max = max(type_counts.values()) or 1
        for typ_lbl, typ_cnt in sorted(type_counts.items(), key=lambda x: x[1], reverse=True):
            bar_w = round(100 * typ_cnt / type_max)
            st.markdown(
                f"<div style='display:flex;align-items:center;gap:8px;margin-bottom:5px;padding-left:10px'>"
                f"<span style='font-size:12px;color:var(--color-text-secondary);flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>{typ_lbl}</span>"
                f"<div style='background:rgba(128,128,128,.12);border-radius:2px;height:3px;flex:2'>"
                f"<div style='height:3px;border-radius:2px;width:{bar_w}%;background:{fam_col};opacity:0.65'></div></div>"
                f"<span style='font-size:12px;color:var(--color-text-tertiary);min-width:24px;text-align:right'>{typ_cnt}</span>"
                f"</div>",
                unsafe_allow_html=True
            )
        st.markdown("<div style='margin-bottom:14px'></div>", unsafe_allow_html=True)

    with _pm_left:
        for fam_name, fam_col, kw_map in _fam_left:
            _render_family(fam_name, fam_col, kw_map)

    with _pm_right:
        for fam_name, fam_col, kw_map in _fam_right:
            _render_family(fam_name, fam_col, kw_map)

else:
    st.info("No project type data available.")

# ── Phase Distribution ────────────────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Phase Distribution</div>', unsafe_allow_html=True)

if "phase" in _active.columns:
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
    _ph_counts = {}
    for ph in _active["phase"].fillna("Unassigned"):
        _ph_clean = str(ph).strip()
        _ph_counts[_ph_clean] = _ph_counts.get(_ph_clean, 0) + 1

    _ph_sorted = sorted(
        _ph_counts.items(),
        key=lambda x: (_pidx(x[0]) if _pidx(x[0]) >= 0 else 999)
    )
    _ph_total = sum(v for _, v in _ph_sorted)
    _max_ph   = max(v for _, v in _ph_sorted) if _ph_sorted else 1

    _ph_cols = st.columns(2)
    _ph_left  = [(ph, cnt) for ph, cnt in _ph_sorted if _pidx(ph) < 5 or _pidx(ph) == -1]
    _ph_right = [(ph, cnt) for ph, cnt in _ph_sorted if _pidx(ph) >= 5 and _pidx(ph) != -1]

    for col_items, col_widget in [(_ph_left, _ph_cols[0]), (_ph_right, _ph_cols[1])]:
        with col_widget:
            for ph, cnt in col_items:
                _abbr = _PHASE_ABBREV.get(str(ph).strip().lower(), str(ph).split(".")[-1].strip()[:24])
                if ph in ("Unassigned", ""):
                    _abbr = "Unassigned"
                _pct  = round(100 * cnt / _ph_total) if _ph_total else 0
                _bar_w = round(100 * cnt / _max_ph)
                st.markdown(f"""
                <div style='margin-bottom:10px'>
                  <div style='display:flex;justify-content:space-between;font-size:14px;margin-bottom:4px'>
                    <span>{ _abbr}</span>
                    <span style='opacity:.55'>{cnt} · {_pct}%</span>
                  </div>
                  <div class='bar-track'>
                    <div class='bar-fill' style='width:{_bar_w}%;background:#08A9B7'></div>
                  </div>
                </div>""", unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── Consultant Workload Table ─────────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Consultant Workload</div>', unsafe_allow_html=True)

# Pre-build scope data for all active projects for overrun/near-limit lookup
_SHORT_DELIVERY = {"approvals", "cc", "sftp", "additional subsidiary", "cc statement import"}
_scope_by_pm = {}  # pm_name → {overrun: n, near_limit: n}

def _safe_f(v):
    try: return float(str(v).replace(",","").strip() or 0)
    except: return 0.0

for _, _sr in _active.iterrows():
    _pm_s  = str(_sr.get("project_manager","") or "")
    _pt_s  = str(_sr.get("project_type","") or "")
    _pn_s  = str(_sr.get("project_name","") or "")
    _actual = _safe_f(_sr.get("actual_hours",0))
    _budget = _safe_f(_sr.get("budgeted_hours",0))
    _co     = _safe_f(_sr.get("change_order",0))
    _scope  = get_ff_scope(_pt_s, _pn_s) or ((_budget + _co) if (_budget + _co) > 0 else None)
    try:
        _scope_f = float(_scope) if _scope is not None else 0.0
    except (TypeError, ValueError):
        _scope_f = 0.0
    if _scope_f <= 0 or _actual <= 0: continue
    try:
        _burn_raw = 100 * _actual / _scope_f
        if not math.isfinite(_burn_raw): continue
        _burn = round(_burn_raw)
    except (ZeroDivisionError, ValueError, OverflowError):
        continue
    _pm_key = _pm_s.strip()
    if _pm_key not in _scope_by_pm:
        _scope_by_pm[_pm_key] = {"overrun": 0, "near_limit": 0}
    if _burn > 100:
        _scope_by_pm[_pm_key]["overrun"] += 1
    elif _burn >= 80:
        _scope_by_pm[_pm_key]["near_limit"] += 1

# Schedule status helper (mirrors Project Health logic)
def _sched_status_pa(row):
    gld = row.get("effective_go_live_date") or row.get("go_live_date")
    phase = str(row.get("phase","") or "").strip().lower()
    pi = _pidx(phase)
    if pd.isna(gld) if gld is None else pd.isna(pd.Timestamp(gld)): return "no_date"
    days = (pd.Timestamp(gld) - today).days
    if pi >= _pidx("06. go-live"): return "delivered"
    if days > 14:   return "on_track"
    if days >= 0:   return "going_live"
    if days >= -30: return "delayed"
    return "at_risk"

team_drs["_sched_pa"] = team_drs.apply(_sched_status_pa, axis=1)
_active = team_drs[~_ioh].copy()  # re-slice so _active has _sched_pa

# Pre-score WHS once for all consultants
_whs_lookup = {}
try:
    from shared.whs import score_projects, build_consultant_summary
    if df_drs is not None and not df_drs.empty:
        _whs_df_all   = score_projects(df_drs)          # returns DataFrame directly
        _whs_summ_all, _ = build_consultant_summary(_whs_df_all)  # returns (df, missing_count)
        for _, _wr in _whs_summ_all.iterrows():
            _wcs = str(_wr.get("project_manager","") or "")
            for _cn2 in _team_consultants:
                if name_matches(_wcs, _cn2):
                    _whs_lookup[_cn2] = round(float(_wr.get("total_score", 0)), 1)
except Exception:
    pass

# Pre-compute month key for util
_month_key_pa = today.strftime("%Y-%m")

_cw_rows = []
for _cn in sorted(_team_consultants):
    _cm = df_drs[pm_col.apply(lambda v: name_matches(v, _cn))].copy()
    if _cm.empty: continue

    _cm_oh     = _cm.get("_on_hold", pd.Series(False, index=_cm.index)).astype(bool)
    _cm_active = _cm[~_cm_oh]
    _n_proj    = len(_cm_active)
    _n_oh      = int(_cm_oh.sum())
    if _n_proj == 0 and _n_oh == 0: continue

    # RAG — includes on-hold projects, matching Daily Briefing behaviour
    _cm_rag = _cm.get("rag", pd.Series(dtype=str)).fillna("").str.strip().str.lower()
    _cm_red = int((_cm_rag == "red").sum())
    _cm_yel = int((_cm_rag == "yellow").sum())

    # Schedule counts — compute directly on this consultant's active projects
    _cm_sched_vals = _cm_active.apply(_sched_status_pa, axis=1)
    _n_on_track = int(_cm_sched_vals.isin(["on_track","going_live","delivered"]).sum())
    _n_delayed  = int((_cm_sched_vals == "delayed").sum())
    _n_at_risk  = int((_cm_sched_vals == "at_risk").sum())

    # WHS score — looked up from pre-scored dict
    _whs_val = _whs_lookup.get(_cn)

    # Scope counts — match consultant name against scope_by_pm dict
    _n_overrun = 0; _n_near_limit = 0
    for _sp_key, _sp_val in _scope_by_pm.items():
        if name_matches(_sp_key, _cn):
            _n_overrun    += _sp_val["overrun"]
            _n_near_limit += _sp_val["near_limit"]

    # Util % MTD — full scope-capped calculation matching Daily Briefing exactly
    _cm_util = "—"
    if df_ns is not None and not df_ns.empty:
        _cm_ns = df_ns[df_ns.get("employee", pd.Series(dtype=str)).fillna("").apply(
            lambda v: name_matches(v, _cn)
        )]
        _cn_loc = EMPLOYEE_LOCATION.get(_cn, "")
        if isinstance(_cn_loc, tuple): _cn_loc = _cn_loc[0]
        _cn_region = PS_REGION_OVERRIDE.get(_cn, PS_REGION_MAP.get(_cn_loc, ""))
        _cn_avail  = (AVAIL_HOURS.get(_cn_loc, {}).get(_month_key_pa)
                   or AVAIL_HOURS.get(_cn_region, {}).get(_month_key_pa))
        if _cn_avail and not _cm_ns.empty:
            _util_result = calc_consultant_util(
                _cm_ns, _month_key_pa, DEFAULT_SCOPE, float(_cn_avail)
            )
            if _util_result["util_pct"] is not None:
                _cm_util = f"{_util_result['util_pct']}%"

    _cn_parts = [p.strip() for p in _cn.split(",")]
    _cn_disp  = f"{_cn_parts[1].split()[0]} {_cn_parts[0]}" if len(_cn_parts) == 2 else _cn

    _cw_rows.append({
        "_cn":         _cn,
        "Consultant":  _cn_disp,
        "_whs":        _whs_val,
        "_n_proj":     _n_proj,
        "_n_oh":       _n_oh,
        "_red":        _cm_red,
        "_yel":        _cm_yel,
        "_on_track":   _n_on_track,
        "_delayed":    _n_delayed,
        "_at_risk":    _n_at_risk,
        "_overrun":    _n_overrun,
        "_near_limit": _n_near_limit,
        "_util":       _cm_util,
    })

if _cw_rows:
    _cw_df = pd.DataFrame(_cw_rows)

    # ── Sort controls ─────────────────────────────────────────────────────────
    _sort_map = {
        "Consultant":    "Consultant",
        "WHS":           "_whs",
        "Active":        "_n_proj",
        "On hold":       "_n_oh",
        "Red RAG":       "_red",
        "Yellow RAG":    "_yel",
        "On track":      "_on_track",
        "Delayed":       "_delayed",
        "At risk (sch)": "_at_risk",
        "Overrun":       "_overrun",
        "Near limit":    "_near_limit",
        "Util %":        "_util",
    }
    _cw_sc1, _cw_sc2, _cw_sc3 = st.columns([2,1,3])
    with _cw_sc1:
        _cw_sort_lbl = st.selectbox("Sort by", list(_sort_map.keys()), index=2,
                                     key="cw_sort_col", label_visibility="collapsed")
    with _cw_sc2:
        _cw_sort_dir = st.selectbox("Direction", ["↓ Desc","↑ Asc"], index=0,
                                     key="cw_sort_dir", label_visibility="collapsed")
    with _cw_sc3:
        st.markdown(f'<div style="font-size:11px;opacity:.4;padding-top:10px">' +
                    f'Sorting by {_cw_sort_lbl} · {len(_cw_rows)} consultants</div>',
                    unsafe_allow_html=True)
    _cw_sort_col = _sort_map[_cw_sort_lbl]
    _cw_asc      = _cw_sort_dir.startswith("↑")
    if _cw_sort_col == "_util":
        # Sort util % numerically
        def _util_key(v):
            try: return int(str(v).replace("%","").strip())
            except: return -1
        _cw_df["_util_sort"] = _cw_df["_util"].apply(_util_key)
        _cw_df = _cw_df.sort_values("_util_sort", ascending=_cw_asc)
    else:
        _cw_df = _cw_df.sort_values(_cw_sort_col, ascending=_cw_asc,
                                     na_position="last" if not _cw_asc else "first")

    def _cv(val, col="", zero_muted=True):
        """Render a table cell value with colour."""
        if val == "—" or val is None:
            return f"<td style='padding:7px 10px;text-align:right;opacity:.35'>—</td>"
        if zero_muted and isinstance(val, (int,float)) and val == 0:
            return f"<td style='padding:7px 10px;text-align:right;opacity:.3'>0</td>"
        _style = f"padding:7px 10px;text-align:right;{col}"
        return f"<td style='{_style}'>{val}</td>"

    def _whs_col(v):
        if v is None: return "opacity:.35"
        from shared.whs import workload_level as _wl
        _, _c = _wl(float(v))
        return f"color:{_c};font-weight:700"

    def _util_col(v):
        if v == "—": return "opacity:.35"
        try:
            p = int(str(v).replace("%",""))
            if p >= 70: return "color:#27AE60;font-weight:700"
            if p >= 60: return "color:#F39C12;font-weight:700"
            return "color:#C0392B;font-weight:700"
        except: return ""

    _BD = "border-left:0.5px solid rgba(128,128,128,.15)"  # group divider

    _cw_html = f"""<div style='border:0.5px solid rgba(128,128,128,.2);border-radius:10px;overflow:hidden'>
<table style='width:100%;border-collapse:collapse;font-family:Manrope,sans-serif;font-size:14px'>
<thead>
<tr style='border-bottom:0.5px solid rgba(128,128,128,.15)'>
  <th rowspan='2' style='padding:7px 10px;text-align:left;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.55;border-right:0.5px solid rgba(128,128,128,.15);vertical-align:bottom'>Consultant</th>
  <th colspan='3' style='padding:4px 10px;text-align:center;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.45;border-bottom:0.5px solid rgba(128,128,128,.15)'>Workload</th>
  <th colspan='2' style='padding:4px 10px;text-align:center;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.45;border-bottom:0.5px solid rgba(128,128,128,.15);border-left:0.5px solid rgba(128,128,128,.15)'>RAG</th>
  <th colspan='3' style='padding:4px 10px;text-align:center;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.45;border-bottom:0.5px solid rgba(128,128,128,.15);border-left:0.5px solid rgba(128,128,128,.15)'>Schedule</th>
  <th colspan='2' style='padding:4px 10px;text-align:center;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.45;border-bottom:0.5px solid rgba(128,128,128,.15);border-left:0.5px solid rgba(128,128,128,.15)'>Scope</th>
  <th rowspan='2' style='padding:7px 10px;text-align:right;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.55;border-left:0.5px solid rgba(128,128,128,.15);vertical-align:bottom'>Util %</th>
</tr>
<tr style='border-bottom:1px solid rgba(128,128,128,.2)'>
  <th style='padding:3px 10px 7px;text-align:right;font-size:12px;font-weight:700;opacity:.55'>WHS</th>
  <th style='padding:3px 10px 7px;text-align:right;font-size:12px;font-weight:700;opacity:.55'>Active</th>
  <th style='padding:3px 10px 7px;text-align:right;font-size:12px;font-weight:700;opacity:.55'>On hold</th>
  <th style='padding:3px 10px 7px;text-align:right;font-size:12px;font-weight:700;opacity:.55;border-left:0.5px solid rgba(128,128,128,.15)'>🔴</th>
  <th style='padding:3px 10px 7px;text-align:right;font-size:12px;font-weight:700;opacity:.55'>🟡</th>
  <th style='padding:3px 10px 7px;text-align:right;font-size:12px;font-weight:700;opacity:.55;border-left:0.5px solid rgba(128,128,128,.15)'>On track</th>
  <th style='padding:3px 10px 7px;text-align:right;font-size:12px;font-weight:700;opacity:.55'>Delayed</th>
  <th style='padding:3px 10px 7px;text-align:right;font-size:12px;font-weight:700;opacity:.55'>At risk</th>
  <th style='padding:3px 10px 7px;text-align:right;font-size:12px;font-weight:700;opacity:.55;border-left:0.5px solid rgba(128,128,128,.15)'>Overrun</th>
  <th style='padding:3px 10px 7px;text-align:right;font-size:12px;font-weight:700;opacity:.55'>Near limit</th>
</tr>
</thead><tbody>"""

    for _, row in _cw_df.iterrows():
        _whs_v   = row["_whs"]
        _whs_disp = str(_whs_v) if _whs_v is not None else "—"
        _util_v  = row["_util"]
        _cw_html += f"<tr style='border-bottom:0.5px solid rgba(128,128,128,.1)'>"
        _cw_html += f"<td style='padding:7px 10px;font-weight:500;border-right:0.5px solid rgba(128,128,128,.15)'>{row['Consultant']}</td>"
        _cw_html += f"<td style='padding:7px 10px;text-align:right;{_whs_col(_whs_v)}'>{_whs_disp}</td>"
        _cw_html += _cv(row["_n_proj"], zero_muted=False)
        _cw_html += _cv(row["_n_oh"], "color:#F39C12;font-weight:700" if row["_n_oh"] > 0 else "")
        _cw_html += f"<td style='padding:7px 10px;text-align:right;border-left:0.5px solid rgba(128,128,128,.15);{'color:#C0392B;font-weight:700' if row['_red']>0 else 'opacity:.3'}'>{row['_red']}</td>"
        _cw_html += _cv(row["_yel"], "color:#F39C12;font-weight:700" if row["_yel"] > 0 else "")
        _cw_html += f"<td style='padding:7px 10px;text-align:right;border-left:0.5px solid rgba(128,128,128,.15);{'color:#27AE60;font-weight:700' if row['_on_track']>0 else 'opacity:.3'}'>{row['_on_track']}</td>"
        _cw_html += _cv(row["_delayed"],  "color:#F39C12;font-weight:700" if row["_delayed"] > 0 else "")
        _cw_html += _cv(row["_at_risk"],  "color:#C0392B;font-weight:700" if row["_at_risk"] > 0 else "")
        _cw_html += f"<td style='padding:7px 10px;text-align:right;border-left:0.5px solid rgba(128,128,128,.15);{'color:#C0392B;font-weight:700' if row['_overrun']>0 else 'opacity:.3'}'>{row['_overrun']}</td>"
        _cw_html += _cv(row["_near_limit"],"color:#F39C12;font-weight:700" if row["_near_limit"] > 0 else "")
        _cw_html += f"<td style='padding:7px 10px;text-align:right;border-left:0.5px solid rgba(128,128,128,.15);{_util_col(_util_v)}'>{_util_v}</td>"
        _cw_html += "</tr>"

    _cw_html += "</tbody></table></div>"
    st.markdown(_cw_html, unsafe_allow_html=True)
    st.caption(f"{len(_cw_rows)} consultants · Schedule based on effective go-live date · Scope excl. short-delivery products · Util % requires NS Time Detail · MTD = month to date")
else:
    st.info("No consultant project data found for this view.")

st.markdown('<div style="font-size:11px;opacity:.4;text-align:center;margin-top:4px">PS Projects & Tools · Internal use only · Data loaded this session only</div>', unsafe_allow_html=True)
