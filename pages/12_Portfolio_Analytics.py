"""
PS Tools — Portfolio Analytics
Manager-level portfolio view: team utilization trends, project risk distribution,
workload by consultant, and product mix analysis.
"""
import streamlit as st
import pandas as pd

st.session_state["current_page"] = "Portfolio Analytics"

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN, ACTIVE_EMPLOYEES,
    MILESTONE_COLS_MAP, get_role, is_manager, name_matches,
    resolve_name, get_ff_scope,
)
from shared.config import (
    EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP,
)

st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
<style>
    html,body,[class*="css"]{font-family:'Manrope',sans-serif!important}
    .section-label{font-size:11px;font-weight:700;text-transform:uppercase;
                   letter-spacing:.8px;color:#4472C4;margin-bottom:8px}
    .metric-card{border:1px solid rgba(128,128,128,.2);border-radius:8px;
                 padding:16px 20px;margin-bottom:4px}
    .metric-val{font-size:26px;font-weight:700;color:inherit}
    .metric-lbl{font-size:12px;opacity:.6;margin-top:2px}
    .metric-sub{font-size:11px;opacity:.5;margin-top:3px}
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

if role not in ("manager", "manager_only", "reporting_only"):
    st.info("Portfolio Analytics is available to managers and above.")
    st.stop()

# ── View As ───────────────────────────────────────────────────────────────────
view_as    = selected
_va_region = None
if role in ("manager", "manager_only", "reporting_only"):
    _pick = st.session_state.get("home_browse", "— My own view —")
    if _pick and _pick.startswith("── ") and _pick.endswith(" ──"):
        _va_region = _pick[3:-3].strip()
    elif _pick and _pick not in ("— My own view —", "— Select —", ""):
        view_as    = _pick
        _va_region = None

# ── Load data ─────────────────────────────────────────────────────────────────
df_drs = st.session_state.get("df_drs")
df_ns  = st.session_state.get("df_ns")

# ── Hero ──────────────────────────────────────────────────────────────────────
_zone_svg = """<svg style='position:absolute;right:-40px;top:50%;transform:translateY(-50%);
opacity:0.06;width:200px;height:200px;pointer-events:none'
viewBox='0 0 1482 1286.25' xmlns='http://www.w3.org/2000/svg'>
<g fill='#3B9EFF' fill-rule='evenodd'><path d='M975.127,924.953c2.608-2.68,1.744-5.496-.42-7.829l-57.415-61.872c-2.463-2.655-5.025-2.878-8.443-.991-10.398,5.739-19.024,12.314-27.949,19.885-83.252,70.621-197.471,155.494-298.93,195.556-17.993,7.105-35.256,13.178-54.191,17.329-62.148,13.627-131.853,15.491-192.702-5.298-64.93-22.183-113.878-68.722-142.715-130.542-28.647-61.415-22.393-131.406,11.352-189.217,2.598-2.793,1.405-6.055-1.389-8.184-35.341-26.918-40.303-33.439-69.367-65.686-1.449-1.607-4.102-2.401-5.903-1.138-13.105,9.189-23.232,20.534-33.172,32.961-16.499,20.629-29.73,42.605-38.718,67.541-5.127,10.469-8.378,20.486-10.885,32.065-13.633,62.973-7.701,128.685,17.402,188.142,23.839,56.463,65.297,103.638,114.77,139.169,32.418,23.283,66.848,42.548,103.476,58.385,25.142,10.871,50.281,18.994,76.934,25.12,96.392,22.153,188.876,4.496,276.774-38.393,42.916-20.94,83.188-45.685,121.922-73.568,75.733-54.514,154.643-126.72,219.571-193.435Z'/></g></svg>"""

_view_label = _va_region + " Team" if _va_region else (
    view_as if view_as != selected else "EMEA Team"
)
st.markdown(f"""
<div style='background:#050D1F;padding:32px 40px 28px;border-radius:10px;
            margin-bottom:24px;font-family:Manrope,sans-serif;
            position:relative;overflow:hidden'>
  {_zone_svg}
  <div style='font-size:10px;font-weight:700;letter-spacing:2.5px;
              text-transform:uppercase;color:#3B9EFF;margin-bottom:10px'>
      Professional Services · Tools</div>
  <h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif'>
      Portfolio Analytics</h1>
  <p style='color:rgba(255,255,255,0.45);margin:6px 0 0;font-size:14px;
            font-family:Manrope,sans-serif'>
      {_view_label} · {today.strftime("%B %Y")}</p>
</div>
""", unsafe_allow_html=True)

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
    _team_consultants = set(CONSULTANT_DROPDOWN)
else:
    # Manager viewing own team — get their region
    _mgr_loc = EMPLOYEE_LOCATION.get(selected, "")
    _mgr_region = PS_REGION_OVERRIDE.get(selected, PS_REGION_MAP.get(_mgr_loc, "Other"))
    _team_consultants = _get_region_consultants(_mgr_region)
    if not _team_consultants:
        _team_consultants = set(CONSULTANT_DROPDOWN)

# ── Filter DRS to team ────────────────────────────────────────────────────────
pm_col = df_drs.get("project_manager", pd.Series(dtype=str)).fillna("")

def _in_team(v):
    vl = resolve_name(str(v)).lower()
    vs = str(v).strip().lower()
    return any(
        vl == _n.lower() or vs == _n.lower() or
        vl == _n.split(",")[0].strip().lower() or
        vs == _n.split(",")[0].strip().lower()
        for _n in _team_consultants
    )

team_drs = df_drs[pm_col.apply(_in_team)].copy() if _team_consultants else df_drs.copy()

# Active projects (not on hold)
_ioh     = team_drs.get("_on_hold", pd.Series(False, index=team_drs.index)).astype(bool)
_active  = team_drs[~_ioh].copy()

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
_n_active   = len(_active)
_n_total    = len(team_drs)
_n_onhold   = int(_ioh.sum())
_n_team     = len(_team_consultants)

_rag_col    = _active.get("rag", pd.Series(dtype=str)).fillna("").str.strip().str.lower()
_n_red      = int((_rag_col == "red").sum())
_n_yellow   = int((_rag_col == "yellow").sum())
_n_at_risk  = _n_red + _n_yellow
_risk_pct   = round(100 * _n_at_risk / _n_active) if _n_active else 0

_oh_pct     = round(100 * _n_onhold / _n_total) if _n_total else 0

# Avg project duration (months from start to today for active)
_durations  = []
if "start_date" in _active.columns:
    _sd = pd.to_datetime(_active["start_date"], errors="coerce")
    _durations = [((today - s).days / 30.44) for s in _sd if pd.notna(s)]
_avg_dur = round(sum(_durations) / len(_durations), 1) if _durations else None

st.markdown('<div class="section-label">Portfolio Snapshot</div>', unsafe_allow_html=True)
ps1, ps2, ps3, ps4, ps5 = st.columns(5)
with ps1:
    st.markdown(f"""<div class="metric-card">
      <div class="metric-val">{_n_active}</div>
      <div class="metric-lbl">Active projects</div>
      <div class="metric-sub">{_n_team} consultants</div>
    </div>""", unsafe_allow_html=True)
with ps2:
    _rc = "#C0392B" if _risk_pct >= 20 else ("#F39C12" if _risk_pct >= 10 else "inherit")
    st.markdown(f"""<div class="metric-card">
      <div class="metric-val" style="color:{_rc}">{_risk_pct}%</div>
      <div class="metric-lbl">Projects at risk</div>
      <div class="metric-sub">{_n_red} Red · {_n_yellow} Yellow RAG</div>
    </div>""", unsafe_allow_html=True)
with ps3:
    _ohc = "#F39C12" if _oh_pct >= 20 else "inherit"
    st.markdown(f"""<div class="metric-card">
      <div class="metric-val" style="color:{_ohc}">{_oh_pct}%</div>
      <div class="metric-lbl">On-hold rate</div>
      <div class="metric-sub">{_n_onhold} of {_n_total} projects</div>
    </div>""", unsafe_allow_html=True)
with ps4:
    _dur_txt = f"{_avg_dur}mo" if _avg_dur else "—"
    _durc = "#F39C12" if (_avg_dur or 0) >= 9 else "inherit"
    st.markdown(f"""<div class="metric-card">
      <div class="metric-val" style="color:{_durc}">{_dur_txt}</div>
      <div class="metric-lbl">Avg project duration</div>
      <div class="metric-sub">active projects</div>
    </div>""", unsafe_allow_html=True)
with ps5:
    # Projects 9mo+ not yet at transition
    _pre_trans = _active[_active["phase"].fillna("").apply(
        lambda p: _pidx(p) < _pidx("08. ready for support transition")
    )] if "phase" in _active.columns else _active
    if "start_date" in _pre_trans.columns:
        _sd2 = pd.to_datetime(_pre_trans["start_date"], errors="coerce")
        _n_9mo = int((((today - _sd2).dt.days / 30.44) >= 9).sum())
    else:
        _n_9mo = 0
    _9moc = "#F39C12" if _n_9mo > 0 else "inherit"
    st.markdown(f"""<div class="metric-card">
      <div class="metric-val" style="color:{_9moc}">{_n_9mo}</div>
      <div class="metric-lbl">9+ months active</div>
      <div class="metric-sub">excl. phase 08+</div>
    </div>""", unsafe_allow_html=True)

# ── Consultant Workload Table ─────────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Consultant Workload</div>', unsafe_allow_html=True)

_cw_rows = []
for _cn in sorted(_team_consultants):
    _cm = df_drs[pm_col.apply(lambda v: name_matches(v, _cn))].copy()
    if _cm.empty: continue

    _cm_oh     = _cm.get("_on_hold", pd.Series(False, index=_cm.index)).astype(bool)
    _cm_active = _cm[~_cm_oh]
    _n_proj    = len(_cm_active)
    if _n_proj == 0: continue

    _cm_rag = _cm_active.get("rag", pd.Series(dtype=str)).fillna("").str.strip().str.lower()
    _cm_red = int((_cm_rag == "red").sum())
    _cm_yel = int((_cm_rag == "yellow").sum())

    # 9mo+ projects
    _cm_9mo = 0
    if "start_date" in _cm_active.columns:
        _cm_sd = pd.to_datetime(_cm_active["start_date"], errors="coerce")
        _cm_ph = _cm_active["phase"].fillna("").apply(_pidx) if "phase" in _cm_active.columns else pd.Series(-1, index=_cm_active.index)
        _trans_idx = _pidx("08. ready for support transition")
        _cm_pre = (_cm_ph < _trans_idx)
        _cm_months = (today - _cm_sd).dt.days / 30.44
        _cm_9mo = int((_cm_pre & (_cm_months >= 9)).sum())

    # Util % from NS if available
    _cm_util = "—"
    if df_ns is not None and not df_ns.empty:
        _month_key = today.strftime("%Y-%m")
        _cm_ns = df_ns[df_ns.get("employee", pd.Series(dtype=str)).fillna("").apply(
            lambda v: name_matches(v, _cn)
        )]
        _cm_month_ns = _cm_ns[
            pd.to_datetime(_cm_ns.get("date", pd.Series(dtype=str)), errors="coerce").dt.strftime("%Y-%m") == _month_key
        ] if not _cm_ns.empty and "date" in _cm_ns.columns else pd.DataFrame()
        if not _cm_month_ns.empty and "hours" in _cm_month_ns.columns:
            _cm_hrs = float(_cm_month_ns["hours"].sum() or 0)
            _work_days = pd.bdate_range(today.replace(day=1), today).shape[0]
            _avail = _work_days * 8
            if _avail > 0:
                _cm_util = f"{round(100 * _cm_hrs / _avail)}%"

    # WHS score if available
    _whs_score = st.session_state.get(f"_whs_{_cn}", "—")

    _cn_parts = [p.strip() for p in _cn.split(",")]
    _cn_disp  = f"{_cn_parts[1].split()[0]} {_cn_parts[0]}" if len(_cn_parts) == 2 else _cn

    _cw_rows.append({
        "Consultant":      _cn_disp,
        "Active projects": _n_proj,
        "Red RAG":         _cm_red,
        "Yellow RAG":      _cm_yel,
        "9mo+ projects":   _cm_9mo,
        "Util % (MTD)":    _cm_util,
        "_red":            _cm_red,
        "_yel":            _cm_yel,
    })

if _cw_rows:
    _cw_df = pd.DataFrame(_cw_rows).sort_values("Active projects", ascending=False)

    def _th(label, align="left"):
        return f"<th style='padding:8px 12px;text-align:{align};font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.55'>{label}</th>"

    _cw_html = """<div style='border:0.5px solid rgba(128,128,128,.2);border-radius:10px;overflow:hidden'>
    <table style='width:100%;border-collapse:collapse;font-family:Manrope,sans-serif;font-size:12px'>
    <thead><tr style='border-bottom:1px solid rgba(128,128,128,.2)'>"""
    _cw_html += _th("Consultant") + _th("Active projects","right") + _th("Red RAG","right") + _th("Yellow RAG","right") + _th("9mo+ projects","right") + _th("Util % (MTD)","right")
    _cw_html += "</tr></thead><tbody>"

    for _, row in _cw_df.iterrows():
        _red_c = f"color:#C0392B;font-weight:700" if row["_red"] > 0 else "opacity:.65"
        _yel_c = f"color:#F39C12;font-weight:700" if row["_yel"] > 0 else "opacity:.65"
        _9mo_c = f"color:#F39C12;font-weight:700" if row["9mo+ projects"] > 0 else "opacity:.65"
        _cw_html += f"""<tr style='border-bottom:0.5px solid rgba(128,128,128,.1)'>
          <td style='padding:8px 12px;font-weight:500'>{row["Consultant"]}</td>
          <td style='padding:8px 12px;text-align:right'>{row["Active projects"]}</td>
          <td style='padding:8px 12px;text-align:right;{_red_c}'>{row["Red RAG"]}</td>
          <td style='padding:8px 12px;text-align:right;{_yel_c}'>{row["Yellow RAG"]}</td>
          <td style='padding:8px 12px;text-align:right;{_9mo_c}'>{row["9mo+ projects"]}</td>
          <td style='padding:8px 12px;text-align:right;opacity:.75'>{row["Util % (MTD)"]}</td>
        </tr>"""
    _cw_html += "</tbody></table></div>"
    st.markdown(_cw_html, unsafe_allow_html=True)
    st.caption(f"{len(_cw_rows)} consultants · Util % requires NS Time Detail to be loaded · MTD = month to date")
else:
    st.info("No consultant project data found for this view.")

# ── Product Mix ───────────────────────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Product Mix</div>', unsafe_allow_html=True)

if "project_type" in _active.columns:
    _pt_counts = {}
    for pt in _active["project_type"].fillna("Unknown"):
        _pt = str(pt).strip()
        _prod = _pt.split(":")[-1].strip() if ":" in _pt else _pt
        _prod = _prod.split()[0] if _prod else "Unknown"
        _pt_counts[_prod] = _pt_counts.get(_prod, 0) + 1

    _pt_sorted = sorted(_pt_counts.items(), key=lambda x: x[1], ascending=False) \
        if False else sorted(_pt_counts.items(), key=lambda x: x[1], reverse=True)
    _pt_total  = sum(v for _, v in _pt_sorted)
    _max_count = max(v for _, v in _pt_sorted) if _pt_sorted else 1

    _pm_cols = st.columns(2)
    with _pm_cols[0]:
        for prod, cnt in _pt_sorted[:len(_pt_sorted)//2 + 1]:
            _pct = round(100 * cnt / _pt_total) if _pt_total else 0
            _bar_w = round(100 * cnt / _max_count)
            st.markdown(f"""
            <div style='margin-bottom:10px'>
              <div style='display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px'>
                <span>{prod}</span>
                <span style='opacity:.55'>{cnt} · {_pct}%</span>
              </div>
              <div class='bar-track'>
                <div class='bar-fill' style='width:{_bar_w}%;background:#3B9EFF'></div>
              </div>
            </div>""", unsafe_allow_html=True)
    with _pm_cols[1]:
        for prod, cnt in _pt_sorted[len(_pt_sorted)//2 + 1:]:
            _pct = round(100 * cnt / _pt_total) if _pt_total else 0
            _bar_w = round(100 * cnt / _max_count)
            st.markdown(f"""
            <div style='margin-bottom:10px'>
              <div style='display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px'>
                <span>{prod}</span>
                <span style='opacity:.55'>{cnt} · {_pct}%</span>
              </div>
              <div class='bar-track'>
                <div class='bar-fill' style='width:{_bar_w}%;background:#3B9EFF'></div>
              </div>
            </div>""", unsafe_allow_html=True)
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
                  <div style='display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px'>
                    <span>{ _abbr}</span>
                    <span style='opacity:.55'>{cnt} · {_pct}%</span>
                  </div>
                  <div class='bar-track'>
                    <div class='bar-fill' style='width:{_bar_w}%;background:#08A9B7'></div>
                  </div>
                </div>""", unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div style="font-size:11px;opacity:.4;text-align:center;margin-top:4px">PS Projects & Tools · Internal use only · Data loaded this session only</div>', unsafe_allow_html=True)
