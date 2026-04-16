"""
PS Tools — Project Health
Delivery performance: schedule variance, milestone health, scope health.
"""
import streamlit as st
import pandas as pd
from datetime import date

st.session_state["current_page"] = "Project Health"

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN,
    MILESTONE_COLS_MAP, get_role, is_manager,
    name_matches, get_ff_scope,
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
</style>
""", unsafe_allow_html=True)

# ── Identity & auth ───────────────────────────────────────────────────────────
selected = st.session_state.get("consultant_name", "")
role     = get_role(selected) if selected else "consultant"
today    = pd.Timestamp.today().normalize()

if not selected:
    st.warning("Sign in on the Home page to view Project Health.")
    st.stop()

# ── Hero ──────────────────────────────────────────────────────────────────────
_zone_svg = """<svg style='position:absolute;right:-40px;top:50%;transform:translateY(-50%);
opacity:0.06;width:200px;height:200px;pointer-events:none'
viewBox='0 0 1482 1286.25' xmlns='http://www.w3.org/2000/svg'>
<g fill='#3B9EFF' fill-rule='evenodd'><path d='M975.127,924.953c2.608-2.68,1.744-5.496-.42-7.829l-57.415-61.872c-2.463-2.655-5.025-2.878-8.443-.991-10.398,5.739-19.024,12.314-27.949,19.885-83.252,70.621-197.471,155.494-298.93,195.556-17.993,7.105-35.256,13.178-54.191,17.329-62.148,13.627-131.853,15.491-192.702-5.298-64.93-22.183-113.878-68.722-142.715-130.542-28.647-61.415-22.393-131.406,11.352-189.217,2.598-2.793,1.405-6.055-1.389-8.184-35.341-26.918-40.303-33.439-69.367-65.686-1.449-1.607-4.102-2.401-5.903-1.138-13.105,9.189-23.232,20.534-33.172,32.961-16.499,20.629-29.73,42.605-38.718,67.541-5.127,10.469-8.378,20.486-10.885,32.065-13.633,62.973-7.701,128.685,17.402,188.142,23.839,56.463,65.297,103.638,114.77,139.169,32.418,23.283,66.848,42.548,103.476,58.385,25.142,10.871,50.281,18.994,76.934,25.12,96.392,22.153,188.876,4.496,276.774-38.393,42.916-20.94,83.188-45.685,121.922-73.568,75.733-54.514,154.643-126.72,219.571-193.435Z'/></g></svg>"""

df_drs = st.session_state.get("df_drs")

# ── View As ───────────────────────────────────────────────────────────────────
view_as    = selected
_va_region = None
if role in ("manager", "manager_only", "reporting_only"):
    _pick = st.session_state.get("home_browse", "— My own view —")
    if _pick.startswith("[ ") and _pick.endswith(" ]"):
        _va_region = _pick[3:-3].strip()
    elif _pick not in ("— My own view —", ""):
        view_as    = _pick
        _va_region = None

# ── Filter DRS to this consultant/region ─────────────────────────────────────
my_drs = pd.DataFrame()
if df_drs is not None and not df_drs.empty:
    pm_col = df_drs.get("project_manager", pd.Series(dtype=str)).fillna("")
    if _va_region and role in ("manager", "manager_only", "reporting_only"):
        from shared.constants import ACTIVE_EMPLOYEES
        _region_pms = {
            n for n, d in EMPLOYEE_LOCATION.items()
            if PS_REGION_MAP.get(PS_REGION_OVERRIDE.get(n, d), "") == _va_region
        }
        my_drs = df_drs[pm_col.apply(lambda v: str(v).strip() in _region_pms)].copy()
    elif role == "manager_only":
        my_drs = df_drs.copy()
    else:
        my_drs = df_drs[pm_col.apply(lambda v: name_matches(v, view_as))].copy()
else:
    my_drs = pd.DataFrame()

# ── Hero render ───────────────────────────────────────────────────────────────
_view_label = _va_region + " Team" if _va_region else (
    view_as if view_as != selected else selected
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
      Project Health</h1>
  <p style='color:rgba(255,255,255,0.45);margin:6px 0 0;font-size:14px;
            font-family:Manrope,sans-serif'>
      {_view_label} · {today.strftime("%A, %B %-d %Y")}</p>
</div>
""", unsafe_allow_html=True)

if df_drs is None:
    st.info("Load SS DRS on the Home page to view Project Health.")
    st.stop()

if my_drs.empty:
    st.info("No projects found for your profile in the DRS.")
    st.stop()

# ── Prepare data ──────────────────────────────────────────────────────────────
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

# Active only (exclude on-hold)
_ioh    = my_drs.get("_on_hold", pd.Series(False, index=my_drs.index)).astype(bool)
_active = my_drs[~_ioh].copy()
_legacy = _active.get("legacy", pd.Series(False, index=_active.index)).astype(bool)

# Normalise dates
for _dc in ["go_live_date","start_date"]:
    if _dc in _active.columns:
        _active[_dc] = pd.to_datetime(_active[_dc], errors="coerce")

# ── Schedule variance ─────────────────────────────────────────────────────────
def _schedule_status(row):
    gld = row.get("go_live_date")
    phase = str(row.get("phase","") or "").strip().lower()
    pi    = _pidx(phase)
    if pd.isna(gld): return "No date", None, "grey"
    days  = (pd.Timestamp(gld) - today).days
    # Already delivered (past go-live phase)
    if pi >= _pidx("06. go-live"):
        return "Delivered", 0, "green"
    if days > 14:   return "On track",   days, "green"
    if days >= 0:   return "Going live",  days, "blue"
    if days >= -30: return "Delayed",    days, "amber"
    return "At risk", days, "red"

_active[["_sched_status","_sched_days","_sched_col"]] = _active.apply(
    lambda r: pd.Series(_schedule_status(r)), axis=1
)

# ── Delivery summary metrics ──────────────────────────────────────────────────
_n_active      = len(_active)
_n_on_track    = len(_active[_active["_sched_status"].isin(["On track","Going live","Delivered"])])
_n_delayed     = len(_active[_active["_sched_status"] == "Delayed"])
_n_at_risk     = len(_active[_active["_sched_status"] == "At risk"])

# Avg schedule variance (negative = overdue) — only for projects with dates
_dated = _active[_active["_sched_days"].notna() & (_active["_sched_status"] != "Delivered")]
_avg_var = round(_dated["_sched_days"].mean()) if not _dated.empty else None

# Days to kickoff — start_date to ms_intro_email
_days_to_ko = []
if "start_date" in _active.columns and "ms_intro_email" in _active.columns:
    for _, r in _active.iterrows():
        sd = r.get("start_date"); ms = r.get("ms_intro_email")
        if pd.notna(sd) and pd.notna(ms):
            d = (pd.Timestamp(ms) - pd.Timestamp(sd)).days
            if 0 <= d <= 90: _days_to_ko.append(d)
_avg_kickoff = round(sum(_days_to_ko) / len(_days_to_ko)) if _days_to_ko else None

# Milestone completion rate
MS_COLS     = ["ms_intro_email","ms_config_start","ms_enablement","ms_session1",
               "ms_session2","ms_uat_signoff","ms_prod_cutover",
               "ms_hypercare_start","ms_close_out","ms_transition"]
MS_PHASES   = {
    "ms_intro_email":   0, "ms_config_start": 2, "ms_enablement": 3,
    "ms_session1":      3, "ms_session2":     3, "ms_uat_signoff": 4,
    "ms_prod_cutover":  5, "ms_hypercare_start": 6,
    "ms_close_out":     7, "ms_transition":   8,
}
_ms_expected = 0; _ms_recorded = 0
for _, r in _active[~_legacy].iterrows():
    pi = _pidx(str(r.get("phase","") or ""))
    for ms, ep in MS_PHASES.items():
        if pi > ep and ms in _active.columns:
            _ms_expected += 1
            if pd.notna(r.get(ms)): _ms_recorded += 1
_ms_rate = round(100 * _ms_recorded / _ms_expected) if _ms_expected > 0 else None

# ── SECTION: Delivery Summary ────────────────────────────────────────────────
st.markdown('<div class="section-label">Delivery Summary</div>', unsafe_allow_html=True)
dc1, dc2, dc3, dc4 = st.columns(4)

with dc1:
    _pct = round(100 * _n_on_track / _n_active) if _n_active else 0
    _col = "#27AE60" if _pct >= 70 else ("#F39C12" if _pct >= 50 else "#C0392B")
    st.markdown(f"""<div class="metric-card">
      <div class="metric-val" style="color:{_col}">{_n_on_track}</div>
      <div class="metric-lbl">On-schedule projects</div>
      <div class="metric-sub">{_pct}% of {_n_active} active</div>
    </div>""", unsafe_allow_html=True)

with dc2:
    if _avg_var is None:
        _var_txt, _var_col = "—", "inherit"
    elif _avg_var >= 0:
        _var_txt, _var_col = f"+{_avg_var}d ahead", "#27AE60"
    else:
        _var_txt, _var_col = f"{_avg_var}d", "#C0392B"
    st.markdown(f"""<div class="metric-card">
      <div class="metric-val" style="color:{_var_col}">{_var_txt}</div>
      <div class="metric-lbl">Avg schedule variance</div>
      <div class="metric-sub">vs planned go-live</div>
    </div>""", unsafe_allow_html=True)

with dc3:
    _ms_txt = f"{_ms_rate}%" if _ms_rate is not None else "—"
    _ms_col = "#27AE60" if (_ms_rate or 0) >= 80 else ("#F39C12" if (_ms_rate or 0) >= 60 else "#C0392B")
    st.markdown(f"""<div class="metric-card">
      <div class="metric-val" style="color:{_ms_col}">{_ms_txt}</div>
      <div class="metric-lbl">Milestone completion rate</div>
      <div class="metric-sub">{_ms_recorded} of {_ms_expected} expected recorded</div>
    </div>""", unsafe_allow_html=True)

with dc4:
    _ko_txt = f"{_avg_kickoff}d" if _avg_kickoff is not None else "—"
    st.markdown(f"""<div class="metric-card">
      <div class="metric-val">{_ko_txt}</div>
      <div class="metric-lbl">Avg days to kickoff</div>
      <div class="metric-sub">start date → intro email</div>
    </div>""", unsafe_allow_html=True)

# ── SECTION: Schedule Variance Table ─────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Schedule Variance by Project</div>', unsafe_allow_html=True)

if not _active.empty:
    _sched_rows = []
    for _, r in _active.sort_values("_sched_days", na_position="last").iterrows():
        _pn   = str(r.get("project_name","") or "")
        _cust = _pn.split(" - ")[0].strip() if " - " in _pn else _pn
        _pt   = str(r.get("project_type","") or "")
        _prod = _pt.split(":")[-1].strip().split()[0] if ":" in _pt else (_pt.split()[0] if _pt else "—")
        _ph   = str(r.get("phase","") or "—")
        _ph_lbl = _ph.split(".")[-1].strip() if "." in _ph else _ph

        _gld  = r.get("go_live_date")
        _gld_str = pd.Timestamp(_gld).strftime("%-d %b %Y") if pd.notna(_gld) else "—"
        _status = r.get("_sched_status","—")
        _days   = r.get("_sched_days")
        _scol   = r.get("_sched_col","grey")

        if _days is None or pd.isna(_days):
            _var_str = "—"
        elif _status == "Delivered":
            _var_str = "Delivered"
        elif _days >= 0:
            _var_str = f"+{int(_days)}d"
        else:
            _var_str = f"{int(_days)}d"

        _rag = str(r.get("rag","") or "").strip().lower()
        _rag_pill = (
            '<span class="pill pill-red">Red</span>' if _rag == "red" else
            '<span class="pill pill-amber">Yellow</span>' if _rag == "yellow" else
            '<span class="pill pill-green">Green</span>' if _rag == "green" else
            '<span class="pill pill-grey">—</span>'
        )
        _status_pill = (
            f'<span class="pill pill-{"green" if _scol=="green" else "amber" if _scol=="amber" else "red" if _scol=="red" else "blue" if _scol=="blue" else "grey"}">{_status}</span>'
        )
        _pm = str(r.get("project_manager","") or "")
        _pm_parts = [p.strip() for p in _pm.split(",")]
        _pm_short = f"{_pm_parts[1].split()[0]} {_pm_parts[0]}" if len(_pm_parts) == 2 else _pm

        _sched_rows.append({
            "Customer":      _cust[:30],
            "Product":       _prod,
            "Phase":         _ph_lbl[:22],
            "Go-Live":       _gld_str,
            "Variance":      _var_str,
            "RAG":           _rag_pill,
            "Status":        _status_pill,
            "Consultant":    _pm_short,
        })

    _sched_df = pd.DataFrame(_sched_rows)

    # Show table as HTML for coloured pills
    _cols_show = ["Customer","Product","Phase","Go-Live","Variance","RAG","Status"]
    if role in ("manager","manager_only","reporting_only") or _va_region:
        _cols_show.append("Consultant")

    _tbl_html = """<div style='border:0.5px solid rgba(128,128,128,.2);border-radius:10px;
                               overflow:hidden;margin-bottom:4px'>
    <table style='width:100%;border-collapse:collapse;font-family:Manrope,sans-serif;font-size:12px'>
    <thead><tr style='border-bottom:1px solid rgba(128,128,128,.2)'>"""
    for c in _cols_show:
        _tbl_html += f"<th style='padding:8px 12px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.55'>{c}</th>"
    _tbl_html += "</tr></thead><tbody>"
    for i, row in _sched_df[_cols_show].iterrows():
        _bg = "background:rgba(192,57,43,.04)" if row.get("Status","") and "risk" in row.get("Status","").lower() else ""
        _tbl_html += f"<tr style='border-bottom:0.5px solid rgba(128,128,128,.1);{_bg}'>"
        for c in _cols_show:
            _val = row[c]
            _tbl_html += f"<td style='padding:8px 12px;color:inherit'>{_val}</td>"
        _tbl_html += "</tr>"
    _tbl_html += "</tbody></table></div>"
    st.markdown(_tbl_html, unsafe_allow_html=True)
    st.caption(f"{len(_active)} active projects · excludes on-hold")

# ── SECTION: Milestone Health ─────────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Milestone Health</div>', unsafe_allow_html=True)

_MS_DISPLAY = {
    "ms_intro_email":     ("Intro email",          0),
    "ms_config_start":    ("Config start",         2),
    "ms_enablement":      ("Enablement session",   3),
    "ms_uat_signoff":     ("UAT signoff",           4),
    "ms_prod_cutover":    ("Prod cutover",          5),
    "ms_hypercare_start": ("Hypercare start",       6),
    "ms_close_out":       ("Close-out tasks",       7),
    "ms_transition":      ("Support transition",    8),
}

_ms_missing = {}
for ms_col, (ms_lbl, exp_phase) in _MS_DISPLAY.items():
    if ms_col not in _active.columns: continue
    _flagged = _active[
        ~_legacy &
        (_active["phase"].fillna("").apply(lambda p: _pidx(p) > exp_phase)) &
        (~_active[ms_col].notna())
    ]
    if len(_flagged) > 0:
        _ms_missing[ms_lbl] = len(_flagged)

if _ms_missing:
    _ms_cols = st.columns(min(len(_ms_missing), 4))
    for i, (lbl, cnt) in enumerate(_ms_missing.items()):
        _col_idx = i % 4
        _col_c = "#C0392B" if cnt >= 3 else ("#F39C12" if cnt >= 1 else "#27AE60")
        with _ms_cols[_col_idx]:
            st.markdown(f"""<div class="metric-card">
              <div class="metric-val" style="color:{_col_c}">{cnt}</div>
              <div class="metric-lbl">Missing: {lbl}</div>
              <div class="metric-sub">past expected phase</div>
            </div>""", unsafe_allow_html=True)
else:
    st.success("All expected milestones are recorded — no gaps detected.")

# ── SECTION: Scope Health ─────────────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Scope Health</div>', unsafe_allow_html=True)

_scope_rows = []
for _, r in _active.iterrows():
    _ptype = str(r.get("project_type","") or "")
    _pname = str(r.get("project_name","") or "")
    _actual = float(r.get("actual_hours",0) or 0)
    _budget = float(r.get("budgeted_hours",0) or 0)
    _co     = float(r.get("change_order",0) or 0)
    _total_scope = _budget + _co
    _ff_scope = get_ff_scope(_ptype, _pname)
    _scope = _ff_scope or (_total_scope if _total_scope > 0 else None)
    if _scope and _scope > 0 and _actual > 0:
        _burn = round(100 * _actual / _scope)
        _cust = _pname.split(" - ")[0].strip()[:28] if " - " in _pname else _pname[:28]
        _scope_rows.append({
            "Customer":    _cust,
            "Scope hrs":   int(_scope),
            "Used hrs":    round(_actual,1),
            "Burn %":      _burn,
            "_burn":       _burn,
        })

if _scope_rows:
    _scope_df = pd.DataFrame(_scope_rows).sort_values("_burn", ascending=False)
    sc1, sc2, sc3 = st.columns(3)
    _overrun = _scope_df[_scope_df["_burn"] > 100]
    _at_limit = _scope_df[(_scope_df["_burn"] >= 80) & (_scope_df["_burn"] <= 100)]
    _healthy  = _scope_df[_scope_df["_burn"] < 80]
    with sc1:
        st.markdown(f"""<div class="metric-card">
          <div class="metric-val" style="color:{'#C0392B' if len(_overrun)>0 else 'inherit'}">{len(_overrun)}</div>
          <div class="metric-lbl">Overrun</div>
          <div class="metric-sub">hours exceed scope</div>
        </div>""", unsafe_allow_html=True)
    with sc2:
        st.markdown(f"""<div class="metric-card">
          <div class="metric-val" style="color:{'#F39C12' if len(_at_limit)>0 else 'inherit'}">{len(_at_limit)}</div>
          <div class="metric-lbl">At or near limit</div>
          <div class="metric-sub">80–100% scope used</div>
        </div>""", unsafe_allow_html=True)
    with sc3:
        st.markdown(f"""<div class="metric-card">
          <div class="metric-val" style="color:#27AE60">{len(_healthy)}</div>
          <div class="metric-lbl">Within budget</div>
          <div class="metric-sub">below 80% scope used</div>
        </div>""", unsafe_allow_html=True)

    # Top 8 by burn %
    st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)
    _tbl2 = """<div style='border:0.5px solid rgba(128,128,128,.2);border-radius:10px;overflow:hidden'>
    <table style='width:100%;border-collapse:collapse;font-family:Manrope,sans-serif;font-size:12px'>
    <thead><tr style='border-bottom:1px solid rgba(128,128,128,.2)'>
    <th style='padding:8px 12px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.55'>Customer</th>
    <th style='padding:8px 12px;text-align:right;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.55'>Scope</th>
    <th style='padding:8px 12px;text-align:right;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.55'>Used</th>
    <th style='padding:8px 12px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;opacity:.55'>Burn %</th>
    </tr></thead><tbody>"""
    for _, row in _scope_df.head(8).iterrows():
        _b = row["_burn"]
        _bc = "#C0392B" if _b > 100 else ("#F39C12" if _b >= 80 else "#27AE60")
        _bar_w = min(_b, 120)
        _bar_html = f"""<div style='display:flex;align-items:center;gap:8px'>
          <div style='flex:1;max-width:120px;background:rgba(128,128,128,.1);
                      border-radius:4px;height:6px;overflow:hidden'>
            <div style='width:{_bar_w}%;height:6px;background:{_bc};border-radius:4px'></div>
          </div>
          <span style='color:{_bc};font-weight:700;font-size:11px'>{_b}%</span>
        </div>"""
        _tbl2 += f"""<tr style='border-bottom:0.5px solid rgba(128,128,128,.1)'>
          <td style='padding:8px 12px'>{row["Customer"]}</td>
          <td style='padding:8px 12px;text-align:right;opacity:.65'>{row["Scope hrs"]}h</td>
          <td style='padding:8px 12px;text-align:right;opacity:.65'>{row["Used hrs"]}h</td>
          <td style='padding:8px 12px'>{_bar_html}</td>
        </tr>"""
    _tbl2 += "</tbody></table></div>"
    st.markdown(_tbl2, unsafe_allow_html=True)
else:
    st.info("No scope data available — ensure budgeted hours are set in the DRS.")

st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div style="font-size:11px;opacity:.4;text-align:center;margin-top:4px">PS Projects & Tools · Internal use only · Data loaded this session only</div>', unsafe_allow_html=True)
