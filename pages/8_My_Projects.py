"""
PS Tools — My Projects
Per-consultant project working list: snapshot metrics, needs-action items,
active projects table, on-hold projects, and a change export flow.
"""
import streamlit as st
import pandas as pd
import io
from datetime import date, timedelta

st.session_state["current_page"] = "My Projects"

# ── Customer name extraction (shared logic with Customer Profile) ─────────────
_PC = ["ZEP","ZoneBilling","ZBilling","ZonePayroll","ZPayroll","ZoneCapture",
       "ZoneApprovals","ZoneReconcile","ZA","ZC","ZR","ZB","ZP"]
_PW = ["Payroll","Billing","Capture","Approvals","Reconcile","Implementation",
       "Optimization","Migration","Integration","Training","Support","MSA"]

def _extract_customer_name(project_name):
    import re as _re
    n = str(project_name).strip()
    # "Customer - XX - Description"
    m = _re.match(r'^(.+?)\s*-\s*[A-Z]{1,4}\s*-\s*.+$', n)
    if m: return m.group(1).strip()
    # "Customer- ProductCode ..."
    _pc_pat = '|'.join(_PC)
    m = _re.match(r'^(.+?)\s*-\s*(?:' + _pc_pat + r')(?:\s|$|-)', n, _re.IGNORECASE)
    if m: return m.group(1).strip()
    # "Customer- ProductWord ..."
    _pw_pat = '|'.join(_PW)
    m = _re.match(r'^(.+?)\s*-\s*(?:' + _pw_pat + r').+$', n, _re.IGNORECASE)
    if m: return m.group(1).strip()
    # "Customer : Customer" (duplicate separated by colon)
    m = _re.match(r'^(.+?)\s*:\s*\1\s*$', n, _re.IGNORECASE)
    if m: return m.group(1).strip()
    # "Customer : Something" (colon separator)
    if ' : ' in n: return n.split(' : ')[0].strip()
    # "Customer ProductCode" no dash
    for code in sorted(_PC, key=len, reverse=True):
        m = _re.search(r'\s+' + _re.escape(code) + r'(?:\s|$|-)', n, _re.IGNORECASE)
        if m and m.start() > 2: return n[:m.start()].strip()
    # "Customer ProductWord" no dash
    for word in _PW:
        m = _re.search(r'\s+' + _re.escape(word) + r'(?:\s|$)', n, _re.IGNORECASE)
        if m and m.start() > 3: return n[:m.start()].strip().rstrip('-').strip()
    return n

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN, ACTIVE_EMPLOYEES,
    MILESTONE_COLS_MAP, get_role, is_manager,
    get_ff_scope, resolve_name, name_matches,
)
from shared.config import (
    EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP,
)
from shared.loaders import calc_days_inactive, calc_last_milestone
from shared.template_utils import suggest_tier, TEMPLATES

st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
<style>
    html,body,[class*="css"]{font-family:'Manrope',sans-serif!important}
    .section-label { font-size: 13px;font-weight:700;text-transform:uppercase;
                   letter-spacing:.8px;color:#4472C4;margin-bottom:8px}
    .metric-card{border:1px solid rgba(128,128,128,.2);border-radius:8px;padding:16px 20px;margin-bottom:12px}
    .metric-val { font-size: 32px;font-weight:700}
    .metric-lbl { font-size: 14px;opacity:.6;margin-top:2px}
    .pf{display:inline-block;font-size:11px;font-weight:600;padding:2px 7px;
        border-radius:4px;margin-right:4px;margin-bottom:3px}
    .pf-e{background:rgba(192,57,43,0.15);color:#C0392B}
    .pf-w{background:rgba(243,156,18,.15);color:#D68910}
    .divider{border:none;border-top:1px solid rgba(128,128,128,.2);margin:20px 0}
    .eform{background:rgba(68,114,196,.06);border:1px solid rgba(68,114,196,.2);
           border-radius:8px;padding:14px 16px;margin-top:8px}
</style>
""", unsafe_allow_html=True)

# ── Identity ──────────────────────────────────────────────────────────────────
selected = st.session_state.get("consultant_name", "")
role     = get_role(selected) if selected else "consultant"
today    = pd.Timestamp.today().normalize()

PHASE_ORDER = [
    "00. onboarding","01. requirements and design","02. configuration",
    "03. enablement/training","04. uat","05. prep for go-live",
    "06. go-live","07. data migration","08. ready for support transition",
    "09. phase 2 scoping",
]
def _pidx(p):
    pl = str(p).strip().lower()
    for i,ph in enumerate(PHASE_ORDER):
        if pl.startswith(ph[:6]) or ph in pl or pl in ph: return i
    return -1

PHASE_OPTIONS = [
    "00. Onboarding","01. Requirements and Design","02. Configuration",
    "03. Enablement/Training","04. UAT","05. Prep for Go-Live",
    "06. Go-Live (Hypercare)","07. Data Migration",
    "08. Ready for Support Transition","09. Phase 2 Scoping",
    "10. Complete/Pending Final Billing",
]
SCHEDULE_OPTIONS = ["On Track","At Risk","Behind","Significantly Behind"]
MS_TO_SS = {
    "ms_intro_email":"Intro. Email Sent","ms_config_start":"Standard Config Start",
    "ms_enablement":"Enablement Session","ms_session1":"Session #1",
    "ms_session2":"Session #2","ms_uat_signoff":"UAT Signoff",
    "ms_prod_cutover":"Prod Cutover","ms_hypercare_start":"Hypercare Start",
    "ms_close_out":"Close Out Remaining Tasks","ms_transition":"Transition to Support",
}

# ── Sidebar View As (managers only) ──────────────────────────────────────────
view_as = selected
_va_region = None
if role == "manager":
    # Read from Home's View As dropdown — same widget, same position, no duplication
    _pick = st.session_state.get("home_browse", "— My own view —")
    if _pick and _pick.startswith("── ") and _pick.endswith(" ──"):
        _va_region = _pick[3:-3].strip()
    elif _pick in ("👥 All team", "All team"):
        _va_region = "__ALL__"   # special flag for all-team view
    elif _pick and _pick not in ("— My own view —", "— Select —", ""):
        view_as    = _pick
        _va_region = None

# ── Data ─────────────────────────────────────────────────────────────────────
df_drs = st.session_state.get("df_drs")
df_ns  = st.session_state.get("df_ns")
if df_drs is None:
    st.info("Upload SS DRS Export in the sidebar to see your projects.")
    st.stop()

# Filter to consultant or region
pm_col = df_drs.get("project_manager", pd.Series(dtype="object"))

if _va_region and role == "manager":
    if _va_region == "__ALL__":
        # All team — show every project
        my_drs = df_drs.copy()
    else:
        # Region view — build name set from ACTIVE_EMPLOYEES + PS_REGION_OVERRIDE aliases
        _region_consultants = set()
        # Add canonical names from ACTIVE_EMPLOYEES
        for n in ACTIVE_EMPLOYEES:
            _nl = EMPLOYEE_LOCATION.get(n, "")
            if isinstance(_nl, tuple): _nl = _nl[0]
            _nr = PS_REGION_OVERRIDE.get(n, PS_REGION_MAP.get(_nl, "Other"))
            if _nr == _va_region:
                _region_consultants.add(n.lower())
                _vp2 = [p.strip() for p in n.split(",")]
                _region_consultants.add(_vp2[0].lower())
                if len(_vp2) == 2:
                    _region_consultants.add(f"{_vp2[1].strip()} {_vp2[0]}".lower())
                    _region_consultants.add(_vp2[1].strip().lower())
        # Also add DRS display name aliases from PS_REGION_OVERRIDE (e.g. "Caroline Tuazon")
        for display_name, region in PS_REGION_OVERRIDE.items():
            if region == _va_region:
                _region_consultants.add(display_name.lower())
                _dparts = display_name.strip().split()
                if len(_dparts) >= 2:
                    _region_consultants.add(_dparts[-1].lower())  # last name only
        my_drs = df_drs[pm_col.apply(lambda v: str(v).strip().lower() in _region_consultants
                                     or resolve_name(str(v)).lower() in _region_consultants
                                     or any(str(v).strip().lower().startswith(ns + " ") or
                                            str(v).strip().lower().endswith(" " + ns)
                                            for ns in _region_consultants if len(ns) > 3)
                                     )].copy()
        if my_drs.empty:
            st.info(f"No projects found for the {_va_region} region in DRS.")
            st.stop()
elif view_as == selected and role == "manager_only":
    # Pure manager (no own projects) — prompt to use View As
    st.info(f"You are logged in as a manager. Use 'View as' to browse a consultant or region.")
    st.stop()
else:
    _vp = [p.strip() for p in view_as.split(",")]
    _vv = {view_as.lower(), _vp[0].lower()}
    if len(_vp) == 2: _vv.add(f"{_vp[1].strip()} {_vp[0]}".lower())
    my_drs = df_drs[pm_col.apply(lambda v: name_matches(v, view_as))].copy()
    if my_drs.empty:
        st.info(f"No projects found for {view_as} in DRS.")
        st.stop()

if df_ns is not None:
    try: my_drs = calc_days_inactive(my_drs,df_ns)
    except Exception: pass

_ioh   = my_drs.get("_on_hold",pd.Series(False,index=my_drs.index)).astype(bool)
on_hold= my_drs[_ioh].copy().reset_index(drop=True)
active = my_drs[~_ioh].copy().reset_index(drop=True)
# ── Deduped project counts for metrics (avoids multi-row DRS inflation) ──
_id_col_dc    = "project_id" if "project_id" in active.columns else "project_name"
_n_active_dc  = int(active[_id_col_dc].nunique()) if not active.empty else 0
_n_onhold_dc  = int(on_hold[_id_col_dc].nunique()) if not on_hold.empty else 0


# ── Flags ─────────────────────────────────────────────────────────────────────
def _flags(row):
    out=[]; phase=str(row.get("phase","")or"").strip()
    go_live=row.get("effective_go_live_date") or row.get("go_live_date"); start_dt=row.get("start_date")
    is_leg=bool(row.get("legacy",False)); pi=_pidx(phase)
    if pd.notna(go_live) and pi>=0:
        if pd.Timestamp(go_live)<today and pi<_pidx("06. go-live"):
            out.append(("error","phase","Go Live passed — advance Phase to 06. Go-Live",True))
    if pd.notna(go_live) and pd.notna(start_dt):
        if pd.Timestamp(go_live)<pd.Timestamp(start_dt):
            out.append(("error",None,"Go Live before Start date — raise with admin",False))
    if not is_leg and pi>=_pidx("00. onboarding") and pi>=0 and not pd.notna(row.get("ms_intro_email")):
        out.append(("warn","ms_intro_email","Intro email date not recorded",True))
    if not str(row.get("schedule_health","")or"").strip():
        out.append(("warn","schedule_health","Schedule Health not set",True))
    # Missing status or phase
    if not str(row.get("status","")or"").strip():
        out.append(("warn","status","Status not set",True))
    if not phase:
        out.append(("warn","phase","Phase not set",True))
    if not is_leg and pi>=0:
        for ms,ep in [("ms_uat_signoff",_pidx("04. uat")),
                       ("ms_prod_cutover",_pidx("05. prep for go-live")),
                       ("ms_hypercare_start",_pidx("06. go-live"))]:
            if pi>ep and not pd.notna(row.get(ms)):
                out.append(("warn",ms,f"Missing: {MILESTONE_COLS_MAP.get(ms,ms)}",True))
    return out

if not active.empty:
    active["_flags"]   = active.apply(_flags,axis=1)
    active["_ne"]      = active["_flags"].apply(lambda f:sum(1 for s,_,_m,_e in f if s=="error"))
    active["_nw"]      = active["_flags"].apply(lambda f:sum(1 for s,_,_m,_e in f if s=="warn"))
    active["_needs"]   = (active["_ne"]>0)|(active["_nw"]>0)|(active.get("days_inactive",pd.Series(-1,index=active.index))>=14)
else:
    for c in ["_flags","_ne","_nw","_needs"]: active[c]=None

# ── On Hold flags ──────────────────────────────────────────────────────────────
def _oh_flags(row):
    out = []
    days     = int(row.get("days_inactive", -1) or -1)
    resp     = str(row.get("client_responsiveness","") or "").strip().lower()
    sent     = str(row.get("client_sentiment","") or "").strip().lower()
    reason   = str(row.get("on_hold_reason","") or "").strip()
    delay_by = str(row.get("responsible_for_delay","") or "").strip()

    # Missing mandatory on-hold fields
    if not reason or reason in ("—","nan","None"):
        out.append(("warn","on_hold_reason","On Hold Reason not set — required for all on-hold projects",True))
    if not delay_by or delay_by in ("—","nan","None"):
        out.append(("warn","responsible_for_delay","Responsible for Delay not set — required for all on-hold projects",True))

    # Engagement/sentiment inconsistent with on-hold status
    if days >= 14:
        if resp in ("highly engaged","highly responsive","responsive"):
            out.append(("warn","client_responsiveness",
                f"Client Responsiveness '{row.get('client_responsiveness','')}' inconsistent with {days}d on hold",True))
        if sent in ("positive",):
            out.append(("warn","client_sentiment",
                f"Client Sentiment 'Positive' inconsistent with {days}d on hold",True))
    return out

if not on_hold.empty:
    on_hold["_flags"] = on_hold.apply(_oh_flags, axis=1)
else:
    on_hold["_flags"] = None

# ── Header ────────────────────────────────────────────────────────────────────
_dn = ("Global Team" if _va_region == "__ALL__" else _va_region + " Team" if _va_region
       else view_as.split(",")[1].strip()+" "+view_as.split(",")[0] if "," in view_as
       else view_as)

# ── Hero metrics ──────────────────────────────────────────────────────────────
_gl14_count = 0; _gl_next_name = ""
if not active.empty:
    _gl_col = next((c for c in ["effective_go_live_date","go_live_date"] if c in active.columns), None)
    if _gl_col:
        _gl_dates = pd.to_datetime(active[_gl_col], errors="coerce")
        _gl_mask  = (_gl_dates >= pd.Timestamp(today)) & (_gl_dates <= pd.Timestamp(today) + pd.Timedelta(days=14))
        _gl14_count = int(_gl_mask.sum())
        if _gl14_count > 0:
            _next_gl = active[_gl_mask].copy(); _next_gl["_gldt"] = _gl_dates[_gl_mask]
            _gl_next_name = _extract_customer_name(str(_next_gl.sort_values("_gldt").iloc[0].get("project_name","")))
_rag_at_risk = int(active["rag"].fillna("").str.strip().str.lower().isin(["red","yellow"]).sum()) if not active.empty and "rag" in active.columns else 0

_metric_tile = (
    "<div style='display:flex;gap:24px;margin-top:20px;flex-wrap:wrap'>"
    f"<div><div style='font-size:10px;font-weight:600;letter-spacing:1px;text-transform:uppercase;"
    f"color:rgba(255,255,255,.5);margin-bottom:4px'>Open Projects</div>"
    f"<div style='font-size:24px;font-weight:700;color:#fff;line-height:1'>{_n_active_dc}</div></div>"
    f"<div style='border-left:1px solid rgba(255,255,255,.12);padding-left:24px'>"
    f"<div style='font-size:10px;font-weight:600;letter-spacing:1px;text-transform:uppercase;"
    f"color:rgba(255,255,255,.5);margin-bottom:4px'>On Hold</div>"
    f"<div style='font-size:24px;font-weight:700;color:#fff;line-height:1'>{_n_onhold_dc}</div></div>"
    f"<div style='border-left:1px solid rgba(255,255,255,.12);padding-left:24px'>"
    f"<div style='font-size:10px;font-weight:600;letter-spacing:1px;text-transform:uppercase;"
    f"color:rgba(255,255,255,.5);margin-bottom:4px'>Go-Lives Next 14d</div>"
    f"<div style='font-size:24px;font-weight:700;color:{'#4ade80' if _gl14_count > 0 else '#fff'};line-height:1'>{_gl14_count}</div>"
    + (f"<div style='font-size:10px;color:rgba(255,255,255,.5);margin-top:2px'>Next: {_gl_next_name}</div>" if _gl_next_name else "") +
    f"</div>"
    f"<div style='border-left:1px solid rgba(255,255,255,.12);padding-left:24px'>"
    f"<div style='font-size:10px;font-weight:600;letter-spacing:1px;text-transform:uppercase;"
    f"color:rgba(255,255,255,.5);margin-bottom:4px'>RAG at Risk</div>"
    f"<div style='font-size:24px;font-weight:700;color:{'#f87171' if _rag_at_risk > 0 else '#fff'};line-height:1'>{_rag_at_risk}</div></div>"
    f"<div style='border-left:1px solid rgba(255,255,255,.12);padding-left:24px;opacity:.45'>"
    f"<div style='font-size:10px;font-weight:600;letter-spacing:1px;text-transform:uppercase;"
    f"color:rgba(255,255,255,.5);margin-bottom:4px'>Not Updated 14d</div>"
    f"<div style='font-size:24px;font-weight:700;color:#fff;line-height:1'>—</div>"
    f"<div style='font-size:10px;color:rgba(255,255,255,.4);margin-top:2px'>Requires hosting</div></div>"
    "</div>"
)

_hero = st.empty()
_hero.markdown(
    f"<div style='background:linear-gradient(135deg,#1a56db 0%,#050D1F 55%,#050D1F 100%);"
    f"padding:28px 32px 24px;border-radius:10px;margin-bottom:16px;font-family:Manrope,sans-serif'>"
    f"<div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;"
    f"color:#7dd3fc;margin-bottom:8px'>Professional Services · My Work</div>"
    f"<h1 style='color:#fff;margin:0;font-size:26px;font-weight:700;font-family:Manrope,sans-serif'>"
    f"My Projects — {_dn}</h1>"
    f"<p style='color:rgba(255,255,255,.45);margin:6px 0 0;font-size:13px;font-family:Manrope,sans-serif'>"
    f"{today.strftime('%A, %B %-d %Y')}</p>"
    f"{_metric_tile}"
    "</div>",
    unsafe_allow_html=True,
)
st.markdown('<hr class="divider">',unsafe_allow_html=True)

st.markdown("""
<div style='background:var(--color-background-secondary, rgba(59,158,255,0.05));
            border-left:4px solid #4472C4;border-radius:6px;
            padding:16px 20px;margin:0 0 20px;font-family:Manrope,sans-serif'>
    <div style='font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;
                color:#4472C4;margin-bottom:10px'>How to update your projects</div>
    <div style='display:flex;gap:32px;flex-wrap:wrap'>
        <div style='flex:1;min-width:220px;border-left:2px solid rgba(68,114,196,.4);padding-left:14px'>
            <span style='background:#1E2C63;color:#fff;font-size:10px;font-weight:700;
                         padding:2px 8px;border-radius:10px;letter-spacing:1px'>OPTION 1 &middot; QUICK UPDATES</span>
            <p style='margin:8px 0 0;font-size:13px;color:inherit;line-height:1.6'>
                <strong>Open projects &amp; On hold tabs</strong> &mdash; edit phase, status,
                milestone dates, and on-hold fields directly in the table. Sync the whole
                batch to Smartsheet in one click.
            </p>
        </div>
        <div style='flex:1;min-width:220px;border-left:2px solid rgba(245,158,11,.4);padding-left:14px'>
            <span style='background:rgba(245,158,11,0.15);color:#f59e0b;font-size:10px;font-weight:700;
                         padding:2px 8px;border-radius:10px;letter-spacing:1px;
                         border:1px solid rgba(245,158,11,0.4)'>OPTION 2 &middot; PROJECT DETAIL</span>
            <p style='margin:8px 0 0;font-size:13px;color:inherit;opacity:0.85;line-height:1.6'>
                <strong>Project detail tab</strong> &mdash; select any project to see its full
                DRS record on the left and all editable fields on the right. Save directly
                to Smartsheet.
            </p>
        </div>
        <div style='flex:1;min-width:220px;border-left:2px solid rgba(34,197,94,.4);padding-left:14px'>
            <span style='background:rgba(34,197,94,0.12);color:#22c55e;font-size:10px;font-weight:700;
                         padding:2px 8px;border-radius:10px;letter-spacing:1px;
                         border:1px solid rgba(34,197,94,0.35)'>BOTH OPTIONS</span>
            <p style='margin:8px 0 0;font-size:13px;color:inherit;opacity:0.85;line-height:1.6'>
                <strong>Sync directly to Smartsheet DRS</strong> &mdash; only the fields you
                change are written back. Nothing is overwritten by accident.
            </p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS — At a glance / Open Projects / On Hold / Project Detail
# ══════════════════════════════════════════════════════════════════════════════

# ── NS lookups — must be defined before overrun metric ───────────────────────
def _clean_pid(v):
    try:
        s = str(v).strip()
        if s in ("", "nan", "None"): return ""
        return str(int(float(s)))
    except: return str(v).strip()

_ns_htd: dict        = {}
_ns_tm_hrs: dict     = {}
_ns_tm_pids: set     = set()
_ns_period_hrs: dict = {}

if df_ns is not None and not active.empty:
    _ns_id_col = "project_id" if "project_id" in df_ns.columns else None
    if _ns_id_col and "hours_to_date" in df_ns.columns:
        for _pid, _grp in df_ns.groupby(_ns_id_col):
            _k = _clean_pid(_pid)
            if _k:
                try: _ns_htd[_k] = round(float(_grp["hours_to_date"].dropna().astype(float).max() or 0), 2)
                except: pass
    if _ns_id_col and "tm_scope" in df_ns.columns:
        for _pid, _grp in df_ns.groupby(_ns_id_col):
            _k = _clean_pid(_pid)
            if _k:
                try:
                    _v = _grp["tm_scope"].dropna().astype(float)
                    if not _v.empty:
                        _ns_tm_hrs[_k] = round(float(_v.max()), 2)
                        _ns_tm_pids.add(_k)
                except: pass
    if _ns_id_col and "billing_type" in df_ns.columns:
        _tm_ns = df_ns[df_ns["billing_type"].fillna("").str.strip().str.lower() == "t&m"]
        for _pid in _tm_ns[_ns_id_col].dropna().unique():
            _k = _clean_pid(_pid)
            if _k: _ns_tm_pids.add(_k)
    if _ns_id_col and "hours" in df_ns.columns:
        for _pid, _grp in df_ns.groupby(_ns_id_col):
            _k = _clean_pid(_pid)
            if _k:
                try: _ns_period_hrs[_k] = round(float(_grp["hours"].dropna().astype(float).sum() or 0), 2)
                except: pass

# ── At a glance metrics ───────────────────────────────────────────────────────
_n_flagged   = int((active["_ne"] > 0).sum() | (active["_nw"] > 0).sum()) if not active.empty and "_ne" in active.columns else 0
_n_flagged   = int(active["_flags"].apply(lambda f: bool(f)).sum()) if not active.empty and "_flags" in active.columns else 0

# FF overrun — count projects where NS hours-to-date exceed DEFAULT_SCOPE
# Uses same logic as Utilization Report: substring match on project_type
_n_overrun = 0
if not active.empty and _ns_htd:
    try:
        from shared.config import DEFAULT_SCOPE as _DS_mp
        def _scope_for(ptype):
            _pt = str(ptype or "").strip().lower()
            _best = None; _blen = 0
            for k,v in _DS_mp.items():
                if k.strip().lower() in _pt and len(k) > _blen:
                    _best = float(v); _blen = len(k)
            return _best
        def _is_overrun(r):
            _bt = str(r.get("billing_type","") or "").lower()
            if "t&m" in _bt or "time and material" in _bt: return False
            _pid = _clean_pid(str(r.get("project_id","") or ""))
            _scope = _scope_for(r.get("project_type",""))
            if _scope is None or _scope <= 0: return False
            _htd = _ns_htd.get(_pid, 0) or 0
            return float(_htd) > float(_scope)
        _n_overrun = int(active.apply(_is_overrun, axis=1).sum())
    except Exception:
        _n_overrun = 0

# No time entry 30d: days_inactive >= 30
_n_no_time_30 = 0
if not active.empty and "days_inactive" in active.columns:
    _n_no_time_30 = int((active["days_inactive"].fillna(-1) >= 30).sum())




tab_glance, tab_open, tab_hold, tab_intake = st.tabs([
    "At a glance",
    f"Open Projects · {_n_active_dc}",
    f"On Hold · {_n_onhold_dc}",
    "Project Detail",
])

# ═══════════════════════════════════════════════════════════════════
# TAB 1 — At a glance
# ═══════════════════════════════════════════════════════════════════
with tab_glance:
    # ── WHERE TO LOOK — matches Utilization Report callout style ────────────
    def _wtl_card(col, icon, title, num, body, is_red=False, is_amber=False):
        if num == 0:
            bc="#22c55e"; tc="#15803d"; ck="✓"
        elif is_red:
            bc="#ef4444"; tc="#b91c1c"; ck=icon
        else:
            bc="#f59e0b"; tc="#b45309"; ck=icon
        col.markdown(
            f"<div style='border:0.5px solid rgba(128,128,128,.15);border-radius:8px;"
            f"border-left:3px solid {bc};padding:14px 16px'>"
            f"<div style='display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px'>"
            f"<span style='font-size:12px;font-weight:600;color:{tc}'>{ck} {title}</span>"
            f"<span style='font-size:22px;font-weight:600;color:{tc};font-variant-numeric:tabular-nums'>{num}</span></div>"
            f"<div style='font-size:12px;opacity:.75;line-height:1.4;color:var(--color-text-primary)'>{body}</div>"
            "</div>", unsafe_allow_html=True
        )

    st.markdown("<div style='margin-top:4px;font-size:11px;opacity:.6;text-transform:uppercase;letter-spacing:.6px;margin-bottom:8px'>Where to look</div>", unsafe_allow_html=True)
    _wtl1,_wtl2,_wtl3 = st.columns(3)
    _wtl_card(_wtl1,"⚠","Projects with flags",_n_flagged,
        f"Date issues · missing milestones · phase gaps{'  · '+str(int(_n_flagged/_n_active_dc*100))+'% of portfolio' if _n_active_dc else ''}",
        is_red=_n_flagged>0)
    _wtl_card(_wtl2,"↑","FF overrun",_n_overrun,
        "Hours exceed contracted scope." if _n_overrun==0 else f"{_n_overrun} project{'s' if _n_overrun>1 else ''} over scope — check Hours to Date",
        is_amber=_n_overrun>0)
    _wtl_card(_wtl3,"○","No time entry 30d",_n_no_time_30,
        "All projects have recent time entries." if _n_no_time_30==0 else f"{_n_no_time_30} project{'s' if _n_no_time_30>1 else ''} with no NS time entry in 30 days",
        is_amber=_n_no_time_30>0)


    st.markdown('<hr class="divider">',unsafe_allow_html=True)

    # ── Projects at risk — identical to Utilization Report ────────────────
    if not active.empty:
        try:
            from shared.config import DEFAULT_SCOPE as _DS_at
            def _at_scope(ptype):
                _pt = str(ptype or "").strip().lower()
                _best = None; _blen = 0
                for k, v in _DS_at.items():
                    if k.strip().lower() in _pt and len(k) > _blen:
                        _best = float(v); _blen = len(k)
                return _best
        except Exception:
            def _at_scope(ptype): return None

        _risk_rows = []
        for _, _r in active.iterrows():
            _pid_k    = _clean_pid(str(_r.get("project_id", "") or ""))
            _scope    = _at_scope(_r.get("project_type", ""))
            _htd_v    = float(_ns_htd.get(_pid_k, 0) or 0)
            _logged   = float(_ns_period_hrs.get(_pid_k, 0) or 0)
            _burn     = (_htd_v / _scope) if (_scope and _scope > 0) else None
            _overrun  = max(0.0, round(_htd_v - _scope, 2)) if (_scope and _scope > 0) else 0.0
            _rag_v    = str(_r.get("rag", "") or "").strip().lower()
            _is_over  = _overrun > 0
            _is_burn  = _burn is not None and _burn >= 0.80 and not _is_over
            _is_noscope = (
                _scope is None and
                str(_r.get("billing_type", "") or "").strip().lower()
                not in ("t&m", "time and material", "internal")
            )
            if not (_is_over or _is_burn or _is_noscope or _rag_v in ("red", "yellow")):
                continue
            if _is_noscope:         _stag = ("blue",  "No scope")
            elif _is_over:          _stag = ("red",   "Overrun")
            elif _is_burn:          _stag = ("amber", f"Burn {int(_burn*100)}%")
            elif _rag_v == "red":   _stag = ("red",   "Red RAG")
            else:                   _stag = ("amber", "Amber RAG")
            _srank = 1 if _stag[0] == "red" else 2 if _stag[0] == "amber" else 3
            _risk_rows.append({
                "project":      _extract_customer_name(str(_r.get("project_name", ""))),
                "project_type": str(_r.get("project_type", "") or ""),
                "consultant":   str(_r.get("project_manager", "") or ""),
                "scoped_hrs":   _scope,
                "htd_hrs":      _htd_v,
                "hours_logged": _logged,
                "burn_pct":     _burn,
                "overrun_hrs":  _overrun,
                "_stag":        _stag,
                "_srank":       _srank,
            })

        if not _risk_rows:
            st.markdown(
                "<div style=\"border:1px solid rgba(128,128,128,.25);border-radius:8px;"
                "padding:32px;text-align:center\">"
                "<div style=\"font-size:28px;margin-bottom:8px\">✓</div>"
                "<div style=\"font-weight:600;margin-bottom:4px\">No projects at risk</div>"
                "<div style=\"opacity:.7;font-size:13px\">No FF overruns, no scope burn over 80%,"
                " no missing scope records.</div>"
                "</div>", unsafe_allow_html=True)
        else:
            import pandas as _rpd
            _risk_df = _rpd.DataFrame(_risk_rows)

            _mp_sort_opts = {
                "Status, then overrun": ("_srank",      "overrun_hrs", "burn_pct"),
                "Overrun (high → low)":  ("overrun_hrs",  "burn_pct",    None),
                "Burn % (high → low)":   ("burn_pct",     "overrun_hrs", None),
                "HTD (high → low)":      ("htd_hrs",      "overrun_hrs", None),
                "Logged this period":   ("hours_logged", "overrun_hrs", None),
                "Project (A → Z)": ("project_asc",  None,          None),
            }
            _mp_sc1, _ = st.columns([1.6, 4])
            with _mp_sc1:
                _mp_sc = st.selectbox("Sort by", list(_mp_sort_opts.keys()),
                    key="mp_risk_sort", label_visibility="visible")
            if _mp_sc == "Project (A → Z)":
                _risk_df = _risk_df.sort_values("project", ascending=True, na_position="last")
            else:
                _mk = [k for k in _mp_sort_opts[_mp_sc] if k]
                _ma = [True if k == "_srank" else False for k in _mk]
                _risk_df = _risk_df.sort_values(_mk, ascending=_ma, na_position="last")
            _risk_df = _risk_df.reset_index(drop=True)

            def _mp_avatar(name):
                parts = [p.strip() for p in str(name).replace(",", " ").split() if p.strip()]
                ini = (parts[0][0]+parts[1][0]).upper() if len(parts)>=2 else (parts[0][:2].upper() if parts else "??")
                pals = [("rgba(34,197,94,0.18)","#15803d"),("rgba(59,130,246,0.18)","#1d4ed8"),
                        ("rgba(245,158,11,0.18)","#b45309"),("rgba(168,85,247,0.18)","#7c3aed"),
                        ("rgba(236,72,153,0.18)","#be185d"),("rgba(20,184,166,0.18)","#0f766e")]
                bg, fg = pals[hash(str(name)) % len(pals)]
                return (f"<span style=\"display:inline-flex;align-items:center;justify-content:center;"
                        f"width:24px;height:24px;border-radius:50%;font-size:10px;font-weight:600;"
                        f"margin-right:8px;background:{bg};color:{fg}\">{ini}</span>")

            def _mp_short(n):
                p = str(n).split(",", 1)
                if len(p) == 2:
                    last = p[0].strip()
                    first = p[1].strip().split()[0] if p[1].strip() else ""
                    return f"{last}, {first[:1]}." if first else last
                return n

            _tbl_rows = []
            for _, _row in _risk_df.iterrows():
                _scls, _sl = _row["_stag"]
                _sc_str  = f"{_row['scoped_hrs']:,.2f}" if _row["scoped_hrs"] is not None else "—"
                _htd_str = f"{_row['htd_hrs']:,.2f}"    if _row["htd_hrs"]   else "—"
                _log_str = f"{_row['hours_logged']:,.2f}" if _row["hours_logged"] else "—"
                _bp = _row["burn_pct"]
                try: _bp = float(_bp) if _bp is not None else None
                except: _bp = None
                if _bp is not None and not (isinstance(_bp, float) and (_bp != _bp)):
                    _bpv  = min(_bp, 1.5)
                    _bcol = "#ef4444" if _bpv > 1.0 else "#f59e0b" if _bpv >= 0.8 else "#22c55e"
                    _bar = (f"<div style=\"display:inline-block;width:60px;height:6px;"
                            f"background:rgba(128,128,128,.2);border-radius:3px;overflow:hidden;"
                            f"vertical-align:middle;margin-right:6px\">"
                            f"<div style=\"width:{min(_bpv*100/1.5,100):.0f}%;height:100%;"
                            f"background:{_bcol}\"></div></div>")
                    _bs = f"{int(_bp*100)}%"
                else:
                    _bar = ""; _bs = "—"
                _cn = _row["consultant"]
                _ch = (f"<span style=\"display:inline-flex;align-items:center\">"
                       f"{_mp_avatar(_cn)}{_mp_short(_cn)}</span>") if _cn else "<span style=\"opacity:.4\">—</span>"
                _pbg = ("rgba(239,68,68,.18)" if _scls=="red"
                        else "rgba(245,158,11,.18)" if _scls=="amber"
                        else "rgba(59,130,246,.18)")
                _pfg = "#b91c1c" if _scls=="red" else "#b45309" if _scls=="amber" else "#1d4ed8"
                _tbl_rows.append(
                    f"<tr style=\"border-bottom:1px solid rgba(128,128,128,.15)\">"
                    f"<td style=\"padding:9px 12px;vertical-align:middle\">{_row['project']}</td>"
                    f"<td style=\"padding:9px 12px;vertical-align:middle;opacity:.75\">{_row['project_type']}</td>"
                    f"<td style=\"padding:9px 12px;vertical-align:middle\">{_ch}</td>"
                    f"<td style=\"padding:9px 12px;vertical-align:middle;text-align:right\">{_sc_str}</td>"
                    f"<td style=\"padding:9px 12px;vertical-align:middle;text-align:right\">{_htd_str}</td>"
                    f"<td style=\"padding:9px 12px;vertical-align:middle;text-align:right\">{_log_str}</td>"
                    f"<td style=\"padding:9px 12px;vertical-align:middle;text-align:right\">"
                    f"{_bar}<span style=\"vertical-align:middle\">{_bs}</span></td>"
                    f"<td style=\"padding:9px 12px;vertical-align:middle;text-align:right\">"
                    f"{_row['overrun_hrs']:,.2f}</td>"
                    f"<td style=\"padding:9px 12px;vertical-align:middle;text-align:center\">"
                    f"<span style=\"padding:3px 9px;border-radius:999px;font-size:12px;"
                    f"font-weight:500;background:{_pbg};color:{_pfg}\">{_sl}</span></td>"
                    f"</tr>"
                )

            _n_risk = len(_risk_df)
            st.markdown(
                f"<div style=\"border:1px solid rgba(128,128,128,.25);border-radius:8px 8px 0 0;"
                f"padding:10px 14px;display:flex;justify-content:space-between;font-size:12px\">"
                f"<span style=\"font-weight:600\">Projects requiring attention</span>"
                f"<span style=\"opacity:.7\">{_n_risk} project{{'s' if _n_risk!=1 else ''}} · sorted by {_mp_sc.lower()}</span>"
                f"</div>"
                f"<div style=\"border:1px solid rgba(128,128,128,.25);border-top:none;"
                f"border-radius:0 0 8px 8px;overflow:hidden\">"
                f"<table style=\"width:100%;border-collapse:collapse;font-family:Manrope,sans-serif;"
                f"font-size:13px;font-variant-numeric:tabular-nums;color:inherit\">"
                f"<thead><tr style=\"background:rgba(128,128,128,.08);"
                f"border-bottom:1px solid rgba(128,128,128,.25)\">"
                f"<th style=\"padding:10px 12px;font-weight:600;text-align:left;opacity:.75\">Project</th>"
                f"<th style=\"padding:10px 12px;font-weight:600;text-align:left;opacity:.75\">Type</th>"
                f"<th style=\"padding:10px 12px;font-weight:600;text-align:left;opacity:.75\">Consultant</th>"
                f"<th style=\"padding:10px 12px;font-weight:600;text-align:right;opacity:.75\">Scoped</th>"
                f"<th style=\"padding:10px 12px;font-weight:600;text-align:right;opacity:.75\">HTD</th>"
                f"<th style=\"padding:10px 12px;font-weight:600;text-align:right;opacity:.75\">Logged</th>"
                f"<th style=\"padding:10px 12px;font-weight:600;text-align:right;opacity:.75\">Burn</th>"
                f"<th style=\"padding:10px 12px;font-weight:600;text-align:right;opacity:.75\">Overrun</th>"
                f"<th style=\"padding:10px 12px;font-weight:600;text-align:center;opacity:.75\">Status</th>"
                f"</tr></thead>"
                f"<tbody>{''.join(_tbl_rows)}</tbody></table>"
                f"<div style=\"padding:8px 14px;border-top:1px solid rgba(128,128,128,.15);"
                f"display:flex;justify-content:space-between;font-size:11px;opacity:.6\">"
                f"<span>Risk = overrun &gt; 0, burn ≥ 80% (HTD ÷ scope),"
                f" Red/Amber RAG, or no scope record. HTD = hours-to-date all time.</span>"
                f"<span>Source: NS + DRS · this session</span>"
                f"</div></div>",
                unsafe_allow_html=True
            )

# ═══════════════════════════════════════════════════════════════════
# TAB 2 — Open Projects
# ═══════════════════════════════════════════════════════════════════
with tab_open:


    # ══════════════════════════════════════════════════════════════════════════════
    # SECTION 1 — Snapshot
    # ══════════════════════════════════════════════════════════════════════════════
    # ══════════════════════════════════════════════════════════════════════════════
    # SECTION 2+3 — Active Projects (editable table + export)
    # ══════════════════════════════════════════════════════════════════════════════
    st.markdown('<div class="section-label">Open Projects</div>',unsafe_allow_html=True)

    if active.empty:
        st.info("No active projects found.")
    else:
        pass  # NS lookups now at module level above tabs
        # ── Build editable dataframe ──────────────────────────────────────────────
    def _rag_emoji(val):
        v = str(val or "").strip().lower()
        if v == "red":    return "🔴"
        if v == "yellow": return "🟡"
        if v == "green":  return "🟢"
        return "—"

    def _engagement_flag(row):
        flags = []
        _days  = int(row.get("days_inactive", -1) or -1)
        _leg   = str(row.get("legacy","")).strip().lower() in ("true","yes","1")
        _phase = str(row.get("phase","") or "").strip().lower()
        _htd   = row.get("_htd_val", 0) or 0  # hours to date — if >0 work has started
        _start = row.get("start_date")

        # Determine if project is genuinely new/early enough to flag missing intro:
        # Skip flag if: legacy, has hours logged (work started = intro likely happened),
        # or phase is past onboarding
        _past_onboarding = any(p in _phase for p in ["config","enablement","training",
                               "uat","prep","go-live","hypercare","support","transition","ready"])
        _has_hours = float(_htd) > 0 if _htd != "" else False

        _no_i = (not pd.notna(row.get("ms_intro_email")) or
                 str(row.get("ms_intro_email","")).strip() in ("","nan","None","NaT"))

        if not _leg and _no_i and not _past_onboarding and not _has_hours:
            flags.append("No intro")

        if _days >= 30:
            flags.append(f"{_days}d inactive")
        elif _days >= 14:
            flags.append(f"{_days}d inactive")
        return " · ".join(flags) if flags else "✓"

    def _to_edit_row(row):
        fl    = row.get("_flags",[]) or []
        needs = "⚠️" if any(sev=="error" for sev,_,_m,_ in fl) else ("⚠️" if any(sev=="warn" for sev,_,_m,_ in fl) else "")
        def _ms(col):
            v = row.get(col)
            return pd.Timestamp(v).date() if pd.notna(v) else None


        def _dt(col):
            v = row.get(col)
            return pd.Timestamp(v).strftime("%Y-%m-%d") if pd.notna(v) else ""
        _pn    = str(row.get("project_name","") or "")
        _cust  = _extract_customer_name(_pn)
        # Scope: FF → from DEFAULT_SCOPE table by project_type; T&M → total NS hours logged
        _ptype_raw = str(row.get("project_type", "") or "")
        _bill_raw  = str(row.get("billing_type", "") or "").strip().lower()
        _pid_key   = _clean_pid(row.get("project_id", ""))
        _is_tm     = ("t&m" in _bill_raw or "time" in _bill_raw
                      or _pid_key in _ns_tm_pids)  # confirmed T&M from NS
        _ff_scope  = get_ff_scope(_ptype_raw, _pn)
        if _is_tm:
            # T&M scope from NS Time Detail "T&M Scope" column (max per project_id)
            # Falls back to hours sum if column not present in export
            _scope = round(_ns_tm_hrs.get(_pid_key, 0.0), 2) if _pid_key and _pid_key in _ns_tm_hrs else ""
        elif _ff_scope is not None:
            _scope = float(_ff_scope)
        else:
            _scope = ""
        _htd = round(_ns_htd.get(_pid_key, 0.0), 2) if _pid_key and _pid_key in _ns_htd else ""
        return {
            "Flags":                needs,
            "Customer":             _cust,
            "Consultant":           str(row.get("project_manager","") or ""),
            "Project Type":         str(row.get("project_type","") or ""),
            "Status":               str(row.get("status","") or ""),
            "Phase":                str(row.get("phase","") or ""),
            "Start Date":           _ms("start_date"),
            "Go-Live Date":         _ms("go_live_date"),
            "Intro Email Sent":     _ms("ms_intro_email"),
            "Config Start":         _ms("ms_config_start"),
            "Enablement Session":   _ms("ms_enablement"),
            "Session #1":           _ms("ms_session1"),
            "Session #2":           _ms("ms_session2"),
            "UAT Signoff":          _ms("ms_uat_signoff"),
            "Prod Cutover":         _ms("ms_prod_cutover"),
            "Go-Live (Actual)":     _ms("ms_prod_cutover"),
            "Hypercare Start":      _ms("ms_hypercare_start"),
            "Close Out Tasks":      _ms("ms_close_out"),
            "Transition to Support":_ms("ms_transition"),
        }

    try:
        edit_df = pd.DataFrame([_to_edit_row(r) for _,r in active.iterrows()])
    except Exception as _e:
        st.error(f"Table build error: {_e}")
        st.stop()
    # Hide Consultant column for single-person views
    if not _va_region:
        edit_df = edit_df.drop(columns=["Consultant"], errors="ignore")

    # ── Column config ─────────────────────────────────────────────────────────
    _ms_cols = ["Intro Email Sent","Config Start","Enablement Session","Session #1","Session #2",
                "UAT Signoff","Prod Cutover","Go-Live (Actual)","Hypercare Start","Close Out Tasks","Transition to Support"]
    col_cfg = {
        "Flags":                 st.column_config.TextColumn("Flags",             disabled=True, width="small"),
        "Customer":              st.column_config.TextColumn("Customer",          disabled=True),
        "Consultant":            st.column_config.TextColumn("Consultant",        disabled=True),
        "Project Type":          st.column_config.TextColumn("Project Type",      disabled=True),
        "Status":                st.column_config.SelectboxColumn("Status", options=["In Progress","On Hold","Closed","Complete","Cancelled"], width="medium"),
        "Phase":                 st.column_config.SelectboxColumn("Phase",        options=PHASE_OPTIONS, width="medium"),
        "Start Date":            st.column_config.DateColumn("Start Date",        min_value=date(2020,1,1), max_value=date(2030,12,31)),
        "Go-Live Date":          st.column_config.DateColumn("Go-Live Date",      min_value=date(2020,1,1), max_value=date(2030,12,31)),
        **{c: st.column_config.DateColumn(c, min_value=date(2020,1,1), max_value=date(2030,12,31)) for c in _ms_cols},
    }

    st.caption("Columns with the edit icon sync back to Smartsheet — edit and export to CSV to update DRS. Greyed columns are derived or read-only.")
    st.markdown('<span style="font-size:11.5px;opacity:.6">⚠️ Flags indicate date issues, missing milestones, or phase gaps. For a deeper look at data quality issues, use the DRS Health Check page.</span>', unsafe_allow_html=True)
    _btn_col1, _btn_col2 = st.columns([1, 1])
    with _btn_col1:
        if False:  # removed
            if _va_region:
                st.session_state["_va_passthrough"]    = f"── {_va_region} ──"
                st.session_state["_browse_passthrough"] = f"── {_va_region} ──"
            elif view_as and view_as != selected:
                st.session_state["_va_passthrough"]    = view_as
                st.session_state["_browse_passthrough"] = view_as
            st.switch_page("pages/6_DRS_Health_Check.py")
    with _btn_col2:
        if False:  # removed
            if _va_region:
                st.session_state["_browse_passthrough"] = f"── {_va_region} ──"
            elif view_as and view_as != selected:
                st.session_state["_browse_passthrough"] = view_as
            st.switch_page("pages/2_Customer_Reengagement.py")

    _display_df = edit_df.copy()
    edited = st.data_editor(
        _display_df,
        column_config=col_cfg,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="mp_edit_table",
    )

    # ── Detect changes vs original ────────────────────────────────────────────
    editable_cols = ["Start Date","Go-Live Date","Phase","Intro Email Sent","Config Start","Enablement Session","Session #1","Session #2","UAT Signoff","Prod Cutover","Go-Live (Actual)","Hypercare Start","Close Out Tasks","Transition to Support"]
    changed = edited[editable_cols].fillna("").ne(edit_df[editable_cols].fillna("")).any(axis=1)
    changed_df = edited[changed].copy() if changed.any() else pd.DataFrame()

    # ── Export bar ────────────────────────────────────────────────────────────
    from shared.smartsheet_api import ss_available, write_row_updates, WRITEBACK_FIELDS

    _ss_ready       = ss_available()
    _loaded_via_api = st.session_state.get("_drs_source") == "api"
    _has_row_ids    = "_ss_row_id" in active.columns

    # Status label
    ex1, ex2, ex3 = st.columns([3, 1, 1])
    with ex1:
        if changed.any():
            st.markdown(f'<span style="font-size:13px;color:#27AE60;font-weight:600">✓ {changed.sum()} project(s) edited — ready to export</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span style="font-size:12px;opacity:.5">No edits yet — edit cells above then export or sync</span>', unsafe_allow_html=True)

    # CSV export (always available)
    with ex2:
        _export_df = changed_df if not changed_df.empty else edited
        _buf = io.BytesIO()
        _export_df.to_csv(_buf, index=False)
        st.download_button(
            label="⬇ Export to CSV" if not changed_df.empty else "⬇ Export all",
            data=_buf.getvalue(),
            file_name=f"drs_updates_{date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            type="primary" if (not changed_df.empty and not (_ss_ready and _has_row_ids)) else "secondary",
            use_container_width=True,
        )

    # Smartsheet sync (only when DRS loaded via API — row IDs available)
    with ex3:
        if _ss_ready and _has_row_ids:
            _sync_disabled = not changed.any()
            if st.button(
                "↑ Sync to Smartsheet",
                key="mp_ss_sync",
                disabled=_sync_disabled,
                type="primary" if changed.any() else "secondary",
                use_container_width=True,
                help="Write edited fields directly back to the Smartsheet DRS" if not _sync_disabled
                     else "No edits to sync — edit cells above first",
            ):
                # Build update payload — map display column names back to internal keys
                _display_to_internal = {v: k for k, v in WRITEBACK_FIELDS.items()}
                _editable_display_cols = list(WRITEBACK_FIELDS.values())

                _changed_positions  = [i for i, v in enumerate(changed) if v]
                _sync_rows          = active.iloc[_changed_positions].copy()
                _edited_changed     = edited.iloc[_changed_positions].copy()
                _orig_changed       = edit_df.iloc[_changed_positions].copy()

                updates = []
                for _ci in range(len(_sync_rows)):
                    row_id    = _sync_rows.iloc[_ci]["_ss_row_id"]
                    proj_name = _sync_rows.iloc[_ci].get("project_name", str(row_id))

                    changes = {}
                    for disp_col in _editable_display_cols:
                        if disp_col not in _edited_changed.columns:
                            continue
                        internal_key = _display_to_internal.get(disp_col)
                        if not internal_key:
                            continue
                        new_val  = _edited_changed.iloc[_ci][disp_col]
                        orig_val = _orig_changed.iloc[_ci][disp_col] if disp_col in _orig_changed.columns else None
                        if str(new_val) != str(orig_val):
                            import datetime as _dt_mod
                            if isinstance(new_val, (_dt_mod.date, _dt_mod.datetime)):
                                new_val = new_val.isoformat()
                            changes[internal_key] = new_val

                    if changes:
                        updates.append({
                            "_ss_row_id":   row_id,
                            "project_name": proj_name,
                            "changes":      changes,
                        })

                if updates:
                    with st.spinner(f"Syncing {len(updates)} row(s) to Smartsheet…"):
                        _ok, _errs = write_row_updates(updates)
                    if _ok:
                        st.success(f"✓ {_ok} row(s) synced to Smartsheet.")
                    if _errs:
                        for _e in _errs:
                            st.warning(f"⚠ {_e}")
                else:
                    st.info("No writable field changes detected.")
        elif _ss_ready and not _has_row_ids:
            st.button(
                "↑ Sync to Smartsheet",
                key="mp_ss_sync_disabled",
                disabled=True,
                use_container_width=True,
                help="Reload DRS using 'Load from Smartsheet' on the Home page to enable direct sync",
            )

    # ── Re-engagement shortcuts for inactive projects ─────────────────────────
    _inactive_projs = active[active["days_inactive"].fillna(0)>=14].sort_values("days_inactive", ascending=False)
    if not _inactive_projs.empty:
        pass  # surfaced in At a glance tab

# ── On Hold helpers (defined before tab_hold block) ──────────────────────────
def _clean(val):
    if val is None: return "—"
    s = str(val).strip()
    return s if s and s not in ("nan","None","NaT","") else "—"

def _risk_emoji(val):
    v = str(val or "").strip().lower()
    if v in ("high","critical"): return "🔴 " + v.capitalize()
    if v == "medium": return "🟡 Medium"
    if v == "low":    return "🟢 Low"
    return _clean(val)

def _delay_summary_prompt(r):
    reason = _clean(r.get("on_hold_reason"))
    days   = int(r.get("days_inactive", 0) or 0)
    phase  = _clean(r.get("phase"))
    resp   = _clean(r.get("responsible_for_delay"))
    parts  = []
    if reason != "—": parts.append(f"Reason: {reason}")
    if days > 0:      parts.append(f"{days}d inactive")
    if phase != "—":  parts.append(f"Phase: {phase}")
    if resp != "—":   parts.append(f"Responsible: {resp}")
    return " · ".join(parts) if parts else "—"

_OH_REASON_OPTS    = ["None","Customer delay","Internal delay","Technical blocker","Commercial","Other"]
_OH_RESP_OPTS      = ["None","Customer","Internal","Shared"]
_OH_SENTIMENT_OPTS = ["None","Good","Neutral","Concern","At Risk"]
_OH_RISK_OPTS      = ["None","Low","Medium","High","Critical"]
_OH_OWNER_OPTS     = ["None","Consultant","Customer","Management","Partner"]
_OH_DELAY_OPTS     = ["None","Customer","Zone","Shared","Partner"]

# ═══════════════════════════════════════════════════════════════
# TAB 3 — On Hold
# ═══════════════════════════════════════════════════════════════
with tab_hold:
    # SECTION 4 — On Hold
    # ══════════════════════════════════════════════════════════════════════════════
    st.markdown('<div class="section-label">On Hold</div>', unsafe_allow_html=True)
    if on_hold.empty:
        st.markdown("No on-hold projects.")
    else:
        _oh_rows = []
        for _, r in on_hold.iterrows():
            _ds_existing = str(r.get("delay_summary","") or "").strip()
            _ds = _ds_existing if _ds_existing and _ds_existing not in ("—","nan","None") else _delay_summary_prompt(r)
            _rag_v = _rag_emoji(r.get("rag"))
            _oh_row = {
                    "Flags":                 ("⚠️" if any(s=="error" for s,*_ in (r.get("_flags") or []))
                                      else ("⚠️" if (r.get("_flags") or []) else "")),
                "Customer":              _extract_customer_name(str(r.get("project_name",""))),
                "Project Type":          str(r.get("project_type","") or "—"),
                "Start Date":            pd.Timestamp(r["start_date"]).strftime("%Y-%m-%d") if pd.notna(r.get("start_date")) else "—",
                "Est. Go-Live":          pd.Timestamp(r["go_live_date"]).strftime("%Y-%m-%d") if pd.notna(r.get("go_live_date")) else "—",
                "Phase":                 str(r.get("phase", "—")),
                "On Hold Reason":        _clean(r.get("on_hold_reason")) if _clean(r.get("on_hold_reason")) != "—" else None,
                "Last Milestone":        _clean(r.get("last_milestone")),
                "Client Responsiveness": _clean(r.get("client_responsiveness")) if _clean(r.get("client_responsiveness")) != "—" else None,
                "Client Sentiment":      _clean(r.get("client_sentiment")) if _clean(r.get("client_sentiment")) != "—" else None,
                "Risk Level":            _risk_emoji(r.get("risk_level")),
                "Risk Owner":            _clean(r.get("risk_owner")) if _clean(r.get("risk_owner")) != "—" else None,
                "Responsible for Delay": _clean(r.get("responsible_for_delay")) if _clean(r.get("responsible_for_delay")) != "—" else None,
            }
            if _va_region:
                _oh_row["Consultant"] = str(r.get("project_manager", "") or "")
            _oh_rows.append(_oh_row)

        _oh_df = pd.DataFrame(_oh_rows)

        # Column order — insert Consultant after RAG if region view
        if "Consultant" in _oh_df.columns:
            _col_order = ["Flags","Customer","Consultant","Project Type","Start Date","Est. Go-Live",
                           "Phase","Last Milestone","On Hold Reason","Responsible for Delay",
                           "Client Responsiveness","Client Sentiment","Risk Level","Risk Owner"]
        else:
            _col_order = ["Flags","Customer","Project Type","Start Date","Est. Go-Live",
                          "Phase","Last Milestone","On Hold Reason","Responsible for Delay",
                          "Client Responsiveness","Client Sentiment","Risk Level","Risk Owner"]
        _oh_df = _oh_df[[c for c in _col_order if c in _oh_df.columns]]

        # ✦ = SS syncable (editable) | no mark = derived/read-only
        st.caption("Columns with the edit icon sync back to Smartsheet — edit and export to CSV to update DRS. Greyed columns are derived or read-only.")
        _oh_edited = st.data_editor(
            _oh_df,
            column_config={
                "Flags":                 st.column_config.TextColumn("Flags",                  disabled=True, width="small"),
                    "Customer":              st.column_config.TextColumn("Customer",                disabled=True, width="medium"),
                "Project Type":          st.column_config.TextColumn("Project Type",            disabled=True, width="medium"),
                "Start Date":            st.column_config.TextColumn("Start Date",              disabled=True, width="small"),
                "Est. Go-Live":          st.column_config.TextColumn("Est. Go-Live",            disabled=True, width="small"),
                "Phase":                 st.column_config.SelectboxColumn("Phase", options=PHASE_OPTIONS, width="medium"),
                "On Hold Reason":        st.column_config.SelectboxColumn("On Hold Reason",  options=_OH_REASON_OPTS,     width="medium"),
                "Last Milestone":        st.column_config.TextColumn("Last Milestone",          disabled=True),
                "Client Responsiveness": st.column_config.SelectboxColumn("Client Responsiveness", options=_OH_RESP_OPTS, width="medium"),
                "Client Sentiment":      st.column_config.SelectboxColumn("Client Sentiment", options=_OH_SENTIMENT_OPTS, width="small"),
                "Risk Level":            st.column_config.SelectboxColumn("Risk Level",       options=_OH_RISK_OPTS,       width="small"),
                "Risk Owner":            st.column_config.SelectboxColumn("Risk Owner",       options=_OH_OWNER_OPTS,      width="small"),
                "Responsible for Delay": st.column_config.SelectboxColumn("Responsible for Delay", options=_OH_DELAY_OPTS, width="medium"),
            },
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="oh_edit_table",
        )

        # Export bar
        _oh_sync_cols = ["Phase","Last Milestone","On Hold Reason","Responsible for Delay","Client Responsiveness","Client Sentiment",
                         "Risk Level","Risk Owner"]
        _oh_changed = _oh_edited[_oh_sync_cols].fillna("").ne(_oh_df[[c for c in _oh_sync_cols if c in _oh_df.columns]].fillna("")).any(axis=1) if not _oh_edited.empty else pd.Series(False, index=_oh_edited.index)
        _oh_ex1, _oh_ex2 = st.columns([3,1])
        with _oh_ex1:
            if _oh_changed.any():
                st.markdown(f'<span style="font-size:13px;color:#27AE60;font-weight:600">✓ {_oh_changed.sum()} on-hold project(s) edited — ready to export</span>', unsafe_allow_html=True)
            else:
                st.markdown('<span style="font-size:12px;opacity:.5">Edit ✦ columns above then export to sync with DRS</span>', unsafe_allow_html=True)
        with _oh_ex2:
            _oh_buf = __import__("io").BytesIO()
            _oh_edited[_oh_sync_cols].to_csv(_oh_buf, index=False)
            st.download_button(
                label="⬇ Export to CSV",
                data=_oh_buf.getvalue(),
                file_name=f"on_hold_updates_{__import__('datetime').date.today().isoformat()}.csv",
                mime="text/csv",
                use_container_width=True,
            )


# ═══════════════════════════════════════════════════════════════════
# TAB 4 — Project Intake
# ═══════════════════════════════════════════════════════════════════
with tab_intake:
    # Project picker — scoped to logged-in consultant / view-as
    _all_proj = pd.concat([active, on_hold], ignore_index=True) if not on_hold.empty else active.copy()
    _pid_col_d = "project_id" if "project_id" in _all_proj.columns else "project_name"
    _all_pids  = [str(r.get(_pid_col_d,"")) for _,r in _all_proj.iterrows()]
    _all_labels= {}
    for _,r in _all_proj.iterrows():
        _k = str(r.get(_pid_col_d,""))
        _oh_tag = " [On Hold]" if str(r.get("status","")).lower() in ("on hold","onhold") else ""
        _all_labels[_k] = _extract_customer_name(str(r.get("project_name","")))+" — "+str(r.get("project_type",""))+_oh_tag

    _sel_pid = st.selectbox(
        "Select project",
        options=_all_pids,
        format_func=lambda p: _all_labels.get(p,p),
        key="_mp_detail_pid",
    )

    if _sel_pid:
        _dm = _all_proj[_all_proj[_pid_col_d].astype(str)==str(_sel_pid)]
        if not _dm.empty:
            _dr = _dm.iloc[0]
            _dr_name = str(_dr.get("project_name",""))
            _dr_rag  = str(_dr.get("rag","") or "").strip().lower()
            _rag_col = {"red":"#E24B4A","yellow":"#EF9F27","green":"#639922"}.get(_dr_rag,"rgba(128,128,128,.3)")
            _is_oh   = str(_dr.get("status","")).lower() in ("on hold","onhold") or _dr.get("_on_hold",False)

            def _dv(key, fallback="—"):
                v = _dr.get(key)
                if v is None or str(v).strip() in ("","nan","None","NaT"): return fallback
                return str(v).strip()
            def _dfmt(key):
                v = _dr.get(key)
                try: return pd.Timestamp(v).strftime("%-d %b %Y")
                except: return "—"

            _left, _right = st.columns([1,1], gap="medium")

            with _left:
                # ── READ fields ───────────────────────────────────────────────
                def _ri(label, val, link=False):
                    _vc = "var(--color-text-info)" if link else ("var(--color-text-primary)" if val!="—" else "var(--color-text-secondary)")
                    _op = "" if val!="—" else ";opacity:.45"
                    return (f"<tr><td style='color:var(--color-text-secondary);padding:5px 0;"
                            f"font-size:12px;width:48%'>{label}</td>"
                            f"<td style='text-align:right;color:{_vc};font-size:12px{_op}'>{val}</td></tr>")

                _project_rows = (
                    _ri("Project ID",       _dv("project_id")) +
                    _ri("Project name",     _extract_customer_name(_dr_name)) +
                    _ri("Project type",     _dv("project_type")) +
                    _ri("Status",           _dv("status")) +
                    _ri("Phase",            _dv("phase")) +
                    _ri("Start date",       _dfmt("start_date")) +
                    _ri("Est. go-live",     _dfmt("go_live_date")) +
                    _ri("Territory",        _dv("territory")) +
                    _ri("Project manager",  _dv("project_manager"))
                )
                _intake_rows = (
                    _ri("Customer",         _dv("account")) +
                    _ri("Customer ID",      _dv("customer_id")) +
                    _ri("Program ID",       _dv("program_id")) +
                    _ri("Program name",     _dv("program_name")) +
                    _ri("Partner name",     _dv("partner_name")) +
                    _ri("Implementer",      _dv("implementer")) +
                    _ri("Signed date",      _dfmt("signed_date")) +
                    _ri("Sub. start date",  _dfmt("subscription_start_date")) +
                    _ri("Original go-live", _dfmt("original_go_live_date"))
                )
                _team_rows = (
                    _ri("Sales rep",        _dv("sales_rep")) +
                    _ri("CSM",              _dv("csm")) +
                    _ri("Account manager",  _dv("account_manager")) +
                    _ri("Partner PM",       _dv("partner_pm")) +
                    _ri("Solution architect",_dv("solution_architect")) +
                    _ri("Lead consultant",  _dv("lead_consultant")) +
                    _ri("Support consultant",_dv("support_consultant")) +
                    _ri("Partner team",     _dv("partner_team")) +
                    _ri("Impl. contact",    _dv("implementation_contact_email"), link=True) +
                    _ri("NS Account #",     _dv("customer_netsuite_account")) +
                    _ri("Sandbox Account #",_dv("customer_sandbox_account"))
                )

                st.markdown(
                    f"<div style='border-left:3px solid {_rag_col};"
                    f"border-radius:0 10px 10px 0;"
                    f"border-top:0.5px solid rgba(128,128,128,.15);"
                    f"border-right:0.5px solid rgba(128,128,128,.15);"
                    f"border-bottom:0.5px solid rgba(128,128,128,.15);"
                    f"padding:16px 18px;'>"

                    f"<div style='font-size:10px;font-weight:600;text-transform:uppercase;"
                    f"letter-spacing:.8px;color:var(--color-text-secondary);margin-bottom:8px'>Project</div>"
                    f"<table style='width:100%;border-collapse:collapse'>{_project_rows}</table>"

                    f"<div style='font-size:10px;font-weight:600;text-transform:uppercase;"
                    f"letter-spacing:.8px;color:var(--color-text-secondary);margin:14px 0 8px;"
                    f"padding-top:12px;border-top:0.5px solid rgba(128,128,128,.1)'>Intake</div>"
                    f"<table style='width:100%;border-collapse:collapse'>{_intake_rows}</table>"

                    f"<div style='font-size:10px;font-weight:600;text-transform:uppercase;"
                    f"letter-spacing:.8px;color:var(--color-text-secondary);margin:14px 0 8px;"
                    f"padding-top:12px;border-top:0.5px solid rgba(128,128,128,.1)'>Team</div>"
                    f"<table style='width:100%;border-collapse:collapse'>{_team_rows}</table>"
                    f"</div>",
                    unsafe_allow_html=True
                )

            with _right:
                # ── WRITE fields ──────────────────────────────────────────────
                _badge = lambda t: f"<span style='font-size:9px;padding:2px 6px;border-radius:4px;background:var(--color-background-info);color:var(--color-text-info);margin-left:6px'>{t}</span>"

                # Project dates — order: Start · Go-Live · End
                st.markdown(f"<div class='section-label' style='margin-bottom:10px'>Project Dates {_badge('editable')}</div>",unsafe_allow_html=True)
                _sd_cur = _dr.get("start_date")
                _sd_val = pd.Timestamp(_sd_cur).date() if pd.notna(_sd_cur) else None
                _w_substart = st.date_input("Start date", value=_sd_val, key=f"dp_sd_{_sel_pid}")
                _wc1,_wc2 = st.columns(2)
                with _wc1:
                    _gl_cur = _dr.get("go_live_date")
                    _gl_val = pd.Timestamp(_gl_cur).date() if pd.notna(_gl_cur) else None
                    _w_golive = st.date_input("Go-Live date", value=_gl_val, key=f"dp_gl_{_sel_pid}")
                with _wc2:
                    _fd_cur = _dr.get("finish_date")
                    _fd_val = pd.Timestamp(_fd_cur).date() if pd.notna(_fd_cur) else None
                    _w_finish = st.date_input("End date", value=_fd_val, key=f"dp_fd_{_sel_pid}")

                st.markdown("<div style='height:10px'></div>",unsafe_allow_html=True)

                # Weekly health
                st.markdown(f"<div class='section-label' style='margin:14px 0 8px;padding-top:12px;border-top:0.5px solid rgba(128,128,128,.15)'>Weekly Health {_badge('editable')}</div>",unsafe_allow_html=True)
                _opts_status    = ["In Progress","On Hold","Complete","Closed","Cancelled"]
                _opts_phase     = PHASE_OPTIONS
                _opts_sentiment = ["","Good","Neutral","Concern","At Risk"]
                _opts_health    = ["","Green","Amber","Red"]
                _opts_risk      = ["","Low","Medium","High","Critical"]

                _w_status = st.selectbox("Status",_opts_status,
                    index=_opts_status.index(_dv("status","In Progress")) if _dv("status","In Progress") in _opts_status else 0,
                    key=f"w_status_{_sel_pid}")
                _w_phase  = st.selectbox("Phase",_opts_phase,
                    index=_opts_phase.index(_dv("phase")) if _dv("phase") in _opts_phase else 0,
                    key=f"w_phase_{_sel_pid}")
                _w_summary = st.text_area("Overall summary",value=_dv("overall_summary",""),height=72,
                    placeholder="Brief project status update...",key=f"w_sum_{_sel_pid}")

                # ── Milestone dates — 5×2 grid ────────────────────────────────
                st.markdown(
                    "<div class='section-label' style='margin:14px 0 8px;padding-top:12px;border-top:0.5px solid rgba(128,128,128,.15)'>Milestones</div>",
                    unsafe_allow_html=True
                )
                _ms_write_cols = [
                    ("ms_intro_email","Intro Email Sent"),
                    ("ms_config_start","Config"),
                    ("ms_enablement","Enablement Session"),
                    ("ms_session1","Session #1"),
                    ("ms_session2","Session #2"),
                    ("ms_uat_signoff","UAT Signoff"),
                    ("ms_prod_cutover","Prod Cutover"),
                    ("ms_hypercare_start","Hypercare Start"),
                    ("ms_close_out","Task Close Out"),
                    ("ms_transition","Transition to Support"),
                ]
                _row1 = st.columns(5)
                _row2 = st.columns(5)
                for _i, (_mk, _ml) in enumerate(_ms_write_cols):
                    _mv = _dr.get(_mk)
                    _mv_val = pd.Timestamp(_mv).date() if pd.notna(_mv) else None
                    _target_col = _row1[_i] if _i < 5 else _row2[_i - 5]
                    with _target_col:
                        st.date_input(_ml, value=_mv_val, key=f"w_ms_{_mk}_{_sel_pid}")

                st.markdown("<div style='margin:14px 0 8px;padding-top:12px;border-top:0.5px solid rgba(128,128,128,.2)'></div>", unsafe_allow_html=True)
                st.markdown(f"<div class='section-label' style='margin-bottom:8px'>Project Health {_badge('editable')}</div>", unsafe_allow_html=True)
                _hc1,_hc2 = st.columns(2)
                with _hc1:
                    _w_sched = st.selectbox("Schedule health",_opts_health,
                        index=_opts_health.index(_dv("schedule_health","")) if _dv("schedule_health","") in _opts_health else 0,
                        key=f"w_sch_{_sel_pid}")
                    _w_res   = st.selectbox("Resource health",_opts_health,
                        index=_opts_health.index(_dv("resource_health","")) if _dv("resource_health","") in _opts_health else 0,
                        key=f"w_res_{_sel_pid}")
                with _hc2:
                    _w_scope = st.selectbox("Scope health",_opts_health,
                        index=_opts_health.index(_dv("scope_health","")) if _dv("scope_health","") in _opts_health else 0,
                        key=f"w_sco_{_sel_pid}")
                    _w_risk  = st.selectbox("Risk level",_opts_risk,
                        index=_opts_risk.index(_dv("risk_level","")) if _dv("risk_level","") in _opts_risk else 0,
                        key=f"w_rsk_{_sel_pid}")
                _w_riskd = st.text_area("Risk detail",value=_dv("risk_detail",""),height=56,
                    placeholder="Describe risk or mitigation...",key=f"w_rkd_{_sel_pid}")
                _w_cresp = st.selectbox("Client responsiveness",_opts_sentiment,
                    index=_opts_sentiment.index(_dv("client_responsiveness","")) if _dv("client_responsiveness","") in _opts_sentiment else 0,
                    key=f"w_crsp_{_sel_pid}")
                _w_csent = st.selectbox("Client sentiment",_opts_sentiment,
                    index=_opts_sentiment.index(_dv("client_sentiment","")) if _dv("client_sentiment","") in _opts_sentiment else 0,
                    key=f"w_csnt_{_sel_pid}")

                # On Hold fields — conditional
                if _is_oh:
                    st.markdown(f"<div class='section-label' style='margin:14px 0 8px;padding-top:12px;border-top:0.5px solid rgba(128,128,128,.15)'>On Hold {_badge('editable')}</div>",unsafe_allow_html=True)
                    _oh_reason_opts = ["None","Customer delay","Internal delay","Technical blocker","Commercial","Other"]
                    _w_oh_reason = st.selectbox("On Hold reason",_oh_reason_opts,
                        index=_oh_reason_opts.index(_dv("on_hold_reason","None")) if _dv("on_hold_reason","None") in _oh_reason_opts else 0,
                        key=f"w_ohr_{_sel_pid}")
                    _oh_resume = _dr.get("resume_date")
                    _oh_resume_val = pd.Timestamp(_oh_resume).date() if pd.notna(_oh_resume) else None
                    _w_resume = st.date_input("Resume date",value=_oh_resume_val,key=f"w_ohrd_{_sel_pid}")
                    _w_oh_resp = st.text_area("On Hold response",value=_dv("on_hold_response",""),height=56,
                        placeholder="Customer/internal response...",key=f"w_ohrs_{_sel_pid}")
                    _w_trans_notes = st.text_area("Support transition notes",value=_dv("support_transition_notes",""),height=56,
                        placeholder="Notes for support handoff...",key=f"w_trn_{_sel_pid}")

                # On Hold fields — conditional
                if _is_oh:
                    st.markdown(f"<div class='section-label' style='margin:14px 0 8px;padding-top:12px;border-top:0.5px solid rgba(128,128,128,.15)'>On Hold {_badge('editable')}</div>",unsafe_allow_html=True)
                    _oh_reason_opts = ["None","Customer delay","Internal delay","Technical blocker","Commercial","Other"]
                    _w_oh_reason = st.selectbox("On Hold reason",_oh_reason_opts,
                        index=_oh_reason_opts.index(_dv("on_hold_reason","None")) if _dv("on_hold_reason","None") in _oh_reason_opts else 0,
                        key=f"w_ohr_{_sel_pid}")
                    _oh_resume = _dr.get("resume_date")
                    _oh_resume_val = pd.Timestamp(_oh_resume).date() if pd.notna(_oh_resume) else None
                    _w_resume = st.date_input("Resume date",value=_oh_resume_val,key=f"w_ohrd_{_sel_pid}")
                    _w_oh_resp = st.text_area("On Hold response",value=_dv("on_hold_response",""),height=56,
                        placeholder="Customer/internal response...",key=f"w_ohrs_{_sel_pid}")
                    _w_trans_notes = st.text_area("Support transition notes",value=_dv("support_transition_notes",""),height=56,
                        placeholder="Notes for support handoff...",key=f"w_trn_{_sel_pid}")
                    _w_delay_sum = st.text_area("Delay summary",value=_dv("delay_summary",""),height=56,
                        placeholder="Summary of delay cause and current status...",key=f"w_dsum_{_sel_pid}")
                else:
                    _w_oh_reason = _w_oh_resp = _w_trans_notes = _w_resume = None
                    _w_delay_sum = ""

                # JIRA Links — optional, shown at bottom for all projects
                st.markdown(f"<div class='section-label' style='margin:14px 0 8px;padding-top:12px;border-top:0.5px solid rgba(128,128,128,.15)'>JIRA <span style='font-size:9px;padding:2px 6px;border-radius:4px;background:rgba(128,128,128,.12);color:rgba(128,128,128,.7);margin-left:6px'>optional</span></div>",unsafe_allow_html=True)
                _w_jira = st.text_input("JIRA links",value=_dv("jira_links",""),
                    placeholder="e.g. https://zone.atlassian.net/browse/ZPS-123",
                    key=f"w_jira_{_sel_pid}")

                # Save button
                if st.button("Save to Smartsheet", key=f"mp_det_save_{_sel_pid}",
                             type="primary", use_container_width=True):
                    _row_id = _dr.get("_ss_row_id")
                    if _row_id and _ss_ready:
                        _changes = {}

                        def _norm(v):
                            if v is None: return ""
                            s = str(v).strip()
                            return "" if s in ("nan","None","NaT","—") else s

                        def _orig_date(key):
                            v = _dr.get(key)
                            try: return pd.Timestamp(v).date().isoformat() if pd.notna(v) else ""
                            except: return ""

                        # Dates — diff only
                        if _w_golive and _w_golive.isoformat() != _orig_date("go_live_date"):
                            _changes["go_live_date"] = _w_golive.isoformat()
                        if _w_finish and _w_finish.isoformat() != _orig_date("finish_date"):
                            _changes["finish_date"] = _w_finish.isoformat()
                        if _w_substart and _w_substart.isoformat() != _orig_date("start_date"):
                            _changes["start_date"] = _w_substart.isoformat()

                        # Health fields — diff only
                        _health_fields = {
                            "status":                _w_status,
                            "phase":                 _w_phase,
                            "overall_summary":       _w_summary,
                            "schedule_health":       _w_sched,
                            "resource_health":       _w_res,
                            "scope_health":          _w_scope,
                            "risk_level":            _w_risk,
                            "risk_detail":           _w_riskd,
                            "client_responsiveness": _w_cresp,
                            "client_sentiment":      _w_csent,
                        }
                        for _fk, _fv in _health_fields.items():
                            if _norm(_fv) != _norm(_dv(_fk)):
                                _changes[_fk] = _fv

                        # On Hold fields — diff only
                        if _is_oh:
                            for _fk, _fv in {
                                "on_hold_reason":           _w_oh_reason,
                                "on_hold_response":         _w_oh_resp,
                                "support_transition_notes": _w_trans_notes,
                                "delay_summary":            _w_delay_sum,
                            }.items():
                                if _norm(_fv) != _norm(_dv(_fk)):
                                    _changes[_fk] = _fv
                            if _w_resume and _w_resume.isoformat() != _orig_date("resume_date"):
                                _changes["resume_date"] = _w_resume.isoformat()

                        # JIRA — diff only, all projects
                        if _norm(_w_jira) != _norm(_dv("jira_links")):
                            _changes["jira_links"] = _w_jira

                        # Milestone dates — diff only
                        for _mk, _ in _ms_write_cols:
                            _mw = st.session_state.get(f"w_ms_{_mk}_{_sel_pid}")
                            if _mw and _mw.isoformat() != _orig_date(_mk):
                                _changes[_mk] = _mw.isoformat()

                        # Drop empty
                        _changes = {k: v for k, v in _changes.items() if _norm(v) != ""}

                        if _changes:
                            with st.spinner("Saving to Smartsheet..."):
                                _ok,_errs = write_row_updates([{
                                    "_ss_row_id":   int(_row_id),
                                    "project_name": _dr_name,
                                    "changes":      _changes,
                                }])
                            if _ok: st.success(f"✓ Saved {len(_changes)} field(s) to Smartsheet")
                            for _e in (_errs or []): st.warning(f"⚠ {_e}")
                        else:
                            st.info("No changes to save.")
                    else:
                        st.info("Sync DRS via API on Home page to enable Smartsheet writeback.")

st.markdown('<div style="font-size:11px;opacity:.4;text-align:center;margin-top:20px">PS Projects & Tools · Internal use only · Data loaded this session only</div>',unsafe_allow_html=True)
