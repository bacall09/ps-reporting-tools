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
on_hold= my_drs[_ioh].copy()
active = my_drs[~_ioh].copy()
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
    # ── Build NS lookups keyed by project_id ─────────────────────────────────
    def _clean_pid(v):
        try:
            s = str(v).strip()
            if s in ("", "nan", "None"): return ""
            return str(int(float(s)))
        except: return str(v).strip()

    _ns_htd: dict    = {}  # project_id → max hours_to_date
    _ns_tm_hrs: dict = {}  # project_id → T&M scope hours
    _ns_tm_pids: set = set()  # project_ids confirmed T&M from NS billing_type

    if df_ns is not None:
        _ns_id_col = "project_id" if "project_id" in df_ns.columns else None
        if _ns_id_col and "hours_to_date" in df_ns.columns:
            for _pid, _grp in df_ns.groupby(_ns_id_col):
                _k = _clean_pid(_pid)
                if _k:
                    try:
                        _ns_htd[_k] = round(float(_grp["hours_to_date"].dropna().astype(float).max() or 0), 2)
                    except Exception:
                        pass
        # T&M scope — read directly from the "T&M Scope" column in NS Time Detail
        # (only populated for T&M projects — replaces the hours-sum proxy)
        if _ns_id_col and "tm_scope" in df_ns.columns:
            for _pid, _grp in df_ns.groupby(_ns_id_col):
                _k = _clean_pid(_pid)
                if _k:
                    try:
                        _v = _grp["tm_scope"].dropna().astype(float)
                        if not _v.empty:
                            _ns_tm_hrs[_k] = round(float(_v.max()), 2)
                            _ns_tm_pids.add(_k)
                    except Exception:
                        pass
        # Also flag T&M by billing_type in NS (catches projects where tm_scope is present)
        if _ns_id_col and "billing_type" in df_ns.columns:
            _tm_ns = df_ns[df_ns["billing_type"].fillna("").str.strip().str.lower() == "t&m"]
            for _pid in _tm_ns[_ns_id_col].dropna().unique():
                _k = _clean_pid(_pid)
                if _k: _ns_tm_pids.add(_k)
        elif _ns_id_col and "hours" in df_ns.columns and "billing_type" in df_ns.columns:
            # Fallback: sum T&M hours if T&M Scope column not present
            _tm_mask = df_ns["billing_type"].fillna("").str.strip().str.lower() == "t&m"
            for _pid, _grp in df_ns[_tm_mask].groupby(_ns_id_col):
                _k = _clean_pid(_pid)
                if _k:
                    try:
                        _ns_tm_hrs[_k] = round(float(_grp["hours"].sum() or 0), 2)
                    except Exception:
                        pass

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
    # Balance: scope - htd. If no NS data yet (htd=""), show full scope as remaining
    if _scope != "" and _htd != "":
        _bal = round(float(_scope) - float(_htd), 2)
    elif _scope != "":
        _bal = float(_scope)  # no hours logged yet — full scope remaining
        _htd = 0.0
    else:
        _bal = ""
    # Balance cell flag logic
    _phase_raw = str(row.get("phase","") or "").strip().lower()
    _closed_phases = {"08. ready for support transition","09. phase 2 scoping","closed","complete"}
    _late_phases   = {"05. prep for go-live","06. go-live (hypercare)","07. hypercare","08. ready for support transition"}
    _bal_flag = ""
    if _bal != "" and _scope not in ("", 0):
        _pct_remaining = float(_bal) / float(_scope) if float(_scope) != 0 else 0
        if float(_bal) < 0 and _phase_raw not in _closed_phases:
            _bal_flag = "red"
        elif _pct_remaining <= 0.10 and _phase_raw not in _late_phases and _phase_raw not in _closed_phases:
            _bal_flag = "yellow"

    # No RAG flag
    _rag_val = _rag_emoji(row.get("rag"))
    if _rag_val == "—":
        _rag_val = "⚠️ No RAG"

    return {
        "Flags":                needs,
        "RAG":                  _rag_val,
        "Customer":             _cust,
        "Consultant":           str(row.get("project_manager","") or ""),
        "Project Type":         str(row.get("project_type","") or ""),
        "Status":               str(row.get("status","") or ""),
        "Phase":                str(row.get("phase","") or ""),
        "Start Date":           _dt("start_date"),
        "Est. Go-Live":         _dt("go_live_date"),
        "Scope Hrs":            _scope,
        "Hours to Date":        _htd,
        "Balance":              _bal,
        "_bal_flag":            _bal_flag,
        "Engagement":           _engagement_flag({**dict(row), "_htd_val": _htd}),
        "Intro Email Sent":     _ms("ms_intro_email"),
        "Config Start":         _ms("ms_config_start"),
        "Enablement Session":   _ms("ms_enablement"),
        "Session #1":           _ms("ms_session1"),
        "Session #2":           _ms("ms_session2"),
        "UAT Signoff":          _ms("ms_uat_signoff"),
        "Prod Cutover":         _ms("ms_prod_cutover"),
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
            "UAT Signoff","Prod Cutover","Hypercare Start","Close Out Tasks","Transition to Support"]
col_cfg = {
    "Flags":                 st.column_config.TextColumn("Flags",             disabled=True, width="small"),
    "RAG":                   st.column_config.TextColumn("RAG",               disabled=True, width="small"),
    "Customer":              st.column_config.TextColumn("Customer",          disabled=True),
    "Consultant":            st.column_config.TextColumn("Consultant",        disabled=True),
    "Project Type":          st.column_config.TextColumn("Project Type",      disabled=True),
    "Status":                st.column_config.SelectboxColumn("Status", options=["In Progress","On Hold","Closed","Complete","Cancelled"], width="medium"),
    "Phase":                 st.column_config.SelectboxColumn("Phase",        options=PHASE_OPTIONS, width="medium"),
    "Start Date":            st.column_config.TextColumn("Start Date",        disabled=True, width="small"),
    "Est. Go-Live":          st.column_config.TextColumn("Est. Go-Live",      disabled=True, width="small"),
    "Scope Hrs":             st.column_config.NumberColumn("Scope Hrs",         disabled=True, width="small"),
    "Hours to Date":         st.column_config.NumberColumn("Hours to Date",     disabled=True, width="small"),
    "Balance":               st.column_config.TextColumn("Balance",            disabled=True, width="small"),
    "Engagement":            st.column_config.TextColumn("Engagement",         disabled=True, width="medium"),
    "_bal_flag":             None,
    **{c: st.column_config.DateColumn(c, min_value=date(2020,1,1), max_value=date(2030,12,31), width="small") for c in _ms_cols},
}

st.caption("Columns with the edit icon sync back to Smartsheet — edit and export to CSV to update DRS. Greyed columns are derived or read-only.")
st.markdown('<span style="font-size:11.5px;opacity:.6">⚠️ Flags indicate date issues, missing milestones, or phase gaps. For a deeper look at data quality issues, use the DRS Health Check page.</span>', unsafe_allow_html=True)
_btn_col1, _btn_col2 = st.columns([1, 1])
with _btn_col1:
    if st.button("→ Go to DRS Health Check", key="mp_drs_link", use_container_width=True):
        if _va_region:
            st.session_state["_va_passthrough"]    = f"── {_va_region} ──"
            st.session_state["_browse_passthrough"] = f"── {_va_region} ──"
        elif view_as and view_as != selected:
            st.session_state["_va_passthrough"]    = view_as
            st.session_state["_browse_passthrough"] = view_as
        st.switch_page("pages/6_DRS_Health_Check.py")
with _btn_col2:
    if st.button("→ Draft Outreach", key="mp_engagement_link", use_container_width=True):
        if _va_region:
            st.session_state["_browse_passthrough"] = f"── {_va_region} ──"
        elif view_as and view_as != selected:
            st.session_state["_browse_passthrough"] = view_as
        st.switch_page("pages/2_Customer_Reengagement.py")

# Apply Balance cell background styling
def _style_balance(df):
    styles = pd.DataFrame("", index=df.index, columns=df.columns)
    if "Balance" in df.columns and "_bal_flag" in df.columns:
        for i, row in df.iterrows():
            flag = str(row.get("_bal_flag",""))
            if flag == "red":
                styles.at[i, "Balance"] = "background-color:#FDECED;color:#C0392B;font-weight:600"
            elif flag == "yellow":
                styles.at[i, "Balance"] = "background-color:#FFF3CD;color:#B7770D;font-weight:600"
    return styles

# st.data_editor doesn't support Styler — apply Balance colour as a display-only
# text prefix flag instead, keep _bal_flag hidden via col_cfg None
_display_df = edit_df.copy()
# Inject colour signal into Balance value as a suffixed marker for display
if "Balance" in _display_df.columns and "_bal_flag" in _display_df.columns:
    def _fmt_bal(row):
        val = row["Balance"]
        flag = str(row.get("_bal_flag","") or "")
        if val == "" or val is None: return val
        try:
            fval = float(val)
            if flag == "red":    return f"🔴 {fval:,.2f}"
            if flag == "yellow": return f"🟡 {fval:,.2f}"
            return round(fval, 2)
        except Exception:
            return val
    _display_df["Balance"] = _display_df.apply(_fmt_bal, axis=1)
_display_df = _display_df.drop(columns=["_bal_flag"], errors="ignore")
edited = st.data_editor(
    _display_df,
    column_config=col_cfg,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    key="mp_edit_table",
)

# ── Detect changes vs original ────────────────────────────────────────────
editable_cols = ["Phase","Intro Email Sent","Config Start","Enablement Session","Session #1","Session #2","UAT Signoff","Prod Cutover","Hypercare Start","Close Out Tasks","Transition to Support"]
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

            # Join edited table back to active df to recover _ss_row_id
            # active has _ss_row_id; edited has display columns. Join on positional index.
            _sync_rows = active[changed].copy()
            _edited_changed = edited[changed].copy()

            updates = []
            for idx in _sync_rows.index:
                _pos = list(_sync_rows.index).index(idx)
                row_id    = _sync_rows.at[idx, "_ss_row_id"]
                proj_name = _sync_rows.at[idx, "project_name"] if "project_name" in _sync_rows.columns else str(row_id)

                changes = {}
                for disp_col in _editable_display_cols:
                    if disp_col not in _edited_changed.columns:
                        continue
                    internal_key = _display_to_internal.get(disp_col)
                    if not internal_key:
                        continue
                    new_val = _edited_changed.iloc[_pos][disp_col]
                    # Only include fields that actually changed vs original
                    orig_val = edit_df.iloc[_pos][disp_col] if disp_col in edit_df.columns else None
                    if str(new_val) != str(orig_val):
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
    st.markdown('<hr class="divider">', unsafe_allow_html=True)

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ── Project detail drawer ─────────────────────────────────────────────────────
# Project selector drives drawer — rendered as a clean selectbox below the table
if not active.empty:
    _pid_col_d = "project_id" if "project_id" in active.columns else "project_name"
    _all_pids  = [str(r.get(_pid_col_d,"")) for _,r in active.iterrows()]
    _all_labels= {str(r.get(_pid_col_d,"")): _extract_customer_name(str(r.get("project_name","")))+" — "+str(r.get("project_type",""))
                  for _,r in active.iterrows()}

    # ── Project detail card ─────────────────────────────────────────────────
    # Selectbox label acts as section header — no separate st.markdown label
    _sel_pid = st.selectbox(
        "Project detail",
        options=_all_pids,
        format_func=lambda p: _all_labels.get(p,p),
        key="_mp_proj_drawer",
    )

    if _sel_pid:
        _drawer_matches = active[active[_pid_col_d].astype(str)==str(_sel_pid)]
        if not _drawer_matches.empty:
            _dr = _drawer_matches.iloc[0]
            _dr_name  = str(_dr.get("project_name",""))
            _dr_cust  = _extract_customer_name(_dr_name)
            _dr_type  = str(_dr.get("project_type","—"))
            _dr_pm    = str(_dr.get("project_manager","—"))
            _dr_pid   = str(_dr.get("project_id","—"))
            _dr_rag   = str(_dr.get("rag","") or "").strip().lower()
            _rag_color= {"red":"#E24B4A","yellow":"#EF9F27","green":"#639922"}.get(_dr_rag,"rgba(128,128,128,.3)")

            def _dv(key, fallback="—"):
                v = _dr.get(key)
                if v is None or str(v).strip() in ("","nan","None","NaT"): return fallback
                return str(v).strip()

            def _dfmt(key):
                v = _dr.get(key)
                try: return pd.Timestamp(v).strftime("%-d %b %Y")
                except: return "—"

            # Build intake rows — only show populated fields
            _intake_fields = [
                ("Customer ID",      _dv("customer_id")),
                ("Program ID",       _dv("program_id")),
                ("Program name",     _dv("program_name")),
                ("Territory",        _dv("territory")),
                ("Signed date",      _dfmt("signed_date")),
                ("Implementer",      _dv("implementer")),
                ("Sales rep",        _dv("sales_rep")),
                ("CSM",              _dv("csm")),
                ("Account manager",  _dv("account_manager")),
                ("Impl. contact",    _dv("implementation_contact_email")),
                ("NS Account #",     _dv("customer_netsuite_account")),
                ("Sandbox Account #",_dv("customer_sandbox_account")),
            ]
            # Hide rows that are "—" to reduce noise (will populate after Friday's sheet)
            _intake_populated = [(lbl,val) for lbl,val in _intake_fields if val != "—"]
            _intake_empty_count = len(_intake_fields) - len(_intake_populated)

            _intake_rows_html = "".join(
                f"<tr><td style='color:var(--color-text-secondary);padding:5px 0;width:45%;font-size:12px'>{lbl}</td>"
                f"<td style='text-align:right;color:{'var(--color-text-info)' if 'contact' in lbl.lower() else 'var(--color-text-primary)'};font-size:12px'>{val}</td></tr>"
                for lbl,val in _intake_populated
            )
            _intake_empty_note = (
                f"<div style='font-size:11px;color:var(--color-text-secondary);opacity:.5;margin-top:8px'>"
                f"{_intake_empty_count} fields pending provisioned sheet</div>"
                if _intake_empty_count > 0 else ""
            )

            dcol_left, dcol_right = st.columns([1,1], gap="medium")

            with dcol_left:
                st.markdown(
                    f"<div style='border-left:3px solid {_rag_color};"
                    f"border-radius:0 10px 10px 0;border-top:0.5px solid rgba(128,128,128,.15);"
                    f"border-right:0.5px solid rgba(128,128,128,.15);border-bottom:0.5px solid rgba(128,128,128,.15);"
                    f"padding:16px 18px;'>"
                    f"<div style='font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.8px;"
                    f"color:var(--color-text-secondary);margin-bottom:12px'>Intake</div>"
                    f"<table style='width:100%;border-collapse:collapse'>{_intake_rows_html}</table>"
                    f"{_intake_empty_note}"
                    f"</div>",
                    unsafe_allow_html=True
                )

            with dcol_right:
                # ── Weekly health fields (editable) ───────────────────────────
                st.markdown('<div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.8px;color:var(--color-text-secondary);margin-bottom:8px">Weekly health update <span style=\'font-size:9px;padding:1px 5px;border-radius:4px;background:var(--color-background-info);color:var(--color-text-info)\'>editable</span></div>', unsafe_allow_html=True)

                _health_opts_sentiment   = ["","Good","Neutral","Concern","At Risk"]
                _health_opts_health      = ["","Green","Amber","Red"]
                _health_opts_risk        = ["","Low","Medium","High","Critical"]

                _h_summary  = st.text_area("Overall summary",
                    value=_dv("overall_summary",""),
                    height=80, key=f"mp_h_summary_{_sel_pid}",
                    placeholder="Brief project status update...")
                _h_sentiment= st.selectbox("Client sentiment",
                    _health_opts_sentiment,
                    index=_health_opts_sentiment.index(_dv("client_sentiment","")) if _dv("client_sentiment","") in _health_opts_sentiment else 0,
                    key=f"mp_h_sentiment_{_sel_pid}")
                _hcols = st.columns(2)
                with _hcols[0]:
                    _h_sched = st.selectbox("Schedule",_health_opts_health,
                        index=_health_opts_health.index(_dv("schedule_health","")) if _dv("schedule_health","") in _health_opts_health else 0,
                        key=f"mp_h_sched_{_sel_pid}")
                    _h_resource = st.selectbox("Resource",_health_opts_health,
                        index=_health_opts_health.index(_dv("resource_health","")) if _dv("resource_health","") in _health_opts_health else 0,
                        key=f"mp_h_resource_{_sel_pid}")
                with _hcols[1]:
                    _h_scope = st.selectbox("Scope",_health_opts_health,
                        index=_health_opts_health.index(_dv("scope_health","")) if _dv("scope_health","") in _health_opts_health else 0,
                        key=f"mp_h_scope_{_sel_pid}")
                    _h_risk = st.selectbox("Risk level",_health_opts_risk,
                        index=_health_opts_risk.index(_dv("risk_level","")) if _dv("risk_level","") in _health_opts_risk else 0,
                        key=f"mp_h_risk_{_sel_pid}")
                _h_risk_detail = st.text_area("Risk detail",
                    value=_dv("risk_detail",""),
                    height=60, key=f"mp_h_riskd_{_sel_pid}",
                    placeholder="Describe risk or mitigation...")
                _h_client_resp = st.selectbox("Client responsiveness",_health_opts_sentiment,
                    index=_health_opts_sentiment.index(_dv("client_responsiveness","")) if _dv("client_responsiveness","") in _health_opts_sentiment else 0,
                    key=f"mp_h_cresp_{_sel_pid}")

                _health_fields_map = {
                    "overall_summary":      _h_summary,
                    "client_sentiment":     _h_sentiment,
                    "schedule_health":      _h_sched,
                    "resource_health":      _h_resource,
                    "scope_health":         _h_scope,
                    "risk_level":           _h_risk,
                    "risk_detail":          _h_risk_detail,
                    "client_responsiveness":_h_client_resp,
                }

                if st.button("Save weekly update to Smartsheet",
                             key=f"mp_h_save_{_sel_pid}",
                             type="primary", use_container_width=True):
                    _row_id = _dr.get("_ss_row_id")
                    if _row_id and _ss_ready:
                        _changes = {k:v for k,v in _health_fields_map.items() if v and str(v).strip()}
                        if _changes:
                            _ok2,_errs2 = write_row_updates([{
                                "_ss_row_id":   int(_row_id),
                                "project_name": _dr_name,
                                "changes":      _changes,
                            }])
                            if _ok2: st.success("✓ Weekly update saved to Smartsheet")
                            for _e2 in (_errs2 or []): st.warning(f"⚠ {_e2}")
                        else:
                            st.info("No changes to save.")
                    else:
                        st.info("Sync DRS via API on Home page to enable Smartsheet writeback.")

_OH_REASON_OPTS   = ["—","Zone Product Dependency","Zone Program Dependency",
                     "NetSuite Dependency","Client Requested","Client Unresponsive"]
_OH_RESP_OPTS     = ["—","Highly Engaged","Neutral","Not Responsive"]
_OH_SENTIMENT_OPTS= ["—","Positive","Neutral","Concerned"]
_OH_RISK_OPTS     = ["—","Low","Medium","High","Escalated"]
_OH_OWNER_OPTS    = ["—","Client","Product","PS","Sales","Marketing","Support","3rd Party","N/A"]
_OH_DELAY_OPTS    = ["—","Zone","Client","3rd Party"]

def _clean(val):
    """Normalise blank/nan/None to —."""
    v = str(val or "").strip()
    return "—" if v in ("", "nan", "None", "NaT") else v

def _delay_summary_prompt(r):
    """Auto-generate a Delay Summary prompt from available fields."""
    reason   = str(r.get("on_hold_reason","") or "")
    days     = r.get("days_inactive", 0) or 0
    risk     = str(r.get("risk_level","") or "")
    owner    = str(r.get("risk_owner","") or "")
    delay_by = str(r.get("responsible_for_delay","") or "")
    parts = []
    if reason and reason not in ("—",""):
        parts.append(reason)
    if days and int(days) > 0:
        parts.append(f"inactive {int(days)}d")
    if risk and risk not in ("—",""):
        parts.append(f"{risk} risk")
    if owner and owner not in ("—",""):
        parts.append(f"owner: {owner}")
    if delay_by and delay_by not in ("—",""):
        parts.append(f"delay: {delay_by}")
    return " · ".join(parts) if parts else ""

# ── Risk level emoji helper ────────────────────────────────────────────────────
def _risk_emoji(val):
    v = str(val or "").strip().lower()
    if v == "escalated": return "🚨 Escalated"
    if v == "high":      return "🔴 High"
    if v == "medium":    return "🟡 Medium"
    if v == "low":       return "🟢 Low"
    return str(val or "—")

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
            "RAG":                   _rag_v if _rag_v != "—" else "⚠️ No RAG",
            "Flags":                 ("⚠️" if any(s=="error" for s,*_ in (r.get("_flags") or []))
                                  else ("⚠️" if (r.get("_flags") or []) else "")),
            "Customer":              _extract_customer_name(str(r.get("project_name",""))),
            "Project Type":          str(r.get("project_type","") or "—"),
            "Start Date":            pd.Timestamp(r["start_date"]).strftime("%Y-%m-%d") if pd.notna(r.get("start_date")) else "—",
            "Est. Go-Live":          pd.Timestamp(r["go_live_date"]).strftime("%Y-%m-%d") if pd.notna(r.get("go_live_date")) else "—",
            "Phase":                 str(r.get("phase", "—")),
            "On Hold Reason":        _clean(r.get("on_hold_reason")) if _clean(r.get("on_hold_reason")) != "—" else None,
            "Days Inactive":         int(r.get("days_inactive", -1)) if r.get("days_inactive", -1) >= 0 else "—",
            "Inactivity Source":     _clean(r.get("_inactivity_source")),
            "Last Milestone":        _clean(r.get("last_milestone")),
            "Client Responsiveness": _clean(r.get("client_responsiveness")) if _clean(r.get("client_responsiveness")) != "—" else None,
            "Client Sentiment":      _clean(r.get("client_sentiment")) if _clean(r.get("client_sentiment")) != "—" else None,
            "Risk Level":            _risk_emoji(r.get("risk_level")),
            "Risk Owner":            _clean(r.get("risk_owner")) if _clean(r.get("risk_owner")) != "—" else None,
            "Risk Detail":           _clean(r.get("risk_detail")),
            "Responsible for Delay": _clean(r.get("responsible_for_delay")) if _clean(r.get("responsible_for_delay")) != "—" else None,
            "Delay Summary":         _ds,
            "JIRA Links":            _clean(r.get("jira_links")),
        }
        if _va_region:
            _oh_row["Consultant"] = str(r.get("project_manager", "") or "")
        _oh_rows.append(_oh_row)

    _oh_df = pd.DataFrame(_oh_rows)

    # Column order — insert Consultant after RAG if region view
    if "Consultant" in _oh_df.columns:
        _col_order = ["Flags","RAG","Customer","Consultant","Project Type","Start Date","Est. Go-Live",
                       "Phase","On Hold Reason","Responsible for Delay","Client Responsiveness",
                       "Client Sentiment","Days Inactive","Inactivity Source","Last Milestone",
                       "Risk Level","Risk Owner","Risk Detail","Delay Summary","JIRA Links"]
    else:
        _col_order = ["Flags","RAG","Customer","Project Type","Start Date","Est. Go-Live",
                      "Phase","On Hold Reason","Responsible for Delay","Client Responsiveness",
                      "Client Sentiment","Days Inactive","Inactivity Source","Last Milestone",
                      "Risk Level","Risk Owner","Risk Detail","Delay Summary","JIRA Links"]
    _oh_df = _oh_df[[c for c in _col_order if c in _oh_df.columns]]

    # ✦ = SS syncable (editable) | no mark = derived/read-only
    st.caption("Columns with the edit icon sync back to Smartsheet — edit and export to CSV to update DRS. Greyed columns are derived or read-only.")
    _oh_edited = st.data_editor(
        _oh_df,
        column_config={
            "Flags":                 st.column_config.TextColumn("Flags",                  disabled=True, width="small"),
            "RAG":                   st.column_config.TextColumn("RAG",                    disabled=True, width="small"),
            "Customer":              st.column_config.TextColumn("Customer",                disabled=True, width="medium"),
            "Project Type":          st.column_config.TextColumn("Project Type",            disabled=True, width="medium"),
            "Start Date":            st.column_config.TextColumn("Start Date",              disabled=True, width="small"),
            "Est. Go-Live":          st.column_config.TextColumn("Est. Go-Live",            disabled=True, width="small"),
            "Phase":                 st.column_config.SelectboxColumn("Phase", options=PHASE_OPTIONS, width="medium"),
            "On Hold Reason":        st.column_config.SelectboxColumn("On Hold Reason",  options=_OH_REASON_OPTS,     width="medium"),
            "Days Inactive":         st.column_config.NumberColumn("Days Inactive",         disabled=True, width="small"),
            "Inactivity Source":     st.column_config.TextColumn("Inactivity Source",       disabled=True, width="small"),
            "Last Milestone":        st.column_config.TextColumn("Last Milestone",          disabled=True, width="medium"),
            "Client Responsiveness": st.column_config.SelectboxColumn("Client Responsiveness", options=_OH_RESP_OPTS, width="medium"),
            "Client Sentiment":      st.column_config.SelectboxColumn("Client Sentiment", options=_OH_SENTIMENT_OPTS, width="small"),
            "Risk Level":            st.column_config.SelectboxColumn("Risk Level",       options=_OH_RISK_OPTS,       width="small"),
            "Risk Owner":            st.column_config.SelectboxColumn("Risk Owner",       options=_OH_OWNER_OPTS,      width="small"),
            "Risk Detail":           st.column_config.TextColumn("Risk Detail",           width="large"),
            "Responsible for Delay": st.column_config.SelectboxColumn("Responsible for Delay", options=_OH_DELAY_OPTS, width="medium"),
            "Delay Summary":         st.column_config.TextColumn("Delay Summary",         width="large"),
            "JIRA Links":            st.column_config.TextColumn("JIRA Links",             width="medium",
                                         help="Comma-separated JIRA ticket URLs, e.g. https://zone.atlassian.net/browse/ZPS-123"),
        },
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="oh_edit_table",
    )

    # Export bar
    _oh_sync_cols = ["Customer","Project Type","On Hold Reason","Responsible for Delay","Client Responsiveness","Client Sentiment",
                     "Risk Level","Risk Owner","Risk Detail","Delay Summary","JIRA Links"]
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

st.markdown('<div style="font-size:11px;opacity:.4;text-align:center;margin-top:20px">PS Projects & Tools · Internal use only · Data loaded this session only</div>',unsafe_allow_html=True)
