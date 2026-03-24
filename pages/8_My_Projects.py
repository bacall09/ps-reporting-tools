"""
PS Tools — My Projects
Per-consultant project working list: snapshot metrics, needs-action items,
active projects table, on-hold projects, and a change export flow.
"""
import streamlit as st
import pandas as pd
import io
from datetime import date, timedelta

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN,
    MILESTONE_COLS_MAP, get_role, is_manager,
    get_ff_scope,
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
    .section-label{font-size:11px;font-weight:700;text-transform:uppercase;
                   letter-spacing:.8px;color:#4472C4;margin-bottom:8px}
    .metric-card{border:1px solid rgba(128,128,128,.2);border-radius:8px;padding:16px 20px;margin-bottom:12px}
    .metric-val{font-size:26px;font-weight:700}
    .metric-lbl{font-size:12px;opacity:.6;margin-top:2px}
    .pf{display:inline-block;font-size:11px;font-weight:600;padding:2px 7px;
        border-radius:4px;margin-right:4px;margin-bottom:3px}
    .pf-e{background:rgba(231,76,60,.15);color:#E74C3C}
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
if role == "manager":
    with st.sidebar:
        def _gr(n):
            if n in PS_REGION_OVERRIDE: return PS_REGION_OVERRIDE[n]
            return PS_REGION_MAP.get(EMPLOYEE_LOCATION.get(n,""),"Other")
        _ac = sorted([n for n in CONSULTANT_DROPDOWN
                      if get_role(n) in ("consultant","manager")
                      and EMPLOYEE_ROLES.get(n,{}).get("products")])
        _by = {}
        for n in _ac: _by.setdefault(_gr(n),[]).append(n)
        _opts = ["— My own projects —"]
        for rg in sorted(_by): _opts.append(f"── {rg} ──"); _opts.extend(_by[rg])
        st.markdown("**My Projects — View as:**")
        _pick = st.selectbox("mp_va", _opts, key="mp_va_sel", label_visibility="collapsed")
        # Derive view intent from current selection
        if _pick.startswith("── ") and _pick.endswith(" ──"):
            _va_region = _pick[3:-3].strip()
        elif _pick == "— My own projects —":
            _va_region = None
        else:
            view_as    = _pick
            _va_region = None
else:
    _va_region = None

# ── Data ─────────────────────────────────────────────────────────────────────
# DEBUG — remove after confirming view_as works
with st.sidebar:
    st.caption(f"DEBUG: selected={selected} | view_as={view_as} | _va_region={_va_region} | role={role}")
    if st.session_state.get("df_drs") is not None:
        _pm_sample = st.session_state["df_drs"].get("project_manager", pd.Series(dtype=str)).dropna().unique()[:5].tolist()
        st.caption(f"DRS PM samples: {_pm_sample}")

df_drs = st.session_state.get("df_drs")
df_ns  = st.session_state.get("df_ns")
if df_drs is None:
    st.info("Upload SS DRS Export in the sidebar to see your projects.")
    st.stop()

# Filter to consultant or region
pm_col = df_drs.get("project_manager", pd.Series(dtype=str))

if _va_region and role == "manager":
    # Region view — show all consultants in this region
    _region_consultants = set()
    for n in CONSULTANT_DROPDOWN:
        _nl = EMPLOYEE_LOCATION.get(n, "")
        _nr = PS_REGION_OVERRIDE.get(n, PS_REGION_MAP.get(_nl, "Other"))
        if _nr == _va_region:
            _region_consultants.add(n.lower())
            _vp2 = [p.strip() for p in n.split(",")]
            _region_consultants.add(_vp2[0].lower())
            if len(_vp2) == 2:
                _region_consultants.add(f"{_vp2[1].strip()} {_vp2[0]}".lower())
    my_drs = df_drs[pm_col.apply(lambda v: str(v).strip().lower() in _region_consultants)].copy()
    if my_drs.empty:
        st.info(f"No projects found for the {_va_region} region in DRS.")
        st.stop()
elif view_as == selected and is_manager(selected):
    # Manager viewing their own — show nothing (they have no projects)
    st.info(f"You are logged in as a manager. Use 'View as' to browse a consultant or region.")
    st.stop()
else:
    _vp = [p.strip() for p in view_as.split(",")]
    _vv = {view_as.lower(), _vp[0].lower()}
    if len(_vp) == 2: _vv.add(f"{_vp[1].strip()} {_vp[0]}".lower())
    def _mpm(v):
        v = str(v).strip().lower()
        return v in _vv or any(v == nv or v.startswith(nv + " ") or v.endswith(" " + nv) for nv in _vv)
    my_drs = df_drs[pm_col.apply(lambda v: _mpm(str(v)))].copy()
    if my_drs.empty:
        st.info(f"No projects found for {view_as} in DRS.")
        st.stop()

if df_ns is not None:
    try: my_drs = calc_days_inactive(my_drs,df_ns)
    except Exception: pass

_ioh   = my_drs.get("_on_hold",pd.Series(False,index=my_drs.index)).astype(bool)
on_hold= my_drs[_ioh].copy()
active = my_drs[~_ioh].copy()

# ── Flags ─────────────────────────────────────────────────────────────────────
def _flags(row):
    out=[]; phase=str(row.get("phase","")or"").strip()
    go_live=row.get("go_live_date"); start_dt=row.get("start_date")
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

# ── Header ────────────────────────────────────────────────────────────────────
_dn = view_as.split(",")[1].strip()+" "+view_as.split(",")[0] if "," in view_as else view_as
st.markdown(f"""
<div style='background:#1B2B5E;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;font-family:Manrope,sans-serif;position:relative;overflow:hidden'>
    <div style='position:absolute;right:-40px;top:-40px;width:220px;height:220px;border-radius:50%;background:radial-gradient(circle,rgba(91,141,239,0.15) 0%,transparent 70%);pointer-events:none'></div>
    <div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3ECFB2;margin-bottom:10px;font-family:Manrope,sans-serif'>Professional Services · My Work</div>
    <h1 style='color:#fff;margin:0;font-size:28px;font-weight:800;font-family:Manrope,sans-serif;line-height:1.15'>My Projects — {_dn}</h1>
    <p style='color:rgba(255,255,255,0.6);margin:8px 0 0;font-size:14px;font-family:Manrope,sans-serif;line-height:1.6'>{today.strftime("%A, %B %-d %Y")} · {len(active)} active · {len(on_hold)} on hold</p>
</div>
""", unsafe_allow_html=True)
st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Snapshot
# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2+3 — Active Projects (editable table + export)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Active Projects</div>',unsafe_allow_html=True)

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
    _ns_tm_hrs: dict = {}  # project_id → sum of T&M hours

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
        if _ns_id_col and "hours" in df_ns.columns and "billing_type" in df_ns.columns:
            _tm_mask = df_ns["billing_type"].fillna("").str.strip().str.lower() == "t&m"
            for _pid, _grp in df_ns[_tm_mask].groupby(_ns_id_col):
                _k = _clean_pid(_pid)
                if _k:
                    try:
                        _ns_tm_hrs[_k] = round(float(_grp["hours"].sum() or 0), 2)
                    except Exception:
                        pass

    # ── Build editable dataframe ──────────────────────────────────────────────
    def _to_edit_row(row):
        fl    = row.get("_flags",[]) or []
        needs = "⚠️" if any(sev=="error" for sev,_,_m,_ in fl) else ("⚠️" if any(sev=="warn" for sev,_,_m,_ in fl) else "")
        def _ms(col):
            v = row.get(col)
            return pd.Timestamp(v).strftime("%-d %b %Y") if pd.notna(v) else ""
        def _dt(col):
            v = row.get(col)
            return pd.Timestamp(v).strftime("%-d %b %Y") if pd.notna(v) else ""
        _pn    = str(row.get("project_name","") or "")
        _cust  = _pn.split(" - ")[0].strip() if " - " in _pn else _pn
        # Scope: FF → from DEFAULT_SCOPE table by project_type; T&M → total NS hours logged
        _ptype_raw = str(row.get("project_type", "") or "")
        _bill_raw  = str(row.get("billing_type", "") or "").strip().lower()
        _is_tm     = "t&m" in _bill_raw or "time" in _bill_raw
        _ff_scope  = get_ff_scope(_ptype_raw)
        _pid_key = _clean_pid(row.get("project_id", ""))
        if _is_tm:
            # TODO: replace with actual scoped hours per project (SFDC opportunity or DRS field)
            # Current proxy = total NS hours logged (T&M is uncapped so no scope table exists)
            _scope = round(_ns_tm_hrs.get(_pid_key, 0.0), 2) if _pid_key and _pid_key in _ns_tm_hrs else ""
        elif _ff_scope is not None:
            _scope = float(_ff_scope)
        else:
            _scope = ""
        _htd = round(_ns_htd.get(_pid_key, 0.0), 2) if _pid_key and _pid_key in _ns_htd else ""
        _bal = round(float(_scope) - float(_htd), 2) if _scope != "" and _htd != "" else ""
        return {
            "Needs Action":         needs,
            "Customer":             _cust,
            "Project Type":         str(row.get("project_type","") or ""),
            "Status":               str(row.get("status","") or ""),
            "Phase":                str(row.get("phase","") or ""),
            "Start Date":           _dt("start_date"),
            "Est. Go-Live":         _dt("go_live_date"),
            "Scope Hrs":            _scope,
            "Hours to Date":        _htd,
            "Balance":              _bal,
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

    edit_df = pd.DataFrame([_to_edit_row(r) for _,r in active.iterrows()])

    # ── Column config ─────────────────────────────────────────────────────────
    _ms_cols = ["Intro Email Sent","Config Start","Enablement Session","Session #1","Session #2",
                "UAT Signoff","Prod Cutover","Hypercare Start","Close Out Tasks","Transition to Support"]
    col_cfg = {
        "Needs Action":          st.column_config.TextColumn("",                  disabled=True, width="small"),
        "Customer":              st.column_config.TextColumn("Customer",          disabled=True),
        "Project Type":          st.column_config.TextColumn("Project Type",      disabled=True),
        "Status":                st.column_config.TextColumn("Status",            disabled=True),
        "Phase":                 st.column_config.SelectboxColumn("Phase",        options=PHASE_OPTIONS, width="medium"),
        "Start Date":            st.column_config.TextColumn("Start Date",        disabled=True, width="small"),
        "Est. Go-Live":          st.column_config.TextColumn("Est. Go-Live",      disabled=True, width="small"),
        "Scope Hrs":             st.column_config.NumberColumn("Scope Hrs",         disabled=True, width="small"),
        "Hours to Date":         st.column_config.NumberColumn("Hours to Date",     disabled=True, width="small"),
        "Balance":               st.column_config.NumberColumn("Balance",           disabled=True, width="small"),
        **{c: st.column_config.TextColumn(c, width="small") for c in _ms_cols},
    }

    st.caption("Edit Phase, Schedule Health, or milestone dates directly in the table. Export to CSV to update Smartsheet.")
    st.markdown('<span style="font-size:11.5px;opacity:.6">For a deeper look at data quality issues, use the DRS Health Check page.</span>', unsafe_allow_html=True)
    if st.button("→ Go to DRS Health Check", key="mp_drs_link"):
        st.switch_page("pages/6_DRS_Health_Check.py")

    edited = st.data_editor(
        edit_df,
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
    ex1, ex2 = st.columns([3,1])
    with ex1:
        if changed.any():
            st.markdown(f'<span style="font-size:13px;color:#27AE60;font-weight:600">✓ {changed.sum()} project(s) edited — ready to export</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span style="font-size:12px;opacity:.5">No edits yet — edit cells above then export</span>', unsafe_allow_html=True)
    with ex2:
        _export_df = changed_df if not changed_df.empty else edited
        _buf = io.BytesIO()
        _export_df.to_csv(_buf, index=False)
        st.download_button(
            label="⬇ Export to CSV" if not changed_df.empty else "⬇ Export all",
            data=_buf.getvalue(),
            file_name=f"drs_updates_{date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            type="primary" if not changed_df.empty else "secondary",
            use_container_width=True,
        )

    # ── Re-engagement shortcuts for inactive projects ─────────────────────────
    _inactive_projs = active[active["days_inactive"].fillna(0)>=14].sort_values("days_inactive", ascending=False)
    if not _inactive_projs.empty:
        st.markdown('<hr class="divider">', unsafe_allow_html=True)
        st.markdown('<div class="section-label">Re-engagement shortcuts</div>', unsafe_allow_html=True)
        for _ri,(_,row) in enumerate(_inactive_projs.iterrows()):
            pn        = str(row.get("project_name","—"))
            days_inac = int(row.get("days_inactive",0))
            tier      = suggest_tier(days_inac) if days_inac>=14 else None
            if tier and tier in TEMPLATES:
                _rc1,_rc2,_rc3 = st.columns([3,1,1])
                with _rc1:
                    st.markdown(f'<span style="font-size:13px">{pn}</span> <span style="font-size:11px;opacity:.6">{days_inac}d inactive · {tier}</span>', unsafe_allow_html=True)
                with _rc3:
                    if st.button("Draft outreach →", key=f"mp_re_{_ri}", use_container_width=True):
                        st.session_state["_jump_to_proj"] = pn
                        st.session_state["_jump_tier"]    = tier
                        st.switch_page("pages/2_Customer_Reengagement.py")

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — On Hold
# ══════════════════════════════════════════════════════════════════════════════
with st.expander(f"On Hold ({len(on_hold)} projects)",expanded=False):
    if on_hold.empty:
        st.markdown("No on-hold projects.")
    else:
        st.dataframe(pd.DataFrame([{
            "Project":str(r.get("project_name","")),"Phase":str(r.get("phase","—")),
            "RAG":str(r.get("rag","")or"—").strip(),
            "Go Live":pd.Timestamp(r["go_live_date"]).strftime("%-d %b %Y") if pd.notna(r.get("go_live_date")) else "—",
            "Days Inactive":int(r.get("days_inactive",-1)) if r.get("days_inactive",-1)>=0 else "—",
        } for _,r in on_hold.iterrows()]),use_container_width=True,hide_index=True)

st.markdown('<div style="font-size:11px;opacity:.4;text-align:center;margin-top:20px">PS Reporting Tools · Internal use only · Data loaded this session only</div>',unsafe_allow_html=True)
