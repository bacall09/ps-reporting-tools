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
        _prev = st.session_state.get("_mp_va","— My own projects —")
        _di   = _opts.index(_prev) if _prev in _opts else 0
        st.markdown("**My Projects — View as:**")
        _pick = st.selectbox("mp_va",_opts,index=_di,key="mp_va_sel",label_visibility="collapsed")
        st.session_state["_mp_va"] = _pick
        if _pick not in ("— My own projects —",) and not (_pick.startswith("── ") and _pick.endswith(" ──")):
            view_as = _pick

# ── Data ─────────────────────────────────────────────────────────────────────
df_drs = st.session_state.get("df_drs")
df_ns  = st.session_state.get("df_ns")
if df_drs is None:
    st.info("Upload SS DRS Export in the sidebar to see your projects.")
    st.stop()

# Filter to consultant
_vp = [p.strip() for p in view_as.split(",")]
_vv = {view_as.lower(), _vp[0].lower()}
if len(_vp)==2: _vv.add(f"{_vp[1].strip()} {_vp[0]}".lower())
def _mpm(v):
    v=str(v).strip().lower()
    return v in _vv or any(v==nv or v.startswith(nv+" ") or v.endswith(" "+nv) for nv in _vv)
pm_col = df_drs.get("project_manager",pd.Series(dtype=str))
my_drs = df_drs[pm_col.apply(lambda v:_mpm(str(v)))].copy()

if my_drs.empty and view_as==selected and is_manager(selected):
    st.info(f"No projects assigned to {selected} in DRS. Use 'View as' to browse a consultant.")
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
st.markdown('<div class="section-label">Snapshot</div>',unsafe_allow_html=True)

# Phase counts sorted by phase order
_pc = sorted(active["phase"].value_counts().items(),key=lambda x:_pidx(x[0])) if "phase" in active.columns else []

# Going live this week — sorted by date asc
_n7 = today+timedelta(days=7)
gls = (active[active["go_live_date"].notna()&(active["go_live_date"]>=today)&(active["go_live_date"]<=_n7)]
       .sort_values("go_live_date") if "go_live_date" in active.columns else pd.DataFrame())

# In hypercare week 1
_14 = today-timedelta(days=14)
ihc = (active[active["phase"].fillna("").str.lower().str.contains("06|go-live|go live|hypercare",na=False)
              &active["go_live_date"].notna()&(active["go_live_date"]>=_14)]
       if "go_live_date" in active.columns else pd.DataFrame())

# Missing intro email — flag any active non-legacy project in or past onboarding without a date
_leg        = active.get("legacy", pd.Series(False, index=active.index)).astype(bool)
_onb_or_past = active["phase"].fillna("").apply(lambda p: _pidx(p) >= _pidx("00. onboarding") and _pidx(p) >= 0)
_no_intro   = (~active["ms_intro_email"].notna()) if "ms_intro_email" in active.columns else pd.Series(True, index=active.index)
mi          = active[(~_leg) & _onb_or_past & _no_intro] if "ms_intro_email" in active.columns else pd.DataFrame()

c1,c2,c3,c4=st.columns(4)
with c1:
    st.markdown(f'<div class="metric-card"><div class="metric-val">{len(active)}</div><div class="metric-lbl">Active Projects</div></div>',unsafe_allow_html=True)
    for ph,cnt in _pc:
        st.markdown(f'<div style="font-size:11px;opacity:.65;padding:1px 0">{cnt} · {str(ph).split(".")[-1].strip()[:22]}</div>',unsafe_allow_html=True)
with c2:
    col="#E74C3C" if len(gls)>0 else "inherit"
    st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{col}">{len(gls)}</div><div class="metric-lbl">Going live this week</div></div>',unsafe_allow_html=True)
    for _,r in gls.iterrows():
        st.markdown(f'<div style="font-size:11px;opacity:.65;padding:1px 0">{str(r.get("project_name",""))[:28]} · {pd.Timestamp(r["go_live_date"]).strftime("%-d %b")}</div>',unsafe_allow_html=True)
with c3:
    col="#F39C12" if len(ihc)>0 else "inherit"
    st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{col}">{len(ihc)}</div><div class="metric-lbl">In hypercare (week 1)</div></div>',unsafe_allow_html=True)
    for _,r in ihc.iterrows():
        st.markdown(f'<div style="font-size:11px;opacity:.65;padding:1px 0">{str(r.get("project_name",""))[:28]} · {pd.Timestamp(r["go_live_date"]).strftime("%-d %b")}</div>',unsafe_allow_html=True)
with c4:
    col="#E74C3C" if len(mi)>0 else "inherit"
    st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{col}">{len(mi)}</div><div class="metric-lbl">Missing intro email</div></div>',unsafe_allow_html=True)
    if len(mi)>0: st.markdown('<div style="font-size:11px;opacity:.55">Excl. legacy projects</div>',unsafe_allow_html=True)

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Needs Action
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Needs Action</div>',unsafe_allow_html=True)

needs_action = (active[active["_needs"].astype(bool)].sort_values(["_ne","_nw","days_inactive"],ascending=[False,False,False])
                if "_needs" in active.columns and active["_needs"].any() else pd.DataFrame())

if needs_action.empty:
    st.markdown('<span style="font-size:13px;color:#27AE60">✓ No projects currently need attention.</span>',unsafe_allow_html=True)
else:
    st.markdown(f'**{len(needs_action)} project(s)** need attention')
    if "queued_changes" not in st.session_state:
        st.session_state["queued_changes"] = {}

    for _ri,(_,row) in enumerate(needs_action.iterrows()):
        pn        = str(row.get("project_name","—"))
        phase     = str(row.get("phase","—"))
        days_inac = int(row.get("days_inactive",0))
        flist     = row.get("_flags",[]) or []
        tier      = suggest_tier(days_inac) if days_inac>=14 else None
        ne,nw     = int(row.get("_ne",0)),int(row.get("_nw",0))

        _lbl = pn
        if ne: _lbl += f"  ·  {ne} Error{'s' if ne>1 else ''}"
        if nw: _lbl += f"  ·  {nw} Warning{'s' if nw>1 else ''}"
        if days_inac>=14: _lbl += f"  ·  {days_inac}d inactive"

        with st.expander(_lbl,expanded=(ne>0)):
            ci,ca=st.columns([3,2])
            with ci:
                bh=""
                for sev,fk,msg,ed in flist:
                    cls="pf-e" if sev=="error" else "pf-w"
                    bh+=f'<span class="pf {cls}">{"✏️ " if ed else "ℹ️ "}{msg}</span>'
                if bh: st.markdown(bh+"<br>",unsafe_allow_html=True)
                st.markdown(f"**Phase:** {phase}")
                st.markdown(f"**RAG:** {str(row.get('rag','')or'—').strip().upper()[:1] or '—'}")
                if days_inac>=0: st.markdown(f"**Days inactive:** {days_inac}")
                gl=row.get("go_live_date")
                if pd.notna(gl): st.markdown(f"**Go Live:** {pd.Timestamp(gl).strftime('%-d %b %Y')}")
            with ca:
                if tier and tier in TEMPLATES:
                    st.markdown(f"**Re-engagement:** {tier}")
                    if st.button("Draft outreach →",key=f"mp_d_{_ri}",type="primary"):
                        st.session_state["_jump_to_proj"]=pn
                        st.session_state["_jump_tier"]=tier
                        st.switch_page("pages/2_Customer_Reengagement.py")

                ef=[(fk,msg) for _,fk,msg,ed in flist if ed and fk]
                if ef:
                    st.markdown("**Update in DRS:**")
                    # _ri makes every widget key unique across all projects
                    ek=f"mp_{_ri}"
                    _q=st.session_state["queued_changes"].get(pn,{})
                    nv=dict(_q)
                    for fk,lbl in ef:
                        raw=row.get(fk)
                        if fk=="phase":
                            pi2=next((i for i,p in enumerate(PHASE_OPTIONS) if phase.lower().replace(" ","") in p.lower().replace(" ","")),0)
                            v=st.selectbox("Phase",PHASE_OPTIONS,index=pi2,key=f"{ek}_ph")
                            if v!=phase: nv["Project Phase"]=v
                        elif fk=="schedule_health":
                            cur=str(raw or "")
                            si=SCHEDULE_OPTIONS.index(cur) if cur in SCHEDULE_OPTIONS else 0
                            v=st.selectbox("Schedule Health",SCHEDULE_OPTIONS,index=si,key=f"{ek}_sh")
                            if v!=cur: nv["Schedule Health"]=v
                        elif fk in MS_TO_SS:
                            ml=MILESTONE_COLS_MAP.get(fk,lbl)
                            cd=pd.Timestamp(raw).date() if pd.notna(raw) else date.today()
                            v=st.date_input(ml,value=cd,key=f"{ek}_{fk}")
                            if str(v)!=str(cd): nv[MS_TO_SS[fk]]=str(v)

                    b1,b2=st.columns(2)
                    with b1:
                        if st.button("Queue change",key=f"{ek}_sv",use_container_width=True):
                            nv["Project Name"]=pn
                            st.session_state["queued_changes"][pn]=nv
                            st.rerun()
                    with b2:
                        if _q and st.button("Clear",key=f"{ek}_cl",use_container_width=True):
                            st.session_state["queued_changes"].pop(pn,None)
                            st.rerun()
                    if pn in st.session_state["queued_changes"]:
                        st.markdown('<span style="font-size:11px;color:#27AE60">✓ Queued for export</span>',unsafe_allow_html=True)

# ── Export ───────────────────────────────────────────────────────────────────
_qa=st.session_state.get("queued_changes",{})
if _qa:
    st.markdown('<hr class="divider">',unsafe_allow_html=True)
    st.markdown(f'**{len(_qa)} change(s) queued for export**')
    _edf=pd.DataFrame([{"project_name":pn,**flds} for pn,flds in _qa.items()])
    st.dataframe(_edf,use_container_width=True,hide_index=True)
    _buf=io.BytesIO(); _edf.to_csv(_buf,index=False)
    st.download_button("⬇ Export changes to CSV",data=_buf.getvalue(),
                       file_name=f"drs_changes_{date.today().strftime('%Y%m%d')}.csv",
                       mime="text/csv",type="primary")
    if st.button("Clear all queued changes",key="mp_clr_all"):
        st.session_state["queued_changes"]={}; st.rerun()

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — Active Projects
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Active Projects</div>',unsafe_allow_html=True)
if active.empty:
    st.info("No active projects found.")
else:
    def _tr(row):
        fl=row.get("_flags",[]) or []
        fs=" ".join("🔴" if s=="error" else "🟡" for s,_,_m,_e in fl) if fl else "✓"
        di=int(row.get("days_inactive",-1))
        return {"Project":str(row.get("project_name","")),"Phase":str(row.get("phase","—")),
                "RAG":str(row.get("rag","")or"—").strip(),
                "Days Inactive":di if di>=0 else "—",
                "Go Live":pd.Timestamp(row["go_live_date"]).strftime("%-d %b %Y") if pd.notna(row.get("go_live_date")) else "—",
                "Schedule":str(row.get("schedule_health","")or"—").strip(),
                "Last Milestone":str(row.get("last_milestone","—")),"Flags":fs}
    st.dataframe(pd.DataFrame([_tr(r) for _,r in active.iterrows()]),
                 use_container_width=True,hide_index=True)

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
