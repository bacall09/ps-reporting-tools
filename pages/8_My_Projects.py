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
    .section-label{font-size:11px;font-weight:700;text-transform:uppercase;
                   letter-spacing:.8px;color:#4472C4;margin-bottom:8px}
    .metric-card{border:1px solid rgba(128,128,128,.2);border-radius:8px;padding:16px 20px;margin-bottom:12px}
    .metric-val{font-size:26px;font-weight:700}
    .metric-lbl{font-size:12px;opacity:.6;margin-top:2px}
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
    my_drs = df_drs[pm_col.apply(lambda v: resolve_name(str(v)).lower() in _region_consultants or str(v).strip().lower() in _region_consultants)].copy()
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
_dn = (_va_region + " Team" if _va_region
       else view_as.split(",")[1].strip()+" "+view_as.split(",")[0] if "," in view_as
       else view_as)
st.markdown(f"""
<div style='background:#050D1F;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;font-family:Manrope,sans-serif;position:relative;overflow:hidden'>
    <svg style='position:absolute;right:-40px;top:50%;transform:translateY(-50%);opacity:0.06;width:200px;height:200px;pointer-events:none' viewBox='0 0 1482 1286.25' xmlns='http://www.w3.org/2000/svg'><g fill='#3B9EFF' fill-rule='evenodd'><path d='M975.127,924.953c2.608-2.68,1.744-5.496-.42-7.829l-57.415-61.872c-2.463-2.655-5.025-2.878-8.443-.991-10.398,5.739-19.024,12.314-27.949,19.885-83.252,70.621-197.471,155.494-298.93,195.556-17.993,7.105-35.256,13.178-54.191,17.329-62.148,13.627-131.853,15.491-192.702-5.298-64.93-22.183-113.878-68.722-142.715-130.542-28.647-61.415-22.393-131.406,11.352-189.217,2.598-2.793,1.405-6.055-1.389-8.184-35.341-26.918-40.303-33.439-69.367-65.686-1.449-1.607-4.102-2.401-5.903-1.138-13.105,9.189-23.232,20.534-33.172,32.961-16.499,20.629-29.73,42.605-38.718,67.541-5.127,10.469-8.378,20.486-10.885,32.065-13.633,62.973-7.701,128.685,17.402,188.142,23.839,56.463,65.297,103.638,114.77,139.169,32.418,23.283,66.848,42.548,103.476,58.385,25.142,10.871,50.281,18.994,76.934,25.12,96.392,22.153,188.876,4.496,276.774-38.393,42.916-20.94,83.188-45.685,121.922-73.568,75.733-54.514,154.643-126.72,219.571-193.435ZM1445.252,792.261c-7.628-38.507-22.817-74.472-43.124-107.897-35.582-58.566-85.801-106.77-139.329-149.092-69.784-55.176-145.355-102.407-225.163-141.162-2.165-1.052-4.941.388-5.391,1.627-.426,1.171-.463,3.413.931,4.628,20.341,17.734,39.847,35.55,58.599,55.093,13.286,14.465,26.223,28.012,37.022,44.544,19.784,30.289,35.735,62.168,50.127,95.397,34.512,31.926,64.863,67.358,90.813,106.359,42.427,63.765,57.696,142.663,37.453,217.116-11.436,42.061-34.763,80.507-64.388,112.265-55.859,59.882-133.144,94.711-214.71,99.157-32.507,1.773-64.093-.538-96.013-6.503-28.16-5.262-70.299-23.997-96.538-36.626-2.312-1.112-4.605-.743-6.449.974-12.635,11.76-25.076,22.901-39.051,33.146l-43.32,31.757c-2.68,1.965-2.195,5.562.439,7.808,70.707,60.309,165.779,100.179,259.837,97.033,39.996-1.336,78.686-6.594,117.486-16.111,94.178-23.099,174.952-71.91,236.526-146.957,23.873-29.096,44.355-60.51,59.779-94.956,29.172-65.148,38.357-137.461,24.463-207.601ZM601.099,242.903c-12.268,10.522-48.215,44.405-47.219,60.482.993,16.01,10.781,31.195,25.227,38.155,14.47,6.972,41.303-10.055,53.886-18.311l65.495-42.972c26.305-17.259,52.496-32.716,80.08-47.834l57.464-31.494c20.451-11.209,41.123-19.851,63.235-27.448,35.852-12.318,72.313-18.084,110.322-17.747,29.787.263,58.398,3.408,86.939,11.449,44.037,12.405,82.745,35.987,114.027,69.974,20.347,22.106,37.598,45.332,51.026,71.732,6.962,13.688,13.008,27.156,16.103,42.311,6.48,31.729,12.267,85.992-.676,115.916-6.013,13.902-13.009,26.627-18.289,40.753-.847,2.264-.768,4.767,1.387,6.461l81.366,63.967c2.003,1.574,5.098.298,6.46-1.592,19.285-26.745,34.599-55.578,45.667-86.804,10.617-29.953,15.416-60.246,15.218-92.192-.482-77.938-29.055-152.791-79.976-211.891-67.16-77.946-169.264-137.487-272.877-146.244-33.524-2.834-66.192-1.328-99.421,3.091-82.214,10.934-149.21,45.218-216.385,92.267-48.269,33.807-94.373,69.644-139.062,107.973ZM72.687,567.553c20.03,44.974,54.35,86.652,88.718,121.568,19.447,19.756,38.882,38.258,60.393,55.711l73.052,59.268c30.921,25.086,74.954,56.331,111.096,72.278,11.713,5.168,23.385,8.99,35.917,11.295,12.922,2.375,24.878,1.136,37.309-3.088,18.441-6.266,35.538-14.698,52.671-24.006,1.792-.974,2.85-2.213,3.058-3.936.179-1.483-.47-3.163-1.914-4.548-14.129-13.542-27.174-27.284-42.195-40.056l-78.193-66.48-93.5-82.422c-23.176-20.43-44.471-41.737-65.536-64.239-15.19-16.227-28.591-32.64-40.05-51.639-20.601-34.157-31.396-72.282-30.182-112.398.614-20.279,2.364-39.861,7.45-59.369,8.872-34.031,50.72-76.652,77.451-99.125,3.767-7.04,2.459-14.401,2.885-21.735.884-15.227,3.244-29.908,5.647-44.959,4.285-26.824,22.718-58.984,38.899-80.638,1.348-1.805,1.936-3.535.891-4.937-.951-1.277-2.618-2.49-4.589-2.222-52.436,7.145-104.92,34.806-146.088,67.704-25.632,20.484-48.458,43.456-68.934,69.137-46.339,58.118-62.952,131.49-53.428,204.864,4.697,36.186,14.376,70.75,29.171,103.971ZM1196.886,310.029c-4.882-10.39-12.371-18.773-20.659-26.723-18.771-18.007-40.425-31.674-64.291-42.362-57.569-25.783-110.906-28.064-173.214-22.213-61.067,5.735-111.183,25.069-164.567,54.081-24.678,13.412-48.301,26.866-71.885,42.28l-105.247,68.787c-85.308,55.756-195.138,156.138-256.755,237.876-1.598,2.12-2.206,4.81-.222,6.912l76.342,80.886c1.468,1.556,2.9,1.672,4.715,1.249,1.397-.326,1.99-1.717,2.793-3.377,3.117-6.44,6.665-11.977,11.238-17.864,38.52-49.59,82.099-94.54,130.222-135.261,40.87-34.583,82.783-67.442,126.68-98.902,83.71-59.991,188.529-115.793,291.15-127.921,23.653-2.795,46.328-.575,69.656,3.405,27.197,4.641,52.661,12.543,78.69,21.347l38.004,12.855c13.849,4.685,27.221-3.226,30.503-17.755,2.725-12.064,2.293-25.708-3.154-37.301Z'/></g></svg>
    <div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3B9EFF;margin-bottom:10px;font-family:Manrope,sans-serif'>Professional Services · My Work</div>
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
    _days = int(row.get("days_inactive", -1) or -1)
    _leg  = str(row.get("legacy","")).strip().lower() in ("true","yes","1")
    _no_i = not pd.notna(row.get("ms_intro_email")) or str(row.get("ms_intro_email","")).strip() in ("","nan","None","NaT")
    if not _leg and _no_i:
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
        return pd.Timestamp(v).strftime("%Y-%m-%d") if pd.notna(v) else ""


    def _dt(col):
        v = row.get(col)
        return pd.Timestamp(v).strftime("%Y-%m-%d") if pd.notna(v) else ""
    _pn    = str(row.get("project_name","") or "")
    _cust  = _pn.split(" - ")[0].strip() if " - " in _pn else _pn
    # Scope: FF → from DEFAULT_SCOPE table by project_type; T&M → total NS hours logged
    _ptype_raw = str(row.get("project_type", "") or "")
    _bill_raw  = str(row.get("billing_type", "") or "").strip().lower()
    _pid_key   = _clean_pid(row.get("project_id", ""))
    _is_tm     = ("t&m" in _bill_raw or "time" in _bill_raw
                  or _pid_key in _ns_tm_pids)  # confirmed T&M from NS
    _ff_scope  = get_ff_scope(_ptype_raw)
    if _is_tm:
        # T&M scope from NS Time Detail "T&M Scope" column (max per project_id)
        # Falls back to hours sum if column not present in export
        _scope = round(_ns_tm_hrs.get(_pid_key, 0.0), 2) if _pid_key and _pid_key in _ns_tm_hrs else ""
    elif _ff_scope is not None:
        _scope = float(_ff_scope)
    else:
        _scope = ""
    _htd = round(_ns_htd.get(_pid_key, 0.0), 2) if _pid_key and _pid_key in _ns_htd else ""
    _bal = round(float(_scope) - float(_htd), 2) if _scope != "" and _htd != "" else ""
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
        "Engagement":           _engagement_flag(row),
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
    "Status":                st.column_config.TextColumn("Status",            disabled=True),
    "Phase":                 st.column_config.SelectboxColumn("Phase",        options=PHASE_OPTIONS, width="medium"),
    "Start Date":            st.column_config.TextColumn("Start Date",        disabled=True, width="small"),
    "Est. Go-Live":          st.column_config.TextColumn("Est. Go-Live",      disabled=True, width="small"),
    "Scope Hrs":             st.column_config.NumberColumn("Scope Hrs",         disabled=True, width="small"),
    "Hours to Date":         st.column_config.NumberColumn("Hours to Date",     disabled=True, width="small"),
    "Balance":               st.column_config.TextColumn("Balance",            disabled=True, width="small"),
    "Engagement":            st.column_config.TextColumn("Engagement",         disabled=True, width="medium"),
    "_bal_flag":             None,
    **{c: st.column_config.TextColumn(c, width="small") for c in _ms_cols},
}

st.caption("Edit Phase, Schedule Health, or milestone dates directly in the table. Export to CSV to update Smartsheet.")
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


st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ── On Hold field value maps ──────────────────────────────────────────────────
_OH_REASON_OPTS   = ["—","Zone Product Dependency","Zone Program Dependency",
                     "NetSuite Dependency","Client Requested","Client Unresponsive"]
_OH_RESP_OPTS     = ["—","Highly Engaged","Neutral","Not Responsive"]
_OH_SENTIMENT_OPTS= ["—","Positive","Neutral","Concerned"]
_OH_RISK_OPTS     = ["—","Low","Medium","High","Escalated"]
_OH_OWNER_OPTS    = ["—","Client","Product","PS","Sales","Marketing","Support","3rd Party","N/A"]
_OH_DELAY_OPTS    = ["—","Zone","Client","3rd Party"]

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
            "Project":               str(r.get("project_name", "")),
            "Phase":                 str(r.get("phase", "—")),
            "On Hold Reason":        str(r.get("on_hold_reason","") or "—"),
            "Days Inactive":         int(r.get("days_inactive", -1)) if r.get("days_inactive", -1) >= 0 else "—",
            "Inactivity Source":     str(r.get("_inactivity_source","") or "—"),
            "Last Milestone":        str(r.get("last_milestone","") or "—"),
            "Client Responsiveness": str(r.get("client_responsiveness","") or "—"),
            "Client Sentiment":      str(r.get("client_sentiment","") or "—"),
            "Risk Level":            _risk_emoji(r.get("risk_level")),
            "Risk Owner":            str(r.get("risk_owner","") or "—"),
            "Risk Detail":           str(r.get("risk_detail","") or "—"),
            "Responsible for Delay": str(r.get("responsible_for_delay","") or "—"),
            "Delay Summary":         _ds,
        }
        if _va_region:
            _oh_row["Consultant"] = str(r.get("project_manager", "") or "")
        _oh_rows.append(_oh_row)

    _oh_df = pd.DataFrame(_oh_rows)

    # Column order — insert Consultant after RAG if region view
    if "Consultant" in _oh_df.columns:
        _col_order = ["RAG","Project","Consultant","Phase","On Hold Reason","Days Inactive",
                      "Inactivity Source","Last Milestone","Client Responsiveness",
                      "Client Sentiment","Risk Level","Risk Owner","Risk Detail",
                      "Responsible for Delay","Delay Summary"]
    else:
        _col_order = ["RAG","Project","Phase","On Hold Reason","Days Inactive",
                      "Inactivity Source","Last Milestone","Client Responsiveness",
                      "Client Sentiment","Risk Level","Risk Owner","Risk Detail",
                      "Responsible for Delay","Delay Summary"]
    _oh_df = _oh_df[[c for c in _col_order if c in _oh_df.columns]]

    st.dataframe(
        _oh_df,
        column_config={
            "RAG":                   st.column_config.TextColumn("RAG",                   width="small"),
            "Project":               st.column_config.TextColumn("Project",               width="medium"),
            "Phase":                 st.column_config.TextColumn("Phase",                 width="medium"),
            "On Hold Reason":        st.column_config.SelectboxColumn("On Hold Reason",   options=_OH_REASON_OPTS,    width="medium"),
            "Days Inactive":         st.column_config.NumberColumn("Days Inactive",        width="small"),
            "Inactivity Source":     st.column_config.TextColumn("Inactivity Source",     width="small"),
            "Last Milestone":        st.column_config.TextColumn("Last Milestone",        width="medium"),
            "Client Responsiveness": st.column_config.SelectboxColumn("Client Responsiveness", options=_OH_RESP_OPTS, width="medium"),
            "Client Sentiment":      st.column_config.SelectboxColumn("Client Sentiment", options=_OH_SENTIMENT_OPTS, width="small"),
            "Risk Level":            st.column_config.TextColumn("Risk Level",            width="small"),
            "Risk Owner":            st.column_config.SelectboxColumn("Risk Owner",       options=_OH_OWNER_OPTS,     width="small"),
            "Risk Detail":           st.column_config.TextColumn("Risk Detail",           width="large"),
            "Responsible for Delay": st.column_config.SelectboxColumn("Responsible for Delay", options=_OH_DELAY_OPTS, width="medium"),
            "Delay Summary":         st.column_config.TextColumn("Delay Summary",         width="large"),
        },
        use_container_width=True,
        hide_index=True,
    )
    st.caption("On Hold Reason, Client Responsiveness, Client Sentiment, Risk Owner, and Responsible for Delay are editable — changes export to CSV for DRS sync.")

st.markdown('<div style="font-size:11px;opacity:.4;text-align:center;margin-top:20px">PS Reporting Tools · Internal use only · Data loaded this session only</div>',unsafe_allow_html=True)
