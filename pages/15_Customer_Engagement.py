"""
pages/15_Customer_Engagement.py  v4
Redesigned: journey rail + compose/preview two-panel layout.
Theme rules (from RUNBOOK):
  - Hero: hardcoded #050D1F (brand element, never switches)
  - Cards/containers: NO background property — inherit from page
  - Pills: rgba() backgrounds + prefers-color-scheme dark selectors for text
  - CSS vars in <style> blocks unreliable — use inline styles or hex+media pairs
  - Email preview ONLY: hardcoded white (#ffffff) — represents a real email
"""
import streamlit as st
import pandas as pd
import datetime
import re
from rapidfuzz import fuzz

st.session_state["current_page"] = "Customer Engagement"

# ── Hero banner — always dark navy per runbook ────────────────────────────────
st.markdown("""
<div style='background:#050D1F;padding:20px 28px 16px;border-radius:0 0 8px 8px;margin-bottom:20px'>
  <span style='color:#4472C4;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase'>Customer Engagement</span>
  <h2 style='color:#ffffff;margin:4px 0 0;font-size:22px;font-weight:600;letter-spacing:-0.3px'>Lifecycle Email Composer</h2>
</div>
""", unsafe_allow_html=True)

# ── CSS — runbook rules applied throughout ─────────────────────────────────────
# Cards: border + padding only, NO background
# Pills: rgba backgrounds, dark-mode text via prefers-color-scheme + [data-theme]
# Preview pane: hardcoded white (exception — represents real email)
st.markdown("""<style>
.ce-label{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.9px;color:#4472C4;margin:0 0 6px}

/* Cards — no background, inherits page theme */
.ce-card{border:1px solid rgba(128,128,128,.22);border-radius:8px;padding:14px 18px;margin-bottom:12px;color:inherit}
.ce-well{border:1px solid rgba(128,128,128,.18);border-radius:6px;padding:10px 14px;margin-bottom:8px;color:inherit}

/* Journey rail — no background on items, just borders */
.journey-rail{display:flex;gap:0;border:1px solid rgba(128,128,128,.2);border-radius:8px;overflow:hidden;margin-bottom:16px}
.stage-item{flex:1;padding:10px 10px 8px;border-right:1px solid rgba(128,128,128,.15);cursor:pointer;color:inherit;min-width:0}
.stage-item:last-child{border-right:none}
.stage-num{font-size:10px;font-weight:700;letter-spacing:.4px;margin-bottom:2px;color:#4472C4}
.stage-lbl{font-size:12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.stage-sub{font-size:10px;margin-top:2px;opacity:.65;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.stage-item.done .stage-num{color:#16a34a}
.stage-item.done .stage-lbl{color:#16a34a}
.stage-item.active{border-bottom:3px solid #4472C4}
.stage-item.active .stage-num{color:#4472C4}
.stage-item.active .stage-lbl{color:#4472C4}
.stage-item.locked{opacity:.4}

/* Pills — rgba so they work on any bg; text flips for dark mode */
.pill-ok{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:rgba(34,197,94,.14);color:#15803d}
.pill-warn{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:rgba(239,68,68,.14);color:#dc2626}
.pill-info{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:rgba(68,114,196,.14);color:#4472C4}
.pill-due{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;background:rgba(234,179,8,.14);color:#854d0e}
@media(prefers-color-scheme:dark){
  .pill-ok{color:#7ed4a4} .pill-warn{color:#fca5a5}
  .pill-info{color:#93b4e8} .pill-due{color:#fcd34d}
}
.stApp[data-theme="dark"] .pill-ok{color:#7ed4a4}
.stApp[data-theme="dark"] .pill-warn{color:#fca5a5}
.stApp[data-theme="dark"] .pill-info{color:#93b4e8}
.stApp[data-theme="dark"] .pill-due{color:#fcd34d}

/* Pre-send checks */
.check-row{display:flex;align-items:center;gap:8px;font-size:12px;padding:3px 0}
.check-ok{color:#16a34a}.check-bad{color:#dc2626}
@media(prefers-color-scheme:dark){.check-ok{color:#7ed4a4}.check-bad{color:#fca5a5}}
.stApp[data-theme="dark"] .check-ok{color:#7ed4a4}
.stApp[data-theme="dark"] .check-bad{color:#fca5a5}

/* Tip bar — no background */
.ce-tip{border-left:3px solid #4472C4;border-radius:0 6px 6px 0;padding:8px 12px;font-size:12px;margin-bottom:10px;color:inherit}

/* Send log */
.log-row{border-bottom:1px solid rgba(128,128,128,.15);padding:6px 0;font-size:12px}

/* Email preview — EXCEPTION: hardcoded white (represents a real email) */
.email-preview-wrap{background:#ffffff;border-radius:8px;border:1px solid #e2e8f0;overflow:hidden}
.email-preview-chrome{background:#f8fafc;border-bottom:1px solid #e2e8f0;padding:8px 12px;display:flex;align-items:center;gap:6px}
.chrome-dot{width:8px;height:8px;border-radius:50%;background:#e2e8f0;display:inline-block}
.email-preview-header{padding:10px 16px;border-bottom:1px solid #e2e8f0}
.email-meta-row{display:flex;gap:8px;font-size:11px;padding:2px 0;color:#64748b}
.email-meta-lbl{width:36px;color:#94a3b8;flex-shrink:0}
.email-preview-body{padding:16px;font-size:13px;color:#0f172a;line-height:1.65;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}
.email-section-hdr{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#4472C4;margin:12px 0 4px}
.email-ph{display:inline-block;background:rgba(220,38,38,.1);color:#dc2626;border:1px dashed rgba(220,38,38,.4);border-radius:3px;padding:0 4px;font-size:11px;font-family:monospace}
.email-link{display:inline-block;background:rgba(68,114,196,.12);color:#1e40af;border-radius:3px;padding:0 5px;font-size:11px}
.email-footer{font-size:11px;color:#64748b;margin-top:14px;padding-top:10px;border-top:1px solid #e2e8f0}
.ph-bar{padding:8px 14px;background:rgba(220,38,38,.06);border-top:1px solid rgba(220,38,38,.15);font-size:11px;color:#dc2626;display:flex;align-items:center;gap:6px}
</style>""", unsafe_allow_html=True)

# ── Auth ──────────────────────────────────────────────────────────────────────
_logged_in = st.session_state.get("consultant_name","")
if not _logged_in:
    st.warning("Please log in via the Home page.")
    st.stop()

from shared.constants import get_role, is_manager, name_matches
_role   = get_role(_logged_in)
_is_mgr = is_manager(_logged_in) or _role in ("manager_only","reporting_only")

_df_drs = st.session_state.get("df_drs")
if _df_drs is None or _df_drs.empty:
    st.info("Load Smartsheet DRS data on the Home page.")
    st.stop()

try:
    from shared.template_engine import (
        get_welcome_template, list_welcome_templates,
        get_post_session_templates, get_lifecycle_template,
        list_lifecycle_templates, build_auto_context, render_template,
        execute_send, build_ss_writeback, mark_ss_writeback_done,
        get_session_send_log, project_needs_welcome_email, _welcome_library,
        SS_GO_LIVE_DATE, SS_PROD_CUTOVER,
    )
except ImportError as e:
    st.error(f"Template engine not found: {e}")
    st.stop()

try:
    from shared.smartsheet_api import write_row_updates, ss_available
    _ss_ok = ss_available()
except ImportError:
    _ss_ok = False

# ── SFDC helpers (ported from reengagement page) ──────────────────────────────
_SFDC_COL_MAP = {
    "18 digit opportunity id":"opportunity_id","first name":"first_name",
    "last name":"last_name","primary title":"title","title":"title",
    "email":"email","contact email":"email","email address":"email",
    "opportunity owner":"account_manager","account owner":"account_manager",
    "owner":"account_manager","account name":"account","account":"account",
    "opportunity name":"opportunity","opportunity":"opportunity",
    "contact name":"contact_name","full name":"contact_name","name":"contact_name",
    "close date":"close_date","closed date":"close_date",
    "implementation contact":"impl_contact_flag","contact roles":"contact_roles",
    "opportunity owner email":"account_manager_email","owner email":"account_manager_email",
    "primary":"is_primary","is primary":"is_primary",
}
_PROD_KW = ["Capture","Approvals","Reconcile","PSP","Payments","SFTP",
            "E-Invoicing","eInvoicing","CC","ZoneCapture","ZoneApprovals","ZoneReconcile"]

def _clean_acct(text):
    t = str(text).lower()
    for s in ["ltd","limited","inc","llc","plc","gmbh","the ","- za -","& co","co."]:
        t = t.replace(s," ")
    return re.sub(r"[^a-z0-9 ]"," ",t).split()

def _prod_hints(text):
    t = str(text).lower()
    return {k for k in _PROD_KW if k.lower() in t}

def _fuzzy_sfdc(df_sfdc, proj_name, acct_name):
    if df_sfdc is None or df_sfdc.empty: return pd.DataFrame(), None
    df = df_sfdc.copy()
    cm = {c.lower().strip():c for c in df.columns}
    opp = cm.get("opportunity"); acc = cm.get("account"); oid = cm.get("opportunity_id")
    if opp:
        exact = df[df[opp].astype(str).str.lower().str.strip()==str(proj_name).lower().strip()]
        if not exact.empty: return exact, "Exact match"
    dw = set(_clean_acct(acct_name)); dp = _prod_hints(proj_name)
    best = 0; best_id = None; best_nm = None
    for _,row in df.iterrows():
        sa = str(row.get(acc or "account","")); so = str(row.get(opp or "opportunity",""))
        sw = set(_clean_acct(sa))
        ws = len(dw&sw)/max(len(dw),1)*100
        fs = fuzz.token_set_ratio(" ".join(dw)," ".join(sw))
        score = max(ws,fs*0.7) + (30 if bool(dp&_prod_hints(so)) else 0)
        if score > best:
            best=score; best_id=row.get(oid or "opportunity_id") if oid else None
            best_nm=row.get(opp or "opportunity") if opp else None
    lbl = f"Fuzzy match ({int(best)}%)"
    if best >= 60:
        if best_id is not None and oid:
            r = df[df[oid]==best_id]
            if not r.empty: return r, lbl
        if best_nm is not None and opp:
            r = df[df[opp]==best_nm]
            if not r.empty: return r, lbl
    if acc:
        df["_sc"] = df[acc].apply(lambda x: fuzz.token_set_ratio(str(acct_name).lower(),str(x).lower()))
        top = df[df["_sc"]>=75].sort_values("_sc",ascending=False)
        df.drop(columns=["_sc"],inplace=True)
        if not top.empty: return top,"Account match"
    return pd.DataFrame(), None

# ── General helpers ───────────────────────────────────────────────────────────
def _row_dict(row):
    return {k:(None if pd.isna(v) else v) for k,v in row.items()}

def _flip_name(n):
    if "," in n:
        p=[x.strip() for x in n.split(",",1)]
        return f"{p[1]} {p[0]}"
    return n

def _consultant_email(name):
    try:
        from shared.constants import EMPLOYEE_ROLES
        em = EMPLOYEE_ROLES.get(name,{}).get("email","")
        if em: return em
    except Exception: pass
    slug = re.sub(r"[^a-z0-9.]","",_flip_name(name).lower().replace(" ","."))
    return f"{slug}@zoneandco.com"

def _missing_phs(text):
    return list(set(re.findall(r"\{[A-Z_]+\}", text)))

_PMAP = {
    "zoneapp: capture":"ZoneCapture","zoneapp: approvals":"ZoneApprovals",
    "zoneapp: reconcile":"ZoneReconcile","zoneapp: reconcile 2.0":"ZoneReconcile_BankConnect",
    "zoneapp: reconcile with bank connectivity":"ZoneReconcile_BankConnect",
    "zoneapp: reconcile with cc import":"ZoneReconcile_CCImport",
    "zoneapp: e-invoicing":"EInvoicing",
    "zoneapp: capture & e-invoicing":"ZoneCapture_EInvoicing",
    "zoneapp: capture and e-invoicing":"ZoneCapture_EInvoicing",
    "zoneapp: capture & approvals":"ZoneCapture_ZoneApprovals",
    "zoneapp: capture and approvals":"ZoneCapture_ZoneApprovals",
    "zoneapp: capture & reconcile":"ZoneCapture_ZoneReconcile",
    "zoneapp: capture and reconcile":"ZoneCapture_ZoneReconcile",
    "zonecapture":"ZoneCapture","zoneapprovals":"ZoneApprovals",
    "zonereconcile":"ZoneReconcile","zonereconcile with bank connectivity":"ZoneReconcile_BankConnect",
    "zonereconcile 2.0":"ZoneReconcile_BankConnect",
    "zonereconcile with cc import":"ZoneReconcile_CCImport",
    "e-invoicing":"EInvoicing","zone e-invoicing":"EInvoicing",
    "zonecapture with e-invoicing":"ZoneCapture_EInvoicing",
    "zonecapture and zoneapprovals":"ZoneCapture_ZoneApprovals",
    "zonecapture and zonereconcile":"ZoneCapture_ZoneReconcile",
}
def _sku(p): return _PMAP.get(str(p).strip().lower())
def _ps_key(p):
    p=str(p).lower()
    if "capture" in p: return "capture"
    if "approv" in p: return "approvals"
    if "reconcile" in p: return "reconcile"
    return None

# ── Journey stage definitions ─────────────────────────────────────────────────
# Maps lifecycle template IDs to their journey position + milestone field
_JOURNEY_STAGES = [
    {"id":"welcome",           "label":"Welcome",           "ss_field":"Intro. Email Sent",     "tab":"Welcome"},
    {"id":"post_session_1",    "label":"Post-Session #1",   "ss_field":"Enablement Session",     "tab":"Post-Session"},
    {"id":"post_session_2",    "label":"Post-Session #2",   "ss_field":"Session #1",             "tab":"Post-Session"},
    {"id":"uat_signoff",       "label":"UAT Sign-Off",      "ss_field":"UAT Signoff",            "tab":"Lifecycle"},
    {"id":"go_live_hypercare_kickoff","label":"Go-Live",    "ss_field":"Prod Cutover",           "tab":"Lifecycle"},
    {"id":"hypercare_closure", "label":"Hypercare close",   "ss_field":"Transition to Support",  "tab":"Lifecycle"},
]

def _stage_status(stage_id, drs_row):
    """Return 'done'|'active'|'pending'|'locked' for a journey stage."""
    ms_map = {
        "welcome":                  ("ms_intro_email","Intro. Email Sent"),
        "post_session_1":           ("ms_enablement","Enablement Session"),
        "post_session_2":           ("ms_session1","Session #1"),
        "uat_signoff":              ("ms_uat_signoff","UAT Signoff"),
        "go_live_hypercare_kickoff":("ms_prod_cutover","Prod Cutover"),
        "hypercare_closure":        ("ms_transition","Transition to Support"),
    }
    col, _ = ms_map.get(stage_id, (None, None))
    if not col or not drs_row: return "pending"
    v = drs_row.get(col) or drs_row.get(_)
    return "done" if v and str(v).strip() not in ("","None","nan","NaT") else "pending"

def _build_journey_html(drs_row, active_tab):
    """Render journey rail as theme-aware HTML (no background on items)."""
    parts = ['<div class="journey-rail">']
    statuses = [_stage_status(s["id"], drs_row) for s in _JOURNEY_STAGES]
    # Find first pending = active
    first_pending = next((i for i,st in enumerate(statuses) if st=="pending"), len(statuses)-1)
    for i, stage in enumerate(_JOURNEY_STAGES):
        st_class = statuses[i]
        if i == first_pending and st_class == "pending":
            st_class = "active"
        elif i > first_pending and statuses[i] == "pending":
            st_class = "locked"
        num = f"0{i+1}" if i < 9 else str(i+1)
        if st_class == "done":
            icon = "✓"
        elif st_class == "active":
            icon = "▼"
        else:
            icon = ""
        parts.append(
            f'<div class="stage-item {st_class}">'
            f'<div class="stage-num">{icon} {num}</div>'
            f'<div class="stage-lbl">{stage["label"]}</div>'
            f'</div>'
        )
    parts.append("</div>")
    return "\n".join(parts)

# ── Writeback ─────────────────────────────────────────────────────────────────
def _do_write(project_id, ss_field, date_val, drs_row) -> bool:
    if not _ss_ok or not ss_field: return False
    ss_row_id = drs_row.get("_ss_row_id") or drs_row.get("ss_row_id") if drs_row else None
    if not ss_row_id:
        st.warning("⚠️ Writeback requires DRS loaded via **Sync SS DRS data** on Home page.")
        return False
    fields = [ss_field] if isinstance(ss_field, str) else ss_field
    wrote = False
    proj_name = str(drs_row.get("project_name", project_id)) if drs_row else project_id
    for f in fields:
        try:
            p = build_ss_writeback(project_id, f, date_val, current_drs_row=drs_row)
            if p["skipped"]:
                st.info(f"ℹ️ **{f}** already set — not overwritten.")
                continue
            if not p["fields"]: continue
            ok, errs = write_row_updates([{"_ss_row_id":ss_row_id,"project_name":proj_name,"changes":{f:date_val.isoformat()}}])
            if ok:
                st.success(f"✓ Smartsheet: **{f}** → {date_val.strftime('%d %b %Y')}")
                wrote = True
            for e in (errs or []): st.warning(f"Writeback error: {e}")
        except Exception as ex:
            st.warning(f"Writeback failed for **{f}**: {ex}")
    return wrote

# ── Email preview HTML renderer ───────────────────────────────────────────────
# EXCEPTION: hardcoded white per runbook — this represents a real email
def _email_preview_html(subject, body, to_email, cc_email):
    """Render email as a white Gmail-style preview. Highlights unfilled placeholders."""
    def _htmlify(text):
        out = []
        for line in text.split("\n"):
            s = line.strip()
            if not s:
                out.append('<div style="margin:5px 0"></div>')
            elif s == s.upper() and len(s) > 3 and not s.startswith("•"):
                out.append(f'<div class="email-section-hdr">{s}</div>')
            elif s.startswith("•") or (s.startswith("-") and len(s)>2):
                out.append(f'<div style="margin:2px 0 2px 14px">{s}</div>')
            elif s.endswith(":") and len(s)<60 and s[0].isupper():
                out.append(f'<div style="font-weight:600;margin-top:8px">{s}</div>')
            else:
                out.append(f'<div>{s}</div>')
        return "\n".join(out)

    def _highlight(html):
        # Highlight {PLACEHOLDER} tags in red
        return re.sub(r'\{([A-Z_]+)\}',
                      r'<span class="email-ph">{\1}</span>', html)

    body_html = _highlight(_htmlify(body))
    missing = _missing_phs(body + subject)
    ph_bar = ""
    if missing:
        ph_bar = f'<div class="ph-bar">⚠ {len(missing)} placeholder(s) empty · {", ".join(missing)} — fill in the form to enable Send.</div>'

    return f"""
<div class="email-preview-wrap">
  <div class="email-preview-chrome">
    <span class="chrome-dot"></span><span class="chrome-dot"></span><span class="chrome-dot"></span>
    <span style="font-size:10px;color:#94a3b8;margin-left:6px">Gmail preview</span>
  </div>
  <div class="email-preview-header">
    <div class="email-meta-row"><span class="email-meta-lbl">TO</span>{to_email or '—'}</div>
    <div class="email-meta-row"><span class="email-meta-lbl">CC</span>{cc_email or '—'}</div>
    <div class="email-meta-row"><span class="email-meta-lbl">SUBJ</span><strong style="color:#0f172a">{subject}</strong></div>
  </div>
  <div class="email-preview-body">{body_html}</div>
  {ph_bar}
</div>
"""

# ── Build DRS dataframe ───────────────────────────────────────────────────────
df_all = _df_drs.copy()
cm = {c.lower().strip():c for c in df_all.columns}

name_col   = cm.get("project_name")    or cm.get("project name")
cust_col   = cm.get("customer")
prod_col   = cm.get("project_type")    or cm.get("project type") or cm.get("product")
id_col     = cm.get("project_id")      or cm.get("project id")
status_col = cm.get("status")
pm_col     = cm.get("project_manager") or cm.get("project manager")
intro_col  = (cm.get("intro. email sent") or cm.get("intro email sent")
              or cm.get("ms_intro_email") or cm.get("intro_email_sent"))
legacy_col = cm.get("legacy")
ss_rid_col = cm.get("_ss_row_id") or cm.get("ss_row_id") or cm.get("row_id")

# Build row ID lookup
_ss_row_id_map: dict = {}
if ss_rid_col and id_col:
    for _,r in _df_drs.iterrows():
        pid = str(r.get(id_col,"")).strip()
        rid = r.get(ss_rid_col)
        if pid and rid: _ss_row_id_map[pid] = rid

# Role filter
if not _is_mgr and pm_col:
    df_all = df_all[df_all[pm_col].apply(lambda x: name_matches(str(x),_logged_in))]

# View as (managers)
if _is_mgr and pm_col:
    _browse = st.session_state.get("_browse_passthrough") or st.session_state.get("home_browse","")
    if _browse and _browse not in ("— My own view —","— Select —","👥 All team",""):
        if _browse.startswith("── ") and _browse.endswith(" ──"):
            try:
                from shared.constants import get_region_consultants
                from shared.config import EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE, ACTIVE_EMPLOYEES
                _rc = get_region_consultants(_browse[3:-3].strip(),EMPLOYEE_LOCATION,PS_REGION_MAP,PS_REGION_OVERRIDE,ACTIVE_EMPLOYEES)
                _f = df_all[df_all[pm_col].astype(str).str.strip().str.lower().isin([n.lower() for n in _rc])]
                if not _f.empty: df_all = _f
            except Exception: pass
        else:
            _f = df_all[df_all[pm_col].apply(lambda x: name_matches(str(x),_browse))]
            if not _f.empty: df_all = _f

# Active only + exclude legacy
if status_col:
    df_all = df_all[~df_all[status_col].astype(str).str.lower().isin(["closed","cancelled","complete","completed"])]
if legacy_col:
    df_all = df_all[~df_all[legacy_col].astype(str).str.strip().str.lower().isin(["yes","y","true","1"])]

df_all = df_all.reset_index(drop=True)
if df_all.empty:
    st.info("No active projects found.")
    st.stop()

# Welcome flag column
def _needs_intro(row):
    if not intro_col: return False
    v = row.get(intro_col,"")
    return v is None or str(v).strip() in ("","None","nan","NaT")
df_all["_needs_welcome"] = df_all.apply(_needs_intro,axis=1)

# ═══════════════════════════════════════════════════════════════════════════════
# TOP ROW: project picker + recipient (full width, compact)
# ═══════════════════════════════════════════════════════════════════════════════
top_left, top_right = st.columns([2,1], gap="medium")

with top_left:
    # Filter
    df = df_all.copy()
    _active_tab = st.session_state.get("_ce_tab","Welcome")
    if intro_col and _active_tab == "Welcome":
        if st.checkbox("Only projects needing Welcome email", value=True, key="ce_wf"):
            _fil = df[df["_needs_welcome"]]
            if not _fil.empty: df = _fil.reset_index(drop=True)
            else: st.success("All projects have a Welcome email on record.")

    if df.empty:
        st.info("No projects match this filter.")
        st.stop()

    def _plabel(row):
        c = str(row.get(cust_col,"")) if cust_col else ""
        n = str(row.get(name_col,"")) if name_col else ""
        p = str(row.get(prod_col,"")) if prod_col else ""
        return f"{c}  —  {n}" + (f"  · {p}" if p and p not in ("nan","None") else "")

    df["_sid"] = df.apply(lambda r: str(r.get(id_col,r.name)) if id_col else str(r.name), axis=1)
    sids = df["_sid"].tolist()
    sid_map = {s:i for i,s in enumerate(sids)}
    _prev = st.session_state.get("_ce_proj_sid")
    _def  = _prev if _prev in sids else sids[0]

    selected_sid = st.selectbox(
        "Project", options=sids,
        format_func=lambda s: _plabel(df.iloc[sid_map[s]]),
        index=sids.index(_def),
        label_visibility="collapsed", key="ce_proj",
    )
    st.session_state["_ce_proj_sid"] = selected_sid

    if st.session_state.get("_last_proj") != selected_sid:
        st.session_state["_last_proj"] = selected_sid
        for k in list(st.session_state.keys()):
            if any(k.startswith(p) for p in ["ce_to","ce_cn","ce_cc","w_","s_","l_","sess_","lc_","welcome_"]):
                del st.session_state[k]

    sel_idx     = sid_map[selected_sid]
    sel         = _row_dict(df.iloc[sel_idx])
    customer    = sel.get(cust_col,"") if cust_col else ""
    product_raw = sel.get(prod_col,"") if prod_col else ""
    project_id  = str(sel.get(id_col,str(sel_idx))) if id_col else str(sel_idx)
    if "_ss_row_id" not in sel and project_id in _ss_row_id_map:
        sel["_ss_row_id"] = _ss_row_id_map[project_id]

    # Project summary inline
    iv = sel.get(intro_col,"") if intro_col else ""
    _intro_done = iv and str(iv).strip() not in ("","None","nan","NaT")
    st.markdown(
        f'<div style="font-size:12px;margin-top:4px">'
        f'<b>{sel.get(name_col,"")}</b>&nbsp;&nbsp;'
        f'<span style="opacity:.6">{customer}</span>&nbsp;&nbsp;'
        f'<span style="opacity:.6">{product_raw}</span>&nbsp;&nbsp;'
        + (f'<span class="pill-ok">Welcome sent {iv}</span>' if _intro_done else '<span class="pill-warn">Welcome pending</span>')
        + '</div>', unsafe_allow_html=True,
    )

with top_right:
    # SFDC lookup
    df_sfdc = st.session_state.get("df_sfdc")
    sfdc_email = ""; sfdc_cname = ""; sfdc_label = None
    if df_sfdc is not None and not df_sfdc.empty:
        _rn = {c:_SFDC_COL_MAP[c.lower().strip()] for c in df_sfdc.columns if c.lower().strip() in _SFDC_COL_MAP}
        df_sn = df_sfdc.rename(columns=_rn)
        if "first_name" in df_sn.columns and "last_name" in df_sn.columns:
            df_sn["contact_name"] = (df_sn["first_name"].fillna("").astype(str)+" "+df_sn["last_name"].fillna("").astype(str)).str.strip()
        pnm = str(sel.get(name_col,"")) if name_col else ""
        sfdc_match, sfdc_label = _fuzzy_sfdc(df_sn, pnm, str(customer))
        if not sfdc_match.empty:
            ec = "email" if "email" in sfdc_match.columns else None
            nc = "contact_name" if "contact_name" in sfdc_match.columns else None
            fc = "impl_contact_flag" if "impl_contact_flag" in sfdc_match.columns else None
            br = sfdc_match.iloc[0]
            if fc:
                fl = sfdc_match[sfdc_match[fc].astype(str).str.lower().isin(["true","1","yes","x"])]
                if not fl.empty: br = fl.iloc[0]
            sfdc_email = str(br.get(ec,"")).strip() if ec else ""
            sfdc_cname = str(br.get(nc,"")).strip() if nc else ""
            if sfdc_email in ("nan","None",""): sfdc_email = ""
            if sfdc_cname in ("nan","None",""): sfdc_cname = ""

    st.markdown('<p class="ce-label">Recipient</p>', unsafe_allow_html=True)
    if sfdc_label:
        st.caption(f"✅ {sfdc_label}")
    elif df_sfdc is not None:
        st.caption(f"No SFDC match for '{customer}'")
    else:
        st.caption("SFDC not loaded")

    recip = st.text_input("To", value=sfdc_email, placeholder="customer@example.com", key="ce_to")
    cname = st.text_input("Contact name", value=sfdc_cname, placeholder="First name", key="ce_cn")
    cc_in = st.text_input("CC", value=_consultant_email(_logged_in), key="ce_cc")
    cc_emails = [e.strip() for e in cc_in.split(",") if e.strip()]

# ── Journey rail ──────────────────────────────────────────────────────────────
st.markdown(_build_journey_html(sel, _active_tab), unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN: compose left | preview right
# ═══════════════════════════════════════════════════════════════════════════════
compose_col, preview_col = st.columns([1,1], gap="medium")

# Auto-context (name flipped, contact injected before render)
_disp = _flip_name(_logged_in)
auto_ctx = build_auto_context(sel, _disp, {"contact_name":cname} if cname else None)
if cname: auto_ctx["CUSTOMER_CONTACT_NAME"] = cname
auto_ctx["SENDER"] = _disp; auto_ctx["CONSULTANT_NAME"] = _disp

# Shared state for preview
if "ce_preview_subject" not in st.session_state:
    st.session_state["ce_preview_subject"] = ""
if "ce_preview_body" not in st.session_state:
    st.session_state["ce_preview_body"] = ""

with compose_col:
    tab_w, tab_s, tab_l = st.tabs(["Welcome","Post-Session","Lifecycle (UAT → Closure)"])

    # ── TAB: Welcome ──────────────────────────────────────────────────────────
    with tab_w:
        st.session_state["_ce_tab"] = "Welcome"
        _sku_key = _sku(str(product_raw)) if product_raw and str(product_raw) not in ("","nan","None") else None
        tmpl_w = get_welcome_template(_sku_key) if _sku_key else None
        if not tmpl_w:
            st.caption(f"No automatic match for '{product_raw}'. Select manually:")
            opts = list_welcome_templates()
            ch = st.selectbox("Template",[t["display_name"] for t in opts],key="w_manual")
            tmpl_w = get_welcome_template(next(t["sku_key"] for t in opts if t["display_name"]==ch))

        var = st.radio("Sender variant",
                       ["Variant A — PM or automated","Variant B — Consultant sends"],
                       horizontal=True, key="w_var")
        vk = "variant_a" if "A" in var else "variant_b"
        subj_r, body_r = render_template(tmpl_w[vk]["body"], tmpl_w["subject"], auto_ctx)
        lib_meta = _welcome_library(); ssf_w = lib_meta.get("ss_milestone_on_send")

        # Pre-send checks — card with no background
        st.markdown('<div class="ce-card">', unsafe_allow_html=True)
        st.markdown('<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;opacity:.6;margin-bottom:6px">Pre-send checks</div>', unsafe_allow_html=True)
        _has_recip = bool(recip and "@" in recip)
        _has_sfdc  = bool(sfdc_email)
        _has_tmpl  = bool(tmpl_w)
        _no_phs    = not _missing_phs(body_r + subj_r)
        for ok, msg in [
            (_has_sfdc,  "SFDC contact linked" if _has_sfdc else "No SFDC match — enter recipient manually"),
            (_has_recip, "Recipient email set" if _has_recip else "Recipient email missing"),
            (_has_tmpl,  f"Template: {tmpl_w['display_name']}" if _has_tmpl else "No template matched"),
            (not _intro_done, "No Welcome email on record (ready to send)" if not _intro_done else f"Welcome already sent {iv} — check before resending"),
        ]:
            cls = "check-ok" if ok else "check-bad"
            icon = "✓" if ok else "✗"
            st.markdown(f'<div class="check-row {cls}">{icon}&nbsp; {msg}</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        if ssf_w:
            st.markdown(f'<div class="ce-tip">On send, <b>{ssf_w}</b> will be date-stamped in Smartsheet.</div>', unsafe_allow_html=True)

        st.session_state["ce_preview_subject"] = subj_r
        st.session_state["ce_preview_body"]    = body_r

        _can_send = _has_recip and not _missing_phs(body_r + subj_r)
        c1,c2 = st.columns([1,2])
        with c1:
            st.button("Copy plain text", key="w_copy", use_container_width=True)
        with c2:
            lbl = "Send & complete Welcome stage" if st.session_state.get("_gmail_approved") else "📋 Log Send"
            if st.button(lbl, key="btn_w", type="primary", use_container_width=True, disabled=not _has_recip):
                st.session_state["_req_w"] = {"subj":subj_r,"body":body_r,"ssf":ssf_w}
                st.rerun()
        if st.session_state.get("_req_w"):
            r = st.session_state.pop("_req_w")
            try:
                with st.spinner("Logging…"):
                    ok,sid = execute_send(project_id=project_id,template_id=f"welcome_{_sku_key or 'manual'}",
                        template_name=f"Welcome — {tmpl_w['display_name']}",
                        subject=r["subj"],body=r["body"],recipient_email=recip,
                        cc_emails=cc_emails,ss_milestone_field=r["ssf"])
                if ok:
                    st.success(f"✓ Logged — ID: `{sid}`")
                    if r["ssf"] and _do_write(project_id, r["ssf"], datetime.date.today(), sel):
                        mark_ss_writeback_done(sid)
                else: st.error(f"Failed: {sid}")
            except Exception as ex: st.error(f"Error: {ex}"); st.exception(ex)

    # ── TAB: Post-Session ─────────────────────────────────────────────────────
    with tab_s:
        st.session_state["_ce_tab"] = "Post-Session"
        psk = _ps_key(str(product_raw))
        if not psk:
            st.caption(f"No post-session templates for '{product_raw}'.")
        else:
            sessions = get_post_session_templates(psk)
            sopts = {s["id"]:(f"Session {s['session_number']} — {s['name']}"+(f" [{s['variant_note']}]" if s.get("variant_note") else ""),s) for s in sessions}
            cid = st.selectbox("Session",list(sopts.keys()),format_func=lambda k:sopts[k][0],key="s_pick")
            _,tmpl_s = sopts[cid]
            st.caption(f"Audience: {tmpl_s.get('audience','Full project team')}")
            mctx: dict = {}
            if tmpl_s.get("editable_fields"):
                with st.expander("Session details",expanded=True):
                    for f in tmpl_s["editable_fields"]:
                        k,lb,ft=f["key"],f["label"],f.get("type","text")
                        req=f.get("required",False); ph=f.get("placeholder","")
                        if ft=="text":       v=st.text_input(lb+(" *"if req else ""),placeholder=ph,key=f"s_{cid}_{k}")
                        elif ft=="textarea": v=st.text_area(lb+(" *"if req else ""),placeholder=ph,height=70,key=f"s_{cid}_{k}")
                        elif ft=="multiselect":
                            s2=st.multiselect(lb+(" *"if req else ""),options=f.get("options",[]),key=f"s_{cid}_{k}")
                            v="\n".join(f"  • {o}" for o in s2)
                        elif ft=="select": v=st.selectbox(lb,f.get("options",[]),key=f"s_{cid}_{k}")
                        else: v=st.text_input(lb,key=f"s_{cid}_{k}")
                        mctx[k]=v
                        if k=="GO_LIVE_READINESS" and v:
                            rm=tmpl_s.get("go_live_readiness_text",{})
                            res=rm.get(v[0],v)
                            if "{HYPERCARE_DATE}" in res: res=res.replace("{HYPERCARE_DATE}",mctx.get("HYPERCARE_DATE","{HYPERCARE_DATE}"))
                            mctx["GO_LIVE_READINESS_TEXT"]=res
            subj_s,body_s=render_template(tmpl_s["body"],tmpl_s["subject"],{},{**auto_ctx,**mctx})
            ssf_s=tmpl_s.get("ss_milestone_on_send")
            st.session_state["ce_preview_subject"]=subj_s; st.session_state["ce_preview_body"]=body_s
            if ssf_s: st.markdown(f'<div class="ce-tip">On send, <b>{ssf_s}</b> will be date-stamped.</div>',unsafe_allow_html=True)
            if st.button("📋 Log Send",key="btn_s",type="primary",disabled=not recip):
                st.session_state["_req_s"]={"subj":subj_s,"body":body_s,"ssf":ssf_s,"tid":tmpl_s["id"],"tnm":tmpl_s["name"]}
                st.rerun()
            if st.session_state.get("_req_s"):
                r=st.session_state.pop("_req_s")
                try:
                    with st.spinner("Logging…"):
                        ok,sid=execute_send(project_id=project_id,template_id=r["tid"],template_name=r["tnm"],
                            subject=r["subj"],body=r["body"],recipient_email=recip,cc_emails=cc_emails,ss_milestone_field=r["ssf"])
                    if ok:
                        st.success(f"✓ Logged — ID: `{sid}`")
                        if r["ssf"] and _do_write(project_id,r["ssf"],datetime.date.today(),sel): mark_ss_writeback_done(sid)
                    else: st.error(f"Failed: {sid}")
                except Exception as ex: st.error(f"Error: {ex}"); st.exception(ex)

    # ── TAB: Lifecycle ────────────────────────────────────────────────────────
    with tab_l:
        st.session_state["_ce_tab"] = "Lifecycle"
        lc_all=list_lifecycle_templates()
        lc_opts={t["id"]:t for t in lc_all}
        lcid=st.selectbox("Template",list(lc_opts.keys()),format_func=lambda k:f"[{lc_opts[k]['category']}] {lc_opts[k]['name']}",key="lc_pick")
        tmpl_l=get_lifecycle_template(lcid)
        st.caption(f"When to send: {tmpl_l['trigger']}")
        for tip in tmpl_l.get("tips",[]): st.markdown(f'<div class="ce-tip">💡 {tip}</div>',unsafe_allow_html=True)
        vbody=tmpl_l.get("body","")
        if tmpl_l.get("variants"):
            vlbls={v["key"]:f"{v['label']} — {v['description']}" for v in tmpl_l["variants"]}
            cv=st.radio("Scenario",list(vlbls.keys()),format_func=lambda k:vlbls[k],key=f"lv_{lcid}",label_visibility="collapsed")
            vbody=tmpl_l["variant_bodies"][cv]
        mctx_l: dict={}
        if tmpl_l.get("editable_fields"):
            with st.expander("Details",expanded=True):
                for f in tmpl_l["editable_fields"]:
                    k,lb=f["key"],f["label"]; req=f.get("required",False); src=f.get("source",""); default=str(f.get("default",""))
                    if src=="drs_prod_cutover":
                        raw=sel.get("prod_cutover") or sel.get("Prod Cutover")
                        if raw:
                            try: default=pd.to_datetime(raw).date().isoformat()
                            except: pass
                    elif src=="drs_project_link":
                        default=str(sel.get("project_link") or sel.get("Project Link") or "")
                    elif src=="calculated_go_live_plus_14":
                        glr=sel.get("prod_cutover") or sel.get("go_live_date") or mctx_l.get("GO_LIVE_DATE","")
                        if glr:
                            try: default=(pd.to_datetime(glr).date()+datetime.timedelta(days=14)).isoformat()
                            except: pass
                    v=st.text_input(lb+(" *"if req else ""),value=default,placeholder="YYYY-MM-DD"if f.get("type")=="date" else "",key=f"l_{lcid}_{k}")
                    mctx_l[k]=v
        subj_l,body_l=render_template(vbody,tmpl_l["subject"],{},{**auto_ctx,**mctx_l})
        ssf_l=tmpl_l.get("ss_milestone_on_send"); gls=mctx_l.get("GO_LIVE_DATE","")
        st.session_state["ce_preview_subject"]=subj_l; st.session_state["ce_preview_body"]=body_l
        if ssf_l:
            fd=", ".join(ssf_l) if isinstance(ssf_l,list) else ssf_l
            st.markdown(f'<div class="ce-tip">On send, <b>{fd}</b> will be date-stamped (existing values not overwritten).</div>',unsafe_allow_html=True)
        if st.button("📋 Log Send",key="btn_l",type="primary",disabled=not recip):
            st.session_state["_req_l"]={"subj":subj_l,"body":body_l,"ssf":ssf_l,"gls":gls}
            st.rerun()
        if st.session_state.get("_req_l"):
            r=st.session_state.pop("_req_l")
            try:
                with st.spinner("Logging…"):
                    ok,sid=execute_send(project_id=project_id,template_id=lcid,template_name=tmpl_l["name"],
                        subject=r["subj"],body=r["body"],recipient_email=recip,cc_emails=cc_emails,ss_milestone_field=r["ssf"])
                if ok:
                    st.success(f"✓ Logged — ID: `{sid}`")
                    if r["ssf"]:
                        try: gld=datetime.date.fromisoformat(r["gls"][:10]) if r["gls"] else datetime.date.today()
                        except: gld=datetime.date.today()
                        for f in (r["ssf"] if isinstance(r["ssf"],list) else [r["ssf"]]):
                            wd=gld if f in (SS_GO_LIVE_DATE,SS_PROD_CUTOVER) else datetime.date.today()
                            if _do_write(project_id,f,wd,sel): mark_ss_writeback_done(sid)
                else: st.error(f"Failed: {sid}")
            except Exception as ex: st.error(f"Error: {ex}"); st.exception(ex)

# ── RIGHT: Email preview + session log ───────────────────────────────────────
with preview_col:
    st.markdown('<p class="ce-label">Live Preview</p>', unsafe_allow_html=True)
    _prev_subj = st.session_state.get("ce_preview_subject","")
    _prev_body = st.session_state.get("ce_preview_body","")
    if _prev_body:
        st.markdown(
            _email_preview_html(_prev_subj, _prev_body, recip, cc_in),
            unsafe_allow_html=True
        )
        # Editable plain text in expander
        with st.expander("Edit before sending"):
            st.text_area("Subject", value=_prev_subj, key="preview_subj_edit", height=40)
            st.text_area("Body", value=_prev_body, key="preview_body_edit", height=280)
    else:
        st.markdown('<div class="ce-card" style="text-align:center;padding:40px;opacity:.4">Select a template to see preview</div>',unsafe_allow_html=True)

    # Session send log
    log = get_session_send_log()
    proj_log = [e for e in log if e["project_id"]==project_id]
    if proj_log:
        st.markdown('<p class="ce-label" style="margin-top:18px">Sent This Session</p>',unsafe_allow_html=True)
        for e in proj_log:
            dt=e["sent_at"][:16].replace("T"," ")
            st.markdown(
                f'<div class="log-row"><span class="pill-ok">✓ Sent</span>&nbsp;'
                f'<b>{e["template_name"]}</b><br>'
                f'<span style="font-size:11px;opacity:.6">{dt} UTC → {e["recipient_email"]}</span>'
                f'</div>', unsafe_allow_html=True)
