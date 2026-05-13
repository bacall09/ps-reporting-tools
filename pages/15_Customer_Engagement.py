"""
pages/15_Customer_Engagement.py  v5
Layout: project picker row → journey rail row → compose(left) | preview(right)
- Recipient inside compose pane
- Persistent send footer with SS checkbox
- Journey dates from DRS milestones
- Green highlights for auto-filled values in preview
- Bullet rendering fixed
- Multi-product consolidated welcome
"""
import streamlit as st
import pandas as pd
import datetime
import re
from rapidfuzz import fuzz

st.session_state["current_page"] = "Customer Engagement"

st.markdown("""
<div style='background:#050D1F;padding:20px 28px 16px;border-radius:0 0 8px 8px;margin-bottom:20px'>
  <span style='color:#4472C4;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase'>Customer Engagement</span>
  <h2 style='color:#ffffff;margin:4px 0 0;font-size:22px;font-weight:600;letter-spacing:-0.3px'>Lifecycle Email Composer</h2>
</div>
""", unsafe_allow_html=True)

st.markdown("""<style>
.ce-label{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.9px;color:#4472C4;margin:0 0 5px}
.ce-card{border:1px solid rgba(128,128,128,.22);border-radius:8px;padding:14px 18px;margin-bottom:10px;color:inherit}
.ce-tip{border-left:3px solid #4472C4;border-radius:0;padding:7px 12px;font-size:12px;margin-bottom:8px;color:inherit}
.journey-rail{display:flex;border:1px solid rgba(128,128,128,.2);border-radius:8px;overflow:hidden;margin-bottom:0}
.sj{flex:1;padding:9px 10px 7px;border-right:1px solid rgba(128,128,128,.15);min-width:0;color:inherit}
.sj:last-child{border-right:none}
.sj-num{font-size:10px;font-weight:700;letter-spacing:.3px;margin-bottom:1px}
.sj-lbl{font-size:11px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sj-date{font-size:10px;margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;opacity:.7}
.sj.done .sj-num,.sj.done .sj-lbl{color:#16a34a}
.sj.active{border-bottom:2px solid #4472C4}
.sj.active .sj-num,.sj.active .sj-lbl{color:#4472C4}
.sj.locked{opacity:.38}
@media(prefers-color-scheme:dark){.sj.done .sj-num,.sj.done .sj-lbl{color:#7ed4a4}}
.stApp[data-theme="dark"] .sj.done .sj-num,.stApp[data-theme="dark"] .sj.done .sj-lbl{color:#7ed4a4}
.pill-ok{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:rgba(34,197,94,.14);color:#15803d}
.pill-warn{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:rgba(239,68,68,.14);color:#dc2626}
.pill-info{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:rgba(68,114,196,.14);color:#4472C4}
.pill-amber{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:rgba(234,179,8,.14);color:#854d0e}
@media(prefers-color-scheme:dark){
  .pill-ok{color:#7ed4a4}.pill-warn{color:#fca5a5}.pill-info{color:#93b4e8}.pill-amber{color:#fcd34d}
}
.stApp[data-theme="dark"] .pill-ok{color:#7ed4a4}
.stApp[data-theme="dark"] .pill-warn{color:#fca5a5}
.stApp[data-theme="dark"] .pill-info{color:#93b4e8}
.stApp[data-theme="dark"] .pill-amber{color:#fcd34d}
.chk-ok{font-size:12px;color:#16a34a;display:flex;align-items:center;gap:6px;padding:2px 0}
.chk-bad{font-size:12px;color:#dc2626;display:flex;align-items:center;gap:6px;padding:2px 0}
@media(prefers-color-scheme:dark){.chk-ok{color:#7ed4a4}.chk-bad{color:#fca5a5}}
.stApp[data-theme="dark"] .chk-ok{color:#7ed4a4}
.stApp[data-theme="dark"] .chk-bad{color:#fca5a5}
.send-footer{border-top:1px solid rgba(128,128,128,.18);padding:12px 16px;margin-top:12px}
.footer-ready{font-size:12px;font-weight:600;color:#16a34a;margin-bottom:3px}
.footer-blocked{font-size:12px;font-weight:600;color:#dc2626;margin-bottom:3px}
.footer-missing{font-size:11px;color:#dc2626;margin-bottom:6px}
.footer-chk{display:flex;align-items:center;gap:5px;font-size:11px;color:inherit;margin-bottom:3px}
@media(prefers-color-scheme:dark){.footer-ready{color:#7ed4a4}.footer-blocked{color:#fca5a5}.footer-missing{color:#fca5a5}}
.stApp[data-theme="dark"] .footer-ready{color:#7ed4a4}
.stApp[data-theme="dark"] .footer-blocked{color:#fca5a5}
.stApp[data-theme="dark"] .footer-missing{color:#fca5a5}
.log-row{border-bottom:1px solid rgba(128,128,128,.13);padding:5px 0;font-size:12px}
.email-preview-wrap{background:#ffffff;border-radius:8px;border:1px solid #e2e8f0;overflow:hidden}
.ep-chrome{background:#f8fafc;border-bottom:1px solid #e2e8f0;padding:7px 12px;display:flex;align-items:center;gap:5px}
.ep-dot{width:8px;height:8px;border-radius:50%;background:#e2e8f0;display:inline-block}
.ep-hdr{padding:9px 14px;border-bottom:1px solid #e2e8f0}
.ep-row{display:flex;gap:8px;font-size:11px;padding:2px 0;color:#64748b}
.ep-lbl{width:36px;color:#94a3b8;flex-shrink:0}
.ep-body{padding:14px 16px;font-size:12.5px;color:#0f172a;line-height:1.65;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}
.ep-hdr-txt{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#4472C4;margin:10px 0 3px}
.ep-ph{display:inline;background:rgba(220,38,38,.1);color:#dc2626;border:1px dashed rgba(220,38,38,.4);border-radius:3px;padding:1px 4px;font-size:11px;font-family:monospace}
.ep-filled{display:inline;background:rgba(22,163,74,.1);color:#15803d;border-radius:3px;padding:1px 4px;font-size:11px}
.ep-bullet{margin:2px 0 2px 14px}
.ep-bar{padding:7px 14px;background:rgba(220,38,38,.06);border-top:1px solid rgba(220,38,38,.15);font-size:11px;color:#dc2626}
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
        get_session_send_log, _welcome_library,
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

# ── SFDC helpers ──────────────────────────────────────────────────────────────
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
    """
    Exact port of reengagement page fuzzy_match_sfdc.
    1. Exact opportunity name match
    2. Fuzzy account + product keyword boost (+30) — returns ALL rows for matched opp
    3. Account-only fallback at 75%
    """
    if df_sfdc is None or df_sfdc.empty: return pd.DataFrame(), None
    df = df_sfdc.copy()
    col_map = {c.lower().strip(): c for c in df.columns}
    opp_col = col_map.get("opportunity")
    acc_col = col_map.get("account")
    oid_col = col_map.get("opportunity_id")

    # 1. Exact opportunity name match
    if opp_col:
        exact = df[df[opp_col].astype(str).str.lower().str.strip()
                   == str(proj_name).lower().strip()]
        if not exact.empty: return exact, "Exact match"

    # 2. Fuzzy account + product keyword scoring
    drs_words = set(_clean_acct(acct_name))
    drs_prods = _prod_hints(proj_name)
    best_score = 0; best_opp_id = None; best_opp_nm = None
    for _, row in df.iterrows():
        sfdc_acct  = str(row.get(acc_col or "account", ""))
        sfdc_opp   = str(row.get(opp_col or "opportunity", ""))
        sfdc_words = set(_clean_acct(sfdc_acct))
        word_score = len(drs_words & sfdc_words) / max(len(drs_words), 1) * 100
        fuzz_score = fuzz.token_set_ratio(" ".join(drs_words), " ".join(sfdc_words))
        acct_score = max(word_score, fuzz_score * 0.7)
        prod_match = bool(drs_prods & _prod_hints(sfdc_opp))
        score = acct_score + (30 if prod_match else 0)
        if score > best_score:
            best_score = score
            best_opp_id = row.get(oid_col) if oid_col else None
            best_opp_nm = row.get(opp_col) if opp_col else None

    label = f"Fuzzy match ({int(best_score)}%)"
    if best_score >= 60:
        if best_opp_id is not None and oid_col:
            rows = df[df[oid_col] == best_opp_id]
            if not rows.empty: return rows, label
        if best_opp_nm is not None and opp_col:
            rows = df[df[opp_col] == best_opp_nm]
            if not rows.empty: return rows, label

    # 3. Account-only fallback at 75%
    if acc_col:
        df["_sc"] = df[acc_col].apply(
            lambda x: fuzz.token_set_ratio(str(acct_name).lower(), str(x).lower()))
        top = df[df["_sc"] >= 75].sort_values("_sc", ascending=False)
        df.drop(columns=["_sc"], inplace=True, errors="ignore")
        if not top.empty: return top, "Account match"

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
        em=EMPLOYEE_ROLES.get(name,{}).get("email","")
        if em: return em
    except Exception: pass
    return f"{re.sub(r'[^a-z0-9.]','',_flip_name(name).lower().replace(' ','.'))}@zoneandco.com"

def _missing_phs(text):
    return list(set(re.findall(r"\{[A-Z_]+\}",text)))

_PMAP = {
    "zoneapp: capture":"ZoneCapture","zoneapp: approvals":"ZoneApprovals",
    "zoneapp: reconcile":"ZoneReconcile","zoneapp: reconcile 2.0":"ZoneReconcile_BankConnect",
    "zoneapp: reconcile with bank connectivity":"ZoneReconcile_BankConnect",
    "zoneapp: reconcile with cc import":"ZoneReconcile_CCImport",
    "zoneapp: e-invoicing":"EInvoicing",
    "zoneapp: ap payment":"ZoneApprovals",  # AP Payment uses Approvals template
    "zoneapp: payments":"ZoneApprovals",
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

# ── Journey ───────────────────────────────────────────────────────────────────
_JOURNEY = [
    {"id":"welcome",            "label":"Welcome",         "ms_col":"ms_intro_email",   "ms_alt":"Intro. Email Sent"},
    {"id":"post_session_1",     "label":"Post-Session #1", "ms_col":"ms_enablement",    "ms_alt":"Enablement Session"},
    {"id":"post_session_2",     "label":"Post-Session #2", "ms_col":"ms_session1",      "ms_alt":"Session #1"},
    {"id":"uat_signoff",        "label":"UAT Sign-Off",    "ms_col":"ms_uat_signoff",   "ms_alt":"UAT Signoff"},
    {"id":"go_live_hypercare_kickoff","label":"Go-Live",   "ms_col":"ms_prod_cutover",  "ms_alt":"Prod Cutover"},
    {"id":"hypercare_closure",  "label":"Hypercare close", "ms_col":"ms_transition",    "ms_alt":"Transition to Support"},
]

def _ms_date(stage, drs_row):
    """Return formatted date string for a milestone, or ''."""
    if not drs_row: return ""
    v = drs_row.get(stage["ms_col"]) or drs_row.get(stage["ms_alt"])
    if not v or str(v).strip() in ("","None","nan","NaT"): return ""
    try:
        return pd.to_datetime(v).strftime("%-d %b")
    except Exception:
        return str(v)[:10]

def _build_journey(drs_row):
    statuses = []
    for s in _JOURNEY:
        d = _ms_date(s, drs_row)
        statuses.append("done" if d else "pending")
    first_p = next((i for i,st in enumerate(statuses) if st=="pending"), len(_JOURNEY)-1)
    parts = ['<div class="journey-rail">']
    for i,stage in enumerate(_JOURNEY):
        cls = statuses[i]
        if i==first_p and cls=="pending": cls="active"
        elif i>first_p and cls=="pending": cls="locked"
        num_icon = f"✓ 0{i+1}" if statuses[i]=="done" else (f"▼ 0{i+1}" if cls=="active" else f"0{i+1}")
        date_str = _ms_date(stage,drs_row)
        sub = f"Sent {date_str}" if date_str else ("composing" if cls=="active" else "")
        parts.append(
            f'<div class="sj {cls}">'
            f'<div class="sj-num">{num_icon}</div>'
            f'<div class="sj-lbl">{stage["label"]}</div>'
            f'<div class="sj-date">{sub}</div>'
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
    fields=[ss_field] if isinstance(ss_field,str) else ss_field
    wrote=False
    pn=str(drs_row.get("project_name",project_id)) if drs_row else project_id
    for f in fields:
        try:
            p=build_ss_writeback(project_id,f,date_val,current_drs_row=drs_row)
            if p["skipped"]: st.info(f"ℹ️ **{f}** already set — not overwritten."); continue
            if not p["fields"]: continue
            ok,errs=write_row_updates([{"_ss_row_id":ss_row_id,"project_name":pn,"changes":{f:date_val.isoformat()}}])
            if ok: st.success(f"✓ Smartsheet: **{f}** → {date_val.strftime('%d %b %Y')}"); wrote=True
            for e in (errs or []): st.warning(f"Writeback error: {e}")
        except Exception as ex: st.warning(f"Writeback failed for **{f}**: {ex}")
    return wrote

# ── Email preview renderer ────────────────────────────────────────────────────
def _email_html(subject, body, to_email, cc_email, auto_values: set):
    """
    Render white Gmail-style preview.
    auto_values: set of strings that were auto-filled — highlighted green.
    Unfilled {PLACEHOLDERS} highlighted red.
    Bullets preserved. Runbook exception: hardcoded white background.
    """
    def _htmlify(text):
        out=[]
        for line in text.split("\n"):
            s=line.strip()
            if not s:
                out.append('<div style="margin:4px 0"></div>')
            elif s==s.upper() and len(s)>3 and not s.startswith("•") and not s.startswith("-"):
                out.append(f'<div class="ep-hdr-txt">{s}</div>')
            elif s.startswith("•") or s.startswith("* ") or (s.startswith("- ") and len(s)>3) or (len(s)>2 and s[0].isdigit() and s[1] in ".)"):
                out.append(f'<div class="ep-bullet">{s}</div>')
            elif s.endswith(":") and len(s)<60 and s[0].isupper():
                out.append(f'<div style="font-weight:600;margin-top:8px;color:#0f172a">{s}</div>')
            else:
                out.append(f'<div>{s}</div>')
        return "\n".join(out)

    def _highlight(html, auto_vals):
        # Red for unfilled placeholders
        html = re.sub(r'\{([A-Z_]+)\}', r'<span class="ep-ph">{\1}</span>', html)
        # Green for auto-filled values (escape for regex, only highlight standalone occurrences)
        for val in sorted(auto_vals, key=len, reverse=True):
            if val and len(val) > 2 and val not in ("{}", "—"):
                escaped = re.escape(val)
                html = re.sub(
                    f'(?<![>\\w]){escaped}(?![<\\w])',
                    f'<span class="ep-filled">{val}</span>',
                    html, count=1
                )
        return html

    body_html = _highlight(_htmlify(body), auto_values)
    missing = _missing_phs(body+subject)
    ph_bar = f'<div class="ep-bar">⚠ {len(missing)} placeholder(s) empty: {", ".join(missing)}</div>' if missing else ""

    return f"""<div class="email-preview-wrap">
<div class="ep-chrome"><span class="ep-dot"></span><span class="ep-dot"></span><span class="ep-dot"></span>
<span style="font-size:10px;color:#94a3b8;margin-left:6px">Gmail preview</span></div>
<div class="ep-hdr">
<div class="ep-row"><span class="ep-lbl">TO</span>{to_email or '—'}</div>
<div class="ep-row"><span class="ep-lbl">CC</span>{cc_email or '—'}</div>
<div class="ep-row"><span class="ep-lbl">SUBJ</span><strong style="color:#0f172a">{subject}</strong></div>
</div>
<div class="ep-body">{body_html}</div>
{ph_bar}
</div>"""

# ── Build DRS dataframe ───────────────────────────────────────────────────────
df_all=_df_drs.copy()
cm={c.lower().strip():c for c in df_all.columns}
name_col  =cm.get("project_name")    or cm.get("project name")
cust_col  =cm.get("customer")        or cm.get("account")  # DRS loader maps "account name" → "account"
prod_col  =cm.get("project_type")    or cm.get("project type") or cm.get("product")
id_col    =cm.get("project_id")      or cm.get("project id")
status_col=cm.get("status")
pm_col    =cm.get("project_manager") or cm.get("project manager")
intro_col =(cm.get("intro. email sent") or cm.get("intro email sent")
            or cm.get("ms_intro_email") or cm.get("intro_email_sent"))
legacy_col=cm.get("legacy")
ss_rid_col=cm.get("_ss_row_id") or cm.get("ss_row_id") or cm.get("row_id")
sales_rep_col=cm.get("sales rep") or cm.get("sales_rep") or cm.get("account executive") or cm.get("ae")

_ss_row_id_map: dict={}
if ss_rid_col and id_col:
    for _,r in _df_drs.iterrows():
        pid=str(r.get(id_col,"")).strip(); rid=r.get(ss_rid_col)
        if pid and rid: _ss_row_id_map[pid]=rid

if not _is_mgr and pm_col:
    df_all=df_all[df_all[pm_col].apply(lambda x: name_matches(str(x),_logged_in))]

if _is_mgr and pm_col:
    _browse=st.session_state.get("_browse_passthrough") or st.session_state.get("home_browse","")
    if _browse and _browse not in ("— My own view —","— Select —","👥 All team",""):
        if _browse.startswith("── ") and _browse.endswith(" ──"):
            try:
                from shared.constants import get_region_consultants
                from shared.config import EMPLOYEE_LOCATION,PS_REGION_MAP,PS_REGION_OVERRIDE,ACTIVE_EMPLOYEES
                _rc=get_region_consultants(_browse[3:-3].strip(),EMPLOYEE_LOCATION,PS_REGION_MAP,PS_REGION_OVERRIDE,ACTIVE_EMPLOYEES)
                _f=df_all[df_all[pm_col].astype(str).str.strip().str.lower().isin([n.lower() for n in _rc])]
                if not _f.empty: df_all=_f
            except Exception: pass
        else:
            _f=df_all[df_all[pm_col].apply(lambda x: name_matches(str(x),_browse))]
            if not _f.empty: df_all=_f

if status_col:
    df_all=df_all[~df_all[status_col].astype(str).str.lower().isin(["closed","cancelled","complete","completed"])]
if legacy_col:
    df_all=df_all[~df_all[legacy_col].astype(str).str.strip().str.lower().isin(["yes","y","true","1"])]

df_all=df_all.reset_index(drop=True)
if df_all.empty:
    st.info("No active projects found.")
    st.stop()

def _needs_intro(row):
    if not intro_col: return False
    v=row.get(intro_col,"")
    return v is None or str(v).strip() in ("","None","nan","NaT")
df_all["_needs_welcome"]=df_all.apply(_needs_intro,axis=1)

# ═══════════════════════════════════════════════════════════════════════════════
# ROW 1: Project picker (full width)
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<p class="ce-label" style="margin-bottom:4px">Select Project</p>', unsafe_allow_html=True)

df=df_all.copy()
_active_tab=st.session_state.get("_ce_tab","Welcome")
if intro_col and _active_tab=="Welcome":
    if st.checkbox("Only projects needing Welcome email", value=True, key="ce_wf"):
        _fil=df[df["_needs_welcome"]]
        if not _fil.empty: df=_fil.reset_index(drop=True)
        else: st.success("All projects have a Welcome email on record.")

if df.empty:
    st.info("No projects match this filter.")
    st.stop()

def _plabel(row):
    c=str(row.get(cust_col,"")) if cust_col else ""
    n=str(row.get(name_col,"")) if name_col else ""
    p=str(row.get(prod_col,"")) if prod_col else ""
    base=f"{c}  —  {n}" if c and n else (n or c)
    return base+(f"  ·  {p}" if p and p not in ("nan","None") else "")

df["_sid"]=df.apply(lambda r:str(r.get(id_col,r.name)) if id_col else str(r.name),axis=1)
sids=df["_sid"].tolist()
sid_map={s:i for i,s in enumerate(sids)}
_prev=st.session_state.get("_ce_proj_sid")
_def=_prev if _prev in sids else sids[0]

# Multi-project detection for same customer
_cust_groups: dict={}
if cust_col:
    for sid in sids:
        c=str(df.iloc[sid_map[sid]].get(cust_col,""))
        if c and c not in ("nan","None"): _cust_groups.setdefault(c,[]).append(sid)

sel_col, info_col = st.columns([3,1])
with sel_col:
    selected_sid=st.selectbox(
        "Project",options=sids,
        format_func=lambda s:_plabel(df.iloc[sid_map[s]]),
        index=sids.index(_def),
        label_visibility="collapsed",key="ce_proj",
    )
    st.session_state["_ce_proj_sid"]=selected_sid

if st.session_state.get("_last_proj")!=selected_sid:
    st.session_state["_last_proj"]=selected_sid
    for k in list(st.session_state.keys()):
        if any(k.startswith(p) for p in ["ce_to","ce_cn","ce_cc","w_","s_","l_","sess_","lc_"]):
            del st.session_state[k]

sel_idx=sid_map[selected_sid]
sel=_row_dict(df.iloc[sel_idx])
customer=sel.get(cust_col,"") if cust_col else ""
product_raw=sel.get(prod_col,"") if prod_col else ""
project_id=str(sel.get(id_col,str(sel_idx))) if id_col else str(sel_idx)
if "_ss_row_id" not in sel and project_id in _ss_row_id_map:
    sel["_ss_row_id"]=_ss_row_id_map[project_id]

# Multi-product: siblings for same customer
siblings=[s for s in _cust_groups.get(str(customer),[]) if s!=selected_sid] if customer else []

with info_col:
    iv=sel.get(intro_col,"") if intro_col else ""
    _intro_done=iv and str(iv).strip() not in ("","None","nan","NaT")
    if _intro_done:
        st.markdown(f'<div style="margin-top:6px"><span class="pill-ok">Welcome sent {iv}</span></div>',unsafe_allow_html=True)
    else:
        st.markdown('<div style="margin-top:6px"><span class="pill-warn">Welcome pending</span></div>',unsafe_allow_html=True)
    if siblings:
        st.markdown(f'<div style="margin-top:4px"><span class="pill-amber">{len(siblings)+1} products for this customer</span></div>',unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ROW 2: Journey rail (full width, with dates)
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown(_build_journey(sel),unsafe_allow_html=True)
st.markdown('<div style="margin-bottom:12px"></div>',unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ROW 3: Compose (left) | Preview (right)
# ═══════════════════════════════════════════════════════════════════════════════
compose_col, preview_col = st.columns([1,1],gap="medium")

# SFDC lookup (shared across all tabs)
df_sfdc=st.session_state.get("df_sfdc")
sfdc_email=""; sfdc_cname=""; sfdc_label=None; _sfdc_debug={}
sfdc_cc_emails: list = []  # additional impl contacts for CC
if df_sfdc is not None and not df_sfdc.empty:
    _rn={c:_SFDC_COL_MAP[c.lower().strip()] for c in df_sfdc.columns if c.lower().strip() in _SFDC_COL_MAP}
    df_sn=df_sfdc.rename(columns=_rn)
    if "first_name" in df_sn.columns and "last_name" in df_sn.columns:
        df_sn["contact_name"]=(df_sn["first_name"].fillna("").astype(str)+" "+df_sn["last_name"].fillna("").astype(str)).str.strip()
    pnm=str(sel.get(name_col,"")) if name_col else ""
    _sfdc_debug={"sfdc_cols":list(df_sn.columns),"drs_customer":str(customer),"drs_project":pnm,"rows":len(df_sn)}
    sfdc_match,sfdc_label=_fuzzy_sfdc(df_sn,pnm,str(customer))
    _sfdc_debug["match_label"]=sfdc_label; _sfdc_debug["match_rows"]=len(sfdc_match)
    if not sfdc_match.empty:
        ec="email" if "email" in sfdc_match.columns else None
        nc="contact_name" if "contact_name" in sfdc_match.columns else None
        fc="impl_contact_flag" if "impl_contact_flag" in sfdc_match.columns else None
        pc="is_primary" if "is_primary" in sfdc_match.columns else None
        # Prefer primary contact (is_primary==1), then impl contact flag, then first row
        br=sfdc_match.iloc[0]
        if pc:
            prim=sfdc_match[sfdc_match[pc].astype(str).isin(["1","True","true","yes"])]
            if not prim.empty: br=prim.iloc[0]
        elif fc:
            fl=sfdc_match[sfdc_match[fc].astype(str).isin(["1","True","true","yes","x"])]
            if not fl.empty: br=fl.iloc[0]
        sfdc_email=str(br[ec]).strip() if ec and ec in br.index else ""
        sfdc_cname=str(br[nc]).strip() if nc and nc in br.index else ""
        if sfdc_email in ("nan","None",""): sfdc_email=""
        if sfdc_cname in ("nan","None",""): sfdc_cname=""

        # Collect remaining impl contacts for CC (exclude the primary To recipient)
        if ec and fc:
            impl_rows=sfdc_match[sfdc_match[fc].astype(str).isin(["1","True","true","yes","x"])]
            for _,r in impl_rows.iterrows():
                e=str(r.get(ec,"")).strip()
                if e and e not in ("nan","None","") and e != sfdc_email:
                    sfdc_cc_emails.append(e)

        # Populate widget session state keys directly
        _sfdc_key=f"_sfdc_match_{project_id}"
        if st.session_state.get(_sfdc_key) != sfdc_email:
            st.session_state[_sfdc_key]=sfdc_email
            st.session_state["ce_to"]=sfdc_email
            st.session_state["ce_cn"]=sfdc_cname

# Sales Rep from DRS → add to CC
_sales_rep_email=""
if sales_rep_col:
    _sr=str(sel.get(sales_rep_col,"")).strip()
    if _sr and _sr not in ("","nan","None"):
        # Try to build Zone email from name — same pattern as consultant email
        _sales_rep_email=_consultant_email(_sr) if "@" not in _sr else _sr


# ── Build default CC list (consultant + SFDC additional contacts + Sales Rep) ─
def _default_cc() -> str:
    parts = [_consultant_email(_logged_in)]
    for e in sfdc_cc_emails:
        if e not in parts: parts.append(e)
    if _sales_rep_email and _sales_rep_email not in parts:
        parts.append(_sales_rep_email)
    return ", ".join(parts)

# SFDC debug expander — shows when SFDC loaded but no match found
if _sfdc_debug and not sfdc_label:
    with st.expander("🔍 SFDC match debug (remove once confirmed working)", expanded=True):
        st.json(_sfdc_debug)

# Auto-context (shared)
_disp=_flip_name(_logged_in)

# Preview state
if "ce_ss_stamp" not in st.session_state:
    st.session_state["ce_ss_stamp"] = True

# ── Send footer helper (rendered inside compose col after tabs) ───────────────
def _send_footer(tab_key, ss_field_label, subj, body, recip_val, cc_val):
    missing=_missing_phs(body+subj)
    can_send=bool(recip_val and "@" in recip_val) and not missing
    st.markdown('<div class="send-footer">',unsafe_allow_html=True)
    if can_send:
        st.markdown('<div class="footer-ready">Ready to send</div>',unsafe_allow_html=True)
    else:
        reasons=[]
        if not recip_val or "@" not in recip_val: reasons.append("recipient email")
        if missing: reasons.extend(missing)
        st.markdown('<div class="footer-blocked">Complete before sending:</div>',unsafe_allow_html=True)
        st.markdown(f'<div class="footer-missing">{" · ".join(reasons)}</div>',unsafe_allow_html=True)

    # Checkbox row
    if ss_field_label:
        st.session_state["ce_ss_stamp"] = st.checkbox(
            f"Date-stamp **{ss_field_label}** in Smartsheet",
            value=st.session_state.get("ce_ss_stamp", True),
            key=f"ss_chk_{tab_key}",
        )
    # Button row — side by side, full width
    btn_copy, btn_send = st.columns([1, 2])
    with btn_copy:
        st.button("Copy text", key=f"copy_{tab_key}", use_container_width=True)
    with btn_send:
        lbl = "Send & log" if st.session_state.get("_gmail_approved") else "📋 Log Send"
        clicked = st.button(
            lbl, key=f"send_{tab_key}", type="primary",
            use_container_width=True, disabled=not recip_val,
        )
    st.markdown('</div>', unsafe_allow_html=True)
    return clicked

# ══════════════════════════════════════════════════════
with compose_col:
    # ── Template type selector — drives everything below deterministically ────
    _TMPL_TYPES = ["Welcome", "Post-Session", "Lifecycle (UAT → Closure)"]
    _tmpl_type = st.selectbox(
        "Template type", _TMPL_TYPES,
        index=_TMPL_TYPES.index(st.session_state.get("_ce_tmpl_type","Welcome")),
        key="ce_tmpl_type",
    )
    st.session_state["_ce_tmpl_type"] = _tmpl_type
    # Also keep _ce_tab in sync for journey rail / filter compatibility
    _tab_key_map = {"Welcome":"Welcome","Post-Session":"Post-Session","Lifecycle (UAT → Closure)":"Lifecycle"}
    st.session_state["_ce_tab"] = _tab_key_map.get(_tmpl_type,"Welcome")

    st.markdown("---")

    # ── Shared recipient block (single set of widgets, no duplication) ────────
    st.markdown('<p class="ce-label">Recipient</p>',unsafe_allow_html=True)
    st.markdown('<div class="ce-card">',unsafe_allow_html=True)
    if sfdc_label:
        st.markdown(f'<div style="font-size:11px;margin-bottom:4px"><span class="pill-ok">✓ {sfdc_label}</span></div>',unsafe_allow_html=True)
    else:
        _sfdc_loaded=st.session_state.get("df_sfdc")
        _sfdc_msg="No SFDC match — enter manually" if _sfdc_loaded is not None else "SFDC contacts not loaded"
        st.markdown(f'<div style="font-size:11px;margin-bottom:4px"><span class="pill-warn">{_sfdc_msg}</span></div>',unsafe_allow_html=True)
    st.markdown('<div style="font-size:11px;opacity:.7;margin-bottom:2px">To (recipient email)</div>',unsafe_allow_html=True)
    recip=st.text_input("To",value=sfdc_email,placeholder="customer@example.com",key="ce_to",label_visibility="collapsed")
    cname=st.text_input("Contact name",value=sfdc_cname,placeholder="First name",key="ce_cn")
    cc_in=st.text_input("CC",value=_default_cc(),key="ce_cc")
    cc_emails=[e.strip() for e in cc_in.split(",") if e.strip()]
    st.markdown('</div>',unsafe_allow_html=True)

    # Live context from current widget values
    _live_cname=st.session_state.get("ce_cn",cname) or cname
    auto_ctx=build_auto_context(sel,_disp,{"contact_name":_live_cname} if _live_cname else None)
    if _live_cname: auto_ctx["CUSTOMER_CONTACT_NAME"]=_live_cname
    auto_ctx["SENDER"]=_disp; auto_ctx["CONSULTANT_NAME"]=_disp

    # ═══ WELCOME ═══════════════════════════════════════════════════════════════
    if _tmpl_type == "Welcome":
        # Multi-product consolidated toggle
        consolidated=False; all_rows=[sel]
        if siblings:
            consolidated=st.checkbox(f"Consolidated email — {len(siblings)+1} Project Types for {customer}",key="ce_con")
            if consolidated:
                all_rows=[sel]+[_row_dict(df.iloc[sid_map[s]]) for s in siblings]
                prods=[str(r.get(prod_col,"")) for r in all_rows if r.get(prod_col) and str(r.get(prod_col)) not in ("nan","None")]
                st.markdown(f'<div class="ce-tip">Consolidated: {", ".join(prods)}</div>',unsafe_allow_html=True)

        if consolidated and len(all_rows)>1:
            tmpl_w=None
            for r in all_rows:
                k=_sku(str(r.get(prod_col,"")))
                if k and "_" in k: tmpl_w=get_welcome_template(k); break
            if not tmpl_w:
                prods=[r.get(prod_col,"") for r in all_rows if r.get(prod_col)]
                tmpl_w=get_welcome_template(_sku(prods[0])) if prods else None
        else:
            _sk=_sku(str(product_raw)) if product_raw and str(product_raw) not in ("","nan","None") else None
            tmpl_w=get_welcome_template(_sk) if _sk else None

        if not tmpl_w:
            st.caption(f"No auto-match for '{product_raw}'. Select manually:")
            opts=list_welcome_templates()
            ch=st.selectbox("Template",[t["display_name"] for t in opts],key="w_manual")
            tmpl_w=get_welcome_template(next(t["sku_key"] for t in opts if t["display_name"]==ch))

        var=st.radio("Sender variant",
                     ["Variant A — PM or automated","Variant B — Consultant sends"],
                     horizontal=True,key="w_var")
        vk="variant_a" if "A" in var else "variant_b"
        subj_w,body_w=render_template(tmpl_w[vk]["body"],tmpl_w["subject"],auto_ctx)
        auto_vals_w={v for v in auto_ctx.values() if v and str(v).strip() and len(str(v))>2 and "{" not in str(v)}
        lib_meta=_welcome_library(); ssf_w=lib_meta.get("ss_milestone_on_send")

        # Pre-send checks
        iv=sel.get(intro_col,"") if intro_col else ""
        _intro_done_w=iv and str(iv).strip() not in ("","None","nan","NaT")
        st.markdown('<p class="ce-label" style="margin-top:8px">Pre-send checks</p>',unsafe_allow_html=True)
        for ok,msg in [
            (bool(sfdc_email),"SFDC contact linked" if sfdc_email else "No SFDC match — enter recipient manually"),
            (bool(recip and "@" in recip),"Recipient email set" if recip and "@" in recip else "Recipient email missing"),
            (bool(tmpl_w),f"Template: {tmpl_w['display_name']}"),
            (not _intro_done_w,"Welcome email not yet sent — ready" if not _intro_done_w else f"Welcome already sent {iv} — check before resending"),
        ]:
            cls="chk-ok" if ok else "chk-bad"
            st.markdown(f'<div class="{cls}">{"✓" if ok else "✗"}&nbsp; {msg}</div>',unsafe_allow_html=True)

        # Write preview state — deterministic, no tab race condition
        st.session_state["ce_prev_subj"]=subj_w
        st.session_state["ce_prev_body"]=body_w
        st.session_state["ce_prev_auto"]=auto_vals_w
        st.session_state["ce_prev_ssf"]=ssf_w

        if _send_footer("w",ssf_w,subj_w,body_w,recip,cc_in):
            st.session_state["_req_w"]={"subj":subj_w,"body":body_w,"ssf":ssf_w}
            st.rerun()
        if st.session_state.get("_req_w"):
            r=st.session_state.pop("_req_w")
            try:
                with st.spinner("Logging…"):
                    ok,sid=execute_send(project_id=project_id,
                        template_id=f"welcome_{tmpl_w.get('sku_key','manual')}",
                        template_name=f"Welcome — {tmpl_w['display_name']}",
                        subject=r["subj"],body=r["body"],recipient_email=recip,
                        cc_emails=cc_emails,ss_milestone_field=r["ssf"])
                if ok:
                    st.success(f"✓ Logged — ID: `{sid}`")
                    if r["ssf"] and st.session_state.get("ce_ss_stamp",True):
                        if _do_write(project_id,r["ssf"],datetime.date.today(),sel): mark_ss_writeback_done(sid)
                else: st.error(f"Failed: {sid}")
            except Exception as ex: st.error(f"Error: {ex}"); st.exception(ex)

    # ═══ POST-SESSION ══════════════════════════════════════════════════════════
    elif _tmpl_type == "Post-Session":
        psk=_ps_key(str(product_raw))
        if not psk:
            st.info(f"No post-session templates for '{product_raw}'.")
        else:
            sessions=get_post_session_templates(psk)
            sopts={s["id"]:(f"Session {s['session_number']} — {s['name']}"+(f" [{s['variant_note']}]" if s.get("variant_note") else ""),s) for s in sessions}
            cid=st.selectbox("Session",list(sopts.keys()),format_func=lambda k:sopts[k][0],key="s_pick")
            _,tmpl_s=sopts[cid]
            st.caption(f"Audience: {tmpl_s.get('audience','Full project team')}")
            mctx: dict={}
            if tmpl_s.get("editable_fields"):
                st.markdown('<p class="ce-label" style="margin-top:8px">Fill in details</p>',unsafe_allow_html=True)
                _nf=sum(1 for f in tmpl_s["editable_fields"] if f.get("required") and not st.session_state.get(f"s_{cid}_{f['key']}",""))
                if _nf: st.markdown(f'<div style="font-size:11px;color:#dc2626;margin-bottom:6px">{_nf} required field(s) missing</div>',unsafe_allow_html=True)
                for f in tmpl_s["editable_fields"]:
                    k,lb,ft=f["key"],f["label"],f.get("type","text"); req=f.get("required",False); ph=f.get("placeholder","")
                    lbl_d=lb+(" *" if req else "")
                    if ft=="text":       v=st.text_input(lbl_d,placeholder=ph,key=f"s_{cid}_{k}")
                    elif ft=="textarea": v=st.text_area(lbl_d,placeholder=ph,height=70,key=f"s_{cid}_{k}")
                    elif ft=="multiselect":
                        s2=st.multiselect(lbl_d,options=f.get("options",[]),key=f"s_{cid}_{k}")
                        v="\n".join(f"  • {o}" for o in s2)
                    elif ft=="select": v=st.selectbox(lbl_d,f.get("options",[]),key=f"s_{cid}_{k}")
                    else: v=st.text_input(lbl_d,key=f"s_{cid}_{k}")
                    mctx[k]=v
                    if k=="GO_LIVE_READINESS" and v:
                        rm=tmpl_s.get("go_live_readiness_text",{})
                        res=rm.get(v[0],v)
                        if "{HYPERCARE_DATE}" in res: res=res.replace("{HYPERCARE_DATE}",mctx.get("HYPERCARE_DATE","{HYPERCARE_DATE}"))
                        mctx["GO_LIVE_READINESS_TEXT"]=res
            auto_ctx_s={**auto_ctx}
            subj_s,body_s=render_template(tmpl_s["body"],tmpl_s["subject"],{},{**auto_ctx_s,**mctx})
            auto_vals_s={v for v in {**auto_ctx_s,**mctx}.values() if v and str(v).strip() and len(str(v))>2 and "{" not in str(v)}
            ssf_s=tmpl_s.get("ss_milestone_on_send")

            st.session_state["ce_prev_subj"]=subj_s
            st.session_state["ce_prev_body"]=body_s
            st.session_state["ce_prev_auto"]=auto_vals_s
            st.session_state["ce_prev_ssf"]=ssf_s

            if _send_footer("s",ssf_s,subj_s,body_s,recip,cc_in):
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
                        if r["ssf"] and st.session_state.get("ce_ss_stamp",True):
                            if _do_write(project_id,r["ssf"],datetime.date.today(),sel): mark_ss_writeback_done(sid)
                    else: st.error(f"Failed: {sid}")
                except Exception as ex: st.error(f"Error: {ex}"); st.exception(ex)

    # ═══ LIFECYCLE ═════════════════════════════════════════════════════════════
    elif _tmpl_type == "Lifecycle (UAT → Closure)":
        lc_all=list_lifecycle_templates()
        lc_opts={t["id"]:t for t in lc_all}
        lcid=st.selectbox("Template",list(lc_opts.keys()),
                          format_func=lambda k:f"[{lc_opts[k]['category']}] {lc_opts[k]['name']}",key="lc_pick")
        tmpl_l=get_lifecycle_template(lcid)
        st.caption(f"When to send: {tmpl_l['trigger']}")
        for tip in tmpl_l.get("tips",[]): st.markdown(f'<div class="ce-tip">💡 {tip}</div>',unsafe_allow_html=True)
        vbody=tmpl_l.get("body","")
        if tmpl_l.get("variants"):
            vlbls={v["key"]:f"{v['label']} — {v['description']}" for v in tmpl_l["variants"]}
            cv=st.radio("Scenario",list(vlbls.keys()),format_func=lambda k:vlbls[k],key=f"lv_{lcid}",label_visibility="collapsed")
            vbody=tmpl_l["variant_bodies"][cv]
        mctx_l: dict={}
        _lc_fields=tmpl_l.get("editable_fields",[])
        if _lc_fields:
            st.markdown('<p class="ce-label" style="margin-top:8px">Fill in details</p>',unsafe_allow_html=True)
            _nf_l=sum(1 for f in _lc_fields if f.get("required") and not st.session_state.get(f"l_{lcid}_{f['key']}",""))
            if _nf_l: st.markdown(f'<div style="font-size:11px;color:#dc2626;margin-bottom:6px">{_nf_l} required field(s) missing</div>',unsafe_allow_html=True)
            for f in _lc_fields:
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
                tag=""
                if default and default not in ("","None"): tag=' <span style="font-size:10px;background:rgba(22,163,74,.12);color:#15803d;padding:1px 5px;border-radius:8px">from project</span>'
                st.markdown(f'<div style="font-size:12px;margin-bottom:3px">{lb}{" *" if req else ""}{tag}</div>',unsafe_allow_html=True)
                v=st.text_input("",value=default,placeholder="YYYY-MM-DD" if f.get("type")=="date" else "",
                                key=f"l_{lcid}_{k}",label_visibility="collapsed")
                mctx_l[k]=v

        auto_ctx_l={**auto_ctx}
        subj_l,body_l=render_template(vbody,tmpl_l["subject"],{},{**auto_ctx_l,**mctx_l})
        auto_vals_l={v for v in {**auto_ctx_l,**mctx_l}.values() if v and str(v).strip() and len(str(v))>2 and "{" not in str(v)}
        ssf_l=tmpl_l.get("ss_milestone_on_send"); gls=mctx_l.get("GO_LIVE_DATE","")

        st.session_state["ce_prev_subj"]=subj_l
        st.session_state["ce_prev_body"]=body_l
        st.session_state["ce_prev_auto"]=auto_vals_l
        st.session_state["ce_prev_ssf"]=ssf_l

        if _send_footer("l",ssf_l if isinstance(ssf_l,str) else (ssf_l[0] if ssf_l else None),subj_l,body_l,recip,cc_in):
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
                    if r["ssf"] and st.session_state.get("ce_ss_stamp",True):
                        try: gld=datetime.date.fromisoformat(r["gls"][:10]) if r["gls"] else datetime.date.today()
                        except: gld=datetime.date.today()
                        for f in (r["ssf"] if isinstance(r["ssf"],list) else [r["ssf"]]):
                            wd=gld if f in (SS_GO_LIVE_DATE,SS_PROD_CUTOVER) else datetime.date.today()
                            if _do_write(project_id,f,wd,sel): mark_ss_writeback_done(sid)
                else: st.error(f"Failed: {sid}")
            except Exception as ex: st.error(f"Error: {ex}"); st.exception(ex)

# ── Preview column ─────────────────────────────────────────────────────────────
with preview_col:
    st.markdown('<p class="ce-label">Live Preview</p>',unsafe_allow_html=True)

    # Read preview state — now deterministic since only one form renders at a time
    _ps=st.session_state.get("ce_prev_subj","")
    _pb=st.session_state.get("ce_prev_body","")
    _pa=st.session_state.get("ce_prev_auto",set())
    _recip_display=st.session_state.get("ce_to","")
    _cc_display=st.session_state.get("ce_cc","")

    if _pb:
        st.markdown(_email_html(_ps,_pb,_recip_display,_cc_display,_pa),unsafe_allow_html=True)
        # Show CC sources as a caption so consultant knows why those addresses are there
        _cc_parts=[]
        if _consultant_email(_logged_in): _cc_parts.append("you")
        if sfdc_cc_emails: _cc_parts.append(f"{len(sfdc_cc_emails)} other SFDC contact(s)")
        if _sales_rep_email: _cc_parts.append("Sales Rep")
        if len(_cc_parts)>1:
            st.caption(f"CC includes: {', '.join(_cc_parts)}")
        with st.expander("Edit before sending"):
            st.text_area("Subject",value=_ps,key="prev_subj_edit",height=40)
            st.text_area("Body",value=_pb,key="prev_body_edit",height=300)
    else:
        # Placeholder — shown before any template has rendered
        _has_proj = bool(sel.get(name_col,"")) if name_col else False
        if _has_proj:
            st.markdown(
                '<div class="ce-card" style="text-align:center;padding:32px 20px">'
                '<div style="font-size:22px;margin-bottom:8px;opacity:.3">✉</div>'
                '<div style="font-size:13px;opacity:.5">Preview will appear here once the template loads.</div>'
                '<div style="font-size:11px;opacity:.35;margin-top:6px">Switch tabs above to load a template.</div>'
                '</div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                '<div class="ce-card" style="text-align:center;padding:32px 20px">'
                '<div style="font-size:22px;margin-bottom:8px;opacity:.3">✉</div>'
                '<div style="font-size:13px;opacity:.5">Select a project above to begin.</div>'
                '</div>',
                unsafe_allow_html=True
            )

    # Session log
    log=get_session_send_log()
    proj_log=[e for e in log if e["project_id"]==project_id]
    if proj_log:
        st.markdown('<p class="ce-label" style="margin-top:18px">Sent This Session</p>',unsafe_allow_html=True)
        for e in proj_log:
            dt=e["sent_at"][:16].replace("T"," ")
            st.markdown(
                f'<div class="log-row"><span class="pill-ok">✓ Sent</span>&nbsp;'
                f'<b>{e["template_name"]}</b><br>'
                f'<span style="font-size:11px;opacity:.6">{dt} UTC → {e["recipient_email"]}</span>'
                f'</div>',unsafe_allow_html=True)
