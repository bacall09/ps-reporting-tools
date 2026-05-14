"""
pages/15_Customer_Engagement.py  v6
Architecture: Customer-first → project cards → journey rail → compose/preview
- Customer selectbox (all unique customers from DRS)
- Project cards: yours (blue border) + other consultants' (grey, reduced opacity)
- Consolidated welcome checkbox in project section header
- Journey rail: CSS var colours (theme-aware) + #4472C4 active (brand)
- Communication type + template selectboxes side-by-side
- Single recipient block, no tab race conditions
"""
import streamlit as st
import pandas as pd
import datetime
import re
from rapidfuzz import fuzz

st.session_state["current_page"] = "Customer Engagement"

st.markdown(
    "<div style='background:#050D1F;padding:28px 32px 24px;border-radius:10px;margin-bottom:16px;font-family:Manrope,sans-serif;'>"
    "<div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#3B9EFF;margin-bottom:8px;'>Professional Services · Customer Engagement</div>"
    "<h1 style='color:#fff;margin:0;font-size:26px;font-weight:700;font-family:Manrope,sans-serif;'>Lifecycle Email Composer</h1>"
    "<p style='color:rgba(255,255,255,0.4);margin:6px 0 0;font-size:13px;font-family:Manrope,sans-serif;'>Compose and track customer lifecycle communications</p>"
    "</div>",
    unsafe_allow_html=True,
)

# ── CSS — runbook compliant ───────────────────────────────────────────────────
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap');
html,body,[class*="css"]{font-family:'Manrope',sans-serif!important}
.ce-label{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.9px;color:#4472C4;margin:0 0 5px}
.ce-card{border:1px solid rgba(128,128,128,.22);border-radius:8px;padding:12px 16px;margin-bottom:10px;color:inherit}
.ce-tip{border-left:3px solid #4472C4;border-radius:0;padding:7px 12px;font-size:12px;margin-bottom:8px;color:inherit}

/* Project cards — no background, inherits theme */
.proj-card{border:0.5px solid rgba(128,128,128,.25);border-radius:8px;padding:10px 12px;margin-bottom:6px;color:inherit;cursor:pointer}
.proj-card.mine{border:1.5px solid rgba(68,114,196,.5)}
.proj-card.mine.selected{border:2px solid #4472C4}
.proj-card.other{opacity:.55}
.proj-name{font-size:12px;font-weight:600;color:inherit;margin-bottom:2px}
.proj-meta{font-size:11px;opacity:.65}
.proj-consultant{font-size:10px;opacity:.55;margin-top:4px;display:flex;align-items:center;gap:5px}
.avatar{width:18px;height:18px;border-radius:50%;background:rgba(68,114,196,.18);color:#4472C4;font-size:8px;font-weight:700;display:inline-flex;align-items:center;justify-content:center;flex-shrink:0}
.avatar.other{background:rgba(128,128,128,.15);color:inherit}
@media(prefers-color-scheme:dark){.avatar{color:#93b4e8}}
.stApp[data-theme="dark"] .avatar{color:#93b4e8}

/* Journey rail — theme-aware (no hardcoded backgrounds) */
.journey-rail{display:flex;border:0.5px solid rgba(128,128,128,.25);border-radius:8px;overflow:hidden;margin:14px 0 18px}
.sj{flex:1;padding:14px 14px 12px;border-right:0.5px solid rgba(128,128,128,.18);min-width:0;color:inherit}
.sj:last-child{border-right:none}
.sj-num{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;margin-bottom:5px;color:rgba(128,128,128,.6)}
.sj-lbl{font-size:13px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;letter-spacing:-.1px}
.sj-date{font-size:11px;margin-top:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;opacity:.65;font-weight:400}
.sj.done{background:rgba(34,197,94,.10)}
.sj.done .sj-num{color:#16a34a}
.sj.done .sj-lbl{color:#16a34a;font-weight:600}
.sj.done .sj-date{color:#16a34a}
@media(prefers-color-scheme:dark){.sj.done{background:rgba(34,197,94,.08)}.sj.done .sj-num,.sj.done .sj-lbl,.sj.done .sj-date{color:#7ed4a4}}
.stApp[data-theme="dark"] .sj.done{background:rgba(34,197,94,.08)}
.stApp[data-theme="dark"] .sj.done .sj-num,.stApp[data-theme="dark"] .sj.done .sj-lbl,.stApp[data-theme="dark"] .sj.done .sj-date{color:#7ed4a4}
.sj.active{background:rgba(68,114,196,.12);border-bottom:2px solid #4472C4}
.sj.active .sj-num{color:#4472C4}
.sj.active .sj-lbl{color:#4472C4;font-weight:600}
@media(prefers-color-scheme:dark){.sj.active{background:rgba(68,114,196,.15)}}
.stApp[data-theme="dark"] .sj.active{background:rgba(68,114,196,.15)}
.sj.locked{opacity:.32}

/* Pills — runbook pattern: rgba backgrounds + dark mode text */
.pill-ok{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:500;background:rgba(34,197,94,.14);color:#15803d}
.pill-warn{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:500;background:rgba(239,68,68,.14);color:#dc2626}
.pill-info{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:500;background:rgba(68,114,196,.14);color:#4472C4}
.pill-gray{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:500;background:rgba(128,128,128,.12);color:inherit;opacity:.7}
@media(prefers-color-scheme:dark){
  .pill-ok{color:#7ed4a4}.pill-warn{color:#fca5a5}.pill-info{color:#93b4e8}
}
.stApp[data-theme="dark"] .pill-ok{color:#7ed4a4}
.stApp[data-theme="dark"] .pill-warn{color:#fca5a5}
.stApp[data-theme="dark"] .pill-info{color:#93b4e8}

/* Pre-send checks */
.chk-ok{font-size:12px;color:#16a34a;display:flex;align-items:center;gap:6px;padding:2px 0}
.chk-bad{font-size:12px;color:#dc2626;display:flex;align-items:center;gap:6px;padding:2px 0}
@media(prefers-color-scheme:dark){.chk-ok{color:#7ed4a4}.chk-bad{color:#fca5a5}}
.stApp[data-theme="dark"] .chk-ok{color:#7ed4a4}
.stApp[data-theme="dark"] .chk-bad{color:#fca5a5}

/* Send footer */
.send-footer{border-top:1px solid rgba(128,128,128,.18);padding:12px 0;margin-top:10px}
.footer-ready{font-size:12px;font-weight:600;color:#16a34a;margin-bottom:4px}
.footer-blocked{font-size:12px;font-weight:600;color:#dc2626;margin-bottom:2px}
.footer-missing{font-size:11px;color:#dc2626;margin-bottom:6px}
@media(prefers-color-scheme:dark){.footer-ready{color:#7ed4a4}.footer-blocked,.footer-missing{color:#fca5a5}}
.stApp[data-theme="dark"] .footer-ready{color:#7ed4a4}
.stApp[data-theme="dark"] .footer-blocked,.stApp[data-theme="dark"] .footer-missing{color:#fca5a5}

/* Session log */
.log-row{border-bottom:0.5px solid rgba(128,128,128,.13);padding:5px 0;font-size:12px}

/* Email preview — EXCEPTION: hardcoded white (represents real email) */
.email-preview-wrap{background:#ffffff;border-radius:8px;border:1px solid #e2e8f0;overflow:hidden}
.ep-chrome{background:#f8fafc;border-bottom:1px solid #e2e8f0;padding:7px 12px;display:flex;align-items:center;gap:5px}
.ep-dot{width:8px;height:8px;border-radius:50%;background:#e2e8f0;display:inline-block}
.ep-hdr{padding:9px 14px;border-bottom:1px solid #e2e8f0}
.ep-row{display:flex;gap:8px;font-size:11px;padding:2px 0;color:#64748b}
.ep-lbl{width:36px;color:#94a3b8;flex-shrink:0}
.ep-body{padding:16px 18px;font-size:13px;color:#0f172a;line-height:1.7;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}
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
    "implementation contact exists":"impl_contact_flag",
    "implementation contact":"impl_contact_flag","contact roles":"contact_roles",
    "opp contact role count":"role_count","partner contact":"partner_contact",
    "opportunity owner email":"account_manager_email","owner email":"account_manager_email",
    "primary":"is_primary","is primary":"is_primary",
}
_PROD_KW = ["Capture","Approvals","Reconcile","PSP","Payments","SFTP",
            "E-Invoicing","eInvoicing","CC","ZoneCapture","ZoneApprovals","ZoneReconcile"]

def _clean_acct(text):
    t=str(text).lower()
    for s in ["ltd","limited","inc","llc","plc","gmbh","the ","- za -","& co","co."]:
        t=t.replace(s," ")
    return re.sub(r"[^a-z0-9 ]"," ",t).split()

def _prod_hints(text):
    t=str(text).lower()
    return {k for k in _PROD_KW if k.lower() in t}

def _fuzzy_sfdc(df_sfdc, proj_name, acct_name):
    if df_sfdc is None or df_sfdc.empty: return pd.DataFrame(),None
    df=df_sfdc.copy()
    cm2={c.lower().strip():c for c in df.columns}
    opp_col=cm2.get("opportunity"); acc_col=cm2.get("account"); oid_col=cm2.get("opportunity_id")
    if opp_col:
        exact=df[df[opp_col].astype(str).str.lower().str.strip()==str(proj_name).lower().strip()]
        if not exact.empty: return exact,"Exact match"
    drs_words=set(_clean_acct(acct_name)); drs_prods=_prod_hints(proj_name)
    best_score=0; best_opp_id=None; best_opp_nm=None
    for _,row in df.iterrows():
        sfdc_acct=str(row.get(acc_col or "account","")); sfdc_opp=str(row.get(opp_col or "opportunity",""))
        sfdc_words=set(_clean_acct(sfdc_acct))
        word_score=len(drs_words&sfdc_words)/max(len(drs_words),1)*100
        fuzz_score=fuzz.token_set_ratio(" ".join(drs_words)," ".join(sfdc_words))
        score=max(word_score,fuzz_score*0.7)+(30 if bool(drs_prods&_prod_hints(sfdc_opp)) else 0)
        if score>best_score:
            best_score=score; best_opp_id=row.get(oid_col) if oid_col else None
            best_opp_nm=row.get(opp_col) if opp_col else None
    lbl=f"Fuzzy match ({int(best_score)}%)"
    if best_score>=60:
        if best_opp_id is not None and oid_col:
            rows=df[df[oid_col]==best_opp_id]
            if not rows.empty: return rows,lbl
        if best_opp_nm is not None and opp_col:
            rows=df[df[opp_col]==best_opp_nm]
            if not rows.empty: return rows,lbl
    if acc_col:
        df["_sc"]=df[acc_col].apply(lambda x:fuzz.token_set_ratio(str(acct_name).lower(),str(x).lower()))
        top=df[df["_sc"]>=75].sort_values("_sc",ascending=False)
        df.drop(columns=["_sc"],inplace=True,errors="ignore")
        if not top.empty: return top,"Account match"
    return pd.DataFrame(),None

# ── General helpers ───────────────────────────────────────────────────────────
def _row_dict(row):
    return {k:(None if pd.isna(v) else v) for k,v in row.items()}

def _flip_name(n):
    if "," in n:
        p=[x.strip() for x in n.split(",",1)]
        return f"{p[1]} {p[0]}"
    return n

def _initials(name):
    parts=_flip_name(str(name)).split()
    return ("".join(p[0].upper() for p in parts[:2]) if parts else "?")

def _consultant_email(name):
    try:
        from shared.constants import EMPLOYEE_ROLES
        em=EMPLOYEE_ROLES.get(name,{}).get("email","")
        if em: return em
    except Exception: pass
    return f"{re.sub(r'[^a-z0-9.]','',_flip_name(name).lower().replace(' ','.'))}@zoneandco.com"

def _missing_phs(text):
    return list(set(re.findall(r"\{[A-Z_]+\}",text)))

# ── Customer name extraction — mirrors My Projects page ───────────────────────
_PC = ["ZEP","ZoneBilling","ZBilling","ZonePayroll","ZPayroll","ZoneCapture",
       "ZoneApprovals","ZoneReconcile","ZA","ZC","ZR","ZB","ZP"]
_PW = ["Payroll","Billing","Capture","Approvals","Reconcile","Implementation",
       "Optimization","Migration","Integration","Training","Support","MSA"]

def _extract_customer_name(project_name):
    n = str(project_name).strip()
    m = re.match(r'^(.+?)\s*-\s*[A-Z]{1,4}\s*-\s*.+$', n)
    if m: return m.group(1).strip()
    _pc_pat = '|'.join(_PC)
    m = re.match(r'^(.+?)\s*-\s*(?:' + _pc_pat + r')(?:\s|$|-)', n, re.IGNORECASE)
    if m: return m.group(1).strip()
    _pw_pat = '|'.join(_PW)
    m = re.match(r'^(.+?)\s*-\s*(?:' + _pw_pat + r').+$', n, re.IGNORECASE)
    if m: return m.group(1).strip()
    if ' : ' in n: return n.split(' : ')[0].strip()
    for code in sorted(_PC, key=len, reverse=True):
        m = re.search(r'\s+' + re.escape(code) + r'(?:\s|$|-)', n, re.IGNORECASE)
        if m and m.start() > 2: return n[:m.start()].strip()
    for word in _PW:
        m = re.search(r'\s+' + re.escape(word) + r'(?:\s|$)', n, re.IGNORECASE)
        if m and m.start() > 3: return n[:m.start()].strip().rstrip('-').strip()
    return n

def _looks_like_project_name(val:str) -> bool:
    """Returns True if a value looks like a project name rather than a customer name."""
    v = str(val).strip()
    return (len(v) > 40
            or any(kw in v for kw in [" - ZA - "," - ZC - "," - ZR - "," - ZB - "])
            or any(kw.lower() in v.lower() for kw in ["Implementation","ZoneApp","ZApprovals","ZCapture"]))

def _inject_customer_subject(subject:str, customer_name:str) -> str:
    """Insert customer name into subject: 'Welcome to X!' → 'Welcome [Customer] to X!'"""
    if not customer_name or customer_name in ("nan","None",""): return subject
    cust = str(customer_name).strip()
    if subject.lower().startswith("welcome to "):
        return f"Welcome {cust} to {subject[len('Welcome to '):]}"
    if subject.lower().startswith("welcome ") and " to " not in subject.lower():
        return f"Welcome {cust} to ZoneApps!"
    return subject

_PMAP = {
    "zoneapp: capture":"ZoneCapture","zoneapp: approvals":"ZoneApprovals",
    "zoneapp: reconcile":"ZoneReconcile","zoneapp: reconcile 2.0":"ZoneReconcile_BankConnect",
    "zoneapp: reconcile with bank connectivity":"ZoneReconcile_BankConnect",
    "zoneapp: reconcile with cc import":"ZoneReconcile_CCImport",
    "zoneapp: e-invoicing":"EInvoicing","zoneapp: ap payment":"ZoneApprovals",
    "zoneapp: payments":"ZoneApprovals",
    "zoneapp: capture & e-invoicing":"ZoneCapture_EInvoicing",
    "zoneapp: capture and e-invoicing":"ZoneCapture_EInvoicing",
    "zoneapp: capture & approvals":"ZoneCapture_ZoneApprovals",
    "zoneapp: capture and approvals":"ZoneCapture_ZoneApprovals",
    "zoneapp: capture & reconcile":"ZoneCapture_ZoneReconcile",
    "zoneapp: capture and reconcile":"ZoneCapture_ZoneReconcile",
    "zonecapture":"ZoneCapture","zoneapprovals":"ZoneApprovals","zonereconcile":"ZoneReconcile",
    "zonereconcile with bank connectivity":"ZoneReconcile_BankConnect",
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

def _combined_welcome_template(rows, prod_col):
    """
    Consolidated welcome: tries all permutations of combined SKU keys,
    then falls back to a clean merge of individual prep sections.
    Returns (tmpl_obj, is_merged:bool)
    """
    from itertools import permutations as _perms
    skus = [_sku(str(r.get(prod_col,""))) for r in rows if r.get(prod_col)]
    skus = list(dict.fromkeys(s for s in skus if s))  # deduplicate, preserve order
    if not skus: return None, False
    if len(skus)==1: return get_welcome_template(skus[0]), False

    # Try all permutations of combined SKU key
    for perm in _perms(skus):
        t = get_welcome_template("_".join(perm))
        if t: return t, False

    # No combined template — build merged body from individual templates
    base = get_welcome_template(skus[0])
    if not base: return None, False

    prod_names = []
    for sku in skus:
        t = get_welcome_template(sku)
        if t: prod_names.append(t.get("display_name", sku))

    merged_bodies = {}
    for vk in ["variant_a", "variant_b"]:
        base_body = base.get(vk, {}).get("body","") if vk in base else ""
        base_lines = base_body.split("\n")

        # Find intro end (where "Before We Begin" / prep starts)
        prep_start = next(
            (i for i,l in enumerate(base_lines)
             if any(kw in l.strip().upper() for kw in ["BEFORE WE BEGIN","GETTING STARTED"])),
            len(base_lines)
        )
        # Find closing start (Key Resources / Next Steps)
        closing_start = next(
            (i for i,l in enumerate(base_lines)
             if any(kw in l.strip().upper() for kw in ["KEY RESOURCES","NEXT STEPS","WHAT HAPPENS NEXT"])),
            len(base_lines)
        )
        intro = "\n".join(base_lines[:prep_start])
        closing = "\n".join(base_lines[closing_start:])

        # Update intro to reference all products
        if prod_names:
            intro = intro.replace(prod_names[0], " and ".join(prod_names))

        # Build merged prep — shared NetSuite section once, then product-specific
        netsuite_lines = []  # shared across products
        netsuite_seen = False
        product_blocks = []

        for sku in skus:
            t = get_welcome_template(sku)
            if not t or vk not in t: continue
            body = t[vk].get("body","")
            blines = body.split("\n")
            p_start = next(
                (i for i,l in enumerate(blines)
                 if any(kw in l.strip().upper() for kw in ["BEFORE WE BEGIN","GETTING STARTED"])),
                None
            )
            p_end = next(
                (i for i,l in enumerate(blines)
                 if any(kw in l.strip().upper() for kw in ["KEY RESOURCES","NEXT STEPS","WHAT HAPPENS NEXT"])),
                len(blines)
            )
            if p_start is None: continue
            prep = blines[p_start:p_end]
            disp = t.get("display_name", sku)

            # Extract NetSuite env block (shared) and product-specific lines
            in_ns = False; prod_specific = []
            for line in prep:
                u = line.strip().upper()
                if "NETSUITE ENVIRONMENT" in u:
                    in_ns = True
                    if not netsuite_seen:
                        netsuite_lines.append(line)
                    continue
                # Detect transition from NetSuite to product-specific
                if in_ns and line.strip() and line.strip()[0].isupper() and "•" not in line:
                    in_ns = False
                if in_ns:
                    if not netsuite_seen:
                        netsuite_lines.append(line)
                else:
                    prod_specific.append(line)
            netsuite_seen = True

            if prod_specific:
                product_blocks.append(f"\n{disp}:")
                product_blocks.extend(prod_specific)

        # Assemble prep section
        prep_section = []
        prep_section.append("\nBefore We Begin")
        prep_section.append("To help us get started smoothly, please complete the following:")
        if netsuite_lines:
            prep_section.append("\nNetSuite Environment")
            prep_section.extend(l for l in netsuite_lines if l.strip() and "NETSUITE ENVIRONMENT" not in l.upper())
        prep_section.extend(product_blocks)

        merged_bodies[vk] = {"body": intro + "\n" + "\n".join(prep_section) + "\n" + closing}

    merged = dict(base)
    merged["variant_a"] = merged_bodies.get("variant_a", base.get("variant_a",{}))
    merged["variant_b"] = merged_bodies.get("variant_b", base.get("variant_b",{}))
    merged["display_name"] = f"Consolidated — {' + '.join(prod_names)}"
    merged["subject"] = base.get("subject","").replace(
        prod_names[0] if prod_names else "ZZZZ", " + ".join(prod_names)
    )
    return merged, True

# ── Journey ───────────────────────────────────────────────────────────────────
_JOURNEY=[
    {"id":"welcome",              "label":"Welcome",            "ms_col":"ms_intro_email",   "ms_alt":"Intro. Email Sent",       "calc":None},
    {"id":"post_enablement",      "label":"Post-Enablement",    "ms_col":"ms_enablement",    "ms_alt":"Enablement Session",      "calc":None},
    {"id":"post_session_1",       "label":"Post-Session #1",    "ms_col":"ms_session1",      "ms_alt":"Session #1",              "calc":None},
    {"id":"post_session_2",       "label":"Post-Session #2",    "ms_col":"ms_session2",      "ms_alt":"Session #2",              "calc":None},
    {"id":"uat_signoff",          "label":"UAT Sign-Off",       "ms_col":"ms_uat_signoff",   "ms_alt":"UAT Signoff",             "calc":None},
    {"id":"go_live",              "label":"Go-Live",            "ms_col":"ms_prod_cutover",  "ms_alt":"Prod Cutover",            "calc":None},
    {"id":"hypercare_checkin",    "label":"Hypercare Check-in", "ms_col":None,               "ms_alt":None,                      "calc":"go_live_plus_5"},
    {"id":"hypercare_closure",    "label":"Hypercare close",    "ms_col":"ms_transition",    "ms_alt":"Transition to Support",   "calc":None},
]

def _ms_date(stage, drs_row):
    """Returns sent date string for done stages, projected date for calc stages."""
    if not drs_row: return ""
    # Calculated stage — derive from another milestone
    if stage.get("calc") == "go_live_plus_5":
        gl = drs_row.get("ms_prod_cutover") or drs_row.get("Prod Cutover")
        if not gl or str(gl).strip() in ("","None","nan","NaT"): return ""
        try:
            return _add_biz_days(gl, 5).strftime("%-d %b") + " (proj.)"
        except: return ""
    if not stage.get("ms_col"): return ""
    v = drs_row.get(stage["ms_col"]) or drs_row.get(stage["ms_alt"])
    if not v or str(v).strip() in ("","None","nan","NaT"): return ""
    try: return pd.to_datetime(v).strftime("%-d %b")
    except: return str(v)[:10]

def _stage_is_done(stage, drs_row):
    """Returns True only for stages with an actual recorded date (not projected)."""
    if not drs_row or stage.get("calc"): return False
    if not stage.get("ms_col"): return False
    v = drs_row.get(stage["ms_col"]) or drs_row.get(stage["ms_alt"])
    return bool(v and str(v).strip() not in ("","None","nan","NaT"))

def _add_biz_days(start_date, n):
    d = pd.Timestamp(start_date)
    added = 0
    while added < n:
        d += pd.Timedelta(days=1)
        if d.weekday() < 5: added += 1
    return d

def _sub_biz_days(start_date, n):
    d = pd.Timestamp(start_date)
    subtracted = 0
    while subtracted < n:
        d -= pd.Timedelta(days=1)
        if d.weekday() < 5: subtracted += 1
    return d

def _calc_outreach_due(signed_date, project_start):
    """10 biz days from signed date; if within 10 biz days of project start, use start minus 5 biz days."""
    if not signed_date or str(signed_date).strip() in ("","None","nan","NaT"): return None
    try:
        candidate = _add_biz_days(signed_date, 10)
        if project_start and str(project_start).strip() not in ("","None","nan","NaT"):
            buffer = _sub_biz_days(project_start, 10)
            if candidate >= buffer:
                return _sub_biz_days(project_start, 5)
        return candidate
    except Exception: return None

def _build_journey(drs_row):
    """Customer Profile card pattern adapted for comms milestones + outreach due date."""
    lbl_s = "font-size:9px;text-transform:uppercase;letter-spacing:.5px;color:rgba(128,128,128,.5);margin-bottom:2px"
    val_s = "font-size:12px;font-weight:500;color:var(--color-text-primary)"

    statuses = ["done" if _stage_is_done(s, drs_row) else "pending" for s in _JOURNEY]
    first_p  = next((i for i, ss in enumerate(statuses) if ss == "pending"), len(_JOURNEY)-1)

    def _step_col(i):
        if statuses[i] == "done": return "#16a34a"
        if i == first_p: return "#4472C4"
        return "rgba(128,128,128,.18)"

    bar_labels = "".join(
        f'<div style="flex:1;font-size:9px;color:rgba(128,128,128,.45);text-align:center;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">{s["label"]}</div>'
        for s in _JOURNEY
    )
    bar_steps = "".join(
        f'<div style="flex:1;height:4px;border-radius:2px;background:{_step_col(i)}"></div>'
        for i in range(len(_JOURNEY))
    )

    stage_html = []
    for i, stage in enumerate(_JOURNEY):
        cls = statuses[i]
        if i == first_p and cls == "pending": cls = "active"
        elif i > first_p and cls == "pending": cls = "locked"
        if cls == "done":
            bg = "background:rgba(34,197,94,.09);"; nc = "#16a34a"; dc = "rgba(22,163,74,.7)"
        elif cls == "active":
            bg = "background:rgba(68,114,196,.11);border-bottom:2px solid #4472C4;"; nc = "#4472C4"; dc = "rgba(68,114,196,.7)"
        else:
            bg = ""; nc = "var(--color-text-secondary)"; dc = "rgba(128,128,128,.5)"
        opacity = "opacity:.32;" if cls == "locked" else ""
        num     = f"✓ 0{i+1}" if statuses[i]=="done" else (f"▼ 0{i+1}" if cls=="active" else f"0{i+1}")
        date_str= _ms_date(stage, drs_row)
        is_calc = bool(stage.get("calc"))
        sub     = f"Sent {date_str}" if (date_str and not is_calc) else (date_str if (date_str and is_calc) else ("composing" if cls=="active" else "—"))
        extra_style = "border-left:2px dashed rgba(68,114,196,.35);" if is_calc and cls not in ("locked",) else ""
        br      = "" if i == len(_JOURNEY)-1 else "border-right:0.5px solid rgba(128,128,128,.15);"
        stage_html.append(
            f'<div style="flex:1;padding:12px 12px 10px;{br}min-width:0;{opacity}{bg}{extra_style}">' +
            f'<div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px;color:{nc}">{num}</div>' +
            f'<div style="font-size:12px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:{nc}">{stage["label"]}</div>' +
            f'<div style="font-size:10px;margin-top:3px;color:{dc}">{sub}</div></div>'
        )

    def _fd(val):
        if not val or str(val).strip() in ("","None","nan","NaT"): return "—"
        try: return pd.to_datetime(val).strftime("%-d %b %Y")
        except: return str(val)[:10]

    signed_val  = drs_row.get("signed_date")       if drs_row else None
    start_val   = drs_row.get("start_date")        if drs_row else None
    go_live_val = (drs_row.get("effective_go_live_date") or drs_row.get("go_live_date")) if drs_row else None
    due_date    = _calc_outreach_due(signed_val, start_val)
    due_str     = due_date.strftime("%-d %b %Y") if due_date else "—"
    today       = pd.Timestamp.today().normalize()
    overdue     = due_date and due_date < today
    d_color     = "#dc2626" if overdue else "#4472C4"
    d_bg        = "rgba(220,38,38,.1)" if overdue else "rgba(68,114,196,.1)"

    meta = (
        f'<div style="display:flex;gap:20px;margin-top:10px;padding-top:10px;border-top:0.5px solid rgba(128,128,128,.12)">' +
        f'<div><div style="{lbl_s}">Signed date</div><div style="{val_s}">{_fd(signed_val)}</div></div>' +
        f'<div><div style="{lbl_s}">Project start</div><div style="{val_s}">{_fd(start_val)}</div></div>' +
        f'<div><div style="{lbl_s}">Est. go-live</div><div style="{val_s}">{_fd(go_live_val)}</div></div>' +
        f'<div style="margin-left:auto"><div style="{lbl_s}">Outreach due date</div>' +
        f'<div style="display:inline-block;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:600;background:{d_bg};color:{d_color}">{"⚠ " if overdue else ""}{due_str}</div></div></div>'
    )
    return (
        f'<div style="margin:14px 0 18px">' +
        f'<div style="display:flex;gap:2px;margin-bottom:3px">{bar_labels}</div>' +
        f'<div style="display:flex;gap:3px;margin-bottom:10px">{bar_steps}</div>' +
        f'<div style="display:flex;border:0.5px solid rgba(128,128,128,.22);border-radius:8px;overflow:hidden">{"".join(stage_html)}</div>' +
        meta + '</div>'
    )

# ── Writeback ─────────────────────────────────────────────────────────────────
def _do_write(project_id,ss_field,date_val,drs_row)->bool:
    if not _ss_ok or not ss_field: return False
    ss_row_id=drs_row.get("_ss_row_id") or drs_row.get("ss_row_id") if drs_row else None
    if not ss_row_id:
        st.warning("⚠️ Writeback requires DRS loaded via **Sync SS DRS data** on Home page.")
        return False
    fields=[ss_field] if isinstance(ss_field,str) else ss_field
    wrote=False; pn=str(drs_row.get("project_name",project_id)) if drs_row else project_id
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

# ── Email preview ─────────────────────────────────────────────────────────────
def _email_html(subject,body,to_email,cc_email,auto_values:set):
    # Known section headings from YAML templates — rendered as bold blue
    _SECTION_HEADS = {
        # Welcome
        "what to expect","before we begin","netsuite environment",
        "key resources","next step","next steps","your project journey",
        "zonecapture","zoneapprovals","zonereconcile","zone e-invoicing","e-invoicing",
        "zonecapture and zoneapprovals","zonecapture and zonereconcile",
        # Post-session
        "what we discussed","session recording","configuration reference",
        "action items","key highlights","uat preparation","what's next",
        # Lifecycle
        "what happens next","how to confirm","important to note",
        "hypercare support","support transition","go-live summary",
    }
    def _htmlify(text):
        out=[]
        lines=text.split("\n")
        i=0
        while i<len(lines):
            raw=lines[i]
            s=raw.strip()
            if not s:
                out.append('<div style="margin:7px 0"></div>')
                while i+1<len(lines) and not lines[i+1].strip(): i+=1
            elif s.startswith("---") and s.endswith("---"):
                # Merge separator — subtle divider with product label
                label=s.strip("- ").strip()
                out.append(
                    f'<div style="font-size:11px;font-weight:700;text-transform:uppercase;'
                    f'letter-spacing:.6px;color:#94a3b8;margin:14px 0 6px;padding-top:12px;'
                    f'border-top:1px solid #e2e8f0">{label}</div>' if label else
                    '<hr style="border:none;border-top:1px solid #e2e8f0;margin:12px 0">'
                )
            elif s==s.upper() and len(s)>3 and not any(c in s for c in "•→@./"):
                # ALL CAPS → blue uppercase section header (e.g. YOUR PROJECT JOURNEY)
                out.append(
                    f'<div style="font-size:11px;font-weight:700;text-transform:uppercase;'
                    f'letter-spacing:.8px;color:#4472C4;margin:16px 0 6px">{s}</div>'
                )
            elif s.lower() in _SECTION_HEADS:
                # Known sub-heading → bold blue, larger
                out.append(
                    f'<div style="font-size:13px;font-weight:700;color:#4472C4;'
                    f'margin:14px 0 4px">{s}</div>'
                )
            elif s.startswith("•"):
                # Bullet
                out.append(
                    f'<div style="margin:3px 0 3px 14px;color:#0f172a">{s}</div>'
                )
            elif s[0].isdigit() and len(s)>2 and s[1] in ".)":
                # Numbered list
                out.append(f'<div style="margin:3px 0 3px 14px;color:#0f172a">{s}</div>')
            elif s.startswith("Please note:") or s.startswith("Note:"):
                # Note line — slightly muted
                out.append(
                    f'<div style="font-size:12px;color:#64748b;margin:8px 0;'
                    f'border-left:3px solid #e2e8f0;padding-left:10px">{s}</div>'
                )
            elif s in ("Kind regards,","Looking forward to working together,","Best regards,"):
                out.append(f'<div style="margin-top:14px;color:#0f172a">{s}</div>')
            elif s.startswith("Professional Services") or s.startswith("Zone & Co"):
                out.append(f'<div style="color:#64748b;font-size:12px">{s}</div>')
            else:
                out.append(f'<div style="color:#0f172a">{s}</div>')
            i+=1
        return "\n".join(out)
        return "\n".join(out)
    def _highlight(html,auto_vals):
        html=re.sub(r'\{([A-Z_]+)\}',r'<span class="ep-ph">{\1}</span>',html)
        for val in sorted(auto_vals,key=len,reverse=True):
            if val and len(val)>2 and val not in ("{}","—"):
                html=re.sub(f'(?<![>\\w]){re.escape(val)}(?![<\\w])',
                            f'<span class="ep-filled">{val}</span>',html,count=1)
        return html
    body_html=_highlight(_htmlify(body),auto_values)
    missing=_missing_phs(body+subject)
    ph_bar=f'<div class="ep-bar">⚠ {len(missing)} placeholder(s) empty: {", ".join(missing)}</div>' if missing else ""
    # Large heading from subject (e.g. "Welcome to ZoneApprovals!")
    subj_heading=""
    if subject and ("welcome" in subject.lower() or "action required" in subject.lower()):
        subj_heading=f'<div style="font-size:18px;font-weight:700;color:#4472C4;margin:0 0 16px;padding-bottom:12px;border-bottom:1px solid #e2e8f0">{subject}</div>'
    return f"""<div class="email-preview-wrap">
<div class="ep-chrome"><span class="ep-dot"></span><span class="ep-dot"></span><span class="ep-dot"></span>
<span style="font-size:10px;color:#94a3b8;margin-left:6px">Gmail preview</span></div>
<div class="ep-hdr">
<div class="ep-row"><span class="ep-lbl">TO</span>{to_email or '—'}</div>
<div class="ep-row"><span class="ep-lbl">CC</span>{cc_email or '—'}</div>
<div class="ep-row"><span class="ep-lbl">SUBJ</span><strong style="color:#0f172a">{subject}</strong></div>
</div>
<div class="ep-body">{subj_heading}{body_html}</div>
{ph_bar}
</div>"""

# ── Build DRS dataframe ───────────────────────────────────────────────────────
df_all=_df_drs.copy()
cm={c.lower().strip():c for c in df_all.columns}
name_col  =cm.get("project_name")    or cm.get("project name")
cust_col  =(cm.get("customer") or cm.get("account")
            or cm.get("account name") or cm.get("customer name")
            or cm.get("client") or cm.get("client name"))
prod_col  =cm.get("project_type")    or cm.get("project type") or cm.get("product")
id_col    =cm.get("project_id")      or cm.get("project id")
status_col=cm.get("status")
pm_col    =cm.get("project_manager") or cm.get("project manager")
intro_col =(cm.get("intro. email sent") or cm.get("intro email sent")
            or cm.get("ms_intro_email") or cm.get("intro_email_sent"))
legacy_col=cm.get("legacy")
ss_rid_col=cm.get("_ss_row_id") or cm.get("ss_row_id") or cm.get("row_id")
sales_rep_col=cm.get("sales rep") or cm.get("sales_rep") or cm.get("account executive") or cm.get("ae")
start_col    =cm.get("start_date")   or cm.get("start date")
signed_col   =cm.get("signed_date")  or cm.get("signed date")

_ss_row_id_map:dict={}
if ss_rid_col and id_col:
    for _,r in _df_drs.iterrows():
        pid=str(r.get(id_col,"")).strip(); rid=r.get(ss_rid_col)
        if pid and rid: _ss_row_id_map[pid]=rid

# Role filter — consultants see only their own projects
if not _is_mgr and pm_col:
    _f = df_all[df_all[pm_col].apply(lambda x: name_matches(str(x), _logged_in))]
    if not _f.empty: df_all = _f

# View-as filter (managers only)
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
            _f=df_all[df_all[pm_col].apply(lambda x:name_matches(str(x),_browse))]
            if not _f.empty: df_all=_f

# Remove closed/cancelled/legacy
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

# ── Send footer helper ────────────────────────────────────────────────────────
def _send_footer(tab_key,ss_field_label,subj,body,recip_val):
    import urllib.parse
    import json as _json
    import streamlit.components.v1 as _components

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
    if ss_field_label:
        st.session_state["ce_ss_stamp"]=st.checkbox(
            f"Date-stamp **{ss_field_label}** in Smartsheet",
            value=st.session_state.get("ce_ss_stamp",True),
            key=f"ss_chk_{tab_key}",
        )

    # ── Action buttons ────────────────────────────────────────────────────────
    _cc_val = st.session_state.get("ce_cc","")
    _to_str = recip_val or ""
    _subj_enc = urllib.parse.quote(subj)
    _body_enc = urllib.parse.quote(body)
    _mailto = f"mailto:{_to_str}?subject={_subj_enc}&body={_body_enc}"
    if _cc_val:
        _mailto = f"mailto:{_to_str}?cc={urllib.parse.quote(_cc_val)}&subject={_subj_enc}&body={_body_enc}"

    # HTML for copy button with clipboard API
    _body_html = body.replace('\n','<br>')
    _h = _json.dumps(f"<div style='font-family:Manrope,Arial,sans-serif;font-size:14px;line-height:1.7'>{_body_html}</div>")
    _p = _json.dumps(body)

    btn1, btn2, btn3 = st.columns([2,2,2])
    with btn1:
        st.markdown(
            f"<a href='{_mailto}' target='_blank'>"
            f"<button style='background:#1e2c63;color:white;border:none;padding:9px 0;"
            f"border-radius:6px;font-family:Manrope,sans-serif;font-size:13px;font-weight:600;"
            f"cursor:pointer;width:100%;'>✉️ Open in Email</button></a>",
            unsafe_allow_html=True
        )
    with btn2:
        _components.html(f"""<!DOCTYPE html><html><body style="margin:0;padding:0">
<button id="cb_{tab_key}" style="background:#1e2c63;color:white;border:none;padding:9px 0;border-radius:6px;font-family:Manrope,sans-serif;font-size:13px;font-weight:600;cursor:pointer;width:100%">📋 Copy (Formatted)</button>
<span id="st_{tab_key}" style="font-size:12px;margin-left:6px"></span>
<script>
var h={_h};var p={_p};
document.getElementById("cb_{tab_key}").addEventListener("click",function(){{
  navigator.clipboard.write([new ClipboardItem({{"text/html":new Blob([h],{{type:"text/html"}}),"text/plain":new Blob([p],{{type:"text/plain"}})}})]).then(function(){{
    document.getElementById("st_{tab_key}").innerText="✅ Copied!";
    setTimeout(function(){{document.getElementById("st_{tab_key}").innerText=""}},3000);
  }}).catch(function(){{document.getElementById("st_{tab_key}").innerText="⚠️ Use Open in Email"}});
}});
</script></body></html>""", height=50)
    with btn3:
        lbl="Send & log" if st.session_state.get("_gmail_approved") else "📋 Log Send"
        clicked=st.button(lbl,key=f"send_{tab_key}",type="primary",use_container_width=True,disabled=not recip_val)

    st.markdown('</div>',unsafe_allow_html=True)
    return clicked

# ═══════════════════════════════════════════════════════════════════════════════
# ROW 1 — Customer selector
# ═══════════════════════════════════════════════════════════════════════════════

# Build customer list — from df_all (already filtered by view-as/consultant)
# This ensures the customer dropdown only shows customers for the viewed consultant.
if cust_col:
    _raw = df_all[cust_col].fillna("").astype(str).str.strip()
    _populated = _raw[~_raw.isin(["","nan","None"])]
    _coverage = len(_populated) / max(len(_raw), 1)
    if _coverage >= 0.5:
        _all_customers = sorted(
            {v for v in _populated if v not in ("","nan","None")},
            key=str.lower
        )
        _using_extracted = False
    else:
        _all_customers = sorted(
            {_extract_customer_name(str(v)) for v in df_all[name_col].dropna()
             if str(v).strip() not in ("","nan","None")} if name_col else set(),
            key=str.lower
        )
        _using_extracted = True
elif name_col:
    _all_customers = sorted(
        {_extract_customer_name(str(v)) for v in df_all[name_col].dropna()
         if str(v).strip() not in ("","nan","None")},
        key=str.lower
    )
    _using_extracted = True
else:
    _all_customers = []; _using_extracted = True

if not _all_customers:
    st.warning("No customer data found in DRS. Check the 'Account Name' / 'Customer' column.")
    st.stop()

# Clear customer + project state when view-as consultant changes
_cur_browse = st.session_state.get("_browse_passthrough") or st.session_state.get("home_browse","")
if st.session_state.get("_ce_last_browse") != _cur_browse:
    st.session_state["_ce_last_browse"] = _cur_browse
    for k in list(st.session_state.keys()):
        if any(k.startswith(p) for p in ["_ce_proj","ce_proj","ce_to","ce_cn","ce_cc",
                                          "_sfdc_match","_last_proj","ce_ss_stamp"]):
            st.session_state.pop(k, None)
    st.session_state.pop("_ce_customer", None)
    st.session_state.pop("ce_customer", None)
    st.rerun()

_prev_cust=st.session_state.get("_ce_customer")
_def_cust=_prev_cust if _prev_cust in _all_customers else _all_customers[0]
st.markdown('<p class="ce-label" style="margin-bottom:4px">Select customer</p>', unsafe_allow_html=True)
selected_customer=st.selectbox(
    "Customer",options=_all_customers,
    index=_all_customers.index(_def_cust),
    label_visibility="collapsed",key="ce_customer",
)
# Clear project selection when customer changes
if st.session_state.get("_ce_customer") != selected_customer:
    st.session_state["_ce_customer"] = selected_customer
    for k in list(st.session_state.keys()):
        if any(k.startswith(p) for p in ["_ce_proj","ce_to","ce_cn","ce_cc","_sfdc_match","ce_ss_stamp"]):
            st.session_state.pop(k, None)

# ── All projects for this customer ───────────────────────────────────────────
if not _using_extracted and cust_col:
    df_cust_all = _df_drs[
        _df_drs[cust_col].astype(str).str.strip() == selected_customer
    ].copy()
elif name_col:
    df_cust_all = _df_drs[
        _df_drs[name_col].apply(lambda v: _extract_customer_name(str(v)) == selected_customer)
    ].copy()
    if status_col:
        df_cust_all=df_cust_all[~df_cust_all[status_col].astype(str).str.lower().isin(["closed","cancelled","complete","completed"])]
    if legacy_col:
        df_cust_all=df_cust_all[~df_cust_all[legacy_col].astype(str).str.strip().str.lower().isin(["yes","y","true","1"])]
else:
    df_cust_all=df_all.copy()

df_cust_all=df_cust_all.reset_index(drop=True)

# Classify projects: mine vs other
# Determine the effective consultant for "mine" classification
# For managers viewing as another consultant, "mine" = that consultant's projects
_browse = st.session_state.get("_browse_passthrough") or st.session_state.get("home_browse","")
_view_as_name = _logged_in  # default: own projects
if _is_mgr and _browse and _browse not in ("— My own view —","— Select —","👥 All team","") \
        and not (_browse.startswith("── ") and _browse.endswith(" ──")):
    _view_as_name = _browse  # viewing as a specific consultant

def _is_mine(row):
    if not pm_col: return True
    return name_matches(str(row.get(pm_col,"")),_view_as_name)

df_cust_all["_mine"]=df_cust_all.apply(_is_mine,axis=1)
_mine_proj=df_cust_all[df_cust_all["_mine"]].reset_index(drop=True)
_other_proj=df_cust_all[~df_cust_all["_mine"]].reset_index(drop=True)

n_mine=len(_mine_proj); n_other=len(_other_proj)

# Consolidated welcome checkbox — shown only when ≥2 mine projects exist
_consolidated=False
if n_mine>1:
    _consolidated=st.checkbox(
        f"Consolidated welcome ({n_mine} products)",
        key="ce_con",value=False
    )

# ── Project cards ─────────────────────────────────────────────────────────────
# Current selection — persistent via stable project ID
def _proj_sid(row):
    if id_col and row.get(id_col): return str(row[id_col])
    if name_col and row.get(name_col): return str(row[name_col])
    return str(row.name)

_mine_sids=[_proj_sid(_row_dict(r)) for _,r in _mine_proj.iterrows()]
_prev_sid=st.session_state.get("_ce_proj_sid")
_def_sid=_prev_sid if _prev_sid in _mine_sids else (_mine_sids[0] if _mine_sids else None)
_def_idx=_mine_sids.index(_def_sid) if _def_sid in _mine_sids else 0

if not _mine_sids:
    _va_display = _flip_name(_view_as_name) if _view_as_name != _logged_in else "you"
    st.info(f"No active projects assigned to {_va_display} for {selected_customer}.")
    st.stop()

# Build card dicts for card_selector
def _proj_icon(row):
    iv=row.get(intro_col,"") if intro_col else ""
    intro_done=iv and str(iv).strip() not in ("","None","nan","NaT")
    stat=str(row.get(status_col,"")).lower() if status_col else ""
    if "hold" in stat: return ":material/pause_circle:"
    if intro_done: return ":material/check_circle:"
    return ":material/mail:"

def _proj_desc(row):
    stat=str(row.get(status_col,"")) if status_col else ""
    start_raw=row.get(start_col) if start_col else None
    start_str=""
    if start_raw:
        try: start_str=f" · <span style=\'opacity:.55\'>Project start:</span> {pd.to_datetime(start_raw).strftime('%-d %b %Y')}"
        except: pass
    return f"<span style=\'opacity:.55\'>Status:</span> {stat}{start_str}"

def _proj_title(row):
    name=str(row.get(name_col,"")) if name_col else ""
    # Strip customer prefix for cleaner display
    for prefix in [selected_customer, selected_customer.split(" - ")[0] if " - " in selected_customer else ""]:
        if prefix and name.startswith(prefix):
            name=name[len(prefix):].lstrip(" -·—").strip()
            break
    return name or str(row.get(name_col,""))

# ── Build project cards HTML — horizontal scrolling row ──────────────────────
def _badge(row):
    stat=str(row.get(status_col,"")).lower() if status_col else ""
    iv=row.get(intro_col,"") if intro_col else ""
    intro_done=iv and str(iv).strip() not in ("","None","nan","NaT")
    if "hold" in stat:
        return "<span style=\'display:inline-block;font-size:10px;padding:1px 7px;border-radius:10px;background:rgba(214,137,16,.12);color:#854F0B;margin-top:5px\'>On hold</span>"
    if intro_done:
        return "<span style=\'display:inline-block;font-size:10px;padding:1px 7px;border-radius:10px;background:rgba(22,163,74,.12);color:#15803d;margin-top:5px\'>✓ Welcome sent</span>"
    return "<span style=\'display:inline-block;font-size:10px;padding:1px 7px;border-radius:10px;background:rgba(68,114,196,.12);color:#4472C4;margin-top:5px\'>Welcome pending</span>"

def _proj_card_html(row, sid, is_mine, is_selected):
    title=_proj_title(row)
    desc=_proj_desc(row)
    icon=_proj_icon(row)
    # Map material icon to Tabler equivalent
    icon_map={"check_circle":"circle-check","mail":"mail","pause_circle":"player-pause"}
    ti_icon=icon_map.get(icon.replace(":material/","").replace(":",""),"mail")
    badge=_badge(row)
    pm=str(row.get(pm_col,"")) if pm_col else ""
    pm_disp=_flip_name(pm) if pm else ""
    ini=_initials(pm) if pm else "?"
    short_pm=pm_disp.split(" ")[0] if pm_disp else "—"
    iv=row.get(intro_col,"") if intro_col else ""
    intro_done=iv and str(iv).strip() not in ("","None","nan","NaT")

    if is_mine:
        border = "2px solid #4472C4" if is_selected else "1.5px solid rgba(68,114,196,.45)"
        bg     = "background:rgba(68,114,196,.06);" if is_selected else ""
        icon_col = "#4472C4" if is_selected else "var(--color-text-secondary)"
        title_col= "#4472C4" if is_selected else "var(--color-text-primary)"
        cursor = "cursor:pointer;"
        opacity= ""
        avatar_bg="rgba(68,114,196,.15)"; avatar_col="#4472C4"
        try:
            _iv_str = pd.to_datetime(iv).strftime("%-d %b") if iv else ""
        except Exception:
            _iv_str = str(iv)[:10]
        consult_html=(
            f"<div style=\'display:flex;align-items:center;gap:5px;margin-top:6px\'>"
            f"<div style=\'width:16px;height:16px;border-radius:50%;background:{avatar_bg};"
            f"color:{avatar_col};font-size:8px;font-weight:500;display:inline-flex;"
            f"align-items:center;justify-content:center;flex-shrink:0\'>{ini}</div>"
            f"<span style=\'font-size:10px;color:var(--color-text-secondary)\'>{short_pm}</span></div>"
        )
    else:
        border = "1px solid rgba(128,128,128,.3)"
        bg     = ""
        icon_col= "var(--color-text-tertiary)"
        title_col="var(--color-text-secondary)"
        cursor = "cursor:default;"
        opacity= "opacity:0.42;"
        avatar_bg="rgba(128,128,128,.15)"; avatar_col="var(--color-text-tertiary)"
        assigned_badge="<span style=\'display:inline-block;font-size:10px;padding:1px 7px;border-radius:10px;background:rgba(128,128,128,.1);color:var(--color-text-tertiary);margin-top:5px\'>Assigned elsewhere</span>"
        consult_html=(
            f"<div style=\'display:flex;align-items:center;gap:5px;margin-top:6px\'>"
            f"<div style=\'width:16px;height:16px;border-radius:50%;background:{avatar_bg};"
            f"color:{avatar_col};font-size:8px;font-weight:500;display:inline-flex;"
            f"align-items:center;justify-content:center;flex-shrink:0\'>{ini}</div>"
            f"<span style=\'font-size:10px;color:var(--color-text-tertiary)\'>{pm_disp}</span></div>"
        )
        badge = assigned_badge

    onclick = f"onclick=\'this.dispatchEvent(new CustomEvent(\'card_click\',{{bubbles:true,detail:{{sid:\'{sid}\'}}}}))\'" if is_mine else ""
    return (
        f"<div style=\'flex:0 0 210px;border:{border};border-radius:12px;"
        f"padding:12px 14px;{bg}{opacity}{cursor}\' {onclick}>"
        f"<div style=\'font-size:17px;color:{icon_col};margin-bottom:6px\'>"
        f"<i class=\'ti ti-{ti_icon}\'></i></div>"
        f"<div style=\'font-size:12px;font-weight:500;color:{title_col};margin-bottom:3px;"
        f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis\'>{title}</div>"
        f"<div style=\'font-size:11px;color:var(--color-text-secondary);line-height:1.4\'>{desc}</div>"
        f"{badge}{consult_html}</div>"
    )

# Build all card HTML
_all_card_rows=[]
# Use current selectbox value for highlight — falls back to _def_sid on first render
_currently_selected = st.session_state.get("ce_proj_select", _def_sid)
if _currently_selected not in _mine_sids: _currently_selected = _def_sid
for i,(_sid,_row) in enumerate(zip(_mine_sids,[_row_dict(_mine_proj.iloc[j]) for j in range(len(_mine_sids))])):
    _all_card_rows.append(_proj_card_html(_row,_sid,True,_sid==_currently_selected))
for _,_orow in _other_proj.iterrows():
    _all_card_rows.append(_proj_card_html(_row_dict(_orow),"__other__",False,False))

_cards_html=(
    "<div style=\'display:flex;gap:8px;overflow-x:auto;padding-bottom:6px;"
    "scrollbar-width:thin;scrollbar-color:rgba(128,128,128,.2) transparent\'>"
    + "".join(_all_card_rows) +
    "</div>"
)

st.markdown('<p class="ce-label" style="margin-bottom:6px">Select project</p>',unsafe_allow_html=True)
st.markdown(_cards_html,unsafe_allow_html=True)

# Project selection via selectbox (invisible — cards are visual, this drives state)
_mine_labels={_mine_sids[i]:_proj_title(_row_dict(_mine_proj.iloc[i])) for i in range(len(_mine_sids))}
selected_sid=st.selectbox(
    "Select project",options=_mine_sids,
    format_func=lambda s:_mine_labels.get(s,s),
    index=_def_idx,key="ce_proj_select",
    label_visibility="collapsed",
)

st.session_state["_ce_proj_sid"]=selected_sid

# Clear stale state on project change
if st.session_state.get("_last_proj_sid")!=selected_sid:
    st.session_state["_last_proj_sid"]=selected_sid
    for k in list(st.session_state.keys()):
        if any(k.startswith(p) for p in ["ce_to","ce_cn","ce_cc","_sfdc_match"]):
            st.session_state.pop(k,None)

# Get selected project row
_sel_mine_idx=_mine_sids.index(selected_sid)
sel=_row_dict(_mine_proj.iloc[_sel_mine_idx])
customer=sel.get(cust_col,"") if cust_col else selected_customer
product_raw=sel.get(prod_col,"") if prod_col else ""
project_id=str(sel.get(id_col,selected_sid)) if id_col else selected_sid
if "_ss_row_id" not in sel and project_id in _ss_row_id_map:
    sel["_ss_row_id"]=_ss_row_id_map[project_id]

# Render other consultants' projects (read-only, informational)
# other-consultant cards rendered inline in card row above

# ── Row 2: Journey rail ───────────────────────────────────────────────────────
st.markdown(_build_journey(sel),unsafe_allow_html=True)

# ── Row 3: Compose (left) | Preview (right) ───────────────────────────────────
compose_col,preview_col=st.columns([1,1],gap="medium")

# ── SFDC lookup ───────────────────────────────────────────────────────────────
df_sfdc=st.session_state.get("df_sfdc")
sfdc_email=""; sfdc_cname=""; sfdc_label=None; sfdc_cc_emails:list=[]
if df_sfdc is not None and not df_sfdc.empty:
    _rn={c:_SFDC_COL_MAP[c.lower().strip()] for c in df_sfdc.columns if c.lower().strip() in _SFDC_COL_MAP}
    df_sn=df_sfdc.rename(columns=_rn)
    if "first_name" in df_sn.columns and "last_name" in df_sn.columns:
        df_sn["contact_name"]=(df_sn["first_name"].fillna("").astype(str)+" "+df_sn["last_name"].fillna("").astype(str)).str.strip()
    pnm=str(sel.get(name_col,"")) if name_col else ""
    # Use account column for SFDC match when available, extract as fallback
    _sfdc_acct = (str(sel.get(cust_col,"")).strip()
                  if cust_col and sel.get(cust_col) and str(sel.get(cust_col)) not in ("","nan","None")
                  else selected_customer)
    sfdc_match,sfdc_label=_fuzzy_sfdc(df_sn,pnm,str(_sfdc_acct))
    if not sfdc_match.empty:
        ec="email" if "email" in sfdc_match.columns else None
        nc="contact_name" if "contact_name" in sfdc_match.columns else None
        fc="impl_contact_flag" if "impl_contact_flag" in sfdc_match.columns else None
        pc="is_primary" if "is_primary" in sfdc_match.columns else None
        ac="account" if "account" in sfdc_match.columns else None

        # Light account sanity check — only filter when we have a clear better match
        # Trust _fuzzy_sfdc's output; only override if exact account match available
        if ac and sfdc_label and sfdc_label != "Exact match":
            _acct_exact = sfdc_match[
                sfdc_match[ac].astype(str).str.strip().str.lower() == str(_sfdc_acct).strip().lower()
            ]
            if not _acct_exact.empty:
                sfdc_match = _acct_exact
                sfdc_label = "Exact account match"
            # If no exact match, keep _fuzzy_sfdc's result — it already ranked by best score

        if not sfdc_match.empty:
            br=sfdc_match.iloc[0]
            # Priority: is_primary=1 first, then impl_contact_flag=1, then first row
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
            if ec and fc:
                impl_rows=sfdc_match[sfdc_match[fc].astype(str).isin(["1","True","true","yes","x"])]
                for _,r in impl_rows.iterrows():
                    e=str(r.get(ec,"")).strip()
                    if e and e not in ("nan","None","") and e!=sfdc_email:
                        sfdc_cc_emails.append(e)

        # Always update session state on project change — clears stale values
        _sfdc_key=f"_sfdc_match_{project_id}"
        if st.session_state.get(_sfdc_key) != sfdc_email:
            st.session_state[_sfdc_key] = sfdc_email
            st.session_state["ce_to"] = sfdc_email
            st.session_state["ce_cn"] = sfdc_cname

# Sales Rep from DRS
_sales_rep_email=""
if sales_rep_col:
    _sr=str(sel.get(sales_rep_col,"")).strip()
    if _sr and _sr not in ("","nan","None"):
        _sales_rep_email=_consultant_email(_sr) if "@" not in _sr else _sr

def _default_cc()->str:
    parts=[_consultant_email(_logged_in)]
    for e in sfdc_cc_emails:
        if e not in parts: parts.append(e)
    if _sales_rep_email and _sales_rep_email not in parts: parts.append(_sales_rep_email)
    return ", ".join(parts)

_disp=_flip_name(_logged_in)
if "ce_ss_stamp" not in st.session_state: st.session_state["ce_ss_stamp"]=True

with compose_col:
    # ── Communication type + template selector ────────────────────────────────
    _STAGE_TO_COMM={
        "welcome":"Welcome",
        "post_enablement":"Post-Session","post_session_1":"Post-Session","post_session_2":"Post-Session",
        "uat_signoff":"Lifecycle (UAT → Closure)","go_live":"Lifecycle (UAT → Closure)",
        "hypercare_checkin":"Lifecycle (UAT → Closure)","hypercare_closure":"Lifecycle (UAT → Closure)",
    }
    _statuses_d = ["done" if _stage_is_done(s,sel) else "pending" for s in _JOURNEY]
    _last_done  = max((i for i,s in enumerate(_statuses_d) if s=="done"), default=-1)
    # Next actionable stage = first pending after the last done stage
    _next_idx   = next((i for i in range(_last_done+1, len(_JOURNEY))
                        if _statuses_d[i]=="pending"), _last_done)
    _next_stage = _JOURNEY[_next_idx]["id"] if _next_idx >= 0 else "welcome"

    # Filter available comm types based on progress
    # Hide Welcome if the last done stage is past Welcome (index > 0)
    _hide_welcome = _last_done >= 1  # at least Post-Enablement is done
    _COMM_TYPES = ["Post-Session","Lifecycle (UAT → Closure)"] if _hide_welcome else ["Welcome","Post-Session","Lifecycle (UAT → Closure)"]

    # Smart default — only auto-set on project change
    _proj_changed = st.session_state.get("_last_proj_sid_for_comm") != selected_sid
    if _proj_changed:
        _smart_default = _STAGE_TO_COMM.get(_next_stage, _COMM_TYPES[0])
        # Ensure smart default is in the available options
        if _smart_default not in _COMM_TYPES:
            _smart_default = _COMM_TYPES[0]
        st.session_state["_ce_tmpl_type"] = _smart_default
        st.session_state["_last_proj_sid_for_comm"] = selected_sid
    st.markdown('<p class="ce-label" style="margin-top:8px;margin-bottom:4px">Communication type</p>', unsafe_allow_html=True)
    _cur_type = st.session_state.get("_ce_tmpl_type", _COMM_TYPES[0])
    if _cur_type not in _COMM_TYPES:
        _cur_type = _COMM_TYPES[0]
        st.session_state["_ce_tmpl_type"] = _cur_type
    _tmpl_type=st.selectbox(
        "Communication type",options=_COMM_TYPES,
        label_visibility="collapsed",
        index=_COMM_TYPES.index(_cur_type),
        key="ce_tmpl_type",
    )
    st.session_state["_ce_tmpl_type"]=_tmpl_type
    st.session_state["_ce_tab"]={"Welcome":"Welcome","Post-Session":"Post-Session","Lifecycle (UAT → Closure)":"Lifecycle"}.get(_tmpl_type,"Welcome")
    # tmpl_col — template selectbox renders below comm type in each branch
    tmpl_col = None  # placeholder so branch code doesn't break

    # Live context — read after recipient widgets exist
    _live_cname=st.session_state.get("ce_cn","") or ""
    auto_ctx=build_auto_context(sel,_disp,{"contact_name":_live_cname} if _live_cname else None)
    if _live_cname: auto_ctx["CUSTOMER_CONTACT_NAME"]=_live_cname
    auto_ctx["SENDER"]=_disp; auto_ctx["CONSULTANT_NAME"]=_disp

    # ── Compose button — gates template rendering (prevents stale cache) ────────
    _compose_key = f"_ce_composed_{selected_sid}_{_tmpl_type}"
    _is_composed = st.session_state.get(_compose_key, False)

    # ── Send request handlers — run every render, outside branch logic ────────
    # This ensures logging fires even if template type changed since button click
    if st.session_state.get("_req_w"):
        r=st.session_state.pop("_req_w")
        try:
            with st.spinner("Logging…"):
                _tmpl_w_for_log=get_welcome_template(_sku(str(product_raw))) if product_raw and str(product_raw) not in ("","nan","None") else None
                _tid=f"welcome_{_tmpl_w_for_log.get('sku_key','manual')}" if _tmpl_w_for_log else "welcome_manual"
                _tnm=f"Welcome — {_tmpl_w_for_log.get('display_name','')}" if _tmpl_w_for_log else "Welcome"
                ok,sid=execute_send(project_id=project_id,template_id=_tid,template_name=_tnm,
                    subject=r["subj"],body=r["body"],recipient_email=recip,cc_emails=cc_emails,ss_milestone_field=r["ssf"])
            if ok:
                st.success(f"✓ Logged — ID: `{sid}`")
                if r["ssf"] and st.session_state.get("ce_ss_stamp",True):
                    if _do_write(project_id,r["ssf"],datetime.date.today(),sel): mark_ss_writeback_done(sid)
            else: st.error(f"Failed: {sid}")
        except Exception as ex: st.error(f"Error: {ex}"); st.exception(ex)

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

    if st.session_state.get("_req_l"):
        r=st.session_state.pop("_req_l")
        try:
            with st.spinner("Logging…"):
                ok,sid=execute_send(project_id=project_id,template_id=r.get("tid","lifecycle"),
                    template_name=r.get("tnm","Lifecycle"),
                    subject=r["subj"],body=r["body"],recipient_email=recip,cc_emails=cc_emails,ss_milestone_field=r["ssf"])
            if ok:
                st.success(f"✓ Logged — ID: `{sid}`")
                if r["ssf"] and st.session_state.get("ce_ss_stamp",True):
                    try: gld=datetime.date.fromisoformat(r["gls"][:10]) if r.get("gls") else datetime.date.today()
                    except: gld=datetime.date.today()
                    for _sf in (r["ssf"] if isinstance(r["ssf"],list) else [r["ssf"]]):
                        wd=gld if _sf in (SS_GO_LIVE_DATE,SS_PROD_CUTOVER) else datetime.date.today()
                        if _do_write(project_id,_sf,wd,sel): mark_ss_writeback_done(sid)
            else: st.error(f"Failed: {sid}")
        except Exception as ex: st.error(f"Error: {ex}"); st.exception(ex)

    # ── Template selector — right after comm type, before recipient ─────────────
    st.markdown('<p class="ce-label" style="margin-top:8px;margin-bottom:2px">Template</p>', unsafe_allow_html=True)
    if _tmpl_type == "Welcome":
        _sk_pre = _sku(str(product_raw)) if product_raw and str(product_raw) not in ("","nan","None") else None
        _tmpl_pre = get_welcome_template(_sk_pre) if _sk_pre else None
        if not _tmpl_pre:
            _opts_pre = list_welcome_templates()
            # Use same key as branch — this IS the selector
            st.selectbox("Template", [t["display_name"] for t in _opts_pre], key="w_manual", label_visibility="collapsed")
        else:
            st.selectbox("Template", [_tmpl_pre.get("display_name","")], key="w_tmpl_disp", disabled=True, label_visibility="collapsed")
    elif _tmpl_type == "Post-Session":
        _psk_pre = _ps_key(str(product_raw))
        if _psk_pre:
            _sessions_pre = get_post_session_templates(_psk_pre)
            _sopts_pre = {s["id"]: f"Session {s['session_number']} — {s['name']}" for s in _sessions_pre}
            # Use same key as branch
            st.selectbox("Template", list(_sopts_pre.keys()), format_func=lambda k: _sopts_pre[k], key="s_pick", label_visibility="collapsed")
        else:
            st.caption("No post-session templates for this product type.")
    elif _tmpl_type == "Lifecycle (UAT → Closure)":
        _lc_pre = list_lifecycle_templates()
        _lc_pre_opts = {t["id"]: f"[{t['category']}] {t['name']}" for t in _lc_pre}
        # Use same key as branch
        st.selectbox("Template", list(_lc_pre_opts.keys()), format_func=lambda k: _lc_pre_opts[k], key="lc_pick", label_visibility="collapsed")

    # ── Recipient ─────────────────────────────────────────────────────────────
    st.markdown('<p class="ce-label" style="margin-top:12px">Recipient</p>',unsafe_allow_html=True)
    if sfdc_label:
                st.markdown(f'<div style="font-size:11px;margin-bottom:4px"><span class="pill-ok">✓ {sfdc_label}</span></div>',unsafe_allow_html=True)
    else:
                _msg="No SFDC match — enter manually" if df_sfdc is not None else "SFDC contacts not loaded"
                st.markdown(f'<div style="font-size:11px;margin-bottom:4px"><span class="pill-warn">{_msg}</span></div>',unsafe_allow_html=True)
    recip=st.text_input("To (recipient email)",value=sfdc_email,placeholder="customer@example.com",key="ce_to")
    cname=st.text_input("Contact name",value=sfdc_cname,placeholder="First name",key="ce_cn")
    cc_in=st.text_input("CC",value=_default_cc(),key="ce_cc")
    cc_emails=[e.strip() for e in cc_in.split(",") if e.strip()]
    # Update live context with typed contact name
    _live_cname=st.session_state.get("ce_cn",_live_cname) or _live_cname
    if _live_cname:
                auto_ctx["CUSTOMER_CONTACT_NAME"]=_live_cname
                auto_ctx=build_auto_context(sel,_disp,{"contact_name":_live_cname})
                auto_ctx["CUSTOMER_CONTACT_NAME"]=_live_cname
                auto_ctx["SENDER"]=_disp; auto_ctx["CONSULTANT_NAME"]=_disp

    # ── Welcome ───────────────────────────────────────────────────────────────
    if _tmpl_type=="Welcome":
                _all_rows=[sel]
                if _consolidated and n_mine>1:
                    _all_rows=[_row_dict(_mine_proj.iloc[i]) for i in range(n_mine)]
                _is_merged=False
                if _consolidated and n_mine>1:
                    tmpl_w,_is_merged=_combined_welcome_template(_all_rows,prod_col)
                else:
                    _sk=_sku(str(product_raw)) if product_raw and str(product_raw) not in ("","nan","None") else None
                    tmpl_w=get_welcome_template(_sk) if _sk else None
                if not tmpl_w:
                    # Template chosen via pre-flight selector above Recipient
                    opts=list_welcome_templates()
                    _w_ch=st.session_state.get("w_manual","")
                    if _w_ch:
                        tmpl_w=get_welcome_template(next((t["sku_key"] for t in opts if t["display_name"]==_w_ch),None))
                    if not tmpl_w and opts:
                        tmpl_w=get_welcome_template(opts[0]["sku_key"])
                if _is_merged:
                    st.markdown('<div class="ce-tip">Consolidated: one email covering all products. Prep sections merged.</div>',unsafe_allow_html=True)

                var=st.radio("Sender variant",["Variant A — PM or automated","Variant B — Consultant sends"],horizontal=True,key="w_var")
                vk="variant_a" if "A" in var else "variant_b"
                subj_w,body_w=render_template(tmpl_w[vk]["body"],tmpl_w["subject"],auto_ctx)
                subj_w=_inject_customer_subject(subj_w, customer)
                auto_vals_w={v for v in auto_ctx.values() if v and str(v).strip() and len(str(v))>2 and "{" not in str(v)}
                lib_meta=_welcome_library(); ssf_w=lib_meta.get("ss_milestone_on_send")
                _trigger=f"When to send: First contact after project kickoff"
                st.caption(_trigger)

                iv=sel.get(intro_col,"") if intro_col else ""
                _intro_done=iv and str(iv).strip() not in ("","None","nan","NaT")
                st.markdown('<p class="ce-label" style="margin-top:8px">Pre-send checks</p>',unsafe_allow_html=True)
                for ok,msg in [
                    (bool(sfdc_email),"SFDC contact linked" if sfdc_email else "No SFDC match — enter recipient manually"),
                    (bool(recip and "@" in recip),"Recipient email set" if recip and "@" in recip else "Recipient email missing"),
                    (bool(tmpl_w),f"Template: {tmpl_w.get('display_name','')}"),
                    (not _intro_done,"Welcome not yet sent — ready" if not _intro_done else f"Welcome already sent {iv} — check before resending"),
                ]:
                    st.markdown(f'<div class="{"chk-ok" if ok else "chk-bad"}">{"✓" if ok else "✗"}&nbsp; {msg}</div>',unsafe_allow_html=True)

                st.session_state["ce_prev_subj"]=subj_w
                st.session_state["ce_prev_body"]=body_w
                st.session_state["ce_prev_auto"]=auto_vals_w

                if _send_footer("w",ssf_w,subj_w,body_w,recip):
                    st.session_state["_req_w"]={"subj":st.session_state.get("ce_send_subj",subj_w),"body":st.session_state.get("ce_send_body",body_w),"ssf":ssf_w}
                    st.rerun()

    # ── Post-Session ──────────────────────────────────────────────────────────
    elif _tmpl_type=="Post-Session":
                psk=_ps_key(str(product_raw))
                if not psk:
                    st.info(f"No post-session templates for '{product_raw}'.")
                else:
                    sessions=get_post_session_templates(psk)
                    sopts={s["id"]:(f"Session {s['session_number']} — {s['name']}"+(f" [{s['variant_note']}]" if s.get("variant_note") else ""),s) for s in sessions}
                    # cid driven by pre-flight selector key="s_pick"
                    cid=st.session_state.get("s_pick",list(sopts.keys())[0] if sopts else None)
                    if cid not in sopts: cid=list(sopts.keys())[0] if sopts else None
                    _,tmpl_s=sopts[cid]
                    st.caption(f"Audience: {tmpl_s.get('audience','Full project team')} · {tmpl_s.get('trigger','')}")
                    mctx:dict={}
                    if tmpl_s.get("editable_fields"):
                        st.markdown('<p class="ce-label" style="margin-top:8px">Fill in details</p>',unsafe_allow_html=True)
                        _nf=sum(1 for f in tmpl_s["editable_fields"] if f.get("required") and not st.session_state.get(f"s_{cid}_{f['key']}",""))
                        if _nf: st.markdown(f'<div style="font-size:11px;color:#dc2626;margin-bottom:6px">{_nf} required field(s) missing</div>',unsafe_allow_html=True)
                        for f in tmpl_s["editable_fields"]:
                            k,lb,ft=f["key"],f["label"],f.get("type","text"); req=f.get("required",False); ph=f.get("placeholder","")
                            lbl_d=lb+(" *" if req else "")
                            _is_date = "date" in lb.lower() or "date" in k.lower() or "date/time" in lb.lower()
                            if _is_date:
                                import datetime as _dt2
                                _dv=st.session_state.get(f"s_{cid}_{k}")
                                _date_val=st.date_input(lbl_d,value=_dv if isinstance(_dv,_dt2.date) else None,key=f"s_{cid}_{k}")
                                v=_date_val.strftime("%-d %B %Y") if _date_val else ""
                            elif ft=="text":       v=st.text_input(lbl_d,placeholder=ph,key=f"s_{cid}_{k}")
                            elif ft=="textarea": v=st.text_area(lbl_d,placeholder=ph,height=70,key=f"s_{cid}_{k}")
                            elif ft=="multiselect":
                                s2=st.multiselect(lbl_d,options=f.get("options",[]),key=f"s_{cid}_{k}")
                                v="\n".join(f"  • {o}" for o in s2)
                            elif ft=="select": v=st.selectbox(lbl_d,f.get("options",[]),key=f"s_{cid}_{k}")
                            else: v=st.text_input(lbl_d,key=f"s_{cid}_{k}")
                            mctx[k]=v
                            if k=="GO_LIVE_READINESS" and v:
                                rm=tmpl_s.get("go_live_readiness_text",{}); res=rm.get(v[0],v)
                                if "{HYPERCARE_DATE}" in res: res=res.replace("{HYPERCARE_DATE}",mctx.get("HYPERCARE_DATE","{HYPERCARE_DATE}"))
                                mctx["GO_LIVE_READINESS_TEXT"]=res
                    subj_s,body_s=render_template(tmpl_s["body"],tmpl_s["subject"],{},{**auto_ctx,**mctx})
                    subj_s=_inject_customer_subject(subj_s, customer)
                    auto_vals_s={v for v in {**auto_ctx,**mctx}.values() if v and str(v).strip() and len(str(v))>2 and "{" not in str(v)}
                    ssf_s=tmpl_s.get("ss_milestone_on_send")
                    st.session_state["ce_prev_subj"]=subj_s
                    st.session_state["ce_prev_body"]=body_s
                    st.session_state["ce_prev_auto"]=auto_vals_s
                    if _send_footer("s",ssf_s,subj_s,body_s,recip):
                        st.session_state["_req_s"]={"subj":st.session_state.get("ce_send_subj",subj_s),"body":st.session_state.get("ce_send_body",body_s),"ssf":ssf_s,"tid":tmpl_s["id"],"tnm":tmpl_s["name"]}
                        st.rerun()

    # ── Lifecycle ─────────────────────────────────────────────────────────────
    elif _tmpl_type=="Lifecycle (UAT → Closure)":
                lc_all=list_lifecycle_templates()
                lc_opts={t["id"]:t for t in lc_all}
                # lcid driven by pre-flight selector key="lc_pick"
                lcid=st.session_state.get("lc_pick",list(lc_opts.keys())[0] if lc_opts else None)
                if lcid not in lc_opts: lcid=list(lc_opts.keys())[0] if lc_opts else None
                tmpl_l=get_lifecycle_template(lcid)
                st.caption(f"When to send: {tmpl_l['trigger']}")
                for tip in tmpl_l.get("tips",[]): st.markdown(f'<div class="ce-tip">💡 {tip}</div>',unsafe_allow_html=True)
                vbody=tmpl_l.get("body","")
                if tmpl_l.get("variants"):
                    vlbls={v["key"]:f"{v['label']} — {v['description']}" for v in tmpl_l["variants"]}
                    cv=st.radio("Scenario",list(vlbls.keys()),format_func=lambda k:vlbls[k],key=f"lv_{lcid}",label_visibility="collapsed")
                    vbody=tmpl_l["variant_bodies"][cv]
                mctx_l:dict={}
                _lc_fields=tmpl_l.get("editable_fields",[])
                if _lc_fields:
                    st.markdown('<p class="ce-label" style="margin-top:8px">Fill in details</p>',unsafe_allow_html=True)
                    for f in _lc_fields:
                        k,lb=f["key"],f["label"]; req=f.get("required",False)
                        fsrc=f.get("source",""); default=str(f.get("default",""))
                        if fsrc=="drs_prod_cutover":
                            raw=sel.get("prod_cutover") or sel.get("Prod Cutover")
                            if raw:
                                try: default=pd.to_datetime(raw).date().isoformat()
                                except: pass
                        elif fsrc=="drs_project_link":
                            default=str(sel.get("project_link") or sel.get("Project Link") or "")
                        elif fsrc=="calculated_go_live_plus_14":
                            glr=sel.get("prod_cutover") or sel.get("go_live_date") or mctx_l.get("GO_LIVE_DATE","")
                            if glr:
                                try: default=(pd.to_datetime(glr).date()+datetime.timedelta(days=14)).isoformat()
                                except: pass
                        tag=""
                        if default and default not in ("","None"): tag=' <span style="font-size:10px;background:rgba(22,163,74,.12);color:#15803d;padding:1px 5px;border-radius:8px">from project</span>'
                        st.markdown(f'<div style="font-size:12px;margin-bottom:3px">{lb}{" *" if req else ""}{tag}</div>',unsafe_allow_html=True)
                        _is_lc_date = f.get("type")=="date" or "date" in lb.lower()
                        if _is_lc_date:
                            _ldv=st.session_state.get(f"l_{lcid}_{k}")
                            _ldate=st.date_input(lb+(" *" if req else ""),
                                                  value=_ldv if isinstance(_ldv,datetime.date) else None,
                                                  key=f"l_{lcid}_{k}",
                                                  label_visibility="collapsed")
                            v=_ldate.strftime("%-d %B %Y") if _ldate else default
                        else:
                            v=st.text_input("",value=default,
                                            placeholder="Enter value",
                                            key=f"l_{lcid}_{k}",label_visibility="collapsed")
                        mctx_l[k]=v
                auto_ctx_l={**auto_ctx}
                subj_l,body_l=render_template(vbody,tmpl_l["subject"],{},{**auto_ctx_l,**mctx_l})
                subj_l=_inject_customer_subject(subj_l, customer)
                auto_vals_l={v for v in {**auto_ctx_l,**mctx_l}.values() if v and str(v).strip() and len(str(v))>2 and "{" not in str(v)}
                ssf_l=tmpl_l.get("ss_milestone_on_send"); gls=mctx_l.get("GO_LIVE_DATE","")
                st.session_state["ce_prev_subj"]=subj_l
                st.session_state["ce_prev_body"]=body_l
                st.session_state["ce_prev_auto"]=auto_vals_l
                if _send_footer("l",ssf_l if isinstance(ssf_l,str) else (ssf_l[0] if ssf_l else None),subj_l,body_l,recip):
                    st.session_state["_req_l"]={"subj":st.session_state.get("ce_send_subj",subj_l),"body":st.session_state.get("ce_send_body",body_l),"ssf":ssf_l,"gls":gls}
                    st.rerun()

    # ── Preview column ─────────────────────────────────────────────────────────────
with preview_col:
    st.markdown('<p class="ce-label">Live Preview</p>',unsafe_allow_html=True)
    # Base values from template render
    _ps=st.session_state.get("ce_prev_subj","")
    _pb=st.session_state.get("ce_prev_body","")
    _pa=st.session_state.get("ce_prev_auto",set())
    _recip_display=st.session_state.get("ce_to","")
    _cc_display=st.session_state.get("ce_cc","")

    # Override with manually edited values
    _ps_edit=st.session_state.get("prev_subj_edit","")
    _pb_edit=st.session_state.get("prev_body_edit","")
    _using_edit = False
    if _ps_edit and _ps_edit != _ps:
        _ps = _ps_edit; _using_edit = True
    if _pb_edit and _pb_edit != _pb:
        _pb = _pb_edit; _using_edit = True

    if not _is_composed:
        try:
            from streamlit_extras.skeleton import skeleton
            _sk = skeleton(height=320)
            _sk.markdown(
                '<div style="text-align:center;padding:40px 20px;opacity:.4">'
                '<div style="font-size:24px;margin-bottom:8px">✉</div>'
                '<div style="font-size:12px">Fill in the fields, then generate preview</div>'
                '</div>', unsafe_allow_html=True
            )
        except Exception:
            st.markdown(
                '<div style="text-align:center;padding:40px 20px">'
                '<div style="font-size:28px;margin-bottom:10px;opacity:.2">✉</div>'
                '<div style="font-size:13px;opacity:.45;margin-bottom:20px">'
                'Fill in the fields on the left, then generate the preview.</div>'
                '</div>', unsafe_allow_html=True
            )
        if st.button("✉ Generate preview",
                     key=f"btn_compose_{selected_sid}_{_tmpl_type}",
                     type="primary", use_container_width=True):
            st.session_state[_compose_key] = True
            st.rerun()
    elif _pb:
        st.markdown(_email_html(_ps,_pb,_recip_display,_cc_display,_pa),unsafe_allow_html=True)
        _cc_parts=[]
        if _consultant_email(_logged_in): _cc_parts.append("you")
        if sfdc_cc_emails: _cc_parts.append(f"{len(sfdc_cc_emails)} other SFDC contact(s)")
        if _sales_rep_email: _cc_parts.append("Sales Rep")
        if len(_cc_parts)>1: st.caption(f"CC includes: {', '.join(_cc_parts)}")
        if _using_edit:
            st.caption("✎ Preview showing your edits")
        with st.expander("Edit before sending"):
            st.caption("Edits here update the preview and will be sent instead of the template version.")
            # Reset button — clears edits and restores template version
            if _using_edit:
                if st.button("↺ Reset to template", key="reset_edits"):
                    st.session_state.pop("prev_subj_edit", None)
                    st.session_state.pop("prev_body_edit", None)
                    st.rerun()
            # Use value= from the current _ps/_pb so edits persist across rerenders
            # but reset when template changes (because ce_prev_subj changed)
            st.text_area("Subject", value=_ps, key="prev_subj_edit", height=50)
            st.text_area("Body", value=_pb, key="prev_body_edit", height=320)

        # Expose edited subject/body for send buttons to use
        st.session_state["ce_send_subj"] = _ps
        st.session_state["ce_send_body"] = _pb

    # Session log
    log=get_session_send_log()
    proj_log=[e for e in log if e["project_id"]==project_id]
    if proj_log:
        st.markdown('<p class="ce-label" style="margin-top:18px">Sent This Session</p>',unsafe_allow_html=True)
        for e in proj_log:
            dt=e["sent_at"][:16].replace("T"," ")
            st.markdown(
                f'<div class="log-row"><span class="pill-ok">✓</span>&nbsp;<b>{e["template_name"]}</b><br>'
                f'<span style="font-size:11px;opacity:.6">{dt} UTC → {e["recipient_email"]}</span>'
                f'</div>',unsafe_allow_html=True)
