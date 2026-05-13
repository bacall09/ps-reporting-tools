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

st.markdown("""
<div style='background:#050D1F;padding:20px 28px 16px;border-radius:0 0 8px 8px;margin-bottom:20px'>
  <span style='color:#4472C4;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase'>Customer Engagement</span>
  <h2 style='color:#ffffff;margin:4px 0 0;font-size:22px;font-weight:600;letter-spacing:-0.3px'>Lifecycle Email Composer</h2>
</div>
""", unsafe_allow_html=True)

# ── CSS — runbook compliant ───────────────────────────────────────────────────
st.markdown("""<style>
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
.journey-rail{display:flex;border:0.5px solid rgba(128,128,128,.25);border-radius:8px;overflow:hidden;margin:12px 0 16px}
.sj{flex:1;padding:12px 12px 10px;border-right:0.5px solid rgba(128,128,128,.18);min-width:0;color:inherit}
.sj:last-child{border-right:none}
.sj-num{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px;color:rgba(128,128,128,.7)}
.sj-lbl{font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sj-date{font-size:11px;margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;opacity:.7}
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
    {"id":"welcome",            "label":"Welcome",         "ms_col":"ms_intro_email",  "ms_alt":"Intro. Email Sent"},
    {"id":"post_session_1",     "label":"Post-Session #1", "ms_col":"ms_enablement",   "ms_alt":"Enablement Session"},
    {"id":"post_session_2",     "label":"Post-Session #2", "ms_col":"ms_session1",     "ms_alt":"Session #1"},
    {"id":"uat_signoff",        "label":"UAT Sign-Off",    "ms_col":"ms_uat_signoff",  "ms_alt":"UAT Signoff"},
    {"id":"go_live",            "label":"Go-Live",         "ms_col":"ms_prod_cutover", "ms_alt":"Prod Cutover"},
    {"id":"hypercare_closure",  "label":"Hypercare close", "ms_col":"ms_transition",   "ms_alt":"Transition to Support"},
]

def _ms_date(stage,drs_row):
    if not drs_row: return ""
    v=drs_row.get(stage["ms_col"]) or drs_row.get(stage["ms_alt"])
    if not v or str(v).strip() in ("","None","nan","NaT"): return ""
    try: return pd.to_datetime(v).strftime("%-d %b")
    except: return str(v)[:10]

def _build_journey(drs_row):
    statuses=["done" if _ms_date(s,drs_row) else "pending" for s in _JOURNEY]
    first_p=next((i for i,st in enumerate(statuses) if st=="pending"),len(_JOURNEY)-1)
    parts=['<div class="journey-rail">']
    for i,stage in enumerate(_JOURNEY):
        cls=statuses[i]
        if i==first_p and cls=="pending": cls="active"
        elif i>first_p and cls=="pending": cls="locked"
        num=f"0{i+1}" if i<9 else str(i+1)
        icon="✓ " if statuses[i]=="done" else ("▼ " if cls=="active" else "")
        date_str=_ms_date(stage,drs_row)
        sub=f"Sent {date_str}" if date_str else ("composing" if cls=="active" else "")
        parts.append(
            f'<div class="sj {cls}">'
            f'<div class="sj-num">{icon}{num}</div>'
            f'<div class="sj-lbl">{stage["label"]}</div>'
            f'<div class="sj-date">{sub}</div>'
            f'</div>'
        )
    parts.append("</div>")
    return "\n".join(parts)

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
        "what to expect","before we begin","netsuite environment",
        "key resources","next step","next steps","zonecapture","zoneapprovals",
        "zonereconcile","zone e-invoicing","e-invoicing","your project journey",
        "important to note","how to confirm","what happens next",
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
cust_col  =cm.get("customer")        or cm.get("account")
prod_col  =cm.get("project_type")    or cm.get("project type") or cm.get("product")
id_col    =cm.get("project_id")      or cm.get("project id")
status_col=cm.get("status")
pm_col    =cm.get("project_manager") or cm.get("project manager")
intro_col =(cm.get("intro. email sent") or cm.get("intro email sent")
            or cm.get("ms_intro_email") or cm.get("intro_email_sent"))
legacy_col=cm.get("legacy")
ss_rid_col=cm.get("_ss_row_id") or cm.get("ss_row_id") or cm.get("row_id")
sales_rep_col=cm.get("sales rep") or cm.get("sales_rep") or cm.get("account executive") or cm.get("ae")
start_col =cm.get("start_date") or cm.get("start date")

_ss_row_id_map:dict={}
if ss_rid_col and id_col:
    for _,r in _df_drs.iterrows():
        pid=str(r.get(id_col,"")).strip(); rid=r.get(ss_rid_col)
        if pid and rid: _ss_row_id_map[pid]=rid

# View-as filter (managers)
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
    btn_c,btn_s=st.columns([1,2])
    with btn_c:
        st.button("Copy text",key=f"copy_{tab_key}",use_container_width=True)
    with btn_s:
        lbl="Send & log" if st.session_state.get("_gmail_approved") else "📋 Log Send"
        clicked=st.button(lbl,key=f"send_{tab_key}",type="primary",use_container_width=True,disabled=not recip_val)
    st.markdown('</div>',unsafe_allow_html=True)
    return clicked

# ═══════════════════════════════════════════════════════════════════════════════
# ROW 1 — Customer selector
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<p class="ce-label">Select Customer</p>',unsafe_allow_html=True)

# Build customer list — all unique customers sorted, preserving original casing
if cust_col:
    _all_customers=sorted(
        {str(v).strip() for v in _df_drs[cust_col].dropna()
         if str(v).strip() not in ("","nan","None")},
        key=str.lower
    )
else:
    _all_customers=[]

if not _all_customers:
    st.warning("No customer data found in DRS. Check the 'Account Name' / 'Customer' column.")
    st.stop()

_prev_cust=st.session_state.get("_ce_customer")
_def_cust=_prev_cust if _prev_cust in _all_customers else _all_customers[0]
selected_customer=st.selectbox(
    "Customer",options=_all_customers,
    index=_all_customers.index(_def_cust),
    label_visibility="collapsed",key="ce_customer",
)
# Clear project selection when customer changes
if st.session_state.get("_ce_customer")!=selected_customer:
    st.session_state["_ce_customer"]=selected_customer
    for k in list(st.session_state.keys()):
        if any(k.startswith(p) for p in ["_ce_proj","ce_to","ce_cn","ce_cc","_sfdc_match","ce_ss_stamp"]):
            st.session_state.pop(k,None)

# ── All projects for this customer (across all consultants) ───────────────────
if cust_col:
    df_cust_all=_df_drs[_df_drs[cust_col].astype(str).str.strip()==selected_customer].copy()
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

n_mine=len(_mine_proj); n_other=len(_other_proj); n_total=n_mine+n_other

# ── Project header row ────────────────────────────────────────────────────────
_ph_left,_ph_right=st.columns([2,1])
with _ph_left:
    st.markdown(
        f'<p class="ce-label" style="margin-bottom:4px">'
        f'{n_total} project{"s" if n_total!=1 else ""}'
        f' · {n_mine} yours'
        + (f', {n_other} assigned elsewhere' if n_other else '')
        + '</p>',
        unsafe_allow_html=True
    )
with _ph_right:
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

# Build cards HTML + track clickable sids for mine projects
_mine_sids=[_proj_sid(_row_dict(r)) for _,r in _mine_proj.iterrows()]
_prev_sid=st.session_state.get("_ce_proj_sid")
_def_sid=_prev_sid if _prev_sid in _mine_sids else (_mine_sids[0] if _mine_sids else None)

if not _mine_sids:
    _va_display = _flip_name(_view_as_name) if _view_as_name != _logged_in else "you"
    st.info(f"No active projects assigned to {_va_display} for {selected_customer}.")
    st.stop()

# Render mine projects as radio-style cards using Streamlit radio
def _card_label(row,mine=True):
    """Project name only — clean, no redundant concatenation."""
    name=str(row.get(name_col,"")) if name_col else ""
    return name

# Use selectbox for project selection (clean Streamlit widget, no hacks)
_mine_labels={_mine_sids[i]:_card_label(_row_dict(_mine_proj.iloc[i]),mine=True)
              for i in range(len(_mine_sids))}

selected_sid=st.selectbox(
    "Your projects",
    options=_mine_sids,
    format_func=lambda s:_mine_labels.get(s,s),
    index=_mine_sids.index(_def_sid) if _def_sid in _mine_sids else 0,
    key="ce_proj_select",
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
if n_other>0:
    with st.expander(f"Also at {selected_customer} — {n_other} project{'s' if n_other!=1 else ''} assigned to other consultants",expanded=False):
        for _,r in _other_proj.iterrows():
            rd=_row_dict(r)
            _pm=str(rd.get(pm_col,"")) if pm_col else "—"
            _pm_disp=_flip_name(_pm)
            _ini=_initials(_pm)
            _prod=str(rd.get(prod_col,"")) if prod_col else ""
            _nm=str(rd.get(name_col,"")) if name_col else ""
            _stat=str(rd.get(status_col,"")) if status_col else ""
            _iv=rd.get(intro_col,"") if intro_col else ""
            _intro_done=_iv and str(_iv).strip() not in ("","None","nan","NaT")
            _welcome=f"Welcome sent {_iv}" if _intro_done else "Welcome pending"
            st.markdown(
                f'<div class="proj-card other" style="opacity:.6;margin-bottom:5px">'
                f'<div style="display:flex;justify-content:space-between;align-items:flex-start">'
                f'<div class="proj-name">{_nm}</div>'
                f'<span class="pill-gray">{_prod}</span>'
                f'</div>'
                f'<div class="proj-meta">{_stat}</div>'
                f'<div class="proj-consultant">'
                f'<div class="avatar other">{_ini}</div>'
                f'{_pm_disp} · {_welcome}'
                f'</div></div>',
                unsafe_allow_html=True
            )

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
    sfdc_match,sfdc_label=_fuzzy_sfdc(df_sn,pnm,str(customer))
    if not sfdc_match.empty:
        ec="email" if "email" in sfdc_match.columns else None
        nc="contact_name" if "contact_name" in sfdc_match.columns else None
        fc="impl_contact_flag" if "impl_contact_flag" in sfdc_match.columns else None
        pc="is_primary" if "is_primary" in sfdc_match.columns else None
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
        if ec and fc:
            impl_rows=sfdc_match[sfdc_match[fc].astype(str).isin(["1","True","true","yes","x"])]
            for _,r in impl_rows.iterrows():
                e=str(r.get(ec,"")).strip()
                if e and e not in ("nan","None","") and e!=sfdc_email:
                    sfdc_cc_emails.append(e)
        _sfdc_key=f"_sfdc_match_{project_id}"
        if st.session_state.get(_sfdc_key)!=sfdc_email:
            st.session_state[_sfdc_key]=sfdc_email
            st.session_state["ce_to"]=sfdc_email
            st.session_state["ce_cn"]=sfdc_cname

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
    _COMM_TYPES=["Welcome","Post-Session","Lifecycle (UAT → Closure)"]
    _STAGE_TO_COMM={
        "welcome":"Welcome","post_session_1":"Post-Session","post_session_2":"Post-Session",
        "uat_signoff":"Lifecycle (UAT → Closure)","go_live":"Lifecycle (UAT → Closure)",
        "hypercare_closure":"Lifecycle (UAT → Closure)",
    }
    # Smart default: derive from first pending journey stage
    # Only auto-set on project change (not on every render — respect user's manual choice)
    _proj_changed = st.session_state.get("_last_proj_sid_for_comm") != selected_sid
    if _proj_changed:
        _statuses_d = ["done" if _ms_date(s,sel) else "pending" for s in _JOURNEY]
        _first_pend = next((i for i,st2 in enumerate(_statuses_d) if st2=="pending"), 0)
        _smart_default = _STAGE_TO_COMM.get(_JOURNEY[_first_pend]["id"],"Welcome")
        st.session_state["_ce_tmpl_type"] = _smart_default
        st.session_state["_last_proj_sid_for_comm"] = selected_sid
    comm_col,tmpl_col=st.columns([1,2])
    with comm_col:
        _tmpl_type=st.selectbox(
            "Communication type",options=_COMM_TYPES,
            index=_COMM_TYPES.index(st.session_state.get("_ce_tmpl_type","Welcome")),
            key="ce_tmpl_type",
        )
        st.session_state["_ce_tmpl_type"]=_tmpl_type
        st.session_state["_ce_tab"]={"Welcome":"Welcome","Post-Session":"Post-Session","Lifecycle (UAT → Closure)":"Lifecycle"}.get(_tmpl_type,"Welcome")

    # Live context — read after recipient widgets exist
    _live_cname=st.session_state.get("ce_cn","") or ""
    auto_ctx=build_auto_context(sel,_disp,{"contact_name":_live_cname} if _live_cname else None)
    if _live_cname: auto_ctx["CUSTOMER_CONTACT_NAME"]=_live_cname
    auto_ctx["SENDER"]=_disp; auto_ctx["CONSULTANT_NAME"]=_disp

    # ── Recipient (below comm type so consultant knows what they're composing) ─
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
        with tmpl_col:
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
                opts=list_welcome_templates()
                ch=st.selectbox("Template",[t["display_name"] for t in opts],key="w_manual",label_visibility="collapsed")
                tmpl_w=get_welcome_template(next(t["sku_key"] for t in opts if t["display_name"]==ch))
            else:
                disp=tmpl_w.get("display_name","")
                st.selectbox("Template",[disp],key="w_tmpl_disp",disabled=True,label_visibility="collapsed")
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
            st.session_state["_req_w"]={"subj":subj_w,"body":body_w,"ssf":ssf_w}
            st.rerun()
        if st.session_state.get("_req_w"):
            r=st.session_state.pop("_req_w")
            try:
                with st.spinner("Logging…"):
                    ok,sid=execute_send(project_id=project_id,template_id=f"welcome_{tmpl_w.get('sku_key','manual')}",
                        template_name=f"Welcome — {tmpl_w.get('display_name','')}",
                        subject=r["subj"],body=r["body"],recipient_email=recip,cc_emails=cc_emails,ss_milestone_field=r["ssf"])
                if ok:
                    st.success(f"✓ Logged — ID: `{sid}`")
                    if r["ssf"] and st.session_state.get("ce_ss_stamp",True):
                        if _do_write(project_id,r["ssf"],datetime.date.today(),sel): mark_ss_writeback_done(sid)
                else: st.error(f"Failed: {sid}")
            except Exception as ex: st.error(f"Error: {ex}"); st.exception(ex)

    # ── Post-Session ──────────────────────────────────────────────────────────
    elif _tmpl_type=="Post-Session":
        psk=_ps_key(str(product_raw))
        if not psk:
            with tmpl_col: st.info(f"No post-session templates for '{product_raw}'.")
        else:
            sessions=get_post_session_templates(psk)
            sopts={s["id"]:(f"Session {s['session_number']} — {s['name']}"+(f" [{s['variant_note']}]" if s.get("variant_note") else ""),s) for s in sessions}
            with tmpl_col:
                cid=st.selectbox("Template",list(sopts.keys()),format_func=lambda k:sopts[k][0],key="s_pick",label_visibility="collapsed")
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
                    if ft=="text":       v=st.text_input(lbl_d,placeholder=ph,key=f"s_{cid}_{k}")
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

    # ── Lifecycle ─────────────────────────────────────────────────────────────
    elif _tmpl_type=="Lifecycle (UAT → Closure)":
        lc_all=list_lifecycle_templates()
        lc_opts={t["id"]:t for t in lc_all}
        with tmpl_col:
            lcid=st.selectbox("Template",list(lc_opts.keys()),
                              format_func=lambda k:f"[{lc_opts[k]['category']}] {lc_opts[k]['name']}",
                              key="lc_pick",label_visibility="collapsed")
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
        subj_l=_inject_customer_subject(subj_l, customer)
        auto_vals_l={v for v in {**auto_ctx_l,**mctx_l}.values() if v and str(v).strip() and len(str(v))>2 and "{" not in str(v)}
        ssf_l=tmpl_l.get("ss_milestone_on_send"); gls=mctx_l.get("GO_LIVE_DATE","")
        st.session_state["ce_prev_subj"]=subj_l
        st.session_state["ce_prev_body"]=body_l
        st.session_state["ce_prev_auto"]=auto_vals_l
        if _send_footer("l",ssf_l if isinstance(ssf_l,str) else (ssf_l[0] if ssf_l else None),subj_l,body_l,recip):
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
    _ps=st.session_state.get("ce_prev_subj","")
    _pb=st.session_state.get("ce_prev_body","")
    _pa=st.session_state.get("ce_prev_auto",set())
    _recip_display=st.session_state.get("ce_to","")
    _cc_display=st.session_state.get("ce_cc","")

    if _pb:
        st.markdown(_email_html(_ps,_pb,_recip_display,_cc_display,_pa),unsafe_allow_html=True)
        _cc_parts=[]
        if _consultant_email(_logged_in): _cc_parts.append("you")
        if sfdc_cc_emails: _cc_parts.append(f"{len(sfdc_cc_emails)} other SFDC contact(s)")
        if _sales_rep_email: _cc_parts.append("Sales Rep")
        if len(_cc_parts)>1: st.caption(f"CC includes: {', '.join(_cc_parts)}")
        with st.expander("Edit before sending"):
            st.text_area("Subject",value=_ps,key="prev_subj_edit",height=40)
            st.text_area("Body",value=_pb,key="prev_body_edit",height=300)
    else:
        st.markdown(
            '<div class="ce-card" style="text-align:center;padding:32px 20px">'
            '<div style="font-size:24px;margin-bottom:10px;opacity:.25">✉</div>'
            '<div style="font-size:13px;opacity:.45">Select a communication type to see preview</div>'
            '</div>',unsafe_allow_html=True)

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
