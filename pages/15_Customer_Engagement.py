"""
pages/15_Customer_Engagement.py  v3
Fixes: load time, stale data on project switch, SS writeback, SFDC matching, HTML tables in preview
"""
import streamlit as st
import pandas as pd
import datetime
import re
from rapidfuzz import fuzz

st.session_state["current_page"] = "Customer Engagement"

st.markdown("""
<div style='background:#050D1F;padding:20px 28px 16px;border-radius:0 0 8px 8px;margin-bottom:24px'>
  <span style='color:#4472C4;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase'>Customer Engagement</span>
  <h2 style='color:#ffffff;margin:4px 0 0;font-size:22px;font-weight:600;letter-spacing:-0.3px'>Lifecycle Email Composer</h2>
</div>
""", unsafe_allow_html=True)

st.markdown("""<style>
.ce-label{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.9px;color:#4472C4;margin:0 0 6px}
.ce-card{border:1px solid rgba(128,128,128,.22);border-radius:8px;padding:16px 20px;margin-bottom:14px}
.ce-flag{display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:700;background:rgba(239,68,68,.14);color:#dc2626}
.ce-sent{display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:700;background:rgba(34,197,94,.14);color:#16a34a}
.ce-info{display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:700;background:rgba(68,114,196,.14);color:#4472C4}
.ce-tip{background:rgba(68,114,196,.08);border-left:3px solid #4472C4;border-radius:0 6px 6px 0;padding:10px 14px;font-size:13px;margin-bottom:12px}
.ce-warn{background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.25);border-radius:6px;padding:8px 12px;font-size:12px;color:#dc2626;margin-bottom:10px}
.ce-preview{border:1px solid rgba(128,128,128,.22);border-radius:8px;padding:20px 24px;font-size:13px;line-height:1.7;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-height:500px;overflow-y:auto}
.ce-log-row{border-bottom:1px solid rgba(128,128,128,.15);padding:8px 0;font-size:13px}
.ce-tbl{border-collapse:collapse;width:100%;margin:8px 0 12px;font-size:12px}
.ce-tbl th{background:#1a2c5b;color:#fff;padding:6px 10px;text-align:left;font-weight:600}
.ce-tbl td{padding:5px 10px;border-bottom:1px solid rgba(128,128,128,.2)}
.ce-section-hdr{font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin:14px 0 4px;font-size:12px;color:#4472C4}
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
    from shared.smartsheet_api import update_project_fields, get_row_id_for_project
    _ss_ok = True
    _ss_has_row_id = True
except ImportError:
    try:
        from shared.smartsheet_api import update_project_fields
        _ss_ok = True
        _ss_has_row_id = False
    except ImportError:
        _ss_ok = False
        _ss_has_row_id = False

# ── SFDC column map (ported from reengagement page) ───────────────────────────
_SFDC_COL_MAP = {
    "18 digit opportunity id": "opportunity_id",
    "first name":              "first_name",
    "last name":               "last_name",
    "primary title":           "title",
    "title":                   "title",
    "email":                   "email",
    "contact email":           "email",
    "email address":           "email",
    "opportunity owner":       "account_manager",
    "account owner":           "account_manager",
    "owner":                   "account_manager",
    "account name":            "account",
    "account":                 "account",
    "opportunity name":        "opportunity",
    "opportunity":             "opportunity",
    "contact name":            "contact_name",
    "full name":               "contact_name",
    "name":                    "contact_name",
    "close date":              "close_date",
    "closed date":             "close_date",
    "implementation contact":  "impl_contact_flag",
    "contact roles":           "contact_roles",
    "opportunity owner email": "account_manager_email",
    "owner email":             "account_manager_email",
    "primary":                 "is_primary",
    "is primary":              "is_primary",
}

_PRODUCT_KEYWORDS = [
    "Capture","Approvals","Reconcile","PSP","Payments","SFTP",
    "E-Invoicing","eInvoicing","CC","Premium","ZoneCapture",
    "ZoneApprovals","ZoneReconcile","ZonePayments",
]

def _clean_account(text):
    t = str(text).lower()
    for stop in ["ltd","limited","inc","llc","plc","gmbh","the ","- za -","& co","co."]:
        t = t.replace(stop," ")
    return re.sub(r"[^a-z0-9 ]"," ",t).split()

def _product_hints(text):
    t = str(text).lower()
    return {kw for kw in _PRODUCT_KEYWORDS if kw.lower() in t}

def _fuzzy_sfdc(df_sfdc, project_name, account_name):
    """Port of reengagement fuzzy_match_sfdc."""
    if df_sfdc is None or df_sfdc.empty:
        return pd.DataFrame(), None
    # Normalise columns
    df = df_sfdc.copy()
    col_map = {c.lower().strip(): c for c in df.columns}
    opp_col  = col_map.get("opportunity")
    acc_col  = col_map.get("account")
    oid_col  = col_map.get("opportunity_id")
    # Exact opp match
    if opp_col:
        exact = df[df[opp_col].astype(str).str.lower().str.strip() == str(project_name).lower().strip()]
        if not exact.empty:
            return exact, "Exact match"
    # Fuzzy
    drs_words = set(_clean_account(account_name))
    drs_prods = _product_hints(project_name)
    best_score = 0; best_opp_id = None; best_opp_nm = None
    for _, row in df.iterrows():
        sfdc_acct = str(row.get(acc_col or "account",""))
        sfdc_opp  = str(row.get(opp_col or "opportunity",""))
        sfdc_words = set(_clean_account(sfdc_acct))
        common = drs_words & sfdc_words
        word_score = len(common)/max(len(drs_words),1)*100
        fuzz_score = fuzz.token_set_ratio(" ".join(drs_words)," ".join(sfdc_words))
        acct_score = max(word_score, fuzz_score*0.7)
        prod_match = bool(drs_prods & _product_hints(sfdc_opp))
        score = acct_score + (30 if prod_match else 0)
        if score > best_score:
            best_score = score
            best_opp_id = row.get(oid_col or "opportunity_id") if oid_col else None
            best_opp_nm = row.get(opp_col or "opportunity") if opp_col else None
    label = f"Fuzzy match ({int(best_score)}%)"
    if best_score >= 60:
        if best_opp_id is not None and oid_col:
            rows = df[df[oid_col] == best_opp_id]
            if not rows.empty: return rows, label
        if best_opp_nm is not None and opp_col:
            rows = df[df[opp_col] == best_opp_nm]
            if not rows.empty: return rows, label
    # Last resort: account name only
    if acc_col:
        df["_sc"] = df[acc_col].apply(lambda x: fuzz.token_set_ratio(str(account_name).lower(),str(x).lower()))
        top = df[df["_sc"]>=75].sort_values("_sc",ascending=False)
        df.drop(columns=["_sc"],inplace=True)
        if not top.empty: return top, f"Account match"
    return pd.DataFrame(), None

# ── Helpers ───────────────────────────────────────────────────────────────────

def _row_dict(row):
    return {k:(None if pd.isna(v) else v) for k,v in row.items()}

def _missing(text):
    return list(set(re.findall(r"\{[A-Z_]+\}", text)))

def _flip_name(n):
    if "," in n:
        p = [x.strip() for x in n.split(",",1)]
        return f"{p[1]} {p[0]}"
    return n

def _consultant_email(name):
    try:
        from shared.constants import EMPLOYEE_ROLES
        em = EMPLOYEE_ROLES.get(name,{}).get("email","")
        if em: return em
    except Exception:
        pass
    slug = re.sub(r"[^a-z0-9.]","",_flip_name(name).lower().replace(" ","."))
    return f"{slug}@zoneandco.com"

_PMAP = {
    "zoneapp: capture":                          "ZoneCapture",
    "zoneapp: approvals":                        "ZoneApprovals",
    "zoneapp: reconcile":                        "ZoneReconcile",
    "zoneapp: reconcile 2.0":                    "ZoneReconcile_BankConnect",
    "zoneapp: reconcile with bank connectivity": "ZoneReconcile_BankConnect",
    "zoneapp: reconcile with cc import":         "ZoneReconcile_CCImport",
    "zoneapp: e-invoicing":                      "EInvoicing",
    "zoneapp: capture & e-invoicing":            "ZoneCapture_EInvoicing",
    "zoneapp: capture and e-invoicing":          "ZoneCapture_EInvoicing",
    "zoneapp: capture & approvals":              "ZoneCapture_ZoneApprovals",
    "zoneapp: capture and approvals":            "ZoneCapture_ZoneApprovals",
    "zoneapp: capture & reconcile":              "ZoneCapture_ZoneReconcile",
    "zoneapp: capture and reconcile":            "ZoneCapture_ZoneReconcile",
    "zonecapture":                               "ZoneCapture",
    "zoneapprovals":                             "ZoneApprovals",
    "zonereconcile":                             "ZoneReconcile",
    "zonereconcile with bank connectivity":      "ZoneReconcile_BankConnect",
    "zonereconcile 2.0":                         "ZoneReconcile_BankConnect",
    "zonereconcile with cc import":              "ZoneReconcile_CCImport",
    "e-invoicing":                               "EInvoicing",
    "zone e-invoicing":                          "EInvoicing",
    "zonecapture with e-invoicing":              "ZoneCapture_EInvoicing",
    "zonecapture and zoneapprovals":             "ZoneCapture_ZoneApprovals",
    "zonecapture and zonereconcile":             "ZoneCapture_ZoneReconcile",
}

def _sku(p): return _PMAP.get(str(p).strip().lower())
def _ps_key(p):
    p = str(p).lower()
    if "capture" in p:   return "capture"
    if "approv" in p:    return "approvals"
    if "reconcile" in p: return "reconcile"
    return None

# FIX 3: writeback — pass SS row ID if available, else project_id
def _do_write(project_id, ss_field, date_val, drs_row):
    if not _ss_ok or not ss_field: return False
    fields = [ss_field] if isinstance(ss_field, str) else ss_field
    wrote = False
    for f in fields:
        try:
            p = build_ss_writeback(project_id, f, date_val, current_drs_row=drs_row)
            if p["skipped"]:
                st.info(f"ℹ️ **{f}** already set in Smartsheet — not overwritten.")
                continue
            if p["fields"]:
                # Try to get the actual SS row ID from the DRS row if available
                ss_row_id = None
                if drs_row:
                    ss_row_id = (drs_row.get("ss_row_id") or drs_row.get("row_id")
                                 or drs_row.get("smartsheet_row_id"))
                write_id = ss_row_id if ss_row_id else project_id
                update_project_fields(write_id, p["fields"])
                st.success(f"✓ Smartsheet: **{f}** → {date_val.strftime('%d %b %Y')}")
                wrote = True
        except Exception as ex:
            st.warning(f"Writeback failed for **{f}**: {ex}. Update manually in Smartsheet.")
    return wrote

# FIX 5: HTML body renderer — preserves tables and structure
def _body_to_html(body: str) -> str:
    """
    Convert plain-text email body to structured HTML.
    Handles: section headers (ALL CAPS), bullet lists, line breaks.
    Tables stored in YAML as markdown-ish are rendered as HTML tables.
    """
    lines = body.split("\n")
    html_parts = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Blank line → spacing
        if not stripped:
            html_parts.append("<div style='margin:6px 0'></div>")
            i += 1
            continue

        # Section header: ALL CAPS, no punctuation at end, length > 3
        if (stripped == stripped.upper() and len(stripped) > 3
                and not stripped.startswith("•") and not stripped.startswith("|")):
            html_parts.append(
                f"<div class='ce-section-hdr'>{stripped}</div>"
            )
            i += 1
            continue

        # Markdown-style table: lines starting with |
        if stripped.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                if not re.match(r"^\s*\|[-| :]+\|\s*$", lines[i]):
                    table_lines.append(lines[i].strip())
                i += 1
            if table_lines:
                html_parts.append("<table class='ce-tbl'>")
                for ti, tline in enumerate(table_lines):
                    cells = [c.strip() for c in tline.strip("|").split("|")]
                    if ti == 0:
                        html_parts.append("<tr>" + "".join(f"<th>{c}</th>" for c in cells) + "</tr>")
                    else:
                        html_parts.append("<tr>" + "".join(f"<td>{c}</td>" for c in cells) + "</tr>")
                html_parts.append("</table>")
            continue

        # Bullet
        if stripped.startswith("•") or stripped.startswith("-") and len(stripped) > 2:
            bullet_text = stripped.lstrip("•- ").strip()
            html_parts.append(f"<div style='margin:3px 0 3px 16px'>• {bullet_text}</div>")
            i += 1
            continue

        # Sub-header: ends with colon, not too long
        if stripped.endswith(":") and len(stripped) < 60 and stripped[0].isupper():
            html_parts.append(f"<div style='font-weight:600;margin-top:10px'>{stripped}</div>")
            i += 1
            continue

        # Normal line
        html_parts.append(f"<div>{stripped}</div>")
        i += 1

    return "\n".join(html_parts)

# ── Build DRS dataframe ───────────────────────────────────────────────────────
df_all = _df_drs.copy()
cm = {c.lower().strip(): c for c in df_all.columns}

name_col   = cm.get("project_name")    or cm.get("project name")
cust_col   = cm.get("customer")
prod_col   = cm.get("project_type")    or cm.get("project type") or cm.get("product")
id_col     = cm.get("project_id")      or cm.get("project id")
status_col = cm.get("status")
pm_col     = cm.get("project_manager") or cm.get("project manager")
intro_col  = (cm.get("intro. email sent") or cm.get("intro email sent")
              or cm.get("ms_intro_email") or cm.get("intro_email_sent"))

# FIX 3: store SS row ID column if present
ss_row_id_col = cm.get("ss_row_id") or cm.get("row_id") or cm.get("id")

# Role filter
if not _is_mgr and pm_col:
    df_all = df_all[df_all[pm_col].apply(lambda x: name_matches(str(x), _logged_in))]

# Active only
if status_col:
    df_all = df_all[~df_all[status_col].astype(str).str.lower().isin(
        ["closed","cancelled","complete","completed"])]

df_all = df_all.reset_index(drop=True)
if df_all.empty:
    st.info("No active projects found for your account.")
    st.stop()

# Precompute welcome flag as a simple boolean column — no caching needed, fast
def _needs_intro(row):
    if not intro_col: return False
    v = row.get(intro_col,"")
    return v is None or str(v).strip() in ("","None","nan","NaT")

df_all["_needs_welcome"] = df_all.apply(lambda r: _needs_intro(r), axis=1)

# ── Layout ────────────────────────────────────────────────────────────────────
col_left, col_right = st.columns([1, 2], gap="large")

# ═══════════════════════════════════════════════════════════════════════════════
# LEFT PANEL
# ═══════════════════════════════════════════════════════════════════════════════
with col_left:
    st.markdown('<p class="ce-label">Select Project</p>', unsafe_allow_html=True)

    _active_tab = st.session_state.get("_ce_tab","Welcome")

    # FIX 4: template-aware filter — Welcome tab only
    df = df_all.copy()
    if intro_col and _active_tab == "Welcome":
        if st.checkbox("Only projects needing Welcome email", value=True, key="ce_wf"):
            _filtered = df[df["_needs_welcome"]]
            if not _filtered.empty:
                df = _filtered.reset_index(drop=True)
            else:
                st.success("All active projects have a Welcome email on record.")

    if df.empty:
        st.info("No projects match this filter.")
        st.stop()

    def _plabel(row):
        c = str(row.get(cust_col,"")) if cust_col else ""
        n = str(row.get(name_col,""))  if name_col else ""
        p = str(row.get(prod_col,""))  if prod_col else ""
        return f"{c}  —  {n}" + (f"  · {p}" if p and p not in ("nan","None") else "")

    sel_idx = st.selectbox(
        "Project", options=list(df.index),
        format_func=lambda i: _plabel(df.loc[i]),
        label_visibility="collapsed", key="ce_proj",
    )

    # FIX 1+2: detect project change — clear stale widget state
    _proj_key = f"{sel_idx}_{df.loc[sel_idx, id_col] if id_col else sel_idx}"
    if st.session_state.get("_last_proj_key") != _proj_key:
        st.session_state["_last_proj_key"] = _proj_key
        # Clear all per-project widget states so new project renders fresh
        for k in list(st.session_state.keys()):
            if any(k.startswith(pfx) for pfx in ["ce_to","ce_cn","ce_cc","welcome_","session_","lifecycle_","w_","s_","l_","sess_","lc_"]):
                del st.session_state[k]

    sel         = _row_dict(df.loc[sel_idx])
    customer    = sel.get(cust_col,"") if cust_col else ""
    product_raw = sel.get(prod_col,"") if prod_col else ""
    project_id  = str(sel.get(id_col, str(sel_idx))) if id_col else str(sel_idx)

    # Multi-project / consolidated
    siblings = (
        [i for i in df.index if str(df.loc[i,cust_col])==str(customer) and i!=sel_idx]
        if cust_col and customer and str(customer) not in ("","nan","None") else []
    )
    consolidated = False
    all_rows = [sel]
    if siblings:
        st.markdown(
            f'<span class="ce-info">{len(siblings)+1} projects for this customer</span>',
            unsafe_allow_html=True)
        consolidated = st.checkbox("Consolidated email (all products)", key="ce_con")
        if consolidated:
            all_rows = [sel] + [_row_dict(df.loc[i]) for i in siblings]

    # Project card
    st.markdown('<div class="ce-card">', unsafe_allow_html=True)
    st.markdown('<p class="ce-label">Project Summary</p>', unsafe_allow_html=True)
    if name_col: st.markdown(f"**{sel.get(name_col,'')}**")
    if cust_col: st.caption(f"Customer: {customer}")
    if prod_col:
        if consolidated:
            prods = [r.get(prod_col,"") for r in all_rows if r.get(prod_col) and str(r.get(prod_col)) not in ("nan","None")]
            st.caption(f"Products: {', '.join(prods)}")
        else:
            st.caption(f"Product: {product_raw}")
    if status_col: st.caption(f"Status: {sel.get(status_col,'')}")
    iv = sel.get(intro_col,"") if intro_col else ""
    if not iv or str(iv).strip() in ("","None","nan","NaT"):
        st.markdown('<span class="ce-flag">Welcome email pending</span>', unsafe_allow_html=True)
    else:
        st.caption(f"Welcome sent: {iv}")
    st.markdown("</div>", unsafe_allow_html=True)

    # ── SFDC contact lookup (FIX 4: use fuzzy match from reengagement) ────────
    df_sfdc = st.session_state.get("df_sfdc")

    sfdc_match    = pd.DataFrame()
    match_label   = None
    sfdc_email    = ""
    sfdc_cname    = ""

    if df_sfdc is not None and not df_sfdc.empty:
        # Normalise SFDC columns using the full col map
        _sfdc_cm = {c.lower().strip(): c for c in df_sfdc.columns}
        _rename  = {c: _SFDC_COL_MAP[c.lower().strip()]
                    for c in df_sfdc.columns if c.lower().strip() in _SFDC_COL_MAP}
        df_sfdc_n = df_sfdc.rename(columns=_rename)
        # Build contact_name from first+last if present
        if "first_name" in df_sfdc_n.columns and "last_name" in df_sfdc_n.columns:
            df_sfdc_n["contact_name"] = (
                df_sfdc_n["first_name"].fillna("") + " " + df_sfdc_n["last_name"].fillna("")
            ).str.strip()
        proj_nm = str(sel.get(name_col,"")) if name_col else ""
        sfdc_match, match_label = _fuzzy_sfdc(df_sfdc_n, proj_nm, str(customer))
        if not sfdc_match.empty:
            # Prefer impl contact flag, then primary, then first row
            _ec = "email"       if "email"        in sfdc_match.columns else None
            _nc = "contact_name"if "contact_name" in sfdc_match.columns else None
            _fc = "impl_contact_flag" if "impl_contact_flag" in sfdc_match.columns else None
            _pc = "is_primary"  if "is_primary"   in sfdc_match.columns else None
            best_row = sfdc_match.iloc[0]
            if _fc:
                flagged = sfdc_match[sfdc_match[_fc].astype(str).str.lower().isin(["true","1","yes","x","checked"])]
                if not flagged.empty: best_row = flagged.iloc[0]
            if _ec: sfdc_email = str(best_row.get(_ec,"")).strip()
            if _nc: sfdc_cname = str(best_row.get(_nc,"")).strip()
            if sfdc_email in ("nan","None",""): sfdc_email = ""
            if sfdc_cname in ("nan","None",""): sfdc_cname = ""

    # Recipient fields
    st.markdown('<p class="ce-label" style="margin-top:12px">Recipient</p>', unsafe_allow_html=True)
    if sfdc_email:
        st.caption(f"✅ SFDC matched — {match_label}")
        recip = st.text_input("To", value=sfdc_email, key="ce_to")
        cname = st.text_input("Contact Name", value=sfdc_cname, key="ce_cn")
    else:
        if df_sfdc is not None and not df_sfdc.empty:
            st.caption(f"No SFDC match for '{customer}' — enter manually.")
        else:
            st.caption("SFDC contacts not loaded — enter manually.")
        recip = st.text_input("To (recipient email)", placeholder="customer@example.com", key="ce_to")
        cname = st.text_input("Contact Name", placeholder="First name", key="ce_cn")

    cc_default = _consultant_email(_logged_in)
    cc_in = st.text_input("CC", value=cc_default,
                          help="Comma-separated. Consultant CC'd by default.", key="ce_cc")
    cc_emails = [e.strip() for e in cc_in.split(",") if e.strip()]

    # Session send log
    log = get_session_send_log()
    proj_log = [e for e in log if e["project_id"]==project_id]
    if proj_log:
        st.markdown('<p class="ce-label" style="margin-top:18px">Sent This Session</p>',
                    unsafe_allow_html=True)
        for e in proj_log:
            dt = e["sent_at"][:16].replace("T"," ")
            st.markdown(
                f'<div class="ce-log-row"><span class="ce-sent">✓ Sent</span>&nbsp;'
                f'<strong>{e["template_name"]}</strong><br>'
                f'<span style="font-size:11px;opacity:.65">{dt} UTC → {e["recipient_email"]}</span>'
                f'</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# RIGHT PANEL
# ═══════════════════════════════════════════════════════════════════════════════
with col_right:
    st.markdown('<p class="ce-label">Template</p>', unsafe_allow_html=True)

    tab_w, tab_s, tab_l = st.tabs(["Welcome","Post-Session","Lifecycle (UAT → Closure)"])

    # Build auto_ctx with corrected name BEFORE render
    _disp = _flip_name(_logged_in)
    _sfdc_ctx = {"contact_name": cname} if cname else None
    auto_ctx = build_auto_context(sel, _disp, _sfdc_ctx)
    # Ensure manually-entered values always win
    if cname:
        auto_ctx["CUSTOMER_CONTACT_NAME"] = cname
    auto_ctx["SENDER"]          = _disp
    auto_ctx["CONSULTANT_NAME"] = _disp

    # ── Preview helper ────────────────────────────────────────────────────────
    def _preview(subject, body, key):
        m = _missing(body + subject)
        if m:
            st.markdown(
                f'<div class="ce-warn">⚠️ Unfilled placeholders: {", ".join(m)}</div>',
                unsafe_allow_html=True)
        subj_edit = st.text_input("Subject", value=subject, key=f"{key}_subj")
        st.markdown("**Preview**")
        html_body = _body_to_html(body)
        st.markdown(f'<div class="ce-preview">{html_body}</div>', unsafe_allow_html=True)
        with st.expander("Edit plain text before sending"):
            body_edit = st.text_area("", value=body, height=300,
                                     key=f"{key}_body", label_visibility="collapsed")
        body_edit = st.session_state.get(f"{key}_body", body)
        return subj_edit, body_edit

    # ── Send flow ─────────────────────────────────────────────────────────────
    def _send(key, subj, body, ss_field, tmpl_id, tmpl_name, go_live_str=""):
        if ss_field:
            fd = ", ".join(ss_field) if isinstance(ss_field,list) else ss_field
            st.markdown(
                f'<div class="ce-tip">On send, <strong>{fd}</strong> will be date-stamped '
                f'in Smartsheet (existing values not overwritten).</div>',
                unsafe_allow_html=True)
        lbl = "Send Email" if st.session_state.get("_gmail_approved") else "📋 Log Send"
        if st.button(lbl, key=f"btn_{key}", type="primary", disabled=not recip):
            st.session_state[f"_req_{key}"] = {
                "subj":subj,"body":body,"ss_field":ss_field,
                "tid":tmpl_id,"tname":tmpl_name,"gld":go_live_str,
            }
            st.rerun()
        if st.session_state.get(f"_req_{key}"):
            r = st.session_state.pop(f"_req_{key}")
            try:
                with st.spinner("Logging…"):
                    ok, sid = execute_send(
                        project_id=project_id, template_id=r["tid"],
                        template_name=r["tname"], subject=r["subj"], body=r["body"],
                        recipient_email=recip, cc_emails=cc_emails,
                        ss_milestone_field=r["ss_field"],
                    )
                if ok:
                    st.success(f"✓ Logged — ID: `{sid}`")
                    if r["ss_field"]:
                        try: gld = datetime.date.fromisoformat(r["gld"][:10]) if r["gld"] else datetime.date.today()
                        except: gld = datetime.date.today()
                        for f in (r["ss_field"] if isinstance(r["ss_field"],list) else [r["ss_field"]]):
                            wd = gld if f in (SS_GO_LIVE_DATE,SS_PROD_CUTOVER) else datetime.date.today()
                            if _do_write(project_id, f, wd, sel):
                                mark_ss_writeback_done(sid)
                else:
                    st.error(f"Failed: {sid}")
            except Exception as ex:
                st.error(f"Error: {ex}"); st.exception(ex)

    # ── TAB: Welcome ──────────────────────────────────────────────────────────
    with tab_w:
        st.session_state["_ce_tab"] = "Welcome"
        if consolidated and len(all_rows)>1:
            st.markdown(
                f'<div class="ce-tip">Consolidated mode: one email covering all '
                f'{len(all_rows)} products for {customer}.</div>', unsafe_allow_html=True)
            tmpl_w = None
            for r in all_rows:
                k = _sku(str(r.get(prod_col,"")))
                if k and "_" in k:
                    tmpl_w = get_welcome_template(k); break
            if not tmpl_w:
                prods = [r.get(prod_col,"") for r in all_rows if r.get(prod_col)]
                if prods: tmpl_w = get_welcome_template(_sku(prods[0]))
        else:
            tmpl_w = get_welcome_template(_sku(str(product_raw)))

        if not tmpl_w:
            st.caption(f"No automatic match for '{product_raw}'. Select manually:")
            opts = list_welcome_templates()
            chosen = st.selectbox("Template",[t["display_name"] for t in opts],key="w_manual")
            tmpl_w = get_welcome_template(next(t["sku_key"] for t in opts if t["display_name"]==chosen))

        var = st.radio("Sender variant",
                       ["Variant A — sent by PM or automated","Variant B — sent by Consultant"],
                       horizontal=True, key="w_var")
        vk = "variant_a" if "A" in var else "variant_b"

        subj_r, body_r = render_template(tmpl_w[vk]["body"], tmpl_w["subject"], auto_ctx)
        lib_meta = _welcome_library()
        ssf_w = lib_meta.get("ss_milestone_on_send")
        sf, bf = _preview(subj_r, body_r, "welcome")
        _send("welcome", sf, bf, ssf_w,
              f"welcome_{tmpl_w.get('sku_key','?')}", f"Welcome — {tmpl_w['display_name']}")

    # ── TAB: Post-Session ─────────────────────────────────────────────────────
    with tab_s:
        st.session_state["_ce_tab"] = "Post-Session"
        psk = _ps_key(str(product_raw))
        if not psk:
            st.caption(f"No post-session templates for '{product_raw}'.")
        else:
            sessions = get_post_session_templates(psk)
            sopts = {}
            for s in sessions:
                lbl = f"Session {s['session_number']} — {s['name']}"
                if s.get("variant_note"): lbl += f"  [{s['variant_note']}]"
                sopts[s["id"]] = (lbl,s)
            cid = st.selectbox("Session",list(sopts.keys()),
                               format_func=lambda k:sopts[k][0],key="s_pick")
            _,tmpl_s = sopts[cid]
            st.caption(f"**Audience:** {tmpl_s.get('audience','Full project team')}")
            mctx: dict = {}
            if tmpl_s.get("editable_fields"):
                with st.expander("Fill in session details",expanded=True):
                    for f in tmpl_s["editable_fields"]:
                        k,lbl,ft=f["key"],f["label"],f.get("type","text")
                        req=f.get("required",False); ph=f.get("placeholder","")
                        if ft=="text":       v=st.text_input(lbl+(" *"if req else ""),placeholder=ph,key=f"s_{cid}_{k}")
                        elif ft=="textarea": v=st.text_area(lbl+(" *"if req else ""),placeholder=ph,height=80,key=f"s_{cid}_{k}")
                        elif ft=="multiselect":
                            sel2=st.multiselect(lbl+(" *"if req else ""),options=f.get("options",[]),key=f"s_{cid}_{k}")
                            v="\n".join(f"  • {o}" for o in sel2)
                        elif ft=="select":  v=st.selectbox(lbl,f.get("options",[]),key=f"s_{cid}_{k}")
                        else:               v=st.text_input(lbl,key=f"s_{cid}_{k}")
                        mctx[k]=v
                        if k=="GO_LIVE_READINESS" and v:
                            rm=tmpl_s.get("go_live_readiness_text",{})
                            res=rm.get(v[0],v)
                            if "{HYPERCARE_DATE}" in res:
                                res=res.replace("{HYPERCARE_DATE}",mctx.get("HYPERCARE_DATE","{HYPERCARE_DATE}"))
                            mctx["GO_LIVE_READINESS_TEXT"]=res
            subj_s,body_s=render_template(tmpl_s["body"],tmpl_s["subject"],{},{**auto_ctx,**mctx})
            ssf_s=tmpl_s.get("ss_milestone_on_send")
            sf2,bf2=_preview(subj_s,body_s,"session")
            _send("session",sf2,bf2,ssf_s,tmpl_s["id"],tmpl_s["name"])

    # ── TAB: Lifecycle ────────────────────────────────────────────────────────
    with tab_l:
        st.session_state["_ce_tab"] = "Lifecycle"
        lc_all=list_lifecycle_templates()
        lc_opts={t["id"]:t for t in lc_all}
        lcid=st.selectbox("Lifecycle template",list(lc_opts.keys()),
                          format_func=lambda k:f"[{lc_opts[k]['category']}]  {lc_opts[k]['name']}",
                          key="lc_pick")
        tmpl_l=get_lifecycle_template(lcid)
        st.caption(f"**When to send:** {tmpl_l['trigger']}")
        for tip in tmpl_l.get("tips",[]):
            st.markdown(f'<div class="ce-tip">💡 {tip}</div>',unsafe_allow_html=True)
        vbody=tmpl_l.get("body","")
        if tmpl_l.get("variants"):
            vlbls={v["key"]:f"{v['label']} — {v['description']}" for v in tmpl_l["variants"]}
            cv=st.radio("Scenario",list(vlbls.keys()),format_func=lambda k:vlbls[k],
                        key=f"lv_{lcid}",label_visibility="collapsed")
            vbody=tmpl_l["variant_bodies"][cv]
        mctx_l: dict={}
        if tmpl_l.get("editable_fields"):
            with st.expander("Fill in details",expanded=True):
                for f in tmpl_l["editable_fields"]:
                    k,lbl=f["key"],f["label"]
                    req,src=f.get("required",False),f.get("source","")
                    default=str(f.get("default",""))
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
                    v=st.text_input(lbl+(" *"if req else ""),value=default,
                                    placeholder="YYYY-MM-DD"if f.get("type")=="date" else "",
                                    key=f"l_{lcid}_{k}")
                    mctx_l[k]=v
        subj_l,body_l=render_template(vbody,tmpl_l["subject"],{},{**auto_ctx,**mctx_l})
        ssf_l=tmpl_l.get("ss_milestone_on_send")
        gls=mctx_l.get("GO_LIVE_DATE","")
        sf3,bf3=_preview(subj_l,body_l,"lifecycle")
        _send("lifecycle",sf3,bf3,ssf_l,lcid,tmpl_l["name"],go_live_str=gls)
