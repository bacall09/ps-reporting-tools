"""
pages/15_Customer_Engagement.py  v2
"""
import streamlit as st
import pandas as pd
import datetime
import re

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
.ce-preview{border:1px solid rgba(128,128,128,.22);border-radius:8px;padding:20px 24px;font-size:13px;line-height:1.7;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-height:460px;overflow-y:auto;white-space:pre-wrap}
.ce-log-row{border-bottom:1px solid rgba(128,128,128,.15);padding:8px 0;font-size:13px}
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
    from shared.smartsheet_api import update_project_fields
    _ss_ok = True
except ImportError:
    _ss_ok = False

# ── Helpers ───────────────────────────────────────────────────────────────────

def _row_dict(row):
    return {k: (None if pd.isna(v) else v) for k, v in row.items()}

def _missing(text):
    return list(set(re.findall(r"\{[A-Z_]+\}", text)))

def _flip_name(n):
    """'Swanson, Patti' → 'Patti Swanson'"""
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

def _sfdc_contact(customer):
    df = st.session_state.get("df_sfdc")
    if df is None or df.empty: return None
    cm = {c.lower().strip(): c for c in df.columns}
    cc = (cm.get("customer") or cm.get("account name") or cm.get("account")
          or cm.get("company") or cm.get("customer name"))
    if not cc: return None
    mask = df[cc].astype(str).str.lower().str.contains(re.escape(customer.lower()), na=False)
    m = df[mask]
    if m.empty:
        m = df[df[cc].astype(str).str.lower() == customer.lower()]
    if m.empty: return None
    r = m.iloc[0]
    nc = cm.get("contact name") or cm.get("name") or cm.get("full name") or cm.get("first name")
    ec = cm.get("email") or cm.get("email address") or cm.get("contact email")
    return {
        "contact_name": str(r[nc]).strip() if nc else "",
        "email":        str(r[ec]).strip() if ec else "",
    }

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

def _do_write(project_id, ss_field, date_val, drs_row):
    if not _ss_ok or not ss_field: return False
    fields = [ss_field] if isinstance(ss_field, str) else ss_field
    wrote = False
    for f in fields:
        try:
            p = build_ss_writeback(project_id, f, date_val, current_drs_row=drs_row)
            if p["skipped"]:
                st.info(f"ℹ️ **{f}** already set — not overwritten.")
                continue
            if p["fields"]:
                update_project_fields(project_id, p["fields"])
                st.success(f"✓ Smartsheet: **{f}** → {date_val.strftime('%d %b %Y')}")
                wrote = True
        except Exception as ex:
            st.warning(f"Writeback failed for **{f}**: {ex}")
    return wrote

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
              or cm.get("intro_email_sent"))

if not _is_mgr and pm_col:
    df_all = df_all[df_all[pm_col].apply(lambda x: name_matches(str(x), _logged_in))]

if status_col:
    df_all = df_all[~df_all[status_col].astype(str).str.lower().isin(
        ["closed","cancelled","complete","completed"])]

df_all = df_all.reset_index(drop=True)

if df_all.empty:
    st.info("No active projects found.")
    st.stop()

# FIX 1: cache welcome flags — avoids per-row YAML load on every render
@st.cache_data(ttl=120, show_spinner=False)
def _welcome_flags(sig: str) -> list:
    import json
    rows = json.loads(sig)
    return [i for i, r in enumerate(rows) if project_needs_welcome_email(r)]

try:
    import json
    _sig = df_all[[c for c in df_all.columns if c in [
        intro_col, status_col, prod_col
    ] if c]].to_json(orient="records")
    _needs_welcome_idx = set(_welcome_flags(_sig))
except Exception:
    _needs_welcome_idx = set()

# ── Layout ────────────────────────────────────────────────────────────────────
col_left, col_right = st.columns([1, 2], gap="large")

with col_left:
    st.markdown('<p class="ce-label">Select Project</p>', unsafe_allow_html=True)

    # FIX 4: welcome filter toggle
    _active_tab = st.session_state.get("_ce_tab", "Welcome")
    _show_filter = False
    if intro_col and _active_tab == "Welcome":
        _show_filter = st.checkbox("Only projects needing Welcome email", value=True, key="ce_wf")

    df = df_all.copy()
    if _show_filter and intro_col:
        _mask = (df[intro_col].isna() |
                 df[intro_col].astype(str).str.strip().isin(["","None","nan","NaT"]))
        if _mask.any():
            df = df[_mask].reset_index(drop=True)
        else:
            st.success("All active projects have a Welcome email on record.")

    if df.empty:
        st.info("No projects match this filter.")
        st.stop()

    def _plabel(row):
        n = str(row.get(name_col,"")) if name_col else ""
        p = str(row.get(prod_col,"")) if prod_col else ""
        c = str(row.get(cust_col,"")) if cust_col else ""
        parts = [x for x in [c, n, f"· {p}" if p else ""] if x.strip()]
        return "  —  ".join(parts[:2]) + (f"  {parts[2]}" if len(parts)>2 else "")

    sel_idx = st.selectbox(
        "Project", options=list(df.index),
        format_func=lambda i: _plabel(df.loc[i]),
        label_visibility="collapsed", key="ce_proj",
    )

    sel = _row_dict(df.loc[sel_idx])
    customer    = sel.get(cust_col,"") if cust_col else ""
    product_raw = sel.get(prod_col,"") if prod_col else ""
    project_id  = sel.get(id_col, str(sel_idx)) if id_col else str(sel_idx)

    # FIX 2: multi-project / consolidated
    siblings = (
        [i for i in df.index if df.loc[i,cust_col]==customer and i!=sel_idx]
        if cust_col and customer else []
    )
    consolidated = False
    all_rows = [sel]
    if siblings:
        st.markdown(
            f'<span class="ce-info">{len(siblings)+1} projects for {customer}</span>',
            unsafe_allow_html=True,
        )
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
            prods = [r.get(prod_col,"") for r in all_rows if r.get(prod_col)]
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

    # Recipient
    sfdc = _sfdc_contact(customer)
    st.markdown('<p class="ce-label" style="margin-top:12px">Recipient</p>', unsafe_allow_html=True)
    if sfdc and sfdc.get("email"):
        recip  = st.text_input("To", value=sfdc["email"], key="ce_to")
        cname  = st.text_input("Contact Name", value=sfdc.get("contact_name",""), key="ce_cn")
    else:
        df_s = st.session_state.get("df_sfdc")
        msg = (f"No SFDC match for '{customer}' — enter manually."
               if df_s is not None and not df_s.empty
               else "SFDC contacts not loaded — enter manually.")
        st.caption(msg)
        recip = st.text_input("To (recipient email)", placeholder="customer@example.com", key="ce_to")
        cname = st.text_input("Contact Name", placeholder="First name", key="ce_cn")

    cc_default = _consultant_email(_logged_in)
    cc_in = st.text_input("CC", value=cc_default,
                           help="Comma-separated. Consultant CC'd by default.", key="ce_cc")
    cc_emails = [e.strip() for e in cc_in.split(",") if e.strip()]

    # Session log
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

# ── RIGHT ─────────────────────────────────────────────────────────────────────
with col_right:
    st.markdown('<p class="ce-label">Template</p>', unsafe_allow_html=True)

    tab_w, tab_s, tab_l = st.tabs(["Welcome","Post-Session","Lifecycle (UAT → Closure)"])

    # FIX 9+10: correct name and contact BEFORE render
    _disp = _flip_name(_logged_in)
    _sfdc_ctx = {"contact_name": cname} if cname else None
    auto_ctx = build_auto_context(sel, _disp, _sfdc_ctx)
    if cname:
        auto_ctx["CUSTOMER_CONTACT_NAME"] = cname
    auto_ctx["SENDER"]          = _disp
    auto_ctx["CONSULTANT_NAME"] = _disp

    # ── Preview + copy helper ─────────────────────────────────────────────────
    def _preview(subject, body, key):
        m = _missing(body + subject)
        if m:
            st.markdown(
                f'<div class="ce-warn">⚠️ Unfilled placeholders: {", ".join(m)}</div>',
                unsafe_allow_html=True)
        subj_edit = st.text_input("Subject", value=subject, key=f"{key}_subj")

        # FIX 11: HTML preview preserving structure
        def _htmlify(text):
            out = []
            for line in text.split("\n"):
                s = line.strip()
                if not s:
                    out.append("<br>")
                elif s == s.upper() and len(s) > 3 and not s.startswith("•"):
                    out.append(f"<strong>{s}</strong>")
                elif s.startswith("•"):
                    out.append(f"&nbsp;&nbsp;{s}")
                else:
                    out.append(s)
            return "<br>".join(out)

        st.markdown("**Preview**")
        st.markdown(f'<div class="ce-preview">{_htmlify(body)}</div>',
                    unsafe_allow_html=True)
        with st.expander("Edit plain text before sending"):
            body_edit = st.text_area("", value=body, height=300,
                                     key=f"{key}_body", label_visibility="collapsed")
        body_edit = st.session_state.get(f"{key}_body", body)
        return subj_edit, body_edit

    # ── Send flow helper ──────────────────────────────────────────────────────
    def _send(key, subj, body, ss_field, tmpl_id, tmpl_name, go_live_str=""):
        if ss_field:
            fd = ", ".join(ss_field) if isinstance(ss_field, list) else ss_field
            st.markdown(
                f'<div class="ce-tip">On send, <strong>{fd}</strong> will be date-stamped '
                f'in Smartsheet (existing values not overwritten).</div>',
                unsafe_allow_html=True)
        lbl = "Send Email" if st.session_state.get("_gmail_approved") else "📋 Log Send"
        if st.button(lbl, key=f"btn_{key}", type="primary", disabled=not recip):
            st.session_state[f"_req_{key}"] = {
                "subj": subj, "body": body, "ss_field": ss_field,
                "tid": tmpl_id, "tname": tmpl_name, "gld": go_live_str,
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
                            wd = gld if f in (SS_GO_LIVE_DATE, SS_PROD_CUTOVER) else datetime.date.today()
                            if _do_write(project_id, f, wd, sel): mark_ss_writeback_done(sid)
                else:
                    st.error(f"Failed: {sid}")
            except Exception as ex:
                st.error(f"Error: {ex}"); st.exception(ex)

    # ── TAB: Welcome ──────────────────────────────────────────────────────────
    with tab_w:
        st.session_state["_ce_tab"] = "Welcome"

        if consolidated and len(all_rows) > 1:
            st.markdown(
                f'<div class="ce-tip">Consolidated mode: one email for all '
                f'{len(all_rows)} products for {customer}.</div>', unsafe_allow_html=True)
            prods = [r.get(prod_col,"") for r in all_rows if r.get(prod_col)]
            tmpl_w = None
            for r in all_rows:
                k = _sku(str(r.get(prod_col,"")))
                if k and "_" in k:
                    tmpl_w = get_welcome_template(k); break
            if not tmpl_w and prods:
                tmpl_w = get_welcome_template(_sku(prods[0]))
        else:
            tmpl_w = get_welcome_template(_sku(str(product_raw)))

        if not tmpl_w:
            st.caption(f"No automatic match for '{product_raw}'. Select manually:")
            opts = list_welcome_templates()
            chosen = st.selectbox("Template", [t["display_name"] for t in opts], key="w_manual")
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
                sopts[s["id"]] = (lbl, s)

            cid = st.selectbox("Session", list(sopts.keys()),
                               format_func=lambda k: sopts[k][0], key="s_pick")
            _, tmpl_s = sopts[cid]
            st.caption(f"**Audience:** {tmpl_s.get('audience','Full project team')}")

            mctx: dict = {}
            if tmpl_s.get("editable_fields"):
                with st.expander("Fill in session details", expanded=True):
                    for f in tmpl_s["editable_fields"]:
                        k, lbl, ft = f["key"], f["label"], f.get("type","text")
                        req = f.get("required", False)
                        ph  = f.get("placeholder","")
                        if ft == "text":
                            v = st.text_input(lbl+(" *" if req else ""), placeholder=ph,
                                              key=f"s_{cid}_{k}")
                        elif ft == "textarea":
                            v = st.text_area(lbl+(" *" if req else ""), placeholder=ph,
                                             height=80, key=f"s_{cid}_{k}")
                        elif ft == "multiselect":
                            sel2 = st.multiselect(lbl+(" *" if req else ""),
                                                  options=f.get("options",[]),
                                                  key=f"s_{cid}_{k}")
                            v = "\n".join(f"  • {o}" for o in sel2)
                        elif ft == "select":
                            v = st.selectbox(lbl, f.get("options",[]), key=f"s_{cid}_{k}")
                        else:
                            v = st.text_input(lbl, key=f"s_{cid}_{k}")
                        mctx[k] = v
                        if k == "GO_LIVE_READINESS" and v:
                            rm = tmpl_s.get("go_live_readiness_text",{})
                            res = rm.get(v[0], v)
                            if "{HYPERCARE_DATE}" in res:
                                res = res.replace("{HYPERCARE_DATE}", mctx.get("HYPERCARE_DATE","{HYPERCARE_DATE}"))
                            mctx["GO_LIVE_READINESS_TEXT"] = res

            subj_s, body_s = render_template(tmpl_s["body"], tmpl_s["subject"], {}, {**auto_ctx,**mctx})
            ssf_s = tmpl_s.get("ss_milestone_on_send")
            sf2, bf2 = _preview(subj_s, body_s, "session")
            _send("session", sf2, bf2, ssf_s, tmpl_s["id"], tmpl_s["name"])

    # ── TAB: Lifecycle ────────────────────────────────────────────────────────
    with tab_l:
        st.session_state["_ce_tab"] = "Lifecycle"

        lc_all = list_lifecycle_templates()
        lc_opts = {t["id"]: t for t in lc_all}
        lcid = st.selectbox("Lifecycle template", list(lc_opts.keys()),
                            format_func=lambda k: f"[{lc_opts[k]['category']}]  {lc_opts[k]['name']}",
                            key="lc_pick")
        tmpl_l = get_lifecycle_template(lcid)
        st.caption(f"**When to send:** {tmpl_l['trigger']}")
        for tip in tmpl_l.get("tips",[]):
            st.markdown(f'<div class="ce-tip">💡 {tip}</div>', unsafe_allow_html=True)

        vbody = tmpl_l.get("body","")
        if tmpl_l.get("variants"):
            vlbls = {v["key"]: f"{v['label']} — {v['description']}" for v in tmpl_l["variants"]}
            cv = st.radio("Scenario", list(vlbls.keys()),
                          format_func=lambda k: vlbls[k],
                          key=f"lv_{lcid}", label_visibility="collapsed")
            vbody = tmpl_l["variant_bodies"][cv]

        mctx_l: dict = {}
        if tmpl_l.get("editable_fields"):
            with st.expander("Fill in details", expanded=True):
                for f in tmpl_l["editable_fields"]:
                    k, lbl = f["key"], f["label"]
                    req, src = f.get("required",False), f.get("source","")
                    default = str(f.get("default",""))
                    if src == "drs_prod_cutover":
                        raw = sel.get("prod_cutover") or sel.get("Prod Cutover")
                        if raw:
                            try: default = pd.to_datetime(raw).date().isoformat()
                            except: pass
                    elif src == "drs_project_link":
                        default = str(sel.get("project_link") or sel.get("Project Link") or "")
                    elif src == "calculated_go_live_plus_14":
                        glr = sel.get("prod_cutover") or sel.get("go_live_date") or mctx_l.get("GO_LIVE_DATE","")
                        if glr:
                            try:
                                default = (pd.to_datetime(glr).date() + datetime.timedelta(days=14)).isoformat()
                            except: pass
                    v = st.text_input(lbl+(" *" if req else ""), value=default,
                                      placeholder="YYYY-MM-DD" if f.get("type")=="date" else "",
                                      key=f"l_{lcid}_{k}")
                    mctx_l[k] = v

        subj_l, body_l = render_template(vbody, tmpl_l["subject"], {}, {**auto_ctx,**mctx_l})
        ssf_l = tmpl_l.get("ss_milestone_on_send")
        gls = mctx_l.get("GO_LIVE_DATE","")
        sf3, bf3 = _preview(subj_l, body_l, "lifecycle")
        _send("lifecycle", sf3, bf3, ssf_l, lcid, tmpl_l["name"], go_live_str=gls)
