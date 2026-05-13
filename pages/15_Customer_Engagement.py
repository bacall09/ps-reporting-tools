"""
pages/15_Customer_Engagement.py
─────────────────────────────────────────────────────────────────────────────
PSPT — Customer Engagement

Purpose:
  Lifecycle email composer for PS consultants. Surfaces the right template
  for a project's current stage, auto-fills all available data from DRS /
  SFDC, and lets the consultant review and send from the shared PS inbox.
  On send, writes the milestone date back to Smartsheet and logs the send
  to the session log.

Roles: All consultants (own projects only). Managers see all projects.

Data sources:
  - df_drs   : project metadata, milestone dates, project link (SS)
  - df_ns    : consultant assignment (for project ownership filter)
  - SFDC contacts : recipient email / contact name (via Google Drive)
  - template_engine : template library, rendering, send tracking
"""

import streamlit as st
import pandas as pd
import datetime
import re

st.session_state["current_page"] = "Customer Engagement"

# ── Hero banner ───────────────────────────────────────────────────────────────
st.markdown("""
<div style='background:#050D1F;padding:20px 28px 16px;border-radius:0 0 8px 8px;margin-bottom:24px'>
  <span style='color:#4472C4;font-size:11px;font-weight:700;letter-spacing:1.5px;
               text-transform:uppercase'>Customer Engagement</span>
  <h2 style='color:#ffffff;margin:4px 0 0;font-size:22px;font-weight:600;letter-spacing:-0.3px'>
    Lifecycle Email Composer
  </h2>
</div>
""", unsafe_allow_html=True)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.ce-section-label {
    font-size:12px;font-weight:700;text-transform:uppercase;
    letter-spacing:0.9px;color:#4472C4;margin:0 0 6px;
}
.ce-card {
    border:1px solid rgba(128,128,128,0.22);border-radius:8px;
    padding:16px 20px;margin-bottom:14px;
}
.ce-pill-flag {
    display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;
    font-weight:700;background:rgba(239,68,68,0.14);color:#dc2626;
}
.ce-pill-flag-dark { color:#fca5a5; }
.ce-pill-sent {
    display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;
    font-weight:700;background:rgba(34,197,94,0.14);color:#16a34a;
}
.ce-pill-info {
    display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;
    font-weight:700;background:rgba(68,114,196,0.14);color:#4472C4;
}
.ce-tip {
    background:rgba(68,114,196,0.08);border-left:3px solid #4472C4;
    border-radius:0 6px 6px 0;padding:10px 14px;font-size:13px;
    margin-bottom:12px;color:inherit;
}
.ce-placeholder-missing {
    background:rgba(239,68,68,0.08);border:1px solid rgba(239,68,68,0.25);
    border-radius:6px;padding:8px 12px;font-size:12px;color:#dc2626;
    margin-bottom:10px;
}
.ce-send-log-row {
    border-bottom:1px solid rgba(128,128,128,0.15);padding:8px 0;font-size:13px;
}
</style>
""", unsafe_allow_html=True)

# ── Auth ──────────────────────────────────────────────────────────────────────
_logged_in = st.session_state.get("consultant_name", "")
if not _logged_in:
    st.warning("Please log in via the Home page.")
    st.stop()

from shared.constants import get_role, is_manager
_role = get_role(_logged_in)
_is_mgr = is_manager(_logged_in) or _role in ("manager_only", "reporting_only")

# ── Data check ────────────────────────────────────────────────────────────────
_df_drs = st.session_state.get("df_drs")
if _df_drs is None or _df_drs.empty:
    st.info("Load Smartsheet DRS data on the Home page to use this tool.")
    st.stop()

# ── Template engine ───────────────────────────────────────────────────────────
try:
    from shared.template_engine import (
        get_welcome_template,
        list_welcome_templates,
        get_post_session_templates,
        get_lifecycle_template,
        list_lifecycle_templates,
        build_auto_context,
        render_template,
        execute_send,
        build_ss_writeback,
        mark_ss_writeback_done,
        get_session_send_log,
        project_needs_welcome_email,
        _welcome_library,
    )
except ImportError as e:
    st.error(f"Template engine not found: {e}. Ensure shared/template_engine.py is deployed.")
    st.stop()

# ── Smartsheet writeback ──────────────────────────────────────────────────────
try:
    from shared.smartsheet_api import update_project_fields
    _ss_writeback_available = True
except ImportError:
    _ss_writeback_available = False


def _do_ss_writeback(
    project_id: str,
    ss_field,           # str or list[str]
    date_val: datetime.date,
    drs_row: dict,
) -> bool:
    """
    Attempt SS writeback for one or more milestone fields.
    Respects no-override logic — skipped fields surface a visible warning.
    Returns True if at least one field was written successfully.
    """
    if not _ss_writeback_available or not ss_field:
        return False

    fields = [ss_field] if isinstance(ss_field, str) else ss_field
    any_written = False

    for field in fields:
        try:
            from shared.template_engine import build_ss_writeback
            payload = build_ss_writeback(project_id, field, date_val, current_drs_row=drs_row)

            if payload["skipped"]:
                st.info(
                    f"ℹ️ **{field}** already has a value in Smartsheet — not overwritten. "
                    f"{payload['skip_reason']}"
                )
                continue

            if payload["fields"]:
                update_project_fields(project_id, payload["fields"])
                st.success(f"✓ Smartsheet: **{field}** updated to {date_val.strftime('%d %b %Y')}.")
                any_written = True

        except Exception as e:
            st.warning(f"Email sent but Smartsheet writeback failed for **{field}**: {e}. Please update manually.")

    return any_written


# ── Helper: slugify customer name for ps-support alias ───────────────────────
def _ps_support_email(customer: str) -> str:
    slug = re.sub(r"[^A-Z0-9]", "", customer.upper())
    return f"ps-support+{slug}@zoneandco.com"


# ── Helper: get SFDC contact for a project ───────────────────────────────────
def _get_sfdc_contact(customer: str) -> dict | None:
    """Return first matching SFDC contact for the customer, or None."""
    df_sfdc = st.session_state.get("df_sfdc")
    if df_sfdc is None or df_sfdc.empty:
        return None
    col_map = {c.lower(): c for c in df_sfdc.columns}
    cust_col = col_map.get("customer") or col_map.get("account name")
    if not cust_col:
        return None
    match = df_sfdc[df_sfdc[cust_col].str.lower() == customer.lower()]
    if match.empty:
        return None
    row = match.iloc[0]
    contact_col = col_map.get("contact name") or col_map.get("name")
    email_col   = col_map.get("email") or col_map.get("email address")
    return {
        "contact_name": row[contact_col] if contact_col else "",
        "email":        row[email_col]   if email_col   else "",
    }


# ── Helper: map DRS row to a dict template_engine understands ────────────────
def _drs_row_to_dict(row: pd.Series) -> dict:
    return {k: (None if pd.isna(v) else v) for k, v in row.items()}


# ── Helper: detect unreplaced placeholders in rendered body ──────────────────
def _find_missing_placeholders(text: str) -> list[str]:
    return list(set(re.findall(r"\{[A-Z_]+\}", text)))


# ── Helper: map product string from DRS to a sku_key ─────────────────────────
_PRODUCT_KEY_MAP = {
    # Exact DRS values → sku_key in welcome_templates.yaml
    "zonecapture":                              "ZoneCapture",
    "zone capture":                             "ZoneCapture",
    "zoneapprovals":                            "ZoneApprovals",
    "zone approvals":                           "ZoneApprovals",
    "zonereconcile":                            "ZoneReconcile",
    "zone reconcile":                           "ZoneReconcile",
    "zonereconcile with bank connectivity":     "ZoneReconcile_BankConnect",
    "zone reconcile with bank connectivity":    "ZoneReconcile_BankConnect",
    "zonereconcile with cc import":             "ZoneReconcile_CCImport",
    "zone reconcile with cc import":            "ZoneReconcile_CCImport",
    "e-invoicing":                              "EInvoicing",
    "einvoicing":                               "EInvoicing",
    "zone e-invoicing":                         "EInvoicing",
    "zonecapture with e-invoicing":             "ZoneCapture_EInvoicing",
    "zonecapture and zoneinvoicing":            "ZoneCapture_EInvoicing",
    "zonecapture and zoneapprovals":            "ZoneCapture_ZoneApprovals",
    "zonecapture and zonereconcile":            "ZoneCapture_ZoneReconcile",
}

def _resolve_sku_key(product_raw: str) -> str | None:
    return _PRODUCT_KEY_MAP.get(str(product_raw).strip().lower())


def _product_to_post_session_key(product_raw: str) -> str | None:
    p = str(product_raw).lower()
    if "capture" in p:   return "capture"
    if "approv" in p:    return "approvals"
    if "reconcile" in p: return "reconcile"
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Layout: Project picker + Composer
# ─────────────────────────────────────────────────────────────────────────────

col_left, col_right = st.columns([1, 2], gap="large")

# ── LEFT: Project picker ──────────────────────────────────────────────────────
with col_left:
    st.markdown('<p class="ce-section-label">Select Project</p>', unsafe_allow_html=True)

    # Filter DRS to consultant's projects (managers see all)
    df = _df_drs.copy()
    col_map_drs = {c.lower(): c for c in df.columns}

    if not _is_mgr:
        pm_col = col_map_drs.get("project_manager") or col_map_drs.get("project manager")
        if pm_col:
            from shared.constants import name_matches
            df = df[df[pm_col].apply(lambda x: name_matches(str(x), _logged_in))]

    # Active projects only
    status_col = col_map_drs.get("status")
    if status_col:
        df = df[~df[status_col].str.lower().isin(["closed", "cancelled", "complete", "completed"])]

    if df.empty:
        st.info("No active projects found.")
        st.stop()

    # Build display labels
    name_col  = col_map_drs.get("project_name") or col_map_drs.get("project name")
    cust_col  = col_map_drs.get("customer")
    prod_col  = col_map_drs.get("product") or col_map_drs.get("project_type") or col_map_drs.get("project type")
    id_col    = col_map_drs.get("project_id") or col_map_drs.get("project id")

    def _label(row):
        parts = []
        if name_col and row.get(name_col): parts.append(str(row[name_col]))
        if cust_col and row.get(cust_col): parts.append(f"({row[cust_col]})")
        return " ".join(parts) if parts else f"Project {row.name}"

    project_labels = {i: _label(row) for i, row in df.iterrows()}

    # Flag projects needing welcome email
    needs_welcome = []
    for i, row in df.iterrows():
        if project_needs_welcome_email(_drs_row_to_dict(row)):
            needs_welcome.append(i)

    if needs_welcome:
        st.markdown(
            f'<span class="ce-pill-flag">⚡ {len(needs_welcome)} project{"s" if len(needs_welcome)>1 else ""} '
            f'need{"s" if len(needs_welcome)==1 else ""} a Welcome email</span>',
            unsafe_allow_html=True,
        )
        st.markdown("")

    selected_idx = st.selectbox(
        "Project",
        options=list(project_labels.keys()),
        format_func=lambda i: project_labels[i],
        label_visibility="collapsed",
    )

    selected_row = _drs_row_to_dict(df.loc[selected_idx])
    customer     = selected_row.get(cust_col, "") if cust_col else ""
    product_raw  = selected_row.get(prod_col, "") if prod_col else ""
    project_id   = selected_row.get(id_col, str(selected_idx)) if id_col else str(selected_idx)

    # Project summary card
    with st.container():
        st.markdown('<div class="ce-card">', unsafe_allow_html=True)
        st.markdown(f'<p class="ce-section-label">Project Summary</p>', unsafe_allow_html=True)
        if name_col:  st.markdown(f"**{selected_row.get(name_col, '')}**")
        if cust_col:  st.caption(f"Customer: {customer}")
        if prod_col:  st.caption(f"Product: {product_raw}")
        if status_col: st.caption(f"Status: {selected_row.get(status_col, '')}")
        if project_needs_welcome_email(selected_row):
            st.markdown('<span class="ce-pill-flag">Welcome email pending</span>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # Recipient info (SFDC)
    sfdc_contact = _get_sfdc_contact(customer)
    st.markdown('<p class="ce-section-label" style="margin-top:12px">Recipient</p>', unsafe_allow_html=True)
    if sfdc_contact and sfdc_contact.get("email"):
        recipient_email = st.text_input("To", value=sfdc_contact["email"])
        contact_name    = st.text_input("Contact Name", value=sfdc_contact.get("contact_name", ""))
    else:
        st.caption("SFDC contacts not loaded — enter manually.")
        recipient_email = st.text_input("To (recipient email)", placeholder="customer@example.com")
        contact_name    = st.text_input("Contact Name", placeholder="First name")

    cc_input = st.text_input(
        "CC",
        value=f"{_logged_in.lower().replace(' ', '.')}@zoneandco.com" if _logged_in else "",
        help="Comma-separated. Consultant is CC'd by default.",
    )
    cc_emails = [e.strip() for e in cc_input.split(",") if e.strip()]

    # Session send log
    send_log = get_session_send_log()
    project_sends = [e for e in send_log if e["project_id"] == project_id]
    if project_sends:
        st.markdown('<p class="ce-section-label" style="margin-top:18px">Sent This Session</p>', unsafe_allow_html=True)
        for entry in project_sends:
            sent_dt = entry["sent_at"][:16].replace("T", " ")
            st.markdown(
                f'<div class="ce-send-log-row">'
                f'<span class="ce-pill-sent">✓ Sent</span>&nbsp;'
                f'<strong>{entry["template_name"]}</strong><br>'
                f'<span style="font-size:11px;opacity:0.65">{sent_dt} UTC → {entry["recipient_email"]}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )


# ── RIGHT: Template composer ──────────────────────────────────────────────────
with col_right:
    st.markdown('<p class="ce-section-label">Template</p>', unsafe_allow_html=True)

    # ── Tab structure: Welcome | Post-Session | Lifecycle | Re-engagement ─────
    tab_welcome, tab_session, tab_lifecycle = st.tabs([
        "Welcome", "Post-Session", "Lifecycle (UAT → Closure)"
    ])

    # Build auto-context once (shared across tabs)
    sfdc_ctx = {"contact_name": contact_name} if contact_name else None
    auto_ctx = build_auto_context(selected_row, _logged_in, sfdc_ctx)
    # Inject manually-entered contact name
    if contact_name:
        auto_ctx["CUSTOMER_CONTACT_NAME"] = contact_name

    # ── TAB: Welcome ──────────────────────────────────────────────────────────
    with tab_welcome:
        sku_key = _resolve_sku_key(str(product_raw))
        tmpl_w  = get_welcome_template(sku_key) if sku_key else None

        if not tmpl_w:
            # Manual picker if auto-match fails
            st.caption(f"No automatic match for product '{product_raw}'. Select manually:")
            all_welcome = list_welcome_templates()
            chosen_label = st.selectbox(
                "Template",
                options=[t["display_name"] for t in all_welcome],
                key="welcome_manual_pick",
            )
            sku_key = next(t["sku_key"] for t in all_welcome if t["display_name"] == chosen_label)
            tmpl_w  = get_welcome_template(sku_key)

        # Variant picker
        variant_choice = st.radio(
            "Sender variant",
            options=["Variant A — sent by PM or automated", "Variant B — sent by Consultant"],
            horizontal=True,
            key="welcome_variant",
        )
        variant_key = "variant_a" if "A" in variant_choice else "variant_b"
        variant      = tmpl_w[variant_key]

        # Render preview
        subject_rendered, body_rendered = render_template(
            variant["body"], tmpl_w["subject"], auto_ctx
        )
        missing = _find_missing_placeholders(body_rendered + subject_rendered)

        if missing:
            st.markdown(
                f'<div class="ce-placeholder-missing">⚠️ Unfilled placeholders: '
                f'{", ".join(missing)}</div>',
                unsafe_allow_html=True,
            )

        subject_edit = st.text_input("Subject", value=subject_rendered, key="welcome_subj")
        body_edit    = st.text_area("Email body", value=body_rendered, height=420, key="welcome_body")

        # Milestone writeback info
        lib_meta = _welcome_library()
        ss_field = lib_meta.get("ss_milestone_on_send")
        if ss_field:
            st.markdown(
                f'<div class="ce-tip">On send, <strong>Intro. Email Sent</strong> will be '
                f'date-stamped in Smartsheet for this project.</div>',
                unsafe_allow_html=True,
            )

        _send_disabled = not recipient_email
        if st.button(
            "Send Welcome Email" if st.session_state.get("_gmail_approved") else "📋 Prepare & Log Send",
            key="btn_send_welcome",
            type="primary",
            disabled=_send_disabled,
        ):
            if not recipient_email:
                st.error("Recipient email is required.")
            else:
                st.session_state["_welcome_send_requested"] = {
                    "subject": subject_edit,
                    "body": body_edit,
                    "ss_field": ss_field,
                    "template_id": f"welcome_{sku_key}",
                    "template_name": f"Welcome — {tmpl_w['display_name']}",
                }
                st.rerun()

        if st.session_state.get("_welcome_send_requested"):
            req = st.session_state.pop("_welcome_send_requested")
            try:
                with st.spinner("Sending…"):
                    success, send_id_or_err = execute_send(
                        project_id=project_id,
                        template_id=req["template_id"],
                        template_name=req["template_name"],
                        subject=req["subject"],
                        body=req["body"],
                        recipient_email=recipient_email,
                        cc_emails=cc_emails,
                        ss_milestone_field=req["ss_field"],
                    )
                if success:
                    st.success(f"✓ Logged. Send ID: `{send_id_or_err}`")
                    if req["ss_field"]:
                        did_write = _do_ss_writeback(project_id, req["ss_field"], datetime.date.today(), selected_row)
                        if did_write:
                            mark_ss_writeback_done(send_id_or_err)
                else:
                    st.error(f"Send failed: {send_id_or_err}")
            except Exception as e:
                st.error(f"Error during send: {e}")
                st.exception(e)

    # ── TAB: Post-Session ─────────────────────────────────────────────────────
    with tab_session:
        ps_key = _product_to_post_session_key(str(product_raw))

        if not ps_key:
            st.caption(f"No post-session templates for product '{product_raw}'.")
        else:
            sessions = get_post_session_templates(ps_key)

            # If reconcile, session 1 has two variants (Legacy vs BankConnect)
            session_options = {}
            for s in sessions:
                label = f"Session {s['session_number']} — {s['name']}"
                if s.get("variant_note"):
                    label += f"  [{s['variant_note']}]"
                session_options[s["id"]] = (label, s)

            chosen_session_id = st.selectbox(
                "Session template",
                options=list(session_options.keys()),
                format_func=lambda k: session_options[k][0],
                key="session_pick",
            )
            _, tmpl_s = session_options[chosen_session_id]

            st.caption(f"**Audience:** {tmpl_s.get('audience', 'Full project team')}")
            if tmpl_s.get("variant_note"):
                st.markdown(
                    f'<span class="ce-pill-info">{tmpl_s["variant_note"]}</span>',
                    unsafe_allow_html=True,
                )

            # Editable fields
            manual_ctx_s: dict = {}
            if tmpl_s.get("editable_fields"):
                with st.expander("Fill in session details", expanded=True):
                    for field in tmpl_s["editable_fields"]:
                        key     = field["key"]
                        label   = field["label"]
                        ftype   = field.get("type", "text")
                        req     = field.get("required", False)
                        placeholder = field.get("placeholder", "")

                        if ftype == "text":
                            val = st.text_input(
                                label + (" *" if req else ""),
                                placeholder=placeholder,
                                key=f"sess_{chosen_session_id}_{key}",
                            )
                        elif ftype == "textarea":
                            val = st.text_area(
                                label + (" *" if req else ""),
                                placeholder=placeholder,
                                height=80,
                                key=f"sess_{chosen_session_id}_{key}",
                            )
                        elif ftype == "multiselect":
                            selected_opts = st.multiselect(
                                label + (" *" if req else ""),
                                options=field.get("options", []),
                                key=f"sess_{chosen_session_id}_{key}",
                            )
                            val = "\n".join(f"  • {o}" for o in selected_opts)
                        elif ftype == "select":
                            opts = field.get("options", [])
                            val = st.selectbox(
                                label,
                                options=opts,
                                key=f"sess_{chosen_session_id}_{key}",
                            )
                        else:
                            val = st.text_input(label, key=f"sess_{chosen_session_id}_{key}")

                        manual_ctx_s[key] = val

                        # For go-live readiness, resolve the display text
                        if key == "GO_LIVE_READINESS" and val:
                            option_letter = val[0]  # "A", "B", or "C"
                            readiness_map = tmpl_s.get("go_live_readiness_text", {})
                            resolved = readiness_map.get(option_letter, val)
                            # Inject HYPERCARE_DATE if needed
                            if "{HYPERCARE_DATE}" in resolved:
                                hc_date = manual_ctx_s.get("HYPERCARE_DATE", "{HYPERCARE_DATE}")
                                resolved = resolved.replace("{HYPERCARE_DATE}", hc_date)
                            manual_ctx_s["GO_LIVE_READINESS_TEXT"] = resolved

            merged_ctx_s = {**auto_ctx, **manual_ctx_s}
            subject_s, body_s = render_template(
                tmpl_s["body"], tmpl_s["subject"], {}, merged_ctx_s
            )
            missing_s = _find_missing_placeholders(body_s + subject_s)
            if missing_s:
                st.markdown(
                    f'<div class="ce-placeholder-missing">⚠️ Unfilled: {", ".join(missing_s)}</div>',
                    unsafe_allow_html=True,
                )

            subject_s_edit = st.text_input("Subject", value=subject_s, key="sess_subj")
            body_s_edit    = st.text_area("Email body", value=body_s, height=380, key="sess_body")

            ss_field_s = tmpl_s.get("ss_milestone_on_send")
            if ss_field_s:
                st.markdown(
                    f'<div class="ce-tip">On send, <strong>{ss_field_s}</strong> will be '
                    f'date-stamped in Smartsheet.</div>',
                    unsafe_allow_html=True,
                )

            if st.button(
                "Send" if st.session_state.get("_gmail_approved") else "📋 Log Send",
                key="btn_send_session",
                type="primary",
                disabled=not recipient_email,
            ):
                st.session_state["_session_send_requested"] = {
                    "subject": subject_s_edit,
                    "body": body_s_edit,
                    "ss_field": ss_field_s,
                    "template_id": tmpl_s["id"],
                    "template_name": tmpl_s["name"],
                }
                st.rerun()

            if st.session_state.get("_session_send_requested"):
                req = st.session_state.pop("_session_send_requested")
                try:
                    with st.spinner("Logging…"):
                        success, send_id_or_err = execute_send(
                            project_id=project_id,
                            template_id=req["template_id"],
                            template_name=req["template_name"],
                            subject=req["subject"],
                            body=req["body"],
                            recipient_email=recipient_email,
                            cc_emails=cc_emails,
                            ss_milestone_field=req["ss_field"],
                        )
                    if success:
                        st.success(f"✓ Logged. Send ID: `{send_id_or_err}`")
                        if req["ss_field"]:
                            did_write = _do_ss_writeback(project_id, req["ss_field"], datetime.date.today(), selected_row)
                            if did_write:
                                mark_ss_writeback_done(send_id_or_err)
                    else:
                        st.error(f"Failed: {send_id_or_err}")
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.exception(e)

    # ── TAB: Lifecycle ────────────────────────────────────────────────────────
    with tab_lifecycle:
        lifecycle_templates = list_lifecycle_templates()
        lc_options = {t["id"]: t for t in lifecycle_templates}

        chosen_lc_id = st.selectbox(
            "Lifecycle template",
            options=list(lc_options.keys()),
            format_func=lambda k: f"[{lc_options[k]['category']}]  {lc_options[k]['name']}",
            key="lc_pick",
        )
        tmpl_lc = get_lifecycle_template(chosen_lc_id)

        st.caption(f"**When to send:** {tmpl_lc['trigger']}")

        # Tips
        for tip in tmpl_lc.get("tips", []):
            st.markdown(f'<div class="ce-tip">💡 {tip}</div>', unsafe_allow_html=True)

        # Variant picker (UAT Signoff has 3 scenarios)
        variant_body_lc = tmpl_lc.get("body", "")
        if tmpl_lc.get("variants"):
            st.markdown("**Which scenario applies?**")
            variant_labels = {v["key"]: f"{v['label']} — {v['description']}" for v in tmpl_lc["variants"]}
            chosen_variant = st.radio(
                "Scenario",
                options=list(variant_labels.keys()),
                format_func=lambda k: variant_labels[k],
                key=f"lc_variant_{chosen_lc_id}",
                label_visibility="collapsed",
            )
            variant_body_lc = tmpl_lc["variant_bodies"][chosen_variant]

        # Editable fields
        manual_ctx_lc: dict = {}
        if tmpl_lc.get("editable_fields"):
            with st.expander("Fill in details", expanded=True):
                for field in tmpl_lc["editable_fields"]:
                    key       = field["key"]
                    label     = field["label"]
                    ftype     = field.get("type", "text")
                    req       = field.get("required", False)
                    source    = field.get("source", "")
                    default   = field.get("default", "")

                    # Auto-populate from DRS where available
                    if source == "drs_prod_cutover":
                        raw = selected_row.get("prod_cutover") or selected_row.get("Prod Cutover")
                        if raw:
                            try:
                                default = pd.to_datetime(raw).date().isoformat()
                            except Exception:
                                default = ""
                    elif source == "drs_project_link":
                        default = str(selected_row.get("project_link") or selected_row.get("Project Link") or "")
                    elif source == "calculated_go_live_plus_14":
                        go_live_raw = (selected_row.get("prod_cutover")
                                       or selected_row.get("go_live_date")
                                       or manual_ctx_lc.get("GO_LIVE_DATE", ""))
                        if go_live_raw:
                            try:
                                base = pd.to_datetime(go_live_raw).date()
                                default = (base + datetime.timedelta(days=14)).isoformat()
                            except Exception:
                                default = ""

                    if ftype == "date":
                        val = st.text_input(
                            label + (" *" if req else ""),
                            value=default,
                            placeholder="YYYY-MM-DD",
                            key=f"lc_{chosen_lc_id}_{key}",
                        )
                    elif ftype == "text":
                        val = st.text_input(
                            label + (" *" if req else ""),
                            value=default,
                            key=f"lc_{chosen_lc_id}_{key}",
                        )
                    else:
                        val = st.text_input(
                            label + (" *" if req else ""),
                            value=default,
                            key=f"lc_{chosen_lc_id}_{key}",
                        )
                    manual_ctx_lc[key] = val

        merged_ctx_lc = {**auto_ctx, **manual_ctx_lc}
        subject_lc, body_lc = render_template(
            variant_body_lc, tmpl_lc["subject"], {}, merged_ctx_lc
        )
        missing_lc = _find_missing_placeholders(body_lc + subject_lc)
        if missing_lc:
            st.markdown(
                f'<div class="ce-placeholder-missing">⚠️ Unfilled: {", ".join(missing_lc)}</div>',
                unsafe_allow_html=True,
            )

        subject_lc_edit = st.text_input("Subject", value=subject_lc, key="lc_subj")
        body_lc_edit    = st.text_area("Email body", value=body_lc, height=380, key="lc_body")

        ss_field_lc = tmpl_lc.get("ss_milestone_on_send")
        if ss_field_lc:
            _fields_display = ", ".join(ss_field_lc) if isinstance(ss_field_lc, list) else ss_field_lc
            st.markdown(
                f'<div class="ce-tip">On send, <strong>{_fields_display}</strong> will be '
                f'date-stamped in Smartsheet (existing values will not be overwritten).</div>',
                unsafe_allow_html=True,
            )

        if st.button(
            "Send" if st.session_state.get("_gmail_approved") else "📋 Log Send",
            key="btn_send_lc",
            type="primary",
            disabled=not recipient_email,
        ):
            st.session_state["_lc_send_requested"] = {
                "subject": subject_lc_edit,
                "body": body_lc_edit,
                "ss_field": ss_field_lc,
                "template_id": chosen_lc_id,
                "template_name": tmpl_lc["name"],
                # Carry the consultant-entered Go-Live Date so writeback uses
                # the actual cutover date rather than today's date
                "go_live_date_str": manual_ctx_lc.get("GO_LIVE_DATE", ""),
            }
            st.rerun()

        if st.session_state.get("_lc_send_requested"):
            req = st.session_state.pop("_lc_send_requested")
            try:
                with st.spinner("Logging…"):
                    success, send_id_or_err = execute_send(
                        project_id=project_id,
                        template_id=req["template_id"],
                        template_name=req["template_name"],
                        subject=req["subject"],
                        body=req["body"],
                        recipient_email=recipient_email,
                        cc_emails=cc_emails,
                        ss_milestone_field=req["ss_field"],
                    )
                if success:
                    st.success(f"✓ Logged. Send ID: `{send_id_or_err}`")
                    if req["ss_field"]:
                        # Resolve the write date:
                        # - For Go-Live Date and Prod Cutover, use the consultant-entered
                        #   go_live date from the template if available, otherwise today.
                        # - For all other fields, today's date is correct.
                        from shared.template_engine import SS_GO_LIVE_DATE, SS_PROD_CUTOVER
                        go_live_str = req.get("go_live_date_str", "")
                        try:
                            go_live_date = (
                                datetime.date.fromisoformat(go_live_str[:10])
                                if go_live_str else datetime.date.today()
                            )
                        except ValueError:
                            go_live_date = datetime.date.today()

                        fields = req["ss_field"] if isinstance(req["ss_field"], list) else [req["ss_field"]]
                        for field in fields:
                            write_date = (
                                go_live_date
                                if field in (SS_GO_LIVE_DATE, SS_PROD_CUTOVER)
                                else datetime.date.today()
                            )
                            did_write = _do_ss_writeback(project_id, field, write_date, selected_row)
                            if did_write:
                                mark_ss_writeback_done(send_id_or_err)
                else:
                    st.error(f"Failed: {send_id_or_err}")
            except Exception as e:
                st.error(f"Error: {e}")
                st.exception(e)
