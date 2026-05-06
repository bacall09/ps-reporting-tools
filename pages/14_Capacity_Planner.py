"""
PS Tools — Capacity Planner
Manager tool to model consultant delivery capacity based on product mix
and Apps allocation. No data upload required — planning calculator only.
"""
import streamlit as st
from datetime import date

st.session_state["current_page"] = "Capacity Planner"

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN, get_role, is_manager,
)
from shared.config import EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        h1,h2,h3,h4,p,div,label,button { font-family: 'Manrope', sans-serif !important; }
        .section-label { font-size: 13px; font-weight: 700; text-transform: uppercase;
                         letter-spacing: 0.8px; color: #08A9B7; margin-bottom: 8px; }
        .divider       { border: none; border-top: 1px solid rgba(128,128,128,0.2); margin: 16px 0; }
        .result-card   { background: rgba(8,169,183,0.08); border: 1px solid rgba(8,169,183,0.3);
                         border-radius: 8px; padding: 16px 20px; margin-bottom: 8px; }
        .result-val    { font-size: 36px; font-weight: 700; color: #08A9B7; }
        .result-lbl    { font-size: 12px; opacity: 0.6; margin-top: 2px; }
        .ontrack-card  { border-radius: 8px; padding: 14px 18px; margin-bottom: 8px; }
        .ontrack-green { background: rgba(39,174,96,0.08); border: 1px solid rgba(39,174,96,0.35); }
        .ontrack-amber { background: rgba(243,156,18,0.08); border: 1px solid rgba(243,156,18,0.35); }
        .ontrack-gray  { background: rgba(128,128,128,0.06); border: 1px solid rgba(128,128,128,0.2); }
        .ontrack-val-green { font-size: 22px; font-weight: 700; color: #27AE60; }
        .ontrack-val-amber { font-size: 22px; font-weight: 700; color: #D68910; }
        .ontrack-val-gray  { font-size: 22px; font-weight: 700; opacity: 0.5; }
        .ontrack-lbl   { font-size: 11px; opacity: 0.6; margin-top: 2px; }
        .context-row   { display: flex; align-items: center; gap: 8px;
                         padding: 6px 8px; border-radius: 5px; margin-bottom: 4px; }
        .context-name  { font-size: 12px; min-width: 80px; }
        .context-conc  { font-size: 12px; font-weight: 700; min-width: 48px; text-align: right; }
        .warn-banner   { background: rgba(243,156,18,0.12); border: 1px solid rgba(243,156,18,0.4);
                         border-radius: 6px; padding: 8px 12px; font-size: 12px; color: #D68910;
                         margin-bottom: 12px; }
        .ok-banner     { background: rgba(39,174,96,0.1); border: 1px solid rgba(39,174,96,0.3);
                         border-radius: 6px; padding: 8px 12px; font-size: 12px; color: #27AE60;
                         margin-bottom: 12px; }
        .info-box      { background: rgba(128,128,128,0.06); border-radius: 6px;
                         padding: 10px 14px; font-size: 12px; opacity: 0.75; margin-top: 8px; }
    </style>
""", unsafe_allow_html=True)

# ── Identity ──────────────────────────────────────────────────────────────────
selected = st.session_state.get("consultant_name", "")
role     = get_role(selected) if selected else "consultant"
today    = date.today()

# ── Capacity model constants ──────────────────────────────────────────────────
BILLABLE_HRS  = 28.0
WORKING_WEEKS = 50

PRODUCT_DATA = {
    "Capture":            {"avg_hrs": 17.8, "timeline": 7.0,  "oh_hrswk": 3.05},
    "Approvals":          {"avg_hrs": 17.5, "timeline": 7.0,  "oh_hrswk": 3.00},
    "Reconcile":          {"avg_hrs": 14.9, "timeline": 8.5,  "oh_hrswk": 2.10},
    "e-Invoicing":        {"avg_hrs": 10.7, "timeline": 7.0,  "oh_hrswk": 1.84},
    "Reconcile PSP":      {"avg_hrs": 13.0, "timeline": 8.5,  "oh_hrswk": 1.84},
    "Reconcile 2.0":      {"avg_hrs": 14.1, "timeline": 8.5,  "oh_hrswk": 1.99},
    "SFTP Connector":     {"avg_hrs": 23.0, "timeline": 7.0,  "oh_hrswk": 3.94},
    "CC Statement Import":{"avg_hrs": 13.5, "timeline": 7.0,  "oh_hrswk": 2.31},
    "AP Payments":        {"avg_hrs": 20.0, "timeline": 8.5,  "oh_hrswk": 2.82},
    "Procure":            {"avg_hrs": 20.0, "timeline": 8.5,  "oh_hrswk": 2.82},
}

CORE_PRODUCTS     = ["Capture", "Approvals", "Reconcile", "e-Invoicing",
                     "Reconcile PSP", "Reconcile 2.0"]
OVERRIDE_PRODUCTS = ["SFTP Connector", "CC Statement Import", "AP Payments", "Procure"]
ALL_PRODUCTS      = CORE_PRODUCTS + OVERRIDE_PRODUCTS

PRODUCT_MAP = {
    "Capture":             "Capture",
    "Approvals":           "Approvals",
    "Reconcile":           "Reconcile",
    "e-Invoicing":         "e-Invoicing",
    "PSP":                 "Reconcile PSP",
    "CC Statement Import": "CC Statement Import",
    "SFTP Connector":      "SFTP Connector",
    "Payments":            "AP Payments",
}

APPS_PRODUCT_FAMILIES = {
    "Capture", "Approvals", "Reconcile", "e-Invoicing", "PSP",
    "CC Statement Import", "SFTP Connector", "Payments", "Reconcile 2.0",
}

def get_region(name):
    loc = EMPLOYEE_LOCATION.get(name, "")
    if isinstance(loc, tuple): loc = loc[0]
    return PS_REGION_OVERRIDE.get(name, PS_REGION_MAP.get(loc, "Other"))

def is_apps_consultant(name):
    info = EMPLOYEE_ROLES.get(name, {})
    all_prods = set(info.get("products", [])) | set(info.get("learning", []))
    return bool(all_prods & APPS_PRODUCT_FAMILIES)

def get_consultant_products(name):
    info   = EMPLOYEE_ROLES.get(name, {})
    raw    = info.get("products", []) + info.get("learning", [])
    mapped = []
    for p in raw:
        m = PRODUCT_MAP.get(p)
        if m and m not in mapped and m in CORE_PRODUCTS:
            mapped.append(m)
    return mapped

def calc_capacity(mix_pct: dict, apps_alloc_pct: float):
    apps_hrs  = BILLABLE_HRS * (apps_alloc_pct / 100)
    other_hrs = BILLABLE_HRS - apps_hrs
    total_pct = sum(mix_pct.values())
    if total_pct == 0 or apps_hrs == 0:
        return {"concurrent": 0, "throughput": 0, "apps_hrswk": apps_hrs,
                "other_hrswk": other_hrs, "wt_oh_hrswk": 0, "wt_timeline": 7.0}
    wt_oh = sum((p / total_pct) * PRODUCT_DATA[prod]["oh_hrswk"]
                for prod, p in mix_pct.items() if p > 0)
    wt_tl = sum((p / total_pct) * PRODUCT_DATA[prod]["timeline"]
                for prod, p in mix_pct.items() if p > 0)
    concurrent = apps_hrs / wt_oh if wt_oh > 0 else 0
    throughput = (WORKING_WEEKS / wt_tl) * concurrent if wt_tl > 0 else 0
    return {"concurrent": concurrent, "throughput": throughput,
            "apps_hrswk": apps_hrs, "other_hrswk": other_hrs,
            "wt_oh_hrswk": wt_oh, "wt_timeline": wt_tl}

# ── Apps consultant list ──────────────────────────────────────────────────────
apps_consultants = sorted([c for c in CONSULTANT_DROPDOWN if is_apps_consultant(c)])

# ── Hero banner ───────────────────────────────────────────────────────────────
_vp       = [p.strip() for p in selected.split(",")] if selected else []
_display  = f"{_vp[1].split()[0]} {_vp[0]}" if len(_vp) == 2 else (selected or "Manager")
st.markdown(
    f"<div style='background:#050D1F;padding:32px 40px 28px;border-radius:10px;"
    f"margin-bottom:24px;font-family:Manrope,sans-serif;position:relative;overflow:hidden'>"
    f"<div style='font-size:13px;font-weight:700;letter-spacing:2.5px;text-transform:uppercase;"
    f"color:#3B9EFF;margin-bottom:10px'>Professional Services · Management</div>"
    f"<h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif'>"
    f"Capacity Planner</h1>"
    f"<p style='color:rgba(255,255,255,0.45);margin:6px 0 0;font-size:14px;"
    f"font-family:Manrope,sans-serif'>{_display} · {today.strftime('%B %Y')}</p>"
    f"</div>",
    unsafe_allow_html=True,
)

# ── Layout ────────────────────────────────────────────────────────────────────
left, right = st.columns([1.1, 0.9], gap="large")

with left:
    st.markdown('<p class="section-label">Configure</p>', unsafe_allow_html=True)

    default_idx = apps_consultants.index(selected) if selected in apps_consultants else 0
    consultant  = st.selectbox("Consultant", options=apps_consultants,
                               index=default_idx, key="cp_consultant")
    consultant_products = get_consultant_products(consultant)
    region = get_region(consultant)

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    st.markdown("**Apps allocation** — % of billable hours devoted to ZoneApps FF delivery")
    apps_alloc = st.slider("Apps allocation %", min_value=10, max_value=100,
                           value=100, step=5, key="cp_apps_alloc",
                           label_visibility="collapsed")
    apps_hrs  = round(BILLABLE_HRS * apps_alloc / 100, 1)
    other_hrs = round(BILLABLE_HRS - apps_hrs, 1)

    col_a, col_b = st.columns(2)
    col_a.metric("Apps hrs/wk", f"{apps_hrs}h")
    col_b.metric("T&M / other hrs/wk", f"{other_hrs}h",
                 help="Hours available for T&M, Payroll, Billing, or other non-FF work.")

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    st.markdown("**Product mix** — % within Apps allocation")
    show_all = st.toggle(
        "Show all products (override enabled products)",
        value=False, key="cp_show_all",
        help="Adds SFTP, CC Statement Import, AP Payments, and Procure. "
             "Placeholder data used for products with limited actuals."
    )

    active_products = ALL_PRODUCTS if show_all else (
        consultant_products if consultant_products else CORE_PRODUCTS
    )

    n           = len(active_products)
    default_per = round(100 / n) if n else 0
    remainder   = 100 - default_per * n

    mix_pct = {}
    for i, prod in enumerate(active_products):
        default_val = default_per + (remainder if i == 0 else 0)
        label = f"{prod}  ⚠ placeholder" if prod in OVERRIDE_PRODUCTS else prod
        mix_pct[prod] = st.slider(
            label, min_value=0, max_value=100,
            value=st.session_state.get(f"cp_mix_{prod}", default_val),
            step=5, key=f"cp_mix_{prod}"
        )

    total_mix = sum(mix_pct.values())
    if total_mix == 0:
        st.markdown('<div class="warn-banner">Set at least one product to > 0%.</div>',
                    unsafe_allow_html=True)
    elif total_mix != 100:
        st.markdown(
            f'<div class="warn-banner">Product mix totals <b>{total_mix}%</b> — '
            f'adjust to reach 100% for accurate results.</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown('<div class="ok-banner">Product mix totals 100% ✓</div>',
                    unsafe_allow_html=True)

    if not show_all and consultant_products:
        st.markdown(
            f'<div class="info-box">Pre-populated from enabled products for '
            f'<b>{consultant.split(",")[0]}</b>. '
            f'Toggle above to model additional products.</div>',
            unsafe_allow_html=True
        )

# ── Results ───────────────────────────────────────────────────────────────────
result     = calc_capacity({p: v for p, v in mix_pct.items() if v > 0}, apps_alloc)
conc_lo    = max(1, int(result["concurrent"]))
conc_hi    = conc_lo + 1
throughput = round(result["throughput"])

# ── On-track helper ───────────────────────────────────────────────────────────
df_drs = st.session_state.get("df_drs")
df_ns  = st.session_state.get("df_ns")

def get_active_project_count(name, drs, ns):
    import pandas as pd
    first = name.split(",")[0].lower()
    # Try DRS
    if drs is not None and not drs.empty:
        for col in ["employee", "consultant", "resource", "assigned_to"]:
            if col in drs.columns:
                mask = drs[col].astype(str).str.lower().str.contains(first, na=False)
                if mask.any():
                    id_col = next((c for c in ["project_id", "project_name"] if c in drs.columns), None)
                    return int(drs[mask][id_col].nunique()) if id_col else int(mask.sum())
    # Try NS time — this month only, FF+T&M, distinct project IDs
    if ns is not None and not ns.empty:
        emp_col  = next((c for c in ["Employee", "employee"] if c in ns.columns), None)
        date_col = next((c for c in ["Date", "date"] if c in ns.columns), None)
        proj_col = next((c for c in ["Project ID", "project_id", "Project"] if c in ns.columns), None)
        bt_col   = next((c for c in ["Billing Type", "billing_type"] if c in ns.columns), None)
        if emp_col and date_col and proj_col:
            ns2 = ns.copy()
            ns2["_month"] = pd.to_datetime(ns2[date_col], errors="coerce").dt.strftime("%Y-%m")
            mask = ns2[emp_col].astype(str).str.lower().str.contains(first, na=False)
            this = ns2[mask & (ns2["_month"] == date.today().strftime("%Y-%m"))]
            if bt_col:
                this = this[this[bt_col].isin(["Fixed Fee", "T&M"])]
            return int(this[proj_col].nunique())
    return None

with right:
    st.markdown('<p class="section-label">Results</p>', unsafe_allow_html=True)

    r1, r2 = st.columns(2)
    with r1:
        st.markdown(
            f'<div class="result-card"><div class="result-lbl">Max concurrent projects</div>'
            f'<div class="result-val">{conc_lo}–{conc_hi}</div></div>',
            unsafe_allow_html=True
        )
    with r2:
        st.markdown(
            f'<div class="result-card"><div class="result-lbl">Estimated throughput</div>'
            f'<div class="result-val">~{throughput}/yr</div></div>',
            unsafe_allow_html=True
        )

    m1, m2, m3 = st.columns(3)
    m1.metric("Apps hrs/wk",   f"{round(result['apps_hrswk'], 1)}h")
    m2.metric("Wtd OH hrs/wk", f"{round(result['wt_oh_hrswk'], 2)}h",
              help="Weighted avg hrs/wk per project incl. 20% PM overhead.")
    m3.metric("Avg timeline",  f"{round(result['wt_timeline'], 1)}wk")

    if other_hrs > 0:
        st.markdown(
            f'<div class="info-box"><b>{other_hrs}h/wk</b> unmodelled — '
            f'available for T&M, Payroll, Billing, or other non-FF work.</div>',
            unsafe_allow_html=True
        )

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    # ── On-track ─────────────────────────────────────────────────────────────
    st.markdown('<p class="section-label">On track?</p>', unsafe_allow_html=True)
    active_count = get_active_project_count(consultant, df_drs, df_ns)

    if active_count is None:
        st.markdown(
            '<div class="ontrack-card ontrack-gray">'
            '<div class="ontrack-val-gray">—</div>'
            '<div class="ontrack-lbl">Load DRS or NS time report from Home to see actuals</div>'
            '</div>',
            unsafe_allow_html=True
        )
    else:
        if active_count <= conc_hi:
            card_cls, val_cls = "ontrack-green", "ontrack-val-green"
            status = f"On track — {active_count} active vs. {conc_lo}–{conc_hi} model"
        elif active_count <= conc_hi + 2:
            card_cls, val_cls = "ontrack-amber", "ontrack-val-amber"
            status = f"Near capacity — {active_count} active vs. {conc_lo}–{conc_hi} model"
        else:
            card_cls, val_cls = "ontrack-amber", "ontrack-val-amber"
            status = f"Over model capacity — {active_count} active vs. {conc_lo}–{conc_hi} model"

        st.markdown(
            f'<div class="ontrack-card {card_cls}">'
            f'<div class="{val_cls}">{active_count} active</div>'
            f'<div class="ontrack-lbl">{status}</div>'
            f'</div>',
            unsafe_allow_html=True
        )
        st.markdown(
            '<div class="info-box">Active count from current session data. '
            'Reflects portfolio size — not all may be at full engagement. '
            'See <i>Portfolio Size</i> for NS active vs. meaningfully engaged context.</div>',
            unsafe_allow_html=True
        )

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    # ── Team context ─────────────────────────────────────────────────────────
    st.markdown('<p class="section-label">Team context</p>', unsafe_allow_html=True)

    region_opts    = ["Global", "EMEA", "APAC", "NOAM"]
    default_region = region if region in region_opts else "Global"
    context_region = st.radio(
        "View", options=region_opts,
        index=region_opts.index(default_region),
        horizontal=True, key="cp_context_region",
        label_visibility="collapsed"
    )

    team_rows = []
    for name in apps_consultants:
        r = get_region(name)
        if context_region != "Global" and r != context_region:
            continue
        prods = get_consultant_products(name)
        if not prods:
            continue
        eq_mix = {p: 100 / len(prods) for p in prods}
        res    = calc_capacity(eq_mix, 100)
        team_rows.append({
            "name":       name.split(",")[0],
            "full_name":  name,
            "concurrent": res["concurrent"],
            "conc_lo":    max(1, int(res["concurrent"])),
            "conc_hi":    max(1, int(res["concurrent"])) + 1,
        })

    team_rows.sort(key=lambda x: x["concurrent"], reverse=True)

    if team_rows:
        max_conc = max(r["concurrent"] for r in team_rows) or 1
        for row in team_rows:
            is_active = row["full_name"] == consultant
            bar_pct   = int((row["concurrent"] / max_conc) * 100)
            bar_color = "#08A9B7" if is_active else "rgba(128,128,128,0.3)"
            bg_style  = "background:rgba(8,169,183,0.08);border:1px solid rgba(8,169,183,0.2);" \
                        if is_active else ""
            st.markdown(
                f'<div class="context-row" style="{bg_style}">'
                f'<span class="context-name" style="font-weight:{"700" if is_active else "400"};">'
                f'{row["name"]}</span>'
                f'<div style="flex:1;height:5px;background:rgba(128,128,128,0.15);'
                f'border-radius:3px;overflow:hidden;">'
                f'<div style="width:{bar_pct}%;height:5px;background:{bar_color};'
                f'border-radius:3px;"></div></div>'
                f'<span class="context-conc">{row["conc_lo"]}–{row["conc_hi"]}</span>'
                f'</div>',
                unsafe_allow_html=True
            )
        st.markdown(
            '<div class="info-box">Team context uses each consultant\'s enabled product list '
            'with equal-split default mix at 100% Apps allocation.</div>',
            unsafe_allow_html=True
        )
    else:
        st.caption("No Apps consultants found for this region.")

# ── Methodology note ──────────────────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown(
    '<div class="info-box">'
    '<b>How this is calculated:</b> '
    'Concurrent = Apps hrs/wk ÷ weighted OH hrs/wk per project '
    '(blended global avg ÷ product timeline × 1.2 PM overhead). '
    'Throughput = (50 working weeks ÷ weighted avg timeline) × concurrent. '
    'Source: global blended actuals 2025 + Q1 2026, 827 projects. '
    'See <i>Consultant Capacity Model — Methodology &amp; Product Reference</i> for full detail.'
    '</div>',
    unsafe_allow_html=True
)
