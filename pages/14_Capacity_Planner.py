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
    "ZonePayments":       {"avg_hrs": 20.5, "timeline": 8.5,  "oh_hrswk": 2.89},  # 30h scope, T&M-adjacent
    # Override-only products (limited actuals)
    "SFTP Connector":     {"avg_hrs": 23.0, "timeline": 7.0,  "oh_hrswk": 3.94},
    "CC Statement Import":{"avg_hrs": 13.5, "timeline": 7.0,  "oh_hrswk": 2.31},
    "AP Payments":        {"avg_hrs":  4.0, "timeline": 7.0,  "oh_hrswk": 0.69},  # add-on SKU, 4h scope
    "Procure":            {"avg_hrs": 20.0, "timeline": 8.5,  "oh_hrswk": 2.82},  # placeholder
}

CORE_PRODUCTS     = ["Capture", "Approvals", "Reconcile", "e-Invoicing",
                     "Reconcile PSP", "Reconcile 2.0", "ZonePayments"]
OVERRIDE_PRODUCTS = ["Procure"]
OVERRIDE_EXTRAS   = ["SFTP Connector", "CC Statement Import", "AP Payments"]
ALL_PRODUCTS      = CORE_PRODUCTS + OVERRIDE_EXTRAS + OVERRIDE_PRODUCTS

PRODUCT_MAP = {
    "Capture":             "Capture",
    "Approvals":           "Approvals",
    "Reconcile":           "Reconcile",
    "e-Invoicing":         "e-Invoicing",
    "PSP":                 "Reconcile PSP",
    "CC Statement Import": "CC Statement Import",
    "SFTP Connector":      "SFTP Connector",
    "Payments":            "ZonePayments",   # ZonePayments product (30h scope)
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
    """Return enabled parent product display names for a consultant."""
    info = EMPLOYEE_ROLES.get(name, {})
    raw  = set(info.get("products", [])) | set(info.get("learning", []))
    mapped = []
    if "Capture" in raw:                                        mapped.append("Capture")
    if "Approvals" in raw:                                      mapped.append("Approvals")
    if raw & {"Reconcile","PSP","CC Statement Import","SFTP Connector"}:
                                                                mapped.append("Reconcile")
    if "e-Invoicing" in raw:                                    mapped.append("e-Invoicing")
    if "Payments" in raw:                                       mapped.append("Payments (AR)")
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
    f"<p style='color:rgba(255,255,255,0.3);margin:10px 0 0;font-size:12px;"
    f"font-family:Manrope,sans-serif;line-height:1.6'>"
    f"Concurrent = Apps hrs/wk ÷ weighted OH hrs/wk per project &nbsp;·&nbsp; "
    f"Estimated throughput = (50 working weeks ÷ avg timeline) × concurrent &nbsp;·&nbsp; "
    f"Source: global blended actuals, rolling 12 months"
    f"</p>"
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
        help="Shows products the consultant is not yet enabled on — e.g. upcoming enablement. "
             "Procure uses placeholder data."
    )

    # Determine which parent products to show
    # Parent products only — add-ons are implied under their parent
    PARENT_PRODUCTS = ["Capture", "Approvals", "Reconcile", "e-Invoicing",
                       "Payments (AR)", "Procure"]

    # Map display name → PRODUCT_DATA key
    DISPLAY_MAP = {
        "Capture":      "Capture",
        "Approvals":    "Approvals",
        "Reconcile":    "Reconcile",
        "e-Invoicing":  "e-Invoicing",
        "Payments (AR)":"ZonePayments",
        "Procure":      "Procure",
    }

    # Add-on notes shown beneath parent sliders
    ADDON_NOTES = {
        "Capture":   ("AP Payments add-on", "#08A9B7", "#003D42",
                      "included within Capture scope"),
        "Reconcile": ("Rec 2.0 · CC · PSP · SFTP",  "#3B5998", "#C8D6FF",
                      "add-ons included"),
    }

    # Which parents are enabled for this consultant
    enabled_parents = []
    for disp, key in DISPLAY_MAP.items():
        if key in (consultant_products if not show_all else list(PRODUCT_DATA.keys())):
            enabled_parents.append(disp)
    if not enabled_parents:
        enabled_parents = ["Capture", "Approvals", "Reconcile"]

    n           = len(enabled_parents)
    default_per = round(100 / n) if n else 0
    remainder   = 100 - default_per * n

    mix_pct = {}   # keyed by PRODUCT_DATA key
    for i, disp in enumerate(enabled_parents):
        key         = DISPLAY_MAP[disp]
        default_val = default_per + (remainder if i == 0 else 0)
        is_placeholder = disp == "Procure"

        # Divider between groups
        if i > 0:
            st.markdown('<hr class="divider" style="margin:6px 0;">', unsafe_allow_html=True)

        # Label row
        lbl = f"{disp}  ⚠ placeholder" if is_placeholder else disp
        st.markdown(
            f'<div style="font-size:11px;font-weight:600;text-transform:uppercase;'
            f'letter-spacing:0.5px;color:var(--color-text-secondary);'
            f'margin-bottom:4px;">{lbl}</div>',
            unsafe_allow_html=True
        )

        val = st.slider(
            disp, min_value=0, max_value=100,
            value=st.session_state.get(f"cp_mix_{key}", default_val),
            step=5, key=f"cp_mix_{key}",
            label_visibility="collapsed"
        )
        mix_pct[key] = val

        # Add-on pills beneath parent
        if disp in ADDON_NOTES:
            pill_label, pill_bg, pill_txt, note = ADDON_NOTES[disp]
            st.markdown(
                f'<div style="margin:2px 0 2px 0;display:flex;align-items:center;gap:8px;">'
                f'<span style="background:{pill_bg};color:{pill_txt};font-size:10px;'
                f'padding:2px 7px;border-radius:4px;font-weight:600;">{pill_label}</span>'
                f'<span style="font-size:11px;color:var(--color-text-tertiary);">{note}</span>'
                f'</div>',
                unsafe_allow_html=True
            )

    # Total / validation
    total_mix = sum(mix_pct.values())
    st.markdown('<hr class="divider" style="margin:10px 0 6px;">', unsafe_allow_html=True)
    total_color = "#27AE60" if total_mix == 100 else "#D68910"
    st.markdown(
        f'<div style="display:flex;justify-content:space-between;align-items:center;'
        f'font-size:13px;margin-bottom:8px;">'
        f'<span style="color:var(--color-text-secondary);">Total</span>'
        f'<span style="font-weight:600;color:{total_color};">{total_mix}%'
        f'{" ✓" if total_mix == 100 else " — adjust to reach 100%"}</span>'
        f'</div>',
        unsafe_allow_html=True
    )

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
    # Weighted avg h/project from mix
    wt_avg_hrs = sum(
        (mix_pct.get(p, 0) / max(sum(v for v in mix_pct.values() if v > 0), 1))
        * PRODUCT_DATA[p]["avg_hrs"]
        for p in mix_pct if mix_pct[p] > 0
    )
    m1.metric("Avg h/project",  f"{round(wt_avg_hrs, 1)}h",
              help="Weighted average hours per project based on your product mix.")
    m2.metric("Avg timeline",   f"{round(result['wt_timeline'], 1)}wk",
              help="Weighted average project timeline across the product mix.")
    m3.metric("Wtd OH hrs/wk",  f"{round(result['wt_oh_hrswk'], 2)}h",
              help="Avg h/project ÷ avg timeline × 1.2 PM overhead = hrs committed per project per week. "
                   "Divides into Apps hrs/wk to give concurrent projects.")

    if other_hrs > 0:
        st.markdown(
            f'<div class="info-box"><b>{other_hrs}h/wk</b> unmodelled — '
            f'available for T&M, Payroll, Billing, or other non-FF work.</div>',
            unsafe_allow_html=True
        )

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    # ── Workload snapshot — four-metric framework ─────────────────────────────
    st.markdown('<p class="section-label">Workload snapshot</p>', unsafe_allow_html=True)

    def get_workload_metrics(name, drs, ns):
        import pandas as pd
        first    = name.split(",")[0].strip().lower()
        assigned = active_port = time_booked = None

        if drs is not None and not drs.empty and "project_manager" in drs.columns:
            mask = drs["project_manager"].astype(str).str.lower().str.contains(first, na=False)
            cdrs = drs[mask]
            if not cdrs.empty:
                id_col = next((c for c in ["project_id","project_name"] if c in cdrs.columns), None)
                if id_col:
                    assigned    = int(cdrs[id_col].nunique())
                    active_port = int(
                        cdrs[~cdrs["_on_hold"].astype(bool)][id_col].nunique()
                        if "_on_hold" in cdrs.columns
                        else cdrs[id_col].nunique()
                    )

        if ns is not None and not ns.empty and "employee" in ns.columns:
            emp_mask = ns["employee"].astype(str).str.lower().str.contains(first, na=False)
            if emp_mask.any():
                ns2 = ns[emp_mask].copy()
                if "date" in ns2.columns:
                    ns2["_month"] = pd.to_datetime(ns2["date"], errors="coerce").dt.strftime("%Y-%m")
                    ns2 = ns2[ns2["_month"] == date.today().strftime("%Y-%m")]
                if "billing_type" in ns2.columns:
                    ns2 = ns2[ns2["billing_type"].isin(["Fixed Fee","T&M"])]
                if "project_id" in ns2.columns and not ns2.empty:
                    time_booked = int(ns2["project_id"].nunique())

        return assigned, active_port, time_booked

    assigned, active_port, time_booked = get_workload_metrics(consultant, df_drs, df_ns)

    def _snap_card(val, label, sub, color="inherit"):
        return (
            f"<div style='background:rgba(128,128,128,0.05);border:1px solid "
            f"rgba(128,128,128,0.15);border-radius:8px;padding:12px 14px;'>"
            f"<div style='font-size:24px;font-weight:700;color:{color}'>"
            f"{'—' if val is None else val}</div>"
            f"<div style='font-size:11px;font-weight:600;margin-top:4px;opacity:0.8'>{label}</div>"
            f"<div style='font-size:10px;opacity:0.45;margin-top:2px'>{sub}</div>"
            f"</div>"
        )

    s1, s2, s3, s4 = st.columns(4)

    s1.markdown(_snap_card(assigned,    "Assigned",        "all DRS projects"),
                unsafe_allow_html=True)
    s2.markdown(_snap_card(active_port, "Active portfolio", "excl. on-hold"),
                unsafe_allow_html=True)

    # Time-booked colour — compare against active portfolio
    if time_booked is not None and active_port is not None:
        gap      = active_port - time_booked
        tb_color = "#27AE60" if gap == 0 else ("#D68910" if gap <= 3 else "#C0392B")
        tb_sub   = "all active booked" if gap == 0 else f"{gap} silent this month"
    else:
        tb_color, tb_sub = "inherit", "NS time · this month"

    s3.markdown(_snap_card(time_booked, "Time-booked",    tb_sub, color=tb_color),
                unsafe_allow_html=True)
    s4.markdown(_snap_card(f"{conc_lo}–{conc_hi}", "Delivery capacity",
                            "model · full engagement", color="#08A9B7"),
                unsafe_allow_html=True)

    if assigned is None and time_booked is None:
        st.markdown(
            '<div class="info-box" style="margin-top:8px;">Load DRS and/or NS time report '
            'from Home to populate this section.</div>',
            unsafe_allow_html=True
        )
    elif time_booked is not None and active_port is not None and (active_port - time_booked) > 0:
        gap = active_port - time_booked
        st.markdown(
            f'<div class="warn-banner" style="margin-top:8px;">'
            f'{gap} project{"s" if gap > 1 else ""} in active portfolio with no time booked '
            f'this month — potential silent customer{"s" if gap > 1 else ""}. '
            f'Check Customer Re-engagement for details.</div>',
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

    PILL_LABELS = {
        "Capture":            "Cap",
        "Approvals":          "App",
        "Reconcile":          "Rec",
        "e-Invoicing":        "eInv",
        "Reconcile PSP":      "PSP",
        "Reconcile 2.0":      "Rec2",
        "ZonePayments":       "Pay",
        "SFTP Connector":     "SFTP",
        "CC Statement Import":"CC",
        "AP Payments":        "APay",
        "Procure":            "Proc",
    }
    PILL_COLORS = {
        # (bg, text)
        "Capture":             ("#08A9B7", "#003D42"),   # teal — primary
        "Approvals":           ("#FF4B40", "#4A0E0A"),   # orange — primary
        "Reconcile":           ("#3B5998", "#C8D6FF"),   # navy — primary
        "e-Invoicing":         ("#7C3AED", "#EDE9FE"),   # purple — standalone
        "Reconcile PSP":       ("#3B5998", "#C8D6FF"),   # reconcile family
        "Reconcile 2.0":       ("#3B5998", "#C8D6FF"),   # reconcile family
        "SFTP Connector":      ("#3B5998", "#C8D6FF"),   # reconcile family
        "CC Statement Import": ("#3B5998", "#C8D6FF"),   # reconcile family
        "ZonePayments":        ("#0E6655", "#A2D9CE"),   # dark teal/green — standalone
        "AP Payments":         ("#08A9B7", "#003D42"),   # capture family (add-on)
        "Procure":             ("#555555", "#CCCCCC"),   # grey — placeholder
    }

    if team_rows:
        max_conc = max(r["concurrent"] for r in team_rows) or 1
        for row in team_rows:
            is_active = row["full_name"] == consultant
            bar_pct   = int((row["concurrent"] / max_conc) * 100)
            bar_color = "#08A9B7" if is_active else "rgba(128,128,128,0.3)"
            bg_style  = "background:rgba(8,169,183,0.08);border:1px solid rgba(8,169,183,0.2);" \
                        if is_active else ""

            # Build pill HTML for enabled products
            prods      = get_consultant_products(row["full_name"])
            pills_html = "".join(
                f"<span style='font-size:11px;padding:3px 8px;border-radius:4px;"
                f"background:{PILL_COLORS.get(p, ('#444', '#ccc'))[0]};"
                f"color:{PILL_COLORS.get(p, ('#444', '#ccc'))[1]};"
                f"font-weight:600;margin-right:4px;white-space:nowrap;'>"
                f"{PILL_LABELS.get(p, p)}</span>"
                for p in prods
            )

            st.markdown(
                f'<div class="context-row" style="{bg_style}display:flex;align-items:center;gap:8px;">'
                f'<span style="font-size:12px;min-width:72px;font-weight:{"700" if is_active else "400"};">'
                f'{row["name"]}</span>'
                f'<span style="flex:1;display:flex;align-items:center;gap:0;">{pills_html}</span>'
                f'<div style="width:80px;height:5px;background:rgba(128,128,128,0.15);'
                f'border-radius:3px;overflow:hidden;flex-shrink:0;">'
                f'<div style="width:{bar_pct}%;height:5px;background:{bar_color};'
                f'border-radius:3px;"></div></div>'
                f'<span style="font-size:12px;font-weight:700;min-width:40px;text-align:right;">'
                f'{row["conc_lo"]}–{row["conc_hi"]}</span>'
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

# ── end ──────────────────────────────────────────────────────────────────────
