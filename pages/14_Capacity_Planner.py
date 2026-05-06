"""
PS Tools — Capacity Planner
Manager tool to model consultant delivery capacity based on product mix
and Apps allocation. No data upload required — planning calculator only.
"""
import streamlit as st

st.session_state["current_page"] = "Capacity Planner"

from shared.constants import (
    EMPLOYEE_ROLES, CONSULTANT_DROPDOWN, get_role, is_manager,
)
from shared.config import (
    EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP,
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        h1,h2,h3,h4,p,div,label,button { font-family: 'Manrope', sans-serif !important; }
        .brief-header  { font-size: 24px; font-weight: 700; color: inherit; margin-bottom: 4px; }
        .brief-sub     { font-size: 13px; margin-bottom: 20px; opacity: 0.6; }
        .section-label { font-size: 13px; font-weight: 700; text-transform: uppercase;
                         letter-spacing: 0.8px; color: #08A9B7; margin-bottom: 8px; }
        .metric-card   { background: transparent; border: 1px solid rgba(128,128,128,0.2);
                         border-radius: 8px; padding: 16px 20px; margin-bottom: 12px; }
        .metric-val    { font-size: 32px; font-weight: 700; color: inherit; }
        .metric-lbl    { font-size: 13px; opacity: 0.6; margin-top: 2px; }
        .divider       { border: none; border-top: 1px solid rgba(128,128,128,0.2); margin: 16px 0; }
        .result-card   { background: rgba(8,169,183,0.08); border: 1px solid rgba(8,169,183,0.3);
                         border-radius: 8px; padding: 16px 20px; margin-bottom: 8px; }
        .result-val    { font-size: 36px; font-weight: 700; color: #08A9B7; }
        .result-lbl    { font-size: 12px; opacity: 0.6; margin-top: 2px; }
        .context-row   { display: flex; align-items: center; gap: 8px;
                         padding: 6px 8px; border-radius: 5px; margin-bottom: 4px; }
        .context-row.active { background: rgba(8,169,183,0.1); border: 1px solid rgba(8,169,183,0.25); }
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
        .product-tag   { display: inline-block; font-size: 11px; padding: 2px 8px;
                         border-radius: 4px; margin: 2px; }
    </style>
""", unsafe_allow_html=True)

# ── Identity ──────────────────────────────────────────────────────────────────
selected = st.session_state.get("consultant_name", "")
role     = get_role(selected) if selected else "consultant"
mgr      = is_manager(selected)

# ── Capacity model constants ──────────────────────────────────────────────────
BILLABLE_HRS   = 28.0   # 40h × 70%
WORKING_WEEKS  = 50     # 52 - 2wk PTO
OH_MULTIPLIER  = 1.2    # +20% PM overhead applied to hrs/wk

# Product data: blended avg hrs, timeline (wk midpoint), OH-loaded hrs/wk
# Source: Global blended averages 2025 + Q1 2026, corrected timelines
PRODUCT_DATA = {
    "Capture":        {"avg_hrs": 17.8, "timeline": 7.0,  "hrswk": 2.54, "oh_hrswk": 3.05},
    "Approvals":      {"avg_hrs": 17.5, "timeline": 7.0,  "hrswk": 2.50, "oh_hrswk": 3.00},
    "Reconcile":      {"avg_hrs": 14.9, "timeline": 8.5,  "hrswk": 1.75, "oh_hrswk": 2.10},
    "e-Invoicing":    {"avg_hrs": 10.7, "timeline": 7.0,  "hrswk": 1.53, "oh_hrswk": 1.84},
    "Reconcile PSP":  {"avg_hrs": 13.0, "timeline": 8.5,  "hrswk": 1.53, "oh_hrswk": 1.84},
    "Reconcile 2.0":  {"avg_hrs": 14.1, "timeline": 8.5,  "hrswk": 1.66, "oh_hrswk": 1.99},
    "e-Inv & Capture":{"avg_hrs": 26.9, "timeline": 8.5,  "hrswk": 3.16, "oh_hrswk": 3.79},
}

# Map EMPLOYEE_ROLES product labels → PRODUCT_DATA keys
PRODUCT_MAP = {
    "Capture":             "Capture",
    "Approvals":           "Approvals",
    "Reconcile":           "Reconcile",
    "e-Invoicing":         "e-Invoicing",
    "PSP":                 "Reconcile PSP",
    "CC Statement Import": "Reconcile",     # bundled under Reconcile
    "SFTP Connector":      None,            # too few data points — excluded
    "Payments":            None,            # T&M only — excluded from FF model
}

# Apps products only (non-FF/T&M products are modelled separately)
APPS_PRODUCTS = list(PRODUCT_DATA.keys())

# Region lookup helper
def get_region(name):
    loc = EMPLOYEE_LOCATION.get(name, "")
    if isinstance(loc, tuple): loc = loc[0]
    return PS_REGION_OVERRIDE.get(name, PS_REGION_MAP.get(loc, "Other"))

# Get modellable Apps products for a consultant from their EMPLOYEE_ROLES entry
def get_consultant_products(name):
    info = EMPLOYEE_ROLES.get(name, {})
    raw  = info.get("products", []) + info.get("learning", [])
    mapped = []
    for p in raw:
        m = PRODUCT_MAP.get(p)
        if m and m not in mapped and m in PRODUCT_DATA:
            mapped.append(m)
    return mapped if mapped else list(PRODUCT_DATA.keys())

# Calculate concurrent + throughput from weighted OH hrs/wk
def calc_capacity(mix_pct: dict, apps_alloc_pct: float):
    """
    mix_pct: {product: pct 0-100} within Apps allocation
    apps_alloc_pct: 0-100, portion of 28h billable devoted to Apps FF work
    Returns dict with concurrent, throughput, apps_hrswk, other_hrswk
    """
    apps_hrs = BILLABLE_HRS * (apps_alloc_pct / 100)
    other_hrs = BILLABLE_HRS - apps_hrs

    total_pct = sum(mix_pct.values())
    if total_pct == 0 or apps_hrs == 0:
        return {"concurrent": 0, "throughput": 0, "apps_hrswk": apps_hrs,
                "other_hrswk": other_hrs, "wt_oh_hrswk": 0, "wt_timeline": 7.0}

    # Weighted OH hrs/wk and timeline from product mix
    wt_oh    = sum((p / total_pct) * PRODUCT_DATA[prod]["oh_hrswk"]
                   for prod, p in mix_pct.items() if p > 0)
    wt_tl    = sum((p / total_pct) * PRODUCT_DATA[prod]["timeline"]
                   for prod, p in mix_pct.items() if p > 0)

    concurrent   = apps_hrs / wt_oh if wt_oh > 0 else 0
    throughput   = (WORKING_WEEKS / wt_tl) * concurrent if wt_tl > 0 else 0

    return {
        "concurrent":   concurrent,
        "throughput":   throughput,
        "apps_hrswk":   apps_hrs,
        "other_hrswk":  other_hrs,
        "wt_oh_hrswk":  wt_oh,
        "wt_timeline":  wt_tl,
    }

# ── Page header ───────────────────────────────────────────────────────────────
st.markdown('<p class="brief-header">Capacity Planner</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="brief-sub">Model consultant delivery capacity by product mix and Apps allocation. '
    'Planning calculator only — no data upload required.</p>',
    unsafe_allow_html=True
)
st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── Layout: left (apps delivery profile) | right (results + context) ──────────────────────
left, right = st.columns([1.1, 0.9], gap="large")

with left:
    st.markdown('<p class="section-label">Configure</p>', unsafe_allow_html=True)

    # ── Consultant selector ──────────────────────────────────────────────────
    all_consultants = sorted([
        c for c in CONSULTANT_DROPDOWN
        if EMPLOYEE_ROLES.get(c, {}).get("products") is not None
    ])
    default_idx = all_consultants.index(selected) if selected in all_consultants else 0

    consultant = st.selectbox(
        "Consultant",
        options=all_consultants,
        index=default_idx,
        key="cp_consultant"
    )

    consultant_products = get_consultant_products(consultant)
    region = get_region(consultant)

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    # ── Apps allocation ──────────────────────────────────────────────────────
    st.markdown("**Apps allocation** — % of billable hours devoted to ZoneApps FF delivery")
    apps_alloc = st.slider(
        "Apps allocation %",
        min_value=10, max_value=100, value=100, step=5,
        key="cp_apps_alloc",
        label_visibility="collapsed"
    )
    apps_hrs   = round(BILLABLE_HRS * apps_alloc / 100, 1)
    other_hrs  = round(BILLABLE_HRS - apps_hrs, 1)

    col_a, col_b = st.columns(2)
    col_a.metric("Apps hrs/wk", f"{apps_hrs}h")
    col_b.metric("T&M / other hrs/wk", f"{other_hrs}h",
                 help="Hours available for T&M projects, Payroll, Billing, or other non-FF work. Not modelled here.")

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    # ── Product mix sliders ──────────────────────────────────────────────────
    st.markdown("**Product mix** — % within Apps allocation")

    # Pre-populate from EMPLOYEE_ROLES, allow full override via toggle
    show_all = st.toggle(
        "Show all products (override enabled products)",
        value=False,
        key="cp_show_all",
        help="Enable to model products the consultant is not yet enabled on — e.g. upcoming enablement."
    )

    active_products = APPS_PRODUCTS if show_all else [
        p for p in APPS_PRODUCTS if p in consultant_products
    ]
    if not active_products:
        active_products = APPS_PRODUCTS

    # Distribute default percentages evenly across enabled products
    default_per = round(100 / len(active_products))
    remainder   = 100 - default_per * len(active_products)

    mix_pct = {}
    for i, prod in enumerate(active_products):
        default_val = default_per + (remainder if i == 0 else 0)
        mix_pct[prod] = st.slider(
            f"{prod}",
            min_value=0, max_value=100,
            value=st.session_state.get(f"cp_mix_{prod}", default_val),
            step=5,
            key=f"cp_mix_{prod}"
        )

    total_mix = sum(mix_pct.values())
    if total_mix == 0:
        st.markdown('<div class="warn-banner">Set at least one product to > 0% to calculate capacity.</div>',
                    unsafe_allow_html=True)
    elif total_mix != 100:
        st.markdown(
            f'<div class="warn-banner">Product mix totals <b>{total_mix}%</b> — adjust sliders to reach 100% '
            f'for accurate results. Current output is scaled proportionally.</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown('<div class="ok-banner">Product mix totals 100% ✓</div>', unsafe_allow_html=True)

    if not show_all and consultant_products:
        enabled_str = " · ".join(consultant_products)
        st.markdown(
            f'<div class="info-box">Showing products enabled for <b>{consultant.split(",")[0]}</b>: '
            f'{enabled_str}. Toggle above to model additional products.</div>',
            unsafe_allow_html=True
        )

# ── Results ───────────────────────────────────────────────────────────────────
result = calc_capacity(
    {p: v for p, v in mix_pct.items() if v > 0},
    apps_alloc
)

conc_lo = max(1, int(result["concurrent"]))
conc_hi = conc_lo + 1
throughput = round(result["throughput"])

with right:
    st.markdown('<p class="section-label">Results</p>', unsafe_allow_html=True)

    # Primary result cards
    r1, r2 = st.columns(2)
    with r1:
        st.markdown(
            f'<div class="result-card">'
            f'<div class="result-lbl">Max concurrent projects</div>'
            f'<div class="result-val">{conc_lo}–{conc_hi}</div>'
            f'</div>',
            unsafe_allow_html=True
        )
    with r2:
        st.markdown(
            f'<div class="result-card">'
            f'<div class="result-lbl">Estimated throughput</div>'
            f'<div class="result-val">~{throughput}/yr</div>'
            f'</div>',
            unsafe_allow_html=True
        )

    # Supporting metrics
    m1, m2, m3 = st.columns(3)
    m1.metric("Apps hrs/wk", f"{round(result['apps_hrswk'], 1)}h")
    m2.metric("Wtd OH hrs/wk", f"{round(result['wt_oh_hrswk'], 2)}h",
              help="Weighted average hrs/wk per project including 20% PM overhead — the divisor in the concurrent calculation.")
    m3.metric("Avg timeline", f"{round(result['wt_timeline'], 1)}wk",
              help="Weighted average project timeline across the product mix.")

    if other_hrs > 0:
        st.markdown(
            f'<div class="info-box">'
            f'<b>{other_hrs}h/wk</b> unmodelled — available for T&M, Payroll, Billing, or other non-FF work.'
            f'</div>',
            unsafe_allow_html=True
        )

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    # ── Team context panel ───────────────────────────────────────────────────
    st.markdown('<p class="section-label">Team context</p>', unsafe_allow_html=True)

    region_opts = ["Global", "EMEA", "APAC", "NOAM"]
    context_region = st.radio(
        "View",
        options=region_opts,
        index=region_opts.index(region) if region in region_opts else 0,
        horizontal=True,
        key="cp_context_region",
        label_visibility="collapsed"
    )

    # Build team comparison — use same model for each consultant's default profile
    team_rows = []
    for name in all_consultants:
        r = get_region(name)
        if context_region != "Global" and r != context_region:
            continue
        prods = get_consultant_products(name)
        if not prods:
            continue
        # Equal-split default mix for each consultant
        eq_mix = {p: 100 / len(prods) for p in prods}
        res    = calc_capacity(eq_mix, 100)
        team_rows.append({
            "name":       name.split(",")[0],
            "full_name":  name,
            "region":     r,
            "concurrent": res["concurrent"],
            "conc_lo":    max(1, int(res["concurrent"])),
            "conc_hi":    max(1, int(res["concurrent"])) + 1,
        })

    # Sort by concurrent desc
    team_rows.sort(key=lambda x: x["concurrent"], reverse=True)

    if team_rows:
        max_conc = max(r["concurrent"] for r in team_rows) or 1
        for row in team_rows:
            is_active = row["full_name"] == consultant
            bar_pct   = int((row["concurrent"] / max_conc) * 100)
            bar_color = "#08A9B7" if is_active else "rgba(128,128,128,0.3)"
            bg_style  = "background:rgba(8,169,183,0.08);border:1px solid rgba(8,169,183,0.2);" if is_active else ""

            st.markdown(
                f'<div class="context-row {"active" if is_active else ""}" style="{bg_style}">'
                f'  <span class="context-name" style="font-weight:{"700" if is_active else "400"};">'
                f'    {row["name"]}'
                f'  </span>'
                f'  <div style="flex:1;height:5px;background:rgba(128,128,128,0.15);border-radius:3px;overflow:hidden;">'
                f'    <div style="width:{bar_pct}%;height:5px;background:{bar_color};border-radius:3px;"></div>'
                f'  </div>'
                f'  <span class="context-conc">{row["conc_lo"]}–{row["conc_hi"]}</span>'
                f'</div>',
                unsafe_allow_html=True
            )
    else:
        st.caption("No consultants found for this region.")

    st.markdown(
        '<div class="info-box">Team context uses each consultant\'s enabled product list with '
        'equal-split default mix at 100% Apps allocation. Adjust the left panel to see how the '
        'selected consultant\'s profile compares.</div>',
        unsafe_allow_html=True
    )

# ── Methodology note ──────────────────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown(
    '<div class="info-box">'
    '<b>How this is calculated:</b> '
    'Concurrent = Apps hrs/wk ÷ weighted OH hrs/wk per project (blended global avg ÷ product timeline × 1.2 overhead). '
    'Throughput = (50 working weeks ÷ weighted avg timeline) × concurrent. '
    'Based on global blended actuals — 2025 full year + Q1 2026, 827 projects. '
    'See <i>Consultant Capacity Model — Methodology &amp; Product Reference</i> for full detail.'
    '</div>',
    unsafe_allow_html=True
)
