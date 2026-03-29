"""
PS Tools — Fixed Fee Revenue Report
Straight-line revenue recognition across YTD / QTD / MTD
by Region and Product. All figures in USD.
"""
import streamlit as st
import pandas as pd
import io
from datetime import date

from shared.loaders import load_revenue, calc_monthly_slices
from shared.config import CURRENCY_REGION_MAP, FX_RATES_TO_USD

st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
<style>
    html,body,[class*="css"]{font-family:'Manrope',sans-serif!important}
    .section-label{font-size:11px;font-weight:700;text-transform:uppercase;
                   letter-spacing:.8px;color:#4472C4;margin-bottom:8px}
    .metric-card{border:1px solid rgba(128,128,128,.2);border-radius:8px;
                 padding:16px 20px;margin-bottom:8px}
    .metric-val{font-size:28px;font-weight:700}
    .metric-sub{font-size:12px;font-weight:600;opacity:.7;margin-top:2px}
    .metric-lbl{font-size:11px;opacity:.55;margin-top:1px}
    .divider{border:none;border-top:1px solid rgba(128,128,128,.2);margin:20px 0}
    .fx-note{font-size:11px;opacity:.5;font-style:italic}
</style>
""", unsafe_allow_html=True)

# ── Identity / access ─────────────────────────────────────────────────────────
from shared.constants import get_role
_role = get_role(st.session_state.get("consultant_name",""))
if _role not in ("manager","manager_only"):
    st.warning("This page is available to managers only.")
    st.stop()

today      = pd.Timestamp.today().normalize()
this_month = today.strftime("%Y-%m")
this_q     = (today.month - 1) // 3 + 1
q_months   = [f"{today.year}-{m:02d}" for m in range((this_q-1)*3+1, (this_q-1)*3+4)]
ytd_months = [f"{today.year}-{m:02d}" for m in range(1, today.month+1)]

# ── Data ─────────────────────────────────────────────────────────────────────
df_rev_raw = st.session_state.get("df_revenue")
df_drs     = st.session_state.get("df_drs")

if df_rev_raw is None:
    st.info("Upload the NS FF Revenue Charges export in the sidebar to load this report.")
    st.stop()

# ── Build monthly slices ──────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def _get_slices(df_hash):
    return calc_monthly_slices(st.session_state["df_revenue"])

slices = calc_monthly_slices(df_rev_raw)

# ── Join DRS for project name + consultant ────────────────────────────────────
if df_drs is not None and "project_id" in df_drs.columns:
    drs_lookup = df_drs[["project_id","project_name","project_manager","phase"]].copy()
    drs_lookup["project_id"] = drs_lookup["project_id"].astype(str).str.strip().str.split(".").str[0]
    slices = slices.merge(drs_lookup, on="project_id", how="left")
else:
    slices["project_name"]    = slices["project_id"]
    slices["project_manager"] = ""
    slices["phase"]           = ""

# ── Period helpers ────────────────────────────────────────────────────────────
def _sum(df, periods, col="usd_amount"):
    return df[df["period"].isin(periods)][col].sum()

def _fmt(v):
    """Format USD value: $1,234k or $1.2M"""
    if abs(v) >= 1_000_000:
        return f"${v/1_000_000:.2f}M"
    if abs(v) >= 1_000:
        return f"${v/1_000:.1f}k"
    return f"${v:,.0f}"

# ── MTD: pro-rate current month by days elapsed ──────────────────────────────
days_elapsed  = today.day
days_in_month = pd.Timestamp(today.year, today.month, 1).days_in_month
mtd_scale     = days_elapsed / days_in_month

slices_mtd = slices[slices["period"] == this_month].copy()
slices_mtd["usd_amount"] = slices_mtd["usd_amount"] * mtd_scale

ytd_df  = slices[slices["period"].isin(ytd_months)].copy()
# Replace current month in YTD with pro-rated MTD
ytd_df  = ytd_df[ytd_df["period"] != this_month]
ytd_df  = pd.concat([ytd_df, slices_mtd], ignore_index=True)

qtd_df  = slices[slices["period"].isin(q_months)].copy()
qtd_df  = qtd_df[qtd_df["period"] != this_month]
qtd_df  = pd.concat([qtd_df, slices_mtd], ignore_index=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(
    f'<div style="font-size:24px;font-weight:700;margin-bottom:4px">'
    f'Revenue Report — Fixed Fee Implementation</div>',
    unsafe_allow_html=True)
st.markdown(
    f'<div style="font-size:13px;opacity:.6;margin-bottom:4px">'
    f'All figures in USD · {today.strftime("%A, %B %-d %Y")} · '
    f'MTD pro-rated ({days_elapsed}/{days_in_month} days)</div>',
    unsafe_allow_html=True)
st.markdown(
    '<div class="fx-note">FX rates: monthly averages — update in shared/config.py</div>',
    unsafe_allow_html=True)
st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Top-line metric cards
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Total Revenue</div>',unsafe_allow_html=True)

ytd_total  = ytd_df["usd_amount"].sum()
qtd_total  = qtd_df["usd_amount"].sum()
mtd_total  = slices_mtd["usd_amount"].sum()
full_month = slices[slices["period"]==this_month]["usd_amount"].sum()

c1,c2,c3,c4 = st.columns(4)
with c1:
    st.markdown(
        f'<div class="metric-card"><div class="metric-val">{_fmt(ytd_total)}</div>'
        f'<div class="metric-sub">YTD Revenue</div>'
        f'<div class="metric-lbl">Jan 1 – {today.strftime("%-d %b %Y")}</div></div>',
        unsafe_allow_html=True)
with c2:
    st.markdown(
        f'<div class="metric-card"><div class="metric-val">{_fmt(qtd_total)}</div>'
        f'<div class="metric-sub">QTD Revenue</div>'
        f'<div class="metric-lbl">Q{this_q} {today.year}</div></div>',
        unsafe_allow_html=True)
with c3:
    st.markdown(
        f'<div class="metric-card"><div class="metric-val">{_fmt(mtd_total)}</div>'
        f'<div class="metric-sub">MTD Revenue</div>'
        f'<div class="metric-lbl">{today.strftime("%B")} (actual to date)</div></div>',
        unsafe_allow_html=True)
with c4:
    st.markdown(
        f'<div class="metric-card"><div class="metric-val">{_fmt(full_month)}</div>'
        f'<div class="metric-sub">Full Month Forecast</div>'
        f'<div class="metric-lbl">{today.strftime("%B")} (full month)</div></div>',
        unsafe_allow_html=True)

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — By Region
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Revenue by Region</div>',unsafe_allow_html=True)

def _region_table(df, label):
    if df.empty: return pd.DataFrame()
    t = df.groupby("region")["usd_amount"].sum().reset_index()
    t.columns = ["Region", label]
    t[label] = t[label].apply(_fmt)
    return t.sort_values("Region")

_rt_ytd = _region_table(ytd_df,  "YTD")
_rt_qtd = _region_table(qtd_df,  "QTD")
_rt_mtd = _region_table(slices_mtd, "MTD")

if not _rt_ytd.empty:
    _rt = _rt_ytd.merge(_rt_qtd, on="Region", how="outer").merge(_rt_mtd, on="Region", how="outer").fillna("$0")
    st.dataframe(_rt, use_container_width=True, hide_index=True)
else:
    st.info("No region data available.")

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — By Product
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Revenue by Product</div>',unsafe_allow_html=True)

def _product_table(df, label):
    if df.empty: return pd.DataFrame()
    t = df.groupby("product")["usd_amount"].sum().reset_index()
    t.columns = ["Product", label]
    t[label] = t[label].apply(_fmt)
    return t.sort_values("Product")

_pt_ytd = _product_table(ytd_df,  "YTD")
_pt_qtd = _product_table(qtd_df,  "QTD")
_pt_mtd = _product_table(slices_mtd, "MTD")

if not _pt_ytd.empty:
    _pt = _pt_ytd.merge(_pt_qtd, on="Product", how="outer").merge(_pt_mtd, on="Product", how="outer").fillna("$0")
    st.dataframe(_pt, use_container_width=True, hide_index=True)
else:
    st.info("No product data available.")

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — Monthly trend
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Monthly Trend (USD)</div>',unsafe_allow_html=True)

_trend = slices.groupby("period")["usd_amount"].sum().reset_index()
_trend.columns = ["Period", "Revenue (USD)"]
_trend = _trend.sort_values("Period")
_trend["Revenue (USD)"] = _trend["Revenue (USD)"].round(2)
st.dataframe(_trend, use_container_width=True, hide_index=True)

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — Project detail
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Project Detail</div>',unsafe_allow_html=True)

_detail_cols = ["project_id","project_name","project_manager","product",
                "region","currency","period","local_amount","usd_amount"]
_detail_cols = [c for c in _detail_cols if c in slices.columns]
_detail = slices[_detail_cols].copy()
_detail["usd_amount"]   = _detail["usd_amount"].round(2)
_detail["local_amount"] = _detail["local_amount"].round(2)

st.dataframe(_detail, use_container_width=True, hide_index=True,
             column_config={
                 "project_id":      st.column_config.TextColumn("Project ID",  width="small"),
                 "project_name":    st.column_config.TextColumn("Project",     width="large"),
                 "project_manager": st.column_config.TextColumn("Consultant",  width="medium"),
                 "product":         st.column_config.TextColumn("Product",     width="small"),
                 "region":          st.column_config.TextColumn("Region",      width="small"),
                 "currency":        st.column_config.TextColumn("Currency",    width="small"),
                 "period":          st.column_config.TextColumn("Period",      width="small"),
                 "local_amount":    st.column_config.NumberColumn("Local Amt", width="small", format="%.2f"),
                 "usd_amount":      st.column_config.NumberColumn("USD Amt",   width="small", format="%.2f"),
             })

# ── Excel download ─────────────────────────────────────────────────────────
st.markdown("")
_buf = io.BytesIO()
with pd.ExcelWriter(_buf, engine="xlsxwriter") as _xl:
    # Summary sheet
    _summ = pd.DataFrame({
        "Metric":  ["YTD Revenue","QTD Revenue","MTD Revenue (actual)","Full Month Forecast"],
        "USD":     [round(ytd_total,2), round(qtd_total,2), round(mtd_total,2), round(full_month,2)],
    })
    _summ.to_excel(_xl, sheet_name="Summary",    index=False)
    if not _rt.empty: _rt.to_excel(_xl, sheet_name="By Region",  index=False)
    if not _pt.empty: _pt.to_excel(_xl, sheet_name="By Product", index=False)
    _trend.to_excel(_xl,                          sheet_name="Monthly Trend", index=False)
    _detail.to_excel(_xl,                         sheet_name="Project Detail", index=False)

st.download_button(
    "⬇ Download Revenue Report (Excel)",
    data=_buf.getvalue(),
    file_name=f"ff_revenue_{today.strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.markdown(
    '<div style="font-size:11px;opacity:.4;text-align:center;margin-top:20px">'
    'PS Reporting Tools · Internal use only · Fixed Fee Implementation Revenue</div>',
    unsafe_allow_html=True)
