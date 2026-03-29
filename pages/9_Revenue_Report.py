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
# Guarantee numeric dtype regardless of how slices were constructed
slices["usd_amount"]   = pd.to_numeric(slices["usd_amount"],   errors="coerce").fillna(0)
slices["local_amount"] = pd.to_numeric(slices["local_amount"], errors="coerce").fillna(0)

# Diagnostic — shown only when $0 to help debug
if slices.empty or slices["usd_amount"].sum() == 0:
    st.warning(
        f"⚠️ Revenue data loaded ({len(df_rev_raw)} charge rows) but produced "
        f"{'no slices' if slices.empty else '$0 in slices'}. "
        f"Columns in file: {list(df_rev_raw.columns)}. "
        f"Sample rev_start: {df_rev_raw.get('rev_start', pd.Series()).dropna().head(3).tolist()}. "
        f"Sample recognizable_amount: {df_rev_raw.get('recognizable_amount', pd.Series()).head(3).tolist()}"
    )

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

# ── Build monthly totals for trend calculations ───────────────────────────────
# Use full-month slices (not pro-rated) for MoM and run rate
_monthly_totals = (slices[slices["period"].isin(ytd_months)]
                   .groupby("period")["usd_amount"].sum()
                   .sort_index())

# MoM growth: compare last two complete months
_complete_months = [m for m in sorted(_monthly_totals.index) if m < this_month]
if len(_complete_months) >= 2:
    _prev_mo  = _complete_months[-2]
    _curr_mo  = _complete_months[-1]
    _prev_rev = _monthly_totals.get(_prev_mo, 0)
    _curr_rev = _monthly_totals.get(_curr_mo, 0)
    _mom_pct  = ((_curr_rev - _prev_rev) / _prev_rev * 100) if _prev_rev > 0 else None
    _mom_label = pd.Timestamp(_curr_mo + "-01").strftime("%b")
    _prev_label = pd.Timestamp(_prev_mo + "-01").strftime("%b")
elif len(_complete_months) == 1:
    _mom_pct = None; _curr_rev = _monthly_totals.iloc[0]; _mom_label = ""
else:
    _mom_pct = None; _curr_rev = 0; _mom_label = ""

# Run rate: avg of complete months × 12
_avg_monthly = _monthly_totals[_monthly_totals.index < this_month].mean() if len(_complete_months) >= 1 else 0
_run_rate    = _avg_monthly * 12

ytd_total  = ytd_df["usd_amount"].sum()
qtd_total  = qtd_df["usd_amount"].sum()
mtd_total  = slices_mtd["usd_amount"].sum()
full_month = slices[slices["period"]==this_month]["usd_amount"].sum()

c1,c2,c3,c4,c5,c6 = st.columns(6)
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
with c5:
    if _mom_pct is not None:
        _arrow = "↑" if _mom_pct >= 0 else "↓"
        _col   = "#27AE60" if _mom_pct >= 0 else "#E74C3C"
        st.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-val" style="color:{_col}">{_arrow} {abs(_mom_pct):.1f}%</div>'
            f'<div class="metric-sub">MoM Growth</div>'
            f'<div class="metric-lbl">{_prev_label} → {_mom_label}</div></div>',
            unsafe_allow_html=True)
    else:
        st.markdown(
            '<div class="metric-card"><div class="metric-val">—</div>'
            '<div class="metric-sub">MoM Growth</div>'
            '<div class="metric-lbl">Need 2+ months</div></div>',
            unsafe_allow_html=True)
with c6:
    st.markdown(
        f'<div class="metric-card"><div class="metric-val">{_fmt(_run_rate)}</div>'
        f'<div class="metric-sub">Run Rate (ARR)</div>'
        f'<div class="metric-lbl">Avg monthly × 12 ({len(_complete_months)} mo)</div></div>',
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

_rt_ytd = _region_table(ytd_df,     "YTD")
_rt_qtd = _region_table(qtd_df,     "QTD")
_rt_mtd = _region_table(slices_mtd, "MTD")
_rt = (_rt_ytd.merge(_rt_qtd, on="Region", how="outer")
              .merge(_rt_mtd, on="Region", how="outer")
              .fillna("$0")) if not _rt_ytd.empty else pd.DataFrame()

if not _rt.empty:
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

_pt_ytd = _product_table(ytd_df,     "YTD")
_pt_qtd = _product_table(qtd_df,     "QTD")
_pt_mtd = _product_table(slices_mtd, "MTD")
_pt = (_pt_ytd.merge(_pt_qtd, on="Product", how="outer")
              .merge(_pt_mtd, on="Product", how="outer")
              .fillna("$0")) if not _pt_ytd.empty else pd.DataFrame()

if not _pt.empty:
    st.dataframe(_pt, use_container_width=True, hide_index=True)
else:
    st.info("No product data available.")

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION — Monthly Breakdown by Region & Product
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Monthly Breakdown (USD)</div>', unsafe_allow_html=True)

# ── Build quarterly columnar pivot ───────────────────────────────────────────
# Layout: Jan | Feb | Mar | Q1 | Apr | May | Jun | Q2 | ... | YTD
# Covers all years present in data

def _pivot_quarterly(df, row_dim):
    """Build management accounts layout: months interleaved with quarterly subtotals + YTD."""
    if df.empty or row_dim not in df.columns:
        return pd.DataFrame()

    # All unique periods in data
    all_periods = sorted(df["period"].unique())
    all_years   = sorted({p[:4] for p in all_periods})

    # Base pivot: rows = row_dim, columns = period
    base = (df.groupby([row_dim, "period"])["usd_amount"]
              .sum()
              .unstack(fill_value=0))

    # Build ordered column list with quarter subtotals
    ordered_cols  = []   # (col_label, source_type, source_key)
    # source_type: "month" | "quarter" | "ytd"

    for year in all_years:
        year_periods = [p for p in all_periods if p.startswith(year)]
        year_months  = sorted(year_periods)

        for q in range(1, 5):
            q_month_nums = [(q-1)*3+1, (q-1)*3+2, (q-1)*3+3]
            q_periods    = [f"{year}-{m:02d}" for m in q_month_nums]
            q_present    = [p for p in q_periods if p in year_months]

            if not q_present:
                continue

            # Add individual months
            for p in q_present:
                mo_label = pd.Timestamp(p + "-01").strftime("%b")
                # Add year suffix only if data spans multiple years
                if len(all_years) > 1:
                    mo_label += f" '{year[2:]}"
                ordered_cols.append((mo_label, "month", p))

            # Add quarter subtotal if any months present
            ordered_cols.append((f"Q{q}" + (f" '{year[2:]}" if len(all_years) > 1 else ""), "quarter", q_periods))

        # Add YTD after December (or last month of year)
        if year == all_years[-1]:
            ordered_cols.append(("YTD", "ytd", year_periods))
        else:
            ordered_cols.append((f"FY{year[2:]}", "ytd", year_periods))

    # Build the output DataFrame column by column
    rows_index = base.index.tolist()
    result = {row_dim: list(rows_index) + ["Total"]}

    for col_label, ctype, key in ordered_cols:
        col_vals = []
        if ctype == "month":
            vals = base[key] if key in base.columns else pd.Series(0, index=base.index)
        else:
            # Sum across list of periods
            present = [p for p in key if p in base.columns]
            vals = base[present].sum(axis=1) if present else pd.Series(0, index=base.index)

        for idx in rows_index:
            col_vals.append(float(vals.get(idx, 0)))
        col_vals.append(sum(col_vals))  # Total row
        result[col_label] = [_fmt(v) for v in col_vals]

    return pd.DataFrame(result)

_piv_region  = _pivot_quarterly(slices, "region")
_piv_product = _pivot_quarterly(slices, "product")

if not _piv_region.empty:
    st.markdown("**By Region**")
    st.dataframe(_piv_region, use_container_width=True, hide_index=True)

if not _piv_product.empty:
    st.markdown("**By Product**")
    st.dataframe(_piv_product, use_container_width=True, hide_index=True)

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION — Trend Analysis
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Trend Analysis</div>', unsafe_allow_html=True)

# Build full monthly table with MoM delta and % change
_trend_full = (_monthly_totals.reset_index()
               if not _monthly_totals.empty else pd.DataFrame(columns=["period","usd_amount"]))
_trend_full.columns = ["Period", "Revenue (USD)"]
_trend_full["Revenue (USD)"] = pd.to_numeric(_trend_full["Revenue (USD)"], errors="coerce").fillna(0)
_trend_full["MoM Change"]    = _trend_full["Revenue (USD)"].diff()
_trend_full["MoM %"]         = _trend_full["Revenue (USD)"].pct_change() * 100
_trend_full["Cumulative YTD"] = _trend_full["Revenue (USD)"].cumsum()

# Format for display
_trend_disp = _trend_full.copy()
_trend_disp["Period"]         = _trend_disp["Period"].apply(
    lambda m: pd.Timestamp(m + "-01").strftime("%b %Y"))
_trend_disp["Revenue (USD)"]  = _trend_disp["Revenue (USD)"].apply(_fmt)
_trend_disp["MoM Change"]     = _trend_disp["MoM Change"].apply(
    lambda v: ("↑ " if v >= 0 else "↓ ") + _fmt(abs(v)) if pd.notna(v) else "—")
_trend_disp["MoM %"]          = _trend_disp["MoM %"].apply(
    lambda v: f"{'↑' if v >= 0 else '↓'} {abs(v):.1f}%" if pd.notna(v) else "—")
_trend_disp["Cumulative YTD"] = _trend_disp["Cumulative YTD"].apply(_fmt)

st.dataframe(_trend_disp, use_container_width=True, hide_index=True)

# Region trend table
if "region" in slices.columns and len(_complete_months) >= 2:
    st.markdown("**MoM by Region**")
    _rgn_trend = (slices[slices["period"].isin(_complete_months)]
                  .groupby(["region","period"])["usd_amount"].sum()
                  .unstack(fill_value=0)
                  .reindex(columns=_complete_months, fill_value=0))
    # Add MoM % for last two months
    if len(_complete_months) >= 2:
        _rgn_trend["MoM %"] = (
            (_rgn_trend[_complete_months[-1]] - _rgn_trend[_complete_months[-2]])
            / _rgn_trend[_complete_months[-2]].replace(0, float("nan")) * 100
        ).apply(lambda v: f"{'↑' if v >= 0 else '↓'} {abs(v):.1f}%" if pd.notna(v) else "—")
    _rgn_trend.columns = [
        pd.Timestamp(m + "-01").strftime("%b %Y") if m != "MoM %" else "MoM %"
        for m in _rgn_trend.columns
    ]
    # Format numeric columns
    for _mc in _rgn_trend.columns:
        if _mc != "MoM %":
            _rgn_trend[_mc] = pd.to_numeric(_rgn_trend[_mc], errors="coerce").apply(_fmt)
    _rgn_trend.index.name = "Region"
    st.dataframe(_rgn_trend.reset_index(), use_container_width=True, hide_index=True)

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — Monthly trend
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Monthly Trend (USD)</div>',unsafe_allow_html=True)

_trend = slices.groupby("period")["usd_amount"].sum().reset_index()
_trend.columns = ["Period", "Revenue (USD)"]
_trend = _trend.sort_values("Period")
_trend["Revenue (USD)"] = pd.to_numeric(_trend["Revenue (USD)"], errors="coerce").fillna(0).round(2)
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
for _nc in ("usd_amount", "local_amount"):
    if _nc in _detail.columns:
        _detail[_nc] = pd.to_numeric(_detail[_nc], errors="coerce").fillna(0).round(2)

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
    if not _rt.empty:          _rt.to_excel(_xl,          sheet_name="By Region",         index=False)
    if not _pt.empty:          _pt.to_excel(_xl,          sheet_name="By Product",         index=False)
    if not _piv_region.empty:  _piv_region.to_excel(_xl,  sheet_name="Region by Month",    index=False)
    if not _piv_product.empty: _piv_product.to_excel(_xl, sheet_name="Product by Month",   index=False)
    if not _trend_disp.empty:  _trend_disp.to_excel(_xl,  sheet_name="Trend Analysis",     index=False)
    _detail.to_excel(_xl, sheet_name="Project Detail", index=False)

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
