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
try:
    from shared.loaders import join_tm_to_ns, calc_tm_monthly_actuals, get_billing_mismatches, get_unmatched_sow
except ImportError:
    def join_tm_to_ns(df_sow, df_ns, df_drs=None):
        df_sow["ns_project"] = None; df_sow["ns_hours_worked"] = 0.0
        df_sow["ns_revenue_to_date"] = 0.0; df_sow["match_source"] = "No NS data"
        return df_sow
    def calc_tm_monthly_actuals(df_ns, df_sow): return pd.DataFrame()
    def get_billing_mismatches(df_ns): return pd.DataFrame()
    def get_unmatched_sow(df_tm): return pd.DataFrame()
from shared.config import CURRENCY_REGION_MAP, FX_RATES_TO_USD

st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
<style>
    html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
    h1,h2,h3,h4,p,div,label,button { font-family: 'Manrope', sans-serif !important; }
    .brief-header  { font-size: 24px; font-weight: 700; color: inherit; margin-bottom: 4px; }
    .brief-sub     { font-size: 13px; margin-bottom: 20px; opacity: 0.6; }
    .section-label { font-size: 11px; font-weight: 700; text-transform: uppercase;
                     letter-spacing: 0.8px; color: #4472C4; margin-bottom: 8px; }
    .metric-card   { background: transparent; border: 1px solid rgba(128,128,128,0.2);
                     border-radius: 8px; padding: 16px 20px; margin-bottom: 12px; }
    .metric-val    { font-size: 26px; font-weight: 700; color: inherit; }
    .metric-lbl    { font-size: 12px; opacity: 0.6; margin-top: 2px; }
    .divider { border: none; border-top: 1px solid rgba(128,128,128,0.2); margin: 20px 0; }
    .fx-note { font-size: 11px; opacity: 0.5; font-style: italic; }
</style>
""", unsafe_allow_html=True)

# ── Identity / access ─────────────────────────────────────────────────────────
from shared.constants import get_role
_role = get_role(st.session_state.get("consultant_name",""))
if _role not in ("manager", "manager_only", "reporting_only"):
    st.warning("This page is available to managers and reporting users only.")
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
    st.info("ℹ️ Upload the NS FF Revenue Charges export in the sidebar to see the Fixed Fee section.")
    # Don't stop — T&M section still loads from NS Time Detail alone

# ── Build monthly slices (FF only — skipped if FF Revenue Charges not loaded) ──
@st.cache_data(show_spinner=False)
def _get_slices(df_hash):
    return calc_monthly_slices(st.session_state["df_revenue"])

if df_rev_raw is not None:
    slices = calc_monthly_slices(df_rev_raw)
    slices["usd_amount"]   = pd.to_numeric(slices["usd_amount"],   errors="coerce").fillna(0)
    slices["local_amount"] = pd.to_numeric(slices["local_amount"], errors="coerce").fillna(0)

    if slices.empty or slices["usd_amount"].sum() == 0:
        st.warning(
            f"⚠️ Revenue data loaded ({len(df_rev_raw)} charge rows) but produced "
            f"{'no slices' if slices.empty else '$0 in slices'}. "
            f"Columns in file: {list(df_rev_raw.columns)}. "
            f"Sample rev_start: {df_rev_raw.get('rev_start', pd.Series()).dropna().head(3).tolist()}. "
            f"Sample recognizable_amount: {df_rev_raw.get('recognizable_amount', pd.Series()).head(3).tolist()}"
        )

    # ── Join DRS for project name + consultant ────────────────────────────────
    if df_drs is not None and "project_id" in df_drs.columns:
        drs_lookup = df_drs[["project_id","project_name","project_manager","phase"]].copy()
        drs_lookup["project_id"] = drs_lookup["project_id"].astype(str).str.strip().str.split(".").str[0]
        # Preserve project_name already in slices (from FF loader) before merge
        _existing_pname = slices.get("project_name", pd.Series(dtype=str)).copy() if "project_name" in slices.columns else None
        slices = slices.merge(drs_lookup, on="project_id", how="left")
        # Restore FF-loaded project_name where DRS didn't fill it
        if _existing_pname is not None and "project_name" in slices.columns:
            slices["project_name"] = slices["project_name"].fillna("").where(
                slices["project_name"].fillna("") != "",
                _existing_pname.reindex(slices.index).fillna("")
            )
    else:
        # Don't overwrite project_name if already populated from FF loader
        if "project_name" not in slices.columns or slices["project_name"].fillna("").eq("").all():
            slices["project_name"] = slices["project_id"]
        slices["project_manager"] = ""
        slices["phase"]           = ""
else:
    slices = pd.DataFrame(columns=["period","project_id","project_name","project_manager",
                                    "phase","usd_amount","local_amount"])

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

slices_mtd = slices[slices["period"] == this_month].copy() if not slices.empty else pd.DataFrame(columns=["period","usd_amount"])
slices_mtd["usd_amount"] = slices_mtd["usd_amount"] * mtd_scale if not slices_mtd.empty else 0

ytd_df  = slices[slices["period"].isin(ytd_months)].copy() if not slices.empty else pd.DataFrame(columns=["period","usd_amount"])
ytd_df  = ytd_df[ytd_df["period"] != this_month]
ytd_df  = pd.concat([ytd_df, slices_mtd], ignore_index=True)

qtd_df  = slices[slices["period"].isin(q_months)].copy() if not slices.empty else pd.DataFrame(columns=["period","usd_amount"])
qtd_df  = qtd_df[qtd_df["period"] != this_month]
qtd_df  = pd.concat([qtd_df, slices_mtd], ignore_index=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<div class="brief-header">Revenue Report</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="brief-sub">All figures in USD &nbsp;·&nbsp; {today.strftime("%A, %B %-d %Y")} &nbsp;·&nbsp; '
    f'MTD pro-rated ({days_elapsed}/{days_in_month} days) &nbsp;·&nbsp; '
    f'<span class="fx-note">FX rates: monthly averages — update in shared/config.py</span></div>',
    unsafe_allow_html=True)
st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── Pre-compute T&M metrics here so they can appear alongside FF at the top ───
df_tm_sow = st.session_state.get("df_tm_sow")
df_ns     = st.session_state.get("df_ns")
df_tm     = None

if df_tm_sow is not None:
    df_tm = join_tm_to_ns(df_tm_sow, df_ns, df_drs)
    # Deduplicate columns in case join produced duplicates
    df_tm = df_tm.loc[:, ~df_tm.columns.duplicated()]
    for _nc in ("sow_amount_usd","sow_hours","ns_revenue_to_date","ns_hours_worked","ns_rate"):
        if _nc in df_tm.columns:
            df_tm[_nc] = pd.to_numeric(df_tm[_nc], errors="coerce").fillna(0)
    tm_contracted   = float(df_tm["sow_amount_usd"].sum())
    tm_worked       = float(df_tm["ns_revenue_to_date"].sum()) if "ns_revenue_to_date" in df_tm.columns else 0.0
    tm_hours_sold   = float(df_tm["sow_hours"].sum())
    tm_hours_worked = float(df_tm["ns_hours_worked"].sum()) if "ns_hours_worked" in df_tm.columns else 0.0
    tm_matched      = int(df_tm["ns_project"].notna().sum()) if "ns_project" in df_tm.columns else 0
    tm_unmatched    = len(df_tm) - tm_matched
    _burn           = (tm_worked / tm_contracted * 100) if tm_contracted > 0 else 0.0
    _rem            = tm_contracted - tm_worked

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Top-line metric cards (FF + T&M)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Total Revenue (FF + T&M)</div>',unsafe_allow_html=True)

# ── Pre-compute T&M actuals here so they feed top-line totals ────────────────
# (full computation happens later; this gives us monthly totals for the bubbles)
_tm_monthly_early = pd.Series(dtype=float)
if df_ns is not None:
    try:
        _tma_early = calc_tm_monthly_actuals(df_ns, df_tm_sow)
        if not _tma_early.empty and "period" in _tma_early.columns:
            _tm_monthly_early = (
                _tma_early[_tma_early["period"].isin(ytd_months)]
                .groupby("period")["revenue_usd"].sum()
            )
    except Exception:
        pass

# ── Build monthly totals — FF + T&M combined ─────────────────────────────────
_ff_monthly = (slices[slices["period"].isin(ytd_months)]
               .groupby("period")["usd_amount"].sum()
               .sort_index()) if not slices.empty else pd.Series(dtype=float)

_monthly_totals = _ff_monthly.add(_tm_monthly_early, fill_value=0).sort_index()

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

# FF component
_ff_ytd   = ytd_df["usd_amount"].sum()   if not ytd_df.empty   else 0.0
_ff_qtd   = qtd_df["usd_amount"].sum()   if not qtd_df.empty   else 0.0
_ff_mtd   = slices_mtd["usd_amount"].sum() if not slices_mtd.empty else 0.0
_ff_full  = slices[slices["period"]==this_month]["usd_amount"].sum() if not slices.empty else 0.0

# T&M component (from early pre-computation)
_tm_ytd   = float(_tm_monthly_early[_tm_monthly_early.index.isin(ytd_months)].sum())
_tm_qtd   = float(_tm_monthly_early[_tm_monthly_early.index.isin(q_months)].sum())
_tm_mtd   = float(_tm_monthly_early.get(this_month, 0) * mtd_scale)
_tm_full  = float(_tm_monthly_early.get(this_month, 0))

ytd_total  = _ff_ytd  + _tm_ytd
qtd_total  = _ff_qtd  + _tm_qtd
mtd_total  = _ff_mtd  + _tm_mtd
full_month = _ff_full + _tm_full

# ── Fixed Fee metric cards ────────────────────────────────────────────────────
# Total Revenue cards — combined FF + T&M
_has_tm_data = df_ns is not None and not _tm_monthly_early.empty
c1,c2,c3,c4,c5,c6 = st.columns(6)
with c1:
    st.markdown(
        f'<div class="metric-card"><div class="metric-val">{_fmt(ytd_total)}</div>'
        f'<div class="metric-lbl">YTD · Jan 1 – {today.strftime("%-d %b %Y")}</div></div>',
        unsafe_allow_html=True)
with c2:
    st.markdown(
        f'<div class="metric-card"><div class="metric-val">{_fmt(qtd_total)}</div>'
        f'<div class="metric-lbl">QTD · Q{this_q} {today.year}</div></div>',
        unsafe_allow_html=True)
with c3:
    st.markdown(
        f'<div class="metric-card"><div class="metric-val">{_fmt(mtd_total)}</div>'
        f'<div class="metric-lbl">MTD · {today.strftime("%B")} to date</div></div>',
        unsafe_allow_html=True)
with c4:
    st.markdown(
        f'<div class="metric-card"><div class="metric-val">{_fmt(full_month)}</div>'
        f'<div class="metric-lbl">Full Month · {today.strftime("%B")} forecast</div></div>',
        unsafe_allow_html=True)
with c5:
    if _mom_pct is not None:
        _arrow = "↑" if _mom_pct >= 0 else "↓"
        _col   = "#27AE60" if _mom_pct >= 0 else "#E74C3C"
        st.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-val" style="color:{_col}">{_arrow} {abs(_mom_pct):.1f}%</div>'
            f'<div class="metric-lbl">MoM Growth · {_prev_label} → {_mom_label}</div></div>',
            unsafe_allow_html=True)
    else:
        st.markdown(
            '<div class="metric-card"><div class="metric-val">—</div>'
            '<div class="metric-lbl">MoM Growth · need 2+ months</div></div>',
            unsafe_allow_html=True)
with c6:
    st.markdown(
        f'<div class="metric-card"><div class="metric-val">{_fmt(_run_rate)}</div>'
        f'<div class="metric-lbl">Run Rate (ARR) · avg × 12 ({len(_complete_months)} mo)</div></div>',
        unsafe_allow_html=True)

# ── T&M metric cards ──────────────────────────────────────────────────────────
if df_tm is not None:
    st.markdown('<div class="section-label" style="margin-top:4px">T&amp;M Pipeline (from SFDC SOW)</div>', unsafe_allow_html=True)
    _bc = "#E74C3C" if _burn > 90 else ("#F39C12" if _burn > 70 else "inherit")
    _avg_rate = float(df_tm["sow_rate_usd"].replace(0, float("nan")).mean()) if "sow_rate_usd" in df_tm.columns else 0.0
    _avg_rate = 0.0 if pd.isna(_avg_rate) else _avg_rate
    t1,t2,t3,t4,t5,t6 = st.columns(6)
    with t1:
        st.markdown(
            f'<div class="metric-card"><div class="metric-val">{_fmt(tm_contracted)}</div>'
            f'<div class="metric-lbl">Contracted YTD · {len(df_tm)} opportunities</div></div>',
            unsafe_allow_html=True)
    with t2:
        st.markdown(
            f'<div class="metric-card"><div class="metric-val">{_fmt(tm_worked)}</div>'
            f'<div class="metric-lbl">SOW-matched Earned · {tm_matched} opps matched to NS</div></div>',
            unsafe_allow_html=True)
    with t3:
        st.markdown(
            f'<div class="metric-card"><div class="metric-val" style="color:{_bc}">{_burn:.1f}%</div>'
            f'<div class="metric-lbl">Burn Rate · {tm_hours_worked:,.0f}h of {tm_hours_sold:,.0f}h sold</div></div>',
            unsafe_allow_html=True)
    with t4:
        st.markdown(
            f'<div class="metric-card"><div class="metric-val">{_fmt(_rem)}</div>'
            f'<div class="metric-lbl">Remaining Backlog · contracted not yet earned</div></div>',
            unsafe_allow_html=True)
    with t5:
        st.markdown(
            f'<div class="metric-card"><div class="metric-val">{tm_hours_sold:,.0f}h</div>'
            f'<div class="metric-lbl">Total Hours Sold · across all SOWs</div></div>',
            unsafe_allow_html=True)
    with t6:
        st.markdown(
            f'<div class="metric-card"><div class="metric-val">${_avg_rate:,.0f}</div>'
            f'<div class="metric-lbl">Avg Rate (USD) · PS SOW Rate converted</div></div>',
            unsafe_allow_html=True)
    if tm_unmatched > 0:
        _ns_matched  = int((df_tm.get("match_source","") == "NS match").sum()) if "match_source" in df_tm.columns else tm_matched
        _drs_matched = int((df_tm.get("match_source","") == "DRS match (no hours yet)").sum()) if "match_source" in df_tm.columns else 0
        st.caption(f"⚠️ {tm_unmatched} of {len(df_tm)} opportunities unmatched — "
                   f"{_ns_matched} NS matches, {_drs_matched} DRS-only matches (no hours yet). "
                   f"See Unmatched Opportunities expander below.")

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — FF Summary by Region & Product
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Revenue Summary by Region &amp; Product</div>', unsafe_allow_html=True)

def _ff_summary_table(df_ytd, df_qtd, df_mtd, df_full_mo, dim):
    """Build FF summary table matching T&M structure."""
    if df_ytd.empty or dim not in df_ytd.columns:
        return pd.DataFrame()

    def _agg(df, col_name):
        if df.empty or dim not in df.columns or "usd_amount" not in df.columns:
            return pd.Series(dtype=float, name=col_name)
        return df.groupby(dim)["usd_amount"].sum().rename(col_name)

    ytd_s  = _agg(df_ytd,     "YTD")
    qtd_s  = _agg(df_qtd,     "QTD")
    mtd_s  = _agg(df_mtd,     "MTD")
    full_s = _agg(df_full_mo, "Full Month")

    # Use YTD index as the base (most complete), fill missing with 0
    idx = ytd_s.index
    result = pd.DataFrame({
        "MTD":        mtd_s.reindex(idx, fill_value=0),
        "Full Month": full_s.reindex(idx, fill_value=0),
        "QTD":        qtd_s.reindex(idx, fill_value=0),
        "YTD":        ytd_s.reindex(idx, fill_value=0),
    })

    # Count projects per dim
    if "project_id" in df_ytd.columns:
        proj_counts = df_ytd.groupby(dim)["project_id"].nunique().rename("# Projects")
        result = result.join(proj_counts, how="left").fillna(0)

    # Total row
    total = result.sum(numeric_only=True)
    total.name = "Total"
    result = pd.concat([result, total.to_frame().T])

    # Format
    for c in ["MTD", "Full Month", "QTD", "YTD"]:
        if c in result.columns:
            result[c] = result[c].apply(_fmt)
    if "# Projects" in result.columns:
        result["# Projects"] = result["# Projects"].apply(lambda v: f"{int(v):,}")

    result.index.name = dim.capitalize()
    return result.reset_index()

_ff_full_mo_df = slices[slices["period"] == this_month].copy()
_ff_sum_rgn  = _ff_summary_table(ytd_df, qtd_df, slices_mtd, _ff_full_mo_df, "region")
_ff_sum_prod = _ff_summary_table(ytd_df, qtd_df, slices_mtd, _ff_full_mo_df, "product")

_tab1, _tab2 = st.tabs(["By Region", "By Product"])
with _tab1:
    if not _ff_sum_rgn.empty:
        st.dataframe(_ff_sum_rgn.rename(columns={"region":"Region"}),
                     use_container_width=True, hide_index=True)
    else:
        st.info("No region data available.")
with _tab2:
    if not _ff_sum_prod.empty:
        st.dataframe(_ff_sum_prod.rename(columns={"product":"Product"}),
                     use_container_width=True, hide_index=True)
    else:
        st.info("No product data available.")

# Keep for Excel export
_rt = _ff_sum_rgn.rename(columns={"region":"Region"}) if not _ff_sum_rgn.empty else pd.DataFrame()
_pt = _ff_sum_prod.rename(columns={"product":"Product"}) if not _ff_sum_prod.empty else pd.DataFrame()

st.markdown('<hr class="divider">',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION — T&M Revenue Summary
# ══════════════════════════════════════════════════════════════════════════════
if df_tm is not None:
    st.markdown('<div class="section-label">T&amp;M Revenue Summary</div>', unsafe_allow_html=True)

    def _tm_summary_table(df, dim):
        """Build T&M summary table grouped by dim (region or product)."""
        if df.empty or dim not in df.columns:
            return pd.DataFrame()
        grp = (df.groupby(dim)
               .agg(
                   contracted  = ("sow_amount_usd",      "sum"),
                   earned      = ("ns_revenue_to_date",  "sum"),
                   hours_sold  = ("sow_hours",            "sum"),
                   hours_worked= ("ns_hours_worked",      "sum"),
                   avg_rate    = ("sow_rate_usd",         lambda x: x.replace(0,float("nan")).mean()),
                   n_opps      = ("opportunity_name",     "count"),
               )
               .reset_index())
        # Total row
        tot = grp[["contracted","earned","hours_sold","hours_worked","n_opps"]].sum()
        tot["avg_rate"] = df["sow_rate_usd"].replace(0,float("nan")).mean()
        tot[dim] = "Total"
        grp = pd.concat([grp, tot.to_frame().T], ignore_index=True)

        grp["Burn %"]      = (grp["earned"] / grp["contracted"] * 100).apply(
                               lambda v: f"{v:.1f}%" if pd.notna(v) and v > 0 else "—")
        grp["Remaining"]   = (grp["contracted"] - grp["earned"]).apply(_fmt)
        grp["Contracted"]  = grp["contracted"].apply(_fmt)
        grp["Earned"]      = grp["earned"].apply(_fmt)
        grp["Hours Sold"]  = grp["hours_sold"].apply(lambda v: f"{v:,.0f}h")
        grp["Avg Rate"]    = grp["avg_rate"].apply(lambda v: f"${v:,.0f}" if pd.notna(v) else "—")
        grp["# Opps"]      = grp["n_opps"].apply(lambda v: f"{int(v):,}")
        col_label = dim.capitalize() if dim != "region" else "Region"
        return grp[[dim,"Contracted","Earned","Remaining","Burn %","Hours Sold","Avg Rate","# Opps"]].rename(columns={dim: col_label})

    _tm_sum_rgn  = _tm_summary_table(df_tm, "region")
    _tm_sum_prod = _tm_summary_table(df_tm, "product")

    tab1, tab2 = st.tabs(["By Region", "By Product"])
    with tab1:
        if not _tm_sum_rgn.empty:
            st.dataframe(_tm_sum_rgn, use_container_width=True, hide_index=True)
    with tab2:
        if not _tm_sum_prod.empty:
            st.dataframe(_tm_sum_prod, use_container_width=True, hide_index=True)

    with st.expander(f"Opportunity Detail ({len(df_tm)} rows)", expanded=False):
        # Rate alignment summary — show before detail table
        if "rate_alignment" in df_tm.columns:
            _mismatches = df_tm[df_tm["rate_alignment"].str.startswith("⚠️", na=False)]
            if not _mismatches.empty:
                st.warning(f"⚠️ {len(_mismatches)} matched opportunities have a rate mismatch between NS and SFDC SOW — see Rate Alignment column below.")

        _tm_detail_cols = ["account_name","opportunity_name","opportunity_owner",
                           "product","region","close_date","sow_hours",
                           "sow_amount_usd","sow_rate_usd","ns_project",
                           "ns_rate","ns_hours_worked","ns_revenue_to_date",
                           "rate_alignment","match_source"]
        _tm_detail = df_tm[[c for c in _tm_detail_cols if c in df_tm.columns]].copy()
        if "close_date" in _tm_detail.columns:
            _tm_detail["close_date"] = _tm_detail["close_date"].dt.strftime("%-d %b %Y")
        st.dataframe(_tm_detail, use_container_width=True, hide_index=True,
                     column_config={
                         "account_name":       st.column_config.TextColumn("Account",        width="medium"),
                         "opportunity_name":   st.column_config.TextColumn("Opportunity",    width="large"),
                         "opportunity_owner":  st.column_config.TextColumn("Owner",          width="medium"),
                         "product":            st.column_config.TextColumn("Product",        width="small"),
                         "region":             st.column_config.TextColumn("Region",         width="small"),
                         "close_date":         st.column_config.TextColumn("Close Date",     width="small"),
                         "sow_hours":          st.column_config.NumberColumn("SOW Hrs",      width="small"),
                         "sow_amount_usd":     st.column_config.NumberColumn("SOW Amt USD",  width="small", format="$%.0f"),
                         "sow_rate_usd":       st.column_config.NumberColumn("SOW Rate",     width="small", format="$%.0f"),
                         "ns_project":         st.column_config.TextColumn("NS Match",       width="medium"),
                         "ns_rate":            st.column_config.NumberColumn("NS Rate",      width="small", format="$%.0f"),
                         "ns_hours_worked":    st.column_config.NumberColumn("Hrs Worked",   width="small"),
                         "ns_revenue_to_date": st.column_config.NumberColumn("Rev Earned",   width="small", format="$%.0f"),
                         "rate_alignment":     st.column_config.TextColumn("Rate Alignment", width="medium"),
                         "match_source":       st.column_config.TextColumn("Match Source",   width="medium"),
                     })

    # Unmatched report — for vetting match quality
    _unmatched = get_unmatched_sow(df_tm)
    if not _unmatched.empty:
        with st.expander(f"⚠️ Unmatched Opportunities ({len(_unmatched)} rows) — review match quality", expanded=False):
            st.caption("These SOW opportunities were not matched to any NS project or DRS record. "
                       "Check account name spelling vs NS/DRS, or these may be pre-2026 SOWs with no active project.")
            st.dataframe(_unmatched, use_container_width=True, hide_index=True,
                         column_config={
                             "account_name":      st.column_config.TextColumn("Account",      width="medium"),
                             "opportunity_name":  st.column_config.TextColumn("Opportunity",  width="large"),
                             "opportunity_owner": st.column_config.TextColumn("Owner",        width="medium"),
                             "product":           st.column_config.TextColumn("Product",      width="small"),
                             "region":            st.column_config.TextColumn("Region",       width="small"),
                             "sow_hours":         st.column_config.NumberColumn("SOW Hrs",    width="small"),
                             "sow_amount_usd":    st.column_config.NumberColumn("SOW Amt",    width="small", format="$%.0f"),
                             "sow_rate_usd":      st.column_config.NumberColumn("Rate",       width="small", format="$%.0f"),
                             "match_source":      st.column_config.TextColumn("Status",       width="small"),
                         })
elif df_tm_sow is None:
    st.info("Upload the SFDC T&M SOW export in the sidebar to see T&M breakdown.")

st.markdown('<div class="section-label">T&amp;M Monthly Actuals (from NS Time Detail)</div>',
            unsafe_allow_html=True)

# Currency breakdown table injected after _tm_actuals is built — placeholder rendered below

df_ns_session = st.session_state.get("df_ns")
_tm_actuals   = pd.DataFrame()

if df_ns_session is None:
    st.info("Upload NS Time Detail in the sidebar to see monthly T&M actuals.")
else:
    _tm_actuals = calc_tm_monthly_actuals(df_ns_session, df_tm_sow)


    if _tm_actuals.empty:
        st.info("No T&M time entries found in the NS file.")
    else:
        _matched_pct  = (_tm_actuals["rate_usd"] > 0).mean() * 100
        _total_tm_rev = float(_tm_actuals["revenue_usd"].sum())
        _total_tm_hrs = float(_tm_actuals["hours"].sum())
        _unrated      = int((_tm_actuals["rate_usd"] == 0).sum())
        _no_sfdc      = int((_tm_actuals.get("rate_source", pd.Series()) == "No SFDC Opp").sum())

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(
                f'<div class="metric-card"><div class="metric-val">{_fmt(_total_tm_rev)}</div>'
                f'<div class="metric-lbl">NS-driven Actual Revenue · hours logged × matched rate</div></div>',
                unsafe_allow_html=True)
        with c2:
            st.markdown(
                f'<div class="metric-card"><div class="metric-val">{_total_tm_hrs:,.1f}h</div>'
                f'<div class="metric-lbl">T&M Hours Worked · all projects</div></div>',
                unsafe_allow_html=True)
        with c3:
            _rc = "#F39C12" if _matched_pct < 80 else "inherit"
            st.markdown(
                f'<div class="metric-card"><div class="metric-val" style="color:{_rc}">{_matched_pct:.0f}%</div>'
                f'<div class="metric-lbl">Rate Match · {_unrated} rows at $0 rate</div></div>',
                unsafe_allow_html=True)

        if _no_sfdc > 0:
            st.caption(
                f"ℹ️ {_no_sfdc} project-month rows are T&M in NS but have no matching SFDC opportunity — "
                f"likely pre-2026 SOWs or projects created without an Opp. "
                f"Add a Rate column to your NS Time Detail export to calculate revenue for these rows "
                f"without needing a SFDC match."
            )
        elif _unrated > 0:
            st.caption(f"⚠️ {_unrated} project-month rows have no matched rate — revenue shown as $0.")

        def _tm_actuals_pivot(df, dim):
            if df.empty or dim not in df.columns: return pd.DataFrame()
            all_periods = sorted(df["period"].unique())
            all_years   = sorted({p[:4] for p in all_periods})
            base = (df.groupby([dim,"period"])["revenue_usd"]
                      .sum().unstack(fill_value=0))
            ordered_cols = []
            for year in all_years:
                year_periods = [p for p in all_periods if p.startswith(year)]
                for q in range(1, 5):
                    q_ps = [f"{year}-{m:02d}" for m in range((q-1)*3+1,(q-1)*3+4)]
                    q_present = [p for p in q_ps if p in year_periods]
                    if not q_present: continue
                    for p in q_present:
                        ml = pd.Timestamp(p+"-01").strftime("%b")
                        if len(all_years)>1: ml += f" '{year[2:]}"
                        ordered_cols.append((ml,"month",p))
                    ordered_cols.append((f"Q{q}"+(f" '{year[2:]}" if len(all_years)>1 else ""),"quarter",q_ps))
                ytd_label = "YTD" if year==all_years[-1] else f"FY{year[2:]}"
                ordered_cols.append((ytd_label,"ytd",year_periods))
            rows_index = base.index.tolist()
            result = {dim: list(rows_index)+["Total"]}
            for col_label, ctype, key in ordered_cols:
                if ctype == "month":
                    vals = base[key] if key in base.columns else pd.Series(0,index=base.index)
                else:
                    present = [p for p in key if p in base.columns]
                    vals = base[present].sum(axis=1) if present else pd.Series(0,index=base.index)
                col_vals = [float(vals.get(idx,0)) for idx in rows_index]
                col_vals.append(sum(col_vals))
                result[col_label] = [f"${v:,.0f}" for v in col_vals]
            return pd.DataFrame(result)

        _tm_piv_rgn  = _tm_actuals_pivot(_tm_actuals, "region")
        _tm_piv_prod = _tm_actuals_pivot(_tm_actuals, "product")

        # ── Currency pivot ────────────────────────────────────────────────────
        _tm_piv_cur = pd.DataFrame()
        if "currency" in _tm_actuals.columns:
            from shared.config import get_fx_rate
            _all_periods = sorted(_tm_actuals["period"].unique())
            _cur_rows = []
            for _cur, _cdf in _tm_actuals[_tm_actuals["revenue_usd"] > 0].groupby("currency"):
                _row = {"Currency": _cur}
                _total_usd = 0
                for _p in _all_periods:
                    _pdf = _cdf[_cdf["period"] == _p]
                    _rev_usd   = float(_pdf["revenue_usd"].sum())
                    _rev_local = float((_pdf["rate_local"] * _pdf["hours"]).sum()) if "rate_local" in _pdf.columns else _rev_usd
                    _fx        = get_fx_rate(_cur, _p)
                    _mo        = pd.Timestamp(_p + "-01").strftime("%b")
                    if _cur != "USD":
                        _row[_mo] = f"${_rev_usd:,.0f} ({_cur} {_rev_local:,.0f} · FX: {_fx:.4f})"
                    else:
                        _row[_mo] = f"${_rev_usd:,.0f}"
                    _total_usd += _rev_usd
                _row["Total USD"] = f"${_total_usd:,.0f}"
                _cur_rows.append(_row)
            # Total row — sum of all currencies converted to USD
            _total_row = {"Currency": "Total"}
            _grand_total = 0
            for _p in _all_periods:
                _mo = pd.Timestamp(_p + "-01").strftime("%b")
                _p_total = float(_tm_actuals[_tm_actuals["period"] == _p]["revenue_usd"].sum())
                _total_row[_mo] = f"${_p_total:,.0f}"
                _grand_total += _p_total
            _total_row["Total USD"] = f"${_grand_total:,.0f}"
            _cur_rows.append(_total_row)
            _tm_piv_cur = pd.DataFrame(_cur_rows)


st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">T&amp;M Actuals Summary (from NS Time Detail)</div>', unsafe_allow_html=True)
_cur_tab_label = "By Currency 💱" if not _tm_piv_cur.empty else "By Currency"
tab1, tab2, tab3 = st.tabs(["By Region", "By Product", _cur_tab_label])
with tab1:
    if not _tm_piv_rgn.empty:
        st.dataframe(_style_pivot(_tm_piv_rgn.rename(columns={"region":"Region"})),
                     use_container_width=True, hide_index=True)
with tab2:
    if not _tm_piv_prod.empty:
        st.dataframe(_style_pivot(_tm_piv_prod.rename(columns={"product":"Product"})),
                     use_container_width=True, hide_index=True)
with tab3:
    if not _tm_piv_cur.empty:
        st.caption("Revenue converted to USD using monthly average FX rates (shared/config.py). "
                   "Local currency amounts shown in brackets.")
        st.dataframe(_tm_piv_cur, use_container_width=True, hide_index=True)
    else:
        st.info("Add a 'Currency' column to your NS Time Detail export to see FX breakdown.")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION — Monthly Breakdown by Region & Product
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Monthly Revenue Breakdown — FF + T&amp;M (USD)</div>', unsafe_allow_html=True)

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

def _style_pivot(df):
    """Apply subtle background to quarter subtotal and FY/YTD columns."""
    # Identify which columns are Q or FY/YTD subtotals
    subtotal_cols = [c for c in df.columns
                     if str(c).startswith("Q") or c in ("YTD", "FY26", "FY27", "FY28")]

    def _highlight(col):
        if col.name in subtotal_cols:
            return ["background-color: rgba(68,114,196,0.18); font-weight: 600"] * len(col)
        return [""] * len(col)

    # Also bold the Total row
    def _bold_total(row):
        if str(row.iloc[0]) == "Total":
            return ["font-weight: 700; border-top: 1px solid rgba(128,128,128,0.4)"] * len(row)
        return [""] * len(row)

    return (df.style
              .apply(_highlight, axis=0)
              .apply(_bold_total, axis=1))

if not _piv_region.empty:
    st.markdown("**By Region**")
    st.dataframe(_style_pivot(_piv_region), use_container_width=True, hide_index=True)

if not _piv_product.empty:
    st.markdown("**By Product**")
    st.dataframe(_style_pivot(_piv_product), use_container_width=True, hide_index=True)

# ── T&M Monthly Actuals pivot — appended to monthly breakdown section ────────
if not _tm_actuals.empty:
    st.markdown("**T&M By Region**")
    if not _tm_piv_rgn.empty:
        st.dataframe(_style_pivot(_tm_piv_rgn.rename(columns={"region":"Region"})),
                     use_container_width=True, hide_index=True)
    st.markdown("**T&M By Product**")
    if not _tm_piv_prod.empty:
        st.dataframe(_style_pivot(_tm_piv_prod.rename(columns={"product":"Product"})),
                     use_container_width=True, hide_index=True)

# ── Trend Analysis and MoM by Region — Excel only (removed from UI) ──────────
_trend_full = (_monthly_totals.reset_index()
               if not _monthly_totals.empty else pd.DataFrame(columns=["period","usd_amount"]))
_trend_full.columns = ["Period", "Revenue (USD)"]
_trend_full["Revenue (USD)"] = pd.to_numeric(_trend_full["Revenue (USD)"], errors="coerce").fillna(0)
_trend_full["MoM Change"]    = _trend_full["Revenue (USD)"].diff()
_trend_full["MoM %"]         = _trend_full["Revenue (USD)"].pct_change() * 100
_trend_full["Cumulative YTD"]= _trend_full["Revenue (USD)"].cumsum()
_trend_disp = _trend_full.copy()
_trend_disp["Period"]         = _trend_disp["Period"].apply(lambda m: pd.Timestamp(m+"-01").strftime("%b %Y"))
_trend_disp["Revenue (USD)"]  = _trend_disp["Revenue (USD)"].apply(_fmt)
_trend_disp["MoM Change"]     = _trend_disp["MoM Change"].apply(lambda v: ("↑ " if v>=0 else "↓ ")+_fmt(abs(v)) if pd.notna(v) else "—")
_trend_disp["MoM %"]          = _trend_disp["MoM %"].apply(lambda v: f"{'↑' if v>=0 else '↓'} {abs(v):.1f}%" if pd.notna(v) else "—")
_trend_disp["Cumulative YTD"] = _trend_disp["Cumulative YTD"].apply(_fmt)

# ── Monthly trend and Project Detail — Excel only ────────────────────────────
_trend = slices.groupby("period")["usd_amount"].sum().reset_index()
_trend.columns = ["Period", "Revenue (USD)"]
_trend = _trend.sort_values("Period")
_trend["Revenue (USD)"] = pd.to_numeric(_trend["Revenue (USD)"], errors="coerce").fillna(0).round(2)

_detail_cols = ["project_id","project_name","project_manager","product",
                "region","currency","period","local_amount","usd_amount"]
_detail_cols = [c for c in _detail_cols if c in slices.columns]
_detail = slices[_detail_cols].copy()
for _nc in ("usd_amount","local_amount"):
    if _nc in _detail.columns:
        _detail[_nc] = pd.to_numeric(_detail[_nc], errors="coerce").fillna(0).round(2)

if "project_id" in _detail.columns and "project_name" in _detail.columns:
    _detail["_proj_label"] = _detail["project_id"].astype(str) + ": " + _detail["project_name"].fillna("").astype(str)
else:
    _detail["_proj_label"] = _detail.get("project_id", pd.Series("", index=_detail.index)).astype(str)

_all_periods_sorted = sorted(_detail["period"].unique()) if "period" in _detail.columns else []
_detail_pivot = pd.DataFrame()
if _all_periods_sorted:
    _dp = (_detail.groupby(["_proj_label","period"])["usd_amount"]
                  .sum().unstack(fill_value=0)
                  .reindex(columns=_all_periods_sorted, fill_value=0))
    _dp.columns = [pd.Timestamp(m+"-01").strftime("%b %Y") for m in _dp.columns]
    _dp["Total"] = _dp.sum(axis=1)
    _dp = _dp.round(2)
    _dp.index.name = "Project"
    _detail_pivot = _dp.reset_index()

# ── T&M Monthly Actuals (from NS time entries) ───────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ── Billing Exceptions ────────────────────────────────────────────────────────
if df_ns_session is not None:
    _mismatches = get_billing_mismatches(df_ns_session)
    if not _mismatches.empty:
        st.markdown('<hr class="divider">', unsafe_allow_html=True)
        st.markdown('<div class="section-label">Billing Exceptions</div>', unsafe_allow_html=True)

        _tm_nb  = _mismatches[_mismatches["mismatch_flag"] == "⚠️ T&M / Non-Billable"]
        _ff_bil = _mismatches[_mismatches["mismatch_flag"] == "⚠️ FF / Billable hours"]

        _tm_nb_hrs  = float(_tm_nb["hours"].sum())  if not _tm_nb.empty  else 0.0
        _ff_bil_hrs = float(_ff_bil["hours"].sum()) if not _ff_bil.empty else 0.0

        col1, col2 = st.columns(2)
        with col1:
            _c = "#E74C3C" if _tm_nb_hrs > 0 else "inherit"
            st.markdown(
                f'<div class="metric-card"><div class="metric-val" style="color:{_c}">{_tm_nb_hrs:,.1f}h</div>'
                f'<div class="metric-lbl">T&M / Non-Billable · {len(_tm_nb)} rows · may be misconfigured in NS — review for billing</div></div>',
                unsafe_allow_html=True)
        with col2:
            _c = "#F39C12" if _ff_bil_hrs > 0 else "inherit"
            st.markdown(
                f'<div class="metric-card"><div class="metric-val" style="color:{_c}">{_ff_bil_hrs:,.1f}h</div>'
                f'<div class="metric-lbl">FF / Billable hours · {len(_ff_bil)} rows · may appear on invoice — review</div></div>',
                unsafe_allow_html=True)

        with st.expander(f"Billing Exception Detail ({len(_mismatches)} rows)", expanded=False):
            st.dataframe(_mismatches, use_container_width=True, hide_index=True,
                         column_config={
                             "employee":       st.column_config.TextColumn("Employee",     width="medium"),
                             "project":        st.column_config.TextColumn("Project",      width="large"),
                             "project_type":   st.column_config.TextColumn("Project Type", width="medium"),
                             "billing_type":   st.column_config.TextColumn("Billing Type", width="small"),
                             "non_billable":   st.column_config.TextColumn("Non-Billable", width="small"),
                             "hours":          st.column_config.NumberColumn("Hours",       width="small", format="%.2f"),
                             "mismatch_flag":  st.column_config.TextColumn("Flag",         width="medium"),
                         })

# ── Itemized Time Entry Detail (Excel only) ──────────────────────────────────
# Build _det for Excel sheet — no UI rendering
_det_excel = pd.DataFrame()
if df_ns_session is not None:
    _det = df_ns_session.copy()

    # Format dates
    if "date" in _det.columns:
        if pd.api.types.is_datetime64_any_dtype(_det["date"]):
            _det["date"] = _det["date"].dt.strftime("%Y-%m-%d").fillna("")
        else:
            try:
                _det["date"] = pd.to_datetime(_det["date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
            except Exception:
                _det["date"] = _det["date"].astype(str)

    if "period" not in _det.columns and "date" in _det.columns:
        _det["period"] = _det["date"].astype(str).str[:7]

    # Classification
    _bt_d  = _det.get("billing_type", pd.Series("", index=_det.index)).fillna("").str.lower()
    _nb_d  = _det.get("non_billable", pd.Series("", index=_det.index)).fillna("").astype(str).str.strip().str.lower()
    _is_nb_d  = _nb_d.isin(["true","t","yes","1","y"])
    _is_tm_d  = _bt_d.str.contains("t&m|time", na=False)
    _is_ff_d  = _bt_d.str.contains("fixed fee|fixed.fee|fixed_fee|\bff\b", na=False, regex=True)
    _is_int_d = _bt_d.str.contains("internal", na=False)

    _det["_classification"] = "Other / Unclassified"
    _det.loc[_is_int_d,              "_classification"] = "Internal"
    _det.loc[_is_ff_d  &  _is_nb_d, "_classification"] = "FF / Non-Billable"
    _det.loc[_is_ff_d  & ~_is_nb_d, "_classification"] = "FF / Billable ⚠️"
    _det.loc[_is_tm_d  &  _is_nb_d, "_classification"] = "T&M / Non-Billable ⚠️"
    _det.loc[_is_tm_d  & ~_is_nb_d, "_classification"] = "T&M / Billable ✓"

    # Revenue
    _det["hours"]   = pd.to_numeric(_det.get("hours",   0), errors="coerce").fillna(0)
    _det["ns_rate"] = pd.to_numeric(_det.get("ns_rate", 0), errors="coerce").fillna(0)

    from shared.config import get_fx_rate as _gfx
    def _to_usd_det(row):
        cur = str(row.get("currency", "USD") or "USD").strip().upper()
        per = str(row.get("period", ""))
        fx  = _gfx(cur, per) if cur != "USD" else 1.0
        return round(row["hours"] * row["ns_rate"] * fx, 2)

    _det["_rev_local"] = (_det["hours"] * _det["ns_rate"]).round(2)
    _det["_rev_usd"]   = _det.apply(_to_usd_det, axis=1)

    # Build display DataFrame
    _det_col_map = {
        "internal_id":    "Internal ID",
        "project_id":     "Project ID",
        "project":        "Project",
        "employee":       "Employee",
        "date":           "Date",
        "period":         "Period",
        "billing_type":   "Billing Type",
        "non_billable":   "Non-Billable",
        "currency":       "Currency",
        "hours":          "Hours",
        "ns_rate":        "Rate (Local)",
        "entry_status":   "Entry Status",
        "billing_flag":   "Flag",
        "_classification":"Classification",
        "_rev_local":     "Revenue (Local)",
        "_rev_usd":       "Revenue (USD)",
    }
    _show = [c for c in _det_col_map if c in _det.columns]
    _det_excel = _det[_show].rename(columns=_det_col_map).copy()


# ── Excel download ─────────────────────────────────────────────────────────
st.markdown("")

# Rolling 15-month window for all pivot columns
_today       = pd.Timestamp.today()
_roll_start  = _today.strftime("%Y-%m")
_roll_end    = ((_today + pd.DateOffset(months=14))
                .strftime("%Y-%m"))
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
    # ── FF Revenue by Project × Month pivot ─────────────────────────────────
    if not slices.empty:
        from shared.config import RECONCILE_IMPL_SKU as _RIMPL

        # Separate Reconcile impl rows from all other FF rows
        def _is_reconcile_impl(ci):
            s = str(ci).strip()
            return s.split(" : ")[-1].strip() == _RIMPL

        _recon_slices = slices[slices["charge_item"].apply(_is_reconcile_impl)].copy()
        _other_slices = slices[~slices["charge_item"].apply(_is_reconcile_impl)].copy()

        # ── Reconcile Implementation pivot ───────────────────────────────────
        _recon_pivot = pd.DataFrame()
        if not _recon_slices.empty:
            _meta_cols = ["project_name", "product", "subscription_id",
                          "subscription_item", "currency",
                          "service_start", "service_end_orig",
                          "rev_rec_start", "rev_rec_end",
                          "impl_gross", "license_sku",
                          "license_cost_local", "license_currency", "license_cost_usd",
                          "carve_max", "notes"]
            _meta_cols = [c for c in _meta_cols if c in _recon_slices.columns]
            _meta_r = (_recon_slices.groupby(["project_id","charge_item"])[_meta_cols]
                       .first().reset_index())
            _all_p_r     = sorted(_recon_slices["period"].unique())
            _display_p_r = [p for p in _all_p_r if _roll_start <= p <= _roll_end]

            # Pivot all periods — Total Carve reflects full recognition window
            _piv_r = (_recon_slices.groupby(["project_id","charge_item","period"])["usd_amount"]
                      .sum().unstack(fill_value=0).reset_index())
            _piv_r.columns.name = None

            # Total Carve = sum across ALL periods
            _all_amt_cols = [p for p in _all_p_r if p in _piv_r.columns]
            _piv_r["Total Carve"] = _piv_r[_all_amt_cols].sum(axis=1)

            # Rename only rolling window periods for column headers
            _mo_r = {p: pd.Timestamp(p+"-01").strftime("%b %Y") for p in _display_p_r}
            _piv_r = _piv_r.rename(columns=_mo_r)
            _mc_r = [_mo_r[p] for p in _display_p_r if _mo_r[p] in _piv_r.columns]

            _recon_pivot = _meta_r.merge(
                _piv_r[["project_id","charge_item"] + _mc_r + ["Total Carve"]],
                on=["project_id","charge_item"], how="left")
            # Rename rev_start → charge_start_date for clarity
            if "rev_start" in _recon_pivot.columns:
                _recon_pivot = _recon_pivot.rename(columns={"rev_start": "charge_start_date"})
            # Rename service_start → charge_start_date for Reconcile tab
            if "service_start" in _recon_pivot.columns:
                _recon_pivot = _recon_pivot.rename(columns={"service_start": "charge_start_date",
                                                             "service_end_orig": "service_end"})
            _ord_r = [c for c in ["project_id","project_name","product","subscription_id",
                                    "subscription_item","currency","charge_start_date","service_end",
                                    "rev_rec_start","rev_rec_end","impl_gross",
                                    "license_sku","license_cost_local","license_currency",
                                    "license_cost_usd","carve_max","notes"]
                      if c in _recon_pivot.columns]
            _ord_r += _mc_r + ["Total Carve"]
            _recon_pivot = _recon_pivot[[c for c in _ord_r if c in _recon_pivot.columns]]
            _recon_pivot.to_excel(_xl, sheet_name="Reconcile Carve Detail", index=False)

        # ── All FF rows pivot (standard) — rolling 15-month window ─────────────
        _all_periods_all  = sorted(slices["period"].unique())
        _display_p_all    = [p for p in _all_periods_all if _roll_start <= p <= _roll_end]
        _meta_cols_all = ["project_name", "region", "product", "subscription_id",
                          "subscription_item", "currency",
                          "service_start", "service_end_orig",
                          "rev_rec_start", "rev_rec_end", "notes"]
        _meta_cols_all = [c for c in _meta_cols_all if c in slices.columns]
        _meta_all = (slices.groupby(["project_id","charge_item"])[_meta_cols_all]
                     .first().reset_index())
        _piv = (slices.groupby(["project_id","charge_item","period"])["usd_amount"]
                .sum().unstack(fill_value=0).reset_index())
        _piv.columns.name = None
        # Rev Amount = Total USD = sum across ALL periods (full window)
        _all_amt_all = [p for p in _all_periods_all if p in _piv.columns]
        _piv["Rev Amount"] = _piv[_all_amt_all].sum(axis=1)
        _month_rename = {p: pd.Timestamp(p+"-01").strftime("%b %Y") for p in _display_p_all}
        _piv = _piv.rename(columns=_month_rename)
        _month_cols = [_month_rename[p] for p in _display_p_all if _month_rename[p] in _piv.columns]
        _ff_proj_pivot = _meta_all.merge(
            _piv[["project_id","charge_item","Rev Amount"] + _month_cols],
            on=["project_id","charge_item"], how="left")
        _ord_all = [c for c in ["project_id","project_name","region","product","subscription_id",
                                  "subscription_item","currency",
                                  "service_start","service_end_orig",
                                  "rev_rec_start","rev_rec_end",
                                  "notes","Rev Amount"] if c in _ff_proj_pivot.columns]
        _ord_all += _month_cols
        _ff_proj_pivot = _ff_proj_pivot[[c for c in _ord_all if c in _ff_proj_pivot.columns]]
        _ff_proj_pivot.to_excel(_xl, sheet_name="FF Rev by Project (Rolling)", index=False)

        # ── FF Rev by Project YTD ────────────────────────────────────────────
        _ytd_start     = f"{_today.year}-01"
        _ytd_end       = f"{_today.year}-12"
        _display_p_ytd = [p for p in _all_periods_all if _ytd_start <= p <= _ytd_end]
        _piv_ytd = (slices.groupby(["project_id","charge_item","period"])["usd_amount"]
                    .sum().unstack(fill_value=0).reset_index())
        _piv_ytd.columns.name = None
        _piv_ytd["Rev Amount"] = _piv_ytd[[p for p in _all_periods_all if p in _piv_ytd.columns]].sum(axis=1)
        _mo_ytd = {p: pd.Timestamp(p+"-01").strftime("%b %Y") for p in _display_p_ytd}
        _piv_ytd = _piv_ytd.rename(columns=_mo_ytd)
        _mc_ytd = [_mo_ytd[p] for p in _display_p_ytd if _mo_ytd[p] in _piv_ytd.columns]
        _ff_proj_ytd = _meta_all.merge(
            _piv_ytd[["project_id","charge_item","Rev Amount"] + _mc_ytd],
            on=["project_id","charge_item"], how="left")
        _ord_ytd = [c for c in ["project_id","project_name","region","product","subscription_id",
                                  "subscription_item","currency",
                                  "service_start","service_end_orig",
                                  "rev_rec_start","rev_rec_end",
                                  "notes","Rev Amount"] if c in _ff_proj_ytd.columns]
        _ord_ytd += _mc_ytd
        _ff_proj_ytd = _ff_proj_ytd[[c for c in _ord_ytd if c in _ff_proj_ytd.columns]]
        _ff_proj_ytd.to_excel(_xl, sheet_name="FF Rev by Project YTD", index=False)

    # FF Project Detail tabs removed — superseded by FF Rev by Project and Reconcile Carve Detail
    if df_tm_sow is not None and not df_tm.empty:
        df_tm[[c for c in ["account_name","opportunity_name","opportunity_owner",
                            "product","region","close_date","sow_hours",
                            "sow_amount_usd","sow_rate_usd","ns_project",
                            "ns_hours_worked","ns_revenue_to_date"]
               if c in df_tm.columns]].to_excel(_xl, sheet_name="TM Detail", index=False)
    if not _det_excel.empty:
        _det_excel.to_excel(_xl, sheet_name="Time Entry Detail", index=False)
    if df_rev_raw is not None and "notes" in df_rev_raw.columns:
        # Only show rows with warning flags — exclude successfully matched rows
        _rec_xl = df_rev_raw[df_rev_raw["notes"].str.startswith("⚠️", na=False)].copy()
        if not _rec_xl.empty:
            _rec_xl_cols = [c for c in ["charge_item","subscription_id","subscription_item",
                                         "project_id","gross_amount","currency",
                                         "rev_start","service_end",
                                         "notes"] if c in _rec_xl.columns]
            _rec_xl_out = _rec_xl[_rec_xl_cols].copy()
            if "rev_start" in _rec_xl_out.columns:
                _rec_xl_out = _rec_xl_out.rename(columns={"rev_start": "service_start"})
            _rec_xl_out.to_excel(_xl, sheet_name="Carve Flags", index=False)

# ── Apply Zone brand formatting ──────────────────────────────────────────────
try:
    from shared.excel_formatter import apply_zone_formatting

    # Collect metrics for Dashboard
    _ns_date_str = ""
    if df_ns is not None and "date" in df_ns.columns:
        _ns_latest = df_ns["date"].dropna().max()
        try:
            _ns_date_str = pd.Timestamp(_ns_latest).strftime("%-d %B %Y")
        except Exception:
            _ns_date_str = str(_ns_latest)[:10]

    _blurb = (
        "This report calculates PS revenue recognition from NetSuite charge exports and "
        "NS Time Detail. Fixed Fee projects: revenue recognised straight-line across a "
        "2–3 month window from service start. Reconcile, Capture and Approvals "
        "implementations apply a license-based carve-out capped at the Year 1 license "
        "cost. T&M projects: revenue = hours logged × billing rate, converted to USD "
        "using monthly average FX rates. Carve flags are surfaced for Finance review."
    )

    # By Region summary rows for dashboard
    _dash_region = []
    if not _rt.empty:
        for _, _rrow in _rt.iterrows():
            try:
                _rv = float(str(_rrow.get("YTD", 0)).replace("$","").replace(",","") or 0)
            except Exception:
                _rv = 0.0
            _pct = f"{_rv/max(float(ytd_total),1)*100:.1f}%" if ytd_total else "0.0%"
            _dash_region.append({
                "Region": str(_rrow.get("Region", "")),
                "YTD (USD)": f"${_rv:,.0f}",
                "% of Total": _pct,
            })

    # By Product summary rows for dashboard
    _dash_product = []
    if not _pt.empty:
        for _, _prow in _pt.iterrows():
            try:
                _pv = float(str(_prow.get("YTD", 0)).replace("$","").replace(",","") or 0)
            except Exception:
                _pv = 0.0
            _pct = f"{_pv/max(float(ytd_total),1)*100:.1f}%" if ytd_total else "0.0%"
            _dash_product.append({
                "Product": str(_prow.get("Product", "")),
                "YTD (USD)": f"${_pv:,.0f}",
                "% of Total": _pct,
            })

    # Trend rows for dashboard — stringify all values for safe Excel writing
    _dash_trend = []
    if not _trend_disp.empty:
        for _, _trow in _trend_disp.head(12).iterrows():
            _dash_trend.append({k: (f"${float(v):,.0f}" if isinstance(v, (int, float)) and k != "Period"
                                    else str(v)) for k, v in _trow.items()})

    _flag_count = 0
    if df_rev_raw is not None and "notes" in df_rev_raw.columns:
        _flag_count = int(df_rev_raw["notes"].str.startswith("⚠️", na=False).sum())

    _ff_ytd  = float(ytd_df["usd_amount"].sum()) if not ytd_df.empty else 0.0
    _tm_ytd  = float(_tm_monthly_early.sum()) if not _tm_monthly_early.empty else 0.0
    _recon_ytd = 0.0
    if not slices.empty and "notes" in slices.columns:
        from shared.config import RECONCILE_IMPL_SKU as _RDASH
        _recon_mask = slices["charge_item"].str.contains(_RDASH, na=False)
        _recon_ytd  = float(slices.loc[_recon_mask & slices["period"].isin(ytd_months), "usd_amount"].sum())

    _mom_display = f"{_mom_pct:+.1f}%" if _mom_pct is not None else "—"

    def _fmt_usd(v):
        try: return f"${float(v):,.0f}"
        except Exception: return "$0"

    _dash_metrics = {
        "ytd":         _fmt_usd(ytd_total),
        "qtd":         _fmt_usd(qtd_total),
        "mtd":         _fmt_usd(mtd_total),
        "full_mo":     _fmt_usd(full_month),
        "run_rate":    _fmt_usd(_run_rate),
        "mom":         _mom_display,
        "ff_ytd":      _fmt_usd(_ff_ytd),
        "tm_ytd":      _fmt_usd(_tm_ytd),
        "recon_ytd":   _fmt_usd(_recon_ytd),
        "ff_projects": int(slices["project_id"].nunique()) if not slices.empty else 0,
        "tm_projects": int(df_tm["ns_project"].dropna().nunique()) if df_tm is not None and not df_tm.empty else 0,
        "flag_count":  int(_flag_count),
        "trend_rows":  _dash_trend,
        "region_rows": _dash_region,
        "product_rows":_dash_product,
    }

    _formatted_bytes = apply_zone_formatting(
        _buf.getvalue(), _dash_metrics, _ns_date_str, _blurb
    )
except Exception as _fmt_err:
    # Fallback to unformatted if formatter fails
    _formatted_bytes = _buf.getvalue()
    st.warning(f"Formatting error (unformatted report downloaded): {_fmt_err}")

st.download_button(
    "⬇ Download Revenue Report (Excel)",
    data=_formatted_bytes,
    file_name=f"ps_revenue_report_{today.strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.markdown(
    '<div style="font-size:11px;opacity:.4;text-align:center;margin-top:20px">'
    'PS Reporting Tools · Internal use only · Fixed Fee Implementation Revenue</div>',
    unsafe_allow_html=True)
