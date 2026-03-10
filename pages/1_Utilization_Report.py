"""
PS Utilization Credit Report
Upload a NetSuite time detail export to generate the utilization Excel report.
"""
import streamlit as st
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from shared.config import DEFAULT_SCOPE
from shared.utils import assign_credits, build_excel, auto_detect_columns

# ── Streamlit UI ──────────────────────────────────────────────────────────────
def main():
    st.markdown("""
        <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
            html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
            h1, h2, h3, .stMarkdown, .stDataFrame, label, button { font-family: 'Manrope', sans-serif !important; }
        </style>
        <div style='background-color:#1e2c63;padding:24px 32px;border-radius:8px;margin-bottom:24px;font-family:Manrope,sans-serif'>
            <h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif'>Professional Services Utilization Credit Report</h1>
            <p style='color:#aac4d0;margin:6px 0 0 0;font-size:14px;font-family:Manrope,sans-serif'>
                Upload your NetSuite time detail export to generate a utilization credit report.
                &nbsp;|&nbsp; <a href="https://3838224.app.netsuite.com/app/common/search/searchresults.nl?searchid=66732&saverun=T&whence=" style='color:#7da9f0;font-family:Manrope,sans-serif;'>Report Link</a>
            </p>
            <p style='color:#8ab0c0;margin:8px 0 0 0;font-size:12px;font-family:Manrope,sans-serif;line-height:1.6;'>This tool calculates <b>Utilization Credits</b> from NetSuite time detail exports. Credits are awarded as follows: <b>T&amp;M</b> projects receive full credit for all hours logged. <b>Fixed Fee</b> projects receive credit up to their scoped hours — hours beyond scope are tracked as overrun and excluded from credits. <b>Internal</b> time is excluded from utilization entirely and tracked separately as Admin Hours. Util&nbsp;% = Utilization Credits &divide; Hours This Period.</p>
        </div>
    """, unsafe_allow_html=True)

    # ── Adaptive metric color CSS ────────────────────────────
    # ── Upload ────────────────────────────────────────────────
    st.subheader("Step 1 — Upload NetSuite Time Detail Export")
    st.caption("Supported columns: Employee, Region, Project, Project Type, Billing Type, "
               "Hours to Date, Date, Hours, Approval Status, Case/Task/Event, Non-Billable")

    uploaded = st.file_uploader(
        "Drop your file here or click to browse",
        type=["csv", "xlsx", "xls"],
        help="Supports CSV and Excel files exported from NetSuite"
    )

    if not uploaded:
        # Show stored config as reference
        with st.expander("📋 View stored scope map & available hours"):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Fixed Fee Scope Hours**")
                scope_df = pd.DataFrame(list(DEFAULT_SCOPE.items()), columns=["Project Type","Scoped Hrs"])
                st.dataframe(scope_df, hide_index=True, use_container_width=True)
            with c2:
                st.markdown("**Available Hours by Region (2026)**")
                avail_df = pd.DataFrame([
                    {"Region": r, **months} for r, months in AVAIL_HOURS.items()
                ])
                st.dataframe(avail_df, hide_index=True, use_container_width=True)
        st.info("👆 Upload your NetSuite export to continue.")
        return

    # Load file
    try:
        ext = os.path.splitext(uploaded.name)[1].lower()
        df_raw = pd.read_excel(uploaded) if ext in (".xlsx",".xls") else \
                 pd.read_csv(uploaded, encoding="utf-8") if True else \
                 pd.read_csv(uploaded, encoding="latin-1")
    except Exception:
        try:
            df_raw = pd.read_csv(uploaded, encoding="latin-1")
        except Exception as e:
            st.error(f"Could not read file: {e}"); return

    st.success(f"✅ Loaded **{len(df_raw):,} rows** from `{uploaded.name}`")
    with st.expander("Preview raw data (first 5 rows)"):
        st.dataframe(df_raw.head(), use_container_width=True)

    st.divider()

    # ── Process ───────────────────────────────────────────────
    st.subheader("Step 2 — Generate Report")

    if st.button("▶️ Run Utilization Engine", type="primary"):
        with st.spinner("Processing..."):
            try:
                df, consumed, skipped_df = assign_credits(df_raw.copy(), DEFAULT_SCOPE)
            except Exception as e:
                st.error(f"Processing error: {e}"); return

        st.success("✅ Processing complete!")


        # ── Warn on unmapped employees ────────────────────────
        _unmapped = []
        for _emp in df["employee"].dropna().unique():
            _emp_s = str(_emp).strip()
            _loc = df[df["employee"]==_emp]["region"].iloc[0] if len(df[df["employee"]==_emp]) > 0 else ""
            _matched = any(
                _emp_s.lower().startswith(k.lower()) or k.lower().startswith(_emp_s.lower())
                for k in EMPLOYEE_LOCATION
            )
            if not _matched and not str(_loc).strip():
                _unmapped.append(_emp_s)
        if _unmapped:
            st.warning(
                f"⚠️ **{len(_unmapped)} employee(s) have no location defined** — "
                f"avail hours and PS region will show as Unknown. "
                f"Add them to EMPLOYEE_LOCATION in the app config.\n\n"
                + ", ".join(sorted(_unmapped))
            )

        # Metrics
        total_rows     = len(df[df["credit_tag"] != "SKIPPED"])
        total_credit   = df["credit_hrs"].sum()
        total_variance = df["variance_hrs"].sum()
        total_nb       = len(df[df["credit_tag"] == "NON-BILLABLE"])
        total_overrun  = len(df[df["credit_tag"] == "OVERRUN"])
        hours_this_period = df[df["credit_tag"] != "SKIPPED"]["hours"].sum() if "hours" in df.columns else 0
        total_admin    = df[df["billing_type"].str.lower() == "internal"]["hours"].sum()             if "billing_type" in df.columns else 0
        total_proj_overrun = df[df["credit_tag"] == "OVERRUN"]["variance_hrs"].sum()             if "variance_hrs" in df.columns else 0

        credit_pct  = total_credit       / hours_this_period if hours_this_period else 0
        overrun_pct = total_proj_overrun / hours_this_period if hours_this_period else 0
        admin_pct   = total_admin        / hours_this_period if hours_this_period else 0

        credit_color = "#2ecc71" if credit_pct >= 0.70 else "#f39c12" if credit_pct >= 0.60 else "#e74c3c"
        credit_label = "On target" if credit_pct >= 0.70 else "Below target" if credit_pct >= 0.60 else "At risk"

        # Max date in report
        if "date" in df.columns:
            max_date = pd.to_datetime(df["date"], errors="coerce").max()
            date_str = max_date.strftime("%-d %B %Y") if pd.notna(max_date) else "—"
        else:
            date_str = "—"
        st.markdown(f"<div style='font-size:13px;color:#a0a0a0;font-family:Manrope,sans-serif;margin-bottom:12px'>Data through <strong style='color:#ffffff'>{date_str}</strong></div>", unsafe_allow_html=True)

        def metric_card(label, value, pill_txt=None, pill_fg=None):
            pill = ""
            if pill_txt and pill_fg:
                pill = f"<div style='display:inline-block;margin-top:6px;padding:2px 10px;border-radius:999px;background-color:{pill_fg}33;font-size:13px;font-family:Manrope,sans-serif;color:{pill_fg}'>&#8593; {pill_txt}</div>"
            return f"<div style='font-size:14px;color:#a0a0a0;font-family:Manrope,sans-serif;margin-bottom:4px'>{label}</div><div style='font-size:36px;font-weight:700;color:inherit;font-family:Manrope,sans-serif;line-height:1.1'>{value}</div>{pill}"

        m1,m2,m3,m4,m5 = st.columns(5)
        with m1: st.markdown(metric_card("Projects This Period",   f"{df[df['billing_type'].fillna('').str.lower() != 'internal'].groupby(['project','project_type']).ngroups:,}"), unsafe_allow_html=True)
        with m2: st.markdown(metric_card("Hours This Period",      f"{hours_this_period:,.1f}"), unsafe_allow_html=True)
        with m3: st.markdown(metric_card("Utilization Credits",    f"{total_credit:,.1f}",    f"{credit_pct:.1%} of total hrs · {credit_label}", credit_color), unsafe_allow_html=True)
        with m4: st.markdown(metric_card("FF Project Overrun Hrs", f"{total_proj_overrun:,.1f}", f"{overrun_pct:.1%} of total hrs", "#ff4b4b"), unsafe_allow_html=True)
        with m5: st.markdown(metric_card("Admin Hrs",              f"{total_admin:,.1f}",     f"{admin_pct:.1%} of total hrs",    "#808495"), unsafe_allow_html=True)

        st.markdown("---")
        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            ["By Employee", "By Project", "ZCO Non-Billable", "Task Analysis", "Detail"]
        )

        with tab1:
            _ep = df[df["credit_tag"] != "SKIPPED"]
            emp_sum_ui = _ep.groupby(["employee","period"], as_index=False).agg(
                hours_this_period=("hours","sum"),
                credit_hrs=("credit_hrs","sum"),
                ff_overrun_hrs=("variance_hrs","sum"),
                admin_hrs=("hours", lambda x: df.loc[
                    (df["employee"].isin(_ep["employee"])) &
                    (df["billing_type"].str.lower()=="internal"), "hours"
                ].sum() if "billing_type" in df.columns else 0),
            ).sort_values(["employee","period"])
            # Build region lookup directly from df for UI context
            _emp_region_ui = df.dropna(subset=["region"]).groupby("employee")["region"].first().to_dict() if "region" in df.columns else {}
            emp_sum_ui["location"]   = emp_sum_ui["employee"].map(_emp_region_ui)
            emp_sum_ui["avail_hrs"]  = emp_sum_ui.apply(
                lambda r: get_avail_hours(r["location"], r["period"]) if r["location"] else None, axis=1)
            emp_sum_ui["util_pct"]   = emp_sum_ui.apply(
                lambda r: f"{r['credit_hrs']/r['avail_hrs']*100:.1f}%" if r["avail_hrs"] else "—", axis=1)
            display_cols = ["employee","location","period","avail_hrs",
                            "hours_this_period","credit_hrs","ff_overrun_hrs","util_pct"]
            st.dataframe(emp_sum_ui[[c for c in display_cols if c in emp_sum_ui.columns]],
                         use_container_width=True, hide_index=True)

        with tab2:
            proj_sum_ui = df[df["credit_tag"] != "SKIPPED"].groupby(
                ["project","project_type"], as_index=False
            ).agg(hours_this_period=("hours","sum"), credit_hrs=("credit_hrs","sum"),
                  ff_overrun_hrs=("variance_hrs","sum")).sort_values("project")
            st.dataframe(proj_sum_ui[["project","project_type","hours_this_period",
                         "credit_hrs","ff_overrun_hrs"]],
                         use_container_width=True, hide_index=True)

        with tab3:
            zco_df = df[df["credit_tag"] == "NON-BILLABLE"]
            if "task" in zco_df.columns and len(zco_df) > 0:
                zco_sum = zco_df.groupby(["task","employee","period"], as_index=False
                ).agg(hours=("hours","sum")).sort_values(["task","employee","period"])
                st.dataframe(zco_sum, use_container_width=True, hide_index=True)
            else:
                st.info("No ZCO Non-Billable entries in this dataset.")

        with tab4:
            ff_df = df[df["ff_task"].notna()] if "ff_task" in df.columns else pd.DataFrame()
            if len(ff_df) > 0:
                task_sum = ff_df.groupby(["ff_task","project_type"], as_index=False
                ).agg(hours=("hours","sum")).sort_values(["ff_task","project_type"])
                type_totals = ff_df.groupby("project_type")["hours"].sum()
                task_sum["pct_of_type"] = task_sum.apply(
                    lambda r: f"{r['hours']/type_totals.get(r['project_type'],1)*100:.1f}%", axis=1)
                st.dataframe(task_sum, use_container_width=True, hide_index=True)
            else:
                st.info("No Fixed Fee task data found. Check Billing Type and Task/Case columns.")

        with tab5:
            display_cols = ["employee","region","project","project_type","billing_type",
                            "hours_to_date","date","hours","credit_hrs","variance_hrs",
                            "previous_htd","credit_tag","notes"]
            existing = [c for c in display_cols if c in df.columns]
            st.dataframe(df[existing].head(100), use_container_width=True, hide_index=True)
            if len(df) > 100:
                st.caption(f"Showing first 100 of {len(df):,} rows. Full data in Excel download.")

        st.divider()

        # Download
        st.subheader("Download Report")
        with st.spinner("Building Excel file..."):
            excel_buf = build_excel(df, DEFAULT_SCOPE, consumed)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename  = f"utilization_report_{timestamp}.xlsx"

        st.download_button(
            label="⬇️ Download Excel Report",
            data=excel_buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(f"`{filename}` — 6 tabs: Processed Data · By Employee · By Project · "
                   f"ZCO Non-Billable · Task Analysis · Skipped Rows")


if __name__ == "__main__":
    main()