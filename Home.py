"""
PS Tools — Home
"""
import streamlit as st

st.set_page_config(
    page_title="PS Tools",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        h1, h2, h3, h4, p, div, label, button, caption { font-family: 'Manrope', sans-serif !important; }
        .tool-row {
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 24px 28px;
            background: #ffffff;
            margin-bottom: 16px;
        }
        .tool-row-soon {
            border: 1px dashed #cbd5e0;
            border-radius: 10px;
            padding: 24px 28px;
            background: #f7fafc;
            margin-bottom: 16px;
        }
        .tool-title { font-weight: 700; font-size: 17px; color: #1e2c63; margin-bottom: 6px; font-family: 'Manrope', sans-serif; }
        .tool-title-soon { font-weight: 700; font-size: 17px; color: #a0aec0; margin-bottom: 6px; font-family: 'Manrope', sans-serif; }
        .tool-desc { color: #4a5568; font-size: 13px; font-family: 'Manrope', sans-serif; line-height: 1.65; margin-bottom: 0; }
        .tool-badge { display: inline-block; font-size: 11px; font-weight: 600; color: #4472C4; background: #EBF0FB; border-radius: 4px; padding: 2px 8px; margin-bottom: 10px; font-family: 'Manrope', sans-serif; letter-spacing: 0.5px; text-transform: uppercase; }
        .tool-badge-soon { display: inline-block; font-size: 11px; font-weight: 600; color: #a0aec0; background: #f0f0f0; border-radius: 4px; padding: 2px 8px; margin-bottom: 10px; font-family: 'Manrope', sans-serif; letter-spacing: 0.5px; text-transform: uppercase; }
        .stPageLink a { display: inline-block; margin-top: 14px; color: #ffffff !important; background: #1e2c63; font-weight: 600; font-family: 'Manrope', sans-serif !important; text-decoration: none; padding: 7px 18px; border-radius: 6px; font-size: 13px; }
        .stPageLink a:hover { background: #4472C4; }
    </style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='background:#1e2c63;padding:32px 40px 28px;border-radius:10px;margin-bottom:32px'>
    <div style='font-size:11px;color:#a0aec0;font-family:Manrope,sans-serif;letter-spacing:2px;text-transform:uppercase;margin-bottom:8px'>Professional Services</div>
    <h1 style='color:#ffffff;font-family:Manrope,sans-serif;margin:0;font-size:30px;font-weight:700'>PS Reporting Tools</h1>
    <p style='color:#a0aec0;font-family:Manrope,sans-serif;margin:10px 0 0;font-size:14px'>Internal tooling for the PS Team — select a report below to get started</p>
</div>
""", unsafe_allow_html=True)

st.markdown("#### Available Reports")
st.markdown("<div style='margin-bottom:12px'></div>", unsafe_allow_html=True)

# ── Card 1: Utilization Credit Report ────────────────────────────────────────
col_a, col_b = st.columns([5, 1])
with col_a:
    st.markdown("""
    <div class='tool-row'>
        <div class='tool-badge'>NetSuite Export</div>
        <div class='tool-title'>Utilization Credit Report</div>
        <div class='tool-desc'>
            The Utilization Credit Report tracks billable utilization for each PS consultant
            using NetSuite time detail exports. By distinguishing between in-scope hours vs.
            overrun hours on Fixed Fee projects, it provides clearer visibility into consultant
            performance and potential project scope overruns.
        </div>
    </div>
    """, unsafe_allow_html=True)
with col_b:
    st.markdown("<div style='margin-top:38px'></div>", unsafe_allow_html=True)
    st.page_link("pages/1_Utilization_Report.py", label="Open Report →")

# ── Card 2: FF Workload Score ─────────────────────────────────────────────────
col_c, col_d = st.columns([5, 1])
with col_c:
    st.markdown("""
    <div class='tool-row'>
        <div class='tool-badge'>Smartsheets + NetSuite</div>
        <div class='tool-title'>FF Workload Score</div>
        <div class='tool-desc'>
            This report measures consultant workload across active Fixed Fee (FF) projects
            using a weighted scoring model. It helps PS leadership identify potential overload,
            compare workloads across regions, and flag projects nearing or exceeding
            contractual limits.
        </div>
    </div>
    """, unsafe_allow_html=True)
with col_d:
    st.markdown("<div style='margin-top:38px'></div>", unsafe_allow_html=True)
    st.page_link("pages/2_Workload_Health_Score.py", label="Open Report →")

# ── Card 3: Coming Soon ───────────────────────────────────────────────────────
st.markdown("""
<div class='tool-row-soon'>
    <div class='tool-badge-soon'>Coming Soon</div>
    <div class='tool-title-soon'>Capacity Outlook</div>
    <div class='tool-desc' style='color:#a0aec0'>
        Three-horizon capacity view combining NetSuite actuals, Salesforce contracts,
        and pipeline forecast. Joined on Project ID.
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("---")
st.caption("PS Reporting Tools · Internal use only")
