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
        .tool-card {
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 24px;
            height: 160px;
            background: #ffffff;
        }
        .tool-card-soon {
            border: 1px dashed #cbd5e0;
            border-radius: 8px;
            padding: 24px;
            height: 160px;
            background: #f7fafc;
        }
        .tool-title { font-weight: 700; font-size: 16px; color: #1e2c63; margin-bottom: 6px; font-family: 'Manrope', sans-serif; }
        .tool-title-soon { font-weight: 700; font-size: 16px; color: #a0aec0; margin-bottom: 6px; font-family: 'Manrope', sans-serif; }
        .tool-desc { color: #718096; font-size: 13px; font-family: 'Manrope', sans-serif; line-height: 1.5; }
        .stPageLink a {
            color: #4472C4 !important;
            font-weight: 600;
            font-family: 'Manrope', sans-serif !important;
            text-decoration: none;
        }
    </style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='background:#1e2c63;padding:32px 40px 28px;border-radius:8px;margin-bottom:32px'>
    <div style='font-size:12px;color:#a0aec0;font-family:Manrope,sans-serif;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:8px'>
        Professional Services
    </div>
    <h1 style='color:#ffffff;font-family:Manrope,sans-serif;margin:0;font-size:30px;font-weight:700'>
        PS Reporting Tools
    </h1>
    <p style='color:#a0aec0;font-family:Manrope,sans-serif;margin:10px 0 0;font-size:14px'>
        Internal tooling for the PS Team — select a tool below to get started
    </p>
</div>
""", unsafe_allow_html=True)

# ── Tool cards ────────────────────────────────────────────────────────────────
st.markdown("#### Available Tools")
st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("""
    <div class='tool-card'>
        <div class='tool-title'>Utilization Credit Report</div>
        <div class='tool-desc'>
            Upload a NetSuite time detail export to generate a full utilization
            report with credits, overruns, PS region metrics, and project watch list.
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.page_link("pages/1_Utilization_Report.py", label="Open Report")

with col2:
    st.markdown("""
    <div class='tool-card-soon'>
        <div class='tool-title-soon'>Coming Soon</div>
        <div class='tool-desc'>Next report or tool goes here.</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class='tool-card-soon'>
        <div class='tool-title-soon'>Coming Soon</div>
        <div class='tool-desc'>Next report or tool goes here.</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")
st.caption("PS Reporting Tools · Internal use only")
