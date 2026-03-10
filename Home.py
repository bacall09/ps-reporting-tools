"""
PS Tools — Home
"""
import streamlit as st

st.set_page_config(
    page_title="PS Tools",
    page_icon="🛠️",
    layout="wide"
)

st.markdown("""
<div style='background:#1e2c63;padding:32px 40px 24px;border-radius:8px;margin-bottom:32px'>
    <h1 style='color:#ffffff;font-family:Manrope,sans-serif;margin:0;font-size:32px'>
        🛠️ Professional Services Tools
    </h1>
    <p style='color:#a0aec0;font-family:Manrope,sans-serif;margin:8px 0 0;font-size:15px'>
        Internal tooling for the ZCO PS Team
    </p>
</div>
""", unsafe_allow_html=True)

st.markdown("### Available Tools")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("""
    <div style='border:1px solid #e2e8f0;border-radius:8px;padding:24px;height:160px'>
        <div style='font-size:28px'>📈</div>
        <div style='font-weight:700;font-size:16px;margin:8px 0 4px;font-family:Manrope,sans-serif'>
            Utilization Credit Report
        </div>
        <div style='color:#718096;font-size:13px;font-family:Manrope,sans-serif'>
            Upload NetSuite time detail → generate utilization Excel report with credits, overruns, and PS region metrics.
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.page_link("pages/1_Utilization_Report.py", label="Open Report →")

with col2:
    st.markdown("""
    <div style='border:1px dashed #cbd5e0;border-radius:8px;padding:24px;height:160px;background:#f7fafc'>
        <div style='font-size:28px'>🔜</div>
        <div style='font-weight:700;font-size:16px;margin:8px 0 4px;font-family:Manrope,sans-serif;color:#a0aec0'>
            Coming Soon
        </div>
        <div style='color:#a0aec0;font-size:13px;font-family:Manrope,sans-serif'>
            Next report or tool goes here.
        </div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div style='border:1px dashed #cbd5e0;border-radius:8px;padding:24px;height:160px;background:#f7fafc'>
        <div style='font-size:28px'>🔜</div>
        <div style='font-weight:700;font-size:16px;margin:8px 0 4px;font-family:Manrope,sans-serif;color:#a0aec0'>
            Coming Soon
        </div>
        <div style='color:#a0aec0;font-size:13px;font-family:Manrope,sans-serif'>
            Next report or tool goes here.
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")
st.caption("PS Tools · ZCO Professional Services · Internal use only")
