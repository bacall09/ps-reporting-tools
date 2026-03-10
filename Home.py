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

        /* ── Card styles — light mode ── */
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
        .tool-title      { font-weight: 700; font-size: 17px; color: #1e2c63; margin-bottom: 6px; font-family: 'Manrope', sans-serif; }
        .tool-title-soon { font-weight: 700; font-size: 17px; color: #a0aec0; margin-bottom: 6px; font-family: 'Manrope', sans-serif; }
        .tool-desc      { color: #4a5568; font-size: 13px; font-family: 'Manrope', sans-serif; line-height: 1.65; margin-bottom: 0; }
        .tool-desc-soon { color: #a0aec0; font-size: 13px; font-family: 'Manrope', sans-serif; line-height: 1.65; margin-bottom: 0; }
        .tool-badge      { display: inline-block; font-size: 11px; font-weight: 600; color: #4472C4; background: #EBF0FB; border-radius: 4px; padding: 2px 8px; margin-bottom: 10px; font-family: 'Manrope', sans-serif; letter-spacing: 0.5px; text-transform: uppercase; }
        .tool-badge-soon { display: inline-block; font-size: 11px; font-weight: 600; color: #a0aec0; background: #f0f0f0; border-radius: 4px; padding: 2px 8px; margin-bottom: 10px; font-family: 'Manrope', sans-serif; letter-spacing: 0.5px; text-transform: uppercase; }
        .tool-link { display: inline-block; margin-top: 14px; color: #4da6ff; font-weight: 600; font-family: 'Manrope', sans-serif; font-size: 13px; text-decoration: none; }
        .tool-link:hover { color: #7dc0ff; text-decoration: underline; }

        /* ── Open Report button ── */
        .stPageLink a {
            display: inline-block;
            margin-top: 14px;
            color: #ffffff !important;
            background: #1e2c63;
            font-weight: 600;
            font-family: 'Manrope', sans-serif !important;
            text-decoration: none;
            padding: 7px 18px;
            border-radius: 6px;
            font-size: 13px;
        }
        .stPageLink a:hover { background: #4472C4 !important; }

        /* ── Dark mode overrides ── */
        @media (prefers-color-scheme: dark) {
            .tool-row       { background: #1e1e2e; border-color: #2d2d44; }
            .tool-row-soon  { background: #16161f; border-color: #2d2d44; }
            .tool-title     { color: #c5d0f0; }
            .tool-desc      { color: #9aa3b8; }
            .tool-badge     { color: #7da9f0; background: #1e2c4a; }
            .stPageLink a   { background: #4472C4; color: #ffffff !important; }
            .stPageLink a:hover { background: #5a8ad4 !important; }
        }

        /* ── Streamlit dark theme class-based overrides ── */
        [data-theme="dark"] .tool-row       { background: #1e1e2e; border-color: #2d2d44; }
        [data-theme="dark"] .tool-row-soon  { background: #16161f; border-color: #2d2d44; }
        [data-theme="dark"] .tool-title     { color: #c5d0f0; }
        [data-theme="dark"] .tool-desc      { color: #9aa3b8; }
        [data-theme="dark"] .tool-badge     { color: #7da9f0; background: #1e2c4a; }
        [data-theme="dark"] .stPageLink a   { color: #4da6ff !important; background: none !important; }
        [data-theme="dark"] .stPageLink a:hover { background: #5a8ad4 !important; }

        /* Streamlit sets this on <html> or body in dark mode */
        .st-emotion-cache-bg [data-testid="stAppViewContainer"],
        [data-testid="stAppViewContainer"][class*="dark"] .tool-row { background: #1e1e2e; }
    </style>
""", unsafe_allow_html=True)

# ── JS dark mode detection — Streamlit doesn't always expose prefers-color-scheme ──
st.markdown("""
<script>
(function() {
    function applyTheme() {
        const isDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches
            || document.body.classList.contains('dark')
            || document.documentElement.getAttribute('data-theme') === 'dark';
        const cards = document.querySelectorAll('.tool-row, .tool-row-soon');
        const titles = document.querySelectorAll('.tool-title');
        const descs = document.querySelectorAll('.tool-desc');
        const badges = document.querySelectorAll('.tool-badge');
        const links = document.querySelectorAll('.stPageLink a');
        if (isDark) {
            cards.forEach(c => { c.style.background = '#1e1e2e'; c.style.borderColor = '#2d2d44'; });
            titles.forEach(t => t.style.color = '#c5d0f0');
            descs.forEach(d => d.style.color = '#9aa3b8');
            badges.forEach(b => { b.style.color = '#7da9f0'; b.style.background = '#1e2c4a'; });
            links.forEach(l => { l.style.color = '#4da6ff'; l.style.background = 'none'; });
        } else {
            links.forEach(l => { l.style.color = '#4da6ff'; l.style.background = 'none'; });
        }
    }
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', applyTheme);
    } else {
        applyTheme();
        setTimeout(applyTheme, 500);
    }
    if (window.matchMedia) {
        window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme);
    }
})();
</script>
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
    <a class='tool-link' href='/Utilization_Report'>Open Report →</a>
</div>
""", unsafe_allow_html=True)

# ── Card 2: FF Workload Score ─────────────────────────────────────────────────
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
    <a class='tool-link' href='/Workload_Health_Score'>Open Report →</a>
</div>
""", unsafe_allow_html=True)

# ── Card 3: Coming Soon ───────────────────────────────────────────────────────
st.markdown("""
<div class='tool-row-soon'>
    <div class='tool-badge-soon'>Coming Soon</div>
    <div class='tool-title-soon'>Capacity Outlook</div>
    <div class='tool-desc-soon'>
        Three-horizon capacity view combining NetSuite actuals, Salesforce contracts,
        and pipeline forecast. Joined on Project ID.
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("---")
st.caption("PS Reporting Tools · Internal use only")
