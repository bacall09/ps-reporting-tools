"""
PS Tools — Help & How-To Guide
"""
import streamlit as st

st.set_page_config(page_title="PS Tools · Help", layout="wide")

st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@300;400;600;700;800&display=swap" rel="stylesheet">
<style>
    html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }

    /* ── HERO ── */
    .help-hero {
        background:#050D1F;
        border-radius: 10px;
        padding: 36px 40px;
        margin-bottom: 28px;
        position: relative;
        overflow: hidden;
    }
    .help-hero::after {
        content: '';
        position: absolute;
        right: -40px; top: -40px;
        width: 240px; height: 240px;
        border-radius: 50%;
        background: radial-gradient(circle, rgba(91,141,239,0.15) 0%, transparent 70%);
        pointer-events: none;
    }
    .hero-eyebrow {
        font-size: 11px; font-weight: 700; letter-spacing: 2px;
        text-transform: uppercase; color: #3B9EFF; margin-bottom: 10px;
    }
    .hero-title {
        font-size: 30px; font-weight: 800; color: #fff;
        line-height: 1.15; margin-bottom: 10px;
    }
    .hero-title em { color: #5B8DEF; font-style: italic; }
    .hero-sub {
        font-size: 14px; color: rgba(255,255,255,0.6);
        max-width: 560px; line-height: 1.65; margin-bottom: 20px;
    }
    .hero-pills { display: flex; gap: 10px; flex-wrap: wrap; }
    .hpill {
        display: inline-block; padding: 5px 14px; border-radius: 20px;
        font-size: 11.5px; font-weight: 600;
    }
    .hpill-mint  { background: rgba(59,158,255,0.15); color: #3B9EFF; border: 1px solid rgba(59,158,255,0.35); }
    .hpill-sky   { background: rgba(91,141,239,0.15);  color: #5B8DEF;  border: 1px solid rgba(91,141,239,0.3); }
    .hpill-white { background: rgba(255,255,255,0.08); color: rgba(255,255,255,0.6); border: 1px solid rgba(255,255,255,0.15); }

    /* ── SECTION LABELS ── */
    .section-eyebrow {
        font-size: 10.5px; font-weight: 700; letter-spacing: 1.8px;
        text-transform: uppercase; color: #8494B0; margin-bottom: 14px;
        margin-top: 4px;
    }

    /* ── START STEPS ── */
    .start-grid {
        display: grid; grid-template-columns: repeat(4,1fr); gap: 12px;
        margin-bottom: 4px;
    }
    .start-step {
        background: #F4F7FF; border: 1px solid #DDE4F5;
        border-radius: 10px; padding: 16px 14px;
    }
    .ss-num  { font-size: 22px; font-weight: 800; color: #2E5CE6; margin-bottom: 6px; }
    .ss-title { font-size: 12.5px; font-weight: 700; color: #1A2340; margin-bottom: 4px; }
    .ss-desc  { font-size: 11.5px; color: #8494B0; line-height: 1.55; }

    /* ── UPLOADS TABLE ── */
    .uploads-tbl { width: 100%; border-collapse: collapse; font-size: 13px; font-family: 'Manrope', sans-serif; }
    .uploads-tbl thead tr { background: #F4F7FF; }
    .uploads-tbl thead th {
        padding: 10px 14px; text-align: left; font-size: 10.5px; font-weight: 700;
        text-transform: uppercase; letter-spacing: 1px; color: #8494B0;
        border-bottom: 1px solid #DDE4F5;
    }
    .uploads-tbl tbody td {
        padding: 11px 14px; border-bottom: 1px solid #DDE4F5;
        color: #1A2340; vertical-align: top; line-height: 1.5;
    }
    .uploads-tbl tbody tr:last-child td { border-bottom: none; }

    /* ── TOOL CARDS ── */
    .tool-card {
        border: 1px solid #DDE4F5; border-radius: 10px;
        overflow: hidden; margin-bottom: 14px;
        background: #fff;
    }
    .tc-header {
        padding: 16px 20px 14px;
        border-bottom: 1px solid #DDE4F5;
        display: flex; align-items: flex-start; gap: 12px;
    }
    .tc-icon { font-size: 24px; line-height: 1; flex-shrink: 0; }
    .tc-title { font-size: 15px; font-weight: 700; color: #1A2340; margin-bottom: 3px; }
    .tc-sub   { font-size: 12px; color: #8494B0; line-height: 1.5; }
    .tc-body  {
        padding: 16px 20px;
        display: grid; grid-template-columns: 1fr 1fr; gap: 20px;
    }
    .tc-body.single { grid-template-columns: 1fr; }
    .col-lbl {
        font-size: 10px; font-weight: 700; text-transform: uppercase;
        letter-spacing: 1.2px; color: #8494B0; margin-bottom: 9px;
    }

    /* Steps */
    .steps-list { list-style: none; padding: 0; margin: 0; }
    .steps-list li {
        display: flex; align-items: flex-start; gap: 9px;
        margin-bottom: 8px; font-size: 12.5px; color: #1A2340; line-height: 1.55;
    }
    .sn {
        width: 19px; height: 19px; border-radius: 50%;
        background:#050D1F; color: #fff; font-size: 9.5px; font-weight: 700;
        display: flex; align-items: center; justify-content: center;
        flex-shrink: 0; margin-top: 1px;
    }

    /* Why */
    .why-list { list-style: none; padding: 0; margin: 0; }
    .why-list li {
        font-size: 12.5px; color: #1A2340; line-height: 1.6;
        padding-left: 18px; position: relative; margin-bottom: 7px;
    }
    .why-list li::before { content: '✓'; position: absolute; left: 0; color: #3B9EFF; font-weight: 700; font-size: 11px; }

    /* Tip */
    .tip-box {
        background: rgba(240,165,0,0.07); border-left: 3px solid #F0A500;
        border-radius: 0 8px 8px 0; padding: 10px 13px;
        font-size: 12px; color: #1A2340; line-height: 1.55; margin-top: 10px;
    }

    /* ── AUDIENCE TAGS ── */
    .atag {
        display: inline-block; font-size: 10px; font-weight: 700;
        padding: 2px 9px; border-radius: 10px; letter-spacing: .4px;
        text-transform: uppercase; margin-left: 8px; vertical-align: middle;
    }
    .atag-c { background: rgba(46,92,230,0.1);  color: #2E5CE6; }
    .atag-m { background: rgba(59,158,255,0.12); color: #1a9a80; }
    .atag-a { background: rgba(139,100,220,0.1); color: #6B3FA0; }

    /* ── PHASE ROADMAP ── */
    .phase-row { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 12px; }
    .pdot {
        width: 28px; height: 28px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-size: 11px; font-weight: 700; flex-shrink: 0;
    }
    .pdot-now  { background: #2E5CE6; color: #fff; }
    .pdot-next { background: #DDE4F5; color: #8494B0; }
    .pt strong { display: block; font-size: 12.5px; font-weight: 700; color: #1A2340; margin-bottom: 2px; }
    .pt span   { font-size: 11.5px; color: #8494B0; line-height: 1.5; }

    /* ── DIVIDER ── */
    .help-divider { border: none; border-top: 1px solid #DDE4F5; margin: 28px 0 24px; }

    /* ── FOOTER ── */
    .help-footer {
        text-align: center; padding: 20px;
        font-size: 12px; color: #8494B0; line-height: 1.7;
        background: #F4F7FF; border-radius: 10px; margin-top: 16px;
    }
    .help-footer strong { color: #1A2340; }
</style>
""", unsafe_allow_html=True)


# ── HERO ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="help-hero">
    <div class="hero-eyebrow">Professional Services · Internal Tools</div>
    <div class="hero-title">PS Tools — <em>How-To Guide</em></div>
    <div class="hero-sub">Everything you need to get up and running — whether you're a consultant managing your projects or a manager keeping an eye on the team.</div>
    <div class="hero-pills">
        <span class="hpill hpill-mint">✦ Consultants</span>
        <span class="hpill hpill-sky">✦ Managers</span>
        <span class="hpill hpill-white">2026</span>
    </div>
</div>
""", unsafe_allow_html=True)


# ── GETTING STARTED ───────────────────────────────────────────────────────────
st.markdown('<div class="section-eyebrow">Before you begin · Everyone</div>', unsafe_allow_html=True)

st.markdown("""
<div style="background:#050D1F; border-radius:10px; padding:24px 28px; margin-bottom:16px; position:relative; overflow:hidden;">
<svg style='position:absolute;right:-40px;top:50%;transform:translateY(-50%);opacity:0.06;width:200px;height:200px;pointer-events:none' viewBox='0 0 1482 1286.25' xmlns='http://www.w3.org/2000/svg'><g fill='#3B9EFF' fill-rule='evenodd'><path d='M975.127,924.953c2.608-2.68,1.744-5.496-.42-7.829l-57.415-61.872c-2.463-2.655-5.025-2.878-8.443-.991-10.398,5.739-19.024,12.314-27.949,19.885-83.252,70.621-197.471,155.494-298.93,195.556-17.993,7.105-35.256,13.178-54.191,17.329-62.148,13.627-131.853,15.491-192.702-5.298-64.93-22.183-113.878-68.722-142.715-130.542-28.647-61.415-22.393-131.406,11.352-189.217,2.598-2.793,1.405-6.055-1.389-8.184-35.341-26.918-40.303-33.439-69.367-65.686-1.449-1.607-4.102-2.401-5.903-1.138-13.105,9.189-23.232,20.534-33.172,32.961-16.499,20.629-29.73,42.605-38.718,67.541-5.127,10.469-8.378,20.486-10.885,32.065-13.633,62.973-7.701,128.685,17.402,188.142,23.839,56.463,65.297,103.638,114.77,139.169,32.418,23.283,66.848,42.548,103.476,58.385,25.142,10.871,50.281,18.994,76.934,25.12,96.392,22.153,188.876,4.496,276.774-38.393,42.916-20.94,83.188-45.685,121.922-73.568,75.733-54.514,154.643-126.72,219.571-193.435ZM1445.252,792.261c-7.628-38.507-22.817-74.472-43.124-107.897-35.582-58.566-85.801-106.77-139.329-149.092-69.784-55.176-145.355-102.407-225.163-141.162-2.165-1.052-4.941.388-5.391,1.627-.426,1.171-.463,3.413.931,4.628,20.341,17.734,39.847,35.55,58.599,55.093,13.286,14.465,26.223,28.012,37.022,44.544,19.784,30.289,35.735,62.168,50.127,95.397,34.512,31.926,64.863,67.358,90.813,106.359,42.427,63.765,57.696,142.663,37.453,217.116-11.436,42.061-34.763,80.507-64.388,112.265-55.859,59.882-133.144,94.711-214.71,99.157-32.507,1.773-64.093-.538-96.013-6.503-28.16-5.262-70.299-23.997-96.538-36.626-2.312-1.112-4.605-.743-6.449.974-12.635,11.76-25.076,22.901-39.051,33.146l-43.32,31.757c-2.68,1.965-2.195,5.562.439,7.808,70.707,60.309,165.779,100.179,259.837,97.033,39.996-1.336,78.686-6.594,117.486-16.111,94.178-23.099,174.952-71.91,236.526-146.957,23.873-29.096,44.355-60.51,59.779-94.956,29.172-65.148,38.357-137.461,24.463-207.601ZM601.099,242.903c-12.268,10.522-48.215,44.405-47.219,60.482.993,16.01,10.781,31.195,25.227,38.155,14.47,6.972,41.303-10.055,53.886-18.311l65.495-42.972c26.305-17.259,52.496-32.716,80.08-47.834l57.464-31.494c20.451-11.209,41.123-19.851,63.235-27.448,35.852-12.318,72.313-18.084,110.322-17.747,29.787.263,58.398,3.408,86.939,11.449,44.037,12.405,82.745,35.987,114.027,69.974,20.347,22.106,37.598,45.332,51.026,71.732,6.962,13.688,13.008,27.156,16.103,42.311,6.48,31.729,12.267,85.992-.676,115.916-6.013,13.902-13.009,26.627-18.289,40.753-.847,2.264-.768,4.767,1.387,6.461l81.366,63.967c2.003,1.574,5.098.298,6.46-1.592,19.285-26.745,34.599-55.578,45.667-86.804,10.617-29.953,15.416-60.246,15.218-92.192-.482-77.938-29.055-152.791-79.976-211.891-67.16-77.946-169.264-137.487-272.877-146.244-33.524-2.834-66.192-1.328-99.421,3.091-82.214,10.934-149.21,45.218-216.385,92.267-48.269,33.807-94.373,69.644-139.062,107.973ZM72.687,567.553c20.03,44.974,54.35,86.652,88.718,121.568,19.447,19.756,38.882,38.258,60.393,55.711l73.052,59.268c30.921,25.086,74.954,56.331,111.096,72.278,11.713,5.168,23.385,8.99,35.917,11.295,12.922,2.375,24.878,1.136,37.309-3.088,18.441-6.266,35.538-14.698,52.671-24.006,1.792-.974,2.85-2.213,3.058-3.936.179-1.483-.47-3.163-1.914-4.548-14.129-13.542-27.174-27.284-42.195-40.056l-78.193-66.48-93.5-82.422c-23.176-20.43-44.471-41.737-65.536-64.239-15.19-16.227-28.591-32.64-40.05-51.639-20.601-34.157-31.396-72.282-30.182-112.398.614-20.279,2.364-39.861,7.45-59.369,8.872-34.031,50.72-76.652,77.451-99.125,3.767-7.04,2.459-14.401,2.885-21.735.884-15.227,3.244-29.908,5.647-44.959,4.285-26.824,22.718-58.984,38.899-80.638,1.348-1.805,1.936-3.535.891-4.937-.951-1.277-2.618-2.49-4.589-2.222-52.436,7.145-104.92,34.806-146.088,67.704-25.632,20.484-48.458,43.456-68.934,69.137-46.339,58.118-62.952,131.49-53.428,204.864,4.697,36.186,14.376,70.75,29.171,103.971ZM1196.886,310.029c-4.882-10.39-12.371-18.773-20.659-26.723-18.771-18.007-40.425-31.674-64.291-42.362-57.569-25.783-110.906-28.064-173.214-22.213-61.067,5.735-111.183,25.069-164.567,54.081-24.678,13.412-48.301,26.866-71.885,42.28l-105.247,68.787c-85.308,55.756-195.138,156.138-256.755,237.876-1.598,2.12-2.206,4.81-.222,6.912l76.342,80.886c1.468,1.556,2.9,1.672,4.715,1.249,1.397-.326,1.99-1.717,2.793-3.377,3.117-6.44,6.665-11.977,11.238-17.864,38.52-49.59,82.099-94.54,130.222-135.261,40.87-34.583,82.783-67.442,126.68-98.902,83.71-59.991,188.529-115.793,291.15-127.921,23.653-2.795,46.328-.575,69.656,3.405,27.197,4.641,52.661,12.543,78.69,21.347l38.004,12.855c13.849,4.685,27.221-3.226,30.503-17.755,2.725-12.064,2.293-25.708-3.154-37.301Z'/></g></svg>
<div style='position:relative;z-index:1'>
    <div style="font-size:16px; font-weight:700; color:#fff; margin-bottom:18px; font-family:Manrope,sans-serif;">
        Getting started in 4 steps
    </div>
    <div class="start-grid">
        <div class="start-step">
            <div class="ss-num">1</div>
            <div class="ss-title">Log in</div>
            <div class="ss-desc">Your role (Consultant or Manager) is set automatically from your credentials.</div>
        </div>
        <div class="start-step">
            <div class="ss-num">2</div>
            <div class="ss-title">Go to Daily Briefing</div>
            <div class="ss-desc">The Daily Briefing page is your upload hub. Drop your data exports here.</div>
        </div>
        <div class="start-step">
            <div class="ss-num">3</div>
            <div class="ss-title">Upload your files</div>
            <div class="ss-desc">Smartsheet DRS, Salesforce, and/or NetSuite exports. Each tool tells you what it needs.</div>
        </div>
        <div class="start-step">
            <div class="ss-num">4</div>
            <div class="ss-title">Navigate freely</div>
            <div class="ss-desc">Every page picks up your uploaded data. Your view is filtered to your projects.</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<hr class="help-divider">', unsafe_allow_html=True)


# ── CONSULTANT TOOLS ──────────────────────────────────────────────────────────
st.markdown('<div class="section-eyebrow">Consultant tools</div>', unsafe_allow_html=True)

CONSULTANT_TOOLS = [
    {
        "icon": "☀️",
        "title": "Daily Briefing",
        "sub": "Start here every morning. Your prioritized view of what needs attention today.",
        "steps": [
            "Log in and make sure your DRS is uploaded on the Home page.",
            "Navigate to Daily Briefing — it loads automatically filtered to your name.",
            "Review flagged items: overdue milestones, at-risk projects, stale activity.",
            "Use the links to jump directly to the relevant tool for any item.",
        ],
        "why": [
            "Surfaces the highest-priority items without scanning every project.",
            "Consistent — same logic every day, no missed flags.",
        ],
        "tip": "Managers can use the 'View As' dropdown in the sidebar to pull any consultant's daily briefing.",
    },
    {
        "icon": "📁",
        "title": "My Projects",
        "sub": "Your full portfolio — milestones, health flags, and needs-action items. All in one place.",
        "steps": [
            "Navigate to My Projects. Your active projects load automatically.",
            "Review the snapshot metrics at the top — active projects, overdue milestones, at-risk flags.",
            "Expand any project to see milestone detail, phase status, and days since last activity.",
            "Review on-hold projects at the bottom to keep track of paused work.",
        ],
        "why": [
            "Gives each consultant a focused, personal view of their own portfolio.",
            "Surfaces milestone gaps and inactivity in one glance — no scrolling through a shared sheet.",
            "Works alongside Smartsheet DRS as an interactive layer on top of your existing data.",
        ],
        "tip": "On-hold projects show separately at the bottom — don't lose track of them.",
    },
    {
        "icon": "✉️",
        "title": "Customer Engagement",
        "sub": "Your hub for customer outreach — initial introductions, re-engagement, and future lifecycle communications.",
        "steps": [
            "Make sure SS DRS, SFDC Contacts, and NS Time Detail exports are uploaded on Home.",
            "Navigate to Customer Engagement — your projects load, filtered to you.",
            "Check 'This Week\'s Initial Engagement Actions' — projects with no intro email sent yet.",
            "Check 'This Week\'s Re-Engagement Actions' for stalled projects needing follow-up.",
            "Select a project, review the suggested tier and template, edit as needed, and send.",
        ],
        "why": [
            "Phase 1: Surfaces projects missing an intro email so no new customer falls through the cracks.",
            "Phase 1: Tier 1–4 re-engagement templates — tone escalates with days inactive.",
            "Pre-populated with project name, product, and contact details from SFDC.",
            "Phase 2 (coming soon): Project lifecycle templates — kick-off, go-live, and close communications.",
        ],
        "tip": "You only see your own projects here. Managers: use 'View As' in the sidebar to review any consultant or region.",
    },
]

for tool in CONSULTANT_TOOLS:
    steps_html = "".join(f'<li><span class="sn">{i+1}</span>{s}</li>' for i, s in enumerate(tool["steps"]))
    why_html   = "".join(f"<li>{w}</li>" for w in tool["why"])
    tip_html   = f'<div class="tip-box">💡 <strong>Tip:</strong> {tool["tip"]}</div>' if tool.get("tip") else ""
    st.markdown(
        f'<div class="tool-card">'
        f'<div class="tc-header"><div class="tc-icon">{tool["icon"]}</div>'
        f'<div><div class="tc-title">{tool["title"]}</div><div class="tc-sub">{tool["sub"]}</div></div></div>'
        f'<div class="tc-body">'
        f'<div><div class="col-lbl">How to use it</div><ul class="steps-list">{steps_html}</ul></div>'
        f'<div><div class="col-lbl">Why it helps</div><ul class="why-list">{why_html}</ul>{tip_html}</div>'
        f'</div></div>',
        unsafe_allow_html=True
    )

st.markdown('<hr class="help-divider">', unsafe_allow_html=True)


# ── REPORTING TOOLS ──────────────────────────────────────────────────────────
st.markdown('<div class="section-eyebrow">Reporting tools</div>', unsafe_allow_html=True)

REPORTING_TOOLS = [
    {
        "icon": "📊",
        "title": "Utilization Report",
        "sub": "Consistent utilization credit scoring across T&M and Fixed Fee engagements.",
        "steps": [
            "Upload the NetSuite time detail export on the Home page.",
            "Navigate to Utilization Report and select the reporting period.",
            "Review individual and team-level utilization credit scores.",
            "Export the report for leadership or payroll review.",
        ],
        "why": [
            "T&M hours score 1:1. Fixed Fee hours score up to scoped hours only — overrun hours flagged separately.",
            "Same methodology every period — no manual adjustments, no interpretive drift.",
            "Delivery variance and overruns are surfaced, not buried.",
        ],
        "tip": "Overrun hours appear separately as 'delivery variance' — not counted against a consultant's utilization, but flagged for project review.",
    },
    {
        "icon": "❤️‍🔥",
        "title": "Workload Health Score",
        "sub": "Who's stretched, who has room — surfaced before it becomes a staffing problem.",
        "steps": [
            "Upload NetSuite and SS DRS exports on Home.",
            "Navigate to Workload Health Score.",
            "Review team-level health indicators — project load, active milestones, overdue flags per consultant.",
            "Use filters to drill into a region or product area.",
        ],
        "why": [
            "Aggregates signals from multiple sources into a single health view.",
            "Catches overload before a consultant raises a hand — or before a project slips.",
            "Useful for resourcing conversations: 'who can take the next project?'",
        ],
        "tip": None,
    },
    {
        "icon": "🔍",
        "title": "DRS Health Check",
        "sub": "Catches data quality issues in Smartsheet DRS before they surface in reporting.",
        "steps": [
            "Upload SS DRS export on Home (or directly on this page).",
            "Navigate to DRS Health Check.",
            "Review flagged rows — errors (red), warnings (amber), and info flags (blue).",
            "Fix issues in Smartsheet, re-export, and re-run to confirm clean.",
        ],
        "why": [
            "Flags logical inconsistencies — e.g. a project marked 'complete' with open milestones.",
            "Catches missing fields that break downstream reporting before they reach an exec deck.",
            "Severity-tiered — errors vs. warnings vs. informational, so you know what to fix first.",
        ],
        "tip": "Run this before any leadership reporting cycle. It takes 30 seconds and saves a lot of awkward 'why is this row blank' conversations.",
    },
]

for tool in REPORTING_TOOLS:
    steps_html = "".join(f'<li><span class="sn">{i+1}</span>{s}</li>' for i, s in enumerate(tool["steps"]))
    why_html   = "".join(f"<li>{w}</li>" for w in tool["why"])
    tip_html   = f'<div class="tip-box">💡 <strong>Tip:</strong> {tool["tip"]}</div>' if tool.get("tip") else ""
    st.markdown(
        f'<div class="tool-card">'
        f'<div class="tc-header"><div class="tc-icon">{tool["icon"]}</div>'
        f'<div><div class="tc-title">{tool["title"]}</div><div class="tc-sub">{tool["sub"]}</div></div></div>'
        f'<div class="tc-body">'
        f'<div><div class="col-lbl">How to use it</div><ul class="steps-list">{steps_html}</ul></div>'
        f'<div><div class="col-lbl">Why it helps</div><ul class="why-list">{why_html}</ul>{tip_html}</div>'
        f'</div></div>',
        unsafe_allow_html=True
    )




# ── MANAGEMENT TOOLS ─────────────────────────────────────────────────────────
st.markdown('<div class="section-eyebrow">Management tools</div>', unsafe_allow_html=True)

MGMT_TOOLS = [
    {
        "icon": "🔭",
        "title": "Capacity Outlook",
        "sub": "Forward-looking headcount signal for pipeline and resourcing decisions.",
        "steps": [
            "Upload NetSuite and SS DRS exports on Home.",
            "Navigate to Capacity Outlook and select the forecast window.",
            "Review projected availability by consultant and region.",
        ],
        "why": [
            "Combines current project data with expected close dates to project future bandwidth.",
            "Helps PS and Sales align on when the team can absorb new bookings.",
            "Consultants in learning/training are surfaced as future capacity, not current.",
        ],
        "tip": None,
    },
]

for tool in MGMT_TOOLS:
    steps_html = "".join(f'<li><span class="sn">{i+1}</span>{s}</li>' for i, s in enumerate(tool["steps"]))
    why_html   = "".join(f"<li>{w}</li>" for w in tool["why"])
    tip_html   = f'<div class="tip-box">💡 <strong>Tip:</strong> {tool["tip"]}</div>' if tool.get("tip") else ""
    st.markdown(
        f'<div class="tool-card">'
        f'<div class="tc-header"><div class="tc-icon">{tool["icon"]}</div>'
        f'<div><div class="tc-title">{tool["title"]}</div><div class="tc-sub">{tool["sub"]}</div></div></div>'
        f'<div class="tc-body">'
        f'<div><div class="col-lbl">How to use it</div><ul class="steps-list">{steps_html}</ul></div>'
        f'<div><div class="col-lbl">Why it helps</div><ul class="why-list">{why_html}</ul>{tip_html}</div>'
        f'</div></div>',
        unsafe_allow_html=True
    )

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="help-footer">
    <strong>PS Tools</strong> · Built by PS, for PS · 2026<br>
    Questions about the tool? Reach out to your PS manager.<br>
    Found a bug? We want to know — this thing only gets better with feedback.
</div>
""", unsafe_allow_html=True)
