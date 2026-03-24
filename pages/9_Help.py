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
        background: #1B2B5E;
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
        text-transform: uppercase; color: #ff4b40; margin-bottom: 10px;
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
    .hpill-mint  { background: rgba(255,75,64,0.15); color: #ff4b40;  border: 1px solid rgba(255,75,64,0.35); }
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
        background: #1B2B5E; color: #fff; font-size: 9.5px; font-weight: 700;
        display: flex; align-items: center; justify-content: center;
        flex-shrink: 0; margin-top: 1px;
    }

    /* Why */
    .why-list { list-style: none; padding: 0; margin: 0; }
    .why-list li {
        font-size: 12.5px; color: #1A2340; line-height: 1.6;
        padding-left: 18px; position: relative; margin-bottom: 7px;
    }
    .why-list li::before { content: '✓'; position: absolute; left: 0; color: #ff4b40; font-weight: 700; font-size: 11px; }

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
    .atag-m { background: rgba(255,75,64,0.12); color: #1a9a80; }
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
<div style="background:#1B2B5E; border-radius:10px; padding:24px 28px; margin-bottom:16px;">
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
        "title": "Customer Re-engagement",
        "sub": "Generate tiered outreach letters for on-hold or stalled projects. No copy-paste required.",
        "steps": [
            "Make sure SS DRS and Salesforce exports are uploaded on Home.",
            "Navigate to Customer Re-engagement — your projects load, filtered to you.",
            "Select a project from the dropdown. The tool shows days inactive and suggests a tier.",
            "Review the suggested letter, edit as needed, and copy or download.",
        ],
        "why": [
            "Three tiers of outreach templates — tone scales with how long the project's been stalled.",
            "Pre-populated with project name, product, and contact details from your data.",
            "Consistent messaging across the team — no one reinventing the wheel per email.",
        ],
        "tip": "You only see your own projects here. Managers: use 'View As' in the sidebar to generate letters on behalf of a consultant.",
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
