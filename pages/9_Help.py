"""
PS Tools — Reference Guide
"""
import streamlit as st

st.session_state["current_page"] = "Help"

st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
<style>
    html,body,[class*="css"]{font-family:'Manrope',sans-serif!important}
    .divider{border:none;border-top:1px solid rgba(128,128,128,.2);margin:24px 0}
    .page-card{border:0.5px solid rgba(128,128,128,.2);border-radius:10px;padding:20px 24px;margin-bottom:14px}
    .page-title{font-size:15px;font-weight:700;margin-bottom:4px}
    .page-section{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;opacity:.45;margin-bottom:10px}
    .page-desc{font-size:13px;opacity:.8;line-height:1.6;margin-bottom:12px}
    .meta-row{display:flex;gap:24px;flex-wrap:wrap;margin-top:10px}
    .meta-item{font-size:11px;color:inherit;opacity:.55;line-height:1.5}
    .meta-label{font-weight:700;opacity:.7;display:block}
    .pill-data{display:inline-block;font-size:10px;font-weight:700;padding:1px 7px;border-radius:8px;background:rgba(59,158,255,.1);color:#3B9EFF;letter-spacing:.5px;margin-right:4px}
    .pill-role{display:inline-block;font-size:10px;font-weight:700;padding:1px 7px;border-radius:8px;background:rgba(128,128,128,.1);opacity:.7;letter-spacing:.5px;margin-right:4px}
    .section-label { font-size: 13px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#4472C4;margin-bottom:8px}
</style>
""", unsafe_allow_html=True)

_zone_svg = """"""

_hero = st.empty()
_hero.markdown(
    f"""
<div style='background:#050D1F;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;position:relative;overflow:hidden;font-family:Manrope,sans-serif'>
  {_zone_svg}
  <div style='font-size:10px;font-weight:700;letter-spacing:2.5px;text-transform:uppercase;color:#3B9EFF;margin-bottom:10px'>Professional Services · Tools</div>
  <h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif'>Reference Guide</h1>
  <p style='color:rgba(255,255,255,0.45);margin:6px 0 0;font-size:14px;font-family:Manrope,sans-serif'>What each page does, what data it uses, and what questions it answers.</p>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown('<div class="section-label">Data Sources</div>', unsafe_allow_html=True)
st.markdown("""
<div style='font-size:12px;opacity:.7;line-height:1.9;margin-bottom:20px'>
  <span class='pill-data'>SS DRS</span> Smartsheet DRS export — project list, phases, milestones, RAG, on-hold status<br>
  <span class='pill-data'>NS Time</span> NetSuite Time Detail export — hours logged per project per consultant<br>
  <span class='pill-data'>SFDC</span> Salesforce export — opportunity and account data (Revenue Report only)
</div>
""", unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

def _card(title, section, desc, data_sources, audience, answers, limitations=None):
    _ds  = "".join(f"<span class='pill-data'>{d}</span>" for d in data_sources)
    _aud = "".join(f"<span class='pill-role'>{a}</span>" for a in audience)
    _ans = "".join(f"<li style='margin-bottom:3px'>{a}</li>" for a in answers)
    _lim = f"<div style='font-size:11px;opacity:.4;margin-top:10px;font-style:italic'>{limitations}</div>" if limitations else ""
    st.markdown(f"""<div class='page-card'>
      <div class='page-section'>{section}</div>
      <div class='page-title'>{title}</div>
      <div class='page-desc'>{desc}</div>
      <ul style='margin:0 0 10px 16px;font-size:12px;opacity:.75;line-height:1.7'>{_ans}</ul>
      <div class='meta-row'>
        <div class='meta-item'><span class='meta-label'>Data needed</span>{_ds}</div>
        <div class='meta-item'><span class='meta-label'>Audience</span>{_aud}</div>
      </div>{_lim}
    </div>""", unsafe_allow_html=True)

st.markdown('<div class="section-label">My Tools</div>', unsafe_allow_html=True)

_card("Daily Briefing", "My Tools · Home page",
    "Your daily starting point. Shows utilization pacing for the current month, project risk signals, active project snapshot, and a natural-language briefing summarising what needs attention today.",
    ["SS DRS","NS Time"], ["Consultant","Manager"],
    ["Am I on pace for my utilization target this month?",
     "Which of my projects are Red or Yellow RAG?",
     "Do I have any projects going live this week or in hypercare?",
     "Are there projects missing intro emails or needing re-engagement?",
     "What phase are my active projects in, and are any 9+ or 12+ months old?"],
    "Utilization colour uses a pro-rated daily pacing target, not a flat monthly comparison. Briefing text is rule-based — AI summaries are on the roadmap.")

_card("My Projects", "My Tools",
    "Your working project list. View all active and on-hold projects, edit key fields (RAG, phase, on-hold reason, milestones), and export changes to sync back to Smartsheet DRS.",
    ["SS DRS","NS Time"], ["Consultant","Manager"],
    ["What is the current status of each of my projects?",
     "Which projects are on hold and why?",
     "What is the hours balance remaining on each Fixed Fee project?",
     "Which milestones are missing or overdue?"],
    "Edits are session-only until exported and synced to Smartsheet. Balance calculations require NS Time Detail.")

_card("Project Health", "My Tools",
    "Delivery performance signals for your active portfolio. Shows schedule variance against go-live dates, milestone gaps by type, and scope burn rate with phase-adjusted warnings.",
    ["SS DRS","NS Time"], ["Consultant","Manager"],
    ["Which projects are delayed or at risk of missing their go-live date?",
     "What milestones are missing past their expected phase?",
     "Which projects are overrunning or close to their scope budget?",
     "Is a project burning hours faster than its phase would suggest?"],
    "Schedule uses Est. Go-Live from DRS. Once Original Go-Live Date and Forecast Go-Live Date are added to DRS, true slippage tracking will be available. Approvals, CC, SFTP, and Additional Subsidiary are excluded from phase-burn scope warnings.")

_card("Customer Engagement", "My Tools",
    "Outreach assistant for stale or at-risk projects. Groups projects by contact urgency and provides customisable re-engagement message templates.",
    ["SS DRS","NS Time"], ["Consultant","Manager"],
    ["Which customers haven't had contact in 14+ days?",
     "What should I say in a re-engagement message?",
     "Who is the right contact person for this project?"],
    "Templates are rule-based by inactivity tier. Gmail/Outlook send integration and outreach logging are on the roadmap.")

_card("Utilization Report", "My Tools",
    "Detailed monthly utilization breakdown by project. Calculates Fixed Fee credited hours (scope-capped), T&M hours, overrun, and internal time. Exportable to branded Excel.",
    ["NS Time"], ["Consultant","Manager"],
    ["How many credited hours did I log this month by project?",
     "Which Fixed Fee projects have overrun their scope?",
     "What is my utilization % for a given period?",
     "Which projects have no scope defined?"],
    "Scope is matched by project type against DEFAULT_SCOPE. Premium projects use the Time Item SKU (IMPL10/IMPL20). Requires NS Time Detail — SS DRS optional but improves project name matching.")

_card("Workload Health Score", "My Tools",
    "WHS scoring for your project portfolio. Weighted risk score per project based on phase, RAG, inactivity, milestone gaps, and go-live proximity. Aggregates to a per-consultant score.",
    ["SS DRS","NS Time"], ["Consultant","Manager"],
    ["What is my overall workload health score?",
     "Which projects are contributing most to my workload risk?",
     "How does my score compare across the team? (manager view)"],
    "WHS is a relative signal, not an absolute performance measure. Scores are recalculated on each page load.")

_card("DRS Health Check", "My Tools",
    "Data quality audit for the SS DRS. Flags missing or inconsistent fields across active projects — go-live dates, on-hold fields, milestone gaps, and RAG inconsistencies.",
    ["SS DRS"], ["Consultant","Manager"],
    ["Which projects have missing or incomplete data in the DRS?",
     "Are on-hold projects missing a reason or responsible party?",
     "Are there RAG or sentiment values inconsistent with on-hold status?",
     "Which projects are missing key milestone dates?"],
    "Flags are rule-based and may surface false positives on legacy projects where field completion was not required historically.")

_card("Time Entries", "My Tools",
    "Browse NS time entries filtered to a specific project from your DRS portfolio. Useful for auditing logged hours or investigating utilization discrepancies.",
    ["SS DRS","NS Time"], ["Consultant","Manager"],
    ["What time entries have been logged against a specific project?",
     "Who has logged time on a project and when?"],
    "Read-only. Time entry editing must be done directly in NetSuite.")

st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Management</div>', unsafe_allow_html=True)

_card("Portfolio Analytics", "Management",
    "Team-level portfolio view for managers. Aggregates project risk, schedule health, scope health, and utilization across all consultants. Respects View As for individual consultant or region views.",
    ["SS DRS","NS Time"], ["Manager"],
    ["How many active and on-hold projects does each consultant have?",
     "Which consultants have Red RAG, delayed, or overrun projects?",
     "How is utilization distributed across the team this month?",
     "What products and phases make up the team's current portfolio?",
     "Which consultants have a high WHS score?"],
    "Util % matches Daily Briefing logic (T&M full credit + FF scope-capped). WHS requires SS DRS. NS Time Detail needed for utilization columns.")

_card("Capacity Outlook", "Management",
    "Forward-looking capacity planning. Maps team available hours against confirmed and pipeline bookings to identify over- and under-capacity periods by consultant and region.",
    ["SS DRS","SFDC"], ["Manager"],
    ["Which consultants are over-capacity in the next 4–8 weeks?",
     "Where is there available capacity to absorb new bookings?",
     "How does pipeline demand compare to available hours?"],
    "Requires both SS DRS and Salesforce exports. Pipeline data is point-in-time.")

_card("Revenue Report", "Management",
    "Fixed Fee and T&M revenue pipeline analysis. Shows contracted vs recognised revenue, SOW-to-NS project matching, FX adjustments, and carve-out SKU handling. Exportable to branded Excel.",
    ["SS DRS","NS Time","SFDC"], ["Manager"],
    ["What is the contracted value of the active FF and T&M pipeline?",
     "How much revenue has been recognised vs what remains?",
     "Which SOWs have not been matched to a NetSuite project?",
     "What is the FX-adjusted revenue by currency?"],
    "SOW-to-NS matching uses fuzzy name matching — unmatched projects appear in the Watch List tab. Requires all three data sources for full functionality.")

st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Loading Data</div>', unsafe_allow_html=True)
st.markdown("""
<div class='page-card'>
  <div class='page-desc'>All data is loaded from the sidebar on the Home page (Daily Briefing). Data is session-only — it clears when you close the browser or log out.</div>
  <div style='font-size:12px;line-height:1.9;margin-top:4px'>
    <b>SS DRS Export</b> — Export your full DRS Smartsheet as CSV or Excel. Required for all project-level pages.<br>
    <b>NS Time Detail</b> — NetSuite: Time › Time Detail report, filtered to the relevant period. Required for utilization metrics.<br>
    <b>Salesforce Export</b> — PS pipeline opportunity report. Required for Capacity Outlook and Revenue Report.
  </div>
  <div style='font-size:11px;opacity:.4;margin-top:12px;font-style:italic'>
    Tip: Load the DRS first — it enables View As so managers can switch between consultant and region views before loading NS data.
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div style="font-size:11px;opacity:.4;text-align:center">PS Projects &amp; Tools · Internal use only · Contact your team lead for access queries</div>', unsafe_allow_html=True)
