"""
PS Tools — Template Utilities
Re-engagement email templates and placeholder helpers.
"""
import re
import streamlit as st

from shared.constants import MILESTONE_COLS_MAP


TEMPLATES = {
    "Tier 1 — ~30 Day Check-In": {
        "tier": 1,
        "days_min": 30,
        "days_max": 59,
        "subject": "{PRODUCT NAME} — Checking In on Your Project",
        "cc_guidance": "No additional CC required.",
        "body": """Hi {CUSTOMER CONTACT NAME},

I hope you're doing well. I wanted to reach out as it's been a little while since we last connected on your {PRODUCT NAME} implementation.

When we last spoke, we were in the {CURRENT PHASE} phase. I appreciate that priorities can shift and schedules get busy — I just wanted to check in and see if you're ready to pick things back up, or if anything has changed on your end that I should be aware of.

We're here whenever you're ready. If it's helpful, I'm happy to schedule a short call to realign on next steps and get things moving again.

Looking forward to hearing from you.

Kind regards,
{IMPLEMENTATION CONSULTANT}
Professional Services | Zone & Co"""
    },
    "Tier 2 — ~60 Day Follow-Up": {
        "tier": 2,
        "days_min": 60,
        "days_max": 89,
        "subject": "{PRODUCT NAME} Project — Let's Reconnect",
        "cc_guidance": "CC: PS Leadership",
        "body": """Hi {CUSTOMER CONTACT NAME},

I'm following up on my earlier message regarding your {PRODUCT NAME} implementation. Our last activity together was on {LAST ACTIVITY DATE}, during the {CURRENT PHASE} phase, and I want to make sure we can keep your project on track.

Where Things Stand

As outlined in the implementation scope shared at the start of the project, estimated timelines are based on mutual availability and timely engagement from both teams. Extended pauses can affect a few things I want to flag for your awareness:

• Consultant availability: Prolonged delays may mean that your currently assigned consultant is reassigned to other projects. We'll do our best to maintain continuity, but early re-engagement helps.
• Session utilization: Your implementation includes a defined number of working sessions and enablement activities. These are available through to project closure.
• Environment changes: If your Sandbox environment is refreshed during the pause, reconfiguration may be required, which could impact the remaining scope.

Suggested Next Step

I'd recommend we schedule a 30-minute call to review where we left off, confirm your current priorities, and agree on an updated timeline. Could you let me know your availability over the next week or two?

We're committed to helping you get the most out of {PRODUCT NAME} and want to ensure a smooth path to Go-Live.

Kind regards,
{IMPLEMENTATION CONSULTANT}
Professional Services | Zone & Co"""
    },
    "Tier 3 — ~90 Day Escalation": {
        "tier": 3,
        "days_min": 90,
        "days_max": 179,
        "subject": "{PRODUCT NAME} Implementation — Re-Engagement Needed",
        "cc_guidance": "CC: PS Leadership · Account Manager ({ACCOUNT MANAGER}) · CS Manager",
        "body": """Hi {CUSTOMER CONTACT NAME},

I'm reaching out once more regarding your {PRODUCT NAME} implementation. It's now been over 90 days since our last project activity on {LAST ACTIVITY DATE}, and I want to make sure we can find a way forward together.

Project Status

Your project is currently in the {CURRENT PHASE} phase. We've been holding this open on our side, but I want to be transparent about a few things that extended inactivity can affect:

• Consultant reassignment: After this length of time, your assigned consultant may no longer be available when you're ready to resume. We'll work to ensure a smooth handover if that's the case, but there may be some ramp-up time.
• Environment drift: If your NetSuite Sandbox has been refreshed or modified during this period, reconfiguration of {PRODUCT NAME} may be required. Depending on the extent, this could impact the remaining scope and deliverables.
• Knowledge continuity: The longer the gap, the more likely it is that key decisions or configurations from earlier phases need to be revisited, which can extend the overall timeline.

Recommended Next Steps

I've copied your Customer Success Manager, {ACCOUNT MANAGER}, on this email so we can coordinate a plan together. Specifically, it would be helpful to:

• Schedule a re-alignment call to review the current project status and outstanding deliverables.
• Confirm availability of your project team and NetSuite administrator.
• Agree on an updated timeline to get the project back on track.

We understand that circumstances change, and we're here to help find the right path forward. Please let us know how you'd like to proceed.

Kind regards,
{IMPLEMENTATION CONSULTANT}
Professional Services | Zone & Co"""
    },
    "Tier 4 — ~6 Month Notification": {
        "tier": 4,
        "days_min": 180,
        "days_max": 99999,
        "subject": "Important: {PRODUCT NAME} Implementation — Service Term Update",
        "cc_guidance": "CC: PS Leadership · Account Manager ({ACCOUNT MANAGER}) · CS Manager",
        "body": """Hi {CUSTOMER CONTACT NAME},

I'm writing to provide an important update regarding your {PRODUCT NAME} implementation. Despite our previous outreach, we've been unable to reconnect with your team, and I want to ensure you have full visibility into where things stand so we can plan the best path forward.

Project Status

Your project has been in the {CURRENT PHASE} phase since our last activity on {LAST ACTIVITY DATE}. You have {REMAINING SESSIONS} sessions remaining in scope.

Service Term & Scope Reminders

I'd like to bring a few items from our agreed implementation scope to your attention:

• Service term: Your Professional Services engagement is valid for 12 months from your contract signature date. Your current service term expires on {SERVICE TERM EXPIRY}.
• Session utilization: Any unused sessions or deliverables remaining at the time of transition to Support will be considered forfeited and no longer valid.
• Resource availability: Extended delays may impact the continued availability of your assigned consultant. It is the customer's responsibility to notify Zone when ready to re-engage.
• Environment changes: If a Sandbox refresh has occurred during this period, reconfiguration may be required and could impact remaining scope.

Recommended Next Steps

To make the most of your remaining sessions and service term, I'd strongly recommend we put together a re-engagement plan as soon as possible. I've copied your Account Manager, {ACCOUNT MANAGER}, on this email so we can coordinate together.

• Schedule a re-alignment call to review the current project status and outstanding deliverables.
• Confirm availability of your project team and NetSuite administrator.
• Agree on an updated timeline that allows us to complete the implementation within your service term.

We want to ensure you get full value from your investment in {PRODUCT NAME}. Please let us know how you'd like to proceed.

Kind regards,
{IMPLEMENTATION CONSULTANT}
Professional Services | Zone & Co"""
    },
}

TIER_COLORS = {1: "#EAF9F1", 2: "#FEF9E7", 3: "#FDECED", 4: "#f0e6ff"}
TIER_TEXT   = {1: "#1E8449", 2: "#9C6500", 3: "#C0392B", 4: "#6c3483"}



def suggest_tier(days_inactive):
    if days_inactive is None:
        return None
    for name, tmpl in TEMPLATES.items():
        if tmpl["days_min"] <= days_inactive <= tmpl["days_max"]:
            return name
    return list(TEMPLATES.keys())[-1]


def fill_template(template_key, fields):
    tmpl  = TEMPLATES[template_key]
    body  = tmpl["body"]
    subj  = tmpl["subject"]
    cc    = tmpl["cc_guidance"]

    for k, v in fields.items():
        if v:
            body = body.replace(f"{{{k}}}", str(v))
            subj = subj.replace(f"{{{k}}}", str(v))
            cc   = cc.replace(f"{{{k}}}", str(v))
    return subj, body, cc


def highlight_placeholders(text):
    """Return text with remaining {PLACEHOLDERS} wrapped in HTML spans."""
    def replacer(m):
        return f"<span class='placeholder-missing'>{m.group(0)}</span>"
    return re.sub(r'\{[A-Z _]+\}', replacer, text)


def extract_placeholders(text):
    return re.findall(r'\{([A-Z _]+)\}', text)

# ── Main ──────────────────────────────────────────────────────────────────────
# ── SS DRS column map ────────────────────────────────────────────────────────
SS_COL_MAP_OUT = {
    "project name":           "project_name",
    "name":                   "project_name",   # fallback: exports where header is just "Name"
    "project id":             "project_id",
    "project phase":          "phase",
    "project type":           "project_type",
    "status":                 "status",
    "start date":             "start_date",
    "start date (subscription)": "subscription_start_date",  # kept separate — not project start
    "go live date":           "go_live_date",
    "territory":              "territory",
    "billing type":           "billing_type",
    "billing":                "billing_type",
    "project manager":        "project_manager",
    "consultant":             "project_manager",
    "overall rag":            "rag",
    "schedule health":        "schedule_health",
    "risk level":             "risk_level",
    "client responsiveness":  "client_responsiveness",
    "last updated":           "last_updated",
    "modified":               "last_updated",
    "modified date":          "last_updated",
    "last modified":          "last_updated",
    "date modified":          "last_updated",
    "account name":           "account",
    "customer":               "account",
    "risk detail":            "risk_detail",
    "risk owner":             "risk_owner",
    "responsible for delay":  "responsible_delay",
    "delay summary":          "delay_summary",
    "intro. email sent":           "ms_intro_email",
    "standard config start":       "ms_config_start",
    "enablement session":          "ms_enablement",
    "session #1":                  "ms_session1",
    "session #2":                  "ms_session2",
    "uat signoff":                 "ms_uat_signoff",
    "prod cutover":                "ms_prod_cutover",
    "hypercare start":             "ms_hypercare_start",
    "close out remaining tasks":   "ms_close_out",
    "transition to support":       "ms_transition",
    # ── Financials & scope ────────────────────────────────────────────────────
    "actual hours":                "actual_hours",
    "budgeted hours":              "budgeted_hours",
    "change order":                "change_order",
    "change order hours":          "change_order",
    # ── Legacy flag ───────────────────────────────────────────────────────────
    "legacy":                      "legacy",
    # ── Client sentiment ──────────────────────────────────────────────────────
    "client sentiment":            "client_sentiment",
}

INACTIVE_PHASES_OUT = {
    "10. complete/pending final billing",
    "11. on hold",
    "12. ps review",
}


