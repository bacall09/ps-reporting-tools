"""
PS Tools — Daily Briefing
Consultant-scoped view: my projects, my utilization, my re-engagement actions.
Upload SS DRS + NS Time Detail to generate a personalized briefing.
"""
import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta

from shared.constants import (
    EMPLOYEE_ROLES, ACTIVE_EMPLOYEES, CONSULTANT_DROPDOWN,
    MILESTONE_COLS_MAP, get_role, is_manager, is_consultant,
)
from shared.config import (
    AVAIL_HOURS, EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP, DEFAULT_SCOPE,
)
from shared.loaders import (
    load_drs, load_ns_time, calc_days_inactive, calc_last_milestone,
    normalise_product_name, suggest_tier_from_days,
)
from shared.template_utils import TEMPLATES, suggest_tier

st.set_page_config(page_title="Daily Briefing", page_icon=None, layout="wide")

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        .brief-header { font-size: 22px; font-weight: 700; color: #1e2c63; margin-bottom: 4px; }
        .brief-sub    { font-size: 13px; color: #718096; margin-bottom: 24px; }
        .section-label { font-size: 11px; font-weight: 700; text-transform: uppercase;
                         letter-spacing: 0.8px; color: #4472C4; margin-bottom: 8px; }
        .metric-card  { background: #f7fafc; border: 1px solid #e2e8f0; border-radius: 8px;
                        padding: 16px 20px; margin-bottom: 12px; }
        .metric-val   { font-size: 28px; font-weight: 700; color: #1e2c63; }
        .metric-lbl   { font-size: 12px; color: #718096; margin-top: 2px; }
        .proj-card    { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px;
                        padding: 16px 20px; margin-bottom: 10px; }
        .proj-name    { font-size: 14px; font-weight: 700; color: #1e2c63; margin-bottom: 4px; }
        .proj-meta    { font-size: 12px; color: #718096; }
        .rag-R { display:inline-block; width:10px; height:10px; border-radius:50%;
                 background:#E74C3C; margin-right:6px; }
        .rag-A { display:inline-block; width:10px; height:10px; border-radius:50%;
                 background:#F39C12; margin-right:6px; }
        .rag-G { display:inline-block; width:10px; height:10px; border-radius:50%;
                 background:#27AE60; margin-right:6px; }
        .rag-  { display:inline-block; width:10px; height:10px; border-radius:50%;
                 background:#BDC3C7; margin-right:6px; }
        .action-badge { display:inline-block; font-size:11px; font-weight:600;
                        padding:2px 8px; border-radius:4px; margin-right:6px; }
        .badge-red    { background:#FDECED; color:#C0392B; }
        .badge-amber  { background:#FEF9E7; color:#D68910; }
        .badge-blue   { background:#EBF0FB; color:#2E4DA7; }
        .badge-gray   { background:#F2F2F2; color:#5D6D7E; }
        .divider      { border: none; border-top: 1px solid #e2e8f0; margin: 20px 0; }
        .ai-box       { background: #f0f4ff; border: 1px solid #c3d1f5; border-radius: 8px;
                        padding: 16px 20px; margin-top: 12px; font-size: 13px; line-height: 1.7; }
        .ai-label     { font-size: 11px; font-weight: 700; text-transform: uppercase;
                        letter-spacing: 0.8px; color: #4472C4; margin-bottom: 8px; }
        @media (prefers-color-scheme: dark) {
            .brief-header { color: #c5d0f0; }
            .metric-card  { background: #1e1e2e; border-color: #2d2d44; }
            .metric-val   { color: #c5d0f0; }
            .proj-card    { background: #1e1e2e; border-color: #2d2d44; }
            .proj-name    { color: #c5d0f0; }
            .ai-box       { background: #1a2040; border-color: #2d3f6e; }
        }
    </style>
""", unsafe_allow_html=True)

# ── Role gate ─────────────────────────────────────────────────────────────────
if "consultant_name" not in st.session_state:
    st.session_state["consultant_name"] = None

# ── Sidebar: who are you? ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("#### My Briefing")

    # Name selector — consultants + manager-consultants + managers-only
    all_names = sorted(CONSULTANT_DROPDOWN)
    selected = st.selectbox(
        "Select your name",
        options=["— Select —"] + all_names,
        index=0 if not st.session_state["consultant_name"]
              else (["— Select —"] + all_names).index(st.session_state["consultant_name"])
              if st.session_state["consultant_name"] in all_names else 0,
        key="briefing_name_select",
    )
    if selected != "— Select —":
        st.session_state["consultant_name"] = selected

    # Managers can also browse another consultant's view
    role = get_role(selected) if selected != "— Select —" else None
    view_as = selected

    if role in ("manager", "manager_only") and selected != "— Select —":
        st.markdown("---")
        st.markdown("**View as consultant:**")
        browse = st.selectbox(
            "Browse team member",
            options=["— My own view —"] + sorted(CONSULTANT_DROPDOWN),
            key="briefing_browse",
        )
        if browse != "— My own view —":
            view_as = browse

    st.markdown("---")
    st.markdown("**Upload data**")
    drs_file = st.file_uploader("SS DRS Export (.xlsx / .csv)", type=["xlsx", "csv"], key="brief_drs")
    ns_file  = st.file_uploader("NS Time Detail (.xlsx / .csv)", type=["xlsx", "csv"], key="brief_ns")

# ── Guard: name required ──────────────────────────────────────────────────────
if selected == "— Select —":
    st.markdown('<div class="brief-header">Daily Briefing</div>', unsafe_allow_html=True)
    st.markdown('<div class="brief-sub">Select your name in the sidebar to get started.</div>', unsafe_allow_html=True)
    st.stop()

# ── Load data ─────────────────────────────────────────────────────────────────
df_drs = None
df_ns  = None
load_errors = []

if drs_file:
    try:
        df_drs = load_drs(drs_file)
    except Exception as e:
        load_errors.append(f"DRS load error: {e}")

if ns_file:
    try:
        df_ns = load_ns_time(ns_file)
    except Exception as e:
        load_errors.append(f"NS load error: {e}")

for err in load_errors:
    st.error(err)

# ── Filter to this consultant ─────────────────────────────────────────────────
today = date.today()
month_key = today.strftime("%Y-%m")
view_name = view_as  # could be manager browsing a consultant

# Normalise name for matching (Last, First → various forms)
def _name_variants(full_name: str) -> list[str]:
    """Return matching variants: 'Last, First', 'First Last', 'Last'."""
    parts = [p.strip() for p in full_name.split(",")]
    variants = [full_name.lower()]
    if len(parts) == 2:
        variants.append(f"{parts[1]} {parts[0]}".lower())
        variants.append(parts[0].lower())
    return variants

view_variants = _name_variants(view_name)

def _match_consultant(val: str) -> bool:
    return any(v in str(val).lower() for v in view_variants)

# Filter DRS to projects where this person is PM or assigned consultant
my_projects = pd.DataFrame()
if df_drs is not None and not df_drs.empty:
    pm_mask = df_drs.get("project_manager", pd.Series(dtype=str)).apply(
        lambda v: _match_consultant(str(v))
    )
    # Also check territory / resource cols if available
    my_projects = df_drs[pm_mask].copy()

# Filter NS time to this consultant's hours
my_ns = pd.DataFrame()
if df_ns is not None and not df_ns.empty:
    emp_mask = df_ns.get("employee", pd.Series(dtype=str)).apply(
        lambda v: _match_consultant(str(v))
    )
    my_ns = df_ns[emp_mask].copy()

# ── Header ────────────────────────────────────────────────────────────────────
display_name = view_name.split(",")[1].strip() if "," in view_name else view_name
viewing_as_note = f" (viewing as {display_name})" if view_as != selected and view_as != selected else ""
emp_info = EMPLOYEE_ROLES.get(view_name, {})
emp_role = emp_info.get("role", "Consultant")
emp_products = emp_info.get("products", [])

col_hdr, col_date = st.columns([3, 1])
with col_hdr:
    st.markdown(f'<div class="brief-header">Good morning, {display_name}{viewing_as_note}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="brief-sub">{emp_role} · {", ".join(emp_products) if emp_products else "All Products"} · {today.strftime("%A, %B %-d %Y")}</div>', unsafe_allow_html=True)
with col_date:
    location = EMPLOYEE_LOCATION.get(view_name, "")
    region = PS_REGION_OVERRIDE.get(view_name, PS_REGION_MAP.get(location, ""))
    if region:
        st.markdown(f'<div class="badge-blue action-badge" style="margin-top:10px">{region}</div>', unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Utilization snapshot
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">This Month — Utilization</div>', unsafe_allow_html=True)

avail = None
location_key = EMPLOYEE_LOCATION.get(view_name, "")
if location_key and month_key in AVAIL_HOURS.get(location_key, {}):
    avail = AVAIL_HOURS[location_key][month_key]

# Hours logged this month
if not my_ns.empty and "date" in my_ns.columns and "hours" in my_ns.columns:
    my_ns["date"] = pd.to_datetime(my_ns["date"], errors="coerce")
    month_ns = my_ns[my_ns["date"].dt.strftime("%Y-%m") == month_key].copy()
    # Exclude non-billable / PTO from utilization numerator
    billable_ns = month_ns[
        ~month_ns.get("project", pd.Series(dtype=str)).str.lower().str.contains(
            "vacation|pto|sick|non-billable|admin", na=False
        )
    ]
    hours_logged  = round(billable_ns["hours"].sum(), 1)
    hours_total   = round(month_ns["hours"].sum(), 1)
    util_pct      = round((hours_logged / avail * 100), 1) if avail else None

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="metric-card"><div class="metric-val">{hours_logged}h</div><div class="metric-lbl">Billable hours logged</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><div class="metric-val">{hours_total}h</div><div class="metric-lbl">Total hours logged</div></div>', unsafe_allow_html=True)
    with c3:
        if avail:
            remaining = round(avail - hours_total, 1)
            st.markdown(f'<div class="metric-card"><div class="metric-val">{avail}h</div><div class="metric-lbl">Available this month</div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">Available hours (location not mapped)</div></div>', unsafe_allow_html=True)
    with c4:
        if util_pct is not None:
            color = "#27AE60" if util_pct >= 75 else ("#F39C12" if util_pct >= 50 else "#E74C3C")
            st.markdown(f'<div class="metric-card"><div class="metric-val" style="color:{color}">{util_pct}%</div><div class="metric-lbl">Utilization</div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="metric-card"><div class="metric-val">—</div><div class="metric-lbl">Utilization</div></div>', unsafe_allow_html=True)
else:
    if not ns_file:
        st.info("Upload NS Time Detail to see your utilization snapshot.")
    else:
        st.warning(f"No time entries found for **{view_name}** in the NS file. Check that your name matches the employee column.")

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — My Projects
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">My Active Projects</div>', unsafe_allow_html=True)

if not drs_file:
    st.info("Upload SS DRS Export to see your projects.")
elif my_projects.empty:
    st.warning(f"No projects found where **{view_name}** is listed as Project Manager. Check that your name matches the DRS 'Project Manager' column.")
else:
    # Compute inactivity if NS loaded
    if df_ns is not None and not df_ns.empty:
        try:
            my_projects = calc_days_inactive(my_projects, df_ns)
        except Exception:
            pass

    # Sort: Red → Amber → Green, then by days_inactive desc
    rag_order = {"R": 0, "A": 1, "G": 2, "": 3}
    my_projects["_rag_sort"] = my_projects.get("rag", pd.Series(dtype=str)).fillna("").map(
        lambda x: rag_order.get(str(x).strip().upper()[:1], 3)
    )
    if "days_inactive" in my_projects.columns:
        my_projects = my_projects.sort_values(["_rag_sort", "days_inactive"], ascending=[True, False])
    else:
        my_projects = my_projects.sort_values("_rag_sort")

    st.markdown(f"**{len(my_projects)} project(s)** assigned to you")

    for _, row in my_projects.iterrows():
        proj_name = row.get("project_name", "—")
        rag       = str(row.get("rag", "")).strip().upper()[:1]
        phase     = row.get("phase", "—")
        pct       = row.get("pct_complete", None)
        go_live   = row.get("go_live_date", None)
        actual_h  = row.get("actual_hours", None)
        budget_h  = row.get("budgeted_hours", None)
        days_inac = row.get("days_inactive", None)
        status    = str(row.get("status", "")).strip()

        # RAG dot
        rag_html = f'<span class="rag-{rag}"></span>'

        # Burn % badge
        burn_badge = ""
        if actual_h and budget_h:
            try:
                burn = round(float(actual_h) / float(budget_h) * 100, 0)
                badge_cls = "badge-red" if burn >= 90 else ("badge-amber" if burn >= 75 else "badge-blue")
                burn_badge = f'<span class="action-badge {badge_cls}">{int(burn)}% burn</span>'
            except Exception:
                pass

        # Inactivity badge
        inac_badge = ""
        if days_inac is not None and days_inac >= 0:
            if days_inac >= 30:
                inac_badge = f'<span class="action-badge badge-red">{days_inac}d inactive</span>'
            elif days_inac >= 14:
                inac_badge = f'<span class="action-badge badge-amber">{days_inac}d inactive</span>'

        # Phase badge
        phase_badge = f'<span class="action-badge badge-gray">{phase}</span>' if phase and phase != "—" else ""

        # Go-live
        go_live_str = ""
        if go_live:
            try:
                gl = pd.to_datetime(go_live)
                days_to_gl = (gl.date() - today).days
                gl_fmt = gl.strftime("%-d %b %Y")
                if days_to_gl < 0:
                    go_live_str = f"Go-live: {gl_fmt} <span style='color:#E74C3C'>({abs(days_to_gl)}d overdue)</span>"
                elif days_to_gl <= 14:
                    go_live_str = f"Go-live: {gl_fmt} <span style='color:#F39C12'>({days_to_gl}d away)</span>"
                else:
                    go_live_str = f"Go-live: {gl_fmt} ({days_to_gl}d)"
            except Exception:
                pass

        pct_str = f"{int(float(pct))}% complete · " if pct else ""

        st.markdown(f"""
        <div class="proj-card">
            <div class="proj-name">{rag_html}{proj_name}</div>
            <div class="proj-meta">{pct_str}{go_live_str}</div>
            <div style="margin-top:8px">{phase_badge}{burn_badge}{inac_badge}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — Re-engagement actions
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">Re-Engagement Actions</div>', unsafe_allow_html=True)

if not drs_file:
    st.info("Upload SS DRS Export to see re-engagement actions.")
else:
    # Projects flagged stale (days_inactive >= 14) assigned to this consultant
    if not my_projects.empty and "days_inactive" in my_projects.columns:
        stale = my_projects[my_projects["days_inactive"] >= 14].copy()
        if stale.empty:
            st.success("No stale projects — all projects have recent activity.")
        else:
            stale = stale.sort_values("days_inactive", ascending=False)
            st.markdown(f"**{len(stale)} project(s)** need re-engagement outreach")
            for _, row in stale.iterrows():
                proj_name  = row.get("project_name", "—")
                days_inac  = int(row.get("days_inactive", 0))
                phase      = row.get("phase", "—")
                tier       = suggest_tier(days_inac)
                tier_label = tier if tier else "No template match"

                last_ms = ""
                try:
                    last_ms = calc_last_milestone(row)
                except Exception:
                    pass

                with st.expander(f"{proj_name} — {days_inac}d inactive · {tier_label}"):
                    c_info, c_tmpl = st.columns([1, 2])
                    with c_info:
                        st.markdown(f"**Phase:** {phase}")
                        st.markdown(f"**Days inactive:** {days_inac}")
                        if last_ms:
                            st.markdown(f"**Last milestone:** {last_ms}")
                    with c_tmpl:
                        if tier and tier in TEMPLATES:
                            tmpl = TEMPLATES[tier]
                            subject = tmpl.get("subject", "")
                            body_preview = tmpl.get("body", "")[:300] + "..."
                            st.markdown(f"**Suggested template:** {tier}")
                            st.markdown(f"*Subject:* {subject}")
                            st.code(body_preview, language=None)
                            st.caption("Open Re-Engagement page to fill and send this template.")
                        else:
                            st.markdown("No template matched for this inactivity window.")
    elif not my_projects.empty:
        st.info("Upload NS Time Detail alongside DRS to calculate project inactivity.")
    else:
        st.info("No projects found for this consultant.")

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — AI Summary (Claude-powered)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">AI Briefing Summary</div>', unsafe_allow_html=True)

has_data = (not my_projects.empty) or (not my_ns.empty)

if not has_data:
    st.info("Upload DRS and/or NS Time Detail to generate an AI summary of your day.")
else:
    # Build context payload for Claude
    def _build_context() -> str:
        lines = [
            f"Consultant: {view_name}",
            f"Role: {emp_role}",
            f"Products: {', '.join(emp_products) if emp_products else 'All'}",
            f"Date: {today.strftime('%Y-%m-%d')}",
            "",
        ]

        # Util
        if not my_ns.empty and "date" in my_ns.columns:
            try:
                month_ns = my_ns[my_ns["date"].dt.strftime("%Y-%m") == month_key]
                h = round(month_ns["hours"].sum(), 1)
                avail_h = avail or "unknown"
                lines.append(f"Hours logged this month: {h} / {avail_h} available")
            except Exception:
                pass

        # Projects
        if not my_projects.empty:
            lines.append(f"\nActive projects ({len(my_projects)}):")
            for _, row in my_projects.head(10).iterrows():
                name = row.get("project_name", "—")
                rag  = row.get("rag", "—")
                phase = row.get("phase", "—")
                pct  = row.get("pct_complete", "—")
                di   = row.get("days_inactive", None)
                inac_str = f", {di}d inactive" if di is not None and di >= 0 else ""
                lines.append(f"  - {name}: RAG={rag}, Phase={phase}, {pct}% complete{inac_str}")

        # Stale
        if not my_projects.empty and "days_inactive" in my_projects.columns:
            stale = my_projects[my_projects["days_inactive"] >= 14]
            if not stale.empty:
                lines.append(f"\nProjects needing re-engagement ({len(stale)}):")
                for _, row in stale.iterrows():
                    lines.append(f"  - {row.get('project_name','—')}: {int(row.get('days_inactive',0))}d inactive")

        return "\n".join(lines)

    ctx = _build_context()

    if st.button("Generate AI Briefing", type="primary", key="gen_briefing"):
        with st.spinner("Generating briefing..."):
            try:
                import requests, json
                payload = {
                    "model": "claude-sonnet-4-20250514",
                    "max_tokens": 1000,
                    "messages": [
                        {
                            "role": "user",
                            "content": f"""You are a PS operations assistant. Based on the consultant data below, write a concise daily briefing in plain prose. 

Structure:
1. One sentence summary of their workload status
2. Top 1-2 priorities for today (projects needing attention: high RAG, stale, near go-live)
3. Utilization status if relevant (on track / behind)
4. One re-engagement action if stale projects exist

Keep it under 150 words. Be direct and specific — name the projects. No bullet points, no headers. Write like a morning standup note.

---
{ctx}"""
                        }
                    ]
                }
                resp = requests.post(
                    "https://api.anthropic.com/v1/messages",
                    headers={"Content-Type": "application/json"},
                    json=payload,
                    timeout=30,
                )
                data = resp.json()
                briefing_text = ""
                for block in data.get("content", []):
                    if block.get("type") == "text":
                        briefing_text += block["text"]

                if briefing_text:
                    st.markdown(f'<div class="ai-box"><div class="ai-label">AI Briefing</div>{briefing_text}</div>', unsafe_allow_html=True)
                else:
                    st.error(f"No response from Claude. API error: {data.get('error', data)}")

            except Exception as e:
                st.error(f"AI briefing failed: {e}")
    else:
        st.markdown(
            '<div class="ai-box" style="color:#718096; font-style:italic;">'
            'Click <strong>Generate AI Briefing</strong> to get a Claude-powered summary of your day based on the loaded data.'
            '</div>',
            unsafe_allow_html=True
        )
