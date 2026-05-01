"""
PS Platform — Shared UI components
Zone & Co design system, global filter bar, hero blocks, status pills.

Every page imports inject_css() and render_filter_bar() at the top.
This is the single place to change look-and-feel; pages stay declarative.
"""
from __future__ import annotations
import streamlit as st

# ── Brand tokens ──────────────────────────────────────────────────────────────
ENERGY_BLUE  = "#014EDC"
TRUE_NAVY    = "#0E223D"
TITAN_ORANGE = "#FF4B40"
POP_YELLOW   = "#FFBC2F"
COOL_WHISPER = "#F5F7FB"
BORDER_1     = "#E3E8F0"
SUCCESS      = "#1F9D66"
SUCCESS_BG   = "#E5F4EC"

# ── Global CSS ────────────────────────────────────────────────────────────────
ZONE_CSS = """
<style>
  :root {
    --eb:  #014EDC;
    --nav: #0E223D;
    --org: #FF4B40;
    --pop: #FFBC2F;
    --bg2: #F5F7FB;
    --b1:  #E3E8F0;
    --ok:  #1F9D66;
  }

  /* ── Layout ── */
  .block-container {
    padding-top: 1.25rem;
    padding-bottom: 4rem;
    max-width: 1280px;
  }

  /* ── Typography ── */
  h1, h2, h3 { color: var(--nav); letter-spacing: -0.015em; }
  h1 { font-weight: 800; font-size: 1.75rem; }
  h2 { font-weight: 700; font-size: 1.25rem; }
  h3 { font-weight: 600; font-size: 1rem; }

  /* ── KPI metrics ── */
  [data-testid="stMetricValue"] {
    font-weight: 800;
    color: var(--nav);
    font-size: 2rem;
    letter-spacing: -0.02em;
  }
  [data-testid="stMetricLabel"] {
    font-weight: 600;
    color: #6B7A8F;
    font-size: 0.72rem;
    letter-spacing: 0.06em;
    text-transform: uppercase;
  }
  [data-testid="stMetricDelta"] { font-size: 0.8rem; }

  /* ── Cards ── */
  [data-testid="stContainer"][class*="border"] {
    background: white;
    border-radius: 12px;
    border-color: var(--b1) !important;
  }

  /* ── Status pills ── */
  .pill {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 999px;
    font-size: 11.5px;
    font-weight: 600;
    letter-spacing: 0.02em;
    white-space: nowrap;
  }
  .pill-ok      { background: #E5F4EC; color: #1F9D66; }
  .pill-warn    { background: #FFF4D9; color: #B07A00; }
  .pill-danger  { background: #FFE7E5; color: #D63A30; }
  .pill-info    { background: #E3EFFF; color: #014EDC; }
  .pill-hold    { background: #EEF1F6; color: #41526B; }
  .pill-neutral { background: #EEF1F6; color: #41526B; }

  /* ── Hero block ── */
  .hero-navy {
    background: linear-gradient(135deg, #0E223D 0%, #1A3257 100%);
    color: white;
    padding: 28px 32px;
    border-radius: 14px;
    margin-bottom: 1.5rem;
  }
  .hero-navy h1 {
    color: white;
    font-size: 26px;
    margin: 4px 0 8px;
  }
  .hero-navy .eyebrow {
    color: #73DAE3;
    font-size: 11px;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    font-weight: 700;
  }
  .hero-navy .stat {
    display: inline-block;
    margin-right: 32px;
    margin-top: 8px;
  }
  .hero-navy .stat-label {
    color: rgba(255,255,255,0.55);
    font-size: 11px;
    letter-spacing: 0.06em;
    text-transform: uppercase;
  }
  .hero-navy .stat-value {
    font-size: 30px;
    font-weight: 800;
    line-height: 1.1;
  }
  .hero-navy .stat-value.warn { color: #FFBC2F; }
  .hero-navy .stat-value.danger { color: #FF4B40; }

  /* ── Sidebar ── */
  [data-testid="stSidebar"] { background: #FAFBFD; }
  [data-testid="stSidebar"] .stMarkdown p {
    font-size: 13px;
    color: #6B7A8F;
  }

  /* ── Filter bar ── */
  .filter-bar-label {
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: #6B7A8F;
    margin-bottom: 4px;
  }

  /* ── Eyebrow / section label ── */
  .eyebrow-label {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #6B7A8F;
    margin-bottom: 2px;
  }

  /* ── DRS flag cards ── */
  .flag-icon { font-size: 1.4rem; line-height: 1; }

  /* ── Data tables ── */
  [data-testid="stDataFrame"] { border-radius: 8px; }

  /* ── Buttons ── */
  [data-testid="stButton"] button[kind="primary"] {
    background: var(--eb);
    border-color: var(--eb);
  }
  [data-testid="stButton"] button[kind="primary"]:hover {
    background: #0040B8;
    border-color: #0040B8;
  }

  /* ── Tab styling ── */
  [data-testid="stTabs"] [data-baseweb="tab"] {
    font-weight: 600;
    font-size: 13px;
  }

  /* ── Upload section labels ── */
  .upload-label {
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: #6B7A8F;
    margin-bottom: 4px;
    display: block;
  }
  .upload-link {
    font-size: 11px;
    color: var(--eb);
    text-decoration: none;
    display: block;
    margin-bottom: 4px;
    opacity: 0.75;
  }
  .upload-link:hover { opacity: 1; }
</style>
"""


def inject_css() -> None:
    """Inject Zone brand CSS. Call at the top of every page."""
    st.markdown(ZONE_CSS, unsafe_allow_html=True)


# ── Filter state ──────────────────────────────────────────────────────────────
def init_filters() -> None:
    """Initialise global filter state. Called once in app.py."""
    if "view" not in st.session_state:
        st.session_state["view"] = "ALL"
    if "product" not in st.session_state:
        st.session_state["product"] = "all"


def render_filter_bar(show: bool = True) -> tuple[str, str]:
    """
    Render the View + Product filter bar.
    Returns (view, product) for passing to data loaders.
    Pages without filters (Engagement, Profile, Time) should not call this.
    """
    init_filters()
    if not show:
        return st.session_state["view"], st.session_state["product"]

    from config import ACTIVE_EMPLOYEES, get_role, EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE

    # Build region → consultant map for the view dropdown
    _regions: dict[str, list[str]] = {}
    for n in ACTIVE_EMPLOYEES:
        if get_role(n) not in ("consultant", "manager"):
            continue
        loc = EMPLOYEE_LOCATION.get(n, "")
        if isinstance(loc, tuple):
            loc = loc[0]
        region = PS_REGION_OVERRIDE.get(n, PS_REGION_MAP.get(loc, "Other"))
        _regions.setdefault(region, []).append(n)

    _role = st.session_state.get("_role", "consultant")
    _is_manager = _role in ("manager", "manager_only", "reporting_only")

    col1, col2, col3, col4 = st.columns([2, 2, 5, 1])

    with col1:
        st.markdown('<div class="filter-bar-label">View as</div>', unsafe_allow_html=True)
        current_view = st.session_state["view"]
        _view_label = _view_display_label(current_view)
        with st.popover(f"{_view_label}", use_container_width=True):
            if _is_manager:
                if st.button("🌐 All team", key="v_all", use_container_width=True):
                    st.session_state["view"] = "ALL"
                    st.rerun()
                st.divider()
                st.caption("Region")
                for r in sorted(_regions.keys()):
                    n_count = len(_regions[r])
                    if st.button(f"{r}  ·  {n_count}", key=f"v_r_{r}", use_container_width=True):
                        st.session_state["view"] = f"REGION:{r}"
                        st.rerun()
                st.divider()
                st.caption("Individual")
                for n in sorted(ACTIVE_EMPLOYEES):
                    if get_role(n) not in ("consultant", "manager"):
                        continue
                    parts = n.split(",")
                    label = f"{parts[1].strip()} {parts[0].strip()}" if len(parts) == 2 else n
                    if st.button(label, key=f"v_p_{n}", use_container_width=True):
                        st.session_state["view"] = f"PERSON:{n}"
                        st.rerun()
            else:
                st.caption("You are viewing your own projects.")

    with col2:
        st.markdown('<div class="filter-bar-label">Product</div>', unsafe_allow_html=True)
        PRODUCTS = [
            ("all",        "All products"),
            ("billing",    "ZoneBilling"),
            ("capture",    "ZoneCapture"),
            ("approvals",  "ZoneApprovals"),
            ("reconcile",  "ZoneReconcile"),
            ("payroll",    "ZonePayroll"),
            ("reporting",  "ZoneReporting"),
        ]
        current_prod = st.session_state["product"]
        prod_label = dict(PRODUCTS).get(current_prod, "All products")
        with st.popover(f"{prod_label}", use_container_width=True):
            for pid, plabel in PRODUCTS:
                if st.button(plabel, key=f"p_{pid}", use_container_width=True):
                    st.session_state["product"] = pid
                    st.rerun()

    with col3:
        pass  # spacer

    with col4:
        st.markdown('<div class="filter-bar-label">&nbsp;</div>', unsafe_allow_html=True)
        if st.session_state["view"] != "ALL" or st.session_state["product"] != "all":
            if st.button("✕ Reset", use_container_width=True):
                st.session_state["view"] = "ALL"
                st.session_state["product"] = "all"
                st.rerun()

    return st.session_state["view"], st.session_state["product"]


def _view_display_label(view: str) -> str:
    if view == "ALL":
        return "🌐 All team"
    if view.startswith("REGION:"):
        return f"📍 {view.split(':', 1)[1]}"
    if view.startswith("PERSON:"):
        name = view.split(":", 1)[1]
        parts = name.split(",")
        return f"👤 {parts[1].strip()} {parts[0].strip()}" if len(parts) == 2 else f"👤 {name}"
    return f"👤 {view}"


# ── Shared UI primitives ──────────────────────────────────────────────────────
def page_header(eyebrow: str, title: str) -> None:
    """Render an eyebrow + h1 page header."""
    st.markdown(
        f'<div class="eyebrow-label">{eyebrow}</div>'
        f'<h1 style="margin:2px 0 1.25rem">{title}</h1>',
        unsafe_allow_html=True,
    )


def hero(eyebrow: str, title: str, subtitle: str = "", stats: list[tuple] | None = None) -> None:
    """
    Render the navy hero block.
    stats: list of (label, value, modifier) tuples. modifier: None | 'warn' | 'danger'
    """
    stats_html = ""
    if stats:
        for label, value, mod in stats:
            cls = f" {mod}" if mod else ""
            stats_html += (
                f'<div class="stat">'
                f'<div class="stat-label">{label}</div>'
                f'<div class="stat-value{cls}">{value}</div>'
                f'</div>'
            )
    sub_html = f'<p style="color:rgba(255,255,255,.7);margin:0 0 16px;font-size:14px">{subtitle}</p>' if subtitle else ""
    st.markdown(
        f'<div class="hero-navy">'
        f'<div class="eyebrow">{eyebrow}</div>'
        f'<h1>{title}</h1>'
        f'{sub_html}'
        f'{stats_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


def status_pill(status: str) -> str:
    """Return HTML for a coloured status pill."""
    kind = {
        "On track":     "ok",
        "Green":        "ok",
        "Active":       "ok",
        "At risk":      "warn",
        "Amber":        "warn",
        "Yellow":       "warn",
        "Off track":    "danger",
        "Red":          "danger",
        "Overdue":      "danger",
        "On hold":      "hold",
        "Pending input":"info",
    }.get(status, "neutral")
    return f'<span class="pill pill-{kind}">{status}</span>'


def rag_pill(rag: str) -> str:
    """Return HTML pill for RAG values."""
    v = str(rag).strip().lower()
    cls = {"green": "ok", "amber": "warn", "yellow": "warn", "red": "danger"}.get(v, "neutral")
    return f'<span class="pill pill-{cls}">{rag}</span>'


def metric_row(metrics: list[tuple]) -> None:
    """
    Render a row of st.metric cards.
    metrics: list of (label, value, delta, help) — delta and help are optional.
    """
    cols = st.columns(len(metrics))
    for col, m in zip(cols, metrics):
        label, value = m[0], m[1]
        delta = m[2] if len(m) > 2 else None
        help_text = m[3] if len(m) > 3 else None
        col.metric(label, value, delta=delta, help=help_text)


def section_label(text: str) -> None:
    """Small uppercase section divider label."""
    st.markdown(
        f'<div class="eyebrow-label" style="margin:1.5rem 0 0.5rem">{text}</div>',
        unsafe_allow_html=True,
    )


def empty_state(message: str, icon: str = "📭") -> None:
    """Centred empty state message."""
    st.markdown(
        f'<div style="text-align:center;padding:3rem 0;color:#6B7A8F">'
        f'<div style="font-size:2rem;margin-bottom:8px">{icon}</div>'
        f'<div style="font-size:14px">{message}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
