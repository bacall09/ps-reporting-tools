# PS Reporting Tools — Runbook

A working reference for operating, debugging, and extending the PS reporting tools Streamlit app. Combines operational steps ("if X breaks, do Y") with engineering rationale ("why we did X").

This is a living document. Pages are documented at one of three depths:

- **Deep** — full operational + decisions context. Currently: Utilization Report, shared infrastructure, Streamlit patterns
- **Light** — purpose, role gating, key data sources, known quirks. All other pages
- **TODO** — flagged for future expansion as those areas get attention

When you work on a page, promote its entry from Light → Deep.

---

## Table of contents

1. [System overview](#system-overview)
2. [Roles & permissions](#roles--permissions)
3. [Data sources](#data-sources)
4. [Streamlit patterns we rely on](#streamlit-patterns-we-rely-on)
5. [Streamlit gotchas — hard-won lessons](#streamlit-gotchas--hard-won-lessons)
6. [Page reference](#page-reference)
   - [Utilization Report (deep)](#utilization-report-deep)
   - [Daily Briefing (light)](#daily-briefing-light)
   - [Customer Reengagement (light)](#customer-reengagement-light)
   - [Workload Health Score (light)](#workload-health-score-light)
   - [Capacity Outlook (light)](#capacity-outlook-light)
   - [DRS Health Check (light)](#drs-health-check-light)
   - [Vibe Check (light)](#vibe-check-light)
   - [My Projects (light)](#my-projects-light)
   - [Help (light)](#help-light)
   - [Revenue Report (light)](#revenue-report-light)
   - [Time Entries (light)](#time-entries-light)
   - [Project Health (light)](#project-health-light)
   - [Portfolio Analytics (light)](#portfolio-analytics-light)
   - [Capacity Planner (light)](#capacity-planner-light)
   - [Customer Profile (light)](#customer-profile-light)
7. [Shared infrastructure (deep)](#shared-infrastructure-deep)
8. [Operational playbook](#operational-playbook)
9. [TODO / known unknowns](#todo--known-unknowns)

---

## System overview

The PS reporting tools is a multi-page Streamlit app. The entrypoint is `Home.py`, which handles auth, data uploads, and navigation. Each page in `pages/` is a standalone Streamlit script that reads from session state (populated by Home or by user upload) and renders a page.

Two primary data sources flow through `st.session_state`:

- **`df_ns`** — NetSuite time-entry detail. Per-row: employee, project, task, hours, date, billing_type, project_type, sku, etc.
- **`df_drs`** — Smartsheet DRS (Delivery Resource Scheduler) export. Per-project: schedule, milestones, owners, scope hours, status.

Some pages also accept ad-hoc uploads (e.g. Revenue Report's revenue file, Capacity Outlook's unassigned-projects export).

Shared logic lives in `shared/`:
- `config.py` — color tokens, region maps, scope defaults, capacity tables
- `constants.py` — employee roster, role lookups, view-as resolver
- `utils.py` — credit-assignment engine, Excel report builder, capacity calc
- `loaders.py` — file ingestion (NetSuite, Smartsheet, etc.)
- `whs.py` — Workload Health Score computation
- `smartsheet_api.py` — Smartsheet API client
- `excel_formatter.py` — Excel cell/column formatting helpers
- `template_utils.py` — re-engagement template suggestions
- `activity_log.py` — session activity log used by Time Entries

---

## Roles & permissions

Source of truth: `shared/constants.py`.

### Role types

| Role | Returned by `get_role()` | What they see |
|------|--------------------------|---------------|
| `manager_only` | Hardcoded list at top of `constants.py` (`MANAGERS_ONLY`) | All pages, all data, no per-consultant filter |
| `manager` | In `MANAGER_CONSULTANTS` list | All pages, can switch "View as" to other consultants/regions |
| `reporting_only` | In `REPORTING_ONLY` list | Same as manager for view, but doesn't show in consultant dropdowns |
| `consultant` | Default for anyone in `EMPLOYEE_ROLES` not above | Their own data only; no "View as" picker |
| `no_access` | In `NO_ACCESS` list | Login is rejected |

### Helpers

```python
from shared.constants import get_role, is_manager, is_consultant

is_manager(name)     # → manager OR manager_only
is_consultant(name)  # → consultant OR manager
```

### Pattern for gating UI to managers

```python
_logged_in = st.session_state.get("consultant_name", "")
from shared.constants import get_role as _gr
_role_u = _gr(_logged_in)
_is_mgr_u = _role_u in ("manager", "manager_only", "reporting_only")

if _is_mgr_u:
    # Render manager-only widget
    ...
```

This is the pattern used in Utilization Report for the Excel + Tableau exports. When adding a new manager-only feature, copy this guard. Don't invent new role checks — they drift.

### "View as" picker

Managers and reporting_only roles get a sidebar "View as" picker that re-filters the underlying data. Implementation: `resolve_view_as()` in `shared/constants.py`. The page reads `st.session_state.get("home_browse", "")` (or the page's local override key) and re-applies a filter to `df_ns` / `df_drs` before rendering.

When adding a new page that should respect "View as":
1. Read `_logged_in` and `_home_browse` (or the page's override key)
2. Call `resolve_view_as()` to get `_va_name`, `_va_region`
3. Filter your data before rendering

See lines 535-590 of `pages/3_Utilization_Report.py` for a complete reference implementation.

---

## Data sources

### NetSuite time entries (`df_ns`)

Loaded by `Home.py` (or via per-page upload in some flows). Stored at `st.session_state["df_ns"]`. Loader: `shared/loaders.py::load_ns()`.

Columns the codebase expects (not exhaustive):
- `employee` — full name, "Last, First" format
- `project` — project name (matches DRS `project_name` for joins)
- `project_id`, `project_internal_id` — IDs (sometimes one or the other depending on export)
- `task` — task category (e.g. "Configuration", "Training & UAT")
- `hours` — float
- `date` — datetime
- `period` — string like "2026-04"
- `project_type` — e.g. "ZoneApp: Capture"
- `billing_type` — "Time and Material", "Fixed Fee", "Internal", etc.
- `sku` — for IMPL10/IMPL20 detection in `assign_credits`
- `non_billable` — flag/note column
- `customer` — customer name

If a page hits "KeyError" on one of these columns, it usually means the NetSuite export schema changed. Check the loader first.

### DRS / Smartsheet (`df_drs`)

Loaded by `Home.py` (Smartsheet API) or per-page upload. Stored at `st.session_state["df_drs"]`.

Columns the codebase expects:
- `project_id` — joins to NS `project` or `project_id`
- `project_name`
- `customer`, `region`
- `start_date`, `end_date`, `go_live_date`
- `status` — project status string
- Milestone columns — see `MILESTONE_COLS_MAP` in `constants.py`

### DRS enrichment in Utilization Report

The Utilization Report engine (`_run_utilization_engine`) optionally cross-references `df_drs` to fill in canonical project names when NetSuite has incomplete project labels. Pattern: build a `project_id → project_name` map from DRS, then map NS rows to it.

This means the engine's output can differ depending on whether DRS data is loaded. The cache key includes `df_drs` to handle this.

---

## Streamlit patterns we rely on

### Page layout pattern

Every page starts roughly the same way:

```python
import streamlit as st
import pandas as pd
# ... other imports ...

st.session_state["current_page"] = "Page Name"

# Hero banner — dark navy, brand mark
st.markdown(f"<div style='background:#050D1F;...'>...</div>", unsafe_allow_html=True)

# Auth / role check
_logged_in = st.session_state.get("consultant_name", "")
# ... role logic ...

# Data check
if "df_ns" not in st.session_state or st.session_state["df_ns"] is None:
    st.warning("Upload NetSuite data on Home")
    return

# Main content
def main():
    ...

main()
```

The hero banner is always a hardcoded dark navy (`#050D1F`) and never theme-switches. It's a brand element.

### Card surfaces (theme-aware)

Cards (KPI tiles, callouts, data wrappers) follow the **Portfolio Analytics pattern**: no explicit `background` property. Just border + padding + `color: inherit`. This lets the card surface inherit the page background, which Streamlit flips with theme automatically.

```css
.metric-card {
    border: 1px solid rgba(128,128,128,0.25);
    border-radius: 8px;
    padding: 14px;
    color: inherit;
    /* NO background property — inherits from page */
}
```

This is the pattern used in `pages/13_Portfolio_Analytics.py` and now in `pages/3_Utilization_Report.py`. It works in both light and dark mode without explicit overrides. Don't try to be clever with `var(--color-background-primary)` in `<style>` blocks — see [gotchas](#streamlit-gotchas--hard-won-lessons).

### Section labels

Use the same blue uppercase label across pages. CSS:

```css
.section-label {
    font-size: 13px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.8px; color: #4472C4; margin-bottom: 8px;
}
```

Reference: `pages/1_Daily_Briefing.py` line 54.

### Pills (status indicators)

Use translucent backgrounds so they work on either page bg. Color flips via `prefers-color-scheme: dark` and `[data-theme="dark"]` selectors:

```css
.util-pill-green { background: rgba(34, 197, 94, 0.18); color: #15803d; }
@media (prefers-color-scheme: dark) {
    .util-pill-green { color: #7ed4a4; }
}
.stApp[data-theme="dark"] .util-pill-green { color: #7ed4a4; }
```

The `[data-theme]` selector handles Streamlit's manual theme toggle (Settings → Theme); `prefers-color-scheme` handles OS-level preference. Need both.

### Caching expensive computations

Two layers:

**Cross-user, persistent**: `@st.cache_data(ttl=300)` on pure functions. Used for the utilization engine. Same inputs → same outputs across users. Refresh button calls `function_name.clear()` to invalidate.

**Per-user, session-scoped**: `st.session_state["_some_cache"][key]`. Used for things that depend on the current user's filters (e.g. Excel buffers tied to a specific view). NEVER use `id(some_object)` as a cache key — see [gotchas](#streamlit-gotchas--hard-won-lessons).

### Two-stage button → action → render pattern

When an action takes multiple seconds (Excel build, large query), don't put it inside a `try/except` block alongside `st.rerun()`. The pattern that works:

```python
# Stage 1: button click sets a flag
if st.button("Do expensive thing"):
    st.session_state["_action_requested"] = stable_key
    st.rerun()  # outside try/except

# Stage 2: on next render, check flag, do the work, clear flag
if st.session_state.get("_action_requested") == stable_key:
    st.session_state["_action_requested"] = None
    try:
        with st.spinner("Working..."):
            result = do_expensive_thing()
            st.session_state["_result_cache"][stable_key] = result
    except Exception as e:
        st.error(f"Failed: {e}")
        st.exception(e)

# Stage 3: render normally based on cache state
if st.session_state["_result_cache"].get(stable_key):
    st.download_button(...)
else:
    if st.button("Do expensive thing"):
        ...
```

This avoids the `RerunException` swallowing problem (see gotchas).

### Lazy-load expensive tabs

If a tab does heavy computation that's only relevant when the user actually opens it, gate it behind a "Load" button on first visit:

```python
with tab_heavy:
    if not st.session_state.get("_heavy_visited"):
        if st.button("Load heavy tab"):
            st.session_state._heavy_visited = True
            st.rerun()
    else:
        # Render the actual tab content
        ...
```

Used in Utilization Report's Trend and Task analysis tabs. Saves several seconds per render when the user only wants At-a-glance.

---

## Streamlit gotchas — hard-won lessons

These are real bugs that cost real time. If you see one of these symptoms, jump here first.

### `RerunException` is a subclass of `Exception`

**Symptom**: A button click appears to do nothing. Cache is being populated (verified by adding a `st.write` after) but the page doesn't show the post-action state.

**Cause**: `st.rerun()` works by raising `streamlit.runtime.scriptrunner.RerunException`, which is caught by a bare `except Exception:`. So:

```python
# BROKEN
if st.button("Click"):
    try:
        do_work()
        st.rerun()  # raises RerunException
    except Exception as e:
        # ← RerunException is caught here, swallowed silently
        st.error(f"Failed: {e}")
```

**Fix**: Never wrap `st.rerun()` inside a `try/except Exception`. Either put it outside, or catch only the specific exceptions you expect:

```python
# OK
if st.button("Click"):
    try:
        do_work()
    except SpecificError as e:
        st.error(f"Failed: {e}")
        return
    st.rerun()  # outside try/except
```

Or better yet, use the [two-stage pattern](#two-stage-button--action--render-pattern) so the rerun and the work are in different code paths.

### `id()` is not a stable cache key across reruns

**Symptom**: A cache populated on one render isn't found on the next render, even though the inputs look identical.

**Cause**: Python's `id()` returns the memory address of an object. Streamlit reruns rebind variables, so `id(result)` changes each time even when the underlying value is the same.

**Fix**: Build cache keys from the actual *inputs* that determine the value:

```python
# BROKEN
_cache_key = ("xl", id(result))  # changes every render

# OK
_cache_key = ("xl", _ns_signature(df_ns), str(period_start), str(period_end),
              view_filter_name, view_filter_region)
```

### CSS variables don't propagate to `<style>` blocks reliably

**Symptom**: Cards styled via a CSS class look fine in inline-styled prototypes but render with wrong/no colors when applied via class. Cards may appear pure white in dark mode, or pure transparent.

**Cause**: Streamlit's CSS variables (`--color-background-primary`, `--color-text-primary`) are defined on `:root` but injection timing into `<style>` blocks parsed via `st.markdown(unsafe_allow_html=True)` is unreliable. Class-based rules using `var(--color-background-primary)` may resolve to the fallback value.

**Inline styles work** because they re-evaluate on every paint. CSS classes parse once.

**Fix**: For card surfaces, omit the `background` property entirely. Inherit from the page background, which Streamlit flips natively. See [card surfaces](#card-surfaces-theme-aware).

If you genuinely need a class-based background that flips with theme, use hardcoded hex values with `prefers-color-scheme` and `[data-theme]` selectors:

```css
.my-card { background: #ffffff; color: #1a1a1a; }
@media (prefers-color-scheme: dark) {
    .my-card { background: #0E1117; color: #fafafa; }
}
.stApp[data-theme="dark"] .my-card { background: #0E1117; color: #fafafa; }
```

### `import re as _re_constants` aliases don't follow when functions move

**Symptom**: A function works fine in one file, throws `NameError: name '_re_constants' is not defined` after being moved to another file.

**Cause**: Module-level import aliases are not part of the function. Moving the function loses the alias unless the destination module has the same alias.

**Fix**: When moving functions between modules, audit the destination's imports. Run the function with synthetic test data after moving — `py_compile` won't catch this because it's a runtime name lookup.

### `st.session_state` is per-user; `@st.cache_data` is cross-user

**Symptom**: Concern that two users might see each other's data.

**Reality**: They can't. These are different caches:

- `st.session_state` is per-browser-session. Never shared.
- `@st.cache_data` is keyed by function arguments. Two users with the same inputs get the same cached output (which is *correct* — same inputs always produce same outputs for a pure function). Two users with different inputs (different period, different filter, different DRS data) get fresh computations.

If you're unsure whether something can leak, check: does the cache key include all inputs that distinguish users? If yes, it's safe.

### `st.empty()` only retains the last element written

When using `st.empty()` as a placeholder for a swap (e.g. button → spinner → download_button), each new write replaces the previous content. Don't try to use `with empty:` for multi-element layouts; use `st.container()` instead.

### Long form / multi-step interactions need explicit state machines

Streamlit's "everything reruns top to bottom" model means complex flows (wizard, multi-step form) need explicit state in `st.session_state` to track which step the user is on. Don't try to use Python control flow for it — a click-driven rerun resets the function.

---

## Page reference

### Utilization Report (deep)

**File**: `pages/3_Utilization_Report.py`
**Purpose**: Per-period utilization credit and capacity report. Shows billable/non-billable breakdown, projects-at-risk, trend analysis, task analysis.
**Roles**: All can view the page. Excel + Tableau exports gated to `manager`/`manager_only`/`reporting_only`.
**Data sources**: `df_ns` (required), `df_drs` (optional, for project name enrichment).

#### Architecture

The page has six tabs: At a glance, Consultants, Projects at risk, Trend, Task analysis, Detail.

**Engine** (`_run_utilization_engine`): pure function decorated with `@st.cache_data(ttl=300)`. Takes `df_raw, period_start, period_end, df_drs` and returns `{df, consumed, empty}`. The output `df` has credit_tag assigned per row; `consumed` is a `{project_id: hours_consumed}` dict.

**Tab-level result objects**:
- Main `result` — engine called with the user's selected period
- `_trend_result` — engine called with a wider window for the Trend + Task tabs (lazy-loaded; only computed after user opens one of those tabs)
- Prior-period results inside Trend and Task tabs — engine called with a comparison window (e.g. prior month, prior 4 weeks)

All four call the same `@st.cache_data`-decorated engine, so caching is automatic across reruns.

#### Critical implementation details

**Theme handling**: Cards use the no-background pattern. Don't add explicit `background:` to `.util-card`, `.util-kpi`, `.util-callout`, `.util-legend`, `.util-table-header`, `.util-table-wrap`. They inherit page bg.

**Excel + Tableau exports**: Two-stage button → flag → rerun → build pattern (see [Streamlit patterns](#streamlit-patterns-we-rely-on)). Cache keys are stable inputs, NOT `id(result)`. Manager-gated via `_is_mgr_u`. The Excel build is ~898 lines of openpyxl writes and is slow — caching matters.

**Lazy-load Trend + Task**: First click on either tab shows a "Load" button that sets a session_state flag and reruns. The wider-window engine call only happens after the flag is set. Saves 2-5 seconds per render when the user only wants At-a-glance.

**Section labels**: Both Trend and Task tabs have section labels matching Daily Briefing's "Team Breakdown — This Week" style (uppercase, blue, weight 700). Class is `util-section-label`.

**Trend metric headlines**: Show *period totals* (the credit % across the entire selected period, not the last week's value) with delta vs prior equivalent period. Not vs prior week. The metric headline is what's actionable for a manager looking at "April utilization."

**Task analysis Movers normalization**: When comparing current period vs prior 4 weeks (per-week avg), both sides are normalized to per-week. The card subtitle exposes the divisors explicitly: `Current ÷ 1.1w  vs  prior ÷ 4.0w (3 Apr → 30 Apr)` so you can sanity-check the math.

#### Common operational issues

| Symptom | Cause | Fix |
|---------|-------|-----|
| Excel button does nothing | `RerunException` swallowed by `except Exception` | Use the two-stage pattern (see line ~727) |
| Cards appear pure white in dark mode | CSS class with explicit `background:` | Remove the `background` property — inherit from page |
| Trend / Task tab slow on first click | Wider-window engine call running | This is expected once. Cached for subsequent renders |
| `NameError: _re_constants` | `re as _re_constants` alias not in target module | Add `import re as _re_constants` to `shared/utils.py` |
| Period change is slow | Engine + Excel build running unconditionally | Verify `@st.cache_data` is on engine, exports are gated behind buttons |
| "Showing N projects · X hrs" doesn't match KPIs | Status line and KPIs computed from different filtered frames | Check that both use the same `df` from `result`, after view-as filter |

#### Cache invalidation

The Refresh button calls:
1. `_run_utilization_engine.clear()` — drops all cached engine results across users
2. `st.session_state.pop("_util_excel_cache", None)` — drops Excel buffers
3. `st.session_state.pop("_util_tableau_cache", None)` — drops Tableau buffers

This is the only way to force fresh data without waiting for the 5-min TTL.

---

### Daily Briefing (light)

**File**: `pages/1_Daily_Briefing.py` (1,497 lines)
**Purpose**: Month-to-date utilization snapshot, team breakdown, re-engagement actions. The "morning coffee" page.
**Roles**: All. Manager view shows team breakdown; consultant view shows their own data.
**Data sources**: `df_ns`, `df_drs`.
**Notable**: Reference implementation for the section-label style. Uses `var(--color-background-primary)` in inline styles successfully (works because inline styles re-evaluate per paint).

---

### Customer Reengagement (light)

**File**: `pages/2_Customer_Reengagement.py` (1,950 lines)
**Purpose**: Identify customers with no recent PS activity, suggest tier-appropriate re-engagement templates.
**Roles**: Heavily role-gated (9 references to role checks). Likely managers + reporting_only.
**Data sources**: `df_ns`, `df_drs`. Uses `rapidfuzz` for name matching.
**Notable**: Pulls templates from `shared/template_utils.py::suggest_tier()`.

---

### Workload Health Score (light)

**File**: `pages/4_Workload_Health_Score.py` (1,734 lines)
**Purpose**: Per-consultant workload health score (0-100) based on hours, project count, milestone density.
**Roles**: Role-gated (8 references).
**Data sources**: `df_ns`, `df_drs`. Loader: `load_ns()`.
**Notable**: WHS computation in `shared/whs.py`.

---

### Capacity Outlook (light)

**File**: `pages/5_Capacity_Outlook.py` (2,071 lines — largest page)
**Purpose**: Project consultant availability across upcoming months. Combines DRS schedule with NS unassigned-projects data.
**Roles**: Has at least 1 role check.
**Data sources**: `df_drs`, `df_ns`, plus a separate "unassigned projects" upload.
**Notable**: Original name in docstring is "Resourcing Planner."

---

### DRS Health Check (light)

**File**: `pages/6_DRS_Health_Check.py` (551 lines)
**Purpose**: Logical consistency validator for Smartsheet DRS data — flags missing dates, mismatched statuses, orphaned milestones.
**Roles**: Role-gated (8 references).
**Data sources**: `df_drs` only.

---

### Vibe Check (light)

**File**: `pages/7_Vibe_Check.py` (162 lines)
**Purpose**: Morale/fun page. Uses Giphy API to surface gifs.
**Roles**: No gating.
**Data sources**: Giphy API. Requires `st.secrets["GIPHY_API_KEY"]`.
**Notable**: Not core reporting. Hero banner uses gradient instead of solid navy.

---

### My Projects (light)

**File**: `pages/8_My_Projects.py` (798 lines)
**Purpose**: Per-consultant working list. Snapshot metrics, needs-action items, project links.
**Roles**: Light role check (2 references). Per-user view.
**Data sources**: `df_ns`, `df_drs`, direct Smartsheet API calls.

---

### Help (light)

**File**: `pages/9_Help.py` (176 lines)
**Purpose**: Reference guide / glossary. Static content.
**Roles**: No gating.
**Data sources**: None.

---

### Revenue Report (light)

**File**: `pages/9_Revenue_Report.py` (1,206 lines)
**Purpose**: Fixed Fee straight-line revenue recognition (YTD/QTD/MTD).
**Roles**: Role-gated (2 references).
**Data sources**: `df_ns`, `df_drs`, plus a separate revenue file upload.

⚠️ **Note**: Two pages are numbered `9_*` (Help and Revenue Report). Streamlit may show ordering inconsistencies. Consider renumbering when convenient.

---

### Time Entries (light)

**File**: `pages/10_Time_Entries.py` (258 lines)
**Purpose**: Session-based activity log → draft NS time entries → CSV export.
**Roles**: No role gating. Per-user.
**Data sources**: `df_drs` (for project picker). Activity log in `shared/activity_log.py`.

---

### Project Health (light)

**File**: `pages/11_Project_Health.py` (590 lines)
**Purpose**: Delivery performance — schedule variance, milestone health, scope health.
**Roles**: Role-gated (2 references).
**Data sources**: `df_drs`.

---

### Portfolio Analytics (light)

**File**: `pages/13_Portfolio_Analytics.py` (821 lines)
**Purpose**: Manager-level portfolio view — team utilization trends, project risk distribution.
**Roles**: Role-gated (2 references). Manager-oriented.
**Data sources**: `df_ns`, `df_drs`.
**Notable**: **Reference implementation for theme-aware card styling.** When in doubt about how to style a card to flip cleanly between light and dark mode, look at `.metric-card` in this file. Imports `calc_consultant_util` from `shared/utils.py` — that's the only external consumer of that function.

---

### Capacity Planner (light)

**File**: `pages/14_Capacity_Planner.py` (559 lines)
**Purpose**: Manager tool to model consultant delivery capacity based on product mix.
**Roles**: Role-gated (2 references). Manager-oriented.
**Data sources**: `df_ns`, `df_drs`.

---

### Customer Profile (light)

**File**: `pages/99_Customer_Profile.py` (1,644 lines)
**Purpose**: Gong handover intelligence — pain points, stakeholders, requirements per customer.
**Roles**: Role-gated (4 references).
**Data sources**: `df_ns`, `df_drs`, NetSuite enrichment.

---

## Shared infrastructure (deep)

### `shared/config.py`

Color tokens, region maps, scope defaults, capacity tables.

Key exports:
- `NAVY`, `TEAL`, `WHITE`, `LTGRAY`, `MID_GRAY` — Excel report colors
- `TAG_COLORS`, `TAG_BADGE` — credit tag visual mapping
- `PTO_KEYWORDS` — strings that mark a row as PTO/sick/vacation
- `UTIL_EXEMPT_EMPLOYEES` — list of names excluded from utilization calculation
- `EMPLOYEE_LOCATION` — name → country/region mapping
- `PS_REGION_MAP`, `PS_REGION_OVERRIDE` — country → PS region grouping
- `AVAIL_HOURS` — region+month → available capacity hours
- `DEFAULT_SCOPE` — project_type → scoped hours dict (e.g. `"ZoneApp: Capture": 20.0`)

### `shared/constants.py`

Roles, employee roster, view-as resolver. Source of truth for who's allowed to see what.

Key exports:
- `MANAGERS_ONLY`, `MANAGER_CONSULTANTS`, `REPORTING_ONLY`, `NO_ACCESS` — role lists
- `EMPLOYEE_ROLES` — full roster dict (name → role/products/learning)
- `ACTIVE_EMPLOYEES`, `CONSULTANT_DROPDOWN` — filtered subsets
- `get_role(name) → str` — primary role resolver
- `is_manager(name) → bool`, `is_consultant(name) → bool` — convenience checks
- `resolve_view_as(...)` — used by all manager-aware pages
- `get_region_consultants(...)` — region filter helper
- `name_matches(a, b)` — fuzzy name match (handles "Last, First" vs "First Last")
- Column-mapping dicts: `MILESTONE_COLS_MAP`, `SS_COL_MAP`, `NS_COL_MAP`, `SFDC_COL_MAP`

⚠️ **Drift warning**: `EMPLOYEE_ROLES` is defined in *both* `shared/constants.py` AND `pages/3_Utilization_Report.py`. The page version was edited locally and differs from the shared version (~157 char delta as of last consolidation). Don't try to deduplicate without reviewing the differences first.

### `shared/utils.py`

Excel report builder, credit assignment engine, capacity calc.

Key exports:
- `assign_credits(df, scope_map) → (df_with_credits, consumed_dict, skipped)` — the engine
- `build_excel(df, scope_map, consumed) → BytesIO` — full multi-sheet Excel report
- `auto_detect_columns(df)` — fuzzy column-name resolver for varying NS exports
- `match_ff_task(task)` — task category resolver
- `get_avail_hours(region, period)` — capacity lookup
- `calc_consultant_util(...)` — consumed externally by Portfolio Analytics
- Excel styling helpers: `thin_border`, `hdr_fill`, `row_fill`, `group_bg`, `style_header`, `style_cell`, `write_title`

⚠️ **Performance hot path**: `assign_credits` iterates row-by-row through ~1500-row datasets and is the slowest path on every render. Vectorizing it (groupby + cumsum for scope tracking) is a known optimization, ~5-10× speedup, ~half-day work. Not yet done.

⚠️ **Performance hot path #2**: `build_excel` has 73 `for` loops in 898 lines using openpyxl cell-by-cell writes. Profiling for bulk row writes is another known optimization, ~3-5× speedup, ~half-day work. Not yet done.

### `shared/loaders.py`

File ingestion. Largest shared module (1,952 lines).

Key functions:
- `load_ns(file)` — NetSuite time-entry parser
- `load_drs(file)` — DRS Smartsheet export parser
- `load_revenue(file)` — Revenue file parser

Failure modes typically appear as `KeyError` on a column the loader couldn't find, or schema-mismatch warnings printed to the page.

### `shared/whs.py`

Workload Health Score computation.

Key exports:
- `workload_level(score) → str` — score → label ("Healthy", "Stretched", "Overloaded", etc.)

### `shared/smartsheet_api.py`

Smartsheet API client. Used when DRS is fetched directly from Smartsheet rather than uploaded.

Requires `st.secrets["SMARTSHEET_API_TOKEN"]`.

### `shared/excel_formatter.py`

Cell/column formatting helpers (cross-page, distinct from `utils.py::build_excel` which is Utilization-Report-specific).

### `shared/template_utils.py`

Re-engagement template suggestion logic. Used by Customer Reengagement.

Key exports:
- `suggest_tier(days_inactive) → str` — tier label

### `shared/activity_log.py`

Session-scoped activity log used by Time Entries.

### `Home.py`

Entrypoint. 368 lines. Handles:
- Auth (username/password against `st.secrets`)
- Sets `st.session_state["consultant_name"]`, `["authentication_status"]`, `["name"]`
- File upload hub: NS, DRS, Smartsheet API trigger, revenue file
- Navigation sidebar
- "Browse as" (manager-only) — sets `_browse_passthrough` for downstream pages

When debugging "consultant doesn't see what they should": start here. Check `st.session_state["consultant_name"]` is populated and matches the roster.

---

## Operational playbook

### Deployment

(Currently running on Streamlit Community Cloud / equivalent. Update this section once on dedicated hosting.)

To deploy:
1. Push to the staging branch
2. Streamlit auto-redeploys on push
3. First load = cold start (~5-10s on free tier)

### Cold start mitigation

Until on dedicated hosting, cold starts are unavoidable. Users should expect 5-10s on first load of the day, fast thereafter.

### Cache strategy

Two layers:

1. **`@st.cache_data(ttl=300)`** on pure functions (currently: `_run_utilization_engine` only). Cross-user. Auto-expires after 5 min.

2. **`st.session_state["_*_cache"]`** for per-user buffers (Excel, Tableau). Cleared by Refresh button or by switching period/view.

To force-refresh everything:
- User clicks Refresh button on the page
- This calls `_run_utilization_engine.clear()` and pops all session caches

To restart cleanly:
- Reboot the Streamlit server (host-specific)

### When the Excel button doesn't work

1. Check the page's `_is_mgr_u` evaluates True for the logged-in user
2. Check `st.session_state["_util_excel_prep_requested"]` is being set on click — add `st.write(st.session_state.get("_util_excel_prep_requested"))` to debug
3. Check the cache key matches between set and read — print `_excel_cache_key` in both branches
4. Check no `try/except Exception` is wrapping `st.rerun()` — that swallows the rerun
5. If the spinner shows but no download appears, the build is failing. The error should show on screen now (we added `st.exception()` in the build branch); if not, check Streamlit logs

### When cards render wrong in dark mode

1. Confirm the user is actually in dark mode (Settings → Theme)
2. Check the card's CSS class — if it has an explicit `background:` property, that's likely the issue
3. Match the Portfolio Analytics pattern: no `background`, just border + padding + `color: inherit`
4. If you really need a class-based background: use `prefers-color-scheme` + `[data-theme]` selectors with hardcoded hex values

### When a user sees data they shouldn't

1. Check `get_role(consultant_name)` for that user
2. Check the page's role gating — does it use the standard `_is_mgr_u in (...)` pattern, or does it have a custom check?
3. Check the page applies the view-as filter to its data BEFORE rendering
4. If a manager-only widget is showing for a consultant: search for the widget's render code, confirm it's wrapped in `if _is_mgr_u:`

### When a page errors with `KeyError` on a column

1. The NetSuite or DRS export schema likely changed
2. Check `shared/loaders.py` for the column rename / detection logic
3. The fix is usually to add an alternate column name to the loader's column-detection list

### When `assign_credits` errors with `NameError: _re_constants`

1. `shared/utils.py` is missing `import re as _re_constants` at the top
2. Add it. This was a bug introduced during consolidation; should be fixed now but can recur if someone moves functions again

### When period changes feel slow

1. Confirm `@st.cache_data` is decorating `_run_utilization_engine`
2. Confirm exports (Excel, Tableau) are behind buttons, not eager builds
3. Confirm Trend + Task tabs are lazy-loaded (require explicit click)
4. If still slow: the engine itself is the bottleneck. See the `assign_credits` vectorization TODO

---

## TODO / known unknowns

Things that should be filled in over time. Promote pages from light → deep as you work on them.

### Performance optimizations (deferred)

- **Vectorize `assign_credits`** — currently row-by-row pandas iteration. Estimated 5-10× speedup. Half-day work.
- **Profile `build_excel`** — 73 openpyxl loops in 898 lines. Switch to bulk row writes where possible. Estimated 3-5× speedup. Half-day work.

### Documentation gaps (light → deep candidates)

- **Daily Briefing** — large, complex, the "morning coffee" page. Section-label reference. Worth a deep entry.
- **Customer Reengagement** — second-largest page, complex template logic. Worth a deep entry.
- **Workload Health Score** — has its own scoring algorithm in `shared/whs.py`. Worth a deep entry covering the score calculation.
- **Capacity Outlook** — largest page, complex modeling. Worth a deep entry.

### Schema documentation

We mention the columns the codebase expects on `df_ns` and `df_drs` but haven't formally documented the contract. Worth adding a `docs/SCHEMAS.md` that lists every column, its type, and which pages depend on it.

### Test coverage

There are no automated tests. Smoke testing happens manually via the staging deploy. As the codebase grows, this becomes a liability. Consider adding pytest with synthetic frames for `assign_credits`, `calc_consultant_util`, and the loaders at minimum.

### Drift between shared and per-page constants

`EMPLOYEE_ROLES` is defined in both `shared/constants.py` and `pages/3_Utilization_Report.py` (and possibly other pages — should audit). When safe, consolidate to the shared version. Currently kept divergent because the page version is what's running.

### Numbering collision

Two pages named `9_*`: Help (`9_Help.py`) and Revenue Report (`9_Revenue_Report.py`). Streamlit's auto-ordering may behave unpredictably. Renumber when convenient.

---

*Last updated: by author of consolidation pass. When you make a substantive change, update the affected section.*
