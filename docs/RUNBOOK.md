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

### Hero → divider → legend box pattern

Pages with a hero banner and an explanatory legend box below it follow this three-part structure:

1. **Hero banner** — dark navy gradient (`linear-gradient(135deg,#1a56db 0%,#050D1F 55%,#050D1F 100%)`), rendered via `st.empty()` so metrics can be injected after data loads.
2. **Divider** — `<hr style="border:none;border-top:1px solid rgba(128,128,128,0.2);margin:20px 0">` (or the `.divider` CSS class where defined) separating the hero from page controls.
3. **Legend / how-to box** — shaded background with blue left border, matching pattern:

```python
st.markdown(
    "<div style='background:var(--color-background-secondary,rgba(59,158,255,0.05));"
    "border-left:4px solid #4472C4;border-radius:6px;"
    "padding:14px 18px;margin-bottom:14px;font-family:Manrope,sans-serif'>"
    "<div style='font-size:13px;font-weight:700;text-transform:uppercase;"
    "letter-spacing:.8px;color:#4472C4;margin-bottom:12px'>Legend title</div>"
    "<div style='display:flex;gap:20px;flex-wrap:wrap'>"
    # First column — no left border
    "<div style='flex:1;min-width:130px;padding-left:12px'>...</div>"
    # Subsequent columns — coloured left border matching pill colour
    "<div style='flex:1;min-width:130px;border-left:2px solid rgba(R,G,B,.4);padding-left:12px'>...</div>"
    "</div></div>",
    unsafe_allow_html=True
)
```

Pills inside legend boxes use `font-size:12px`, `padding:3px 10px`, `border-radius:20px`, and a `border:1px solid` matching the pill background colour (at `.4` alpha). Background opacity is `0.18`.

**Pages using this pattern:** Daily Briefing (how-to box), My Projects (how-to box), Utilization Report (credit tags legend), DRS Health Check (flag categories legend).


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

**Available hours helpers** (both defined locally in the page, not imported from `utils.py`):
- `get_avail_hours(region, period, employee=None)` — looks up `AVAIL_HOURS[region][period]`. If `employee` is in `CONTRACTOR_EMPLOYEES`, routes to the `"Contractor"` region regardless of the `region` arg.
- `_prorated_avail(region, period, employee=None)` — calls `get_avail_hours` then applies a working-day proration if `employee` has an exit date in `LEAVER_EXIT_DATES` that falls within `period`. Example: Arestarkhov exited June 9 → June avail = 176 × 7/22 bdays = 56hrs.

These helpers exist in both `3_Utilization_Report.py` (for the Streamlit display path) and `shared/utils.py` (for the Excel build path). Keep them in sync if the logic changes.

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
| Employee avail hours warning at page top | `EMPLOYEE_LOCATION` key doesn't match the NS name for that employee | Add both the roster key **and** the exact NS name variant to `EMPLOYEE_LOCATION` in `config.py`. If the NS name differs from the roster key, add both pointing to the same location. Also add an `NS_name_variant` entry to `PS_REGION_OVERRIDE` if region override applies. |
| PSPT avail hours don't match NS for a contractor | Contractor was using their country's public holiday calendar | Add the employee's roster key to `CONTRACTOR_EMPLOYEES` in `constants.py` |
| PSPT avail hours higher than NS for a mid-month leaver | `LEAVER_EXIT_DATES` has the exit date but proration isn't applied | Confirm the entry is in `LEAVER_EXIT_DATES` with a `"YYYY-MM-DD"` string (not `None`). The `_prorated_avail()` helper in `3_Utilization_Report.py` reads it at render time. |
| PSPT shows fewer monthly rows than NS for a consultant | Consultant had no time entries in that month — `emp_sum` is driven by NS rows | Known limitation: capacity rows are only generated for months with at least one time entry. See TODO section. |
| `ImportError: cannot import name 'CONTRACTOR_EMPLOYEES'` | `constants.py` was deployed without the other files, or vice versa | Always deploy `constants.py`, `utils.py`, and `3_Utilization_Report.py` together when roster/avail changes are made |

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
- `EMPLOYEE_LOCATION` — name → country/region mapping. Keys must match the name as it appears in NS time data, not necessarily the legal name. Where NS uses a different name variant (e.g. `"Cadeliña, Mark Enric"` vs roster key `"Cadelina, Macoy"`), add a second entry for the NS form — both pointing to the same location string.
- `PS_REGION_MAP`, `PS_REGION_OVERRIDE` — country → PS region grouping. Use `PS_REGION_OVERRIDE` for employees whose PS region doesn't follow from their country (e.g. Manila-based PMs reporting into NOAM).
- `AVAIL_HOURS` — region+month → available capacity hours. Keys are location strings matching `EMPLOYEE_LOCATION` values. Includes a `"Contractor"` key with raw Mon–Fri hours (no public holiday deductions) — see `CONTRACTOR_EMPLOYEES` in `constants.py`.
- `DEFAULT_SCOPE` — project_type → scoped hours dict (e.g. `"ZoneApp: Capture": 20.0`)

### `shared/constants.py`

Roles, employee roster, view-as resolver. Source of truth for who's allowed to see what.

Key exports:
- `MANAGERS_ONLY`, `MANAGER_CONSULTANTS`, `REPORTING_ONLY`, `NO_ACCESS` — role lists
- `EMPLOYEE_ROLES` — full roster dict (name → role/products/learning)
- `ACTIVE_EMPLOYEES`, `CONSULTANT_DROPDOWN` — filtered subsets (auto-derived; excludes `NO_ACCESS` and `_LEAVERS`)
- `_LEAVERS` — set of names excluded from active lists but kept in `EMPLOYEE_ROLES` for historical NS data joins
- `LEAVER_EXIT_DATES` — dict of `"Name": "YYYY-MM-DD"` for employees with a known exit date. Used by `_prorated_avail()` in `3_Utilization_Report.py` to scale the capacity denominator for mid-month leavers. Set value to `None` for leavers with unknown exit dates.
- `CONTRACTOR_EMPLOYEES` — set of names whose available hours should use the `"Contractor"` region (raw Mon–Fri, no public holiday deductions) regardless of their location. Currently: Dolha, Jordanova, Zoric. Add new contractors here — no other files need changing.
- `get_role(name) → str` — primary role resolver
- `is_manager(name) → bool`, `is_consultant(name) → bool` — convenience checks
- `resolve_view_as(...)` — used by all manager-aware pages
- `get_region_consultants(...)` — region filter helper
- `name_matches(a, b)` — fuzzy name match (handles "Last, First" vs "First Last")
- Column-mapping dicts: `MILESTONE_COLS_MAP`, `SS_COL_MAP`, `NS_COL_MAP`, `SFDC_COL_MAP`

**Adding a new employee**: Add to `EMPLOYEE_ROLES` with role, products, learning, util_exempt, util_target. Add to `EMPLOYEE_LOCATION` in `config.py` (use the name as it will appear in NS exports). If their location-based region is wrong, add a `PS_REGION_OVERRIDE` entry in `config.py`. If they're a contractor, add to `CONTRACTOR_EMPLOYEES`.

**Adding a leaver**: Add to `NO_ACCESS` (blocks login), `_LEAVERS` (excludes from dropdowns), and `LEAVER_EXIT_DATES` with the exit date. Keep in `EMPLOYEE_ROLES` — removing it breaks historical NS data joins.

### `shared/utils.py`

Excel report builder, credit assignment engine, capacity calc.

Key exports:
- `assign_credits(df, scope_map) → (df_with_credits, consumed_dict, skipped)` — the engine
- `build_excel(df, scope_map, consumed, df_drs=None) → BytesIO` — full multi-sheet Excel report. Pass `df_drs` to enable DRS-sourced PM attribution (DRS takes priority over NS; NS fills gaps where DRS PM is blank).
- `auto_detect_columns(df)` — fuzzy column-name resolver for varying NS exports
- `match_ff_task(task)` — task category resolver
- `get_avail_hours(region, period, employee=None)` — capacity lookup. Pass `employee` to enable contractor routing: if the name is in `CONTRACTOR_EMPLOYEES`, the `"Contractor"` region is used regardless of the `region` arg. All internal call sites pass `employee`; external callers should do the same.
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

### When PSPT avail hours don't match NS

Three distinct causes with different fixes:

**Consultant has fewer months than NS** — PSPT only generates capacity rows for months where the consultant logged at least one time entry. A consultant with zero hours in a given month gets no row for that month, making their period total look low. This is a known limitation, not a bug — the denominator is correct for the months shown. See TODO.

**Avail hours are wrong for a specific month** — likely a wrong or missing location in `EMPLOYEE_LOCATION`, or the NS name variant differs from the roster key. Check: (1) the employee appears in the avail hours warning banner at page top, (2) their key in `EMPLOYEE_LOCATION` matches exactly what NS exports for their name. Fix: add the NS name variant as a second key in both `EMPLOYEE_LOCATION` and (if needed) `PS_REGION_OVERRIDE` in `config.py`.

**Contractor shows lower avail than NS** — their country calendar deducts public holidays but NS uses raw working days. Fix: add the employee to `CONTRACTOR_EMPLOYEES` in `constants.py`. Currently: Dolha, Jordanova, Zoric.

**Mid-month leaver shows higher avail than NS** — PSPT is using full-month capacity instead of prorating. Fix: confirm the employee is in `LEAVER_EXIT_DATES` with a `"YYYY-MM-DD"` string (not `None`). The `_prorated_avail()` helper picks it up at render time — no code change needed if the date is populated.

---

### When PSPT billable hours are lower than the NS Quarterly Utilization Report

**Root cause: the NS Quarterly Utilization Report has no effective approval status filter.**

Confirmed June 2026: the NS report filter set contains date range, generic resource exclusion, department, customer/project, resource name, class, subsidiary, and internal project flag — but no `Approval Status` filter. This means the NS report counts **all time entries regardless of status**, including Rejected entries that have been explicitly declined by a manager.

PSPT reads from `customsearch66732` (Time Detail), which filters on approved/submitted time only. PSPT figures will therefore be **lower** than the NS Quarterly Utilization Report in any period where rejected time entries exist. This is correct behaviour — PSPT is more accurate, not less.

**How to confirm:** drill into the NS report for the consultant and period showing the discrepancy. Cross-reference against the Time Detail saved search — rejected entries visible there but not in PSPT confirm this is the cause.

**Attempted fix and known limitation (June 2026):** Adding `Time Tracked: Approval Status is not equal to Rejected` to the report filters had no effect. The relevant field in this report appears to be `Payroll Time: Approval Status`, not `Time Tracked: Approval Status` — a component mismatch that means the filter applies to a different data join and doesn't suppress rejected time entries. This is a NetSuite report architecture issue. NS admin or NetSuite support would need to investigate the correct component to filter on.

**Do not adjust PSPT** to match NS on this discrepancy. The NS native figure is the one that is wrong. The Time Detail saved search (`customsearch66732`) is the reliable source for reconciliation.

---

### When a project shows the wrong consultant in PSPT

**Root cause (fixed June 2026):** per-project lookup dicts (`proj_pm`, `proj_cust_region`, `proj_ps_region`, `proj_start`, `proj_phase`) were keyed by project name. When the same customer has two concurrent projects with different types (e.g. Acer Europe AG — Approvals and Acer Europe AG — Capture), both collapsed to one key and `.first()` picked whichever employee sorted first alphabetically.

**Fix applied:** all per-project dicts in `shared/utils.py` are now keyed by `project_id` (with fallback to project name if `project_id` is absent from the export). Project ID is unique per project in NS and is the correct grouping key throughout.

**If you see wrong consultant attribution after this fix:** check that the NS Time Detail export includes the `Project ID` column. If that column is missing from the export, PSPT falls back to project name and the collision risk returns. The export schema should always include Project ID — verify `customsearch66732` has it in the column list.

**Do not re-key any per-project dict by project name.** Project name is not unique. Always use `project_id`.

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

### Capacity rows missing for zero-entry months

`emp_sum` in both the Utilization Report page and `build_excel` is built by `groupby(["employee", "period"])` on NS time entries. Consultants with no hours logged in a given month get no capacity row for that month, so their period total understates available hours when comparing against NS (which always shows full-quarter capacity). 

Fix approach: after building `emp_sum`, expand it to include one row per active employee per period in the selected range, filling missing rows with zero hours and the correct avail hours. Requires knowing which employees were active in each period — cross-reference `LEAVER_EXIT_DATES` and `ACTIVE_EMPLOYEES`. Medium complexity, ~half-day work.

### Numbering collision

Two pages named `9_*`: Help (`9_Help.py`) and Revenue Report (`9_Revenue_Report.py`). Streamlit's auto-ordering may behave unpredictably. Renumber when convenient.

### Roster & configuration as database records (target state)

**Problem:** Employee names, locations, regions, products, start/exit dates, and capacity benchmarks are currently hardcoded in `shared/constants.py` and `shared/config.py`. Every roster change (new hire, leaver, location update, product reassignment) requires a code edit and a deployment. This creates drift risk, is error-prone, and is not scalable as the team grows.

**Target state:** All roster and configuration data lives in a persistent database. An admin page in PSPT (director/manager-only) provides a UI to add, edit, and deactivate employees without touching code. Deploys are only needed for logic changes, not data changes.

**Proposed two-phase approach:**

**Phase 1 — Google Sheets bridge (near-term, pre-hosting)**

Use a Google Sheet as a lightweight roster store. PSPT reads it at load time via the Drive API (already in the IT approval pipeline). The sheet is the editable source; `constants.py` and `config.py` become fallback/cache only.

- Schema: one row per employee with columns for name, NS name variant, role, region, location, products (comma-separated), start date, exit date, is_contractor, util_exempt, ps_region_override
- PSPT reads the sheet on Home load, builds the same in-memory dicts currently loaded from Python files
- If the sheet is unavailable, falls back to the hardcoded files with a warning banner
- No new IT approvals required beyond the existing Google Workspace submission
- Roster change = edit the sheet, refresh the app. No deployment.
- Estimated complexity: medium (~1 day). Dependency: Google Drive API IT approval.

**Phase 2 — Postgres database (target state, post-hosting)**

Once on IT-approved hosting with a persistent DB:

- Replace the Google Sheet with a Postgres schema (or IT-approved equivalent)
- Add an admin page in PSPT with forms to add/edit/deactivate employees and update location/capacity data
- `constants.py` and `config.py` are retired or become thin stubs that load from the DB at startup
- Leavers are deactivated via the UI (exit date stamped); `NO_ACCESS`, `_LEAVERS`, `CONTRACTOR_EMPLOYEES` become DB flags
- Requires IT architecture review (employee data in a new persistent store is a new data classification surface — similar process to the current Google Workspace/Claude submission)
- Estimated complexity: large (~3-5 days including schema design, admin page, migration). Dependency: IT-approved hosting, DB provisioning, IT review.

**DB schema (draft — for Phase 2 design)**

```
employees
  id, name, ns_name_variant, role, region, location,
  products[], start_date, exit_date, is_contractor,
  util_exempt, util_target, ps_region_override, active

avail_hours
  location, year_month, available_hours

product_benchmarks
  product_name, avg_hrs, scope_hrs, timeline_weeks, oh_hrswk
```

**Current state of hardcoded data (all need migrating):**
- `shared/constants.py` — `EMPLOYEE_ROLES`, `ACTIVE_EMPLOYEES`, `CONTRACTOR_EMPLOYEES`, `LEAVER_EXIT_DATES`, `NO_ACCESS`, `_LEAVERS`
- `shared/config.py` — `EMPLOYEE_LOCATION`, `PS_REGION_OVERRIDE`, `AVAIL_HOURS`, `DEFAULT_SCOPE`, `PRODUCT_DATA`, `PHASE_END_WEEKS`

Until Phase 1 or 2 is implemented, the rule is: **all roster and configuration changes go through `constants.py` and `config.py` only** — never redefined in individual page files. All page files import from shared; they never hardcode names or product data directly.

---

*Last updated: June 2026 — roster updates (new hires, leavers), contractor avail hours routing, mid-month leaver proration, avail hours reconciliation playbook, NS rejected time entry finding, project attribution bug fix (project_id keying), hardcoded roster/product data centralised to shared/constants.py and shared/config.py, roster-as-records roadmap added. When you make a substantive change, update the affected section.*
