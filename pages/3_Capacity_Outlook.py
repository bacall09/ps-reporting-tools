"""
Capacity Outlook
Upload SS DRS + NS Unassigned Projects to project consultant availability
and surface demand from unassigned (closed) deals.
"""
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Capacity Outlook", page_icon=None, layout="wide")

# ── Colours ───────────────────────────────────────────────────────────────────
NAVY   = "1e2c63"
WHITE  = "FFFFFF"
LTGRAY = "F2F2F2"
RED    = "E74C3C"
GREEN  = "27AE60"
AMBER  = "F39C12"
BLUE   = "4472C4"

# ── Employee config (mirrors other pages) ─────────────────────────────────────
EMPLOYEE_ROLES = {
    # ── Project Managers (no product delivery) ────────────────────────────────
    "Barrio, Nairobi":        {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    "Hughes, Madalyn":        {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    "Porangada, Suraj":       {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    "Cadelina":               {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    # ── Solution Architects ───────────────────────────────────────────────────
    "Bell, Stuart":           {"role": "Solution Architect", "products": ["Billing"],                                                                              "learning": []},
    "DiMarco, Nicole R":      {"role": "Solution Architect", "products": ["Billing"],                                                                              "learning": []},
    "Murphy, Conor":          {"role": "Solution Architect", "products": ["Billing"],                                                                              "learning": [], "util_exempt": True},
    # ── Developer ─────────────────────────────────────────────────────────────
    "Church, Jason G":        {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
    # ── Consultants ───────────────────────────────────────────────────────────
    "Arestarkhov, Yaroslav":  {"role": "Consultant",         "products": ["Billing", "Capture"],                                                                   "learning": []},
    "Carpen, Anamaria":       {"role": "Consultant",         "products": ["Capture", "Approvals", "e-Invoicing"],                                                  "learning": []},
    "Centinaje, Rhodechild":  {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"], "learning": []},
    "Cooke, Ellen":           {"role": "Consultant",         "products": ["Billing", "Payroll"],                                                                   "learning": []},
    "Dolha, Madalina":        {"role": "Consultant",         "products": ["Capture", "Reconcile", "CC Statement Import", "Reconcile PSP", "e-Invoicing"],          "learning": []},
    "Finalle-Newton, Jesse":  {"role": "Solution Architect", "products": ["Reporting"],                                                                            "learning": []},
    "Gardner, Cheryll L":     {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Hopkins, Chris":         {"role": "Consultant",         "products": ["Capture", "Approvals"],                                                                 "learning": []},
    "Ickler, Georganne":      {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Isberg, Eric":           {"role": "Consultant",         "products": ["Reporting"],                                                                            "learning": []},
    "Jordanova, Marija":      {"role": "Consultant",         "products": ["Approvals", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"],     "learning": []},
    "Lappin, Thomas":         {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": ["Capture", "Reconcile"]},
    "Longalong, Santiago":    {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile"],                                                    "learning": ["Billing"]},
    "Mohammad, Manaan":       {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile"],                                                    "learning": []},
    "Morris, Lisa":           {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "NAQVI, SYED":            {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "Olson, Austin D":        {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Pallone, Daniel":        {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "Raykova, Silvia":        {"role": "Consultant",         "products": ["Capture", "Approvals", "e-Invoicing"],                                                  "learning": []},
    "Selvakumar, Sajithan":   {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile"],                                                    "learning": []},
    "Snee, Stefanie J":       {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Swanson":                {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": [], "util_exempt": True},
    "Tuazon, Carol":          {"role": "Consultant",         "products": ["Payroll", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"],       "learning": []},
    "Zoric, Ivan":            {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"], "learning": []},
    "Dunn, Steven":           {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
    "Law, Brandon":           {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
    "Quiambao, Generalyn":    {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
        # ── Leavers (historical) ──────────────────────────────────────────────────
    "Alam, Laisa":            {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Chan, Joven":            {"role": "Consultant",         "products": ["Capture"],                                                                              "learning": []},
    "Cloete, Bronwyn":        {"role": "Consultant",         "products": ["Capture", "Approvals"],                                                                 "learning": []},
    "Eyong, Eyong":           {"role": "Consultant",         "products": ["Capture"],                                                                              "learning": []},
    "Hamilton, Julie C":      {"role": "Consultant",         "products": ["Reporting"],                                                                            "learning": []},
    "Hernandez, Camila":      {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Rushbrook, Emma C":      {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "Strauss, John W":        {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
}

EMPLOYEE_LOCATION = {
    "Arestarkhov, Yaroslav":  ("Czech Republic",      None,       None),
    "Barrio, Nairobi":        ("USA",                 None,       None),
    "Bell, Stuart":           ("USA",                 None,       None),
    "Cadelina":               ("Manila (PH)",         "2026-03",  None),
    "Carpen, Anamaria":       ("Spain",               None,       None),
    "Centinaje, Rhodechild":  ("Manila (PH)",         None,       None),
    "Church, Jason G":        ("USA",                 None,       None),
    "Cooke, Ellen":           ("Northern Ireland",    None,       None),
    "Cruz, Daniel":           ("Manila (PH)",         None,       None),
    "DiMarco, Nicole R":      ("USA",                 None,       None),
    "Dolha, Madalina":        ("Faroe Islands",       None,       None),
    "Finalle-Newton, Jesse":  ("USA",                 None,       None),
    "Gardner, Cheryll L":     ("USA",                 None,       None),
    "Hopkins, Chris":         ("USA",                 None,       None),
    "Hughes, Madalyn":        ("USA",                 None,       None),
    "Ickler, Georganne":      ("USA",                 None,       None),
    "Isberg, Eric":           ("USA",                 None,       None),
    "Jordanova, Marija":      ("North Macedonia",     None,       None),
    "Lappin, Thomas":         ("Northern Ireland",    None,       None),
    "Longalong, Santiago":    ("Manila (PH)",         None,       None),
    "Mohammad, Manaan":       ("Canada",              None,       None),
    "Morris, Lisa":           ("Sydney (NSW)",        None,       None),
    "Murphy, Conor":          ("USA",                 None,       None),
    "NAQVI, SYED":            ("Canada",              None,       None),
    "Olson, Austin D":        ("USA",                 None,       None),
    "Pallone, Daniel":        ("Sydney (NSW)",        None,       None),
    "Porangada, Suraj":       ("USA",                 None,       None),
    "Raykova, Silvia":        ("Netherlands",         None,       None),
    "Selvakumar, Sajithan":   ("Canada",              None,       None),
    "Snee, Stefanie J":       ("USA",                 None,       None),
    "Stone, Matt":            ("USA",                 None,       None),
    "Swanson":                ("USA",                 None,       None),
    "Tuazon, Carol":          ("Manila (PH)",         None,       None),
    "Zoric, Ivan":            ("Serbia",              None,       None),
    "Alam, Laisa":            ("USA",                 None,       "2025-12"),
    "Chan, Joven":            ("Manila (PH)",         None,       "2025-12"),
    "Cloete, Bronwyn":        ("Netherlands",         None,       "2026-02"),
    "Eyong, Eyong":           ("USA",                 None,       "2025-12"),
    "Hamilton, Julie C":      ("USA",                 None,       "2026-01"),
    "Hernandez, Camila":      ("USA",                 None,       "2025-12"),
    "Rushbrook, Emma C":      ("Wales",               None,       "2025-12"),
    "Strauss, John W":        ("USA",                 None,       "2026-03"),
}

PS_REGION_OVERRIDE = {
    "NAQVI, SYED":       "EMEA",
    "Cruz, Daniel":      "NOAM",
    "Chan, Joven":       "NOAM",
    "Rushbrook, Emma C": "EMEA",
}
PS_REGION_MAP = {
    "Sydney (NSW)":     "APAC",
    "Manila (PH)":      "APAC",
    "UK":               "EMEA",
    "Wales":            "EMEA",
    "Spain":            "EMEA",
    "Netherlands":      "EMEA",
    "Northern Ireland": "EMEA",
    "Faroe Islands":    "EMEA",
    "North Macedonia":  "EMEA",
    "Czech Republic":   "EMEA",
    "Serbia":           "EMEA",
    "USA":              "NOAM",
    "Canada":           "NOAM",
}

# ── Phase duration table — end week per phase per product type ─────────────────
# Values = week number from project start when phase completes
PHASE_END_WEEKS = {
    "Approvals": {
        "01. requirements and design": 1,
        "02. configuration":           2,
        "03. enablement/training":     3,
        "04. uat":                     6,
        "05. prep for go-live":        6,
        "06. go-live":                 8,
        "08. ready for support transition": 8,
    },
    "Capture": {
        "01. requirements and design": 1,
        "02. configuration":           2,
        "03. enablement/training":     3,
        "04. uat":                     6,
        "05. prep for go-live":        6,
        "06. go-live":                 8,
        "08. ready for support transition": 8,
    },
    "Capture & e-Invoicing": {
        "01. requirements and design": 1,
        "02. configuration":           4,
        "03. enablement/training":     3,
        "04. uat":                     9,
        "05. prep for go-live":        6,
        "06. go-live":                 10,
        "08. ready for support transition": 10,
    },
    "Payments": {
        "01. requirements and design": 1,
        "02. configuration":           2,
        "03. enablement/training":     3,
        "04. uat":                     8,
        "05. prep for go-live":        8,
        "06. go-live":                 10,
        "08. ready for support transition": 10,
    },
    "Reconcile": {
        "01. requirements and design": 1,
        "02. configuration":           2,
        "03. enablement/training":     3,
        "04. uat":                     7,
        "05. prep for go-live":        8,
        "06. go-live":                 10,
        "08. ready for support transition": 10,
    },
    "Reconcile PSP": {
        "01. requirements and design": 1,
        "02. configuration":           2,
        "03. enablement/training":     3,
        "04. uat":                     7,
        "05. prep for go-live":        8,
        "06. go-live":                 10,
        "08. ready for support transition": 10,
    },
    "e-Invoicing": {
        "01. requirements and design": 1,
        "02. configuration":           4,
        "03. enablement/training":     3,
        "04. uat":                     9,
        "06. go-live":                 10,
        "08. ready for support transition": 10,
    },
    "SFTP Connector": {
        "01. requirements and design": 1,
        "02. configuration":           2,
        "03. enablement/training":     3,
        "04. uat":                     7,
        "05. prep for go-live":        8,
        "06. go-live":                 10,
        "08. ready for support transition": 10,
    },
    "CC Statement Import": {
        "01. requirements and design": 1,
        "02. configuration":           2,
        "03. enablement/training":     3,
        "04. uat":                     7,
        "05. prep for go-live":        8,
        "06. go-live":                 10,
        "08. ready for support transition": 10,
    },
}

# Default total duration (weeks) for unknown product types
DEFAULT_DURATION_WEEKS = 10

# Phase weight for parallel load assessment
PHASE_WEIGHTS = {
    "00. onboarding":                    1.0,
    "01. requirements and design":       1.0,
    "02. configuration":                 2.0,
    "03. enablement/training":           2.5,
    "04. uat":                           3.0,
    "05. prep for go-live":              1.0,
    "06. go-live":                       3.0,
    "07. data migration":                1.0,
    "08. ready for support transition":  0.5,
    "09. phase 2 scoping":               1.0,
    "10. complete/pending final billing": 0.0,
    "11. on hold":                       0.25,
    "12. ps review":                     0.25,
}
PARALLEL_CONFLICT_THRESHOLD = 5.0  # sum of phase weights across concurrent projects

# FF scope map (mirrors utilization report)
FF_SCOPE_MAP = {
    "capture & e-invoicing": 40,
    "e-invoicing":           40,
    "cc statement import":   20,
    "reconcile psp":         20,
    "sftp connector":        20,
    "capture":               20,
    "approvals":             20,
    "payments":              40,
    "reconcile":             40,
    "reporting":             40,
    "billing":               80,
    "payroll":               80,
}

HORIZON_MONTHS = 6  # how many months forward to project

# ── Helper functions ───────────────────────────────────────────────────────────
def _emp_location(name):
    v = EMPLOYEE_LOCATION.get(name)
    if v is None:
        return None
    return v[0] if isinstance(v, tuple) else v

def _emp_ps_region(name):
    if name in PS_REGION_OVERRIDE:
        return PS_REGION_OVERRIDE[name]
    loc = _emp_location(name)
    return PS_REGION_MAP.get(loc, "Unknown") if loc else "Unknown"

def _emp_active(name, period_str):
    v = EMPLOYEE_LOCATION.get(name)
    if v is None or not isinstance(v, tuple):
        return True
    _, start, end = v
    try:
        period = str(period_str)[:7]
        if start and period < start:
            return False
        if end and period > end:
            return False
    except Exception:
        pass
    return True

def _is_delivery_role(name):
    """True if the employee owns product delivery (not a pure PM)."""
    info = EMPLOYEE_ROLES.get(name, {})
    return info.get("role") not in ("Project Manager",)

def _resolve_ff_scope(project_type):
    """Return scoped hours for a FF project type."""
    pt = (project_type or "").strip().lower()
    for key in sorted(FF_SCOPE_MAP.keys(), key=len, reverse=True):
        if key in pt:
            return FF_SCOPE_MAP[key]
    return None

def _project_end_date(start_date, project_type):
    """Return estimated project completion date from start date + phase duration table."""
    pt = (project_type or "").strip()
    phase_map = PHASE_END_WEEKS.get(pt)
    if phase_map:
        total_weeks = max(phase_map.values())
    else:
        # Try case-insensitive match
        for k, v in PHASE_END_WEEKS.items():
            if k.lower() == pt.lower():
                total_weeks = max(v.values())
                break
        else:
            total_weeks = DEFAULT_DURATION_WEEKS
    return start_date + timedelta(weeks=total_weeks)

def _current_phase_end_date(start_date, current_phase, project_type):
    """Return estimated end date of the current phase."""
    pt = (project_type or "").strip()
    phase_map = None
    for k, v in PHASE_END_WEEKS.items():
        if k.lower() == pt.lower():
            phase_map = v
            break
    if phase_map is None:
        phase_map = {}

    phase_key = (current_phase or "").strip().lower()
    # Find this phase's end week
    phase_end_wk = phase_map.get(phase_key)
    if phase_end_wk is None:
        # Use total duration as fallback
        phase_end_wk = max(phase_map.values()) if phase_map else DEFAULT_DURATION_WEEKS

    return start_date + timedelta(weeks=phase_end_wk)

def _months_range(n=HORIZON_MONTHS):
    """Return list of (year, month) tuples for the next n months from today."""
    today = date.today()
    months = []
    for i in range(n):
        d = today + relativedelta(months=i)
        months.append((d.year, d.month))
    return months

def _month_label(year, month):
    return datetime(year, month, 1).strftime("%b %Y")

def _territory_to_ps_region(territory):
    """Map a territory string to NOAM / EMEA / APAC."""
    t = (territory or "").strip().upper()
    if "NOAM" in t or "NORTH AM" in t or "US" in t or "CANADA" in t:
        return "NOAM"
    if "EMEA" in t:
        return "EMEA"
    if "APAC" in t or "ASIA" in t or "PACIFIC" in t:
        return "APAC"
    return t if t else "Unknown"

def fmt_hrs(n):
    return f"{n:,.2f}".rstrip("0").rstrip(".") if "." in f"{n:,.2f}" else f"{n:,}"

# ── SS DRS loader ──────────────────────────────────────────────────────────────
SS_COL_MAP = {
    "project id":        "project_id",
    "project name":      "project_name",
    "name":              "project_name",
    "project":           "project_name",
    "consultant":        "consultant",
    "assigned to":       "consultant",
    "resource":          "consultant",
    "phase":             "phase",
    "project phase":     "phase",
    "milestone":         "phase",
    "project type":      "project_type",
    "type":              "project_type",
    "billing type":      "billing_type",
    "status":            "status",
    "project status":    "status",
    "territory":         "territory",
    "project manager":   "project_manager",
    "pm":                "project_manager",
    "start date":        "start_date",
    "project start date":"start_date",
    "client health":     "client_health",
    "risk":              "risk",
    "risk flag":         "risk",
}

def load_ss(file):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    df.columns = [c.strip().lower() for c in df.columns]
    rename = {}
    for col in df.columns:
        if col in SS_COL_MAP:
            rename[col] = SS_COL_MAP[col]
    df = df.rename(columns=rename)
    # Normalise phase
    if "phase" in df.columns:
        df["phase"] = df["phase"].astype(str).str.strip().str.lower()
    # Parse start_date
    if "start_date" in df.columns:
        df["start_date"] = pd.to_datetime(df["start_date"], errors="coerce")
    # Filter FF only for projection (T&M in unassigned report)
    if "billing_type" in df.columns:
        df = df[df["billing_type"].astype(str).str.strip().str.lower() == "fixed fee"]
    # Exclude complete / inactive phases from active projection
    inactive = {"10. complete/pending final billing", "11. on hold", "12. ps review"}
    if "phase" in df.columns:
        df = df[~df["phase"].isin(inactive)]
    return df

# ── NS Unassigned Projects loader ──────────────────────────────────────────────
NS_COL_MAP = {
    "project id":          "project_id",
    "internal id":         "project_id",
    "project":             "project_name",
    "name":                "project_name",
    "territory":           "territory",
    "status":              "status",
    "billing type":        "billing_type",
    "project type":        "project_type",
    "signed date":         "signed_date",
    "project outreach":    "outreach_date",
    "start date":          "start_date",
    "t&m scope":           "tm_scope",
    "orig estimated work": "tm_scope",
}

def load_ns_unassigned(file):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    df.columns = [c.strip().lower() for c in df.columns]
    rename = {}
    for col in df.columns:
        if col in NS_COL_MAP:
            rename[col] = NS_COL_MAP[col]
    df = df.rename(columns=rename)
    # Parse dates
    for dcol in ["signed_date", "outreach_date", "start_date"]:
        if dcol in df.columns:
            df[dcol] = pd.to_datetime(df[dcol], errors="coerce")
    # Numeric T&M scope
    if "tm_scope" in df.columns:
        df["tm_scope"] = pd.to_numeric(df["tm_scope"], errors="coerce")
    # Derive scoped hours
    df["scoped_hours"] = df.apply(lambda r: (
        r.get("tm_scope") if str(r.get("billing_type", "")).strip().lower() == "time & material"
        else _resolve_ff_scope(str(r.get("project_type", "")))
    ), axis=1)
    # Derive PS region from territory
    if "territory" in df.columns:
        df["ps_region"] = df["territory"].apply(_territory_to_ps_region)
    return df

# ── Projection engine ─────────────────────────────────────────────────────────
def project_consultant_availability(ss_df, months):
    """
    For each consultant, determine which months they are:
    - BUSY (active project in that month)
    - FREE (no active project)
    Returns dict: {consultant: {(year, month): 'busy'|'free'|'partial'}}
    Also returns conflict flags for parallel heavy-phase projects.
    """
    today = date.today()
    results = {}    # consultant -> {month_tuple: status}
    conflicts = []  # list of conflict dicts
    project_spans = {}  # consultant -> list of (project, start, end, phase, weight)

    if ss_df is None or ss_df.empty:
        return results, conflicts

    required = {"consultant", "start_date", "phase"}
    if not required.issubset(ss_df.columns):
        return results, conflicts

    for _, row in ss_df.iterrows():
        consultant = str(row.get("consultant", "")).strip()
        if not consultant or consultant.lower() in ("nan", "none", ""):
            continue
        # Skip pure PMs
        if not _is_delivery_role(consultant):
            continue

        start_dt = row.get("start_date")
        if pd.isna(start_dt):
            continue
        start_d = start_dt.date() if hasattr(start_dt, "date") else start_dt

        phase = str(row.get("phase", "")).strip().lower()
        project_type = str(row.get("project_type", "")).strip()
        project_name = str(row.get("project_name", "")).strip()

        # Estimate project end date from current phase
        end_d = _current_phase_end_date(start_d, phase, project_type)
        # Also get full project end
        full_end_d = _project_end_date(start_d, project_type)

        phase_weight = PHASE_WEIGHTS.get(phase, 1.0)

        if consultant not in project_spans:
            project_spans[consultant] = []
        project_spans[consultant].append({
            "project": project_name,
            "project_type": project_type,
            "phase": phase,
            "phase_weight": phase_weight,
            "start": start_d,
            "phase_end": end_d,
            "project_end": full_end_d,
        })

    # Detect parallel conflicts
    for consultant, spans in project_spans.items():
        if len(spans) > 1:
            for i, s1 in enumerate(spans):
                for s2 in spans[i+1:]:
                    # Check overlap
                    overlap_start = max(s1["start"], s2["start"])
                    overlap_end = min(s1["phase_end"], s2["phase_end"])
                    if overlap_start <= overlap_end:
                        combined_weight = s1["phase_weight"] + s2["phase_weight"]
                        if combined_weight >= PARALLEL_CONFLICT_THRESHOLD:
                            conflicts.append({
                                "consultant": consultant,
                                "project_1": s1["project"],
                                "phase_1": s1["phase"],
                                "project_2": s2["project"],
                                "phase_2": s2["phase"],
                                "combined_weight": combined_weight,
                                "overlap_start": overlap_start,
                                "overlap_end": overlap_end,
                            })

    # Build monthly availability grid
    for consultant, spans in project_spans.items():
        results[consultant] = {}
        for (yr, mo) in months:
            month_start = date(yr, mo, 1)
            month_end = (month_start + relativedelta(months=1)) - timedelta(days=1)
            busy_projects = [
                s for s in spans
                if s["start"] <= month_end and s["project_end"] >= month_start
            ]
            if not busy_projects:
                results[consultant][(yr, mo)] = "free"
            else:
                # Check if all projects are in low-weight phases
                max_weight = max(s["phase_weight"] for s in busy_projects)
                if max_weight <= 1.0:
                    results[consultant][(yr, mo)] = "partial"  # light load
                else:
                    results[consultant][(yr, mo)] = "busy"

    return results, conflicts

def build_demand_summary(ns_df, months):
    """Summarise unassigned project demand by region and month (based on planned start date)."""
    if ns_df is None or ns_df.empty:
        return {}

    demand = {}  # region -> {month_tuple: {"count": n, "hours": h, "projects": []}}
    for _, row in ns_df.iterrows():
        region = str(row.get("ps_region", "Unknown")).strip()
        start_dt = row.get("start_date")
        if pd.isna(start_dt) or start_dt is None:
            # Use outreach date as proxy
            start_dt = row.get("outreach_date")
        if pd.isna(start_dt) or start_dt is None:
            continue
        start_d = start_dt.date() if hasattr(start_dt, "date") else start_dt
        mo_key = (start_d.year, start_d.month)
        if mo_key not in [m for m in months]:
            continue
        scoped_hrs = row.get("scoped_hours") or 0
        project_name = str(row.get("project_name", "")).strip()
        project_type = str(row.get("project_type", "")).strip()

        if region not in demand:
            demand[region] = {}
        if mo_key not in demand[region]:
            demand[region][mo_key] = {"count": 0, "hours": 0, "projects": []}
        demand[region][mo_key]["count"] += 1
        demand[region][mo_key]["hours"] += scoped_hrs
        demand[region][mo_key]["projects"].append(f"{project_name} ({project_type})")

    return demand

# ── Excel export ───────────────────────────────────────────────────────────────
def build_excel(availability, conflicts, ns_df, months, ss_df):
    wb = Workbook()
    navy_fill   = PatternFill("solid", fgColor=NAVY)
    green_fill  = PatternFill("solid", fgColor="C6EFCE")
    amber_fill  = PatternFill("solid", fgColor="FFEB9C")
    red_fill    = PatternFill("solid", fgColor="FFC7CE")
    blue_fill   = PatternFill("solid", fgColor="BDD7EE")
    gray_fill   = PatternFill("solid", fgColor=LTGRAY)
    white_fill  = PatternFill("solid", fgColor=WHITE)
    hdr_font    = Font(bold=True, color=WHITE, name="Calibri", size=11)
    bold_font   = Font(bold=True, name="Calibri", size=10)
    std_font    = Font(name="Calibri", size=10)
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left        = Alignment(horizontal="left",   vertical="center")
    thin        = Side(style="thin", color="CCCCCC")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr_row(ws, row, values, fill=None):
        for col, val in enumerate(values, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = hdr_font
            c.fill = fill or navy_fill
            c.alignment = center
            c.border = border

    def data_cell(ws, row, col, val, fill=None, bold=False, align=None):
        c = ws.cell(row=row, column=col, value=val)
        c.font = bold_font if bold else std_font
        if fill:
            c.fill = fill
        c.alignment = align or left
        c.border = border
        return c

    month_labels = [_month_label(y, m) for y, m in months]

    # ── Tab 1: Heatmap ─────────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Capacity Heatmap"
    hdr_row(ws1, 1, ["Consultant", "Region", "Role"] + month_labels)
    ws1.column_dimensions["A"].width = 28
    ws1.column_dimensions["B"].width = 10
    ws1.column_dimensions["C"].width = 18
    for col in range(4, 4 + len(months)):
        ws1.column_dimensions[get_column_letter(col)].width = 14

    row_n = 2
    for consultant in sorted(availability.keys()):
        region = _emp_ps_region(consultant)
        role_info = EMPLOYEE_ROLES.get(consultant, {})
        role = role_info.get("role", "Consultant")
        data_cell(ws1, row_n, 1, consultant, bold=True)
        data_cell(ws1, row_n, 2, region, align=center)
        data_cell(ws1, row_n, 3, role, align=center)
        for col_i, mo in enumerate(months, 4):
            status = availability[consultant].get(mo, "free")
            fill = green_fill if status == "free" else amber_fill if status == "partial" else red_fill
            label = "Available" if status == "free" else "Light" if status == "partial" else "Busy"
            data_cell(ws1, row_n, col_i, label, fill=fill, align=center)
        row_n += 1

    # ── Tab 2: Consultant Detail ───────────────────────────────────────────────
    ws2 = wb.create_sheet("Consultant Detail")
    hdr_row(ws2, 1, ["Consultant", "Region", "Role", "Products", "Current Projects",
                     "Current Phase", "Est. Free From", "Learning (Future)"])
    for col, w in zip("ABCDEFGH", [28, 10, 18, 35, 35, 25, 15, 30]):
        ws2.column_dimensions[get_column_letter(col)].width = w

    row_n = 2
    if ss_df is not None and not ss_df.empty:
        active_consultants = ss_df["consultant"].dropna().unique() if "consultant" in ss_df.columns else []
    else:
        active_consultants = []

    for name, info in sorted(EMPLOYEE_ROLES.items(), key=lambda x: x[0]):
        if info.get("role") == "Project Manager":
            continue
        if not _emp_active(name, datetime.today().strftime("%Y-%m")):
            continue
        region = _emp_ps_region(name)
        role = info.get("role", "Consultant")
        products = ", ".join(info.get("products", []) or ["All"])
        learning = ", ".join(info.get("learning", []))

        # Find their active projects from SS
        if ss_df is not None and "consultant" in ss_df.columns:
            emp_rows = ss_df[ss_df["consultant"].str.strip() == name]
            projects = "; ".join(emp_rows["project_name"].dropna().unique()) if "project_name" in emp_rows.columns else ""
            phases   = "; ".join(emp_rows["phase"].dropna().unique()) if "phase" in emp_rows.columns else ""
            # Estimate free from
            free_dates = []
            for _, r in emp_rows.iterrows():
                sd = r.get("start_date")
                if pd.notna(sd):
                    sd = sd.date() if hasattr(sd, "date") else sd
                    pt = str(r.get("project_type", "")).strip()
                    free_dates.append(_project_end_date(sd, pt))
            free_from = max(free_dates).strftime("%b %Y") if free_dates else "Available Now"
        else:
            projects = ""
            phases = ""
            free_from = "Available Now"

        fill = gray_fill if not projects else white_fill
        data_cell(ws2, row_n, 1, name, bold=True, fill=fill)
        data_cell(ws2, row_n, 2, region, fill=fill, align=center)
        data_cell(ws2, row_n, 3, role, fill=fill)
        data_cell(ws2, row_n, 4, products, fill=fill)
        data_cell(ws2, row_n, 5, projects, fill=fill)
        data_cell(ws2, row_n, 6, phases, fill=fill)
        data_cell(ws2, row_n, 7, free_from, fill=fill, align=center)
        data_cell(ws2, row_n, 8, learning, fill=fill)
        row_n += 1

    # ── Tab 3: Regional Rollup ─────────────────────────────────────────────────
    ws3 = wb.create_sheet("Regional Rollup")
    regions = ["NOAM", "EMEA", "APAC"]
    hdr_row(ws3, 1, ["Region", "Metric"] + month_labels)
    ws3.column_dimensions["A"].width = 10
    ws3.column_dimensions["B"].width = 22
    for col in range(3, 3 + len(months)):
        ws3.column_dimensions[get_column_letter(col)].width = 14

    demand = build_demand_summary(ns_df, months)
    row_n = 2
    for region in regions:
        # Capacity: count of consultants available in each month
        region_consultants = [
            c for c in availability
            if _emp_ps_region(c) == region
        ]
        # Available count
        data_cell(ws3, row_n, 1, region, bold=True, fill=blue_fill, align=center)
        data_cell(ws3, row_n, 2, "Consultants Available", bold=True)
        for col_i, mo in enumerate(months, 3):
            avail = sum(1 for c in region_consultants
                       if availability[c].get(mo, "free") in ("free", "partial"))
            fill = green_fill if avail > 2 else amber_fill if avail > 0 else red_fill
            data_cell(ws3, row_n, col_i, avail, fill=fill, align=center)
        row_n += 1

        # Demand: unassigned project count
        data_cell(ws3, row_n, 1, "", fill=blue_fill)
        data_cell(ws3, row_n, 2, "Unassigned Projects", bold=True)
        reg_demand = demand.get(region, {})
        for col_i, mo in enumerate(months, 3):
            cnt = reg_demand.get(mo, {}).get("count", 0)
            fill = red_fill if cnt > 2 else amber_fill if cnt > 0 else green_fill
            data_cell(ws3, row_n, col_i, cnt, fill=fill, align=center)
        row_n += 1

        # Demand: scoped hours
        data_cell(ws3, row_n, 1, "", fill=blue_fill)
        data_cell(ws3, row_n, 2, "Scoped Hours Demand")
        for col_i, mo in enumerate(months, 3):
            hrs = reg_demand.get(mo, {}).get("hours", 0)
            data_cell(ws3, row_n, col_i, round(hrs, 1) if hrs else 0, align=center)
        row_n += 2  # gap between regions

    # ── Tab 4: Unassigned Projects ─────────────────────────────────────────────
    ws4 = wb.create_sheet("Unassigned Projects")
    hdr_row(ws4, 1, ["Project", "Project Type", "Billing Type", "Territory", "PS Region",
                     "Signed Date", "Outreach Deadline", "Planned Start", "Scoped Hours", "Urgency"])
    for col, w in zip("ABCDEFGHIJ", [35, 22, 15, 15, 10, 14, 18, 14, 14, 12]):
        ws4.column_dimensions[get_column_letter(col)].width = w

    if ns_df is not None and not ns_df.empty:
        today_d = date.today()
        for row_i, (_, r) in enumerate(ns_df.iterrows(), 2):
            outreach = r.get("outreach_date")
            outreach_d = outreach.date() if pd.notna(outreach) and hasattr(outreach, "date") else None
            days_to_outreach = (outreach_d - today_d).days if outreach_d else None
            urgency = ("Overdue" if days_to_outreach is not None and days_to_outreach < 0
                      else "This Week" if days_to_outreach is not None and days_to_outreach <= 7
                      else "This Month" if days_to_outreach is not None and days_to_outreach <= 30
                      else "Upcoming")
            urg_fill = (red_fill if urgency == "Overdue" else
                       amber_fill if urgency in ("This Week", "This Month") else white_fill)

            vals = [
                r.get("project_name", ""),
                r.get("project_type", ""),
                r.get("billing_type", ""),
                r.get("territory", ""),
                r.get("ps_region", ""),
                r.get("signed_date").strftime("%Y-%m-%d") if pd.notna(r.get("signed_date")) else "",
                outreach_d.strftime("%Y-%m-%d") if outreach_d else "",
                r.get("start_date").strftime("%Y-%m-%d") if pd.notna(r.get("start_date")) else "",
                fmt_hrs(r.get("scoped_hours") or 0),
                urgency,
            ]
            for col_i, val in enumerate(vals, 1):
                fill = urg_fill if col_i == 10 else white_fill
                data_cell(ws4, row_i, col_i, val, fill=fill,
                         align=center if col_i in (5, 6, 7, 8, 9, 10) else left)

    # ── Tab 5: Parallel Conflicts ──────────────────────────────────────────────
    ws5 = wb.create_sheet("Parallel Conflicts")
    hdr_row(ws5, 1, ["Consultant", "Project 1", "Phase 1", "Project 2", "Phase 2",
                     "Combined Weight", "Overlap Start", "Overlap End", "Flag"])
    for col, w in zip("ABCDEFGHI", [28, 30, 25, 30, 25, 16, 14, 14, 12]):
        ws5.column_dimensions[get_column_letter(col)].width = w

    if conflicts:
        for row_i, cf in enumerate(conflicts, 2):
            vals = [
                cf["consultant"],
                cf["project_1"], cf["phase_1"],
                cf["project_2"], cf["phase_2"],
                cf["combined_weight"],
                cf["overlap_start"].strftime("%Y-%m-%d") if cf.get("overlap_start") else "",
                cf["overlap_end"].strftime("%Y-%m-%d") if cf.get("overlap_end") else "",
                "Capacity Conflict",
            ]
            for col_i, val in enumerate(vals, 1):
                fill = red_fill if col_i == 9 else amber_fill if col_i == 6 else white_fill
                data_cell(ws5, row_i, col_i, val, fill=fill,
                         align=center if col_i in (6, 7, 8, 9) else left)
    else:
        ws5.cell(row=2, column=1, value="No parallel conflicts detected.").font = std_font

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── UI ─────────────────────────────────────────────────────────────────────────
def main():
    st.markdown("""
        <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
            html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
            h1, h2, h3, .stMarkdown, .stDataFrame, label, button { font-family: 'Manrope', sans-serif !important; }
        </style>
        <div style='background-color:#1e2c63;padding:24px 32px;border-radius:8px;margin-bottom:24px;font-family:Manrope,sans-serif'>
            <h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif'>Capacity Outlook</h1>
            <p style='color:#aac4d0;margin:6px 0 0 0;font-size:14px;font-family:Manrope,sans-serif'>
                Project consultant availability over the next 6 months — active FF projects from Smartsheets + unassigned closed deals from NetSuite
            </p>
        </div>
    """, unsafe_allow_html=True)

    with st.expander("Exclusions & Limitations", expanded=False):
        st.markdown("""
        <style>
        .excl-list { font-family: Manrope, sans-serif; font-size: 13px; color: #4a5568; line-height: 1.8; padding-left: 4px; }
        .excl-list li { margin-bottom: 4px; }
        </style>
        <ul class='excl-list'>
            <li><b>Project Managers</b> are excluded from capacity calculations — they do not own product delivery.</li>
            <li><b>Projection accuracy</b> depends on SS project start dates being kept current by consultants and PMs.</li>
            <li><b>T&M active projects</b> are excluded from the projection engine — only FF active projects are used. T&M demand is captured via the NS Unassigned Projects report.</li>
            <li><b>Learning products</b> (Lappin: Capture/Reconcile; Longalong: Billing) are surfaced as future capacity only and not included in current assignment matching.</li>
            <li><b>Parallel conflicts</b> are flagged when two concurrent projects have a combined phase weight of 5.0 or higher (e.g. UAT + Go-Live simultaneously).</li>
        </ul>
        """, unsafe_allow_html=True)

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Step 1 — Upload Smartsheets DRS Export")
        st.caption("Required: Project Name, Consultant, Phase, Project Type, Billing Type, Start Date")
        ss_file = st.file_uploader("Drop SS DRS file here", type=["xlsx", "xls", "csv"], key="ss_cap_p3")
    with col2:
        st.subheader("Step 2 — Upload NS Unassigned Projects")
        st.caption("Required: Project, Territory, Billing Type, Project Type, Signed Date, Project Outreach, Start Date, T&M Scope")
        ns_file = st.file_uploader("Drop NS Unassigned Projects file here", type=["xlsx", "xls", "csv"], key="ns_unassigned_p3")

    ss_df = None
    ns_df = None

    if ss_file:
        try:
            ss_df = load_ss(ss_file)
            st.success(f"Loaded {len(ss_df):,} active FF project rows from SS DRS.")
        except Exception as e:
            st.error(f"Error loading SS file: {e}")

    if ns_file:
        try:
            ns_df = load_ns_unassigned(ns_file)
            st.success(f"Loaded {len(ns_df):,} unassigned projects from NS.")
        except Exception as e:
            st.error(f"Error loading NS file: {e}")

    if not ss_file and not ns_file:
        st.info("Upload SS DRS to project current capacity. Add NS Unassigned Projects to overlay demand.")
        return

    months = _months_range(HORIZON_MONTHS)
    month_labels = [_month_label(y, m) for y, m in months]

    availability, conflicts = project_consultant_availability(ss_df, months)

    # ── Metrics row ────────────────────────────────────────────────────────────
    today_str = datetime.today().strftime("%Y-%m")
    active_consultants = sum(
        1 for name, info in EMPLOYEE_ROLES.items()
        if info.get("role") not in ("Project Manager",)
        and _emp_active(name, today_str)
        and not info.get("util_exempt")
    )
    currently_busy = sum(
        1 for c, months_map in availability.items()
        if months_map.get(months[0], "free") == "busy"
    )
    available_now = sum(
        1 for c, months_map in availability.items()
        if months_map.get(months[0], "free") in ("free", "partial")
    )
    unassigned_count = len(ns_df) if ns_df is not None else 0
    conflict_count = len(conflicts)

    m1, m2, m3, m4, m5 = st.columns(5)
    def metric_card(label, value, sub=None, pill_color=None):
        sub_html = f"<div style='font-size:12px;color:#888;margin-top:2px'>{sub}</div>" if sub else ""
        pill_html = f"<div style='display:inline-block;margin-top:6px;padding:2px 10px;border-radius:999px;background:{pill_color}22;color:{pill_color};font-size:12px'>{sub}</div>" if pill_color and sub else ""
        return f"""<div style='background:#f8f9fb;border-radius:8px;padding:16px 20px;border:1px solid #e8eaed'>
            <div style='font-size:12px;color:#a0a0a0;font-family:Manrope,sans-serif;margin-bottom:4px'>{label}</div>
            <div style='font-size:28px;font-weight:700;color:#1e2c63;font-family:Manrope,sans-serif;line-height:1.1'>{value}</div>
            {pill_html if pill_color else sub_html}
        </div>"""

    with m1: st.markdown(metric_card("Active Delivery Staff", active_consultants), unsafe_allow_html=True)
    with m2: st.markdown(metric_card("Available Now", available_now, "free or light load", "#27AE60"), unsafe_allow_html=True)
    with m3: st.markdown(metric_card("Fully Busy", currently_busy, "heavy phase load", "#E74C3C"), unsafe_allow_html=True)
    with m4: st.markdown(metric_card("Unassigned Projects", unassigned_count, "closed, not started", "#F39C12"), unsafe_allow_html=True)
    with m5: st.markdown(metric_card("Parallel Conflicts", conflict_count, "flagged for review", "#E74C3C" if conflict_count else "#27AE60"), unsafe_allow_html=True)

    st.markdown("<div style='margin-top:16px'></div>", unsafe_allow_html=True)

    # ── Tabs ───────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "Capacity Heatmap",
        "Consultant Detail",
        "Regional Rollup",
        "Unassigned Projects",
    ])

    # ── Tab 1: Heatmap ─────────────────────────────────────────────────────────
    with tab1:
        st.markdown("#### Consultant Availability — Next 6 Months")
        st.caption("Green = Available · Amber = Light load (low-weight phase) · Red = Busy")

        if not availability:
            st.info("No active project data found. Upload SS DRS to generate heatmap.")
        else:
            heatmap_data = []
            for consultant in sorted(availability.keys()):
                region = _emp_ps_region(consultant)
                row_data = {"Consultant": consultant, "Region": region}
                for (yr, mo), label in zip(months, month_labels):
                    status = availability[consultant].get((yr, mo), "free")
                    row_data[label] = "Available" if status == "free" else "Light" if status == "partial" else "Busy"
                heatmap_data.append(row_data)

            hm_df = pd.DataFrame(heatmap_data)

            def color_cell(val):
                if val == "Available":
                    return "background-color:#C6EFCE;color:#276221"
                if val == "Light":
                    return "background-color:#FFEB9C;color:#9C6500"
                if val == "Busy":
                    return "background-color:#FFC7CE;color:#9C0006"
                return ""

            styled = hm_df.style.applymap(color_cell, subset=month_labels)
            st.dataframe(styled, hide_index=True, use_container_width=True)

        if conflicts:
            st.markdown("---")
            st.warning(f"{len(conflicts)} parallel conflict(s) detected — consultants in concurrent heavy-phase projects.")
            conflict_df = pd.DataFrame(conflicts)[["consultant", "project_1", "phase_1", "project_2", "phase_2", "combined_weight"]]
            conflict_df.columns = ["Consultant", "Project 1", "Phase 1", "Project 2", "Phase 2", "Combined Weight"]
            st.dataframe(conflict_df, hide_index=True, use_container_width=True)

    # ── Tab 2: Consultant Detail ───────────────────────────────────────────────
    with tab2:
        st.markdown("#### Consultant Capacity Detail")
        detail_rows = []
        for name, info in sorted(EMPLOYEE_ROLES.items(), key=lambda x: x[0]):
            if info.get("role") == "Project Manager":
                continue
            if not _emp_active(name, today_str):
                continue
            region = _emp_ps_region(name)
            role = info.get("role", "Consultant")
            products = ", ".join(info.get("products", []) or ["All"])
            learning = ", ".join(info.get("learning", []))

            projects, phases, free_from = "", "", "Available Now"
            if ss_df is not None and "consultant" in ss_df.columns:
                emp_rows = ss_df[ss_df["consultant"].str.strip() == name]
                if not emp_rows.empty:
                    projects = "; ".join(emp_rows["project_name"].dropna().unique()) if "project_name" in emp_rows.columns else ""
                    phases   = "; ".join(emp_rows["phase"].dropna().unique()) if "phase" in emp_rows.columns else ""
                    free_dates = []
                    for _, r in emp_rows.iterrows():
                        sd = r.get("start_date")
                        if pd.notna(sd):
                            sd = sd.date() if hasattr(sd, "date") else sd
                            pt = str(r.get("project_type", "")).strip()
                            free_dates.append(_project_end_date(sd, pt))
                    if free_dates:
                        free_from = max(free_dates).strftime("%b %Y")

            detail_rows.append({
                "Consultant": name,
                "Region": region,
                "Role": role,
                "Products": products,
                "Active Projects": projects,
                "Current Phase": phases,
                "Est. Free From": free_from,
                "Learning (Future)": learning,
            })

        if detail_rows:
            detail_df = pd.DataFrame(detail_rows)

            def highlight_free(row):
                color = "#C6EFCE" if row["Active Projects"] == "" else ""
                return [f"background-color:{color}" if color else "" for _ in row]

            styled_detail = detail_df.style.apply(highlight_free, axis=1)
            st.dataframe(styled_detail, hide_index=True, use_container_width=True)
        else:
            st.info("Upload SS DRS to populate consultant detail.")

    # ── Tab 3: Regional Rollup ─────────────────────────────────────────────────
    with tab3:
        st.markdown("#### Regional Capacity vs Demand")
        demand = build_demand_summary(ns_df, months) if ns_df is not None else {}

        for region in ["NOAM", "EMEA", "APAC"]:
            st.markdown(f"**{region}**")
            region_consultants = [c for c in availability if _emp_ps_region(c) == region]
            reg_demand = demand.get(region, {})

            rollup_rows = [
                {"Metric": "Consultants Available (Free/Light)"} | {
                    lbl: sum(1 for c in region_consultants
                             if availability[c].get(mo, "free") in ("free", "partial"))
                    for mo, lbl in zip(months, month_labels)
                },
                {"Metric": "Unassigned Projects (Demand)"} | {
                    lbl: reg_demand.get(mo, {}).get("count", 0)
                    for mo, lbl in zip(months, month_labels)
                },
                {"Metric": "Scoped Hours Demand"} | {
                    lbl: round(reg_demand.get(mo, {}).get("hours", 0), 1)
                    for mo, lbl in zip(months, month_labels)
                },
            ]
            rollup_df = pd.DataFrame(rollup_rows)

            def highlight_rollup(row):
                colors = [""] * len(row)
                if row["Metric"] == "Consultants Available (Free/Light)":
                    for i, col in enumerate(month_labels, 1):
                        v = row.get(col, 0)
                        colors[i] = "background-color:#C6EFCE" if v > 2 else "background-color:#FFEB9C" if v > 0 else "background-color:#FFC7CE"
                elif row["Metric"] == "Unassigned Projects (Demand)":
                    for i, col in enumerate(month_labels, 1):
                        v = row.get(col, 0)
                        colors[i] = "background-color:#FFC7CE" if v > 2 else "background-color:#FFEB9C" if v > 0 else "background-color:#C6EFCE"
                return colors

            styled_rollup = rollup_df.style.apply(highlight_rollup, axis=1)
            st.dataframe(styled_rollup, hide_index=True, use_container_width=True)
            st.markdown("<div style='margin-bottom:12px'></div>", unsafe_allow_html=True)

    # ── Tab 4: Unassigned Projects ─────────────────────────────────────────────
    with tab4:
        st.markdown("#### Unassigned Projects — Closed Deals Awaiting Staffing")
        if ns_df is None or ns_df.empty:
            st.info("Upload NS Unassigned Projects file to view demand.")
        else:
            today_d = date.today()
            display_df = ns_df.copy()
            if "outreach_date" in display_df.columns:
                def urgency(d):
                    if pd.isna(d): return "Unknown"
                    diff = (d.date() - today_d).days if hasattr(d, "date") else (d - today_d).days
                    return "Overdue" if diff < 0 else "This Week" if diff <= 7 else "This Month" if diff <= 30 else "Upcoming"
                display_df["Urgency"] = display_df["outreach_date"].apply(urgency)

            col_rename = {
                "project_name":   "Project",
                "project_type":   "Project Type",
                "billing_type":   "Billing Type",
                "territory":      "Territory",
                "ps_region":      "PS Region",
                "signed_date":    "Signed Date",
                "outreach_date":  "Outreach Deadline",
                "start_date":     "Planned Start",
                "scoped_hours":   "Scoped Hours",
            }
            show_cols = [c for c in col_rename if c in display_df.columns] + (["Urgency"] if "Urgency" in display_df.columns else [])
            display_df = display_df[show_cols].rename(columns=col_rename)

            def color_urgency(val):
                if val == "Overdue":    return "background-color:#FFC7CE;color:#9C0006"
                if val == "This Week":  return "background-color:#FFEB9C;color:#9C6500"
                if val == "This Month": return "background-color:#FFEB9C;color:#9C6500"
                return ""

            if "Urgency" in display_df.columns:
                styled_ns = display_df.style.applymap(color_urgency, subset=["Urgency"])
                st.dataframe(styled_ns, hide_index=True, use_container_width=True)
            else:
                st.dataframe(display_df, hide_index=True, use_container_width=True)

    # ── Export ─────────────────────────────────────────────────────────────────
    st.markdown("---")
    if st.button("Download Capacity Report (Excel)", type="primary"):
        excel_buf = build_excel(availability, conflicts, ns_df, months, ss_df)
        st.download_button(
            label="Save Excel",
            data=excel_buf,
            file_name=f"capacity_outlook_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    main()

main()
