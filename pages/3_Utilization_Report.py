"""
PS Utilization Credit Report — Self-contained page
"""
import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ════════════════════════════════════════════════════════
# CONFIG
# ════════════════════════════════════════════════════════

# ── Constants ─────────────────────────────────────────────────────────────────
NAVY     = "1e2c63"
TEAL     = "4472C4"
WHITE    = "FFFFFF"
LTGRAY   = "F2F2F2"
MID_GRAY = "BDC3C7"

TAG_COLORS = {
    "CREDITED":     "EAF9F1",
    "OVERRUN":      "FDECED",
    "PARTIAL":      "FEF9E7",
    "NON-BILLABLE": "EBEDEE",
    "FF: NO SCOPE DEFINED": "F2F2F2",
}
TAG_BADGE = {
    "CREDITED":     "🟢",
    "OVERRUN":      "🔴",
    "PARTIAL":      "🟡",
    "NON-BILLABLE": "⚫",
    "FF: NO SCOPE DEFINED": "⚪",
}

# ── Stored scope map ──────────────────────────────────────────────────────────
# ── PTO / Vacation keywords (matched against task/case column) ───────────────
PTO_KEYWORDS = ["vacation", "pto", "sick", "vacation/pto"]

# ── Employee → Location lookup (drives avail hours + PS region) ──────────────
EMPLOYEE_ROLES = {
    # Structure: name -> {role, products, learning, util_exempt (optional)}
    # roles: Consultant | Project Manager | Solution Architect | Developer
    # products: assignable product types
    # learning: in-training — surface as future capacity only, not current
    # util_exempt: True = excluded from utilization targets

    # ── Project Managers (no product delivery) ────────────────────────────────
    "Barrio, Nairobi":          {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    "Hughes, Madalyn":          {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    "Porangada, Suraj":       {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    "Cadelina":             {"role": "Project Manager",    "products": [],                                                                                      "learning": []},

    # ── Solution Architects ───────────────────────────────────────────────────
    "Bell, Stuart":            {"role": "Solution Architect", "products": ["Billing"],                                                                              "learning": []},
    "DiMarco, Nicole R":         {"role": "Solution Architect", "products": ["Billing"],                                                                              "learning": []},
    "Murphy, Conor":          {"role": "Solution Architect", "products": ["Billing"],                                                                              "learning": [], "util_exempt": True},

    # ── Developer ─────────────────────────────────────────────────────────────
    "Church, Jason G":      {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},

    # ── Consultants ───────────────────────────────────────────────────────────
    "Arestarkhov, Yaroslav":     {"role": "Consultant",         "products": ["Billing", "Capture"],                                                                   "learning": []},
    "Carpen, Anamaria":          {"role": "Consultant",         "products": ["Capture", "Approvals", "e-Invoicing"],                                                  "learning": []},
    "Centinaje, Rhodechild":       {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"], "learning": []},
    "Cooke, Ellen":           {"role": "Consultant",         "products": ["Billing", "Payroll"],                                                                   "learning": []},
    "Dolha, Madalina":           {"role": "Consultant",         "products": ["Capture", "Reconcile", "CC Statement Import", "Reconcile PSP", "e-Invoicing"],          "learning": []},
    "Finalle-Newton, Jesse":  {"role": "Solution Architect", "products": ["Reporting"],                                                                            "learning": []},
    "Gardner, Cheryll L":         {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Hopkins, Chris":         {"role": "Consultant",         "products": ["Capture", "Approvals"],                                                                 "learning": []},
    "Ickler, Georganne":          {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Isberg, Eric":          {"role": "Consultant",         "products": ["Reporting"],                                                                            "learning": []},
    "Jordanova, Marija":       {"role": "Consultant",         "products": ["Approvals", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"],     "learning": []},
    "Lappin, Thomas":          {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": ["Capture", "Reconcile"]},
    "Longalong, Santiago":       {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile"],                                                    "learning": ["Billing"]},
    "Mohammad, Manaan":        {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile"],                                                    "learning": []},
    "Morris, Lisa":          {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "NAQVI, SYED":          {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "Olson, Austin D":           {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Pallone, Daniel":         {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "Raykova, Silvia":         {"role": "Consultant",         "products": ["Capture", "Approvals", "e-Invoicing"],                                                  "learning": []},
    "Selvakumar, Sajithan":      {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile"],                                                    "learning": []},
    "Snee, Stefanie J":            {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Swanson, Patti":  {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": [], "util_exempt": True},
    "Tuazon, Carol":          {"role": "Consultant",         "products": ["Payroll", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"],       "learning": []},
    "Zoric, Ivan":           {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"], "learning": []},

    "Dunn, Steven":           {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
    "Law, Brandon":           {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
    "Quiambao, Generalyn":    {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
        "Cruz, Daniel":           {"role": "Consultant",         "products": ["All"],                                                                                  "learning": []},
    # ── Leavers (historical data only) ────────────────────────────────────────
    "Alam, Laisa":          {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Chan, Joven":          {"role": "Consultant",         "products": ["Capture"],                                                                              "learning": []},
    "Cloete, Bronwyn":      {"role": "Consultant",         "products": ["Capture", "Approvals"],                                                                 "learning": []},
    "Eyong, Eyong":         {"role": "Consultant",         "products": ["Capture"],                                                                              "learning": []},
    "Hamilton, Julie C":    {"role": "Consultant",         "products": ["Reporting"],                                                                            "learning": []},
    "Hernandez, Camila":    {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Rushbrook, Emma C":    {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "Strauss, John W":      {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
}

# ── Employees excluded from utilization targets ──────────────────────────────
UTIL_EXEMPT_EMPLOYEES = [
    k.lower() for k, v in EMPLOYEE_ROLES.items()
    if isinstance(v, dict) and v.get("util_exempt")
]  # derived from EMPLOYEE_ROLES util_exempt flag

# ── Employee roster — {name: (location, start_date, end_date)}
# start/end as "YYYY-MM" strings or None (None = no limit).
# Employees with an end_date will be excluded from utilization targets
# for any period after their exit month and flagged as Alumni.
EMPLOYEE_LOCATION = {
    #  Name                        Location              Start       End
    # ── Active employees ──────────────────────────────────────────────────────
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
    "Swanson, Patti":         ("UK",                  None,       None),
    "Tuazon, Carol":          ("Manila (PH)",         None,       None),
    "Zoric, Ivan":            ("Serbia",              None,       None),
    "Dunn, Steven":           ("USA",                 None,       None),
    "Law, Brandon":           ("USA",                 None,       None),
    "Quiambao, Generalyn":    ("Manila (PH)",         None,       None),
        # ── Leavers ───────────────────────────────────────────────────────────────
    "Alam, Laisa":            ("USA",                 None,       "2025-12"),
    "Chan, Joven":            ("Manila (PH)",         None,       "2025-12"),
    "Cloete, Bronwyn":        ("Netherlands",         None,       "2026-02"),
    "Eyong, Eyong":           ("USA",                 None,       "2025-12"),
    "Hamilton, Julie C":      ("USA",                 None,       "2026-01"),
    "Hernandez, Camila":      ("USA",                 None,       "2025-12"),
    "Rushbrook, Emma C":      ("Wales",               None,       "2025-12"),
    "Strauss, John W":        ("USA",                 None,       "2026-03"),
}

def _emp_location(name):
    """Return location string regardless of whether value is tuple or plain string."""
    v = EMPLOYEE_LOCATION.get(name)
    if v is None:
        return None
    return v[0] if isinstance(v, tuple) else v

def _emp_active(name, period_str):
    """
    Return True if the employee should be included for the given period.
    period_str: "YYYY-MM" or "YYYYMM" or a pandas Period string.
    Returns True if no tenure dates are set, or if period falls within [start, end].
    """
    v = EMPLOYEE_LOCATION.get(name)
    if v is None or not isinstance(v, tuple):
        return True   # unknown or plain-string legacy entry — don't exclude
    _, start, end = v
    try:
        # Normalise period to "YYYY-MM"
        p = str(period_str).strip()
        if len(p) == 6 and "-" not in p:
            p = p[:4] + "-" + p[4:]
        if start and p < start[:7]:
            return False
        if end and p > end[:7]:
            return False
    except Exception:
        pass
    return True

# ── PS Region overrides (employee name → region, bypasses location mapping) ──
PS_REGION_OVERRIDE = {
    "NAQVI, SYED":  "EMEA",  # Canada-based but reports into EMEA
    "Cruz, Daniel": "NOAM",  # Manila-based but reports into NOAM
    "Chan, Joven":    "NOAM",   # Manila-based but reports into NOAM
    "Rushbrook, Emma C": "EMEA",  # Wales → EMEA
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

DEFAULT_SCOPE = {
    "Capture":                 20,
    "Approvals":               17,
    "Reconcile":               17,
    "PSP":                     18,
    "Payments":                30,
    "Reconcile 2.0":           20,
    "CC":                       6,
    "SFTP":                    12,
    "Premium - 10":            10,
    "Premium - 20":            20,
    "E-Invoicing":             15,
    "Capture and E-Invoicing": 30,
    "Additional Subsidiary":    2,
}

# ── Available hours by region/month (2026) ────────────────────────────────────
AVAIL_HOURS = {
    "Spain": {"2025-01":165.0,"2025-02":150.0,"2025-03":157.5,"2025-04":150.0,"2025-05":157.5,"2025-06":157.5,"2025-07":172.5,"2025-08":150.0,"2025-09":165.0,"2025-10":165.0,"2025-11":142.5,"2025-12":157.5,"2026-01":155.0,"2026-02":155.0,"2026-03":170.5,"2026-04":162.75,"2026-05":155.0,"2026-06":170.5,"2026-07":178.25,"2026-08":162.75,"2026-09":170.5,"2026-10":162.75,"2026-11":155.0,"2026-12":155.0},
    "UK": {"2025-01":165.0,"2025-02":150.0,"2025-03":157.5,"2025-04":150.0,"2025-05":150.0,"2025-06":157.5,"2025-07":172.5,"2025-08":150.0,"2025-09":165.0,"2025-10":172.5,"2025-11":150.0,"2025-12":157.5,"2026-01":157.5,"2026-02":150.0,"2026-03":165.0,"2026-04":150.0,"2026-05":142.5,"2026-06":165.0,"2026-07":172.5,"2026-08":150.0,"2026-09":165.0,"2026-10":165.0,"2026-11":157.5,"2026-12":157.5},
    "Northern Ireland": {"2025-01":165.0,"2025-02":150.0,"2025-03":150.0,"2025-04":150.0,"2025-05":150.0,"2025-06":157.5,"2025-07":165.0,"2025-08":150.0,"2025-09":165.0,"2025-10":172.5,"2025-11":150.0,"2025-12":157.5,"2026-01":157.5,"2026-02":150.0,"2026-03":157.5,"2026-04":150.0,"2026-05":142.5,"2026-06":165.0,"2026-07":165.0,"2026-08":150.0,"2026-09":165.0,"2026-10":165.0,"2026-11":157.5,"2026-12":157.5},
    "Netherlands": {"2025-01":176.0,"2025-02":160.0,"2025-03":168.0,"2025-04":152.0,"2025-05":160.0,"2025-06":160.0,"2025-07":184.0,"2025-08":168.0,"2025-09":176.0,"2025-10":184.0,"2025-11":160.0,"2025-12":168.0,"2026-01":168.0,"2026-02":160.0,"2026-03":176.0,"2026-04":152.0,"2026-05":144.0,"2026-06":176.0,"2026-07":184.0,"2026-08":168.0,"2026-09":176.0,"2026-10":176.0,"2026-11":168.0,"2026-12":176.0},
    "Faroe Islands": {"2025-01":176.0,"2025-02":160.0,"2025-03":168.0,"2025-04":152.0,"2025-05":160.0,"2025-06":152.0,"2025-07":176.0,"2025-08":168.0,"2025-09":176.0,"2025-10":184.0,"2025-11":160.0,"2025-12":168.0,"2026-01":168.0,"2026-02":160.0,"2026-03":176.0,"2026-04":144.0,"2026-05":144.0,"2026-06":176.0,"2026-07":168.0,"2026-08":168.0,"2026-09":176.0,"2026-10":176.0,"2026-11":168.0,"2026-12":168.0},
    "North Macedonia": {"2025-01":168.0,"2025-02":160.0,"2025-03":168.0,"2025-04":160.0,"2025-05":152.0,"2025-06":168.0,"2025-07":184.0,"2025-08":160.0,"2025-09":168.0,"2025-10":176.0,"2025-11":160.0,"2025-12":176.0,"2026-01":160.0,"2026-02":160.0,"2026-03":168.0,"2026-04":168.0,"2026-05":160.0,"2026-06":176.0,"2026-07":184.0,"2026-08":168.0,"2026-09":168.0,"2026-10":168.0,"2026-11":168.0,"2026-12":176.0},
    "Czech Republic": {"2025-01":176.0,"2025-02":160.0,"2025-03":168.0,"2025-04":160.0,"2025-05":160.0,"2025-06":168.0,"2025-07":168.0,"2025-08":168.0,"2025-09":168.0,"2025-10":176.0,"2025-11":152.0,"2025-12":160.0,"2026-01":168.0,"2026-02":160.0,"2026-03":176.0,"2026-04":160.0,"2026-05":152.0,"2026-06":176.0,"2026-07":176.0,"2026-08":168.0,"2026-09":168.0,"2026-10":168.0,"2026-11":160.0,"2026-12":168.0},
    "Serbia": {"2025-01":168.0,"2025-02":144.0,"2025-03":168.0,"2025-04":152.0,"2025-05":160.0,"2025-06":168.0,"2025-07":184.0,"2025-08":168.0,"2025-09":176.0,"2025-10":184.0,"2025-11":152.0,"2025-12":184.0,"2026-01":152.0,"2026-02":152.0,"2026-03":176.0,"2026-04":160.0,"2026-05":160.0,"2026-06":176.0,"2026-07":184.0,"2026-08":168.0,"2026-09":176.0,"2026-10":176.0,"2026-11":160.0,"2026-12":184.0},
    "Canada": {"2025-01":176.0,"2025-02":152.0,"2025-03":168.0,"2025-04":160.0,"2025-05":168.0,"2025-06":168.0,"2025-07":176.0,"2025-08":160.0,"2025-09":168.0,"2025-10":176.0,"2025-11":152.0,"2025-12":168.0,"2026-01":168.0,"2026-02":160.0,"2026-03":176.0,"2026-04":168.0,"2026-05":160.0,"2026-06":176.0,"2026-07":176.0,"2026-08":160.0,"2026-09":160.0,"2026-10":168.0,"2026-11":160.0,"2026-12":168.0},
    "USA": {"2025-01":168.0,"2025-02":152.0,"2025-03":168.0,"2025-04":176.0,"2025-05":168.0,"2025-06":160.0,"2025-07":176.0,"2025-08":168.0,"2025-09":168.0,"2025-10":184.0,"2025-11":144.0,"2025-12":176.0,"2026-01":160.0,"2026-02":152.0,"2026-03":176.0,"2026-04":176.0,"2026-05":160.0,"2026-06":168.0,"2026-07":176.0,"2026-08":168.0,"2026-09":168.0,"2026-10":168.0,"2026-11":152.0,"2026-12":176.0},
    "Sydney (NSW)": {"2025-01":159.6,"2025-02":152.0,"2025-03":159.6,"2025-04":136.8,"2025-05":167.2,"2025-06":152.0,"2025-07":174.8,"2025-08":152.0,"2025-09":167.2,"2025-10":167.2,"2025-11":152.0,"2025-12":159.6,"2026-01":152.0,"2026-02":152.0,"2026-03":167.2,"2026-04":144.4,"2026-05":159.6,"2026-06":159.6,"2026-07":174.8,"2026-08":152.0,"2026-09":167.2,"2026-10":159.6,"2026-11":159.6,"2026-12":159.6},
    "Manila (PH)": {"2025-01":176.0,"2025-02":152.0,"2025-03":168.0,"2025-04":152.0,"2025-05":168.0,"2025-06":160.0,"2025-07":184.0,"2025-08":160.0,"2025-09":176.0,"2025-10":184.0,"2025-11":136.0,"2025-12":160.0,"2026-01":168.0,"2026-02":152.0,"2026-03":176.0,"2026-04":152.0,"2026-05":160.0,"2026-06":168.0,"2026-07":184.0,"2026-08":152.0,"2026-09":176.0,"2026-10":176.0,"2026-11":152.0,"2026-12":144.0},
}

# Fixed fee task keywords (Case/Task/Event column)
FF_TASKS = ["Configuration", "Enablement", "Training", "Post Go-live", "Project Management"]

def _emp_role(name):
    """Return role string for an employee, handling new dict structure."""
    v = EMPLOYEE_ROLES.get(name)
    if v is None:
        return None
    return v["role"] if isinstance(v, dict) else v

def _emp_products(name, include_learning=False):
    """Return list of assignable products for an employee."""
    v = EMPLOYEE_ROLES.get(name)
    if v is None or not isinstance(v, dict):
        return []
    products = list(v.get("products", []))
    if include_learning:
        products += v.get("products_learning", [])
    return products

def _emp_util_exempt(name):
    """Return True if employee is exempt from utilization targets.
    Falls back to last-name-only match for entries stored without first name (e.g. 'Swanson')."""
    v = EMPLOYEE_ROLES.get(name)
    if isinstance(v, dict):
        return v.get("util_exempt", False)
    # Try last-name-only fallback (e.g. NS exports 'Swanson, Jeff' but key is 'Swanson')
    last = name.split(",")[0].strip() if "," in name else name.strip()
    v2 = EMPLOYEE_ROLES.get(last)
    if isinstance(v2, dict):
        return v2.get("util_exempt", False)
    return False


def get_avail_hours(region, period):
    """Look up available hours for a region/period. Returns None if not found."""
    region_clean = str(region).strip()
    # Try exact match first, then case-insensitive
    for r, months in AVAIL_HOURS.items():
        if r.lower() == region_clean.lower():
            return months.get(str(period), None)
    return None

# ════════════════════════════════════════════════════════
# UTILS
# ════════════════════════════════════════════════════════

import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def match_ff_task(task_val):
    """Return raw task value for display — no categorization."""
    t = str(task_val).strip()
    return t if t and t.lower() not in ("", "nan", "none") else None

# ── Excel helpers ─────────────────────────────────────────────────────────────
def thin_border():
    s = Side(style="thin", color=MID_GRAY)
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_fill(hex_color):  return PatternFill("solid", fgColor=hex_color)
def row_fill(hex_color):  return PatternFill("solid", fgColor=hex_color)

GROUP_COLORS = ["EEF2FB", "FFFFFF"]  # alternating soft blue/white for grouped rows

def group_bg(value, prev_value, group_idx):
    """Return bg color and updated group index for grouped first-column display."""
    if value != prev_value:
        group_idx = 1 - group_idx  # toggle group
    return GROUP_COLORS[group_idx], group_idx

def style_header(ws, row, headers, fill_color=NAVY):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font      = Font(name="Manrope", bold=True, color=WHITE, size=10)
        c.fill      = hdr_fill(fill_color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = thin_border()
    ws.row_dimensions[row].height = 22

def style_cell(cell, bg, fmt=None, bold=False, align="left"):
    cell.fill      = row_fill(bg)
    cell.font      = Font(name="Manrope", size=10, bold=bold)
    cell.border    = thin_border()
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if fmt:
        cell.number_format = fmt

def write_title(ws, title, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font      = Font(name="Manrope", bold=True, size=14, color=WHITE)
    c.fill      = hdr_fill(NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

# ── Column detection ──────────────────────────────────────────────────────────
def auto_detect_columns(df):
    cols_lower = {c.lower().strip(): c for c in df.columns}
    checks = {
        "employee":      ["employee", "employee name", "name", "resource"],
        "project":       ["project", "project name", "job"],
        "project_type":  ["project type", "type", "project_type"],
        "date":          ["date", "time entry date", "entry date", "work date"],
        "hours":         ["hours", "duration", "hours logged", "time", "qty"],
        "approval":      ["approval status", "approval", "status"],
        "task":          ["case/task/event", "cask/task/event", "task", "case", "event", "memo"],
        "non_billable":  ["non-billable", "non billable", "nonbillable",
                          "non_billable", "is non billable"],
        "billing_type":  ["billing type", "billing_type", "bill type", "billtype"],
        "hours_to_date": ["hours to date", "hours_to_date", "htd", "prior hours",
                          "cumulative hours", "hours booked to date"],
        "region":           ["employee location", "location", "region", "country", "office"],
        "customer_region":  ["customer region", "customer_region", "cust region", "client region"],
        "project_manager":  ["project manager", "project_manager", "pm", "manager"],
        "project_phase":    ["project phase", "phase", "project_phase", "stage"],
        "start_date":       ["start date", "project start date", "start_date",
                             "project start", "commenced"],
    }
    mapping = {}
    unmatched = []
    for standard, candidates in checks.items():
        for candidate in candidates:
            if candidate in cols_lower:
                mapping[cols_lower[candidate]] = standard
                break
        else:
            if standard in ["employee", "project", "project_type", "hours", "non_billable"]:
                unmatched.append(standard)
    return mapping, unmatched

# ── Core credit logic ─────────────────────────────────────────────────────────
def assign_credits(df, scope_map):
    col_map, unmatched = auto_detect_columns(df)
    if unmatched:
        st.warning(f"Could not auto-detect columns: {unmatched}. Check your file headers.")

    df = df.rename(columns=col_map)
    df["non_billable"] = df["non_billable"].astype(str).str.strip().str.upper()
    # Normalize project names — collapse internal whitespace and strip edges
    if "project" in df.columns:
        df["project"] = df["project"].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
    df["hours"]  = pd.to_numeric(df["hours"], errors="coerce").fillna(0)
    df["date"]   = pd.to_datetime(df["date"], errors="coerce")
    df["period"] = df["date"].dt.strftime("%Y-%m").fillna("Unknown")
    if "start_date" in df.columns:
        df["start_date"] = pd.to_datetime(df["start_date"], errors="coerce")
    # Backfill location from hardcoded lookup when column is missing or blank
    if "region" not in df.columns:
        df["region"] = ""
    if "employee" in df.columns:
        def _resolve_location(row):
            loc = str(row.get("region","")).strip()
            if loc: return loc
            emp = str(row.get("employee","")).strip()
            # Normalize: strip extra spaces, try exact → prefix → last-name match
            emp_n = " ".join(emp.split())
            if emp_n in EMPLOYEE_LOCATION:
                return _emp_location(emp_n)
            emp_lower = emp_n.lower()
            for key, val in EMPLOYEE_LOCATION.items():
                key_lower = key.lower()
                if emp_lower == key_lower: return _emp_location(key)
                if emp_lower.startswith(key_lower) or key_lower.startswith(emp_lower): return _emp_location(key)
                # last name only match (before the comma)
                last = key_lower.split(",")[0].strip()
                if emp_lower.startswith(last): return val
            return ""
        df["region"] = df.apply(_resolve_location, axis=1)

    # Map employee location → PS Region, with per-employee overrides
    def _resolve_ps_region(row):
        emp = str(row.get("employee","")).strip()
        # Check override first
        for key in PS_REGION_OVERRIDE:
            if emp.lower().startswith(key.lower()) or key.lower() == emp.lower():
                return PS_REGION_OVERRIDE[key]
        return PS_REGION_MAP.get(str(row.get("region","")).strip(), "Other")
    df["ps_region"] = df.apply(_resolve_ps_region, axis=1)

    # ── Sort by date so entries are processed chronologically per project ────
    if "date" in df.columns:
        df = df.sort_values(["project", "date"], na_position="first").reset_index(drop=True)

    # ── Pre-compute prior hours per project ───────────────────────────────────
    # hours_to_date in NetSuite is cumulative INCLUDING this period's rows.
    # We want the baseline BEFORE this period starts, so:
    #   prior_htd[proj] = max(hours_to_date) - sum(hours this period)
    # Using max HTD as the ceiling snapshot, then subtracting all period hours
    # gives us hours on the project before any row in this export.
    _htd_col = "hours_to_date"
    prior_htd = {}
    if _htd_col in df.columns:
        # Group by (project, project_type) — same composite key as consumed dict
        # Prevents cross-product-type HTD bleed on multi-product accounts
        _htd_grp_cols = [c for c in ["project", "project_type"] if c in df.columns]
        for _key, _grp in df.groupby(_htd_grp_cols):
            _pk = tuple(" ".join(str(k).strip().split()) for k in (_key if isinstance(_key, tuple) else (_key,)))
            try:
                _max_htd    = float(_grp[_htd_col].dropna().astype(float).max() or 0)
                _period_hrs = float(_grp["hours"].dropna().astype(float).sum() or 0)
                prior_htd[_pk] = max(0.0, _max_htd - _period_hrs)
            except Exception:
                prior_htd[_pk] = 0.0

    consumed = {}
    credit_hrs_list    = []
    variance_hrs_list  = []
    credit_tag_list    = []
    notes_list         = []
    htd_start_list     = []  # track starting HTD per row for output

    for _, row in df.iterrows():
        proj      = " ".join(str(row.get("project", "")).split())  # normalize whitespace
        ptype     = str(row.get("project_type", "")).strip()
        hrs       = float(row.get("hours", 0))
        nb        = str(row.get("non_billable", "NO")).strip().upper()
        bill_type = str(row.get("billing_type", "")).strip().lower()
        is_tm     = bill_type == "t&m"

        if hrs <= 0:
            credit_hrs_list.append(0); variance_hrs_list.append(0)
            credit_tag_list.append("SKIPPED"); notes_list.append("Zero or missing hours")
            htd_start_list.append(0)
            continue

        # Rule 1: Internal = always 0 credit
        if bill_type == "internal":
            credit_hrs_list.append(0); variance_hrs_list.append(0)
            credit_tag_list.append("NON-BILLABLE"); notes_list.append("Internal: excluded from utilization")
            htd_start_list.append(0)
            continue

        # Rule 2: T&M = always full credit, no cap
        if is_tm:
            credit_hrs_list.append(hrs); variance_hrs_list.append(0)
            credit_tag_list.append("CREDITED"); notes_list.append("T&M: full credit")
            htd_start_list.append(0)
            continue

        # Rule 3: Fixed Fee = capped at scope (longest match wins)
        _ptype_lower = ptype.strip().lower()
        _matches = [(k, float(v)) for k, v in scope_map.items()
                    if k.strip().lower() in _ptype_lower]
        scope_hrs = max(_matches, key=lambda x: len(x[0]))[1] if _matches else None

        if scope_hrs is None:
            # UNCONFIGURED: count as credited (matches Daily Briefing logic)
            # Hours flagged separately via the no-scope banner below the metrics
            credit_hrs_list.append(hrs); variance_hrs_list.append(0)
            credit_tag_list.append("UNCONFIGURED"); notes_list.append(f"Fixed Fee but no scope defined for: {ptype}")
            htd_start_list.append(0)
            continue

        # Seed starting balance — composite key: (project, project_type)
        # Prevents cross-product-type HTD bleed on multi-product accounts
        _ck = (proj, ptype)
        if _ck not in consumed:
            consumed[_ck] = prior_htd.get(_ck, 0.0)

        already    = consumed[_ck]
        remaining  = scope_hrs - already
        htd_start_list.append(already)

        if remaining <= 0:
            credit_hrs_list.append(0); variance_hrs_list.append(hrs)
            credit_tag_list.append("OVERRUN"); notes_list.append(f"Scope exhausted (cap: {scope_hrs:.0f}h)")
        elif hrs <= remaining:
            consumed[_ck] = already + hrs
            credit_hrs_list.append(hrs); variance_hrs_list.append(0)
            credit_tag_list.append("CREDITED"); notes_list.append(f"NB within scope ({already:.1f}/{scope_hrs:.0f}h used)")
        else:
            consumed[_ck] = already + remaining
            credit_hrs_list.append(remaining); variance_hrs_list.append(hrs - remaining)
            credit_tag_list.append("PARTIAL")
            notes_list.append(f"Split: {remaining:.2f}h credited / {hrs - remaining:.2f}h overrun")

    df["credit_hrs"]    = credit_hrs_list
    df["variance_hrs"]  = variance_hrs_list
    df["credit_tag"]    = credit_tag_list
    df["notes"]         = notes_list
    df["htd_start"]     = htd_start_list

    # Updated hours to date = htd_start + total hours booked this period (credited + overrun)
    df["previous_htd"] = df["htd_start"] + df["hours"]

    # Tag FF tasks
    # ff_task — COL_MAP already renamed case/task/event -> task
    if "task" in df.columns:
        df["ff_task"] = df["task"].fillna("").apply(match_ff_task)
    else:
        df["ff_task"] = ""

    # Collect skipped rows for reporting
    skipped_df = df[df["credit_tag"] == "SKIPPED"][
        [c for c in ["employee","project","project_type","billing_type","date","hours","notes"] if c in df.columns]
    ].copy()

    return df, consumed, skipped_df


# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(df, scope_map, consumed):
    wb  = Workbook()
    wb.remove(wb.active)
    bgs = [WHITE, LTGRAY]

    # ── 1. PROCESSED DATA ─────────────────────────────────────

    ws = wb.create_sheet("PROCESSED_DATA")
    ws.sheet_properties.tabColor = TEAL
    ws.freeze_panes = "A3"

    headers = ["Employee","Location","Customer Region","Project Manager","Project",
               "Project Type","Billing Type","Hrs to Date","Date","Hours Logged",
               "Approval","Task/Case","Non-Billable","Credit Hrs","Variance Hrs",
               "Previous Hrs to Date","Credit Tag","Period","Notes",
               "Project Phase","Start Date","Days Active","Scoped Hrs","Variance Flag"]
    widths  = [20,16,18,20,35,20,14,13,14,13,14,25,13,12,12,18,16,12,45,16,14,12,12,16]
    cols    = ["employee","region","customer_region","project_manager","project",
               "project_type","billing_type","hours_to_date","date","hours",
               "approval","task","non_billable","credit_hrs","variance_hrs",
               "previous_htd","credit_tag","period","notes",
               "project_phase","start_date_display","days_active","_scoped_hrs","_variance_flag"]

    write_title(ws, "PROCESSED DATA — Utilization Credit Detail", len(headers))
    style_header(ws, 2, headers, TEAL)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def _get_scoped(ptype):
        """Lookup scoped hours for a project type from scope_map."""
        _m = [(k, float(v)) for k, v in scope_map.items()
              if k.strip().lower() in str(ptype).strip().lower()]
        return max(_m, key=lambda x: len(x[0]))[1] if _m else None

    def _variance_flag(tag, scoped):
        """Derive human-readable variance flag from credit_tag."""
        t = str(tag).strip().upper()
        if t in ("SKIPPED", "NON-BILLABLE", "CREDITED"):
            return "N/A"
        if t == "UNCONFIGURED":
            return "No Scope Set"
        if t in ("OVERRUN", "PARTIAL"):
            return "True Overrun" if scoped and scoped > 0 else "No Scope Set"
        return "Within Budget"

    for r_idx, (_, row) in enumerate(df.iterrows(), 3):
        tag      = str(row.get("credit_tag","")).strip()
        bg       = TAG_COLORS.get(tag, "F2F2F2")
        _scoped  = _get_scoped(row.get("project_type",""))
        _vflag   = _variance_flag(tag, _scoped)
        # Inject derived values so row.get() resolves them
        _row_ext = dict(row)
        _row_ext["_scoped_hrs"]    = round(_scoped, 2) if _scoped is not None else ""
        _row_ext["_variance_flag"] = _vflag
        for c_idx, col in enumerate(cols, 1):
            val  = _row_ext.get(col, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            fmt, bold, align = None, False, "left"
            if col == "date" and pd.notna(val):
                fmt = "YYYY-MM-DD"; align = "center"
            elif col in ("hours","credit_hrs","variance_hrs","hours_to_date","previous_htd"):
                fmt = "#,##0.00"; align = "right"
            elif col == "credit_tag":
                bold = True; align = "center"
            elif col in ("period","billing_type","region"):
                align = "center"
            elif col == "_scoped_hrs":
                fmt = "#,##0.00"; align = "right"
            elif col == "_variance_flag":
                align = "center"
                bold = True
                val_str = str(val)
                if val_str == "True Overrun":
                    bg = "FFC7CE"
                elif val_str == "No Scope Set":
                    bg = "FFEB9C"
                elif val_str == "Within Budget":
                    bg = "C6EFCE"
                else:
                    bg = TAG_COLORS.get(tag, "F2F2F2")
            style_cell(cell, bg, fmt=fmt, bold=bold, align=align)

    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}2"

    # ── 2. EMPLOYEE SUMMARY ───────────────────────────────────
    ws2 = wb.create_sheet("SUMMARY - By Employee")
    ws2.sheet_properties.tabColor = NAVY
    ws2.freeze_panes = "A3"

    eh = ["Employee","Location","PS Region","Period",
          "Avail Hrs (Capacity)","Hours This Period","Utilization Credits",
          "FF Project Overrun Hrs","Admin Hrs",
          "Util % (vs Hours Logged)","Util % (vs Capacity)",
          "Project Util %","Gap (pts)",
          "Projected Full Month Util %"]
    ew = [22,16,14,12,16,15,18,18,14,20,20,16,12,22]
    write_title(ws2, "SUMMARY — Utilization by Employee", len(eh))
    style_header(ws2, 2, eh, TEAL)
    ws2.auto_filter.ref = "A2:L2"

    for i, w in enumerate(ew, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # Get region per employee (first occurrence)
    emp_region = {}
    emp_cust_region = {}
    emp_pm = {}
    if "region" in df.columns:
        emp_region = df.dropna(subset=["region"]).groupby("employee")["region"].first().to_dict()
    if "customer_region" in df.columns:
        emp_cust_region = df.dropna(subset=["customer_region"]).groupby("employee")["customer_region"].first().to_dict()
    if "project_manager" in df.columns:
        emp_pm = df.dropna(subset=["project_manager"]).groupby("employee")["project_manager"].first().to_dict()

    # Admin hours = all Internal billing type rows
    admin_hrs = df[df["billing_type"].str.lower() == "internal"].groupby(
        ["employee","period"])["hours"].sum().reset_index().rename(columns={"hours":"admin_hrs"})

    emp_sum = df[df["credit_tag"] != "SKIPPED"].groupby(
        ["employee","period"], as_index=False
    ).agg(
        hours_this_period=("hours","sum"),
        credit_hrs=("credit_hrs","sum"),
        ff_overrun_hrs=("variance_hrs","sum"),
    ).sort_values(["employee","period"])

    emp_sum = emp_sum.merge(admin_hrs, on=["employee","period"], how="left")
    emp_sum["admin_hrs"] = emp_sum["admin_hrs"].fillna(0)

    # PTO hours per employee+period
    _pto_lookup = {}
    _pto_task_col = None
    for _tc in ["task", "case_task_event"]:
        if _tc in df.columns and not df[_tc].fillna("").str.strip().eq("").all():
            _pto_task_col = _tc; break
    if _pto_task_col:
        _pto_mask = df[_pto_task_col].fillna("").str.lower().apply(
            lambda t: any(k in t for k in PTO_KEYWORDS))
        _pto_df = df[_pto_mask].groupby(["employee","period"])["hours"].sum()
        _pto_lookup = {(emp, per): hrs for (emp, per), hrs in _pto_df.items()}

    # ── Partial month detection ──────────────────────────────────────────────
    _period_days = None
    _total_period_days = None
    if "date" in df.columns:
        _dates = df["date"].dropna()
        if len(_dates) > 0:
            import calendar
            _min_d = _dates.min()
            _max_d = _dates.max()
            _period_days = len(pd.bdate_range(_min_d, _max_d))
            _yr, _mo = _min_d.year, _min_d.month
            _month_start = pd.Timestamp(_yr, _mo, 1)
            _month_end   = pd.Timestamp(_yr, _mo, calendar.monthrange(_yr, _mo)[1])
            _total_period_days = len(pd.bdate_range(_month_start, _month_end))

    _prev_emp = None; _grp_idx = 0
    for r_idx, (_, row) in enumerate(emp_sum.iterrows(), 3):
        emp     = row["employee"]
        period  = row["period"]
        region  = emp_region.get(emp, "")
        avail   = get_avail_hours(region, period) if region else None
        util    = row["credit_hrs"] / row["hours_this_period"] if row["hours_this_period"] > 0 else 0
        bg, _grp_idx = group_bg(emp, _prev_emp, _grp_idx)
        _prev_emp = emp
        util_bg = ("EAF9F1" if util >= 0.8 else "FEF9E7" if util >= 0.6 else "FDECED")

        _ps_reg = df[df["employee"]==emp]["ps_region"].iloc[0] \
            if len(df[df["employee"]==emp]) > 0 else ""
        # Capacity util: credits / available hours in period
        cap_util = row["credit_hrs"] / avail if avail and avail > 0 else None
        # Project util: (credits + ff overrun) / available hours
        proj_util_pct = (row["credit_hrs"] + row["ff_overrun_hrs"]) / avail if avail and avail > 0 else None
        # Gap: project util - capacity util (percentage points)
        gap_pts = (proj_util_pct - cap_util) if proj_util_pct is not None and cap_util is not None else None
        # Partial month projection
        proj_util = None
        if _period_days is not None and _total_period_days is not None and _period_days < _total_period_days:
            daily_rate = row["credit_hrs"] / _period_days if _period_days > 0 else 0
            proj_util = (daily_rate * _total_period_days) / avail if avail and avail > 0 else None

        vals = [emp, region, _ps_reg, period, avail or "—",
                row["hours_this_period"], row["credit_hrs"],
                row["ff_overrun_hrs"], row.get("admin_hrs", 0),
                util if row["hours_this_period"] > 0 else "—",
                cap_util if cap_util is not None else "—",
                proj_util_pct if proj_util_pct is not None else "—",
                gap_pts if gap_pts is not None else "—",
                proj_util if proj_util is not None else "—"]
        fmts = [None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","0.0%","0.0%","0.0%","+0.0%;-0.0%;-","0.0%"]

        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            is_util_col = c_idx in (10, 11, 12)
            # Highlight Gap (pts) col 13 — amber if gap > 0.10, red if > 0.20
            if c_idx == 13 and isinstance(val, float):
                gap_bg = ("FCE4D6" if val >= 0.20 else "FFF2CC" if val >= 0.10 else bg)
                style_cell(cell, gap_bg, fmt=fmt, bold=(val >= 0.10),
                           align="right")
            else:
                style_cell(cell, util_bg if is_util_col else bg, fmt=fmt,
                           align="right" if c_idx > 4 else "center" if c_idx == 4 else "left")

    # ── 3. PROJECT SUMMARY ────────────────────────────────────
    ws3 = wb.create_sheet("FF By Project Utilization")
    ws3.sheet_properties.tabColor = "E67E22"
    ws3.freeze_panes = "A3"

    ph = ["Project","Project Type","Project Manager",
          "Scoped Hrs","Previous Hrs to Date","Hours This Period","Credit Hrs",
          "FF Project Overrun Hrs","Hours to Date","Hours Balance","Burn %","Status"]
    pw = [35,22,22,12,15,15,12,18,18,16,10,12]
    write_title(ws3, "SUMMARY — Utilization by Project (Fixed Fee)", len(ph))
    style_header(ws3, 2, ph, TEAL)
    ws3.auto_filter.ref = "A2:L2"

    for i, w in enumerate(pw, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # Project Summary: Fixed Fee only
    ff_proj_df = df[
        (df["credit_tag"] != "SKIPPED") &
        (df["billing_type"].str.lower() == "fixed fee")
    ] if "billing_type" in df.columns else df[df["credit_tag"] != "SKIPPED"]

    proj_sum = ff_proj_df.groupby(
        ["project","project_type"], as_index=False
    ).agg(
        hours_this_period=("hours","sum"),
        credit_hrs=("credit_hrs","sum"),
        variance_hrs=("variance_hrs","sum"),
        htd_start=("htd_start","first"),
    ).sort_values("project")

    # HTD seed now comes directly from aggregation above
    htd_seeds = dict(zip(proj_sum["project"], proj_sum["htd_start"]))

    # Project-level lookups for Customer Region and Project Manager
    proj_cust_region = {}
    proj_pm = {}
    if "customer_region" in df.columns:
        proj_cust_region = df.dropna(subset=["customer_region"]).groupby("project")["customer_region"].first().to_dict()
    if "project_manager" in df.columns:
        proj_pm = df.dropna(subset=["project_manager"]).groupby("project")["project_manager"].first().to_dict()
    # PS region per project — from first employee on that project
    proj_ps_region = df.groupby("project")["ps_region"].first().to_dict() \
        if "ps_region" in df.columns else {}
    proj_phase = {}
    if "project_phase" in df.columns:
        proj_phase = df.dropna(subset=["project_phase"]).groupby("project")["project_phase"].first().to_dict()
    proj_start = {}
    if "start_date" in df.columns:
        proj_start = df.dropna(subset=["start_date"]).groupby("project")["start_date"].min().to_dict()
    # Max date in import = "as of" date for Days Active calc
    _as_of = pd.to_datetime(df["date"], errors="coerce").max() if "date" in df.columns else pd.Timestamp.now()
    # Add project_phase, start_date, days_active columns to df for PROCESSED_DATA
    df["project_phase"] = df["project"].map(proj_phase) if proj_phase else ""
    if proj_start:
        df["start_date_mapped"] = df["project"].map(proj_start)
        df["days_active"] = df["start_date_mapped"].apply(
            lambda s: int((_as_of - s).days) if pd.notna(s) and pd.notna(_as_of) else None)
        # Use mapped start_date for display
        df["start_date_display"] = df["start_date_mapped"]
    else:
        df["start_date_display"] = df.get("start_date", None)
        df["days_active"] = None


    _prev_ptype = None; _grp_idx_p = 0
    for r_idx, (_, row) in enumerate(proj_sum.iterrows(), 3):
        ptype   = str(row["project_type"]).strip()
        _pm     = [(k, float(v)) for k, v in scope_map.items() if k.strip().lower() in ptype.lower()]
        scope_h = max(_pm, key=lambda x: len(x[0]))[1] if _pm else 0

        seed      = float(row["htd_start"]) if row["htd_start"] else 0
        # htd_start already includes hours this period (per NetSuite export)
        # Previous HTD = htd_start minus hours booked this period
        previous_h = max(0.0, seed - row["hours_this_period"])

        # Burn % = htd_start / scoped hrs (htd_start includes this period)
        burn = seed / scope_h if scope_h > 0 else 0

        vari_h   = row["variance_hrs"]
        htd_total = previous_h + row["hours_this_period"]
        if scope_h > 0 and htd_total > scope_h:   status = "OVERRUN"
        elif scope_h > 0 and htd_total == scope_h: status = "AT LIMIT"
        elif scope_h == 0 and htd_total > 0:       status = "REVIEW"
        elif burn >= 0.9:                          status = "REVIEW"
        elif burn > 0:                             status = "ON TRACK"
        else:                                      status = "—"

        status_bg = {"OVERRUN":"FDECED","AT LIMIT":"FDECED","REVIEW":"FEF9E7","ON TRACK":"EAF9F1"}.get(status, LTGRAY)
        bg, _grp_idx_p = group_bg(ptype, _prev_ptype, _grp_idx_p)
        _prev_ptype = ptype

        cust_reg   = proj_cust_region.get(row["project"], "")
        pm_name    = proj_pm.get(row["project"], "")
        phase      = proj_phase.get(row["project"], "")
        start_dt   = proj_start.get(row["project"])
        days_active = int((_as_of - start_dt).days) if pd.notna(start_dt) and pd.notna(_as_of) else "—"
        vals = [row["project"], ptype, pm_name,
                scope_h or "—", previous_h,
                row["hours_this_period"], row["credit_hrs"], vari_h,
                previous_h + row["hours_this_period"],
                (previous_h + row["hours_this_period"]) - scope_h if scope_h > 0 else "—",
                burn if scope_h > 0 else "—", status]
        fmts = [None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","0.0%",None]

        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, status_bg if c_idx == 12 else bg,
                       fmt=fmt, bold=(c_idx == 12),
                       align="right" if c_idx in (5,6,7,8,9,10,11) else "center" if c_idx == 12 else "left")

    # ── 4. FF OVERRUN BY PROJECT TYPE ────────────────────────
    ws_ot = wb.create_sheet("FF Overrun by Project Type")
    ws_ot.sheet_properties.tabColor = "8E44AD"
    ws_ot.freeze_panes = "A3"

    oth = ["Project Type","# Projects","# Over Budget","% Over Budget",
           "Total Overrun Hrs","Avg Scoped Hrs","Avg Actual Hrs","Avg +/− vs Scope"]
    otw = [28,11,14,14,17,14,14,17]
    write_title(ws_ot, "FIXED FEE OVERRUN BY PROJECT TYPE", len(oth))
    style_header(ws_ot, 2, oth, TEAL)
    ws_ot.auto_filter.ref = "A2:H2"
    for i, w in enumerate(otw, 1):
        ws_ot.column_dimensions[get_column_letter(i)].width = w

    # Build from FF projects with configured scopes only
    _ot_ff = df[
        (df["credit_tag"] != "SKIPPED") &
        (df.get("billing_type", pd.Series(dtype=str)).str.lower() == "fixed fee")
    ].copy() if "billing_type" in df.columns else df[df["credit_tag"] != "SKIPPED"].copy()

    # Compute per-project totals for over-budget detection
    _ot_proj = _ot_ff.groupby(["project","project_type"], as_index=False).agg(
        hours_total=("hours","sum"),
        overrun_hrs=("variance_hrs","sum"),
    )
    # Attach scoped hours
    def _ot_scope(ptype):
        _m = [(k, float(v)) for k, v in scope_map.items()
              if k.strip().lower() in str(ptype).strip().lower()]
        return max(_m, key=lambda x: len(x[0]))[1] if _m else None
    _ot_proj["scoped_hrs"] = _ot_proj["project_type"].apply(_ot_scope)
    # Only include projects where scope is configured
    _ot_proj = _ot_proj[_ot_proj["scoped_hrs"].notna()].copy()
    _ot_proj["over_under"] = _ot_proj["hours_total"] - _ot_proj["scoped_hrs"]
    _ot_proj["is_over"] = (_ot_proj["over_under"] > 0).astype(int)

    _ot_type = _ot_proj.groupby("project_type", as_index=False).agg(
        total_projects=("project","count"),
        over_count=("is_over","sum"),
        total_overrun_hrs=("overrun_hrs","sum"),
        avg_scoped=("scoped_hrs","mean"),
        avg_actual=("hours_total","mean"),
        avg_over_under=("over_under","mean"),
    ).sort_values("total_overrun_hrs", ascending=False)

    _ot_type["over_pct"] = _ot_type["over_count"] / _ot_type["total_projects"]

    _ot_totals_proj   = int(_ot_type["total_projects"].sum())
    _ot_totals_over   = int(_ot_type["over_count"].sum())
    _ot_totals_ovhrs  = float(_ot_type["total_overrun_hrs"].sum())

    for r_idx_ot, (_, row_ot) in enumerate(_ot_type.iterrows(), 3):
        bg_ot = GROUP_COLORS[r_idx_ot % 2]
        avg_ou = float(row_ot["avg_over_under"])
        over_pct = float(row_ot["over_pct"])

        vals_ot = [
            row_ot["project_type"],
            int(row_ot["total_projects"]),
            int(row_ot["over_count"]),
            over_pct,
            float(row_ot["total_overrun_hrs"]),
            float(row_ot["avg_scoped"]),
            float(row_ot["avg_actual"]),
            avg_ou,
        ]
        fmts_ot = [None,"#,##0","#,##0","0.0%","#,##0.0","0.0","0.0","+0.0;-0.0;\"-\""]

        for c_idx_ot, (val_ot, fmt_ot) in enumerate(zip(vals_ot, fmts_ot), 1):
            cell_ot = ws_ot.cell(row=r_idx_ot, column=c_idx_ot, value=val_ot)
            # % Over Budget colouring: red ≥40%, amber 25–39%
            if c_idx_ot == 4:
                if over_pct >= 0.40:
                    style_cell(cell_ot, "FCE4D6", fmt=fmt_ot, bold=True, align="right")
                    cell_ot.font = Font(name="Manrope", size=10, bold=True, color="9C2A00")
                elif over_pct >= 0.25:
                    style_cell(cell_ot, "FFF2CC", fmt=fmt_ot, bold=True, align="right")
                    cell_ot.font = Font(name="Manrope", size=10, bold=True, color="7F4F00")
                else:
                    style_cell(cell_ot, bg_ot, fmt=fmt_ot, align="right")
            # Avg +/− colouring: red if positive (running over on average), grey if negative
            elif c_idx_ot == 8:
                style_cell(cell_ot, bg_ot, fmt=fmt_ot, bold=(avg_ou > 0),
                           align="right")
                if avg_ou > 0:
                    cell_ot.font = Font(name="Manrope", size=10, bold=True, color="9C2A00")
                else:
                    cell_ot.font = Font(name="Manrope", size=10, color="595959")
            # Total overrun hrs — bold red if ≥ 500
            elif c_idx_ot == 5 and float(row_ot["total_overrun_hrs"]) >= 500:
                style_cell(cell_ot, bg_ot, fmt=fmt_ot, bold=True, align="right")
                cell_ot.font = Font(name="Manrope", size=10, bold=True, color="9C2A00")
            else:
                style_cell(cell_ot, bg_ot, fmt=fmt_ot,
                           align="left" if c_idx_ot == 1 else "right")
        ws_ot.row_dimensions[r_idx_ot].height = 15

    # Totals row
    _ot_tr = 3 + len(_ot_type)
    for c_idx_ot, (val_ot, fmt_ot) in enumerate(zip(
        ["Total", _ot_totals_proj, _ot_totals_over, None, _ot_totals_ovhrs, None, None, None],
        [None, "#,##0", "#,##0", None, "#,##0.0", None, None, None]
    ), 1):
        cell_ot = ws_ot.cell(row=_ot_tr, column=c_idx_ot, value=val_ot)
        cell_ot.font  = Font(name="Manrope", size=10, bold=True, color="FFFFFF")
        cell_ot.fill  = PatternFill("solid", fgColor=NAVY)
        cell_ot.border = thin_border()
        cell_ot.alignment = Alignment(
            horizontal="left" if c_idx_ot == 1 else "right", vertical="center")
        if fmt_ot and val_ot is not None:
            cell_ot.number_format = fmt_ot
    ws_ot.row_dimensions[_ot_tr].height = 18

    # ── 5. ZCO NON-BILLABLE BREAKDOWN ─────────────────────────
    ws5 = wb.create_sheet("Non-Billable")
    ws5.sheet_properties.tabColor = "95A5A6"
    ws5.freeze_panes = "A3"

    znh = ["Task / Activity","Employee","Period","Hours"]
    znw = [35,22,12,12]
    write_title(ws5, "NON-BILLABLE — Hours by Employee by Activity", len(znh))
    style_header(ws5, 2, znh, TEAL)
    ws5.auto_filter.ref = "A2:E2"

    for i, w in enumerate(znw, 1):
        ws5.column_dimensions[get_column_letter(i)].width = w

    # Use billing_type == internal as primary filter (reliable regardless of Non-Billable flag)
    if "billing_type" in df.columns:
        zco_df = df[df["billing_type"].fillna("").str.lower() == "internal"].copy()
    else:
        zco_df = df[df["credit_tag"] == "NON-BILLABLE"].copy()

    # task col already mapped from case/task/event via COL_MAP
    if len(zco_df) > 0:
        if "task" not in zco_df.columns:
            zco_df["task"] = "Internal (no task)"
        else:
            zco_df["task"] = zco_df["task"].fillna("").replace("", "Internal (no task)")

    if "task" in zco_df.columns and len(zco_df) > 0:
        zco_sum = zco_df.groupby(
            ["task","employee","period"], as_index=False
        ).agg(hours=("hours","sum")).sort_values(["task","employee","period"])

        # Total hours per employee per period (all rows incl billable)
        emp_period_totals = df[df["credit_tag"] != "SKIPPED"].groupby(
            ["employee","period"])["hours"].sum().to_dict()

        znh[3] = "Hours"
        # Update headers to include % col
        ws5.cell(row=2, column=5, value="% of Total Hrs").font = Font(name="Manrope", bold=True, color=WHITE, size=10)
        ws5.cell(row=2, column=5).fill = hdr_fill(TEAL)
        ws5.cell(row=2, column=5).alignment = Alignment(horizontal="center", vertical="center")
        ws5.cell(row=2, column=5).border = thin_border()
        ws5.column_dimensions["E"].width = 16

        _prev_task_z = None; _grp_idx_z = 0; r_idx = 3
        for _, row in zco_sum.iterrows():
            task = row.get("task","")
            # Navy section header when task changes
            if task != _prev_task_z:
                _task_hrs = zco_sum[zco_sum["task"]==task]["hours"].sum()
                for ci, (hval, hfmt) in enumerate([
                    (task, None), ("", None), ("", None),
                    (_task_hrs, "#,##0.00"), ("", None)], 1):
                    hcell = ws5.cell(row=r_idx, column=ci, value=hval)
                    hcell.font  = Font(name="Manrope", size=10, bold=True, color="FFFFFF")
                    hcell.fill  = PatternFill("solid", fgColor=NAVY)
                    hcell.border = thin_border()
                    hcell.alignment = Alignment(
                        horizontal="right" if ci==4 else "left", vertical="center")
                    if hfmt: hcell.number_format = hfmt
                r_idx += 1
                _prev_task_z = task
                _grp_idx_z = 0
            bg, _grp_idx_z = group_bg(task, task, _grp_idx_z)
            total_hrs  = emp_period_totals.get((row["employee"], row["period"]), 0)
            pct        = row["hours"] / total_hrs if total_hrs > 0 else 0
            vals = ["", row["employee"], row["period"], row["hours"], pct]
            fmts = [None, None, None, "#,##0.00", "0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws5.cell(row=r_idx, column=c_idx, value=val)
                style_cell(cell, bg, fmt=fmt,
                           align="right" if c_idx in (4,5) else "center" if c_idx == 3 else "left")
            r_idx += 1
    else:
        ws5.cell(row=3, column=1, value="No Non-Billable (Internal) entries in this period.")

    # ── 5. TASK ANALYSIS ───────────────────────
    ws6 = wb.create_sheet("Task Analysis")
    ws6.sheet_properties.tabColor = "27AE60"
    ws6.freeze_panes = "A3"

    tah = ["Task Category","Project Type","Hours This Period","Avg Hrs / Project","% of Type Hrs"]
    taw = [28,25,16,18,16]
    write_title(ws6, "TASK ANALYSIS — Hours by Task › Project Type", len(tah))
    style_header(ws6, 2, tah, TEAL)
    ws6.auto_filter.ref = "A2:E2"

    for i, w in enumerate(taw, 1):
        ws6.column_dimensions[get_column_letter(i)].width = w

    # Only Fixed Fee rows with a matched FF task
    ff_df = df[(df["billing_type"].str.lower() == "fixed fee") & (df["ff_task"].notna())].copy() \
        if "billing_type" in df.columns else df[df["ff_task"].notna()].copy()

    if len(ff_df) > 0:
        task_sum = ff_df.groupby(
            ["ff_task","project_type"], as_index=False
        ).agg(
            hours=("hours","sum")
        ).sort_values(["ff_task","project_type"])

        # Distinct project count per type using ALL fixed fee rows (not just task-tagged rows)
        all_ff = df[df["billing_type"].str.lower() == "fixed fee"] if "billing_type" in df.columns else df
        proj_count_by_type = all_ff.groupby("project_type")["project"].nunique().to_dict()

        # Total hours per type for % calc
        type_totals = ff_df.groupby("project_type")["hours"].sum().to_dict()

        _prev_task_t = None; _grp_idx_t = 0; r_idx_t = 3
        for _, row in task_sum.iterrows():
            ff_task = row["ff_task"]
            type_total = type_totals.get(row["project_type"], 0)
            pct        = row["hours"] / type_total if type_total > 0 else 0
            proj_cnt   = proj_count_by_type.get(row["project_type"], 1)
            raw_avg    = row["hours"] / proj_cnt if proj_cnt > 0 else 0
            avg_hrs    = round(raw_avg * 4) / 4
            # Navy section header when task changes
            if ff_task != _prev_task_t:
                _task_total_hrs = task_sum[task_sum["ff_task"]==ff_task]["hours"].sum()
                for ci, (hval, hfmt) in enumerate([
                    (ff_task, None), ("— ALL TYPES —", None),
                    (_task_total_hrs, "#,##0.00"), ("", None), ("", None)], 1):
                    hcell = ws6.cell(row=r_idx_t, column=ci, value=hval)
                    hcell.font  = Font(name="Manrope", size=10, bold=True, color="FFFFFF")
                    hcell.fill  = PatternFill("solid", fgColor=NAVY)
                    hcell.border = thin_border()
                    hcell.alignment = Alignment(
                        horizontal="right" if ci==3 else "left", vertical="center")
                    if hfmt: hcell.number_format = hfmt
                r_idx_t += 1
                _prev_task_t = ff_task
                _grp_idx_t = 0
            task_colors = {
                "Configuration":          "EBF5FB",
                "Post Go-Live Consulting": "FEF9E7",
                "Project Management":     "F4ECF7",
                "Training & UAT":         "EAF9F1",
                "Customer Communication": "FDF2F8",
            }
            bg, _grp_idx_t = group_bg(ff_task, ff_task, _grp_idx_t)
            task_bg = task_colors.get(ff_task, bg)
            vals = ["", row["project_type"], row["hours"], avg_hrs, pct]
            fmts = [None, None, "#,##0.00", "#,##0.00", "0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws6.cell(row=r_idx_t, column=c_idx, value=val)
                style_cell(cell, task_bg if c_idx > 1 else bg, fmt=fmt,
                           align="right" if c_idx > 2 else "left")
            r_idx_t += 1
    else:
        ws6.cell(row=3, column=1, value="No Fixed Fee task data found. Check Billing Type and Task/Case columns.")

    # ── 6. PROJECT COUNT BY TYPE ─────────────────────────────
    ws_pc = wb.create_sheet("Project Count")
    ws_pc.sheet_properties.tabColor = "2980B9"
    ws_pc.freeze_panes = "A3"

    pch = ["Project Type","Billing Type","Project Count"]
    pcw = [35, 14, 14]
    write_title(ws_pc, "PROJECT COUNT — Distinct Projects by Type (excl. Internal)", len(pch))
    style_header(ws_pc, 2, pch, TEAL)
    ws_pc.auto_filter.ref = "A2:D2"

    for i, w in enumerate(pcw, 1):
        ws_pc.column_dimensions[get_column_letter(i)].width = w

    # Exclude internal, count distinct projects per type only
    pc_df = df[df["billing_type"].str.lower() != "internal"].copy()         if "billing_type" in df.columns else df.copy()

    pc_sum = pc_df.groupby(["project_type","billing_type"], as_index=False).agg(
        project_count=("project","nunique"),
    ).sort_values(["project_type","billing_type"])

    grand_total = pc_sum["project_count"].sum()

    for r_idx, (_, row) in enumerate(pc_sum.iterrows(), 3):
        bg = LTGRAY if r_idx % 2 == 0 else WHITE
        vals = [row["project_type"], row["billing_type"], row["project_count"]]
        fmts = [None, None, "#,##0"]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_pc.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, bg, fmt=fmt, align="center" if c_idx == 2 else "right" if c_idx == 3 else "left")

    # Grand total row
    total_row = r_idx + 1 if len(pc_sum) > 0 else 3
    for c_idx, (val, fmt, bold) in enumerate([
        ("Grand Total", None, True),
        ("", None, False),
        (grand_total, "#,##0", True),
    ], 1):
        cell = ws_pc.cell(row=total_row, column=c_idx, value=val)
        cell.font   = Font(name="Manrope", bold=True, size=10, color=WHITE)
        cell.fill   = hdr_fill(NAVY)
        cell.border = thin_border()
        cell.alignment = Alignment(horizontal="right" if c_idx == 2 else "left", vertical="center")
        if fmt:
            cell.number_format = fmt

    # ── 7. CUSTOMER REGION SUMMARY ───────────────────────────
    # ── FF PROJECT TYPE ANALYSIS ─────────────────────────────
    ws_pta = wb.create_sheet("FF Project Type Analysis")
    ws_pta.sheet_properties.tabColor = "8E44AD"
    ws_pta.freeze_panes = "A4"

    ptah = ["Project Type","Task Category","Hours This Period","Avg Hrs / Project","% of Type Hrs"]
    ptaw = [28,25,16,18,16]
    write_title(ws_pta, "FF PROJECT TYPE ANALYSIS — Hours by Project Type › Task", len(ptah))
    style_header(ws_pta, 2, ptah, TEAL)
    ws_pta.auto_filter.ref = "A2:E2"
    ws_pta.cell(row=3, column=1,
        value="Grouped by Project Type → Task Category").font = Font(
        name="Manrope", size=9, italic=True, color="808080")
    for i, w in enumerate(ptaw, 1):
        ws_pta.column_dimensions[get_column_letter(i)].width = w

    if len(ff_df) > 0:
        pta_sum = ff_df.groupby(
            ["project_type","ff_task"], as_index=False
        ).agg(hours=("hours","sum"))
        pta_sum = pta_sum[pta_sum["ff_task"].notna() & (pta_sum["ff_task"] != "")]
        pta_sum = pta_sum.sort_values(["project_type","ff_task"])

        # Distinct project count per type (all FF rows)
        _all_ff_pta = df[df["billing_type"].fillna("").str.lower() == "fixed fee"]             if "billing_type" in df.columns else df.copy()
        _proj_count_pta = _all_ff_pta.groupby("project_type")["project"].nunique().to_dict()
        _type_totals_pta = ff_df.groupby("project_type")["hours"].sum().to_dict()

        _prev_ptype_pta = None; _grp_idx_pta = 0; r_idx_pta = 4
        for _, row in pta_sum.iterrows():
            ptype_pta   = row["project_type"]
            ff_task_pta = row["ff_task"]
            type_total  = _type_totals_pta.get(ptype_pta, 0)
            pct         = row["hours"] / type_total if type_total > 0 else 0
            proj_cnt    = _proj_count_pta.get(ptype_pta, 1)
            raw_avg     = row["hours"] / proj_cnt if proj_cnt > 0 else 0
            avg_hrs     = round(raw_avg * 4) / 4

            # Navy section header when project type changes
            if ptype_pta != _prev_ptype_pta:
                _ptype_total = pta_sum[pta_sum["project_type"]==ptype_pta]["hours"].sum()
                for ci, (hval, hfmt) in enumerate([
                    (ptype_pta, None), ("— ALL TASKS —", None),
                    (_ptype_total, "#,##0.00"), ("", None), ("", None)], 1):
                    hcell = ws_pta.cell(row=r_idx_pta, column=ci, value=hval)
                    hcell.font   = Font(name="Manrope", size=10, bold=True, color="FFFFFF")
                    hcell.fill   = PatternFill("solid", fgColor=NAVY)
                    hcell.border = thin_border()
                    hcell.alignment = Alignment(
                        horizontal="right" if ci==3 else "left", vertical="center")
                    if hfmt: hcell.number_format = hfmt
                r_idx_pta += 1
                _prev_ptype_pta = ptype_pta
                _grp_idx_pta = 0

            task_colors = {
                "Configuration":          "EBF5FB",
                "Post Go-Live Consulting": "FEF9E7",
                "Project Management":     "F4ECF7",
                "Training & UAT":         "EAF9F1",
                "Customer Communication": "FDF2F8",
            }
            bg_pta, _grp_idx_pta = group_bg(ptype_pta, ptype_pta, _grp_idx_pta)
            task_bg_pta = task_colors.get(ff_task_pta, bg_pta)

            vals = ["", ff_task_pta, row["hours"], avg_hrs, pct]
            fmts = [None, None, "#,##0.00", "#,##0.00", "0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws_pta.cell(row=r_idx_pta, column=c_idx, value=val)
                style_cell(cell, task_bg_pta if c_idx > 1 else bg_pta, fmt=fmt,
                           align="right" if c_idx > 2 else "left")
            r_idx_pta += 1
    else:
        ws_pta.cell(row=4, column=1, value="No Fixed Fee task data available.")


    ws_cr = wb.create_sheet("By Customer Region (WIP)")
    ws_cr.sheet_properties.tabColor = "1e2c63"
    ws_cr.freeze_panes = "A3"

    crh = ["Customer Region","Hours This Period","Utilization Credits",
           "FF Project Overrun Hrs","Util %"]
    crw = [22,16,18,20,10]
    write_title(ws_cr, "SUMMARY — Utilization by Customer Region", len(crh))
    style_header(ws_cr, 2, crh, TEAL)
    ws_cr.auto_filter.ref = "A2:E2"

    for i, w in enumerate(crw, 1):
        ws_cr.column_dimensions[get_column_letter(i)].width = w

    if "customer_region" in df.columns:
        cr_base = df[df["credit_tag"] != "SKIPPED"].copy()
        cr_base["customer_region"] = cr_base["customer_region"].fillna("Unassigned")

        cr_sum = cr_base.groupby("customer_region", as_index=False).agg(
            hours_this_period=("hours","sum"),
            credit_hrs=("credit_hrs","sum"),
            ff_overrun_hrs=("variance_hrs","sum"),
        ).sort_values("customer_region")

        for r_idx, (_, row) in enumerate(cr_sum.iterrows(), 3):
            cr      = row["customer_region"]
            total_h = row["hours_this_period"]
            util    = row["credit_hrs"] / total_h if total_h > 0 else 0
            util_bg = ("EAF9F1" if util >= 0.8 else "FEF9E7" if util >= 0.6 else "FDECED")
            bg      = bgs[r_idx % 2]
            vals = [cr, total_h, row["credit_hrs"], row["ff_overrun_hrs"],
                    util if total_h > 0 else "—"]
            fmts = [None,"#,##0.00","#,##0.00","#,##0.00","0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws_cr.cell(row=r_idx, column=c_idx, value=val)
                style_cell(cell, util_bg if c_idx == 5 else bg, fmt=fmt,
                           align="right" if c_idx > 1 else "left")
    else:
        ws_cr.cell(row=3, column=1, value="No 'Customer Region' column found in import.")

    # ── 8. PS REGION SUMMARY ─────────────────────────────────
    ws_ps = wb.create_sheet("By PS Region")
    ws_ps.sheet_properties.tabColor = "4472C4"
    ws_ps.freeze_panes = "A4"

    psh = ["PS Region","Project Type","Billing Type",
           "Avail Hrs","Hours This Period","Utilization Credits",
           "FF Project Overrun Hrs","Admin Hrs","Util %"]
    psw = [14,28,14,12,16,18,20,14,10]
    write_title(ws_ps, "SUMMARY — Utilization by PS Region (APAC / EMEA / NOAM)", len(psh))
    style_header(ws_ps, 2, psh, TEAL)
    ws_ps.auto_filter.ref = "A3:I3"

    # Sub-header note
    ws_ps.cell(row=3, column=1,
        value="Grouped by PS Region → Project Type → Billing Type").font = Font(
        name="Manrope", size=9, italic=True, color="808080")
    for i, w in enumerate(psw, 1):
        ws_ps.column_dimensions[get_column_letter(i)].width = w

    # Avail hrs per region (unique employee+period)
    ps_avail = {}
    _seen_ep = set()
    for _emp, _grp in df.groupby("employee"):
        _loc  = emp_region.get(_emp, "")
        _ps   = PS_REGION_MAP.get(_loc, "Other")
        for _p in _grp["period"].unique():
            if (_emp, _p) not in _seen_ep:
                _seen_ep.add((_emp, _p))
                ps_avail[_ps] = ps_avail.get(_ps, 0) + (get_avail_hours(_loc, _p) or 0)

    # Admin hrs per region
    ps_admin = {}
    if "billing_type" in df.columns:
        for _, _ar in df[df["billing_type"].str.lower()=="internal"].iterrows():
            _ps = PS_REGION_MAP.get(_ar.get("region",""), "Other")
            ps_admin[_ps] = ps_admin.get(_ps, 0) + _ar.get("hours", 0)

    # Build 3-level aggregation
    _ps_base = df[df["credit_tag"] != "SKIPPED"].copy()
    if "billing_type" not in _ps_base.columns:
        _ps_base["billing_type"] = "Unknown"
    _ps_detail = _ps_base.groupby(
        ["ps_region","project_type","billing_type"], as_index=False
    ).agg(
        hours_this_period=("hours","sum"),
        credit_hrs=("credit_hrs","sum"),
        ff_overrun_hrs=("variance_hrs","sum"),
    )

    # Region subtotals
    _ps_reg_total = _ps_base.groupby("ps_region", as_index=False).agg(
        hours_this_period=("hours","sum"),
        credit_hrs=("credit_hrs","sum"),
        ff_overrun_hrs=("variance_hrs","sum"),
    )

    region_order = ["APAC","EMEA","NOAM","Other"]
    _ps_detail["_rord"] = _ps_detail["ps_region"].map(
        {r:i for i,r in enumerate(region_order)}).fillna(99)
    _ps_detail = _ps_detail.sort_values(
        ["_rord","ps_region","project_type","billing_type"]).drop(columns=["_rord"])

    r_idx = 4
    _last_region = None
    for _, row in _ps_detail.iterrows():
        ps_reg = row["ps_region"]

        # ── Region header row ──────────────────────────────────
        if ps_reg != _last_region:
            _last_region = ps_reg
            _rt = _ps_reg_total[_ps_reg_total["ps_region"]==ps_reg]
            _rh = _rt.iloc[0]["hours_this_period"] if len(_rt) else 0
            _rc = _rt.iloc[0]["credit_hrs"]         if len(_rt) else 0
            _ro = _rt.iloc[0]["ff_overrun_hrs"]     if len(_rt) else 0
            _ra = ps_admin.get(ps_reg, 0)
            _rv = ps_avail.get(ps_reg, 0)
            _ru = _rc / _rh if _rh > 0 else 0
            _ru_bg = "EAF9F1" if _ru>=0.7 else "FEF9E7" if _ru>=0.6 else "FDECED"

            reg_vals = [ps_reg, "— ALL TYPES —", "",
                        _rv or "—", _rh, _rc, _ro, _ra,
                        _ru if _rh > 0 else "—"]
            reg_fmts = [None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(reg_vals, reg_fmts), 1):
                cell = ws_ps.cell(row=r_idx, column=c_idx, value=val)
                cell.font  = Font(name="Manrope", size=10, bold=True,
                                  color="FFFFFF" if c_idx <= 2 else "000000")
                cell.fill  = PatternFill("solid", fgColor=NAVY if c_idx <= 2 else (
                                  _ru_bg if c_idx == 9 else "D6DCF0"))
                cell.border = thin_border()
                if fmt: cell.number_format = fmt
                cell.alignment = Alignment(horizontal="right" if c_idx > 3 else "left",
                                           vertical="center")
            r_idx += 1

        # ── Detail row ─────────────────────────────────────────
        hrs  = row["hours_this_period"]
        util = row["credit_hrs"] / hrs if hrs > 0 else 0
        util_bg = "EAF9F1" if util >= 0.7 else "FEF9E7" if util >= 0.6 else "FDECED"
        bg = bgs[r_idx % 2]

        vals = ["", row["project_type"], row["billing_type"],
                "", hrs, row["credit_hrs"], row["ff_overrun_hrs"], "",
                util if hrs > 0 else "—"]
        fmts = [None,None,None,None,"#,##0.00","#,##0.00","#,##0.00",None,"0.0%"]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_ps.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, util_bg if c_idx == 9 else bg, fmt=fmt,
                       align="right" if c_idx > 3 else "left")
        r_idx += 1


    # ── 9. PROJECT WATCH LIST ────────────────────────────────
    ws_wl = wb.create_sheet("Watch List")
    ws_wl.sheet_properties.tabColor = "E74C3C"
    ws_wl.freeze_panes = "A3"

    # Section A: Top 10 overrun projects
    wlh = ["Project","Project Type","PS Region","Project Manager",
           "Scoped Hrs","Previous Hrs to Date","Hours to Date","Hours Balance","Burn %","FF Overrun Hrs","Status"]
    wlw = [35,20,14,22,12,18,14,16,10,14,12]
    write_title(ws_wl, "PROJECT WATCH LIST — Overrun & At-Risk Projects", len(wlh))
    style_header(ws_wl, 2, wlh, "E74C3C")
    ws_wl.auto_filter.ref = f"A2:{get_column_letter(len(wlh))}2"
    for i, w in enumerate(wlw, 1):
        ws_wl.column_dimensions[get_column_letter(i)].width = w

    # Build project-level data from ff_proj_df
    wl_df = ff_proj_df.groupby(["project","project_type"], as_index=False).agg(
        hours_this_period=("hours","sum"),
        credit_hrs=("credit_hrs","sum"),
        variance_hrs=("variance_hrs","sum"),
        htd_start=("htd_start","first"),
    )
    wl_df["previous_htd"] = wl_df.apply(
        lambda r: max(0.0, (float(r["htd_start"]) if r["htd_start"] else 0.0) - r["hours_this_period"]), axis=1)
    wl_df["hours_to_date"] = wl_df.apply(
        lambda r: (float(r["htd_start"]) if r["htd_start"] else 0.0), axis=1)

    def get_scope(ptype):
        _pm = [(k, float(v)) for k, v in scope_map.items() if k.strip().lower() in str(ptype).strip().lower()]
        return max(_pm, key=lambda x: len(x[0]))[1] if _pm else 0

    wl_df["scope_h"]  = wl_df["project_type"].apply(get_scope)
    wl_df["burn_pct"] = wl_df.apply(
        lambda r: (float(r["htd_start"]) if r["htd_start"] else 0) / r["scope_h"] if r["scope_h"] > 0 else None, axis=1)
    def _wl_status(r):
        s_h = r["scope_h"] or 0
        htd = (float(r["htd_start"]) if r["htd_start"] else 0)
        if s_h > 0 and htd > s_h:   return "OVERRUN"
        if s_h > 0 and htd == s_h:  return "AT LIMIT"
        if s_h == 0 and htd > 0:    return "REVIEW"
        if (r["burn_pct"] or 0) >= 0.9: return "REVIEW"
        return "ON TRACK"
    wl_df["status"] = wl_df.apply(_wl_status, axis=1)

    # Filter to OVERRUN + AT RISK, sort by burn desc
    watchlist = wl_df[wl_df["status"].isin(["OVERRUN","AT LIMIT","REVIEW"])].sort_values(
        "burn_pct", ascending=False, na_position="last")

    r_idx = 3
    for _, row in watchlist.iterrows():
        status   = row["status"]
        status_bg = "FDECED" if status == "OVERRUN" else "FEF9E7"
        bg       = status_bg
        burn_val = row["burn_pct"] if row["burn_pct"] is not None else "—"
        cust_reg   = proj_cust_region.get(row["project"], "")
        pm_name    = proj_pm.get(row["project"], "")
        phase      = proj_phase.get(row["project"], "")
        start_dt   = proj_start.get(row["project"])
        days_active = int((_as_of - start_dt).days) if pd.notna(start_dt) and pd.notna(_as_of) else "—"
        _htd_wl   = row["previous_htd"] + row["hours_this_period"]
        _tot_ov   = _htd_wl - row["scope_h"] if row["scope_h"] and row["scope_h"] > 0 else "—"
        _ps_reg_wl = proj_ps_region.get(row["project"], "")
        vals = [row["project"], row["project_type"], _ps_reg_wl, pm_name,
                row["scope_h"] or "—", row["previous_htd"],
                _htd_wl, _tot_ov,
                burn_val, row["variance_hrs"], status]
        fmts = [None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","0.0%","#,##0.00",None]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_wl.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, status_bg if c_idx == 11 else bg, fmt=fmt,
                       bold=(c_idx == 9),
                       align="right" if c_idx in (5,6,7,8) else "center" if c_idx == 9 else "left")
        r_idx += 1

    # Section B: FF: NO SCOPE DEFINED projects (no scope defined)
    r_idx += 1
    unconf_title_cell = ws_wl.cell(row=r_idx, column=1,
        value="FF: NO SCOPE DEFINED (Hours at Risk)")
    unconf_title_cell.font  = Font(name="Manrope", bold=True, size=11, color="FFFFFF")
    unconf_title_cell.fill  = hdr_fill("E67E22")
    ws_wl.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=len(wlh))
    r_idx += 1

    unconf_df = df[df["credit_tag"] == "UNCONFIGURED"].groupby(
        ["project","project_type"], as_index=False
    ).agg(hours=("hours","sum")).sort_values("hours", ascending=False)

    for _, row in unconf_df.iterrows():
        bg = "FEF3E2"
        vals = [row["project"], row["project_type"], proj_cust_region.get(row["project"],""),
                proj_pm.get(row["project"],""), "—", "—", "—", row["hours"], "FF: NO SCOPE DEFINED"]
        fmts = [None,None,None,None,None,None,None,"#,##0.00",None]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_wl.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, bg, fmt=fmt,
                       align="right" if c_idx == 8 else "left")
        r_idx += 1

    # ── 10. SKIPPED ROWS ──────────────────────────────────────
    ws7 = wb.create_sheet("Skipped Rows")
    ws7.sheet_properties.tabColor = "E74C3C"
    ws7.freeze_panes = "A3"

    skh = ["Employee","Project","Project Type","Billing Type","Date","Hours","Reason"]
    skw = [20,35,20,14,14,10,45]
    write_title(ws7, "SKIPPED ROWS — Not Included in Utilization Calculations", len(skh))
    style_header(ws7, 2, skh, "C0392B")
    for i, w in enumerate(skw, 1):
        ws7.column_dimensions[get_column_letter(i)].width = w

    skip_cols = ["employee","project","project_type","billing_type","date","hours","notes"]
    skipped = df[df["credit_tag"] == "SKIPPED"]
    if len(skipped) > 0:
        for r_idx, (_, row) in enumerate(skipped.iterrows(), 3):
            for c_idx, col in enumerate(skip_cols, 1):
                val  = row.get(col, "")
                cell = ws7.cell(row=r_idx, column=c_idx, value=val)
                fmt  = "YYYY-MM-DD" if col == "date" else "#,##0.00" if col == "hours" else None
                style_cell(cell, "FDECED", fmt=fmt,
                           align="right" if col == "hours" else "center" if col == "date" else "left")
    else:
        ws7.cell(row=3, column=1, value="No skipped rows — all entries were processed.")

    # ── DASHBOARD (exec summary) ─────────────────────────────
    ws_dash = wb.create_sheet("Dashboard")
    ws_dash.sheet_properties.tabColor = "1e2c63"
    ws_dash.sheet_view.showGridLines = False

    def dash_label(ws, row, col, text, size=10, bold=False, color="808080"):
        c = ws.cell(row=row, column=col, value=text)
        c.font = Font(name="Manrope", size=size, bold=bold, color=color)
        return c

    def dash_value(ws, row, col, value, fmt=None, size=18, bold=True, color="1e2c63"):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(name="Manrope", size=size, bold=bold, color=color)
        if fmt: c.number_format = fmt
        return c

    def dash_section(ws, row, col, text, ncols=4):
        c = ws.cell(row=row, column=col, value=text)
        c.font  = Font(name="Manrope", size=11, bold=True, color="FFFFFF")
        c.fill  = hdr_fill(NAVY)
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+ncols-1)
        return c

    def rag_cell(ws, row, col, value, fmt=None, status="green"):
        colors = {"green":"EAF9F1","yellow":"FEF9E7","red":"FDECED"}
        txt    = {"green":"2ECC71","yellow":"F39C12","red":"E74C3C"}
        c = ws.cell(row=row, column=col, value=value)
        c.font  = Font(name="Manrope", size=14, bold=True, color=txt.get(status,"000000"))
        c.fill  = PatternFill("solid", fgColor=colors.get(status,"FFFFFF"))
        c.alignment = Alignment(horizontal="center", vertical="center")
        if fmt: c.number_format = fmt
        return c

    for col, w in [(1,3),(2,22),(3,18),(4,18),(5,18),(6,18),(7,18),(8,3)]:
        ws_dash.column_dimensions[get_column_letter(col)].width = w
    for row in range(1, 45):
        ws_dash.row_dimensions[row].height = 18

    # Title
    tc = ws_dash.cell(row=2, column=2, value="Professional Services — Utilization Credit Report")
    tc.font = Font(name="Manrope", size=16, bold=True, color="FFFFFF")
    tc.fill = hdr_fill(NAVY)
    ws_dash.merge_cells(start_row=2, start_column=2, end_row=2, end_column=7)
    ws_dash.row_dimensions[2].height = 30

    if "date" in df.columns:
        max_dt   = pd.to_datetime(df["date"], errors="coerce").max()
        date_str = max_dt.strftime("%d %B %Y") if pd.notna(max_dt) else "—"
    else:
        date_str = "—"
    sc = ws_dash.cell(row=3, column=2, value=f"Data through {date_str}")
    sc.font = Font(name="Manrope", size=10, color="808080")
    ws_dash.merge_cells(start_row=3, start_column=2, end_row=3, end_column=7)
    kc = ws_dash.cell(row=4, column=2, value="This report calculates Utilization Credits from NetSuite time detail exports. T&M projects: full credit for all hours logged. Fixed Fee projects: credit up to scoped hours; hours beyond scope tracked as overrun (excluded from credits). Internal time: excluded from utilization, tracked as Admin Hours. Util % = Utilization Credits / Hours This Period.")
    kc.font = Font(name="Manrope", size=9, italic=True, color="808080")
    kc.alignment = Alignment(wrap_text=True)
    ws_dash.merge_cells(start_row=4, start_column=2, end_row=4, end_column=7)
    ws_dash.row_dimensions[4].height = 30
    # Push all section starts down 1 row to accommodate key text


    # Key Metrics
    dash_section(ws_dash, 6, 2, "KEY METRICS", ncols=6)
    ws_dash.row_dimensions[5].height = 22
    hours_tp_d    = df[df["credit_tag"] != "SKIPPED"]["hours"].sum()
    credit_hrs_d  = df[df["credit_tag"].isin(["CREDITED","PARTIAL"])]["credit_hrs"].sum()
    overrun_hrs_d = df[df["credit_tag"].isin(["OVERRUN", "PARTIAL"])]["variance_hrs"].sum()
    admin_hrs_d   = df[df["billing_type"].str.lower()=="internal"]["hours"].sum() if "billing_type" in df.columns else 0
    total_rows_d  = len(df[df["credit_tag"] != "SKIPPED"])
    util_pct_d    = credit_hrs_d / hours_tp_d if hours_tp_d > 0 else 0
    util_status_d = "green" if util_pct_d >= 0.70 else "yellow" if util_pct_d >= 0.60 else "red"

    for i, (label, value, fmt, status) in enumerate([
        ("Hours This Period", hours_tp_d, "#,##0.00", None),
        ("Utilization Credits", credit_hrs_d, "#,##0.00", None),
        ("Util % (target 70%)", util_pct_d, "0.0%", util_status_d),
        ("FF Overrun Hrs", overrun_hrs_d, "#,##0.00", None),
        ("Admin Hrs", admin_hrs_d, "#,##0.00", None),
        ("Projects This Period", df[df["billing_type"].fillna("").str.lower() != "internal"].groupby(["project","project_type"]).ngroups, "#,##0", None),
    ]):
        col = 2 + i
        dash_label(ws_dash, 7, col, label)
        if status:
            rag_cell(ws_dash, 8, col, value, fmt=fmt, status=status)
        else:
            dash_value(ws_dash, 8, col, value, fmt=fmt, size=14)
    ws_dash.row_dimensions[8].height = 28

    # PS Region
    dash_section(ws_dash, 10, 2, "UTILIZATION BY PS REGION", ncols=6)
    ws_dash.row_dimensions[9].height = 22
    for ci, hdr in enumerate(["PS Region","Hours This Period","Credit Hrs","Util %","FF Overrun Hrs","Admin Hrs"], 2):
        c = ws_dash.cell(row=11, column=ci, value=hdr)
        c.font = Font(name="Manrope", size=9, bold=True, color="FFFFFF")
        c.fill = hdr_fill(TEAL)

    ps_base_d = df[df["credit_tag"] != "SKIPPED"]
    ps_sum_d  = ps_base_d.groupby("ps_region").agg(
        hours=("hours","sum"), credit=("credit_hrs","sum"), overrun=("variance_hrs","sum"))
    ps_admin_d = df[df["billing_type"].str.lower()=="internal"].groupby("ps_region")["hours"].sum() if "billing_type" in df.columns else pd.Series(dtype=float)
    ps_avail_d = {}
    _seen_emp_period = set()
    for _emp2, _grp2 in df.groupby("employee"):
        _loc2  = emp_region.get(_emp2,"")
        _ps2   = PS_REGION_MAP.get(_loc2,"Other")
        for _p2 in _grp2["period"].unique():
            if (_emp2, _p2) not in _seen_emp_period:
                _seen_emp_period.add((_emp2, _p2))
                ps_avail_d[_ps2] = ps_avail_d.get(_ps2,0) + (get_avail_hours(_loc2,_p2) or 0)

    for ri, reg in enumerate(["APAC","EMEA","NOAM","Other"], 12):
        if reg not in ps_sum_d.index: continue
        _row = ps_sum_d.loc[reg]
        _adm = float(ps_admin_d.get(reg,0)) if reg in ps_admin_d.index else 0
        _avl = ps_avail_d.get(reg,0)
        _util= _row["credit"] / _row["hours"] if _row["hours"] > 0 else None
        _us  = "green" if _util>=0.70 else "yellow" if _util>=0.60 else "red"
        _bg  = bgs[ri % 2]
        _util_color = ("E74C3C" if _util<0.60 else "2ECC71" if _util>=0.70 else "F39C12") if _util is not None else "808080"
        _dash_ps_vals = [
            (2, reg,                    None,        False, "000000"),
            (3, _row["hours"],          "#,##0.00",  False, "000000"),
            (4, _row["credit"],         "#,##0.00",  False, "000000"),
            (5, _util if _util is not None else "—", "0.0%" if _util is not None else None, True, _util_color),
            (6, _row["overrun"],        "#,##0.00",  False, "000000"),
            (7, _adm,                   "#,##0.00",  False, "000000"),
        ]
        for ci2, val2, fmt2, bold2, color2 in _dash_ps_vals:
            _c = ws_dash.cell(row=ri, column=ci2, value=val2)
            _c.font  = Font(name="Manrope", size=10, bold=bold2, color=color2)
            _c.fill  = PatternFill("solid", fgColor=_bg)
            _c.border = thin_border()
            if fmt2: _c.number_format = fmt2

    # Watch List summary
    dash_section(ws_dash, 17, 2, "WATCH LIST SUMMARY", ncols=6)
    ws_dash.row_dimensions[17].height = 22
    # Align with Watch List — count projects where HTD > scope
    n_overrun  = len(wl_df[wl_df["status"]=="OVERRUN"]) if "wl_df" in dir() and len(wl_df) > 0 else \
                 len(df[df["credit_tag"]=="OVERRUN"]["project"].unique())
    _wl_at_risk = wl_df[(wl_df["burn_pct"].notna()) & (wl_df["burn_pct"]>=0.9) & (wl_df["status"]!="OVERRUN")] if "wl_df" in dir() else pd.DataFrame()
    n_at_risk  = len(_wl_at_risk["project"].unique()) if len(_wl_at_risk) > 0 else 0
    n_unconf   = len(df[df["credit_tag"]=="UNCONFIGURED"]["project"].unique())
    unconf_hrs_d = df[df["credit_tag"]=="UNCONFIGURED"]["hours"].sum()

    for i, (label, value, fmt, status) in enumerate([
        ("Projects in Overrun", n_overrun, "#,##0", "red" if n_overrun>0 else "green"),
        ("Projects (≥90% burn)", n_at_risk, "#,##0", "yellow" if n_at_risk>0 else "green"),
        ("FF: No Scope Defined Projects", n_unconf, "#,##0", "yellow" if n_unconf>0 else "green"),
        ("FF: No Scope Defined Hours", unconf_hrs_d, "#,##0.00", "yellow" if unconf_hrs_d>0 else "green"),
    ]):
        col = 2 + i
        dash_label(ws_dash, 18, col, label)
        rag_cell(ws_dash, 19, col, value, fmt=fmt, status=status)
    ws_dash.row_dimensions[19].height = 28

    # Low utilization employees
    dash_section(ws_dash, 21, 2, "EMPLOYEES BELOW 60% UTILIZATION — Action Required", ncols=6)
    ws_dash.row_dimensions[21].height = 22
    for ci, hdr in enumerate(["Employee","Location","PS Region","Period","Util %","Credit Hrs"], 2):
        _c = ws_dash.cell(row=22, column=ci, value=hdr)
        _c.font = Font(name="Manrope", size=9, bold=True, color="FFFFFF")
        _c.fill = hdr_fill(TEAL)

    _low_rows = []
    for _, _erow in emp_sum.iterrows():
        _emp3  = _erow["employee"]
        # Skip util-exempt employees
        if any(_emp3.lower().startswith(ex.lower()) for ex in UTIL_EXEMPT_EMPLOYEES):
            continue
        _loc3  = emp_region.get(_emp3,"")
        _ps3   = PS_REGION_MAP.get(_loc3,"Other")
        _p3    = _erow["period"]
        _avl3  = get_avail_hours(_loc3, _p3) or 0
        _util3 = _erow["credit_hrs"] / _avl3 if _avl3 > 0 else 0
        if _util3 < 0.60 and _avl3 > 0:
            _low_rows.append((_emp3, _loc3, _ps3, _p3, _util3, _erow["credit_hrs"]))

    for ri, (_e,_l,_ps,_per,_u,_c) in enumerate(sorted(_low_rows, key=lambda x:x[4])[:15], 23):
        for ci2, (val2,fmt2) in enumerate([(_e,None),(_l,None),(_ps,None),(_per,None),(_u,"0.0%"),(_c,"#,##0.00")], 2):
            _cv = ws_dash.cell(row=ri, column=ci2, value=val2)
            _cv.font  = Font(name="Manrope", size=10, color="E74C3C" if ci2==6 else "000000")
            _cv.fill  = PatternFill("solid", fgColor="FDECED")
            _cv.border = thin_border()
            if fmt2: _cv.number_format = fmt2



    # ── Reorder sheets: Project Count first, Processed Data last ────────────
    sheet_order = [
        "Dashboard",
        "Project Count",
        "SUMMARY - By Employee",
        "FF By Project Utilization",
        "FF Overrun by Project Type",
        "By Customer Region (WIP)",
        "By PS Region",
        "Watch List",
        "Non-Billable",
        "Task Analysis",
        "FF Project Type Analysis",
        "Skipped Rows",
        "PROCESSED_DATA",
    ]
    # Rebuild workbook sheet order directly
    existing = [s for s in sheet_order if s in wb.sheetnames]
    # Any sheets not in our list go at the end
    remaining = [s for s in wb.sheetnames if s not in existing]
    wb._sheets = [wb[s] for s in existing + remaining]

    # Save
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ════════════════════════════════════════════════════════
# PAGE
# ════════════════════════════════════════════════════════

import streamlit as st
import pandas as pd
import io




# Ensure shared/ is on the path regardless of how Streamlit Cloud runs the file


# ── Streamlit UI ──────────────────────────────────────────────────────────────
def main():
    st.markdown("""
        <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
            html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
            h1, h2, h3, .stMarkdown, .stDataFrame, label, button { font-family: 'Manrope', sans-serif !important; }
        </style>
    """, unsafe_allow_html=True)

    # Dynamic title suffix from View As
    _b = st.session_state.get("home_browse", "") or ""
    if _b.startswith("── ") and _b.endswith(" ──"):
        _title_sfx = f" — {_b[3:-3].strip()} Team"
    elif _b and _b not in ("— My own view —", "— Select —", "👥 All team"):
        _bp = [p.strip() for p in _b.split(",")]
        _title_sfx = f" — {_bp[1] + ' ' + _bp[0] if len(_bp)==2 else _b}"
    else:
        _title_sfx = ""

    st.markdown(f"""
        <div style='background:#1B2B5E;padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;font-family:Manrope,sans-serif;position:relative;overflow:hidden'>
            <div style='position:absolute;right:-40px;top:-40px;width:220px;height:220px;border-radius:50%;background:radial-gradient(circle,rgba(91,141,239,0.15) 0%,transparent 70%);pointer-events:none'></div>
            <div style='font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#ff4b40;margin-bottom:10px;font-family:Manrope,sans-serif'>Professional Services · Reporting</div>
            <h1 style='color:#fff;margin:0;font-size:28px;font-weight:800;font-family:Manrope,sans-serif;line-height:1.15'>Utilization Report{_title_sfx}</h1>
            <p style='color:rgba(255,255,255,0.6);margin:8px 0 0;font-size:14px;font-family:Manrope,sans-serif;line-height:1.6;max-width:520px'>Upload your NetSuite time detail export to generate a consistent utilization credit report across T&amp;M and Fixed Fee engagements.</p>
        </div>
    """, unsafe_allow_html=True)

    # ── Credit tag reference ───────────────────────────────────
    with st.expander("Credit Tag Logic", expanded=False):
        st.markdown("""
        <style>
        .ref-table { width: 100%; border-collapse: collapse; font-family: Manrope, sans-serif; font-size: 13px; margin-bottom: 8px; }
        .ref-table th { background: #4472C4; color: white; padding: 8px 12px; text-align: left; font-weight: 600; }
        .ref-table td { padding: 8px 12px; border-bottom: 1px solid #e2e8f0; vertical-align: top; }
        .tag-credited  { background: #EAF9F1; color: #1E8449; font-weight: 700; }
        .tag-partial   { background: #FEF9E7; color: #E67E22; font-weight: 700; }
        .tag-overrun   { background: #FDECED; color: #C0392B; font-weight: 700; }
        .tag-nonbill   { background: #F2F2F2; color: #555555; font-weight: 700; }
        .tag-noscope   { background: #F2F2F2; color: #777777; font-weight: 700; }
        </style>
        <table class='ref-table'>
            <tr><th>Credit Tag</th><th>Billing Type</th><th>Logic</th></tr>
            <tr style='background:#EAF9F1'><td style='color:#1E8449;font-weight:700'>CREDITED</td><td style='color:#1a1a1a'>T&amp;M</td><td style='color:#1a1a1a'>Full hours credited. No cap.</td></tr>
            <tr style='background:#EAF9F1'><td style='color:#1E8449;font-weight:700'>CREDITED</td><td style='color:#1a1a1a'>Fixed Fee</td><td style='color:#1a1a1a'>Hours credited up to scoped amount.</td></tr>
            <tr style='background:#FEF9E7'><td style='color:#E67E22;font-weight:700'>PARTIAL</td><td style='color:#1a1a1a'>Fixed Fee</td><td style='color:#1a1a1a'>Project has used some but not all scope. Hours credited up to remaining scope.</td></tr>
            <tr style='background:#FDECED'><td style='color:#C0392B;font-weight:700'>OVERRUN</td><td style='color:#1a1a1a'>Fixed Fee</td><td style='color:#1a1a1a'>Hours logged beyond contracted scope. Not credited. Tracked separately on the Watch List.</td></tr>
            <tr style='background:#F2F2F2'><td style='color:#555555;font-weight:700'>NON-BILLABLE</td><td style='color:#1a1a1a'>Internal</td><td style='color:#1a1a1a'>Internal time (admin, training, etc.). Excluded from utilization entirely. Tracked as Admin Hours.</td></tr>
            <tr style='background:#F2F2F2'><td style='color:#777777;font-weight:700'>FF: NO SCOPE DEFINED</td><td style='color:#1a1a1a'>Fixed Fee</td><td style='color:#1a1a1a'>FF project with no matching scope entry in the config. Hours not credited. Flagged for follow-up.</td></tr>
        </table>
        """, unsafe_allow_html=True)

    with st.expander("Utilization Target & RAG Status", expanded=False):
        st.markdown("""
        <table class='ref-table'>
            <tr><th>Status</th><th>Threshold</th><th>Description</th></tr>
            <tr style='background:#EAF9F1'><td style='color:#1E8449;font-weight:700'>Green</td><td style='color:#1a1a1a'>≥ 70%</td><td style='color:#1a1a1a'>At or above target utilization.</td></tr>
            <tr style='background:#FEF9E7'><td style='color:#E67E22;font-weight:700'>Amber</td><td style='color:#1a1a1a'>60% – 69%</td><td style='color:#1a1a1a'>Below target. Monitor and assess project mix.</td></tr>
            <tr style='background:#FDECED'><td style='color:#C0392B;font-weight:700'>Red</td><td style='color:#1a1a1a'>&lt; 60%</td><td style='color:#1a1a1a'>Significantly below target. Action required.</td></tr>
        </table>
        """, unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)

    # ── Upload — session state from Home takes priority ──────
    _ns_from_session = st.session_state.get("df_ns")
    uploaded = None

    if _ns_from_session is not None:
        st.success("✓ NS Time Detail loaded from Home page. Expand below to upload a different file.")
        with st.expander("Override NS Time Detail for this page", expanded=False):
            st.markdown("[Open NS Time Details Report ↗](https://3838224.app.netsuite.com/app/common/search/searchresults.nl?searchid=66732&saverun=T&whence=)")
            st.caption("Supported columns: Employee, Region, Project, Project Type, Billing Type, "
                       "Hours to Date, Date, Hours, Approval Status, Case/Task/Event, Non-Billable")
            uploaded = st.file_uploader(
                "Drop your file here or click to browse",
                type=["csv", "xlsx", "xls"],
                help="Supports CSV and Excel files exported from NetSuite",
                key="util_ns_override"
            )
    else:
        st.subheader("Step 1 — Upload NetSuite Time Detail Export")
        st.markdown("[Open NS Time Details Report ↗](https://3838224.app.netsuite.com/app/common/search/searchresults.nl?searchid=66732&saverun=T&whence=)")
        st.caption("Supported columns: Employee, Region, Project, Project Type, Billing Type, "
                   "Hours to Date, Date, Hours, Approval Status, Case/Task/Event, Non-Billable")
        uploaded = st.file_uploader(
            "Drop your file here or click to browse",
            type=["csv", "xlsx", "xls"],
            help="Supports CSV and Excel files exported from NetSuite"
        )

    if not uploaded and _ns_from_session is None:
        # Show stored config as reference
        with st.expander(" View stored scope map & available hours"):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Fixed Fee Scope Hours**")
                scope_df = pd.DataFrame(list(DEFAULT_SCOPE.items()), columns=["Project Type","Scoped Hrs"])
                st.dataframe(scope_df, hide_index=True, use_container_width=True)
            with c2:
                st.markdown("**Available Hours by Region (2026)**")
                avail_df = pd.DataFrame([
                    {"Region": r, **months} for r, months in AVAIL_HOURS.items()
                ])
                st.dataframe(avail_df, hide_index=True, use_container_width=True)
                st.info(" Upload your NetSuite export here, or load it on the Home page first.")
        return

    # Load file — local upload overrides session state
    if uploaded:
        try:
            ext = os.path.splitext(uploaded.name)[1].lower()
            df_raw = pd.read_excel(uploaded) if ext in (".xlsx",".xls") else \
                     pd.read_csv(uploaded, encoding="utf-8") if True else \
                     pd.read_csv(uploaded, encoding="latin-1")
        except Exception:
            try:
                df_raw = pd.read_csv(uploaded, encoding="latin-1")
            except Exception as e:
                st.error(f"Could not read file: {e}"); return
    else:
        df_raw = _ns_from_session

    # ── Apply view_as filter (respects home_browse for managers) ──
    _session_name = st.session_state.get("consultant_name", "")
    if _session_name:
        from shared.constants import get_role as _get_role, resolve_view_as, get_region_consultants
        from shared.config import EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE
        from shared.constants import ACTIVE_EMPLOYEES, EMPLOYEE_ROLES
        _home_browse = st.session_state.get("home_browse", "— My own view —")
        _va_name, _va_region, _is_group = resolve_view_as(
            _session_name, _home_browse, EMPLOYEE_ROLES,
            EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE, ACTIVE_EMPLOYEES
        )
        _role = _get_role(_session_name)
        _is_mgr = _role in ("manager", "manager_only")
        if _va_region and "employee" in df_raw.columns:
            _rc = get_region_consultants(_va_region, EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE, ACTIVE_EMPLOYEES)
            _filtered = df_raw[df_raw["employee"].astype(str).str.strip().str.lower().isin(_rc)]
            if not _filtered.empty: df_raw = _filtered.copy()
        elif _va_name and "employee" in df_raw.columns:
            _filtered = df_raw[df_raw["employee"].astype(str).str.strip().apply(lambda v: __import__("shared.constants", fromlist=["resolve_name"]).resolve_name(v)) == _va_name]
            if not _filtered.empty: df_raw = _filtered.copy()
        elif not _is_mgr and "employee" in df_raw.columns:
            _filtered = df_raw[df_raw["employee"].astype(str).str.strip().apply(lambda v: __import__("shared.constants", fromlist=["resolve_name"]).resolve_name(v)) == _session_name]
            if not _filtered.empty: df_raw = _filtered.copy()

    st.divider()

    # ── Process ───────────────────────────────────────────────
    st.subheader("Step 2 — Generate Report")

    if st.button("Run Utilization Engine", type="primary"):
        with st.spinner("Processing..."):
            try:
                df, consumed, skipped_df = assign_credits(df_raw.copy(), DEFAULT_SCOPE)
            except Exception as e:
                st.error(f"Processing error: {e}"); return

                st.success(" Processing complete!")


        # ── Warn on unmapped employees + alumni in period ────
        _unmapped = []
        _alumni   = []
        # Detect the reporting period (most common period in data)
        _period_str = str(df["period"].mode().iloc[0]) if "period" in df.columns and len(df) > 0 else None
        for _emp in df["employee"].dropna().unique():
            _emp_s = str(_emp).strip()
            _loc = df[df["employee"]==_emp]["region"].iloc[0] if len(df[df["employee"]==_emp]) > 0 else ""
            # Find matching key in roster
            _matched_key = None
            for k in EMPLOYEE_LOCATION:
                if _emp_s.lower() == k.lower() or _emp_s.lower().startswith(k.lower()) or k.lower().startswith(_emp_s.lower()):
                    _matched_key = k
                    break
            if not _matched_key and not str(_loc).strip():
                _unmapped.append(_emp_s)
            elif _matched_key and _period_str:
                if not _emp_active(_matched_key, _period_str):
                    _alumni.append(_emp_s)
        if _unmapped:
            st.warning(
                f"**{len(_unmapped)} employee(s) have no location defined** — "
                f"avail hours and PS region will show as Unknown. "
                f"Add them to EMPLOYEE_LOCATION in the app config.\n\n"
                + ", ".join(sorted(_unmapped))
            )
        if _alumni:
            st.info(
                f"**{len(_alumni)} employee(s) have time entries but are outside their active tenure** — "
                f"excluded from utilization targets. Check exit dates in EMPLOYEE_LOCATION if incorrect.\n\n"
                + ", ".join(sorted(_alumni))
            )

        # Metrics
        total_rows     = len(df[df["credit_tag"] != "SKIPPED"])
        total_credit   = df["credit_hrs"].sum()
        total_variance = df["variance_hrs"].sum()
        total_nb       = len(df[df["credit_tag"] == "NON-BILLABLE"])
        total_overrun  = len(df[df["credit_tag"] == "OVERRUN"])
        hours_this_period = df[df["credit_tag"] != "SKIPPED"]["hours"].sum() if "hours" in df.columns else 0
        total_admin    = df[df["billing_type"].str.lower() == "internal"]["hours"].sum()             if "billing_type" in df.columns else 0
        total_proj_overrun = df[df["credit_tag"].isin(["OVERRUN", "PARTIAL"])]["variance_hrs"].sum()             if "variance_hrs" in df.columns else 0

        credit_pct  = total_credit       / hours_this_period if hours_this_period else 0
        overrun_pct = total_proj_overrun / hours_this_period if hours_this_period else 0
        admin_pct   = total_admin        / hours_this_period if hours_this_period else 0

        credit_color = "#2ecc71" if credit_pct >= 0.70 else "#f39c12" if credit_pct >= 0.60 else "#C0392B"
        credit_label = "On target · Goal: 70%" if credit_pct >= 0.70 else "Below target · Goal: 70%" if credit_pct >= 0.60 else "At risk · Goal: 70%"

        # Date range banner
        if "date" in df.columns:
            _dates_clean = pd.to_datetime(df["date"], errors="coerce").dropna()
            if not _dates_clean.empty:
                _min_date = _dates_clean.min().strftime("%-d %b %Y")
                _max_date = _dates_clean.max().strftime("%-d %b %Y")
                date_str  = _dates_clean.max().strftime("%-d %B %Y")
            else:
                _min_date = _max_date = date_str = "—"
        else:
            _min_date = _max_date = date_str = "—"
        st.markdown(
            f"<div style='font-size:12px;color:#a0a0a0;font-family:Manrope,sans-serif;margin-bottom:12px'>"
            f"Reporting period: <strong style='color:inherit'>{_min_date} — {_max_date}</strong></div>",
            unsafe_allow_html=True
        )

        def fmt_hrs(n):
            """Show 2 decimals only if needed — drops .00 and .X0 trailing zeros."""
            return f"{n:,.2f}".rstrip("0").rstrip(".")  if "." in f"{n:,.2f}" else f"{n:,}"

        def metric_card(label, value, pill_txt=None, pill_fg=None, pill_icon="&#8593;"):
            pill = ""
            if pill_txt and pill_fg:
                pill = f"<div style='display:inline-block;margin-top:6px;padding:2px 10px;border-radius:999px;background-color:{pill_fg}33;font-size:13px;font-family:Manrope,sans-serif;color:{pill_fg}'>{pill_icon} {pill_txt}</div>"
            return f"""<div style='font-size:12px;color:#a0a0a0;font-family:Manrope,sans-serif;margin-bottom:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>{label}</div>
<div style='font-size:28px;font-weight:700;color:inherit;font-family:Manrope,sans-serif;line-height:1.15;white-space:nowrap'>{value}</div>{pill}"""

        # ── Metric row: 3 left + 2 right, with gap column in between ──
        st.markdown("""
        <style>
        [data-testid="stHorizontalBlock"] > div { padding-left: 6px !important; padding-right: 6px !important; }
        </style>""", unsafe_allow_html=True)

        m1, m2, m3, _gap, m4, m5 = st.columns([2, 2, 2, 0.3, 2, 2])
        with m1: st.markdown(metric_card("Projects This Period", f"{df[df['billing_type'].fillna('').str.lower() != 'internal'].groupby(['project','project_type']).ngroups:,}"), unsafe_allow_html=True)
        with m2: st.markdown(metric_card("Hours This Period", fmt_hrs(hours_this_period)), unsafe_allow_html=True)
        _cred_arrow = "↑" if credit_pct >= 0.60 else "↓"
        with m3: st.markdown(metric_card("Utilization Credits", fmt_hrs(total_credit), f"{credit_pct:.1%} of hrs · {credit_label}", credit_color, pill_icon=_cred_arrow), unsafe_allow_html=True)
        with m4: st.markdown(metric_card("FF Project Overrun Hrs", fmt_hrs(total_proj_overrun), f"{overrun_pct:.1%} of hrs", "#ff4b4b", pill_icon=""), unsafe_allow_html=True)
        with m5: st.markdown(metric_card("Admin Hrs", fmt_hrs(total_admin), f"{admin_pct:.1%} of hrs", "#808495", pill_icon="·"), unsafe_allow_html=True)

        # ── No-scope banner ───────────────────────────────────────
        _noscope_hrs = df[df["credit_tag"] == "UNCONFIGURED"]["hours"].sum() if "credit_tag" in df.columns else 0
        if _noscope_hrs > 0:
            st.markdown("<div style='margin-top:12px'></div>", unsafe_allow_html=True)
            st.warning(f"{fmt_hrs(_noscope_hrs)} hour(s) on FF projects with NO SCOPE DEFINED — see Watch List tab in the downloaded report.")

            st.markdown("---")
        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            ["By Employee", "By Project", "Non-Billable", "Task Analysis", "Detail"]
        )

        with tab1:
            _ep = df[df["credit_tag"] != "SKIPPED"]
            emp_sum_ui = _ep.groupby(["employee","period"], as_index=False).agg(
                hours_this_period=("hours","sum"),
                credit_hrs=("credit_hrs","sum"),
                ff_overrun_hrs=("variance_hrs","sum"),
                admin_hrs=("hours", lambda x: df.loc[
                    (df["employee"].isin(_ep["employee"])) &
                    (df["billing_type"].str.lower()=="internal"), "hours"
                ].sum() if "billing_type" in df.columns else 0),
            ).sort_values(["employee","period"])
            # Build region lookup directly from df for UI context
            _emp_region_ui = df.dropna(subset=["region"]).groupby("employee")["region"].first().to_dict() if "region" in df.columns else {}
            emp_sum_ui["location"]   = emp_sum_ui["employee"].map(_emp_region_ui)
            emp_sum_ui["avail_hrs"]  = emp_sum_ui.apply(
                lambda r: get_avail_hours(r["location"], r["period"]) if r["location"] else None, axis=1)
            # Util % vs hours logged — numeric for per-column RAG colouring
            emp_sum_ui["util_vs_logged_n"] = emp_sum_ui.apply(
                lambda r: r["credit_hrs"] / r["hours_this_period"]
                if r["hours_this_period"] > 0 else None, axis=1)
            # Util % vs available capacity — numeric for per-column RAG colouring
            emp_sum_ui["util_vs_capacity_n"] = emp_sum_ui.apply(
                lambda r: r["credit_hrs"] / r["avail_hrs"]
                if r["avail_hrs"] and r["avail_hrs"] > 0 else None, axis=1)
            # String display versions
            emp_sum_ui["util_vs_logged"] = emp_sum_ui["util_vs_logged_n"].apply(
                lambda v: f"{v*100:.1f}%" if v is not None else "—")
            emp_sum_ui["util_vs_capacity"] = emp_sum_ui["util_vs_capacity_n"].apply(
                lambda v: f"{v*100:.1f}%" if v is not None else "—")
            # Partial month projection — per period so multi-month uploads work correctly
            import calendar as _cal
            _period_bdays = {}
            if "date" in df.columns:
                for _per, _grp in df.groupby("period"):
                    _dates = _grp["date"].dropna()
                    if len(_dates) == 0: continue
                    _days_in = len(pd.bdate_range(_dates.min(), _dates.max()))
                    _yr2, _mo2 = _dates.min().year, _dates.min().month
                    _ms = pd.Timestamp(_yr2, _mo2, 1)
                    _me = pd.Timestamp(_yr2, _mo2, _cal.monthrange(_yr2, _mo2)[1])
                    _total = len(pd.bdate_range(_ms, _me))
                    _period_bdays[str(_per)] = (_days_in, _total)

            def _proj_util(row):
                per = str(row["period"])
                if per not in _period_bdays: return "—"
                days_in, total = _period_bdays[per]
                if days_in >= total: return "—"  # full month — no projection
                avail = row["avail_hrs"]
                if not avail or avail <= 0 or days_in <= 0: return "—"
                return f"{(row['credit_hrs'] / days_in * total) / avail * 100:.1f}%"

            emp_sum_ui["proj_full_month"] = emp_sum_ui.apply(_proj_util, axis=1)
            _is_partial = (emp_sum_ui["proj_full_month"] != "—").any()

            display_cols = ["employee", "location", "period", "avail_hrs",
                            "hours_this_period", "credit_hrs", "ff_overrun_hrs",
                            "util_vs_logged", "util_vs_capacity"]
            if _is_partial:
                display_cols.append("proj_full_month")
                st.caption("Partial period detected for one or more months. "
                           "Projected Full Month Util extrapolates credits at current daily run rate.")
            col_labels = {
                "employee":          "Employee",
                "location":          "Location",
                "period":            "Period",
                "avail_hrs":         "Avail Hrs (Capacity)",
                "hours_this_period": "Hours Logged",
                "credit_hrs":        "Util Credits",
                "ff_overrun_hrs":    "FF Overrun Hrs",
                "util_vs_logged":    "Util % (vs Logged)",
                "util_vs_capacity":  "Util % (vs Capacity)",
                "proj_full_month":   "Projected Full Month",
            }
            # Round raw numeric columns before display
            for _rc in ["avail_hrs", "hours_this_period", "credit_hrs", "ff_overrun_hrs"]:
                if _rc in emp_sum_ui.columns:
                    emp_sum_ui[_rc] = pd.to_numeric(emp_sum_ui[_rc], errors="coerce").round(2)
            show_df = emp_sum_ui[
                [col for col in display_cols if col in emp_sum_ui.columns]
            ].rename(columns=col_labels)

            # RAG colouring — separate thresholds per util column, based on numeric values
            _logged_col   = "Util % (vs Logged)"
            _capacity_col = "Util % (vs Capacity)"
            _proj_col     = "Projected Full Month"

            def _rag(val_str, numeric_series, col_name):
                """Return style string based on numeric lookup."""
                return ""  # placeholder — applied via apply below

            def _style_util_row(row):
                styles = [""] * len(row)
                for col_name, num_col in [
                    (_logged_col,   "util_vs_logged_n"),
                    (_capacity_col, "util_vs_capacity_n"),
                ]:
                    if col_name in row.index:
                        idx_pos = list(row.index).index(col_name)
                        # Look up numeric value from emp_sum_ui
                        emp = row.get("Employee", "")
                        per = row.get("Period", "")
                        match = emp_sum_ui[
                            (emp_sum_ui["employee"] == emp) &
                            (emp_sum_ui["period"] == per)
                        ]
                        if not match.empty:
                            n = match.iloc[0].get(num_col)
                            if n is not None and not _emp_util_exempt(emp):
                                bg    = ("#C6EFCE" if n >= 0.70
                                         else "#FFEB9C" if n >= 0.60
                                         else "#FFC7CE")
                                fg    = ("#276221" if n >= 0.70
                                         else "#9C6500" if n >= 0.60
                                         else "#9C0006")
                                styles[idx_pos] = f"background-color:{bg};color:{fg};font-weight:600"
                return styles

            styled_df = show_df.style.apply(_style_util_row, axis=1)
            st.dataframe(styled_df, use_container_width=True, hide_index=True)


        with tab2:
            proj_sum_ui = df[df["credit_tag"] != "SKIPPED"].groupby(
                ["project","project_type"], as_index=False
            ).agg(hours_this_period=("hours","sum"), credit_hrs=("credit_hrs","sum"),
                  ff_overrun_hrs=("variance_hrs","sum")).sort_values("project")
            st.dataframe(proj_sum_ui[["project","project_type","hours_this_period",
                         "credit_hrs","ff_overrun_hrs"]],
                         use_container_width=True, hide_index=True)

        with tab3:
            zco_df = df[df["credit_tag"] == "NON-BILLABLE"]
            if "task" in zco_df.columns and len(zco_df) > 0:
                zco_sum = zco_df.groupby(["task","employee","period"], as_index=False
                ).agg(hours=("hours","sum")).sort_values(["task","employee","period"])
                st.dataframe(zco_sum, use_container_width=True, hide_index=True)
            else:
                st.info("No Non-Billable (Internal) entries in this dataset.")

        with tab4:
            ff_df = df[df["ff_task"].notna()] if "ff_task" in df.columns else pd.DataFrame()
            if len(ff_df) > 0:
                task_sum = ff_df.groupby(["ff_task","project_type"], as_index=False
                ).agg(hours=("hours","sum")).sort_values(["ff_task","project_type"])
                type_totals = ff_df.groupby("project_type")["hours"].sum()
                task_sum["pct_of_type"] = task_sum.apply(
                    lambda r: f"{r['hours']/type_totals.get(r['project_type'],1)*100:.1f}%", axis=1)
                st.dataframe(task_sum, use_container_width=True, hide_index=True)
            else:
                st.info("No Fixed Fee task data found. Check Billing Type and Task/Case columns.")

        with tab5:
            display_cols = ["employee","region","project","project_type","billing_type",
                            "hours_to_date","date","hours","credit_hrs","variance_hrs",
                            "previous_htd","credit_tag","notes"]
            existing = [c for c in display_cols if c in df.columns]
            st.dataframe(df[existing].head(100), use_container_width=True, hide_index=True)
            if len(df) > 100:
                st.caption(f"Showing first 100 of {len(df):,} rows. Full data in Excel download.")

        st.divider()

        # Download
        st.subheader("Download Report")
        with st.spinner("Building Excel file..."):
            excel_buf = build_excel(df, DEFAULT_SCOPE, consumed)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename  = f"utilization_report_{timestamp}.xlsx"

        st.download_button(
label="⬇ Download Excel Report",
            data=excel_buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(f"`{filename}` — 6 tabs: Processed Data · By Employee · By Project · "
                   f"Non-Billable · Task Analysis · Skipped Rows")

        st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)
        with st.spinner("Building Tableau export..."):
            tableau_buf = build_tableau_excel(df, DEFAULT_SCOPE, consumed)
        tableau_filename = f"utilization_tableau_{timestamp}.xlsx"
        st.download_button(
            label="⬇ Download Tableau Export",
            data=tableau_buf,
            file_name=tableau_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption("`" + tableau_filename + "` — 3 flat sheets: fact_utilization · fact_processed_time_entries · fact_ff_overrun_by_type")


def build_tableau_excel(df, scope_map, consumed):
    """
    Clean flat export for Tableau — no formatting, no merged cells, no dashboard.
    Two sheets:
      fact_utilization  — one row per employee per period (summary metrics)
      fact_transactions — one row per time entry (raw grain)
    """
    import io as _io
    from openpyxl import Workbook as _WB

    wb = _WB()
    wb.remove(wb.active)

    # ── helpers ──────────────────────────────────────────────────────────────
    def _flat_sheet(wb, name, headers, rows):
        ws = wb.create_sheet(name)
        ws.append(headers)
        for r in rows:
            ws.append(r)
        return ws

    def _pct(num, denom):
        try:
            return round(num / denom, 6) if denom and denom > 0 else None
        except Exception:
            return None

    # ── pre-compute summary grain ─────────────────────────────────────────────
    _skipped = df["credit_tag"] == "SKIPPED" if "credit_tag" in df.columns else df.index == -1
    _df      = df[~_skipped].copy()

    # admin = internal billing rows
    _admin = (_df[_df["billing_type"].str.lower() == "internal"]
              .groupby(["employee","period"])["hours"].sum()
              .reset_index().rename(columns={"hours":"admin_hrs"}))

    _sum = (_df.groupby(["employee","period"], as_index=False)
            .agg(hours_logged=("hours","sum"),
                 credit_hrs=("credit_hrs","sum"),
                 ff_overrun_hrs=("variance_hrs","sum"))
            .sort_values(["employee","period"]))

    _sum = _sum.merge(_admin, on=["employee","period"], how="left")
    _sum["admin_hrs"] = _sum["admin_hrs"].fillna(0)

    # ── Sheet 1: fact_utilization ─────────────────────────────────────────────
    fu_headers = [
        "employee", "location", "ps_region", "role",
        "period", "hours_capacity", "hours_logged", "credit_hrs",
        "admin_hrs", "ff_overrun_hrs",
        "util_pct_vs_logged", "util_pct_vs_capacity",
        "project_util_pct", "gap_pts",
        "util_rag",
    ]
    fu_rows = []
    for _, row in _sum.iterrows():
        emp    = row["employee"]
        period = row["period"]
        loc    = df[df["employee"] == emp]["region"].iloc[0] if len(df[df["employee"] == emp]) > 0 else ""
        ps_reg = df[df["employee"] == emp]["ps_region"].iloc[0] if "ps_region" in df.columns and len(df[df["employee"] == emp]) > 0 else ""
        info   = EMPLOYEE_ROLES.get(emp, {})
        role        = info.get("role", "Consultant")

        avail  = get_avail_hours(loc, period) if loc else None
        u_log  = _pct(row["credit_hrs"], row["hours_logged"])
        u_cap  = _pct(row["credit_hrs"], avail)
        u_proj = _pct(row["credit_hrs"] + row["ff_overrun_hrs"], avail)
        gap    = round(u_proj - u_cap, 6) if u_proj is not None and u_cap is not None else None
        # RAG based on util vs capacity (primary metric)
        if u_cap is None:
            rag = "Unknown"
        elif u_cap >= 0.70:
            rag = "Green"
        elif u_cap >= 0.60:
            rag = "Amber"
        else:
            rag = "Red"

        fu_rows.append([
            emp, loc, ps_reg, role,
            period, avail, round(row["hours_logged"], 2), round(row["credit_hrs"], 2),
            round(row["admin_hrs"], 2), round(row["ff_overrun_hrs"], 2),
            u_log, u_cap, u_proj, gap, rag,
        ])
    _flat_sheet(wb, "fact_utilization", fu_headers, fu_rows)

    # ── Sheet 2: fact_processed_time_entries ─────────────────────────────────
    # Mirrors PROCESSED_DATA tab exactly — all same columns, Tableau-friendly names
    ft_headers = [
        "employee", "location", "ps_region", "customer_region", "project_manager", "project",
        "project_type", "billing_type", "hrs_to_date", "date", "hours_logged",
        "approval", "task_case", "non_billable", "credit_hrs", "variance_hrs",
        "previous_hrs_to_date", "credit_tag", "period", "notes",
        "project_phase", "start_date", "days_active",
        "scoped_hrs", "variance_flag",
    ]
    def _tbl_scoped(ptype):
        _m = [(k, float(v)) for k, v in scope_map.items()
              if k.strip().lower() in str(ptype).strip().lower()]
        return round(max(_m, key=lambda x: len(x[0]))[1], 2) if _m else None

    def _tbl_vflag(tag, scoped):
        t = str(tag).strip().upper()
        if t in ("SKIPPED", "NON-BILLABLE", "CREDITED"):
            return "N/A"
        if t == "UNCONFIGURED":
            return "No Scope Set"
        if t in ("OVERRUN", "PARTIAL"):
            return "True Overrun" if scoped and scoped > 0 else "No Scope Set"
        return "Within Budget"

    ft_rows = []
    for _, row in df.iterrows():
        def _g(col, default=""):
            v = row.get(col, default)
            if v is None or (hasattr(v, "__class__") and v.__class__.__name__ == "float" and str(v) == "nan"):
                return default
            return v

        _tag    = str(_g("credit_tag")).strip()
        _scoped = _tbl_scoped(_g("project_type"))
        ft_rows.append([
            _g("employee"), _g("region"), _g("ps_region", ""), _g("customer_region"), _g("project_manager"),
            _g("project"), _g("project_type"), _g("billing_type"),
            round(float(_g("hours_to_date", 0) or 0), 2),
            str(_g("date"))[:10], round(float(_g("hours", 0) or 0), 2),
            _g("approval"),
            _g("task", _g("case_task_event", "")),
            1 if str(_g("non_billable","")).lower() in ("true","yes","1","x") else 0,
            round(float(_g("credit_hrs", 0) or 0), 2),
            round(float(_g("variance_hrs", 0) or 0), 2),
            round(float(_g("previous_htd", 0) or 0), 2),
            _g("credit_tag"), _g("period"), _g("notes",""),
            _g("project_phase"), str(_g("start_date_display", _g("start_date", "")))[:10],
            _g("days_active",""),
            _scoped if _scoped is not None else "",
            _tbl_vflag(_tag, _scoped),
        ])
    _flat_sheet(wb, "fact_processed_time_entries", ft_headers, ft_rows)

    # ── Sheet 3: fact_ff_overrun_by_type ─────────────────────────────────────
    # One row per project type — aggregated FF overrun metrics.
    # Grain: project_type (join key to fact_utilization via project_type on
    # fact_processed_time_entries, or use as a standalone dimension in Tableau).
    fot_headers = [
        "project_type",
        "scoped_hrs",          # configured scope for this type (from scope_map)
        "total_projects",      # distinct FF projects of this type
        "projects_over_budget",
        "pct_over_budget",     # projects_over_budget / total_projects (0–1)
        "total_overrun_hrs",   # sum of variance_hrs across all overrun/partial rows
        "avg_scoped_hrs",      # mean scoped hrs per project
        "avg_actual_hrs",      # mean actual hours per project
        "avg_over_under_hrs",  # mean of (actual - scoped) — negative = typically under
    ]

    # Build from FF rows with configured scopes only
    _fot_ff = _df[
        _df.get("billing_type", pd.Series(dtype=str)).str.lower() == "fixed fee"
    ].copy() if "billing_type" in _df.columns else _df.copy()

    def _fot_scope(ptype):
        _m = [(k, float(v)) for k, v in scope_map.items()
              if k.strip().lower() in str(ptype).strip().lower()]
        return max(_m, key=lambda x: len(x[0]))[1] if _m else None

    # Per-project totals
    _fot_proj = _fot_ff.groupby(["project", "project_type"], as_index=False).agg(
        hours_total=("hours", "sum"),
        overrun_hrs=("variance_hrs", "sum"),
    )
    _fot_proj["scoped_hrs"] = _fot_proj["project_type"].apply(_fot_scope)
    _fot_proj = _fot_proj[_fot_proj["scoped_hrs"].notna()].copy()
    _fot_proj["over_under"] = _fot_proj["hours_total"] - _fot_proj["scoped_hrs"]
    _fot_proj["is_over"]    = (_fot_proj["over_under"] > 0).astype(int)

    # Aggregate to project-type grain
    _fot_type = _fot_proj.groupby("project_type", as_index=False).agg(
        scoped_hrs=("scoped_hrs", "mean"),       # same value for all rows of a type
        total_projects=("project", "count"),
        projects_over_budget=("is_over", "sum"),
        total_overrun_hrs=("overrun_hrs", "sum"),
        avg_scoped_hrs=("scoped_hrs", "mean"),
        avg_actual_hrs=("hours_total", "mean"),
        avg_over_under_hrs=("over_under", "mean"),
    ).sort_values("total_overrun_hrs", ascending=False)

    _fot_type["pct_over_budget"] = _fot_type.apply(
        lambda r: round(r["projects_over_budget"] / r["total_projects"], 6)
        if r["total_projects"] > 0 else None, axis=1
    )

    fot_rows = []
    for _, r in _fot_type.iterrows():
        fot_rows.append([
            r["project_type"],
            round(float(r["scoped_hrs"]), 2),
            int(r["total_projects"]),
            int(r["projects_over_budget"]),
            r["pct_over_budget"],
            round(float(r["total_overrun_hrs"]), 2),
            round(float(r["avg_scoped_hrs"]), 2),
            round(float(r["avg_actual_hrs"]), 2),
            round(float(r["avg_over_under_hrs"]), 2),
        ])
    _flat_sheet(wb, "fact_ff_overrun_by_type", fot_headers, fot_rows)

    buf = _io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


if __name__ == "__main__":
    main()