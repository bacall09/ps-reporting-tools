"""
PS Utilization Credit Report — Self-contained page
"""
import streamlit as st
import pandas as pd
import re as _re_constants
import io
import os
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.session_state["current_page"] = "Utilization Report"

# ════════════════════════════════════════════════════════
# CONFIG
# ════════════════════════════════════════════════════════

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

PTO_KEYWORDS = ["vacation", "pto", "sick", "vacation/pto"]

EMPLOYEE_ROLES = {
    "Barrio, Nairobi":          {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    "Hughes, Madalyn":          {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    "Porangada, Suraj":         {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    "Cadelina":                 {"role": "Project Manager",    "products": [],                                                                                      "learning": []},
    "Bell, Stuart":             {"role": "Solution Architect", "products": ["Billing"],                                                                              "learning": []},
    "DiMarco, Nicole R":        {"role": "Solution Architect", "products": ["Billing"],                                                                              "learning": []},
    "Murphy, Conor":            {"role": "Solution Architect", "products": ["Billing"],                                                                              "learning": [], "util_exempt": True},
    "Church, Jason G":          {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
    "Arestarkhov, Yaroslav":    {"role": "Consultant",         "products": ["Billing", "Capture"],                                                                   "learning": []},
    "Carpen, Anamaria":         {"role": "Consultant",         "products": ["Capture", "Approvals", "e-Invoicing"],                                                  "learning": []},
    "Centinaje, Rhodechild":    {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"], "learning": []},
    "Cooke, Ellen":             {"role": "Consultant",         "products": ["Billing", "Payroll"],                                                                   "learning": []},
    "Dolha, Madalina":          {"role": "Consultant",         "products": ["Capture", "Reconcile", "CC Statement Import", "Reconcile PSP", "e-Invoicing"],          "learning": []},
    "Finalle-Newton, Jesse":    {"role": "Solution Architect", "products": ["Reporting"],                                                                            "learning": []},
    "Gardner, Cheryll L":       {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Hopkins, Chris":           {"role": "Consultant",         "products": ["Capture", "Approvals"],                                                                 "learning": []},
    "Ickler, Georganne":        {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Isberg, Eric":             {"role": "Consultant",         "products": ["Reporting"],                                                                            "learning": []},
    "Jordanova, Marija":        {"role": "Consultant",         "products": ["Approvals", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"],     "learning": []},
    "Lappin, Thomas":           {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": ["Capture", "Reconcile"]},
    "Longalong, Santiago":      {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile"],                                                    "learning": ["Billing"]},
    "Mohammad, Manaan":         {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile"],                                                    "learning": []},
    "Morris, Lisa":             {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "NAQVI, SYED":              {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "Olson, Austin D":          {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Pallone, Daniel":          {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "Raykova, Silvia":          {"role": "Consultant",         "products": ["Capture", "Approvals", "e-Invoicing"],                                                  "learning": []},
    "Selvakumar, Sajithan":     {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile"],                                                    "learning": []},
    "Snee, Stefanie J":         {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Swanson, Patti":           {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": [], "util_exempt": True},
    "Tuazon, Carol":            {"role": "Consultant",         "products": ["Payroll", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"],       "learning": []},
    "Zoric, Ivan":              {"role": "Consultant",         "products": ["Capture", "Approvals", "Reconcile", "CC Statement Import", "Reconcile PSP", "SFTP Connector"], "learning": []},
    "Dunn, Steven":             {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
    "Law, Brandon":             {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
    "Quiambao, Generalyn":      {"role": "Developer",          "products": ["All"],                                                                                  "learning": []},
    "Cruz, Daniel":             {"role": "Consultant",         "products": ["All"],                                                                                  "learning": []},
    "Alam, Laisa":              {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Chan, Joven":              {"role": "Consultant",         "products": ["Capture"],                                                                              "learning": []},
    "Cloete, Bronwyn":          {"role": "Consultant",         "products": ["Capture", "Approvals"],                                                                 "learning": []},
    "Eyong, Eyong":             {"role": "Consultant",         "products": ["Capture"],                                                                              "learning": []},
    "Hamilton, Julie C":        {"role": "Consultant",         "products": ["Reporting"],                                                                            "learning": []},
    "Hernandez, Camila":        {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
    "Rushbrook, Emma C":        {"role": "Consultant",         "products": ["Payroll"],                                                                              "learning": []},
    "Strauss, John W":          {"role": "Consultant",         "products": ["Billing"],                                                                              "learning": []},
}

UTIL_EXEMPT_EMPLOYEES = [
    k.lower() for k, v in EMPLOYEE_ROLES.items()
    if isinstance(v, dict) and v.get("util_exempt")
]

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
    "Swanson, Patti":         ("UK",                  None,       None),
    "Tuazon, Carol":          ("Manila (PH)",         None,       None),
    "Zoric, Ivan":            ("Serbia",              None,       None),
    "Dunn, Steven":           ("USA",                 None,       None),
    "Law, Brandon":           ("USA",                 None,       None),
    "Quiambao, Generalyn":    ("Manila (PH)",         None,       None),
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
    v = EMPLOYEE_LOCATION.get(name)
    if v is None: return None
    return v[0] if isinstance(v, tuple) else v

def _emp_active(name, period_str):
    v = EMPLOYEE_LOCATION.get(name)
    if v is None or not isinstance(v, tuple): return True
    _, start, end = v
    try:
        p = str(period_str).strip()
        if len(p) == 6 and "-" not in p: p = p[:4] + "-" + p[4:]
        if start and p < start[:7]: return False
        if end and p > end[:7]: return False
    except Exception: pass
    return True

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

FF_TASKS = ["Configuration", "Enablement", "Training", "Post Go-live", "Project Management"]

def _emp_role(name):
    v = EMPLOYEE_ROLES.get(name)
    if v is None: return None
    return v["role"] if isinstance(v, dict) else v

def _emp_products(name, include_learning=False):
    v = EMPLOYEE_ROLES.get(name)
    if v is None or not isinstance(v, dict): return []
    products = list(v.get("products", []))
    if include_learning: products += v.get("products_learning", [])
    return products

def _emp_util_exempt(name):
    v = EMPLOYEE_ROLES.get(name)
    if isinstance(v, dict): return v.get("util_exempt", False)
    last = name.split(",")[0].strip() if "," in name else name.strip()
    v2 = EMPLOYEE_ROLES.get(last)
    if isinstance(v2, dict): return v2.get("util_exempt", False)
    return False

def get_avail_hours(region, period):
    region_clean = str(region).strip()
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
    t = str(task_val).strip()
    return t if t and t.lower() not in ("", "nan", "none") else None

def thin_border():
    s = Side(style="thin", color=MID_GRAY)
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_fill(hex_color):  return PatternFill("solid", fgColor=hex_color)
def row_fill(hex_color):  return PatternFill("solid", fgColor=hex_color)

GROUP_COLORS = ["EEF2FB", "FFFFFF"]

def group_bg(value, prev_value, group_idx):
    if value != prev_value: group_idx = 1 - group_idx
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
    if fmt: cell.number_format = fmt

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
        "project":       ["project", "project name", "job", "customer:project", "customer: project", "project: name"],
        "project_type":  ["project type", "type", "project_type"],
        "time_item_sku": ["time item sku", "item sku", "sku", "time item"],
        "date":          ["date", "time entry date", "entry date", "work date"],
        "hours":         ["hours", "duration", "hours logged", "time", "qty"],
        "approval":      ["approval status", "approval", "status"],
        "task":          ["case/task/event", "cask/task/event", "task", "case", "event", "memo"],
        "non_billable":  ["non-billable", "non billable", "nonbillable", "non_billable", "is non billable"],
        "billing_type":  ["billing type", "billing_type", "bill type", "billtype"],
        "hours_to_date": ["hours to date", "hours_to_date", "htd", "prior hours", "cumulative hours", "hours booked to date"],
        "project_id":    ["project id", "project_id", "projectid", "project internal id"],
        "region":        ["employee location", "location", "region", "country", "office"],
        "customer_region":  ["customer region", "customer_region", "cust region", "client region"],
        "project_manager":  ["project manager", "project_manager", "pm", "manager"],
        "project_phase":    ["project phase", "phase", "project_phase", "stage"],
        "start_date":       ["start date", "project start date", "start_date", "project start", "commenced"],
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

    df = df.copy()
    for _c in list(df.columns):
        _dt = str(df[_c].dtype).lower()
        if _dt in ('string', 'str') or 'string' in _dt or 'arrow' in _dt:
            try:
                df[_c] = df[_c].astype(object)
            except Exception:
                df[_c] = df[_c].to_numpy(dtype=object, na_value=None)

    df["non_billable"] = df["non_billable"].fillna("").astype(str).str.strip().str.upper()
    if "project" in df.columns:
        df["project"] = df["project"].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
    # Normalize project_id — strip float suffix e.g. 709470.0 -> 709470
    if "project_id" in df.columns:
        df["project_id"] = df["project_id"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    df["hours"]  = pd.to_numeric(df["hours"], errors="coerce").fillna(0)
    df["date"]   = pd.to_datetime(df["date"], errors="coerce")
    df["period"] = df["date"].dt.strftime("%Y-%m").fillna("Unknown")
    if "start_date" in df.columns:
        df["start_date"] = pd.to_datetime(df["start_date"], errors="coerce")
    if "region" not in df.columns:
        df["region"] = pd.array([""] * len(df), dtype=object)

    if "employee" in df.columns:
        _loc_map_norm = {}
        for _k, _v in EMPLOYEE_LOCATION.items():
            _loc_map_norm[_k.lower().strip()] = _v
            _loc_map_norm[" ".join(_k.split()).lower()] = _v

        def _fast_loc(emp_raw):
            emp_n = " ".join(str(emp_raw or "").split()).strip()
            emp_l = emp_n.lower()
            if not emp_l: return ""
            def _extract(v): return v[0] if isinstance(v, tuple) else (v or "")
            if emp_l in _loc_map_norm: return _extract(_loc_map_norm[emp_l])
            for _kl, _vl in _loc_map_norm.items():
                if emp_l.startswith(_kl) or _kl.startswith(emp_l): return _extract(_vl)
                if emp_l.startswith(_kl.split(",")[0].strip()): return _extract(_vl)
            return ""

        _needs_loc = df["region"].fillna("").str.strip() == ""
        if _needs_loc.any():
            df["region"] = df["region"].astype(object)
            df.loc[_needs_loc, "region"] = df.loc[_needs_loc, "employee"].map(_fast_loc)

    _ps_ovr_l = {_k.lower(): _v for _k, _v in PS_REGION_OVERRIDE.items()}
    if "employee" in df.columns:
        def _fast_ps(emp_raw):
            e = " ".join(str(emp_raw or "").split()).lower()
            for _kl, _vl in _ps_ovr_l.items():
                if e == _kl or e.startswith(_kl): return _vl
            return ""
        _emp_ps = df["employee"].map(_fast_ps)
        df["ps_region"] = _emp_ps.astype(object)
        _no_ovr = _emp_ps.fillna("") == ""
        if _no_ovr.any():
            df["ps_region"] = df["ps_region"].astype(object)
            df.loc[_no_ovr, "ps_region"] = df.loc[_no_ovr, "region"].map(
                lambda r: PS_REGION_MAP.get(str(r).strip(), "Other"))
    else:
        df["ps_region"] = df["region"].map(lambda r: PS_REGION_MAP.get(str(r).strip(), "Other"))

    if "date" in df.columns:
        df = df.sort_values(["project", "date"], na_position="first").reset_index(drop=True)

    # ── Pre-compute prior hours per project ───────────────────────────────────
    # Key prior_htd by BOTH project_id and project_name so downstream lookup
    # works regardless of which key _con_key resolves to.
    _htd_col = "hours_to_date"
    prior_htd = {}  # keyed by project_id AND project_name — both point to same value
    if _htd_col in df.columns:
        _pid_col = "project_id" if "project_id" in df.columns else "project"
        for _proj_key, _grp in df.groupby(_pid_col):
            _proj_n = " ".join(str(_grp["project"].iloc[0]).strip().split()) \
                if _pid_col == "project_id" else " ".join(str(_proj_key).strip().split())
            _pid_str = str(_proj_key).strip().replace(".0", "") \
                if _pid_col == "project_id" else None
            try:
                _max_htd    = float(_grp[_htd_col].dropna().astype(float).max() or 0)
                _period_hrs = float(_grp["hours"].dropna().astype(float).sum() or 0)
                _val = max(0.0, _max_htd - _period_hrs)
                # Store under both project name AND project_id so lookup always hits
                prior_htd[_proj_n] = _val
                if _pid_str and _pid_str not in ("nan", "None", ""):
                    prior_htd[_pid_str] = _val
            except Exception:
                prior_htd[_proj_n] = 0.0
                if _pid_str and _pid_str not in ("nan", "None", ""):
                    prior_htd[_pid_str] = 0.0

    consumed = {}
    credit_hrs_list   = []
    variance_hrs_list = []
    credit_tag_list   = []
    notes_list        = []
    htd_start_list     = []  # track starting HTD per row for output
    consumed_after_list = []  # track running consumed position after each row

    _proj_idx  = df.columns.get_loc("project")       if "project"           in df.columns else None
    _ptype_idx = df.columns.get_loc("project_type")  if "project_type"      in df.columns else None
    _hrs_idx   = df.columns.get_loc("hours")         if "hours"             in df.columns else None
    _nb_idx    = df.columns.get_loc("non_billable")  if "non_billable"      in df.columns else None
    _bt_idx    = df.columns.get_loc("billing_type")  if "billing_type"      in df.columns else None
    _sku_idx   = df.columns.get_loc("time_item_sku") if "time_item_sku"     in df.columns else None
    _drs_idx   = df.columns.get_loc("_drs_project_name") if "_drs_project_name" in df.columns else None
    _pid_idx   = df.columns.get_loc("project_id")    if "project_id"        in df.columns else None

    for _tup in df.itertuples(index=False, name=None):
        proj      = " ".join(str(_tup[_proj_idx]  if _proj_idx  is not None else "").split())
        ptype     = str(_tup[_ptype_idx] if _ptype_idx is not None else "").strip()
        hrs       = float(_tup[_hrs_idx]  if _hrs_idx  is not None else 0)
        nb        = str(_tup[_nb_idx]    if _nb_idx   is not None else "NO").strip().upper()
        bill_type = str(_tup[_bt_idx]    if _bt_idx   is not None else "").strip().lower()
        is_tm     = bill_type == "t&m"

        if hrs <= 0:
            credit_hrs_list.append(0); variance_hrs_list.append(0)
            credit_tag_list.append("SKIPPED"); notes_list.append("Zero or missing hours")
            htd_start_list.append(0)
            consumed_after_list.append(0.0)
            continue

        if bill_type == "internal":
            credit_hrs_list.append(0); variance_hrs_list.append(0)
            credit_tag_list.append("NON-BILLABLE"); notes_list.append("Internal: excluded from utilization")
            htd_start_list.append(0)
            consumed_after_list.append(0.0)
            continue

        if is_tm:
            credit_hrs_list.append(hrs); variance_hrs_list.append(0)
            credit_tag_list.append("CREDITED"); notes_list.append("T&M: full credit")
            htd_start_list.append(0)
            consumed_after_list.append(0.0)
            continue

        _ptype_lower = ptype.strip().lower()
        if "premium" in _ptype_lower:
            _sku = str(_tup[_sku_idx] if _sku_idx is not None else "") or ""
            _sku_nums = _re_constants.findall(r"IMPL(\d+)", _sku.upper())
            if _sku_nums:
                scope_hrs = float(_sku_nums[0])
            else:
                _proj_name_scope = str(_tup[_drs_idx] if _drs_idx is not None else "") or str(_tup[_proj_idx] if _proj_idx is not None else "") or ""
                _name_nums = _re_constants.findall(r"(?<![\d])(10|20)(?![\d])", _proj_name_scope)
                scope_hrs = float(_name_nums[0]) if _name_nums else None
        else:
            _matches = [(k, float(v)) for k, v in scope_map.items()
                        if k.strip().lower() in _ptype_lower]
            scope_hrs = max(_matches, key=lambda x: len(x[0]))[1] if _matches else None

        if scope_hrs is None:
            credit_hrs_list.append(0); variance_hrs_list.append(hrs)
            credit_tag_list.append("UNCONFIGURED"); notes_list.append(f"Fixed Fee but no scope defined for: {ptype}")
            htd_start_list.append(0)
            consumed_after_list.append(0.0)
            continue

        # Key on project_id if available
        _raw_pid = str(_tup[_pid_idx] if _pid_idx is not None else "").strip().replace(".0", "")
        _con_key = _raw_pid if _raw_pid and _raw_pid not in ("nan", "None", "") else proj

        if _con_key not in consumed:
            consumed[_con_key] = prior_htd.get(_con_key, prior_htd.get(proj, 0.0))

        already   = consumed[_con_key]
        remaining = scope_hrs - already
        htd_start_list.append(already)

        if remaining <= 0:
            credit_hrs_list.append(0); variance_hrs_list.append(hrs)
            credit_tag_list.append("OVERRUN"); notes_list.append(f"Scope exhausted (cap: {scope_hrs:.0f}h)")
            consumed[_con_key] = already + hrs
            consumed_after_list.append(float(already))
        elif hrs <= remaining:
            consumed[_con_key] = already + hrs
            credit_hrs_list.append(hrs); variance_hrs_list.append(0)
            credit_tag_list.append("CREDITED"); notes_list.append(f"NB within scope ({already:.1f}/{scope_hrs:.0f}h used)")
            consumed_after_list.append(float(already))
        else:
            consumed[_con_key] = already + remaining
            credit_hrs_list.append(remaining); variance_hrs_list.append(hrs - remaining)
            credit_tag_list.append("PARTIAL")
            notes_list.append(f"Split: {remaining:.2f}h credited / {hrs - remaining:.2f}h overrun")
            consumed_after_list.append(float(already))

    for _col in df.select_dtypes(include="string").columns:
        df[_col] = df[_col].astype(object)

    df["credit_hrs"]   = credit_hrs_list
    df["variance_hrs"] = variance_hrs_list
    df["credit_tag"]   = credit_tag_list
    df["notes"]        = notes_list
    df["htd_start"]    = htd_start_list
    df["previous_htd"] = [float(v) for v in consumed_after_list]

    if "task" in df.columns:
        df["ff_task"] = df["task"].fillna("").apply(match_ff_task)
    else:
        df["ff_task"] = ""

    skipped_df = df[df["credit_tag"] == "SKIPPED"][
        [c for c in ["employee","project","project_type","billing_type","date","hours","notes"] if c in df.columns]
    ].copy()

    return df, consumed, skipped_df


# ── Excel builder ─────────────────────────────────────────────────────────────
def _xl_val(val):
    if val is None: return ""
    if isinstance(val, tuple): return val[0] if val else ""
    try:
        if pd.isna(val): return ""
    except (TypeError, ValueError): pass
    return val

def build_excel(df, scope_map, consumed):
    wb  = Workbook()
    wb.remove(wb.active)
    bgs = [WHITE, LTGRAY]

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
        _m = [(k, float(v)) for k, v in scope_map.items()
              if k.strip().lower() in str(ptype).strip().lower()]
        return max(_m, key=lambda x: len(x[0]))[1] if _m else None

    def _variance_flag(tag, scoped):
        t = str(tag).strip().upper()
        if t in ("SKIPPED", "NON-BILLABLE", "CREDITED"): return "N/A"
        if t == "UNCONFIGURED": return "No Scope Set"
        if t in ("OVERRUN", "PARTIAL"): return "True Overrun" if scoped and scoped > 0 else "No Scope Set"
        return "Within Budget"

    for r_idx, (_, row) in enumerate(df.iterrows(), 3):
        tag     = str(row.get("credit_tag","")).strip()
        bg      = TAG_COLORS.get(tag, "F2F2F2")
        _scoped = _get_scoped(row.get("project_type",""))
        _vflag  = _variance_flag(tag, _scoped)
        _row_ext = dict(row)
        _row_ext["_scoped_hrs"]    = round(_scoped, 2) if _scoped is not None else ""
        _row_ext["_variance_flag"] = _vflag
        for c_idx, col in enumerate(cols, 1):
            val  = _xl_val(_row_ext.get(col, ""))
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            fmt, bold, align = None, False, "left"
            if col == "date" and pd.notna(val): fmt = "YYYY-MM-DD"; align = "center"
            elif col in ("hours","credit_hrs","variance_hrs","hours_to_date","previous_htd"):
                fmt = "#,##0.00"; align = "right"
            elif col == "credit_tag": bold = True; align = "center"
            elif col in ("period","billing_type","region"): align = "center"
            elif col == "_scoped_hrs": fmt = "#,##0.00"; align = "right"
            elif col == "_variance_flag":
                align = "center"; bold = True
                val_str = str(val)
                if val_str == "True Overrun": bg = "FFC7CE"
                elif val_str == "No Scope Set": bg = "FFEB9C"
                elif val_str == "Within Budget": bg = "C6EFCE"
                else: bg = TAG_COLORS.get(tag, "F2F2F2")
            style_cell(cell, bg, fmt=fmt, bold=bold, align=align)

    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}2"

    ws2 = wb.create_sheet("SUMMARY - By Employee")
    ws2.sheet_properties.tabColor = NAVY
    ws2.freeze_panes = "A3"

    eh = ["Employee","Location","PS Region","Period",
          "Avail Hrs (Capacity)","Hours This Period","Utilization Credits",
          "FF Project Overrun Hrs","Admin Hrs",
          "Util % (vs Hours Logged)","Util % (vs Capacity)",
          "Project Util %","Gap (pts)","Projected Full Month Util %"]
    ew = [22,16,14,12,16,15,18,18,14,20,20,16,12,22]
    write_title(ws2, "SUMMARY — Utilization by Employee", len(eh))
    style_header(ws2, 2, eh, TEAL)
    ws2.auto_filter.ref = "A2:L2"
    for i, w in enumerate(ew, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    emp_region = {}
    emp_cust_region = {}
    emp_pm = {}
    if "region" in df.columns:
        emp_region = df.dropna(subset=["region"]).groupby("employee")["region"].first().to_dict()
    if "customer_region" in df.columns:
        emp_cust_region = df.dropna(subset=["customer_region"]).groupby("employee")["customer_region"].first().to_dict()
    if "project_manager" in df.columns:
        emp_pm = df.dropna(subset=["project_manager"]).groupby("employee")["project_manager"].first().to_dict()

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

    _period_days = None
    _total_period_days = None
    if "date" in df.columns:
        _dates = df["date"].dropna()
        if len(_dates) > 0:
            import calendar
            _min_d = _dates.min(); _max_d = _dates.max()
            _period_days = len(pd.bdate_range(_min_d, _max_d))
            _yr, _mo = _min_d.year, _min_d.month
            _month_start = pd.Timestamp(_yr, _mo, 1)
            _month_end   = pd.Timestamp(_yr, _mo, calendar.monthrange(_yr, _mo)[1])
            _total_period_days = len(pd.bdate_range(_month_start, _month_end))

    _prev_emp = None; _grp_idx = 0
    for r_idx, (_, row) in enumerate(emp_sum.iterrows(), 3):
        emp    = row["employee"]
        period = row["period"]
        region = emp_region.get(emp, "")
        avail  = get_avail_hours(region, period) if region else None
        util   = row["credit_hrs"] / row["hours_this_period"] if row["hours_this_period"] > 0 else 0
        bg, _grp_idx = group_bg(emp, _prev_emp, _grp_idx)
        _prev_emp = emp
        util_bg = ("EAF9F1" if util >= 0.8 else "FEF9E7" if util >= 0.6 else "FDECED")
        _ps_reg = df[df["employee"]==emp]["ps_region"].iloc[0] if len(df[df["employee"]==emp]) > 0 else ""
        cap_util = row["credit_hrs"] / avail if avail and avail > 0 else None
        proj_util_pct = (row["credit_hrs"] + row["ff_overrun_hrs"]) / avail if avail and avail > 0 else None
        gap_pts = (proj_util_pct - cap_util) if proj_util_pct is not None and cap_util is not None else None
        proj_util = None
        if _period_days is not None and _total_period_days is not None and _period_days < _total_period_days:
            daily_rate = row["credit_hrs"] / _period_days if _period_days > 0 else 0
            proj_util = (daily_rate * _total_period_days) / avail if avail and avail > 0 else None

        vals = [emp, region, _ps_reg, period, avail or "—",
                row["hours_this_period"], row["credit_hrs"], row["ff_overrun_hrs"], row.get("admin_hrs", 0),
                util if row["hours_this_period"] > 0 else "—",
                cap_util if cap_util is not None else "—",
                proj_util_pct if proj_util_pct is not None else "—",
                gap_pts if gap_pts is not None else "—",
                proj_util if proj_util is not None else "—"]
        fmts = [None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","0.0%","0.0%","0.0%","+0.0%;-0.0%;-","0.0%"]

        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=_xl_val(val))
            is_util_col = c_idx in (10, 11, 12)
            if c_idx == 13 and isinstance(val, float):
                gap_bg = ("FCE4D6" if val >= 0.20 else "FFF2CC" if val >= 0.10 else bg)
                style_cell(cell, gap_bg, fmt=fmt, bold=(val >= 0.10), align="right")
            else:
                style_cell(cell, util_bg if is_util_col else bg, fmt=fmt,
                           align="right" if c_idx > 4 else "center" if c_idx == 4 else "left")

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

    htd_seeds = dict(zip(proj_sum["project"], proj_sum["htd_start"]))
    proj_cust_region = {}
    proj_pm = {}
    if "customer_region" in df.columns:
        proj_cust_region = df.dropna(subset=["customer_region"]).groupby("project")["customer_region"].first().to_dict()
    if "project_manager" in df.columns:
        proj_pm = df.dropna(subset=["project_manager"]).groupby("project")["project_manager"].first().to_dict()
    proj_ps_region = df.groupby("project")["ps_region"].first().to_dict() if "ps_region" in df.columns else {}
    proj_phase = {}
    if "project_phase" in df.columns:
        proj_phase = df.dropna(subset=["project_phase"]).groupby("project")["project_phase"].first().to_dict()
    proj_start = {}
    if "start_date" in df.columns:
        proj_start = df.dropna(subset=["start_date"]).groupby("project")["start_date"].min().to_dict()
    _as_of = pd.to_datetime(df["date"], errors="coerce").max() if "date" in df.columns else pd.Timestamp.now()
    df["project_phase"] = df["project"].map(proj_phase) if proj_phase else ""
    if proj_start:
        df["start_date_mapped"] = df["project"].map(proj_start)
        df["days_active"] = df["start_date_mapped"].apply(
            lambda s: int((_as_of - s).days) if pd.notna(s) and pd.notna(_as_of) else None)
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
        previous_h = max(0.0, seed - row["hours_this_period"])
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
        pm_name = proj_pm.get(row["project"], "")
        start_dt = proj_start.get(row["project"])
        vals = [row["project"], ptype, pm_name, scope_h or "—", previous_h,
                row["hours_this_period"], row["credit_hrs"], vari_h,
                previous_h + row["hours_this_period"],
                (previous_h + row["hours_this_period"]) - scope_h if scope_h > 0 else "—",
                burn if scope_h > 0 else "—", status]
        fmts = [None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","0.0%",None]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=_xl_val(val))
            style_cell(cell, status_bg if c_idx == 12 else bg, fmt=fmt, bold=(c_idx == 12),
                       align="right" if c_idx in (5,6,7,8,9,10,11) else "center" if c_idx == 12 else "left")

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

    _ot_ff = df[
        (df["credit_tag"] != "SKIPPED") &
        (df.get("billing_type", pd.Series(dtype="object")).str.lower() == "fixed fee")
    ].copy() if "billing_type" in df.columns else df[df["credit_tag"] != "SKIPPED"].copy()

    _ot_proj = _ot_ff.groupby(["project","project_type"], as_index=False).agg(
        hours_total=("hours","sum"), overrun_hrs=("variance_hrs","sum"))
    def _ot_scope(ptype):
        _m = [(k, float(v)) for k, v in scope_map.items()
              if k.strip().lower() in str(ptype).strip().lower()]
        return max(_m, key=lambda x: len(x[0]))[1] if _m else None
    _ot_proj["scoped_hrs"] = _ot_proj["project_type"].apply(_ot_scope)
    _ot_proj = _ot_proj[_ot_proj["scoped_hrs"].notna()].copy()
    _ot_proj["over_under"] = _ot_proj["hours_total"].astype(float) - _ot_proj["scoped_hrs"].astype(float)
    _ot_proj["is_over"] = (_ot_proj["over_under"] > 0).astype(int)
    _ot_type = _ot_proj.groupby("project_type", as_index=False).agg(
        total_projects=("project","count"), over_count=("is_over","sum"),
        total_overrun_hrs=("overrun_hrs","sum"), avg_scoped=("scoped_hrs","mean"),
        avg_actual=("hours_total","mean"), avg_over_under=("over_under","mean"),
    ).sort_values("total_overrun_hrs", ascending=False)
    _ot_type["over_pct"] = _ot_type["over_count"] / _ot_type["total_projects"]
    _ot_totals_proj  = int(_ot_type["total_projects"].sum())
    _ot_totals_over  = int(_ot_type["over_count"].sum())
    _ot_totals_ovhrs = float(_ot_type["total_overrun_hrs"].sum())

    for r_idx_ot, (_, row_ot) in enumerate(_ot_type.iterrows(), 3):
        bg_ot = GROUP_COLORS[r_idx_ot % 2]
        avg_ou = float(row_ot["avg_over_under"])
        over_pct = float(row_ot["over_pct"])
        vals_ot = [row_ot["project_type"], int(row_ot["total_projects"]), int(row_ot["over_count"]),
                   over_pct, float(row_ot["total_overrun_hrs"]),
                   float(row_ot["avg_scoped"]), float(row_ot["avg_actual"]), avg_ou]
        fmts_ot = [None,"#,##0","#,##0","0.0%","#,##0.0","0.0","0.0","+0.0;-0.0;\"-\""]
        for c_idx_ot, (val_ot, fmt_ot) in enumerate(zip(vals_ot, fmts_ot), 1):
            cell_ot = ws_ot.cell(row=r_idx_ot, column=c_idx_ot, value=val_ot)
            if c_idx_ot == 4:
                if over_pct >= 0.40:
                    style_cell(cell_ot, "FCE4D6", fmt=fmt_ot, bold=True, align="right")
                    cell_ot.font = Font(name="Manrope", size=10, bold=True, color="9C2A00")
                elif over_pct >= 0.25:
                    style_cell(cell_ot, "FFF2CC", fmt=fmt_ot, bold=True, align="right")
                    cell_ot.font = Font(name="Manrope", size=10, bold=True, color="7F4F00")
                else:
                    style_cell(cell_ot, bg_ot, fmt=fmt_ot, align="right")
            elif c_idx_ot == 8:
                style_cell(cell_ot, bg_ot, fmt=fmt_ot, bold=(avg_ou > 0), align="right")
                cell_ot.font = Font(name="Manrope", size=10, bold=(avg_ou>0),
                                    color="9C2A00" if avg_ou > 0 else "595959")
            elif c_idx_ot == 5 and float(row_ot["total_overrun_hrs"]) >= 500:
                style_cell(cell_ot, bg_ot, fmt=fmt_ot, bold=True, align="right")
                cell_ot.font = Font(name="Manrope", size=10, bold=True, color="9C2A00")
            else:
                style_cell(cell_ot, bg_ot, fmt=fmt_ot, align="left" if c_idx_ot == 1 else "right")
        ws_ot.row_dimensions[r_idx_ot].height = 15

    _ot_tr = 3 + len(_ot_type)
    for c_idx_ot, (val_ot, fmt_ot) in enumerate(zip(
        ["Total", _ot_totals_proj, _ot_totals_over, None, _ot_totals_ovhrs, None, None, None],
        [None, "#,##0", "#,##0", None, "#,##0.0", None, None, None]), 1):
        cell_ot = ws_ot.cell(row=_ot_tr, column=c_idx_ot, value=val_ot)
        cell_ot.font  = Font(name="Manrope", size=10, bold=True, color="FFFFFF")
        cell_ot.fill  = PatternFill("solid", fgColor=NAVY)
        cell_ot.border = thin_border()
        cell_ot.alignment = Alignment(horizontal="left" if c_idx_ot == 1 else "right", vertical="center")
        if fmt_ot and val_ot is not None: cell_ot.number_format = fmt_ot
    ws_ot.row_dimensions[_ot_tr].height = 18

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

    if "billing_type" in df.columns:
        zco_df = df[df["billing_type"].fillna("").str.lower() == "internal"].copy()
    else:
        zco_df = df[df["credit_tag"] == "NON-BILLABLE"].copy()

    if len(zco_df) > 0:
        if "task" not in zco_df.columns:
            zco_df["task"] = "Internal (no task)"
        else:
            zco_df["task"] = zco_df["task"].fillna("").replace("", "Internal (no task)")

    if "task" in zco_df.columns and len(zco_df) > 0:
        zco_sum = zco_df.groupby(["task","employee","period"], as_index=False
        ).agg(hours=("hours","sum")).sort_values(["task","employee","period"])
        emp_period_totals = df[df["credit_tag"] != "SKIPPED"].groupby(
            ["employee","period"])["hours"].sum().to_dict()
        znh[3] = "Hours"
        ws5.cell(row=2, column=5, value="% of Total Hrs").font = Font(name="Manrope", bold=True, color=WHITE, size=10)
        ws5.cell(row=2, column=5).fill = hdr_fill(TEAL)
        ws5.cell(row=2, column=5).alignment = Alignment(horizontal="center", vertical="center")
        ws5.cell(row=2, column=5).border = thin_border()
        ws5.column_dimensions["E"].width = 16
        _prev_task_z = None; _grp_idx_z = 0; r_idx = 3
        for _, row in zco_sum.iterrows():
            task = row.get("task","")
            if task != _prev_task_z:
                _task_hrs = zco_sum[zco_sum["task"]==task]["hours"].sum()
                for ci, (hval, hfmt) in enumerate([(task, None), ("", None), ("", None), (_task_hrs, "#,##0.00"), ("", None)], 1):
                    hcell = ws5.cell(row=r_idx, column=ci, value=hval)
                    hcell.font  = Font(name="Manrope", size=10, bold=True, color="FFFFFF")
                    hcell.fill  = PatternFill("solid", fgColor=NAVY)
                    hcell.border = thin_border()
                    hcell.alignment = Alignment(horizontal="right" if ci==4 else "left", vertical="center")
                    if hfmt: hcell.number_format = hfmt
                r_idx += 1; _prev_task_z = task; _grp_idx_z = 0
            bg, _grp_idx_z = group_bg(task, task, _grp_idx_z)
            total_hrs = emp_period_totals.get((row["employee"], row["period"]), 0)
            pct = row["hours"] / total_hrs if total_hrs > 0 else 0
            vals = ["", row["employee"], row["period"], row["hours"], pct]
            fmts = [None, None, None, "#,##0.00", "0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws5.cell(row=r_idx, column=c_idx, value=val)
                style_cell(cell, bg, fmt=fmt, align="right" if c_idx in (4,5) else "center" if c_idx == 3 else "left")
            r_idx += 1
    else:
        ws5.cell(row=3, column=1, value="No Non-Billable (Internal) entries in this period.")

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

    ff_df = df[(df["billing_type"].str.lower() == "fixed fee") & (df["ff_task"].notna())].copy() \
        if "billing_type" in df.columns else df[df["ff_task"].notna()].copy()

    if len(ff_df) > 0:
        task_sum = ff_df.groupby(["ff_task","project_type"], as_index=False).agg(hours=("hours","sum")).sort_values(["ff_task","project_type"])
        all_ff = df[df["billing_type"].str.lower() == "fixed fee"] if "billing_type" in df.columns else df
        proj_count_by_type = all_ff.groupby("project_type")["project"].nunique().to_dict()
        type_totals = ff_df.groupby("project_type")["hours"].sum().to_dict()
        _prev_task_t = None; _grp_idx_t = 0; r_idx_t = 3
        for _, row in task_sum.iterrows():
            ff_task = row["ff_task"]
            type_total = type_totals.get(row["project_type"], 0)
            pct = row["hours"] / type_total if type_total > 0 else 0
            proj_cnt = proj_count_by_type.get(row["project_type"], 1)
            avg_hrs = round((row["hours"] / proj_cnt if proj_cnt > 0 else 0) * 4) / 4
            if ff_task != _prev_task_t:
                _task_total_hrs = task_sum[task_sum["ff_task"]==ff_task]["hours"].sum()
                for ci, (hval, hfmt) in enumerate([(ff_task, None), ("— ALL TYPES —", None), (_task_total_hrs, "#,##0.00"), ("", None), ("", None)], 1):
                    hcell = ws6.cell(row=r_idx_t, column=ci, value=hval)
                    hcell.font  = Font(name="Manrope", size=10, bold=True, color="FFFFFF")
                    hcell.fill  = PatternFill("solid", fgColor=NAVY)
                    hcell.border = thin_border()
                    hcell.alignment = Alignment(horizontal="right" if ci==3 else "left", vertical="center")
                    if hfmt: hcell.number_format = hfmt
                r_idx_t += 1; _prev_task_t = ff_task; _grp_idx_t = 0
            task_colors = {"Configuration":"EBF5FB","Post Go-Live Consulting":"FEF9E7",
                           "Project Management":"F4ECF7","Training & UAT":"EAF9F1","Customer Communication":"FDF2F8"}
            bg, _grp_idx_t = group_bg(ff_task, ff_task, _grp_idx_t)
            task_bg = task_colors.get(ff_task, bg)
            vals = ["", row["project_type"], row["hours"], avg_hrs, pct]
            fmts = [None, None, "#,##0.00", "#,##0.00", "0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws6.cell(row=r_idx_t, column=c_idx, value=val)
                style_cell(cell, task_bg if c_idx > 1 else bg, fmt=fmt, align="right" if c_idx > 2 else "left")
            r_idx_t += 1
    else:
        ws6.cell(row=3, column=1, value="No Fixed Fee task data found. Check Billing Type and Task/Case columns.")

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
    pc_df = df[df["billing_type"].str.lower() != "internal"].copy() if "billing_type" in df.columns else df.copy()
    pc_sum = pc_df.groupby(["project_type","billing_type"], as_index=False).agg(
        project_count=("project","nunique")).sort_values(["project_type","billing_type"])
    grand_total = pc_sum["project_count"].sum()
    for r_idx, (_, row) in enumerate(pc_sum.iterrows(), 3):
        bg = LTGRAY if r_idx % 2 == 0 else WHITE
        vals = [row["project_type"], row["billing_type"], row["project_count"]]
        fmts = [None, None, "#,##0"]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_pc.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, bg, fmt=fmt, align="center" if c_idx == 2 else "right" if c_idx == 3 else "left")
    total_row = r_idx + 1 if len(pc_sum) > 0 else 3
    for c_idx, (val, fmt, bold) in enumerate([("Grand Total", None, True), ("", None, False), (grand_total, "#,##0", True)], 1):
        cell = ws_pc.cell(row=total_row, column=c_idx, value=val)
        cell.font = Font(name="Manrope", bold=True, size=10, color=WHITE)
        cell.fill = hdr_fill(NAVY); cell.border = thin_border()
        cell.alignment = Alignment(horizontal="right" if c_idx == 2 else "left", vertical="center")
        if fmt: cell.number_format = fmt

    ws_pta = wb.create_sheet("FF Project Type Analysis")
    ws_pta.sheet_properties.tabColor = "8E44AD"
    ws_pta.freeze_panes = "A4"
    ptah = ["Project Type","Task Category","Hours This Period","Avg Hrs / Project","% of Type Hrs"]
    ptaw = [28,25,16,18,16]
    write_title(ws_pta, "FF PROJECT TYPE ANALYSIS — Hours by Project Type › Task", len(ptah))
    style_header(ws_pta, 2, ptah, TEAL)
    ws_pta.auto_filter.ref = "A2:E2"
    ws_pta.cell(row=3, column=1, value="Grouped by Project Type → Task Category").font = Font(name="Manrope", size=9, italic=True, color="808080")
    for i, w in enumerate(ptaw, 1):
        ws_pta.column_dimensions[get_column_letter(i)].width = w

    if len(ff_df) > 0:
        pta_sum = ff_df.groupby(["project_type","ff_task"], as_index=False).agg(hours=("hours","sum"))
        pta_sum = pta_sum[pta_sum["ff_task"].notna() & (pta_sum["ff_task"] != "")].sort_values(["project_type","ff_task"])
        _all_ff_pta = df[df["billing_type"].fillna("").str.lower() == "fixed fee"] if "billing_type" in df.columns else df.copy()
        _proj_count_pta = _all_ff_pta.groupby("project_type")["project"].nunique().to_dict()
        _type_totals_pta = ff_df.groupby("project_type")["hours"].sum().to_dict()
        _prev_ptype_pta = None; _grp_idx_pta = 0; r_idx_pta = 4
        for _, row in pta_sum.iterrows():
            ptype_pta = row["project_type"]; ff_task_pta = row["ff_task"]
            type_total = _type_totals_pta.get(ptype_pta, 0)
            pct = row["hours"] / type_total if type_total > 0 else 0
            proj_cnt = _proj_count_pta.get(ptype_pta, 1)
            avg_hrs = round((row["hours"] / proj_cnt if proj_cnt > 0 else 0) * 4) / 4
            if ptype_pta != _prev_ptype_pta:
                _ptype_total = pta_sum[pta_sum["project_type"]==ptype_pta]["hours"].sum()
                for ci, (hval, hfmt) in enumerate([(ptype_pta, None), ("— ALL TASKS —", None), (_ptype_total, "#,##0.00"), ("", None), ("", None)], 1):
                    hcell = ws_pta.cell(row=r_idx_pta, column=ci, value=hval)
                    hcell.font = Font(name="Manrope", size=10, bold=True, color="FFFFFF")
                    hcell.fill = PatternFill("solid", fgColor=NAVY); hcell.border = thin_border()
                    hcell.alignment = Alignment(horizontal="right" if ci==3 else "left", vertical="center")
                    if hfmt: hcell.number_format = hfmt
                r_idx_pta += 1; _prev_ptype_pta = ptype_pta; _grp_idx_pta = 0
            task_colors = {"Configuration":"EBF5FB","Post Go-Live Consulting":"FEF9E7",
                           "Project Management":"F4ECF7","Training & UAT":"EAF9F1","Customer Communication":"FDF2F8"}
            bg_pta, _grp_idx_pta = group_bg(ptype_pta, ptype_pta, _grp_idx_pta)
            task_bg_pta = task_colors.get(ff_task_pta, bg_pta)
            vals = ["", ff_task_pta, row["hours"], avg_hrs, pct]
            fmts = [None, None, "#,##0.00", "#,##0.00", "0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws_pta.cell(row=r_idx_pta, column=c_idx, value=val)
                style_cell(cell, task_bg_pta if c_idx > 1 else bg_pta, fmt=fmt, align="right" if c_idx > 2 else "left")
            r_idx_pta += 1
    else:
        ws_pta.cell(row=4, column=1, value="No Fixed Fee task data available.")

    ws_cr = wb.create_sheet("By Customer Region (WIP)")
    ws_cr.sheet_properties.tabColor = "1e2c63"
    ws_cr.freeze_panes = "A3"
    crh = ["Customer Region","Hours This Period","Utilization Credits","FF Project Overrun Hrs","Util %"]
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
            hours_this_period=("hours","sum"), credit_hrs=("credit_hrs","sum"),
            ff_overrun_hrs=("variance_hrs","sum")).sort_values("customer_region")
        for r_idx, (_, row) in enumerate(cr_sum.iterrows(), 3):
            cr = row["customer_region"]; total_h = row["hours_this_period"]
            util = row["credit_hrs"] / total_h if total_h > 0 else 0
            util_bg = ("EAF9F1" if util >= 0.8 else "FEF9E7" if util >= 0.6 else "FDECED")
            bg = bgs[r_idx % 2]
            vals = [cr, total_h, row["credit_hrs"], row["ff_overrun_hrs"], util if total_h > 0 else "—"]
            fmts = [None,"#,##0.00","#,##0.00","#,##0.00","0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws_cr.cell(row=r_idx, column=c_idx, value=val)
                style_cell(cell, util_bg if c_idx == 5 else bg, fmt=fmt, align="right" if c_idx > 1 else "left")
    else:
        ws_cr.cell(row=3, column=1, value="No 'Customer Region' column found in import.")

    ws_ps = wb.create_sheet("By PS Region")
    ws_ps.sheet_properties.tabColor = "4472C4"
    ws_ps.freeze_panes = "A4"
    psh = ["PS Region","Project Type","Billing Type","Avail Hrs","Hours This Period",
           "Utilization Credits","FF Project Overrun Hrs","Admin Hrs","Util %"]
    psw = [14,28,14,12,16,18,20,14,10]
    write_title(ws_ps, "SUMMARY — Utilization by PS Region (APAC / EMEA / NOAM)", len(psh))
    style_header(ws_ps, 2, psh, TEAL)
    ws_ps.auto_filter.ref = "A3:I3"
    ws_ps.cell(row=3, column=1, value="Grouped by PS Region → Project Type → Billing Type").font = Font(name="Manrope", size=9, italic=True, color="808080")
    for i, w in enumerate(psw, 1):
        ws_ps.column_dimensions[get_column_letter(i)].width = w

    ps_avail = {}
    _seen_ep = set()
    for _emp, _grp in df.groupby("employee"):
        _loc = emp_region.get(_emp, ""); _ps = PS_REGION_MAP.get(_loc, "Other")
        for _p in _grp["period"].unique():
            if (_emp, _p) not in _seen_ep:
                _seen_ep.add((_emp, _p))
                ps_avail[_ps] = ps_avail.get(_ps, 0) + (get_avail_hours(_loc, _p) or 0)

    ps_admin = {}
    if "billing_type" in df.columns:
        for _, _ar in df[df["billing_type"].str.lower()=="internal"].iterrows():
            _ps = PS_REGION_MAP.get(_ar.get("region",""), "Other")
            ps_admin[_ps] = ps_admin.get(_ps, 0) + _ar.get("hours", 0)

    _ps_base = df[df["credit_tag"] != "SKIPPED"].copy()
    if "billing_type" not in _ps_base.columns: _ps_base["billing_type"] = "Unknown"
    _ps_detail = _ps_base.groupby(["ps_region","project_type","billing_type"], as_index=False).agg(
        hours_this_period=("hours","sum"), credit_hrs=("credit_hrs","sum"), ff_overrun_hrs=("variance_hrs","sum"))
    _ps_reg_total = _ps_base.groupby("ps_region", as_index=False).agg(
        hours_this_period=("hours","sum"), credit_hrs=("credit_hrs","sum"), ff_overrun_hrs=("variance_hrs","sum"))
    region_order = ["APAC","EMEA","NOAM","Other"]
    _ps_detail["_rord"] = _ps_detail["ps_region"].map({r:i for i,r in enumerate(region_order)}).fillna(99)
    _ps_detail = _ps_detail.sort_values(["_rord","ps_region","project_type","billing_type"]).drop(columns=["_rord"])

    r_idx = 4; _last_region = None
    for _, row in _ps_detail.iterrows():
        ps_reg = row["ps_region"]
        if ps_reg != _last_region:
            _last_region = ps_reg
            _rt = _ps_reg_total[_ps_reg_total["ps_region"]==ps_reg]
            _rh = _rt.iloc[0]["hours_this_period"] if len(_rt) else 0
            _rc = _rt.iloc[0]["credit_hrs"] if len(_rt) else 0
            _ro = _rt.iloc[0]["ff_overrun_hrs"] if len(_rt) else 0
            _ra = ps_admin.get(ps_reg, 0); _rv = ps_avail.get(ps_reg, 0)
            _ru = _rc / _rh if _rh > 0 else 0
            _ru_bg = "EAF9F1" if _ru>=0.7 else "FEF9E7" if _ru>=0.6 else "FDECED"
            reg_vals = [ps_reg, "— ALL TYPES —", "", _rv or "—", _rh, _rc, _ro, _ra, _ru if _rh > 0 else "—"]
            reg_fmts = [None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","#,##0.00","0.0%"]
            for c_idx, (val, fmt) in enumerate(zip(reg_vals, reg_fmts), 1):
                cell = ws_ps.cell(row=r_idx, column=c_idx, value=val)
                cell.font  = Font(name="Manrope", size=10, bold=True, color="FFFFFF" if c_idx <= 2 else "000000")
                cell.fill  = PatternFill("solid", fgColor=NAVY if c_idx <= 2 else (_ru_bg if c_idx == 9 else "D6DCF0"))
                cell.border = thin_border()
                if fmt: cell.number_format = fmt
                cell.alignment = Alignment(horizontal="right" if c_idx > 3 else "left", vertical="center")
            r_idx += 1
        hrs  = row["hours_this_period"]
        util = row["credit_hrs"] / hrs if hrs > 0 else 0
        util_bg = "EAF9F1" if util >= 0.7 else "FEF9E7" if util >= 0.6 else "FDECED"
        bg = bgs[r_idx % 2]
        vals = ["", row["project_type"], row["billing_type"], "", hrs, row["credit_hrs"], row["ff_overrun_hrs"], "", util if hrs > 0 else "—"]
        fmts = [None,None,None,None,"#,##0.00","#,##0.00","#,##0.00",None,"0.0%"]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_ps.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, util_bg if c_idx == 9 else bg, fmt=fmt, align="right" if c_idx > 3 else "left")
        r_idx += 1

    ws_wl = wb.create_sheet("Watch List")
    ws_wl.sheet_properties.tabColor = "E74C3C"
    ws_wl.freeze_panes = "A3"
    wlh = ["Project","Project Type","PS Region","Project Manager",
           "Scoped Hrs","Previous Hrs to Date","Hours to Date","Hours Balance","Burn %","FF Overrun Hrs","Status"]
    wlw = [35,20,14,22,12,18,14,16,10,14,12]
    write_title(ws_wl, "PROJECT WATCH LIST — Overrun & At-Risk Projects", len(wlh))
    style_header(ws_wl, 2, wlh, "E74C3C")
    ws_wl.auto_filter.ref = f"A2:{get_column_letter(len(wlh))}2"
    for i, w in enumerate(wlw, 1):
        ws_wl.column_dimensions[get_column_letter(i)].width = w

    wl_df = ff_proj_df.groupby(["project","project_type"], as_index=False).agg(
        hours_this_period=("hours","sum"), credit_hrs=("credit_hrs","sum"),
        variance_hrs=("variance_hrs","sum"), htd_start=("htd_start","first"))
    wl_df["previous_htd"] = wl_df.apply(
        lambda r: max(0.0, (float(r["htd_start"]) if r["htd_start"] else 0.0) - r["hours_this_period"]), axis=1)
    wl_df["hours_to_date"] = wl_df.apply(
        lambda r: (float(r["htd_start"]) if r["htd_start"] else 0.0), axis=1)

    def get_scope_wl(row):
        ptype = str(row.get("project_type","") or ""); pname = str(row.get("project","") or "")
        if "premium" in ptype.strip().lower() and pname:
            _nums = _re_constants.findall(r"(?<![\d])(10|20)(?![\d])", pname)
            if _nums: return float(_nums[0])
        _pm = [(k, float(v)) for k, v in scope_map.items() if k.strip().lower() in ptype.strip().lower()]
        return max(_pm, key=lambda x: len(x[0]))[1] if _pm else 0

    wl_df["scope_h"] = wl_df.apply(get_scope_wl, axis=1)
    wl_df["burn_pct"] = wl_df.apply(
        lambda r: (float(r["htd_start"]) if r["htd_start"] else 0) / r["scope_h"] if r["scope_h"] > 0 else None, axis=1)
    def _wl_status(r):
        s_h = r["scope_h"] or 0; htd = (float(r["htd_start"]) if r["htd_start"] else 0)
        if s_h > 0 and htd > s_h:   return "OVERRUN"
        if s_h > 0 and htd == s_h:  return "AT LIMIT"
        if s_h == 0 and htd > 0:    return "REVIEW"
        if (r["burn_pct"] or 0) >= 0.9: return "REVIEW"
        return "ON TRACK"
    wl_df["status"] = wl_df.apply(_wl_status, axis=1)
    watchlist = wl_df[wl_df["status"].isin(["OVERRUN","AT LIMIT","REVIEW"])].sort_values(
        "burn_pct", ascending=False, na_position="last")

    r_idx = 3
    for _, row in watchlist.iterrows():
        status = row["status"]
        status_bg = "FDECED" if status == "OVERRUN" else "FEF9E7"
        burn_val = row["burn_pct"] if row["burn_pct"] is not None else "—"
        pm_name = proj_pm.get(row["project"], "")
        start_dt = proj_start.get(row["project"])
        _htd_wl = float(row["previous_htd"]) + float(row["hours_this_period"])
        _tot_ov = _htd_wl - row["scope_h"] if row["scope_h"] and row["scope_h"] > 0 else "—"
        _ps_reg_wl = proj_ps_region.get(row["project"], "")
        vals = [row["project"], row["project_type"], _ps_reg_wl, pm_name,
                row["scope_h"] or "—", row["previous_htd"], _htd_wl, _tot_ov,
                burn_val, row["variance_hrs"], status]
        fmts = [None,None,None,None,"#,##0.00","#,##0.00","#,##0.00","#,##0.00","0.0%","#,##0.00",None]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_wl.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, status_bg if c_idx == 11 else status_bg, fmt=fmt, bold=(c_idx == 9),
                       align="right" if c_idx in (5,6,7,8) else "center" if c_idx == 9 else "left")
        r_idx += 1

    r_idx += 1
    unconf_title_cell = ws_wl.cell(row=r_idx, column=1, value="FF: NO SCOPE DEFINED (Hours at Risk)")
    unconf_title_cell.font  = Font(name="Manrope", bold=True, size=11, color="FFFFFF")
    unconf_title_cell.fill  = hdr_fill("E67E22")
    ws_wl.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=len(wlh))
    r_idx += 1
    unconf_df = df[df["credit_tag"] == "UNCONFIGURED"].groupby(
        ["project","project_type"], as_index=False).agg(hours=("hours","sum")).sort_values("hours", ascending=False)
    for _, row in unconf_df.iterrows():
        bg = "FEF3E2"
        vals = [row["project"], row["project_type"], proj_cust_region.get(row["project"],""),
                proj_pm.get(row["project"],""), "—", "—", "—", row["hours"], "FF: NO SCOPE DEFINED"]
        fmts = [None,None,None,None,None,None,None,"#,##0.00",None]
        for c_idx, (val, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws_wl.cell(row=r_idx, column=c_idx, value=val)
            style_cell(cell, bg, fmt=fmt, align="right" if c_idx == 8 else "left")
        r_idx += 1

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
        c.font = Font(name="Manrope", size=11, bold=True, color="FFFFFF")
        c.fill = hdr_fill(NAVY)
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

    dash_section(ws_dash, 6, 2, "KEY METRICS", ncols=6)
    ws_dash.row_dimensions[5].height = 22
    hours_tp_d   = df[df["credit_tag"] != "SKIPPED"]["hours"].sum()
    credit_hrs_d = df[df["credit_tag"].isin(["CREDITED","PARTIAL"])]["credit_hrs"].sum()
    overrun_hrs_d = df[df["credit_tag"].isin(["OVERRUN", "PARTIAL"])]["variance_hrs"].sum()
    admin_hrs_d  = df[df["billing_type"].str.lower()=="internal"]["hours"].sum() if "billing_type" in df.columns else 0
    util_pct_d   = credit_hrs_d / hours_tp_d if hours_tp_d > 0 else 0
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
        if status: rag_cell(ws_dash, 8, col, value, fmt=fmt, status=status)
        else: dash_value(ws_dash, 8, col, value, fmt=fmt, size=14)
    ws_dash.row_dimensions[8].height = 28

    dash_section(ws_dash, 10, 2, "UTILIZATION BY PS REGION", ncols=6)
    ws_dash.row_dimensions[9].height = 22
    for ci, hdr in enumerate(["PS Region","Hours This Period","Credit Hrs","Util %","FF Overrun Hrs","Admin Hrs"], 2):
        c = ws_dash.cell(row=11, column=ci, value=hdr)
        c.font = Font(name="Manrope", size=9, bold=True, color="FFFFFF")
        c.fill = hdr_fill(TEAL)

    ps_base_d = df[df["credit_tag"] != "SKIPPED"]
    ps_sum_d  = ps_base_d.groupby("ps_region").agg(hours=("hours","sum"), credit=("credit_hrs","sum"), overrun=("variance_hrs","sum"))
    ps_admin_d = df[df["billing_type"].str.lower()=="internal"].groupby("ps_region")["hours"].sum() if "billing_type" in df.columns else pd.Series(dtype=float)
    ps_avail_d = {}
    _seen_emp_period = set()
    for _emp2, _grp2 in df.groupby("employee"):
        _loc2 = emp_region.get(_emp2,""); _ps2 = PS_REGION_MAP.get(_loc2,"Other")
        for _p2 in _grp2["period"].unique():
            if (_emp2, _p2) not in _seen_emp_period:
                _seen_emp_period.add((_emp2, _p2))
                ps_avail_d[_ps2] = ps_avail_d.get(_ps2,0) + (get_avail_hours(_loc2,_p2) or 0)

    for ri, reg in enumerate(["APAC","EMEA","NOAM","Other"], 12):
        if reg not in ps_sum_d.index: continue
        _row = ps_sum_d.loc[reg]
        _adm = float(ps_admin_d.get(reg,0)) if reg in ps_admin_d.index else 0
        _util= _row["credit"] / _row["hours"] if _row["hours"] > 0 else None
        _bg  = bgs[ri % 2]
        _util_color = ("E74C3C" if _util<0.60 else "2ECC71" if _util>=0.70 else "F39C12") if _util is not None else "808080"
        _dash_ps_vals = [
            (2, reg, None, False, "000000"),
            (3, _row["hours"], "#,##0.00", False, "000000"),
            (4, _row["credit"], "#,##0.00", False, "000000"),
            (5, _util if _util is not None else "—", "0.0%" if _util is not None else None, True, _util_color),
            (6, _row["overrun"], "#,##0.00", False, "000000"),
            (7, _adm, "#,##0.00", False, "000000"),
        ]
        for ci2, val2, fmt2, bold2, color2 in _dash_ps_vals:
            _c = ws_dash.cell(row=ri, column=ci2, value=val2)
            _c.font = Font(name="Manrope", size=10, bold=bold2, color=color2)
            _c.fill = PatternFill("solid", fgColor=_bg)
            _c.border = thin_border()
            if fmt2: _c.number_format = fmt2

    dash_section(ws_dash, 17, 2, "WATCH LIST SUMMARY", ncols=6)
    ws_dash.row_dimensions[17].height = 22
    n_overrun  = len(wl_df[wl_df["status"]=="OVERRUN"]) if len(wl_df) > 0 else len(df[df["credit_tag"]=="OVERRUN"]["project"].unique())
    _wl_at_risk = wl_df[(wl_df["burn_pct"].notna()) & (wl_df["burn_pct"]>=0.9) & (wl_df["status"]!="OVERRUN")]
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

    dash_section(ws_dash, 21, 2, "EMPLOYEES BELOW 60% UTILIZATION — Action Required", ncols=6)
    ws_dash.row_dimensions[21].height = 22
    for ci, hdr in enumerate(["Employee","Location","PS Region","Period","Util %","Credit Hrs"], 2):
        _c = ws_dash.cell(row=22, column=ci, value=hdr)
        _c.font = Font(name="Manrope", size=9, bold=True, color="FFFFFF")
        _c.fill = hdr_fill(TEAL)

    _low_rows = []
    for _, _erow in emp_sum.iterrows():
        _emp3  = _erow["employee"]
        if any(_emp3.lower().startswith(ex.lower()) for ex in UTIL_EXEMPT_EMPLOYEES): continue
        _loc3  = emp_region.get(_emp3,""); _ps3 = PS_REGION_MAP.get(_loc3,"Other")
        _p3    = _erow["period"]; _avl3 = get_avail_hours(_loc3, _p3) or 0
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

    sheet_order = [
        "Dashboard", "Project Count", "SUMMARY - By Employee", "FF By Project Utilization",
        "FF Overrun by Project Type", "By Customer Region (WIP)", "By PS Region",
        "Watch List", "Non-Billable", "Task Analysis", "FF Project Type Analysis",
        "Skipped Rows", "PROCESSED_DATA",
    ]
    existing  = [s for s in sheet_order if s in wb.sheetnames]
    remaining = [s for s in wb.sheetnames if s not in existing]
    wb._sheets = [wb[s] for s in existing + remaining]

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


def _run_utilization_engine(df_raw, period_start, period_end):
    """
    Single cached compute path. Returns dict with df (post-credit-assignment), consumed, empty.
    DRS enrichment preserved from original.
    """
    import pandas as _pd
    _ts = _pd.Timestamp(period_start)
    _te = _pd.Timestamp(period_end) + _pd.Timedelta(days=1) - _pd.Timedelta(seconds=1)

    if "date" in df_raw.columns:
        _dated = _pd.to_datetime(df_raw["date"], errors="coerce")
        _mask = (_dated >= _ts) & (_dated <= _te)
        _df_period = df_raw[_mask].copy()
    else:
        _df_period = df_raw.copy()

    if _df_period.empty:
        return {"df": _df_period, "consumed": {}, "empty": True}

    _df_drs_enrich = st.session_state.get("df_drs")
    if _df_drs_enrich is not None and not _df_drs_enrich.empty:
        if "project_id" in _df_drs_enrich.columns and "project_name" in _df_drs_enrich.columns:
            _drs_name_map = dict(zip(
                _df_drs_enrich["project_id"].astype(str).str.strip(),
                _df_drs_enrich["project_name"].astype(str).str.strip()
            ))
            if "project" in _df_period.columns:
                _df_period["_drs_project_name"] = _df_period["project"].astype(str).str.strip().map(
                    lambda pid: _drs_name_map.get(pid, ""))
                _drs_name_by_name = {v: v for v in _df_drs_enrich["project_name"].astype(str).str.strip()}
                _empty_mask = _df_period["_drs_project_name"] == ""
                _df_period.loc[_empty_mask, "_drs_project_name"] = \
                    _df_period.loc[_empty_mask, "project"].map(
                        lambda n: next((v for v in _drs_name_by_name if n.lower() in v.lower() or v.lower() in n.lower()), ""))

    df, consumed, _skipped = assign_credits(_df_period, DEFAULT_SCOPE)
    return {"df": df, "consumed": consumed, "empty": False}


def _ns_signature(df_ns):
    if df_ns is None or df_ns.empty:
        return ("empty",)
    try:
        _max_d = pd.to_datetime(df_ns["date"], errors="coerce").max()
        return (len(df_ns), str(_max_d))
    except Exception:
        return (len(df_ns),)


def _proj_id_col(df):
    """Return the column name to use as the project ID for distinct counting."""
    for c in ("project_id", "project_internal_id", "project"):
        if c in df.columns:
            return c
    return None


def _billable_mask(df):
    """Boolean mask: rows that are billable (non-internal, non-skipped)."""
    if "billing_type" in df.columns and "credit_tag" in df.columns:
        return (df["billing_type"].fillna("").str.lower() != "internal") & (df["credit_tag"] != "SKIPPED")
    if "credit_tag" in df.columns:
        return df["credit_tag"].isin(["CREDITED", "PARTIAL", "OVERRUN", "UNCONFIGURED"])
    return pd.Series([True] * len(df), index=df.index)


def main():
    # ─────────────────────────────────────────────────────
    # Theme-aware styles using Streamlit's CSS variables
    # ─────────────────────────────────────────────────────
    st.markdown("""
        <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
            h1, h2, h3, .stMarkdown, .stDataFrame, label, button { font-family: 'Manrope', sans-serif !important; }

            /* Theme-aware card surfaces */
            /* Card surfaces — no explicit background; inherit page bg so cards
               flip with theme automatically (mirrors Portfolio Analytics pattern). */
            .util-card,
            .util-control-bar,
            .util-kpi,
            .util-legend,
            .util-callout,
            .util-table-header,
            .util-table-wrap {
                border: 1px solid rgba(128,128,128,0.25);
                color: inherit;
            }
            .util-card        { border-radius: 8px; padding: 14px; }
            .util-control-bar { border-radius: 10px; padding: 12px 16px; margin-bottom: 10px; }
            .util-kpi         { border-radius: 8px; padding: 14px; }
            .util-legend      { border-radius: 8px; padding: 10px 14px; margin-bottom: 8px;
                                display: flex; flex-wrap: wrap; gap: 8px;
                                align-items: center; font-size: 12px; }
            .util-callout     { border-radius: 8px; padding: 12px 14px; }
            .util-table-header{ border-radius: 8px 8px 0 0; padding: 10px 14px;
                                display: flex; justify-content: space-between; font-size: 12px; }
            .util-table-wrap  { border-top: none; border-radius: 0 0 8px 8px; overflow: hidden; }

            .util-meta-row {
                margin-top: 10px; padding-top: 10px;
                border-top: 1px solid rgba(128,128,128,0.18);
                display: flex; gap: 14px; font-size: 12px;
                opacity: 0.85;
                align-items: center; flex-wrap: wrap;
            }
            .util-live-dot {
                width: 7px; height: 7px; border-radius: 50%;
                background: #22c55e; display: inline-block; margin-right: 5px;
            }

            .util-kpi-label { font-size: 11px; opacity: 0.75; margin-bottom: 4px;
                              white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
            .util-kpi-value { font-size: 24px; font-weight: 600; line-height: 1.1;
                              font-variant-numeric: tabular-nums; }
            .util-kpi-sub   { font-size: 11px; opacity: 0.65; margin-top: 4px; }
            .util-kpi-pill  { display: inline-block; margin-top: 4px;
                              padding: 2px 9px; border-radius: 999px; font-size: 11px; }

            /* Pills (translucent so they work on either page bg) */
            .util-pill { padding: 3px 9px; border-radius: 999px;
                         font-size: 12px; white-space: nowrap; font-weight: 500; }
            .util-pill-green { background: rgba(34, 197, 94, 0.18); color: #15803d; }
            .util-pill-amber { background: rgba(245, 158, 11, 0.18); color: #b45309; }
            .util-pill-red   { background: rgba(239, 68, 68, 0.18); color: #b91c1c; }
            .util-pill-blue  { background: rgba(59, 130, 246, 0.18); color: #1d4ed8; }
            .util-pill-grey  { background: rgba(128, 128, 128, 0.18); color: inherit; opacity: 0.75; }
            @media (prefers-color-scheme: dark) {
                .util-pill-green { color: #7ed4a4; }
                .util-pill-amber { color: #f5b958; }
                .util-pill-red   { color: #f08585; }
                .util-pill-blue  { color: #6fa8dc; }
            }
            .stApp[data-theme="dark"] .util-pill-green { color: #7ed4a4; }
            .stApp[data-theme="dark"] .util-pill-amber { color: #f5b958; }
            .stApp[data-theme="dark"] .util-pill-red   { color: #f08585; }
            .stApp[data-theme="dark"] .util-pill-blue  { color: #6fa8dc; }

            .util-callout-red    { border-left: 3px solid #ef4444; }
            .util-callout-amber  { border-left: 3px solid #f59e0b; }
            .util-callout-blue   { border-left: 3px solid #3b82f6; }
            .util-callout-green  { border-left: 3px solid #22c55e; }
            .util-callout-title  { font-size: 12px; font-weight: 600; }
            .util-callout-num    { font-size: 22px; font-weight: 600;
                                   font-variant-numeric: tabular-nums; }
            .util-callout-body   { font-size: 12px; opacity: 0.8; line-height: 1.4; }

            .util-legend-label {
                opacity: 0.6; text-transform: uppercase;
                letter-spacing: 0.6px; font-size: 11px; margin-right: 4px;
            }

            /* Employee HTML table */
            .util-emp-table {
                width: 100%; border-collapse: collapse;
                font-family: 'Manrope', sans-serif; font-size: 13px;
                font-variant-numeric: tabular-nums;
                color: inherit;
            }

            .util-emp-table thead tr {
                background: rgba(128,128,128,0.08);
            }
            .util-emp-table th {
                padding: 10px 12px; font-weight: 600; text-align: left;
                opacity: 0.75;
                border-bottom: 1px solid rgba(128,128,128,0.25);
            }
            .util-emp-table th.num    { text-align: right; }
            .util-emp-table th.center { text-align: center; }
            .util-emp-table tbody tr  { border-bottom: 1px solid rgba(128,128,128,0.15); }
            .util-emp-table td        { padding: 9px 12px; vertical-align: middle; }
            .util-emp-table td.num    { text-align: right; }
            .util-emp-table td.center { text-align: center; }
            .util-emp-table td.muted  { opacity: 0.6; }
            .util-emp-avatar {
                display: inline-flex; align-items: center; justify-content: center;
                width: 24px; height: 24px; border-radius: 50%;
                font-size: 10px; font-weight: 600; margin-right: 8px;
                vertical-align: middle; flex-shrink: 0;
            }
            .util-emp-name { display: inline-flex; align-items: center; }
            .util-table-foot {
                padding: 8px 14px;
                border-top: 1px solid rgba(128,128,128,0.15);
                display: flex; justify-content: space-between;
                font-size: 11px; opacity: 0.6;
            }

            /* Sleek reference card (replaces rainbow tables) */
            .util-ref-grid {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 16px;
            }
            .util-ref-section h4 {
                font-size: 11px; font-weight: 600; opacity: 0.65;
                text-transform: uppercase; letter-spacing: 0.8px;
                margin: 0 0 10px 0;
            }
            .util-ref-row {
                display: grid;
                grid-template-columns: auto 1fr;
                gap: 12px; padding: 8px 0;
                border-bottom: 1px solid rgba(128,128,128,0.12);
                align-items: start;
                font-size: 12px;
            }
            .util-ref-row:last-child { border-bottom: none; }
            .util-ref-tag {
                padding: 2px 8px; border-radius: 999px;
                font-size: 11px; font-weight: 500;
                white-space: nowrap;
            }
            .util-ref-desc { opacity: 0.85; line-height: 1.5; }

            /* Movers row */
            .util-mover-row {
                display: flex; justify-content: space-between; align-items: center;
                padding: 7px 0; font-size: 12px;
                border-bottom: 1px solid rgba(128,128,128,0.1);
            }
            .util-mover-row:last-child { border-bottom: none; }
            .util-mover-vals {
                display: inline-flex; align-items: center; gap: 6px;
                font-variant-numeric: tabular-nums;
            }
            .util-mover-from { opacity: 0.6; }

            /* Share bar */
            .util-share-row {
                margin-bottom: 8px;
            }
            .util-share-head {
                display: flex; justify-content: space-between;
                font-size: 12px; margin-bottom: 3px;
            }
            .util-share-bar-bg {
                background: rgba(128,128,128,0.15);
                height: 8px; border-radius: 4px; overflow: hidden;
            }
            .util-share-bar-fg {
                height: 100%; background: #3b82f6; border-radius: 4px;
            }
        </style>
    """, unsafe_allow_html=True)

    # Hero (intentionally dark — brand element, like other pages)
    st.markdown(
        "<div style='background:#050D1F;padding:28px 36px 24px;border-radius:10px;"
        "margin-bottom:14px;font-family:Manrope,sans-serif;'>"
        "<div style='font-size:10px;font-weight:700;letter-spacing:2.5px;"
        "text-transform:uppercase;color:#3B9EFF;margin-bottom:8px'>Professional Services · Tools</div>"
        "<h1 style='color:white;margin:0;font-size:26px;'>Utilization Report</h1>"
        "<p style='color:#aac4d0;margin:6px 0 0 0;font-size:13px;'>"
        "Live utilization credits and capacity from NetSuite. Adjust the period — everything below recomputes automatically.</p>"
        "</div>", unsafe_allow_html=True)

    # ─────────────────────────────────────────────────────
    # Data sourcing (unchanged)
    # ─────────────────────────────────────────────────────
    _ns_from_session = st.session_state.get("df_ns")
    if _ns_from_session is None:
        st.info("Upload NS Time Detail on the Home page to generate the Utilization Report.")
        return

    df_raw = _ns_from_session.copy()

    _logged_in = st.session_state.get("consultant_name", "")
    from shared.constants import get_role as _gr, resolve_view_as, get_region_consultants
    from shared.config import EMPLOYEE_LOCATION as _EL_u, PS_REGION_MAP as _RM_u, PS_REGION_OVERRIDE as _RO_u
    from shared.constants import ACTIVE_EMPLOYEES as _AE_u
    _role_u = _gr(_logged_in)
    _is_mgr_u = _role_u in ("manager", "manager_only", "reporting_only")

    _home_browse_u = (st.session_state.get("_browse_passthrough") or
                      st.session_state.get("home_browse", ""))

    if _is_mgr_u and not _home_browse_u:
        from shared.constants import CONSULTANT_DROPDOWN as _CD_u
        _active_c_u = sorted([n for n in _CD_u if _gr(n) in ("consultant", "manager")])
        _by_rgn_u = {}
        for _cn_u in _active_c_u:
            _loc_u = _EL_u.get(_cn_u, "")
            _rg_u  = _RO_u.get(_cn_u, _RM_u.get(_loc_u, "Other"))
            _by_rgn_u.setdefault(_rg_u, []).append(_cn_u)
        _bopts_u = ["👥 All team"]
        for _rg_u in sorted(_by_rgn_u.keys()):
            _bopts_u.append(f"── {_rg_u} ──")
            _bopts_u.extend(_by_rgn_u[_rg_u])
        with st.sidebar:
            st.markdown("**View as:**")
            _home_browse_u = st.selectbox("View as", _bopts_u, key="util_view_as", label_visibility="collapsed")

    _browse_clean = _home_browse_u.replace("👥", "").strip() if _home_browse_u else ""
    _is_all_team  = _browse_clean.lower() in ("all team", "") or _home_browse_u in ("👥 All team", "All team")

    if _is_mgr_u and _is_all_team:
        _va_name_u, _va_region_u = None, None
    else:
        _va_name_u, _va_region_u, _ = resolve_view_as(
            _logged_in, _home_browse_u, EMPLOYEE_ROLES, _EL_u, _RM_u, _RO_u, _AE_u)

    if "employee" in df_raw.columns:
        from shared.constants import name_matches
        if _va_region_u:
            _rc_u = get_region_consultants(_va_region_u, _EL_u, _RM_u, _RO_u, _AE_u)
            _filtered_u = df_raw[df_raw["employee"].astype(str).str.strip().str.lower().isin(_rc_u)]
            if not _filtered_u.empty: df_raw = _filtered_u
        elif _va_name_u:
            _filtered_u = df_raw[df_raw["employee"].apply(lambda v: name_matches(v, _va_name_u))]
            if not _filtered_u.empty: df_raw = _filtered_u
        elif not _is_mgr_u and _logged_in:
            _filtered_u = df_raw[df_raw["employee"].apply(lambda v: name_matches(v, _logged_in))]
            if not _filtered_u.empty: df_raw = _filtered_u

    _raw_dates = pd.to_datetime(df_raw["date"], errors="coerce").dropna() if "date" in df_raw.columns else pd.Series(dtype="datetime64[ns]")
    _min_date  = _raw_dates.min().date() if not _raw_dates.empty else date(date.today().year, 1, 1)
    _max_date  = _raw_dates.max().date() if not _raw_dates.empty else date.today()

    # ─────────────────────────────────────────────────────
    # Sticky control bar
    # ─────────────────────────────────────────────────────
    today = date.today()
    _first_this_month_ts = pd.Timestamp(today.year, today.month, 1)
    _last_month_end = (_first_this_month_ts - pd.Timedelta(days=1)).date()
    _last_month_start = _last_month_end.replace(day=1)
    period_options = {
        "This month":    (date(today.year, today.month, 1), today),
        "Last month":    (_last_month_start, _last_month_end),
        "QTD":           (date(today.year, ((today.month - 1) // 3) * 3 + 1, 1), today),
        "YTD":           (date(today.year, 1, 1), today),
        "All available": (_min_date, _max_date),
        "Custom":        (None, None),
    }

    c1, c2, c3, c4, c5 = st.columns([1.2, 1.2, 1.2, 1.0, 1.0])
    with c1:
        preset = st.selectbox("Period", list(period_options.keys()),
                              index=0, key="util_period_preset")
    with c2:
        if preset == "Custom":
            _default_start = st.session_state.get("util_period_start", _min_date)
            period_start = st.date_input("Start", value=_default_start,
                                         min_value=date(2020, 1, 1), max_value=_max_date,
                                         key="util_period_start")
        else:
            _ps, _ = period_options[preset]
            period_start = max(_ps, _min_date)
            st.markdown(f"<div style='font-size:11px;opacity:0.6;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:4px'>Start</div>"
                        f"<div style='border:1px solid rgba(128,128,128,0.25);border-radius:6px;padding:6px 10px;font-size:13px;opacity:0.85'>{period_start.strftime('%Y-%m-%d')}</div>",
                        unsafe_allow_html=True)
    with c3:
        if preset == "Custom":
            _default_end = st.session_state.get("util_period_end", _max_date)
            period_end = st.date_input("End", value=_default_end,
                                       min_value=period_start, max_value=date(2030, 12, 31),
                                       key="util_period_end")
        else:
            _, _pe = period_options[preset]
            period_end = min(_pe, _max_date)
            st.markdown(f"<div style='font-size:11px;opacity:0.6;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:4px'>End</div>"
                        f"<div style='border:1px solid rgba(128,128,128,0.25);border-radius:6px;padding:6px 10px;font-size:13px;opacity:0.85'>{period_end.strftime('%Y-%m-%d')}</div>",
                        unsafe_allow_html=True)
    with c4:
        st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
        refresh_clicked = st.button("↻ Refresh", key="util_refresh",
                                    help="Re-pull from source. Only needed if NS data has changed but inputs haven't.")
    with c5:
        st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
        _download_slot = st.empty()

    # ─────────────────────────────────────────────────────
    # Auto-run cached engine
    # ─────────────────────────────────────────────────────
    if "_util_cache" not in st.session_state:
        st.session_state._util_cache = {}

    _cache_key = (
        _ns_signature(_ns_from_session),
        str(period_start), str(period_end),
        _va_name_u, _va_region_u, _is_all_team,
    )
    if refresh_clicked:
        st.session_state._util_cache.pop(_cache_key, None)
        # Also clear any wider-window engine results (used by trend/movers)
        for _k in list(st.session_state._util_cache.keys()):
            st.session_state._util_cache.pop(_k, None)

    if _cache_key in st.session_state._util_cache:
        result = st.session_state._util_cache[_cache_key]
    else:
        with st.spinner("Computing..."):
            try:
                result = _run_utilization_engine(df_raw, period_start, period_end)
                st.session_state._util_cache[_cache_key] = result
            except Exception as e:
                st.error(f"Processing error: {e}")
                st.exception(e)
                return

    if result.get("empty"):
        st.markdown(
            f"<div class='util-meta-row'>"
            f"<span><span class='util-live-dot'></span><b>Live</b></span>"
            f"<span>NS data: {_min_date.strftime('%-d %b %Y')} → {_max_date.strftime('%-d %b %Y')}</span>"
            f"<span>·</span><span>No entries for {period_start.strftime('%-d %b %Y')} → {period_end.strftime('%-d %b %Y')}</span>"
            f"</div>", unsafe_allow_html=True)
        return

    df = result["df"]
    consumed = result["consumed"]

    # ─────────────────────────────────────────────────────
    # KPI / summary numbers
    # ─────────────────────────────────────────────────────
    bmask = _billable_mask(df)
    df_bill = df[bmask]

    hours_this_period   = df[df["credit_tag"] != "SKIPPED"]["hours"].sum() if "hours" in df.columns else 0
    total_credit        = df["credit_hrs"].sum()
    total_admin         = df[df["billing_type"].str.lower() == "internal"]["hours"].sum() if "billing_type" in df.columns else 0
    total_proj_overrun  = df[df["credit_tag"].isin(["OVERRUN", "PARTIAL"])]["variance_hrs"].sum() if "variance_hrs" in df.columns else 0

    credit_pct  = total_credit       / hours_this_period if hours_this_period else 0
    overrun_pct = total_proj_overrun / hours_this_period if hours_this_period else 0
    admin_pct   = total_admin        / hours_this_period if hours_this_period else 0
    credit_label = "On target" if credit_pct >= 0.70 else "Below target" if credit_pct >= 0.60 else "At risk"

    # Billable project count by project ID (the fix)
    _pid_col = _proj_id_col(df_bill)
    billable_proj_count = df_bill[_pid_col].astype(str).str.strip().nunique() if _pid_col else 0
    consultant_count = df["employee"].nunique() if "employee" in df.columns else 0

    # Capacity
    _avail_total = 0
    if "employee" in df.columns and "region" in df.columns and "period" in df.columns:
        _emp_region_ui = df.dropna(subset=["region"]).groupby("employee")["region"].first().to_dict()
        _emp_period_ui = df.groupby("employee")["period"].first().to_dict()
        for _e, _r in _emp_region_ui.items():
            _av = get_avail_hours(_r, _emp_period_ui.get(_e, ""))
            if _av: _avail_total += _av
    capacity_pct = hours_this_period / _avail_total if _avail_total else 0

    # Partial period detection
    import calendar as _cal
    _bdays_in = len(pd.bdate_range(period_start, period_end))
    _yr2, _mo2 = period_start.year, period_start.month
    _ms = pd.Timestamp(_yr2, _mo2, 1)
    _me = pd.Timestamp(_yr2, _mo2, _cal.monthrange(_yr2, _mo2)[1])
    _bdays_total_month = len(pd.bdate_range(_ms, _me))
    is_partial = (period_start.month == period_end.month and period_start.year == period_end.year and _bdays_in < _bdays_total_month)

    _last_refresh = st.session_state.get("_util_last_refresh", datetime.now())
    if refresh_clicked or _cache_key not in st.session_state._util_cache:
        st.session_state._util_last_refresh = datetime.now()
        _last_refresh = st.session_state._util_last_refresh
    _ago_min = max(0, int((datetime.now() - _last_refresh).total_seconds() // 60))
    _ago_str = "just now" if _ago_min == 0 else f"{_ago_min} min ago"

    _partial_str = f" · partial period ({_bdays_in} of {_bdays_total_month} days)" if is_partial else ""
    st.markdown(
        f"<div class='util-meta-row'>"
        f"<span><span class='util-live-dot'></span><b>Live</b></span>"
        f"<span>NS data: {_min_date.strftime('%-d %b %Y')} → {_max_date.strftime('%-d %b %Y')}</span>"
        f"<span>·</span>"
        f"<span>Showing <b>{billable_proj_count} billable projects</b> · <b>{hours_this_period:,.2f} hrs</b>{_partial_str}</span>"
        f"<span style='margin-left:auto;opacity:0.6'>Updated {_ago_str}</span>"
        f"</div>", unsafe_allow_html=True)

    # Excel download in reserved slot
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"utilization_report_{timestamp}.xlsx"
    try:
        excel_buf = build_excel(df, DEFAULT_SCOPE, consumed)
        with _download_slot:
            st.download_button(
                "⬇ Excel", data=excel_buf, file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", key="util_dl_excel_top",
            )
    except Exception:
        pass

    # ─────────────────────────────────────────────────────
    # Unmapped / alumni warnings
    # ─────────────────────────────────────────────────────
    _unmapped = []; _alumni = []
    _period_str = str(df["period"].mode().iloc[0]) if "period" in df.columns and len(df) > 0 else None
    for _emp in df["employee"].dropna().unique():
        _emp_s = str(_emp).strip()
        _loc = df[df["employee"]==_emp]["region"].iloc[0] if len(df[df["employee"]==_emp]) > 0 else ""
        _matched_key = None
        for k in EMPLOYEE_LOCATION:
            if _emp_s.lower() == k.lower() or _emp_s.lower().startswith(k.lower()) or k.lower().startswith(_emp_s.lower()):
                _matched_key = k; break
        if not _matched_key and not str(_loc).strip(): _unmapped.append(_emp_s)
        elif _matched_key and _period_str:
            if not _emp_active(_matched_key, _period_str): _alumni.append(_emp_s)
    if _unmapped:
        st.warning(f"**{len(_unmapped)} employee(s) have no location defined** — avail hours and PS region will show as Unknown.\n\n" + ", ".join(sorted(_unmapped)))
    if _alumni:
        st.info(f"**{len(_alumni)} employee(s) have time entries but are outside their active tenure** — excluded from utilization targets.\n\n" + ", ".join(sorted(_alumni)))

    # ─────────────────────────────────────────────────────
    # Helpers used in tabs
    # ─────────────────────────────────────────────────────
    def fmt_hrs(n):
        s = f"{n:,.2f}"
        if "." in s: s = s.rstrip("0").rstrip(".")
        return s

    def rag_pill_html(v, exempt=False):
        if exempt: return "<span class='util-pill util-pill-grey' style='opacity:0.7'>exempt</span>"
        if v is None: return "<span style='opacity:0.5'>—</span>"
        if v >= 0.70:   cls = "util-pill-green"
        elif v >= 0.60: cls = "util-pill-amber"
        else:           cls = "util-pill-red"
        return f"<span class='util-pill {cls}'>{v*100:.1f}%</span>"

    def avatar_html(name):
        parts = [p.strip() for p in str(name).replace(",", " ").split() if p.strip()]
        initials = (parts[0][0] + parts[1][0]).upper() if len(parts) >= 2 else (parts[0][:2].upper() if parts else "??")
        palettes = [
            ("rgba(34,197,94,0.18)", "#15803d"),
            ("rgba(59,130,246,0.18)", "#1d4ed8"),
            ("rgba(245,158,11,0.18)", "#b45309"),
            ("rgba(168,85,247,0.18)", "#7c3aed"),
            ("rgba(236,72,153,0.18)", "#be185d"),
            ("rgba(20,184,166,0.18)", "#0f766e"),
        ]
        bg, fg = palettes[hash(str(name)) % len(palettes)]
        return f"<span class='util-emp-avatar' style='background:{bg};color:{fg}'>{initials}</span>"

    def short_name(n):
        parts = str(n).split(",", 1)
        if len(parts) == 2:
            last = parts[0].strip()
            first = parts[1].strip().split()[0] if parts[1].strip() else ""
            return f"{last}, {first[:1]}." if first else last
        return n

    # ─────────────────────────────────────────────────────
    # Action callouts (for At a glance tab)
    # ─────────────────────────────────────────────────────
    # Build employee summary first — needed by callouts and Consultants tab
    _ep = df[df["credit_tag"] != "SKIPPED"]
    emp_sum = _ep.groupby(["employee", "period"], as_index=False).agg(
        hours_this_period=("hours", "sum"),
        credit_hrs=("credit_hrs", "sum"),
        ff_overrun_hrs=("variance_hrs", "sum"),
    ).sort_values(["employee", "period"])
    _emp_region_ui = df.dropna(subset=["region"]).groupby("employee")["region"].first().to_dict() if "region" in df.columns else {}
    emp_sum["location"]  = emp_sum["employee"].map(_emp_region_ui)
    emp_sum["avail_hrs"] = emp_sum.apply(
        lambda r: get_avail_hours(r["location"], r["period"]) if r["location"] else None, axis=1)
    emp_sum["util_vs_logged"]   = emp_sum.apply(lambda r: r["credit_hrs"] / r["hours_this_period"] if r["hours_this_period"] > 0 else None, axis=1)
    emp_sum["util_vs_capacity"] = emp_sum.apply(lambda r: r["credit_hrs"] / r["avail_hrs"] if r["avail_hrs"] and r["avail_hrs"] > 0 else None, axis=1)
    emp_sum["exempt"] = emp_sum["employee"].apply(_emp_util_exempt)

    # Below-60% consultants
    _below_60 = emp_sum[(~emp_sum["exempt"]) & emp_sum["util_vs_capacity"].notna() & (emp_sum["util_vs_capacity"] < 0.60)]
    below60_count = len(_below_60)
    below60_names = ", ".join(sorted([short_name(n) for n in _below_60["employee"].unique()]))

    # Projects in overrun
    # credit_tag = OVERRUN is only assigned to FF projects by assign_credits, so this is the
    # canonical signal for "FF in overrun" — no project_type string parsing needed.
    _proj_overrun = pd.DataFrame()
    if "variance_hrs" in df.columns and _pid_col and "credit_tag" in df.columns:
        _overrun_pids = df_bill[df_bill["credit_tag"] == "OVERRUN"][_pid_col].astype(str).str.strip().unique()
        _po_grp = df_bill[df_bill[_pid_col].astype(str).str.strip().isin(_overrun_pids)].groupby([_pid_col], as_index=False).agg(
            project=("project", "first") if "project" in df.columns else (_pid_col, "first"),
            project_type=("project_type", "first") if "project_type" in df.columns else (_pid_col, "first"),
            hours_logged=("hours", "sum"),
            credit_hrs=("credit_hrs", "sum"),
            overrun_hrs=("variance_hrs", "sum"),
        )
        _proj_overrun = _po_grp[_po_grp["overrun_hrs"] > 0].copy()
    overrun_count = len(_proj_overrun)
    overrun_total = _proj_overrun["overrun_hrs"].sum() if not _proj_overrun.empty else 0
    overrun_top = ""
    if not _proj_overrun.empty:
        _top = _proj_overrun.sort_values("overrun_hrs", ascending=False).iloc[0]
        overrun_top = f"Largest: <b>{_top.get('project', '')}</b> (+{_top['overrun_hrs']:.1f}h)"

    # No-scope FF
    _noscope = df[df["credit_tag"] == "UNCONFIGURED"] if "credit_tag" in df.columns else df.iloc[0:0]
    noscope_hrs = _noscope["hours"].sum() if not _noscope.empty else 0
    noscope_proj_count = _noscope[_pid_col].astype(str).str.strip().nunique() if _pid_col and not _noscope.empty else 0

    # ─────────────────────────────────────────────────────
    # Wider-window data prep (used by Trend + Task tabs)
    # ─────────────────────────────────────────────────────
    _trend_df_raw = _ns_from_session.copy()
    if "employee" in _trend_df_raw.columns:
        from shared.constants import name_matches as _nm
        if _va_region_u:
            _rc = get_region_consultants(_va_region_u, _EL_u, _RM_u, _RO_u, _AE_u)
            _filt = _trend_df_raw[_trend_df_raw["employee"].astype(str).str.strip().str.lower().isin(_rc)]
            if not _filt.empty: _trend_df_raw = _filt
        elif _va_name_u:
            _filt = _trend_df_raw[_trend_df_raw["employee"].apply(lambda v: _nm(v, _va_name_u))]
            if not _filt.empty: _trend_df_raw = _filt
        elif not _is_mgr_u and _logged_in:
            _filt = _trend_df_raw[_trend_df_raw["employee"].apply(lambda v: _nm(v, _logged_in))]
            if not _filt.empty: _trend_df_raw = _filt

    _trend_cache_key = ("trend", _ns_signature(_ns_from_session), str(period_start), str(period_end), _va_name_u, _va_region_u)
    if _trend_cache_key in st.session_state._util_cache:
        _trend_result = st.session_state._util_cache[_trend_cache_key]
    else:
        try:
            _trend_result = _run_utilization_engine(_trend_df_raw, period_start, period_end)
            st.session_state._util_cache[_trend_cache_key] = _trend_result
        except Exception:
            _trend_result = {"empty": True}

    # ─────────────────────────────────────────────────────
    # Tabs
    # ─────────────────────────────────────────────────────
    overrun_badge = f" · <span style='color:#b91c1c'>{overrun_count}</span>" if overrun_count > 0 else ""
    tab_at_glance, tab_consult, tab_risk, tab_trend, tab_task, tab_detail = st.tabs([
        "At a glance",
        f"Consultants · {consultant_count}",
        f"Projects at risk{' · ' + str(overrun_count + noscope_proj_count) if (overrun_count + noscope_proj_count) > 0 else ''}",
        "Trend",
        "Task analysis",
        f"Detail · {len(df):,}",
    ])

    # ═══════════════════════════════════════════════════════════════════
    # TAB 1 — At a glance
    # ═══════════════════════════════════════════════════════════════════
    with tab_at_glance:
        # KPI strip
        m1, m2, m3, m4, m5 = st.columns(5)
        def _kpi_card(label, value, sub=None, sub_pill_class=None):
            sub_html = ""
            if sub:
                if sub_pill_class:
                    sub_html = f"<div style='display:inline-block;margin-top:4px;padding:2px 9px;border-radius:999px;font-size:11px' class='util-pill {sub_pill_class}'>{sub}</div>"
                else:
                    sub_html = f"<div style='font-size:11px;opacity:0.65;margin-top:4px'>{sub}</div>"
            return (f"<div class='util-kpi'>"
                    f"<div style='font-size:11px;opacity:0.75;margin-bottom:4px;"
                    f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>{label}</div>"
                    f"<div style='font-size:24px;font-weight:600;line-height:1.1;"
                    f"font-variant-numeric:tabular-nums'>{value}</div>"
                    f"{sub_html}</div>")

        with m1: st.markdown(_kpi_card("Billable projects", f"{billable_proj_count:,}", f"across {consultant_count} consultants"), unsafe_allow_html=True)
        with m2: st.markdown(_kpi_card("Hours logged", fmt_hrs(hours_this_period), f"of {_avail_total:,.0f} capacity ({capacity_pct:.1%})" if _avail_total else None), unsafe_allow_html=True)

        _credit_cls = "util-pill-green" if credit_pct >= 0.70 else "util-pill-amber" if credit_pct >= 0.60 else "util-pill-red"
        with m3: st.markdown(_kpi_card("Util credits", fmt_hrs(total_credit), f"{credit_pct:.1%} · {credit_label}", _credit_cls), unsafe_allow_html=True)
        with m4: st.markdown(_kpi_card("FF overrun", fmt_hrs(total_proj_overrun), f"{overrun_pct:.1%} of hrs", "util-pill-amber" if total_proj_overrun > 0 else None), unsafe_allow_html=True)
        with m5: st.markdown(_kpi_card("Admin hrs", fmt_hrs(total_admin), f"{admin_pct:.1%} of hrs"), unsafe_allow_html=True)

        # Where to look — three callouts, always shown, green when clean
        st.markdown("<div style='margin-top:18px;font-size:11px;opacity:0.6;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:8px'>Where to look</div>", unsafe_allow_html=True)

        callout_cols = st.columns(3)

        def _callout(border_color, title_color, icon, title, num, body):
            return (f"<div class='util-card' style='border-left:3px solid {border_color}'>"
                    f"<div style='display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px'>"
                    f"<span style='font-size:12px;font-weight:600;color:{title_color}'>{icon} {title}</span>"
                    f"<span style='font-size:22px;font-weight:600;color:{title_color};font-variant-numeric:tabular-nums'>{num}</span></div>"
                    f"<div style='font-size:12px;opacity:0.75;line-height:1.4'>{body}</div>"
                    f"</div>")

        # Callout 1 — Below 60%
        if below60_count > 0:
            _names_short = ", ".join(sorted([short_name(n) for n in _below_60["employee"].unique()])[:5])
            _and_more = f" +{below60_count - 5} more" if below60_count > 5 else ""
            _c1 = _callout("#ef4444", "#b91c1c", "⚠", "Consultants below 60%", below60_count, f"{_names_short}{_and_more}")
        else:
            _c1 = _callout("#22c55e", "#15803d", "✓", "Consultants below 60%", 0, "All consultants at or above 60% utilization on capacity.")
        with callout_cols[0]: st.markdown(_c1, unsafe_allow_html=True)

        # Callout 2 — FF in overrun
        if overrun_count > 0:
            _c2 = _callout("#f59e0b", "#b45309", "↑", "FF projects in overrun", overrun_count, f"{overrun_total:.1f} hrs beyond contracted scope. {overrun_top}")
        else:
            _c2 = _callout("#22c55e", "#15803d", "✓", "FF projects in overrun", 0, "No Fixed Fee projects exceeded scope this period.")
        with callout_cols[1]: st.markdown(_c2, unsafe_allow_html=True)

        # Callout 3 — No scope
        if noscope_proj_count > 0:
            _c3 = _callout("#3b82f6", "#1d4ed8", "⊘", "FF: no scope defined", noscope_proj_count, f"{noscope_hrs:.1f} hrs on FF projects with no scope record. Finance follow-up needed.")
        else:
            _c3 = _callout("#22c55e", "#15803d", "✓", "FF: no scope defined", 0, "All FF time logged against scoped projects.")
        with callout_cols[2]: st.markdown(_c3, unsafe_allow_html=True)

        # Inline legend + popover
        st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)
        legend_col, ref_col = st.columns([8, 1])
        with legend_col:
            st.markdown(
                "<div class='util-legend'>"
                "<span class='util-legend-label'>Credit tags</span>"
                "<span class='util-pill util-pill-green'>● Credited</span>"
                "<span class='util-pill util-pill-amber'>● Partial</span>"
                "<span class='util-pill util-pill-red'>● Overrun</span>"
                "<span class='util-pill util-pill-grey'>● Non-billable</span>"
                "<span class='util-pill util-pill-blue'>● FF: no scope</span>"
                "<span class='util-legend-label' style='margin-left:12px'>RAG</span>"
                "<span class='util-pill util-pill-green'>● ≥70%</span>"
                "<span class='util-pill util-pill-amber'>● 60–69%</span>"
                "<span class='util-pill util-pill-red'>● &lt;60%</span>"
                "</div>", unsafe_allow_html=True)
        with ref_col:
            _popover_fn = getattr(st, "popover", None)
            _ref_ctx = _popover_fn("ⓘ How?", help="How are these calculated?") if _popover_fn else st.expander("ⓘ How?", expanded=False)
            with _ref_ctx:
                st.markdown("""
                <div class='util-ref-grid'>
                  <div class='util-ref-section'>
                    <h4>Credit tags</h4>
                    <div class='util-ref-row'>
                      <span class='util-ref-tag util-pill-green'>Credited</span>
                      <span class='util-ref-desc'>T&amp;M: full hours credited, no cap. Fixed Fee: hours credited up to scoped amount.</span>
                    </div>
                    <div class='util-ref-row'>
                      <span class='util-ref-tag util-pill-amber'>Partial</span>
                      <span class='util-ref-desc'>Fixed Fee: hours credited up to remaining scope only.</span>
                    </div>
                    <div class='util-ref-row'>
                      <span class='util-ref-tag util-pill-red'>Overrun</span>
                      <span class='util-ref-desc'>Fixed Fee: hours beyond contracted scope. Not credited.</span>
                    </div>
                    <div class='util-ref-row'>
                      <span class='util-ref-tag util-pill-grey'>Non-billable</span>
                      <span class='util-ref-desc'>Internal time. Excluded from utilization, tracked as Admin Hours.</span>
                    </div>
                    <div class='util-ref-row'>
                      <span class='util-ref-tag util-pill-blue'>FF: no scope</span>
                      <span class='util-ref-desc'>No matching scope entry. Flagged for finance follow-up.</span>
                    </div>
                  </div>
                  <div class='util-ref-section'>
                    <h4>RAG status</h4>
                    <div class='util-ref-row'>
                      <span class='util-ref-tag util-pill-green'>Green ≥ 70%</span>
                      <span class='util-ref-desc'>At or above target utilization.</span>
                    </div>
                    <div class='util-ref-row'>
                      <span class='util-ref-tag util-pill-amber'>Amber 60–69%</span>
                      <span class='util-ref-desc'>Below target. Monitor and assess project mix.</span>
                    </div>
                    <div class='util-ref-row'>
                      <span class='util-ref-tag util-pill-red'>Red &lt; 60%</span>
                      <span class='util-ref-desc'>Significantly below target. Action required.</span>
                    </div>
                  </div>
                </div>
                """, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # TAB 2 — Consultants
    # ═══════════════════════════════════════════════════════════════════
    with tab_consult:
        # Projection
        _period_bdays = {}
        if "date" in df.columns:
            for _per, _grp in df.groupby("period"):
                _dates = _grp["date"].dropna()
                if len(_dates) == 0: continue
                _days_in = len(pd.bdate_range(_dates.min(), _dates.max()))
                _yr3, _mo3 = _dates.min().year, _dates.min().month
                _ms3 = pd.Timestamp(_yr3, _mo3, 1)
                _me3 = pd.Timestamp(_yr3, _mo3, _cal.monthrange(_yr3, _mo3)[1])
                _total = len(pd.bdate_range(_ms3, _me3))
                _period_bdays[str(_per)] = (_days_in, _total)

        def _proj_util(row):
            per = str(row["period"])
            if per not in _period_bdays: return None
            days_in, total = _period_bdays[per]
            if days_in >= total: return None
            avail = row["avail_hrs"]
            if not avail or avail <= 0 or days_in <= 0: return None
            return (row['credit_hrs'] / days_in * total) / avail

        emp_sum["proj_full_month"] = emp_sum.apply(_proj_util, axis=1)
        _is_partial_emp = emp_sum["proj_full_month"].notna().any()

        _consult_sort_opts = {
            "Util % capacity (low → high)":  ("util_vs_capacity", True),
            "Util % capacity (high → low)":  ("util_vs_capacity", False),
            "Util % logged (low → high)":    ("util_vs_logged", True),
            "Util % logged (high → low)":    ("util_vs_logged", False),
            "Hours logged (high → low)":     ("hours_this_period", False),
            "FF overrun (high → low)":       ("ff_overrun_hrs", False),
            "Consultant (A → Z)":            ("employee", True),
        }
        cs_col, _ = st.columns([1.6, 4])
        with cs_col:
            consult_sort = st.selectbox("Sort by", list(_consult_sort_opts.keys()),
                                        index=1, key="util_consult_sort")
        _csk, _csa = _consult_sort_opts[consult_sort]
        emp_sum_sorted = emp_sum.sort_values(_csk, ascending=_csa, na_position="last").reset_index(drop=True)

        rows_html = []
        for _, r in emp_sum_sorted.iterrows():
            emp = r["employee"]; ex = bool(r["exempt"])
            avail_str = f"{r['avail_hrs']:,.1f}" if r["avail_hrs"] else "—"
            pj = r["proj_full_month"]
            pj_str = f"{pj*100:.1f}%" if pj is not None else "—"
            rows_html.append(
                f"<tr>"
                f"<td><span class='util-emp-name'>{avatar_html(emp)}{short_name(emp)}</span></td>"
                f"<td class='muted'>{r['location'] or '—'}</td>"
                f"<td class='muted'>{r['period']}</td>"
                f"<td class='num'>{avail_str}</td>"
                f"<td class='num'>{r['hours_this_period']:,.2f}</td>"
                f"<td class='num'>{r['credit_hrs']:,.2f}</td>"
                f"<td class='num'>{r['ff_overrun_hrs']:,.2f}</td>"
                f"<td class='center'>{rag_pill_html(r['util_vs_logged'], ex)}</td>"
                f"<td class='center'>{rag_pill_html(r['util_vs_capacity'], ex)}</td>"
                + (f"<td class='num muted'>{pj_str}</td>" if _is_partial_emp else "")
                + "</tr>"
            )

        proj_th = "<th class='num'>Proj full mo</th>" if _is_partial_emp else ""
        st.markdown(
            f"<div class='util-table-header'>"
            f"<span style='font-weight:600'>By consultant{' · projected to full month' if _is_partial_emp else ''}</span>"
            f"<span style='opacity:0.7'>{len(emp_sum_sorted)} of {len(emp_sum_sorted)} · sorted by {consult_sort.lower()}</span>"
            f"</div>"
            f"<div class='util-table-wrap'>"
            f"<table class='util-emp-table'>"
            f"<thead><tr>"
            f"<th>Consultant</th><th>Location</th><th>Period</th>"
            f"<th class='num'>Avail</th><th class='num'>Logged</th>"
            f"<th class='num'>Credits</th><th class='num'>Overrun</th>"
            f"<th class='center'>Util % logged</th><th class='center'>Util % cap</th>"
            f"{proj_th}"
            f"</tr></thead>"
            f"<tbody>{''.join(rows_html)}</tbody>"
            f"</table>"
            f"<div class='util-table-foot'>"
            f"<span>{'Partial period: ' + str(_bdays_in) + ' of ' + str(_bdays_total_month) + ' days. Projected at current daily run rate.' if is_partial else '&nbsp;'}</span>"
            f"<span>Last refresh: {_ago_str} · NetSuite source · {len(df):,} entries</span>"
            f"</div></div>",
            unsafe_allow_html=True
        )

    # ═══════════════════════════════════════════════════════════════════
    # TAB 3 — Projects at risk
    # ═══════════════════════════════════════════════════════════════════
    with tab_risk:
        # Build full project summary
        if _pid_col and len(df_bill) > 0:
            proj_sum = df_bill.groupby([_pid_col], as_index=False).agg(
                project=("project", "first") if "project" in df.columns else (_pid_col, "first"),
                project_type=("project_type", "first") if "project_type" in df.columns else (_pid_col, "first"),
                hours_logged=("hours", "sum"),
                credit_hrs=("credit_hrs", "sum"),
                overrun_hrs=("variance_hrs", "sum"),
            )

            # Consultants per project, ranked by hours logged (this period)
            if "employee" in df_bill.columns:
                _emp_by_proj = (df_bill.groupby([_pid_col, "employee"], as_index=False)["hours"].sum()
                                .sort_values([_pid_col, "hours"], ascending=[True, False]))
                _proj_consultants = _emp_by_proj.groupby(_pid_col)["employee"].apply(list).to_dict()
                proj_sum["consultants"] = proj_sum[_pid_col].map(_proj_consultants).apply(lambda v: v if isinstance(v, list) else [])
            else:
                proj_sum["consultants"] = [[] for _ in range(len(proj_sum))]

            # Resolve scope per project type via DEFAULT_SCOPE substring match
            def _scope_lookup(ptype):
                _matches = [(k, float(v)) for k, v in DEFAULT_SCOPE.items()
                            if k.strip().lower() in str(ptype).strip().lower()]
                return max(_matches, key=lambda x: len(x[0]))[1] if _matches else None
            proj_sum["scoped_hrs"] = proj_sum["project_type"].apply(_scope_lookup)

            # HTD per project from the consumed dict (keyed by project_id, falling back to project name)
            def _htd(r):
                pid = str(r[_pid_col]).strip().replace(".0", "")
                if pid and pid in consumed:
                    return consumed[pid]
                pname = str(r.get("project", "")).strip()
                return consumed.get(pname, None)
            proj_sum["htd_hrs"] = proj_sum.apply(_htd, axis=1)

            # Tag noscope
            _noscope_pids = set(_noscope[_pid_col].astype(str).str.strip().unique()) if _pid_col and not _noscope.empty else set()
            proj_sum["is_noscope"] = proj_sum[_pid_col].astype(str).str.strip().isin(_noscope_pids)

            # Burn % uses HTD vs scope (shows true project health, not just this period)
            def _burn(r):
                sc = r.get("scoped_hrs")
                htd = r.get("htd_hrs")
                if sc is None or sc <= 0 or htd is None: return None
                return htd / sc
            proj_sum["burn_pct"] = proj_sum.apply(_burn, axis=1)

            # Status pill
            def _status(r):
                if r["is_noscope"]:                                                         return ("blue",  "No scope")
                if r["overrun_hrs"] > 0:                                                    return ("red",   "Overrun")
                if r.get("burn_pct") is not None and r["burn_pct"] >= 0.80:                 return ("amber", f"Burn {r['burn_pct']*100:.0f}%")
                return None

            proj_sum["_status"] = proj_sum.apply(_status, axis=1)
            at_risk = proj_sum[proj_sum["_status"].notna()].copy()

            if len(at_risk) == 0:
                st.markdown(
                    "<div class='util-card' style='text-align:center;padding:32px'>"
                    "<div style='font-size:32px;margin-bottom:8px'>✓</div>"
                    "<div style='font-weight:600;margin-bottom:4px'>No projects at risk</div>"
                    "<div style='opacity:0.7;font-size:13px'>No FF overruns, no scope burn over 80%, no missing scope records.</div>"
                    "</div>", unsafe_allow_html=True)
            else:
                # Sort dropdown
                _sort_options = {
                    "Status, then overrun":  ("_status_rank", "overrun_hrs", "burn_pct"),
                    "Overrun (high → low)":  ("overrun_hrs", "burn_pct", None),
                    "Burn % (high → low)":   ("burn_pct", "overrun_hrs", None),
                    "HTD (high → low)":      ("htd_hrs", "overrun_hrs", None),
                    "Logged this period":    ("hours_logged", "overrun_hrs", None),
                    "Project (A → Z)":       ("project_asc", None, None),
                }
                sc1, _ = st.columns([1.6, 4])
                with sc1:
                    sort_choice = st.selectbox("Sort by", list(_sort_options.keys()),
                                               key="util_risk_sort", label_visibility="visible")

                # Apply sort
                at_risk["_status_rank"] = at_risk["_status"].apply(
                    lambda s: 1 if s[0] == "red" else 2 if s[0] == "amber" else 3)
                if sort_choice == "Project (A → Z)":
                    at_risk = at_risk.sort_values("project", ascending=True, na_position="last")
                else:
                    keys = [k for k in _sort_options[sort_choice] if k]
                    ascending = [True if k == "_status_rank" else False for k in keys]
                    at_risk = at_risk.sort_values(keys, ascending=ascending, na_position="last")
                at_risk = at_risk.reset_index(drop=True)

                rows = []
                for _, r in at_risk.iterrows():
                    _scls, _sl = r["_status"]
                    _sc_str = f"{r['scoped_hrs']:,.2f}" if pd.notna(r.get("scoped_hrs")) else "—"
                    _htd_str = f"{r['htd_hrs']:,.2f}" if pd.notna(r.get("htd_hrs")) else "—"
                    _burn_str = f"{r['burn_pct']*100:.0f}%" if pd.notna(r.get("burn_pct")) else "—"
                    _burn_bar = ""
                    if pd.notna(r.get("burn_pct")):
                        _bp = min(r["burn_pct"], 1.5)
                        _bcol = "#ef4444" if _bp > 1.0 else "#f59e0b" if _bp >= 0.8 else "#22c55e"
                        _burn_bar = (f"<div style='display:inline-block;width:60px;height:6px;background:rgba(128,128,128,0.2);"
                                     f"border-radius:3px;overflow:hidden;vertical-align:middle;margin-right:6px'>"
                                     f"<div style='width:{min(_bp*100/1.5, 100)}%;height:100%;background:{_bcol}'></div></div>")

                    # Consultants cell: solo = avatar+name, multi = stacked avatars + count
                    _consultants = r.get("consultants") or []
                    if len(_consultants) == 0:
                        _consult_html = "<span style='opacity:0.4'>—</span>"
                    elif len(_consultants) == 1:
                        _consult_html = f"<span class='util-emp-name'>{avatar_html(_consultants[0])}{short_name(_consultants[0])}</span>"
                    else:
                        # Stack first 3 avatars overlapping, then show count
                        _shown = _consultants[:3]
                        _avatars_html = "".join([
                            f"<span style='display:inline-block;margin-right:-8px' title='{c}'>{avatar_html(c)}</span>"
                            for c in _shown
                        ])
                        _all_names = ", ".join([short_name(c) for c in _consultants])
                        _consult_html = (f"<span class='util-emp-name' title='{_all_names}' style='gap:0'>"
                                         f"{_avatars_html}"
                                         f"<span style='margin-left:14px;font-size:12px;opacity:0.85;white-space:nowrap'>{len(_consultants)} consultants</span>"
                                         f"</span>")

                    rows.append(
                        f"<tr>"
                        f"<td>{r.get('project', '')}</td>"
                        f"<td class='muted'>{r.get('project_type', '')}</td>"
                        f"<td>{_consult_html}</td>"
                        f"<td class='num'>{_sc_str}</td>"
                        f"<td class='num'>{_htd_str}</td>"
                        f"<td class='num'>{r['hours_logged']:,.2f}</td>"
                        f"<td class='num'>{_burn_bar}<span style='vertical-align:middle'>{_burn_str}</span></td>"
                        f"<td class='num'>{r['overrun_hrs']:,.2f}</td>"
                        f"<td class='center'><span class='util-pill util-pill-{_scls}'>{_sl}</span></td>"
                        f"</tr>")

                st.markdown(
                    f"<div class='util-table-header'>"
                    f"<span style='font-weight:600'>Projects requiring attention</span>"
                    f"<span style='opacity:0.7'>{len(at_risk)} of {len(proj_sum)} billable projects · sorted by {sort_choice.lower()}</span>"
                    f"</div>"
                    f"<div class='util-table-wrap'>"
                    f"<table class='util-emp-table'>"
                    f"<thead><tr>"
                    f"<th>Project</th><th>Type</th><th>Consultant</th><th class='num'>Scoped</th>"
                    f"<th class='num'>HTD</th><th class='num'>Logged</th>"
                    f"<th class='num'>Burn</th><th class='num'>Overrun</th>"
                    f"<th class='center'>Status</th>"
                    f"</tr></thead><tbody>{''.join(rows)}</tbody></table>"
                    f"<div class='util-table-foot'>"
                    f"<span>Risk = overrun &gt; 0, burn ≥ 80% (HTD ÷ scope), or no scope record. HTD = hours-to-date all time.</span>"
                    f"<span>Last refresh: {_ago_str}</span>"
                    f"</div></div>",
                    unsafe_allow_html=True)
        else:
            st.info("No billable project data in this period.")

    # ═══════════════════════════════════════════════════════════════════
    # TAB 4 — Trend (matches the Period filter)
    # ═══════════════════════════════════════════════════════════════════
    with tab_trend:
        if _trend_result.get("empty") or "date" not in _trend_result.get("df", pd.DataFrame()).columns:
            st.info("Not enough data to render a trend for this period.")
        else:
            _td = _trend_result["df"].copy()
            _td["_date"] = pd.to_datetime(_td["date"], errors="coerce")
            _td = _td.dropna(subset=["_date"])
            _td["_week"] = _td["_date"].dt.to_period("W").apply(lambda p: p.start_time.date())

            # Weekly aggregates
            weekly = _td.groupby("_week", as_index=False).agg(
                hours=("hours", "sum"),
                credit_hrs=("credit_hrs", "sum"),
                overrun_hrs=("variance_hrs", "sum"),
            )
            # Billable project count per week (distinct project IDs, billable rows only)
            _td_pid = _proj_id_col(_td)
            if _td_pid:
                _bmask_td = _billable_mask(_td)
                _proj_weekly = _td[_bmask_td].groupby("_week", as_index=False).agg(
                    n_projects=(_td_pid, lambda s: s.astype(str).str.strip().nunique())
                )
                weekly = weekly.merge(_proj_weekly, on="_week", how="left")
                weekly["n_projects"] = weekly["n_projects"].fillna(0).astype(int)
            else:
                weekly["n_projects"] = 0
            # noscope hrs computed separately and merged in
            if "credit_tag" in _td.columns:
                _ns_weekly = _td[_td["credit_tag"] == "UNCONFIGURED"].groupby("_week", as_index=False).agg(noscope_hrs=("hours", "sum"))
                weekly = weekly.merge(_ns_weekly, on="_week", how="left")
                weekly["noscope_hrs"] = weekly["noscope_hrs"].fillna(0)
            else:
                weekly["noscope_hrs"] = 0
            weekly["credit_pct"] = weekly.apply(lambda r: r["credit_hrs"] / r["hours"] if r["hours"] > 0 else 0, axis=1)
            weekly = weekly.sort_values("_week").reset_index(drop=True)

            if len(weekly) < 2:
                st.info("Need at least 2 weeks of data to show a trend. Widen the period filter.")
            else:
                # Three side-by-side trend cards
                t1, t2, t3 = st.columns(3)

                def _mini_chart(values, target=None, fmt="pct", color="#3b82f6"):
                    """Render an inline SVG bar chart for trend cards."""
                    if not values: return ""
                    n = len(values)
                    max_v = max([v for v in values if v is not None] + [0.01])
                    bar_w = max(40 / n, 4)
                    gap = bar_w * 0.2
                    chart_w = 100
                    chart_h = 40
                    bar_w_actual = (chart_w - gap * (n - 1)) / n
                    bars = []
                    for i, v in enumerate(values):
                        if v is None or max_v == 0:
                            h = 0
                        else:
                            h = (v / max_v) * chart_h
                        x = i * (bar_w_actual + gap)
                        y = chart_h - h
                        bars.append(f"<rect x='{x:.1f}' y='{y:.1f}' width='{bar_w_actual:.1f}' height='{h:.1f}' fill='{color}' opacity='0.8' rx='1' />")
                    target_line = ""
                    if target is not None and max_v > 0:
                        ty = chart_h - (target / max_v) * chart_h
                        if 0 <= ty <= chart_h:
                            target_line = f"<line x1='0' y1='{ty:.1f}' x2='{chart_w}' y2='{ty:.1f}' stroke='rgba(128,128,128,0.5)' stroke-dasharray='2 2' />"
                    return f"<svg viewBox='0 0 {chart_w} {chart_h}' style='width:100%;height:50px;display:block'>{target_line}{''.join(bars)}</svg>"

                _credit_pcts  = weekly["credit_pct"].tolist()
                _overrun_hrs  = weekly["overrun_hrs"].tolist()
                _noscope_hrs_l= weekly["noscope_hrs"].tolist()

                # ─── Period totals (this period) ───
                _curr_total_hrs        = weekly["hours"].sum()
                _curr_total_credits    = weekly["credit_hrs"].sum()
                _curr_credit_pct       = _curr_total_credits / _curr_total_hrs if _curr_total_hrs > 0 else 0
                _curr_total_overrun    = sum(_overrun_hrs)
                _curr_total_noscope    = sum(_noscope_hrs_l)

                # ─── Prior equivalent period totals ───
                _period_days = (period_end - period_start).days + 1
                _prior_p_end = (pd.Timestamp(period_start) - pd.Timedelta(days=1)).date()
                _prior_p_start = (pd.Timestamp(_prior_p_end) - pd.Timedelta(days=_period_days - 1)).date()

                _prior_p_cache_key = ("trend_prior", _ns_signature(_ns_from_session), str(_prior_p_start), str(_prior_p_end), _va_name_u, _va_region_u)
                if _prior_p_cache_key in st.session_state._util_cache:
                    _prior_p_result = st.session_state._util_cache[_prior_p_cache_key]
                else:
                    try:
                        _prior_p_result = _run_utilization_engine(_trend_df_raw, _prior_p_start, _prior_p_end)
                        st.session_state._util_cache[_prior_p_cache_key] = _prior_p_result
                    except Exception:
                        _prior_p_result = {"empty": True}

                _prior_credit_pct = None
                _prior_total_overrun = None
                _prior_total_noscope = None
                if not _prior_p_result.get("empty") and not _prior_p_result.get("df", pd.DataFrame()).empty:
                    _pdf2 = _prior_p_result["df"]
                    _p_hrs    = _pdf2["hours"].sum() if "hours" in _pdf2.columns else 0
                    _p_credit = _pdf2["credit_hrs"].sum() if "credit_hrs" in _pdf2.columns else 0
                    _prior_credit_pct = _p_credit / _p_hrs if _p_hrs > 0 else 0
                    _prior_total_overrun = _pdf2[_pdf2["credit_tag"].isin(["OVERRUN", "PARTIAL"])]["variance_hrs"].sum() if "variance_hrs" in _pdf2.columns and "credit_tag" in _pdf2.columns else 0
                    _prior_total_noscope = _pdf2[_pdf2["credit_tag"] == "UNCONFIGURED"]["hours"].sum() if "credit_tag" in _pdf2.columns else 0

                # ─── Deltas (current − prior) ───
                _delta_credit  = (_curr_credit_pct - _prior_credit_pct) if _prior_credit_pct is not None else None
                _delta_overrun = (_curr_total_overrun - _prior_total_overrun) if _prior_total_overrun is not None else None
                _delta_noscope = (_curr_total_noscope - _prior_total_noscope) if _prior_total_noscope is not None else None

                _prior_label = f"vs {_prior_p_start.strftime('%-d %b')}–{_prior_p_end.strftime('%-d %b')}"

                # Headline color (use period total for credit, since that's what RAG measures)
                _headline_color = "#22c55e" if _curr_credit_pct >= 0.70 else "#f59e0b" if _curr_credit_pct >= 0.60 else "#ef4444"

                def _delta_html(delta, unit, lower_is_better=False, threshold=0.005):
                    """Render arrow + magnitude. Higher = better unless lower_is_better."""
                    if delta is None:
                        return f"<div style='font-size:12px;opacity:0.6'>no prior data</div>"
                    if abs(delta) < threshold:
                        return f"<div style='font-size:12px;opacity:0.6'>→ flat</div>"
                    is_up = delta > 0
                    is_good = (is_up and not lower_is_better) or (not is_up and lower_is_better)
                    color = "#15803d" if is_good else "#b91c1c"
                    arrow = "↑" if is_up else "↓"
                    return f"<div style='font-size:12px;color:{color}'>{arrow} {abs(delta):.1f}{unit}</div>"

                with t1:
                    _curr_credit_pp = _curr_credit_pct * 100
                    _delta_credit_pp = _delta_credit * 100 if _delta_credit is not None else None
                    st.markdown(
                        f"<div class='util-card'>"
                        f"<div style='font-size:11px;opacity:0.65;margin-bottom:4px'>Credit % · period total</div>"
                        f"<div style='display:flex;align-items:baseline;gap:8px;margin-bottom:8px'>"
                        f"<div style='font-size:22px;font-weight:600;color:{_headline_color}'>{_curr_credit_pp:.1f}%</div>"
                        f"{_delta_html(_delta_credit_pp, 'pp', lower_is_better=False, threshold=0.5)}"
                        f"</div>"
                        f"{_mini_chart(_credit_pcts, target=0.70, color='#3b82f6')}"
                        f"<div style='font-size:11px;opacity:0.6;margin-top:4px'>{len(weekly)} weeks · target ≥ 70% · {_prior_label}</div>"
                        f"</div>", unsafe_allow_html=True)

                with t2:
                    st.markdown(
                        f"<div class='util-card'>"
                        f"<div style='font-size:11px;opacity:0.65;margin-bottom:4px'>FF overrun · period total</div>"
                        f"<div style='display:flex;align-items:baseline;gap:8px;margin-bottom:8px'>"
                        f"<div style='font-size:22px;font-weight:600'>{fmt_hrs(_curr_total_overrun)} h</div>"
                        f"{_delta_html(_delta_overrun, 'h', lower_is_better=True, threshold=0.5)}"
                        f"</div>"
                        f"{_mini_chart(_overrun_hrs, color='#f59e0b')}"
                        f"<div style='font-size:11px;opacity:0.6;margin-top:4px'>{len(weekly)} weeks · {_prior_label}</div>"
                        f"</div>", unsafe_allow_html=True)

                with t3:
                    st.markdown(
                        f"<div class='util-card'>"
                        f"<div style='font-size:11px;opacity:0.65;margin-bottom:4px'>FF: no scope · period total</div>"
                        f"<div style='display:flex;align-items:baseline;gap:8px;margin-bottom:8px'>"
                        f"<div style='font-size:22px;font-weight:600'>{fmt_hrs(_curr_total_noscope)} h</div>"
                        f"{_delta_html(_delta_noscope, 'h', lower_is_better=True, threshold=0.5)}"
                        f"</div>"
                        f"{_mini_chart(_noscope_hrs_l, color='#3b82f6')}"
                        f"<div style='font-size:11px;opacity:0.6;margin-top:4px'>{len(weekly)} weeks · {_prior_label}</div>"
                        f"</div>", unsafe_allow_html=True)

                # Weekly detail table
                st.markdown("<div style='margin-top:14px;font-size:11px;opacity:0.6;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:6px'>Weekly breakdown</div>", unsafe_allow_html=True)
                _weekly_disp = weekly.copy()
                _weekly_disp["Week"] = _weekly_disp["_week"].apply(lambda d: d.strftime("%Y-W%V") if hasattr(d, "strftime") else str(d))
                _weekly_disp["credit_pct"] = _weekly_disp["credit_pct"] * 100  # Show as percentage points
                _weekly_disp = _weekly_disp[["Week", "n_projects", "hours", "credit_hrs", "credit_pct", "overrun_hrs", "noscope_hrs"]]
                st.dataframe(
                    _weekly_disp, use_container_width=True, hide_index=True,
                    column_config={
                        "Week":         st.column_config.TextColumn("Week", pinned=True),
                        "n_projects":   st.column_config.NumberColumn("Projects", format="%d", help="Distinct billable projects with hours logged this week"),
                        "hours":        st.column_config.NumberColumn("Logged", format="%.2f"),
                        "credit_hrs":   st.column_config.NumberColumn("Credits", format="%.2f"),
                        "credit_pct":   st.column_config.ProgressColumn("Credit %", format="%.1f%%", min_value=0, max_value=100, help="Credits ÷ hours logged"),
                        "overrun_hrs":  st.column_config.NumberColumn("FF overrun", format="%.2f"),
                        "noscope_hrs":  st.column_config.NumberColumn("No scope", format="%.2f"),
                    },
                )

    # ═══════════════════════════════════════════════════════════════════
    # TAB 5 — Task analysis (Billable / Non-billable toggle)
    # ═══════════════════════════════════════════════════════════════════
    with tab_task:
        # Mode selector
        _bill_hrs_total = df_bill["hours"].sum() if not df_bill.empty else 0
        _nonbill_hrs_total = df[df["billing_type"].fillna("").str.lower() == "internal"]["hours"].sum() if "billing_type" in df.columns else 0

        ts_col1, ts_col2, ts_col3 = st.columns([1.2, 1.2, 2.5])
        with ts_col1:
            task_mode = st.radio(
                "Mode",
                options=["Billable", "Non-billable"],
                horizontal=True, label_visibility="collapsed",
                key="util_task_mode",
                format_func=lambda x: f"Billable · {fmt_hrs(_bill_hrs_total)}h" if x == "Billable" else f"Non-billable · {fmt_hrs(_nonbill_hrs_total)}h",
            )
        with ts_col2:
            compare_options = {
                "Prior equivalent period": "prior_eq",
                "Prior 4 weeks (per-week avg)": "prior_4w",
                "Prior 8 weeks (per-week avg)": "prior_8w",
            }
            compare_label = st.selectbox("Compare vs", list(compare_options.keys()), key="util_task_compare", label_visibility="collapsed")
            compare_mode = compare_options[compare_label]
        with ts_col3:
            _norm_note = " · normalized to /week" if compare_mode != "prior_eq" else ""
            st.markdown(f"<div style='padding:8px 0;font-size:11px;opacity:0.6;text-align:right'>vs <b>{compare_label}</b>{_norm_note}</div>", unsafe_allow_html=True)

        # Determine task column
        if task_mode == "Billable":
            _task_df = df_bill.copy()
            _task_col = "ff_task" if "ff_task" in _task_df.columns and _task_df["ff_task"].notna().any() else ("task" if "task" in _task_df.columns else None)
        else:
            _task_df = df[df["billing_type"].fillna("").str.lower() == "internal"].copy() if "billing_type" in df.columns else df.iloc[0:0].copy()
            _task_col = "task" if "task" in _task_df.columns else None

        if not _task_col or len(_task_df) == 0:
            st.info(f"No {task_mode.lower()} time entries in this period.")
        else:
            # ─── Top-left: share by task ───
            share = _task_df.groupby(_task_col, as_index=False).agg(hours=("hours", "sum"))
            share = share[share["hours"] > 0].sort_values("hours", ascending=False).reset_index(drop=True)
            total_share = share["hours"].sum()
            # Collapse < 5% into "Other"
            _big = share[share["hours"] / total_share >= 0.05].copy() if total_share > 0 else share.copy()
            _small = share[share["hours"] / total_share < 0.05].copy() if total_share > 0 else share.iloc[0:0]
            if len(_small) > 0:
                _other_row = pd.DataFrame([{_task_col: f"Other ({len(_small)} tasks)", "hours": _small["hours"].sum()}])
                _big = pd.concat([_big, _other_row], ignore_index=True)
            _big["pct"] = _big["hours"] / total_share if total_share > 0 else 0

            share_rows = []
            for i, r in _big.iterrows():
                _opacity = max(0.4, 1.0 - (i * 0.1))
                _muted = " style='opacity:0.65'" if str(r[_task_col]).startswith("Other") else ""
                share_rows.append(
                    f"<div class='util-share-row'>"
                    f"<div class='util-share-head'{_muted}>"
                    f"<span>{r[_task_col]}</span>"
                    f"<span style='opacity:0.7;font-variant-numeric:tabular-nums'>{r['hours']:,.2f} h · {r['pct']*100:.1f}%</span>"
                    f"</div>"
                    f"<div class='util-share-bar-bg'>"
                    f"<div class='util-share-bar-fg' style='width:{min(r['pct']*100, 100):.1f}%;opacity:{_opacity}'></div>"
                    f"</div></div>")

            # ─── Top-right: movers vs prior period ───
            if compare_mode == "prior_eq":
                _delta = (period_end - period_start).days + 1
                _prior_end = (pd.Timestamp(period_start) - pd.Timedelta(days=1)).date()
                _prior_start = (pd.Timestamp(_prior_end) - pd.Timedelta(days=_delta - 1)).date()
            elif compare_mode == "prior_4w":
                _prior_end = (pd.Timestamp(period_start) - pd.Timedelta(days=1)).date()
                _prior_start = (pd.Timestamp(_prior_end) - pd.Timedelta(weeks=4)).date()
            else:  # prior_8w
                _prior_end = (pd.Timestamp(period_start) - pd.Timedelta(days=1)).date()
                _prior_start = (pd.Timestamp(_prior_end) - pd.Timedelta(weeks=8)).date()

            _curr_weeks  = max(((period_end - period_start).days + 1) / 7, 1/7)
            _prior_weeks = max(((_prior_end - _prior_start).days + 1) / 7, 1/7)

            # Use the wider df (df_raw before period filter) for comparison
            _prior_cache_key = ("prior_task", _ns_signature(_ns_from_session), str(_prior_start), str(_prior_end), _va_name_u, _va_region_u)
            if _prior_cache_key in st.session_state._util_cache:
                _prior_result = st.session_state._util_cache[_prior_cache_key]
            else:
                try:
                    _prior_result = _run_utilization_engine(_trend_df_raw, _prior_start, _prior_end)
                    st.session_state._util_cache[_prior_cache_key] = _prior_result
                except Exception:
                    _prior_result = {"empty": True}

            movers_html = ""
            if _prior_result.get("empty") or _prior_result.get("df", pd.DataFrame()).empty:
                movers_html = "<div style='opacity:0.6;font-size:12px;padding:20px 0;text-align:center'>No data for the comparison window.</div>"
            else:
                _pdf = _prior_result["df"]
                if task_mode == "Billable":
                    _pdf_filt = _pdf[_billable_mask(_pdf)]
                    _pcol = "ff_task" if "ff_task" in _pdf_filt.columns and _pdf_filt["ff_task"].notna().any() else ("task" if "task" in _pdf_filt.columns else None)
                else:
                    _pdf_filt = _pdf[_pdf["billing_type"].fillna("").str.lower() == "internal"] if "billing_type" in _pdf.columns else _pdf.iloc[0:0]
                    _pcol = "task" if "task" in _pdf_filt.columns else None

                if _pcol and len(_pdf_filt) > 0:
                    prior_share = _pdf_filt.groupby(_pcol, as_index=False).agg(prior_hours=("hours", "sum"))
                    merged = share.merge(prior_share, left_on=_task_col, right_on=_pcol, how="outer").fillna(0)
                    if _task_col != _pcol and _pcol in merged.columns:
                        merged[_task_col] = merged[_task_col].where(merged[_task_col] != 0, merged[_pcol])

                    # Per-week normalization for rolling options
                    _normalize = (compare_mode != "prior_eq")
                    if _normalize:
                        merged["curr_per_wk"]  = merged["hours"] / _curr_weeks
                        merged["prior_per_wk"] = merged["prior_hours"] / _prior_weeks
                        merged["delta"] = merged["curr_per_wk"] - merged["prior_per_wk"]
                        _curr_disp_col, _prior_disp_col = "curr_per_wk", "prior_per_wk"
                        _unit = "h/wk"
                    else:
                        merged["delta"] = merged["hours"] - merged["prior_hours"]
                        _curr_disp_col, _prior_disp_col = "hours", "prior_hours"
                        _unit = "h"

                    merged["abs_delta"] = merged["delta"].abs()
                    movers = merged.sort_values("abs_delta", ascending=False).head(5)

                    _rows = []
                    for _, r in movers.iterrows():
                        _d = r["delta"]
                        _cls = "util-pill-green" if _d > 0 else "util-pill-red" if _d < 0 else "util-pill-grey"
                        _sign = "+" if _d > 0 else ""
                        _rows.append(
                            f"<div class='util-mover-row'>"
                            f"<span>{r[_task_col]}</span>"
                            f"<span class='util-mover-vals'>"
                            f"<span class='util-mover-from'>{r[_prior_disp_col]:.2f} → {r[_curr_disp_col]:.2f} {_unit}</span>"
                            f"<span class='util-pill {_cls}'>{_sign}{_d:.2f}</span>"
                            f"</span></div>"
                        )

                    # Auto-interpretation
                    _interp = ""
                    if not movers.empty:
                        _top = movers.iloc[0]
                        _delta_unit = "h/wk" if _normalize else "h"
                        if _top["delta"] > 0:
                            _interp = f"{_top[_task_col]} growing fastest (+{_top['delta']:.1f} {_delta_unit})."
                        else:
                            _interp = f"{_top[_task_col]} declining most ({_top['delta']:.1f} {_delta_unit})."

                    movers_html = "".join(_rows) + (f"<div style='font-size:11px;opacity:0.6;margin-top:10px;padding-top:8px;border-top:1px solid rgba(128,128,128,0.15)'>{_interp}</div>" if _interp else "")
                else:
                    movers_html = "<div style='opacity:0.6;font-size:12px;padding:20px 0;text-align:center'>No comparable task data.</div>"

            # Render top row
            tr1, tr2 = st.columns([1.4, 1])
            with tr1:
                st.markdown(
                    f"<div class='util-card'>"
                    f"<div style='font-size:12px;font-weight:600;margin-bottom:4px'>{task_mode} hours by task · this period</div>"
                    f"<div style='font-size:11px;opacity:0.65;margin-bottom:12px'>{total_share:,.2f} hrs across {len(share)} task type{'s' if len(share) != 1 else ''}</div>"
                    f"{''.join(share_rows)}"
                    f"</div>", unsafe_allow_html=True)
            with tr2:
                if compare_mode != "prior_eq":
                    _subtitle = (f"Current ÷ {_curr_weeks:.1f}w  vs  prior ÷ {_prior_weeks:.1f}w "
                                 f"({_prior_start.strftime('%-d %b')} → {_prior_end.strftime('%-d %b')})")
                else:
                    _subtitle = f"{_prior_start.strftime('%-d %b')} → {_prior_end.strftime('%-d %b')}"
                st.markdown(
                    f"<div class='util-card'>"
                    f"<div style='font-size:12px;font-weight:600;margin-bottom:4px'>Movers · vs {compare_label.lower()}</div>"
                    f"<div style='font-size:11px;opacity:0.65;margin-bottom:8px'>{_subtitle}</div>"
                    f"{movers_html}"
                    f"</div>", unsafe_allow_html=True)

            # ─── Bottom: stacked weekly trend ───
            if not _trend_result.get("empty"):
                _td2 = _trend_result["df"].copy()
                if task_mode == "Billable":
                    _td2 = _td2[_billable_mask(_td2)]
                else:
                    _td2 = _td2[_td2["billing_type"].fillna("").str.lower() == "internal"] if "billing_type" in _td2.columns else _td2.iloc[0:0]
                _td2["_date"] = pd.to_datetime(_td2["date"], errors="coerce")
                _td2 = _td2.dropna(subset=["_date"])
                _td2["_week"] = _td2["_date"].dt.to_period("W").apply(lambda p: p.start_time.date())

                _tcol_trend = _task_col if _task_col in _td2.columns else None
                if _tcol_trend and len(_td2) > 0:
                    # Top 4 tasks by total hours; rest = Other
                    _totals = _td2.groupby(_tcol_trend)["hours"].sum().sort_values(ascending=False)
                    _top4 = _totals.head(4).index.tolist()
                    _td2["_task_grp"] = _td2[_tcol_trend].apply(lambda v: v if v in _top4 else "Other")
                    _stack = _td2.groupby(["_week", "_task_grp"], as_index=False).agg(hours=("hours", "sum"))
                    _pivot = _stack.pivot(index="_week", columns="_task_grp", values="hours").fillna(0).sort_index()
                    _categories = _top4 + (["Other"] if "Other" in _pivot.columns else [])
                    _categories = [c for c in _categories if c in _pivot.columns]

                    if not _pivot.empty and len(_categories) > 0:
                        _palette = ["#1d4ed8", "#3b82f6", "#60a5fa", "#93c5fd", "#bfdbfe"]
                        _max_total = _pivot[_categories].sum(axis=1).max()
                        _w = 700; _h = 200
                        _bar_area_w = _w - 50
                        _n = len(_pivot.index)
                        _bw = (_bar_area_w / _n) * 0.7
                        _gap = (_bar_area_w / _n) * 0.3

                        bars_svg = []
                        labels_svg = []
                        for i, (wk, row) in enumerate(_pivot.iterrows()):
                            x = 50 + i * (_bw + _gap)
                            _y_cursor = _h - 20
                            for j, cat in enumerate(_categories):
                                v = row.get(cat, 0)
                                if v <= 0: continue
                                bh = (v / _max_total) * (_h - 40) if _max_total > 0 else 0
                                _y_cursor -= bh
                                bars_svg.append(f"<rect x='{x:.1f}' y='{_y_cursor:.1f}' width='{_bw:.1f}' height='{bh:.1f}' fill='{_palette[j % len(_palette)]}' />")
                            wk_label = wk.strftime("%-d %b") if hasattr(wk, "strftime") else str(wk)
                            labels_svg.append(f"<text x='{x + _bw/2:.1f}' y='{_h - 5}' fill='currentColor' opacity='0.6' font-size='9' text-anchor='middle'>{wk_label}</text>")

                        # y-axis grid
                        grid = ""
                        if _max_total > 0:
                            for frac in [0.25, 0.5, 0.75, 1.0]:
                                gy = (_h - 20) - frac * (_h - 40)
                                gv = frac * _max_total
                                grid += f"<line x1='50' y1='{gy:.1f}' x2='{_w}' y2='{gy:.1f}' stroke='currentColor' opacity='0.1' stroke-dasharray='2 4' />"
                                grid += f"<text x='45' y='{gy + 3:.1f}' fill='currentColor' opacity='0.5' font-size='9' text-anchor='end'>{gv:.0f}h</text>"

                        legend_html = "".join([
                            f"<span style='display:inline-flex;align-items:center;gap:5px;margin-right:14px;font-size:11px'>"
                            f"<span style='display:inline-block;width:10px;height:10px;background:{_palette[j % len(_palette)]};border-radius:2px'></span>"
                            f"{cat}</span>"
                            for j, cat in enumerate(_categories)
                        ])

                        st.markdown(
                            f"<div class='util-card' style='margin-top:12px'>"
                            f"<div style='display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px'>"
                            f"<span style='font-size:12px;font-weight:600'>{task_mode} hours by task · weekly</span>"
                            f"<span style='font-size:11px;opacity:0.6'>stacked · top 4 + other</span>"
                            f"</div>"
                            f"<svg viewBox='0 0 {_w} {_h}' style='width:100%;height:200px'>{grid}{''.join(bars_svg)}{''.join(labels_svg)}</svg>"
                            f"<div style='margin-top:8px;color:currentColor'>{legend_html}</div>"
                            f"</div>", unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # TAB 6 — Detail
    # ═══════════════════════════════════════════════════════════════════
    with tab_detail:
        display_cols_t5 = ["employee","region","project","project_type","billing_type",
                            "hours_to_date","date","hours","credit_hrs","variance_hrs",
                            "previous_htd","credit_tag","notes"]
        existing_t5 = [c for c in display_cols_t5 if c in df.columns]

        _glyph_map = {"CREDITED":"🟢","OVERRUN":"🔴","PARTIAL":"🟡","NON-BILLABLE":"⚫","UNCONFIGURED":"⚪","SKIPPED":"·"}
        _detail_view = df[existing_t5].head(500).copy()
        if "credit_tag" in _detail_view.columns:
            _detail_view["credit_tag"] = _detail_view["credit_tag"].map(lambda t: f"{_glyph_map.get(t,'')} {t}" if pd.notna(t) else "")

        st.dataframe(
            _detail_view, use_container_width=True, hide_index=True,
            column_config={
                "employee":     st.column_config.TextColumn("Employee", pinned=True),
                "region":       st.column_config.TextColumn("Region"),
                "project":      st.column_config.TextColumn("Project"),
                "project_type": st.column_config.TextColumn("Type"),
                "billing_type": st.column_config.TextColumn("Billing"),
                "hours_to_date":st.column_config.NumberColumn("HTD", format="%.2f"),
                "date":         st.column_config.DateColumn("Date"),
                "hours":        st.column_config.NumberColumn("Hours", format="%.2f"),
                "credit_hrs":   st.column_config.NumberColumn("Credit hrs", format="%.2f"),
                "variance_hrs": st.column_config.NumberColumn("Variance", format="%.2f"),
                "previous_htd": st.column_config.NumberColumn("Prev HTD", format="%.2f"),
                "credit_tag":   st.column_config.TextColumn("Credit tag", help="🟢 Credited · 🟡 Partial · 🔴 Overrun · ⚫ Non-billable · ⚪ FF no scope"),
                "notes":        st.column_config.TextColumn("Notes"),
            },
        )
        if len(df) > 500:
            st.caption(f"Showing first 500 of {len(df):,} rows. Full data in Excel download.")

    # ─────────────────────────────────────────────────────
    # Tableau export — secondary
    # ─────────────────────────────────────────────────────
    st.divider()
    with st.expander("Tableau export (advanced)", expanded=False):
        with st.spinner("Building Tableau export..."):
            tableau_buf = build_tableau_excel(df, DEFAULT_SCOPE, consumed)
        tableau_filename = f"utilization_tableau_{timestamp}.xlsx"
        st.download_button(
            label="⬇ Download Tableau Export",
            data=tableau_buf, file_name=tableau_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="util_dl_tableau",
        )
        st.caption("3 flat sheets: fact_utilization · fact_processed_time_entries · fact_ff_overrun_by_type")
def build_tableau_excel(df, scope_map, consumed):
    import io as _io
    from openpyxl import Workbook as _WB

    wb = _WB()
    wb.remove(wb.active)

    def _flat_sheet(wb, name, headers, rows):
        ws = wb.create_sheet(name)
        ws.append(headers)
        for r in rows: ws.append(r)
        return ws

    def _pct(num, denom):
        try: return round(num / denom, 6) if denom and denom > 0 else None
        except Exception: return None

    _skipped = df["credit_tag"] == "SKIPPED" if "credit_tag" in df.columns else df.index == -1
    _df = df[~_skipped].copy()
    _admin = (_df[_df["billing_type"].str.lower() == "internal"]
              .groupby(["employee","period"])["hours"].sum()
              .reset_index().rename(columns={"hours":"admin_hrs"}))
    _sum = (_df.groupby(["employee","period"], as_index=False)
            .agg(hours_logged=("hours","sum"), credit_hrs=("credit_hrs","sum"), ff_overrun_hrs=("variance_hrs","sum"))
            .sort_values(["employee","period"]))
    _sum = _sum.merge(_admin, on=["employee","period"], how="left")
    _sum["admin_hrs"] = _sum["admin_hrs"].fillna(0)

    fu_headers = ["employee","location","ps_region","role","period","hours_capacity","hours_logged",
                  "credit_hrs","admin_hrs","ff_overrun_hrs","util_pct_vs_logged","util_pct_vs_capacity",
                  "project_util_pct","gap_pts","util_rag"]
    fu_rows = []
    for _, row in _sum.iterrows():
        emp = row["employee"]; period = row["period"]
        loc = df[df["employee"] == emp]["region"].iloc[0] if len(df[df["employee"] == emp]) > 0 else ""
        ps_reg = df[df["employee"] == emp]["ps_region"].iloc[0] if "ps_region" in df.columns and len(df[df["employee"] == emp]) > 0 else ""
        info = EMPLOYEE_ROLES.get(emp, {}); role = info.get("role", "Consultant")
        avail = get_avail_hours(loc, period) if loc else None
        u_log = _pct(row["credit_hrs"], row["hours_logged"])
        u_cap = _pct(row["credit_hrs"], avail)
        u_proj = _pct(row["credit_hrs"] + row["ff_overrun_hrs"], avail)
        gap = round(u_proj - u_cap, 6) if u_proj is not None and u_cap is not None else None
        rag = "Unknown" if u_cap is None else "Green" if u_cap >= 0.70 else "Amber" if u_cap >= 0.60 else "Red"
        fu_rows.append([emp, loc, ps_reg, role, period, avail,
                        round(row["hours_logged"], 2), round(row["credit_hrs"], 2),
                        round(row["admin_hrs"], 2), round(row["ff_overrun_hrs"], 2),
                        u_log, u_cap, u_proj, gap, rag])
    _flat_sheet(wb, "fact_utilization", fu_headers, fu_rows)

    ft_headers = ["employee","location","ps_region","customer_region","project_manager","project",
                  "project_type","billing_type","hrs_to_date","date","hours_logged","approval",
                  "task_case","non_billable","credit_hrs","variance_hrs","previous_hrs_to_date",
                  "credit_tag","period","notes","project_phase","start_date","days_active",
                  "scoped_hrs","variance_flag"]

    def _tbl_scoped(ptype):
        _m = [(k, float(v)) for k, v in scope_map.items() if k.strip().lower() in str(ptype).strip().lower()]
        return round(max(_m, key=lambda x: len(x[0]))[1], 2) if _m else None

    def _tbl_vflag(tag, scoped):
        t = str(tag).strip().upper()
        if t in ("SKIPPED", "NON-BILLABLE", "CREDITED"): return "N/A"
        if t == "UNCONFIGURED": return "No Scope Set"
        if t in ("OVERRUN", "PARTIAL"): return "True Overrun" if scoped and scoped > 0 else "No Scope Set"
        return "Within Budget"

    ft_rows = []
    for _, row in df.iterrows():
        def _g(col, default=""):
            v = row.get(col, default)
            if v is None or (hasattr(v, "__class__") and v.__class__.__name__ == "float" and str(v) == "nan"): return default
            return v
        _tag = str(_g("credit_tag")).strip(); _scoped = _tbl_scoped(_g("project_type"))
        ft_rows.append([
            _g("employee"), _g("region"), _g("ps_region",""), _g("customer_region"), _g("project_manager"),
            _g("project"), _g("project_type"), _g("billing_type"),
            round(float(_g("hours_to_date", 0) or 0), 2),
            str(_g("date"))[:10], round(float(_g("hours", 0) or 0), 2),
            _g("approval"), _g("task", _g("case_task_event", "")),
            1 if str(_g("non_billable","")).lower() in ("true","yes","1","x") else 0,
            round(float(_g("credit_hrs", 0) or 0), 2), round(float(_g("variance_hrs", 0) or 0), 2),
            round(float(_g("previous_htd", 0) or 0), 2),
            _g("credit_tag"), _g("period"), _g("notes",""),
            _g("project_phase"), str(_g("start_date_display", _g("start_date", "")))[:10],
            _g("days_active",""), _scoped if _scoped is not None else "", _tbl_vflag(_tag, _scoped),
        ])
    _flat_sheet(wb, "fact_processed_time_entries", ft_headers, ft_rows)

    fot_headers = ["project_type","scoped_hrs","total_projects","projects_over_budget",
                   "pct_over_budget","total_overrun_hrs","avg_scoped_hrs","avg_actual_hrs","avg_over_under_hrs"]
    _fot_ff = _df[_df.get("billing_type", pd.Series(dtype="object")).str.lower() == "fixed fee"].copy() \
        if "billing_type" in _df.columns else _df.copy()

    def _fot_scope(ptype):
        _m = [(k, float(v)) for k, v in scope_map.items() if k.strip().lower() in str(ptype).strip().lower()]
        return max(_m, key=lambda x: len(x[0]))[1] if _m else None

    _fot_proj = _fot_ff.groupby(["project","project_type"], as_index=False).agg(
        hours_total=("hours","sum"), overrun_hrs=("variance_hrs","sum"))
    _fot_proj["scoped_hrs"] = _fot_proj["project_type"].apply(_fot_scope)
    _fot_proj = _fot_proj[_fot_proj["scoped_hrs"].notna()].copy()
    _fot_proj["over_under"] = _fot_proj["hours_total"].astype(float) - _fot_proj["scoped_hrs"].astype(float)
    _fot_proj["is_over"] = (_fot_proj["over_under"] > 0).astype(int)
    _fot_type = _fot_proj.groupby("project_type", as_index=False).agg(
        scoped_hrs=("scoped_hrs","mean"), total_projects=("project","count"),
        projects_over_budget=("is_over","sum"), total_overrun_hrs=("overrun_hrs","sum"),
        avg_scoped_hrs=("scoped_hrs","mean"), avg_actual_hrs=("hours_total","mean"),
        avg_over_under_hrs=("over_under","mean"),
    ).sort_values("total_overrun_hrs", ascending=False)
    _fot_type["pct_over_budget"] = _fot_type.apply(
        lambda r: round(r["projects_over_budget"] / r["total_projects"], 6) if r["total_projects"] > 0 else None, axis=1)
    fot_rows = []
    for _, r in _fot_type.iterrows():
        fot_rows.append([r["project_type"], round(float(r["scoped_hrs"]),2), int(r["total_projects"]),
                         int(r["projects_over_budget"]), r["pct_over_budget"],
                         round(float(r["total_overrun_hrs"]),2), round(float(r["avg_scoped_hrs"]),2),
                         round(float(r["avg_actual_hrs"]),2), round(float(r["avg_over_under_hrs"]),2)])
    _flat_sheet(wb, "fact_ff_overrun_by_type", fot_headers, fot_rows)

    buf = _io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


if __name__ == "__main__":
    main()
