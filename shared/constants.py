"""
PS Tools — Shared Constants
Employee roster, role assignments, column maps, templates.
Updated: 2026-03-22 v2
"""

# ── Streamlit permission roles ─────────────────────────────────────────────────
# Manager-only: see all pages, do NOT appear in consultant dropdowns
MANAGERS_ONLY = []  # Login handles identity — no manager-only access tier needed

# Manager + Consultant: see all pages AND appear in consultant-scoped views
MANAGER_CONSULTANTS = [
    "Hopkins, Chris",      # Team Lead — Capture & Approvals (NOAM)
    "Ickler, Georganne",   # Consultant & Manager of PS
    "Lappin, Thomas",      # Manager-level Consultant
    "Longi",               # Director of PS
    "Murphy, Conor",       # Solution Architect — manager tier
    "Prince",              # VP of PMO
    "Rusnak",              # VP of PS
    "Snee, Stefanie J",    # Manager-level Consultant
    "Stone, Matt",         # Manager-level Consultant
    "Swanson, Patti",      # Consultant & Director of PS
]

# No access (leavers or no Streamlit access)
NO_ACCESS = [
    "Alam, Laisa",          # No longer employed
    "Chan, Joven",          # No longer employed
    "Eyong, Eyong",         # No longer employed
    "Hernandez, Camila",    # No longer employed
    "Centinaje, Rhodechild",# Left March 16 2026
    "Cloete, Bronwyn",      # Left Feb 23 2026 — not part of PS org
]

def get_role(name: str) -> str:
    """Return 'manager_only' | 'manager' | 'consultant' | 'no_access'."""
    if any(name.startswith(m) for m in MANAGERS_ONLY):
        return "manager_only"
    if name in NO_ACCESS:
        return "no_access"
    if name in MANAGER_CONSULTANTS:
        return "manager"
    return "consultant"

def is_manager(name: str) -> bool:
    return get_role(name) in ("manager", "manager_only")

def is_consultant(name: str) -> bool:
    return get_role(name) in ("consultant", "manager")

# ── Employee roster ────────────────────────────────────────────────────────────
# util_exempt: True = no utilisation targets (managers, solution architects etc.)
EMPLOYEE_ROLES = {
    # ── Project Managers (no product delivery) ────────────────────────────────
    "Barrio, Nairobi":        {"role": "Project Manager",    "products": [], "learning": []},
    "Cadelina, Macoy":        {"role": "Project Manager",    "products": [], "learning": [], "note": "New as of March 3 2026"},
    "Hughes, Madalyn":        {"role": "Project Manager",    "products": [], "learning": []},
    "Porangada, Suraj":       {"role": "Project Manager",    "products": [], "learning": []},
    # ── Solution Architects ───────────────────────────────────────────────────
    "Bell, Stuart":           {"role": "Solution Architect", "products": ["Billing"],     "learning": []},
    "DiMarco, Nicole R":      {"role": "Solution Architect", "products": ["Billing"],     "learning": []},
    "Murphy, Conor":          {"role": "Solution Architect", "products": ["Billing"],     "learning": [], "util_exempt": True},
    "Finalle-Newton, Jesse":  {"role": "Solution Architect", "products": ["Reporting"],   "learning": []},
    # ── Developers ────────────────────────────────────────────────────────────
    "Church, Jason G":        {"role": "Developer",          "products": ["All"],         "learning": [], "util_exempt": True},
    "Dunn, Steven":           {"role": "Developer",          "products": ["All"],         "learning": []},
    "Law, Brandon":           {"role": "Developer",          "products": ["Reporting"],   "learning": []},
    "Quiambao, Generalyn":    {"role": "Developer",          "products": ["All"],         "learning": []},
    # ── Consultants ───────────────────────────────────────────────────────────
    "Arestarkhov, Yaroslav":  {"role": "Consultant", "products": ["Billing", "Capture"],                                                                              "learning": []},
    "Carpen, Anamaria":       {"role": "Consultant", "products": ["Capture", "Approvals", "e-Invoicing"],                                                             "learning": []},
    "Cooke, Ellen":           {"role": "Consultant", "products": ["Billing", "Payroll"],                                                                              "learning": []},
    "Cruz, Daniel":           {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile", "Payments", "e-Invoicing", "SFTP Connector", "CC Statement Import"], "learning": []},
    "Dolha, Madalina":        {"role": "Consultant", "products": ["Capture", "Reconcile", "CC Statement Import", "PSP", "e-Invoicing"],                               "learning": []},
    "Gardner, Cheryll L":     {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Hopkins, Chris":         {"role": "Team Lead",   "products": ["Capture", "Approvals"],                                                                            "learning": []},
    "Ickler, Georganne":      {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Isberg, Eric":           {"role": "Consultant", "products": ["Reporting"],                                                                                       "learning": []},
    "Jordanova, Marija":      {"role": "Consultant", "products": ["Approvals", "Reconcile", "CC Statement Import", "PSP", "SFTP Connector"],                          "learning": []},
    "Lappin, Thomas":         {"role": "Consultant", "products": ["Payroll"],                                                                                         "learning": ["Capture", "Reconcile"]},
    "Longalong, Santiago":    {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile"],                                                               "learning": ["Billing"]},
    "Mohammad, Manaan":       {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile"],                                                               "learning": []},
    "Morris, Lisa":           {"role": "Consultant", "products": ["Payroll"],                                                                                         "learning": []},
    "NAQVI, SYED":            {"role": "Consultant", "products": ["Payroll"],                                                                                         "learning": []},
    "Olson, Austin D":        {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Pallone, Daniel":        {"role": "Consultant", "products": ["Payroll"],                                                                                         "learning": []},
    "Raykova, Silvia":        {"role": "Consultant", "products": ["Capture", "Approvals", "e-Invoicing"],                                                             "learning": []},
    "Selvakumar, Sajithan":   {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile"],                                                               "learning": []},
    "Snee, Stefanie J":       {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Swanson, Patti":         {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": [], "util_exempt": True},
    "Tuazon, Carol":          {"role": "Consultant", "products": ["Payroll", "Reconcile", "CC Statement Import", "PSP", "SFTP Connector"],                            "learning": []},
    "Zoric, Ivan":            {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile", "CC Statement Import", "PSP", "SFTP Connector"],               "learning": []},
    # ── Leadership (managers only — no product delivery) ─────────────────────
    "Longi":                  {"role": "Manager", "products": [], "learning": []},
    "Prince":                 {"role": "Manager", "products": [], "learning": []},
    "Rusnak":                 {"role": "Manager", "products": [], "learning": []},
    # ── Leavers (historical data only — do not remove) ────────────────────────
    "Alam, Laisa":            {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Centinaje, Rhodechild":  {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile", "CC Statement Import", "PSP", "SFTP Connector"],               "learning": []},
    "Chan, Joven":            {"role": "Consultant", "products": ["Capture"],                                                                                         "learning": []},
    "Cloete, Bronwyn":        {"role": "Consultant", "products": ["Capture", "Approvals"],                                                                            "learning": []},
    "Eyong, Eyong":           {"role": "Consultant", "products": ["Capture"],                                                                                         "learning": []},
    "Hamilton, Julie C":      {"role": "Consultant", "products": ["Reporting"],                                                                                       "learning": []},
    "Hernandez, Camila":      {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Rushbrook, Emma C":      {"role": "Consultant", "products": ["Payroll"],                                                                                         "learning": []},
    "Strauss, John W":        {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
}

# Active employees (excludes no-access and leavers)
_LEAVERS = {
    "Alam, Laisa", "Chan, Joven", "Centinaje, Rhodechild", "Cloete, Bronwyn",
    "Eyong, Eyong", "Hamilton, Julie C", "Hernandez, Camila",
    "Rushbrook, Emma C", "Strauss, John W",
}

# Leaver exit dates — used for prorated available hours in team breakdown
# Format: "Name": "YYYY-MM-DD"
LEAVER_EXIT_DATES = {
    "Centinaje, Rhodechild": "2026-03-16",
    "Cloete, Bronwyn":       "2026-02-23",
    "Alam, Laisa":           None,   # date unknown
    "Chan, Joven":           None,
    "Eyong, Eyong":          None,
    "Hamilton, Julie C":     None,
    "Hernandez, Camila":     None,
    "Rushbrook, Emma C":     None,
    "Strauss, John W":       None,
}
ACTIVE_EMPLOYEES = [k for k in EMPLOYEE_ROLES if k not in NO_ACCESS and k not in _LEAVERS]

# Dropdown: consultants + manager-consultants (alphabetical)
CONSULTANT_DROPDOWN = sorted([
    e for e in ACTIVE_EMPLOYEES
    if get_role(e) in ("consultant", "manager")
])

# Dropdown: managers (all tiers)
MANAGER_DROPDOWN = sorted(
    MANAGERS_ONLY +
    [e for e in ACTIVE_EMPLOYEES if e in MANAGER_CONSULTANTS]
)

# ── Column maps ───────────────────────────────────────────────────────────────
SS_COL_MAP = {
    "project name":          "project_name",
    "project id":            "project_id",
    "overall rag":           "rag",
    "start date":            "start_date",
    "go live date":          "go_live_date",
    "% complete":            "pct_complete",
    "project type":          "project_type",
    "status":                "status",
    "project phase":         "phase",
    "client responsiveness": "client_responsiveness",
    "client sentiment":      "client_sentiment",
    "risk level":            "risk_level",
    "schedule health":       "schedule_health",
    "resource health":       "resource_health",
    "scope health":          "scope_health",
    "territory":             "territory",
    "actual hours":          "actual_hours",
    "budgeted hours":        "budgeted_hours",
    "budget":                "budget",
    "change order":          "change_order",
    "partner name":          "partner_name",
    "on hold reason":        "on_hold_reason",
    "project manager":       "project_manager",
    "billing type":          "billing_type",
    "billing":               "billing_type",
}

NS_COL_MAP = {
    "employee":        "employee",
    "name":            "employee",
    "project":         "project",
    "project name":    "project",
    "project id":      "project_id",
    "billing type":    "billing_type",
    "project manager": "project_manager",
    "date":            "date",
    "hours":           "hours",
    "quantity":        "hours",
}

SFDC_COL_MAP = {
    # Exact headers from SFDC contacts export template
    "18 digit opportunity id":   "opportunity_id",
    "first name":                "first_name",
    "last name":                 "last_name",
    "primary":                   "is_primary",
    "title":                     "title",
    "email":                     "email",
    "opportunity owner":         "account_manager",
    "implementation contact exists": "impl_contact_flag",
    "contact roles":             "contact_roles",
    "opp contact role count":    "role_count",
    "partner contact":           "partner_contact",
    "opportunity owner email":   "account_manager_email",
    "account name":              "account",
    "opportunity name":          "opportunity",
    "close date":                "close_date",
    # Fallback aliases
    "opportunity":               "opportunity",
    "account":                   "account",
    "stage":                     "stage",
    "amount":                    "amount",
    "product":                   "product",
    "owner":                     "account_manager",
    "closed date":               "close_date",
    "territory":                 "territory",
    "primary title":             "title",
    "contact name":              "contact_name",
    "contact email":             "email",
    "owner email":               "account_manager_email",
}

# ── WHS phase benchmarks (days) ───────────────────────────────────────────────
PHASE_BENCHMARKS = {
    "Discovery":    14,
    "Planning":     21,
    "Build":        45,
    "UAT":          21,
    "Go-Live":      14,
    "Hypercare":    30,
    "Close-Out":    14,
}

# ── Milestone column map ──────────────────────────────────────────────────────
MILESTONE_COLS_MAP = {
    "ms_intro_email":     "Intro. Email Sent",
    "ms_config_start":    "Standard Config Start",
    "ms_enablement":      "Enablement Session",
    "ms_session1":        "Session #1",
    "ms_session2":        "Session #2",
    "ms_uat_signoff":     "UAT Signoff",
    "ms_prod_cutover":    "Prod Cutover",
    "ms_hypercare_start": "Hypercare Start",
    "ms_close_out":       "Close Out Remaining Tasks",
    "ms_transition":      "Transition to Support",
}

# ── Product keywords ──────────────────────────────────────────────────────────
PRODUCT_KEYWORDS = [
    "Capture", "Approvals", "Reconcile", "PSP", "Payments", "SFTP",
    "E-Invoicing", "eInvoicing", "CC", "Premium", "ZoneCapture",
    "ZoneApprovals", "ZoneReconcile", "ZonePayments", "ZCapture",
    "ZApprovals", "ZReconcile",
]

# ── Fixed Fee scope hours by project type ─────────────────────────────────────
# Source of truth — same table used by Utilization Report engine
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

def get_ff_scope(project_type: str):
    """Return scoped hours for a project type, or None if not found / T&M."""
    if not project_type:
        return None
    pt = str(project_type).strip().lower()
    matches = [(k, float(v)) for k, v in DEFAULT_SCOPE.items() if k.strip().lower() in pt]
    if not matches:
        return None
    return max(matches, key=lambda x: len(x[0]))[1]
